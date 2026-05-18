from __future__ import annotations

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from motor_comissao import calcular_comissoes, salvar_excel

st.set_page_config(
    page_title="Portal de Comissão",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

SENHA_PADRAO = "2026_Geral"
ARQUIVO_PADRAO = "base.xlsx"

CSS = """
<style>
    .block-container {padding-top: 1.3rem; padding-bottom: 2rem; max-width: 100%;}
    .main-title {font-size: 2.15rem; font-weight: 850; margin-bottom: .15rem; color:#111827;}
    .subtitle {color: #6b7280; margin-bottom: 1.2rem; font-size: 1rem;}

    div[data-testid="metric-container"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        padding: 18px 18px 16px 18px;
        border-radius: 18px;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
        min-height: 118px;
        overflow: visible;
    }
    div[data-testid="metric-container"] > label,
    [data-testid="stMetricLabel"] {
        font-size: .92rem !important;
        color: #6b7280 !important;
        white-space: normal !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.55rem !important;
        font-weight: 800 !important;
        color: #111827 !important;
        white-space: normal !important;
        overflow: visible !important;
        line-height: 1.15 !important;
    }
    [data-testid="stMetricDelta"] {
        font-size: .82rem !important;
    }
    .section-title {
        font-size: 1.15rem;
        font-weight: 750;
        margin: .8rem 0 .5rem 0;
        color: #111827;
    }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def moeda(valor: float) -> str:
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def percentual(valor: float) -> str:
    try:
        return f"{float(valor) * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"


def carregar_base(uploaded_file) -> tuple[pd.DataFrame, pd.DataFrame, bytes]:
    """Calcula comissões a partir de upload ou base.xlsx local."""
    if uploaded_file is not None:
        suffix = Path(uploaded_file.name).suffix or ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.getvalue())
            caminho = Path(tmp.name)
    else:
        caminho = Path(__file__).resolve().parent / ARQUIVO_PADRAO
        if not caminho.exists():
            st.error("Envie a base Excel na barra lateral ou coloque o arquivo base.xlsx na pasta do app.")
            st.stop()

    saida, resumo = calcular_comissoes(caminho)

    buffer = io.BytesIO()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
        salvar_excel(saida, resumo, Path(tmp_out.name))
        buffer.write(Path(tmp_out.name).read_bytes())

    return saida, resumo, buffer.getvalue()


def preparar_acessos(df: pd.DataFrame) -> list[str]:
    nomes = set()
    for col in ["Vendedor", "Gerente"]:
        if col in df.columns:
            nomes.update(str(x).strip() for x in df[col].dropna().unique() if str(x).strip())
    return sorted(nomes)


def filtrar_por_acesso(df: pd.DataFrame, usuario: str, perfil_acesso: str) -> pd.DataFrame:
    if perfil_acesso == "Diretoria / Controladoria":
        return df.copy()
    if perfil_acesso == "Gerente":
        return df[df["Gerente"].astype(str).str.upper() == usuario.upper()].copy()
    return df[df["Vendedor"].astype(str).str.upper() == usuario.upper()].copy()


def app_login(nomes: list[str]) -> tuple[str, str]:
    st.sidebar.markdown("### Acesso")
    perfil = st.sidebar.selectbox(
        "Perfil",
        ["Vendedor", "Gerente", "Diretoria / Controladoria"],
    )

    if perfil == "Diretoria / Controladoria":
        usuario = "GERAL"
    else:
        usuario = st.sidebar.selectbox("Usuário", nomes if nomes else ["Sem usuários encontrados"])

    senha = st.sidebar.text_input("Senha", type="password")
    if senha != SENHA_PADRAO:
        st.info("Informe a senha para acessar o mural de comissão.")
        st.stop()
    return usuario, perfil


def aplicar_filtros(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.markdown("### Filtros")

    status = st.sidebar.multiselect(
        "Status da comissão",
        sorted(df["Status Comissão"].dropna().unique()),
        default=sorted(df["Status Comissão"].dropna().unique()),
    )

    operacoes = st.sidebar.multiselect(
        "Operação",
        sorted(df["Operação"].dropna().unique()),
        default=sorted(df["Operação"].dropna().unique()),
    )

    tipo_regra = st.sidebar.multiselect(
        "Tipo de regra",
        sorted(df["Tipo Regra"].dropna().unique()),
        default=sorted(df["Tipo Regra"].dropna().unique()),
    )

    filtrado = df.copy()
    if status:
        filtrado = filtrado[filtrado["Status Comissão"].isin(status)]
    if operacoes:
        filtrado = filtrado[filtrado["Operação"].isin(operacoes)]
    if tipo_regra:
        filtrado = filtrado[filtrado["Tipo Regra"].isin(tipo_regra)]

    pesquisa = st.sidebar.text_input("Pesquisar cliente, NF ou produto")
    if pesquisa:
        p = pesquisa.upper().strip()
        mascara = (
            filtrado["Cliente"].astype(str).str.upper().str.contains(p, na=False)
            | filtrado["Nota Fiscal"].astype(str).str.upper().str.contains(p, na=False)
            | filtrado["Produto"].astype(str).str.upper().str.contains(p, na=False)
            | filtrado["Descrição"].astype(str).str.upper().str.contains(p, na=False)
        )
        filtrado = filtrado[mascara]

    return filtrado


def exibir_cards(df: pd.DataFrame) -> None:
    """Cards principais do mural. Mantidos sem gráficos, com layout maior."""
    faturado = df["Valor Faturado"].sum() if "Valor Faturado" in df.columns else 0
    recebido = df["Valor Recebido NF"].sum() if "Valor Recebido NF" in df.columns else 0
    prev_vend = df["Comissão Prevista Vendedor"].sum() if "Comissão Prevista Vendedor" in df.columns else 0
    lib_vend = df["Comissão Liberada Vendedor"].sum() if "Comissão Liberada Vendedor" in df.columns else 0
    prev_ger = df["Comissão Prevista Gerente"].sum() if "Comissão Prevista Gerente" in df.columns else 0
    lib_ger = df["Comissão Liberada Gerente"].sum() if "Comissão Liberada Gerente" in df.columns else 0
    perc_recebido = 0 if faturado == 0 else min(recebido / faturado, 1)
    saldo_aberto = max(faturado - recebido, 0)

    st.markdown('<div class="section-title">Resumo do período selecionado</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Faturado", moeda(faturado))
    c2.metric("Recebido no período", moeda(recebido))
    c3.metric("Saldo em aberto", moeda(saldo_aberto))
    c4.metric("% Recebido", percentual(perc_recebido))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Comissão prevista vendedor", moeda(prev_vend))
    c6.metric("Comissão liberada vendedor", moeda(lib_vend))
    c7.metric("Comissão prevista gerente", moeda(prev_ger))
    c8.metric("Comissão liberada gerente", moeda(lib_ger))

def main() -> None:
    st.markdown('<div class="main-title">Portal Comercial Financeiro</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Mural de comissão prevista, parcial e liberada com base no faturamento e recebimento.</div>', unsafe_allow_html=True)

    st.sidebar.markdown("### Base de dados")
    uploaded = st.sidebar.file_uploader("Enviar base Excel", type=["xlsx"])

    with st.spinner("Calculando comissões..."):
        saida, resumo, excel_bytes = carregar_base(uploaded)

    nomes = preparar_acessos(saida)
    usuario, perfil_acesso = app_login(nomes)

    df_acesso = filtrar_por_acesso(saida, usuario, perfil_acesso)
    df = aplicar_filtros(df_acesso)

    st.caption(f"Acesso: {perfil_acesso} | Usuário: {usuario} | Linhas exibidas: {len(df):,}".replace(",", "."))

    exibir_cards(df)
    st.divider()

    st.subheader("Detalhamento das comissões")
    colunas_exibir = [
        "Nota Fiscal", "Cliente", "Produto", "Descrição", "Operação", "Finalidade", "Linha",
        "Vendedor", "Perfil Vendedor", "Gerente", "Tipo Regra", "Preço Praticado", "Preço Tabela",
        "Variação %", "Faixa", "% Comissão Vendedor", "% Comissão Gerente", "Valor Faturado",
        "Valor Recebido NF", "% Recebido", "Comissão Prevista Vendedor", "Comissão Liberada Vendedor",
        "Comissão Prevista Gerente", "Comissão Liberada Gerente", "Status Comissão", "Observação",
    ]
    colunas_exibir = [c for c in colunas_exibir if c in df.columns]

    st.dataframe(
        df[colunas_exibir],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Preço Praticado": st.column_config.NumberColumn(format="R$ %.2f"),
            "Preço Tabela": st.column_config.NumberColumn(format="R$ %.2f"),
            "Variação %": st.column_config.NumberColumn(format="%.2f%%"),
            "% Comissão Vendedor": st.column_config.NumberColumn(format="%.2f%%"),
            "% Comissão Gerente": st.column_config.NumberColumn(format="%.2f%%"),
            "Valor Faturado": st.column_config.NumberColumn(format="R$ %.2f"),
            "Valor Recebido NF": st.column_config.NumberColumn(format="R$ %.2f"),
            "% Recebido": st.column_config.NumberColumn(format="%.2f%%"),
            "Comissão Prevista Vendedor": st.column_config.NumberColumn(format="R$ %.2f"),
            "Comissão Liberada Vendedor": st.column_config.NumberColumn(format="R$ %.2f"),
            "Comissão Prevista Gerente": st.column_config.NumberColumn(format="R$ %.2f"),
            "Comissão Liberada Gerente": st.column_config.NumberColumn(format="R$ %.2f"),
        },
    )

    st.download_button(
        "Baixar Excel calculado",
        data=excel_bytes,
        file_name="base_comissoes_calculadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Resumo por vendedor / gerente"):
        st.dataframe(resumo, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
