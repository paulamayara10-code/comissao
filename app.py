from __future__ import annotations

import io
import tempfile
from datetime import date, datetime
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
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    section[data-testid="stSidebar"] {background: linear-gradient(180deg, #f4f7fb 0%, #edf2f7 100%);}
    .main-title {font-size: 2.1rem; font-weight: 850; margin-bottom: .15rem; color: #0f172a;}
    .subtitle {color: #64748b; margin-bottom: 1.1rem; font-size: 1rem;}
    .pill {display:inline-block; padding: .35rem .7rem; border-radius: 999px; background:#e8f0f8; color:#12324a; font-weight:700; margin-right:.4rem; font-size:.85rem;}
    div[data-testid="stMetric"] {background: #ffffff; border: 1px solid #e2e8f0; border-radius: 18px; padding: 16px 16px; box-shadow: 0 8px 22px rgba(15,23,42,.05);}
    [data-testid="stMetricValue"] {font-size: 1.45rem; font-weight: 800; color:#0f172a;}
    [data-testid="stMetricLabel"] {font-size: .88rem; color:#475569;}
    .status-card {border: 1px solid #e2e8f0; border-radius: 18px; padding: 18px; background: #fff; box-shadow: 0 8px 22px rgba(15,23,42,.04);}
    .status-title {font-weight: 800; font-size: 1.05rem; margin-bottom: .3rem; color:#0f172a;}
    .small-muted {font-size:.86rem; color:#64748b;}
    div[data-testid="stDataFrame"] {border: 1px solid #e2e8f0; border-radius: 16px; overflow:hidden;}
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


def periodo_20_20_padrao() -> tuple[date, date]:
    hoje = date.today()
    if hoje.day >= 20:
        inicio = date(hoje.year, hoje.month, 20)
        if hoje.month == 12:
            fim = date(hoje.year + 1, 1, 20)
        else:
            fim = date(hoje.year, hoje.month + 1, 20)
    else:
        fim = date(hoje.year, hoje.month, 20)
        if hoje.month == 1:
            inicio = date(hoje.year - 1, 12, 20)
        else:
            inicio = date(hoje.year, hoje.month - 1, 20)
    return inicio, fim


@st.cache_data(show_spinner=False)
def carregar_base_cache(file_bytes: bytes | None, file_name: str | None, data_inicio: date, data_fim: date) -> tuple[pd.DataFrame, pd.DataFrame, bytes]:
    if file_bytes is not None:
        suffix = Path(file_name or "base.xlsx").suffix or ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(file_bytes)
            caminho = Path(tmp.name)
    else:
        caminho = Path(__file__).resolve().parent / ARQUIVO_PADRAO
        if not caminho.exists():
            raise FileNotFoundError("Envie a base Excel na barra lateral ou coloque o arquivo base.xlsx na pasta do app.")

    saida, resumo = calcular_comissoes(caminho, data_inicio=data_inicio, data_fim=data_fim)

    buffer = io.BytesIO()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
        salvar_excel(saida, resumo, Path(tmp_out.name))
        buffer.write(Path(tmp_out.name).read_bytes())
    return saida, resumo, buffer.getvalue()


def preparar_acessos(df: pd.DataFrame, perfil: str) -> list[str]:
    col = "Gerente" if perfil == "Gerente" else "Vendedor"
    if col not in df.columns:
        return []
    nomes = [str(x).strip() for x in df[col].dropna().unique() if str(x).strip()]
    return sorted(nomes)


def filtrar_por_acesso(df: pd.DataFrame, usuario: str, perfil_acesso: str) -> pd.DataFrame:
    if perfil_acesso == "Diretoria / Controladoria":
        return df.copy()
    if perfil_acesso == "Gerente":
        return df[df["Gerente"].astype(str).str.upper() == usuario.upper()].copy()
    return df[df["Vendedor"].astype(str).str.upper() == usuario.upper()].copy()


def app_login(saida: pd.DataFrame) -> tuple[str, str]:
    st.sidebar.markdown("### Acesso")
    perfil = st.sidebar.selectbox("Perfil", ["Vendedor", "Gerente", "Diretoria / Controladoria"])
    if perfil == "Diretoria / Controladoria":
        usuario = "GERAL"
    else:
        nomes = preparar_acessos(saida, perfil)
        usuario = st.sidebar.selectbox("Usuário", nomes if nomes else ["Sem usuários encontrados"])
    senha = st.sidebar.text_input("Senha", type="password")
    if senha != SENHA_PADRAO:
        st.info("Informe a senha para acessar o mural de comissão.")
        st.stop()
    return usuario, perfil


def aplicar_filtros(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.markdown("### Filtros do mural")
    filtrado = df.copy()

    nf_opcoes = ["Todas"] + sorted([str(x) for x in filtrado["Nota Fiscal"].dropna().unique()])
    nf_especifica = st.sidebar.selectbox("Consultar NF específica", nf_opcoes)
    if nf_especifica != "Todas":
        filtrado = filtrado[filtrado["Nota Fiscal"].astype(str) == nf_especifica]

    status_lista = sorted(filtrado["Status Comissão"].dropna().unique()) if len(filtrado) else []
    status = st.sidebar.multiselect("Status da comissão", status_lista, default=status_lista)
    if status:
        filtrado = filtrado[filtrado["Status Comissão"].isin(status)]

    operacoes_lista = sorted(filtrado["Operação"].dropna().unique()) if len(filtrado) else []
    operacoes = st.sidebar.multiselect("Operação", operacoes_lista, default=operacoes_lista)
    if operacoes:
        filtrado = filtrado[filtrado["Operação"].isin(operacoes)]

    tipo_lista = sorted(filtrado["Tipo Regra"].dropna().unique()) if len(filtrado) else []
    tipo_regra = st.sidebar.multiselect("Tipo de regra", tipo_lista, default=tipo_lista)
    if tipo_regra:
        filtrado = filtrado[filtrado["Tipo Regra"].isin(tipo_regra)]

    pesquisa = st.sidebar.text_input("Pesquisar cliente, produto ou descrição")
    if pesquisa:
        p = pesquisa.upper().strip()
        mascara = (
            filtrado["Cliente"].astype(str).str.upper().str.contains(p, na=False)
            | filtrado["Produto"].astype(str).str.upper().str.contains(p, na=False)
            | filtrado["Descrição"].astype(str).str.upper().str.contains(p, na=False)
        )
        filtrado = filtrado[mascara]
    return filtrado


def exibir_cards(df: pd.DataFrame, perfil_acesso: str) -> None:
    faturado = df["Valor Faturado"].sum()
    recebido_total = df.get("Valor Recebido Item Total", pd.Series(dtype=float)).sum()
    recebido_periodo = df.get("Valor Recebido Item Período", pd.Series(dtype=float)).sum()
    prev_vend = df["Comissão Prevista Vendedor"].sum()
    lib_vend = df["Comissão Liberada Vendedor"].sum()
    prev_ger = df["Comissão Prevista Gerente"].sum()
    lib_ger = df["Comissão Liberada Gerente"].sum()
    perc_recebido = 0 if faturado == 0 else min(recebido_total / faturado, 1)

    if perfil_acesso == "Gerente":
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Faturado da equipe", moeda(faturado))
        c2.metric("Recebido total", moeda(recebido_total))
        c3.metric("Recebido no período", moeda(recebido_periodo))
        c4.metric("Comissão gerente prevista", moeda(prev_ger))
        c5.metric("Comissão gerente liberada", moeda(lib_ger))
    else:
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Faturado", moeda(faturado))
        c2.metric("Recebido total", moeda(recebido_total))
        c3.metric("Recebido período", moeda(recebido_periodo))
        c4.metric("% Recebido", percentual(perc_recebido))
        c5.metric("Comissão prevista", moeda(prev_vend))
        c6.metric("Comissão liberada", moeda(lib_vend), help=f"Gerente liberada: {moeda(lib_ger)}")


def exibir_status_nf(df: pd.DataFrame) -> None:
    if df.empty:
        return
    nfs = df["Nota Fiscal"].nunique()
    if nfs != 1:
        return
    row = df.iloc[0]
    st.markdown("### Consulta da Nota Fiscal")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"<div class='status-card'><div class='status-title'>NF {row['Nota Fiscal']}</div><div class='small-muted'>{row['Cliente']}</div></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='status-card'><div class='status-title'>{row['Status Comissão']}</div><div class='small-muted'>Status atual da comissão</div></div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='status-card'><div class='status-title'>{moeda(df['Comissão Prevista Vendedor'].sum())}</div><div class='small-muted'>Comissão prevista vendedor</div></div>", unsafe_allow_html=True)
    with col4:
        st.markdown(f"<div class='status-card'><div class='status-title'>{moeda(df['Comissão Liberada Vendedor'].sum())}</div><div class='small-muted'>Liberada no período</div></div>", unsafe_allow_html=True)
    st.write("")


def exibir_graficos(df: pd.DataFrame, perfil_acesso: str) -> None:
    if df.empty:
        st.warning("Nenhum registro encontrado para os filtros selecionados.")
        return
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Comissão liberada por status")
        status_df = df.groupby("Status Comissão", as_index=False)[["Comissão Prevista Vendedor", "Comissão Liberada Vendedor"]].sum()
        st.bar_chart(status_df.set_index("Status Comissão"))
    with col2:
        st.subheader("Faturado por operação")
        op_df = df.groupby("Operação", as_index=False)["Valor Faturado"].sum().sort_values("Valor Faturado", ascending=False)
        st.bar_chart(op_df.set_index("Operação"))


def resumo_filtrado(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    return df.groupby(["Vendedor", "Perfil Vendedor", "Gerente", "Status Comissão"], dropna=False).agg(
        Valor_Faturado=("Valor Faturado", "sum"),
        Valor_Recebido_Total=("Valor Recebido Item Total", "sum"),
        Valor_Recebido_Periodo=("Valor Recebido Item Período", "sum"),
        Comissao_Prevista_Vendedor=("Comissão Prevista Vendedor", "sum"),
        Comissao_Liberada_Vendedor_Periodo=("Comissão Liberada Vendedor", "sum"),
        Comissao_Prevista_Gerente=("Comissão Prevista Gerente", "sum"),
        Comissao_Liberada_Gerente_Periodo=("Comissão Liberada Gerente", "sum"),
        Qtde_Notas=("Nota Fiscal", "nunique"),
    ).reset_index()


def main() -> None:
    st.markdown('<div class="main-title">Portal Comercial Financeiro</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Mural de comissão prevista e liberada com apuração por recebimento no período 20 a 20.</div>', unsafe_allow_html=True)

    st.sidebar.markdown("### Base de dados")
    uploaded = st.sidebar.file_uploader("Enviar base Excel", type=["xlsx"])
    file_bytes = uploaded.getvalue() if uploaded is not None else None
    file_name = uploaded.name if uploaded is not None else None

    st.sidebar.markdown("### Período de recebimento")
    ini_padrao, fim_padrao = periodo_20_20_padrao()
    periodo = st.sidebar.date_input("Apuração da comissão", value=(ini_padrao, fim_padrao), format="DD/MM/YYYY")
    if isinstance(periodo, tuple) and len(periodo) == 2:
        data_inicio, data_fim = periodo
    else:
        data_inicio, data_fim = ini_padrao, fim_padrao
    st.sidebar.caption("A comissão liberada considera os recebimentos dentro deste intervalo.")

    try:
        with st.spinner("Calculando comissões e conciliando recebimentos..."):
            saida, resumo, excel_bytes = carregar_base_cache(file_bytes, file_name, data_inicio, data_fim)
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    usuario, perfil_acesso = app_login(saida)
    df_acesso = filtrar_por_acesso(saida, usuario, perfil_acesso)
    df = aplicar_filtros(df_acesso)

    st.markdown(
        f"<span class='pill'>Acesso: {perfil_acesso}</span>"
        f"<span class='pill'>Usuário: {usuario}</span>"
        f"<span class='pill'>Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}</span>"
        f"<span class='pill'>Linhas: {len(df):,}</span>".replace(",", "."),
        unsafe_allow_html=True,
    )
    st.write("")

    exibir_cards(df, perfil_acesso)
    st.divider()
    exibir_status_nf(df)
    exibir_graficos(df, perfil_acesso)
    st.divider()

    st.subheader("Detalhamento das comissões")
    colunas_exibir = [
        "Nota Fiscal", "Cliente", "Produto", "Descrição", "Operação", "Finalidade", "Linha",
        "Vendedor", "Perfil Vendedor", "Gerente", "Tipo Regra", "Preço Praticado", "Preço Tabela",
        "Variação %", "Faixa", "% Comissão Vendedor", "% Comissão Gerente", "Valor Faturado",
        "Valor Faturado NF Total", "Valor Recebido Item Total", "% Recebido Total",
        "Valor Recebido Item Período", "% Recebido Período", "Data Recebimento Período Inicial",
        "Data Recebimento Período Final", "Comissão Prevista Vendedor", "Comissão Liberada Vendedor",
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
            "Valor Faturado NF Total": st.column_config.NumberColumn(format="R$ %.2f"),
            "Valor Recebido Item Total": st.column_config.NumberColumn(format="R$ %.2f"),
            "% Recebido Total": st.column_config.NumberColumn(format="%.2f%%"),
            "Valor Recebido Item Período": st.column_config.NumberColumn(format="R$ %.2f"),
            "% Recebido Período": st.column_config.NumberColumn(format="%.2f%%"),
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

    with st.expander("Resumo filtrado por vendedor / gerente"):
        st.dataframe(resumo_filtrado(df), use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
