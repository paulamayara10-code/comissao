from __future__ import annotations

import io
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from motor_comissao import calcular_comissoes, salvar_excel, carregar_inadimplencia

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
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1500px;}
    section[data-testid="stSidebar"] {background: linear-gradient(180deg, #f7fafc 0%, #edf2f7 100%);}
    .main-title {font-size: 2.2rem; font-weight: 900; margin-bottom: .15rem; color: #0f172a; letter-spacing: -.03em;}
    .subtitle {color: #64748b; margin-bottom: 1.1rem; font-size: 1rem;}
    .pill {display:inline-block; padding: .38rem .72rem; border-radius: 999px; background:#e8f0f8; color:#12324a; font-weight:800; margin-right:.4rem; margin-bottom:.35rem; font-size:.84rem;}
    .section-card {border:1px solid #e2e8f0; border-radius:22px; padding:18px; background:#fff; box-shadow:0 12px 28px rgba(15,23,42,.05); margin-bottom:14px;}
    .kpi-grid {display:grid; grid-template-columns: repeat(auto-fit, minmax(215px, 1fr)); gap:14px; margin: 12px 0 16px 0;}
    .kpi-card {background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%); border:1px solid #e2e8f0; border-radius:20px; padding:18px 18px; box-shadow: 0 10px 24px rgba(15,23,42,.05); min-height:112px; overflow: visible;}
    .kpi-label {font-size:.82rem; color:#64748b; font-weight:800; text-transform:uppercase; letter-spacing:.04em; margin-bottom:8px; white-space:normal;}
    .kpi-value {font-size:clamp(1.25rem, 1.5vw, 1.85rem); color:#0f172a; font-weight:950; line-height:1.14; white-space:normal; word-break:break-word;}
    .kpi-help {font-size:.82rem; color:#64748b; margin-top:7px;}
    .status-card {border: 1px solid #e2e8f0; border-radius: 18px; padding: 18px; background: #fff; box-shadow: 0 8px 22px rgba(15,23,42,.04); min-height: 94px;}
    .status-title {font-weight: 900; font-size: 1.05rem; margin-bottom: .3rem; color:#0f172a; word-break:break-word;}
    .small-muted {font-size:.86rem; color:#64748b;}
    div[data-testid="stDataFrame"] {border: 1px solid #e2e8f0; border-radius: 16px; overflow:hidden;}
    h2, h3 {color:#0f172a; letter-spacing:-.02em;}
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
        fim = date(hoje.year + 1, 1, 20) if hoje.month == 12 else date(hoje.year, hoje.month + 1, 20)
    else:
        fim = date(hoje.year, hoje.month, 20)
        inicio = date(hoje.year - 1, 12, 20) if hoje.month == 1 else date(hoje.year, hoje.month - 1, 20)
    return inicio, fim


@st.cache_data(show_spinner=False)
def carregar_base_cache(file_bytes: bytes | None, file_name: str | None, data_inicio: date, data_fim: date) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, bytes]:
    if file_bytes is not None:
        suffix = Path(file_name or "base.xlsx").suffix or ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(file_bytes)
            caminho = Path(tmp.name)
    else:
        caminho = Path(__file__).resolve().parent / ARQUIVO_PADRAO
        if not caminho.exists():
            raise FileNotFoundError("Envie a base Excel na barra lateral ou coloque o arquivo base.xlsx na pasta do app.")

    # Regra central: a comissão liberada do fechamento é apurada SOMENTE pela DATA DO RECEBIMENTO
    # dentro do período escolhido, independente da data de faturamento.
    saida, resumo = calcular_comissoes(caminho, data_inicio=data_inicio, data_fim=data_fim)
    inad = carregar_inadimplencia(caminho, saida)

    buffer = io.BytesIO()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
        salvar_excel(saida, resumo, Path(tmp_out.name), inad)
        buffer.write(Path(tmp_out.name).read_bytes())
    return saida, resumo, inad, buffer.getvalue()


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
    # Vendedor logado enxerga SOMENTE a própria comissão.
    return df[df["Vendedor"].astype(str).str.upper() == usuario.upper()].copy()


def filtrar_inad_por_acesso(df: pd.DataFrame, usuario: str, perfil_acesso: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    if perfil_acesso == "Diretoria / Controladoria":
        return df.copy()
    if perfil_acesso == "Gerente" and "Gerente" in df.columns:
        return df[df["Gerente"].astype(str).str.upper() == usuario.upper()].copy()
    if "Vendedor" in df.columns:
        return df[df["Vendedor"].astype(str).str.upper() == usuario.upper()].copy()
    return pd.DataFrame()


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

    status_lista = sorted(filtrado["Status Período"].dropna().unique()) if len(filtrado) and "Status Período" in filtrado.columns else []
    status = st.sidebar.multiselect("Status do período", status_lista, default=status_lista)
    if status:
        filtrado = filtrado[filtrado["Status Período"].isin(status)]

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


def kpi_card(label: str, value: str, help_text: str = "") -> str:
    return f"""
    <div class='kpi-card'>
        <div class='kpi-label'>{label}</div>
        <div class='kpi-value'>{value}</div>
        <div class='kpi-help'>{help_text}</div>
    </div>
    """


def exibir_cards(df: pd.DataFrame, perfil_acesso: str) -> None:
    faturado = df.get("Valor Faturado", pd.Series(dtype=float)).sum()
    recebido_periodo = df.get("Valor Recebido Item Período", pd.Series(dtype=float)).sum()
    prev_vend = df.get("Comissão Prevista Vendedor", pd.Series(dtype=float)).sum()
    lib_vend = df.get("Comissão Liberada Vendedor", pd.Series(dtype=float)).sum()
    prev_ger = df.get("Comissão Prevista Gerente", pd.Series(dtype=float)).sum()
    lib_ger = df.get("Comissão Liberada Gerente", pd.Series(dtype=float)).sum()
    perc_recebido_periodo = 0 if faturado == 0 else min(recebido_periodo / faturado, 1)

    if perfil_acesso == "Gerente":
        cards = [
            kpi_card("Faturado da equipe", moeda(faturado), "Base de vendas vinculadas ao gerente"),
            kpi_card("Recebido no período", moeda(recebido_periodo), "Considera somente data de recebimento"),
            kpi_card("% recebido no período", percentual(perc_recebido_periodo), "Sobre as vendas exibidas"),
            kpi_card("Comissão gerente prevista", moeda(prev_ger), "Provisão total da carteira filtrada"),
            kpi_card("Comissão gerente liberada", moeda(lib_ger), "Fechamento do período selecionado"),
        ]
    elif perfil_acesso == "Diretoria / Controladoria":
        cards = [
            kpi_card("Faturado filtrado", moeda(faturado), "Independente da data de faturamento"),
            kpi_card("Recebido no período", moeda(recebido_periodo), "Regra oficial da apuração"),
            kpi_card("Comissão vend. liberada", moeda(lib_vend), "Valor a pagar para vendedores/representantes"),
            kpi_card("Comissão ger. liberada", moeda(lib_ger), "Valor a pagar para gerentes"),
            kpi_card("Total comissão liberada", moeda(lib_vend + lib_ger), "Vendedores + gerentes"),
        ]
    else:
        cards = [
            kpi_card("Faturado", moeda(faturado), "Vendas vinculadas ao usuário logado"),
            kpi_card("Recebido no período", moeda(recebido_periodo), "Somente recebimentos no intervalo"),
            kpi_card("% recebido no período", percentual(perc_recebido_periodo), "Sobre as vendas exibidas"),
            kpi_card("Comissão prevista", moeda(prev_vend), "Provisão das vendas exibidas"),
            kpi_card("Comissão liberada", moeda(lib_vend), "Valor efetivo pelo recebimento do período"),
        ]
    st.markdown("<div class='kpi-grid'>" + "".join(cards) + "</div>", unsafe_allow_html=True)


def exibir_status_nf(df: pd.DataFrame) -> None:
    if df.empty or df["Nota Fiscal"].nunique() != 1:
        return
    row = df.iloc[0]
    st.markdown("### Consulta da Nota Fiscal")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"<div class='status-card'><div class='status-title'>NF {row['Nota Fiscal']}</div><div class='small-muted'>{row['Cliente']}</div></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='status-card'><div class='status-title'>{row.get('Status Período', row.get('Status Comissão', ''))}</div><div class='small-muted'>Status pelo recebimento do período</div></div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='status-card'><div class='status-title'>{moeda(df['Valor Recebido Item Período'].sum())}</div><div class='small-muted'>Recebido no período</div></div>", unsafe_allow_html=True)
    with col4:
        st.markdown(f"<div class='status-card'><div class='status-title'>{moeda(df['Comissão Liberada Vendedor'].sum())}</div><div class='small-muted'>Comissão liberada vendedor</div></div>", unsafe_allow_html=True)
    st.write("")


def resumo_filtrado(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return df.groupby(["Vendedor", "Perfil Vendedor", "Gerente", "Status Período"], dropna=False).agg(
        Valor_Faturado=("Valor Faturado", "sum"),
        Valor_Recebido_Periodo=("Valor Recebido Item Período", "sum"),
        Comissao_Prevista_Vendedor=("Comissão Prevista Vendedor", "sum"),
        Comissao_Liberada_Vendedor_Periodo=("Comissão Liberada Vendedor", "sum"),
        Comissao_Prevista_Gerente=("Comissão Prevista Gerente", "sum"),
        Comissao_Liberada_Gerente_Periodo=("Comissão Liberada Gerente", "sum"),
        Qtde_Notas=("Nota Fiscal", "nunique"),
    ).reset_index()


def resumo_inadimplencia(inad: pd.DataFrame) -> pd.DataFrame:
    if inad is None or inad.empty:
        return pd.DataFrame()
    grupo = ["Vendedor"]
    if "Gerente" in inad.columns:
        grupo.append("Gerente")
    valor_col = "Valor Inadimplente" if "Valor Inadimplente" in inad.columns else None
    if not valor_col:
        return pd.DataFrame()
    return inad.groupby(grupo, dropna=False).agg(
        Valor_Inadimplente=(valor_col, "sum"),
        Qtde_Titulos=("Nota Fiscal", "nunique"),
    ).reset_index().sort_values("Valor_Inadimplente", ascending=False)


def exibir_inadimplencia(inad: pd.DataFrame) -> None:
    st.subheader("Inadimplência por vendedor")
    if inad is None or inad.empty:
        st.info("Inclua uma aba chamada Inadimplência/Inadimplencia na base para este painel aparecer. Se tiver NF, o sistema tenta puxar vendedor e gerente automaticamente.")
        return
    resumo = resumo_inadimplencia(inad)
    if resumo.empty:
        st.warning("A aba de inadimplência foi encontrada, mas não identifiquei valor/vendedor para resumir.")
        st.dataframe(inad, use_container_width=True, hide_index=True)
        return
    c1, c2, c3 = st.columns(3)
    c1.markdown(kpi_card("Total inadimplente", moeda(resumo["Valor_Inadimplente"].sum()), "Conforme aba de inadimplência"), unsafe_allow_html=True)
    c2.markdown(kpi_card("Títulos em aberto", f"{int(resumo['Qtde_Titulos'].sum())}", "Quantidade de notas/títulos"), unsafe_allow_html=True)
    maior = resumo.iloc[0]["Vendedor"] if len(resumo) else "-"
    c3.markdown(kpi_card("Maior concentração", str(maior), "Vendedor com maior saldo"), unsafe_allow_html=True)
    st.dataframe(
        resumo,
        use_container_width=True,
        hide_index=True,
        column_config={"Valor_Inadimplente": st.column_config.NumberColumn(format="R$ %.2f")},
    )
    with st.expander("Detalhamento da inadimplência"):
        st.dataframe(
            inad,
            use_container_width=True,
            hide_index=True,
            column_config={"Valor Inadimplente": st.column_config.NumberColumn(format="R$ %.2f")},
        )


def main() -> None:
    st.markdown('<div class="main-title">Portal Comercial Financeiro</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Comissão liberada sempre pela data de recebimento do período selecionado, independente da data do faturamento.</div>', unsafe_allow_html=True)

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
    st.sidebar.caption("Regra: comissão liberada = recebimentos com data dentro do intervalo.")

    try:
        with st.spinner("Calculando comissões por data de recebimento e conciliando bases..."):
            saida, resumo, inad, excel_bytes = carregar_base_cache(file_bytes, file_name, data_inicio, data_fim)
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    usuario, perfil_acesso = app_login(saida)
    df_acesso = filtrar_por_acesso(saida, usuario, perfil_acesso)
    inad_acesso = filtrar_inad_por_acesso(inad, usuario, perfil_acesso)
    df = aplicar_filtros(df_acesso)

    st.markdown(
        f"<span class='pill'>Acesso: {perfil_acesso}</span>"
        f"<span class='pill'>Usuário: {usuario}</span>"
        f"<span class='pill'>Recebimento: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}</span>"
        f"<span class='pill'>Linhas: {len(df):,}</span>".replace(",", "."),
        unsafe_allow_html=True,
    )

    exibir_cards(df, perfil_acesso)
    exibir_status_nf(df)

    tab1, tab2, tab3 = st.tabs(["Comissões", "Inadimplência", "Resumo"])

    with tab1:
        st.subheader("Detalhamento das comissões")
        if df.empty:
            st.warning("Nenhum registro encontrado para os filtros selecionados.")
        else:
            colunas_exibir = [
                "Nota Fiscal", "Cliente", "Produto", "Descrição", "Operação", "Finalidade", "Linha",
                "Vendedor", "Perfil Vendedor", "Gerente", "Tipo Regra", "Preço Praticado", "Preço Tabela",
                "Variação %", "Faixa", "% Comissão Vendedor", "% Comissão Gerente", "Valor Faturado",
                "Valor Faturado NF Total", "Valor Recebido Item Período", "% Recebido Período",
                "Data Recebimento Período Inicial", "Data Recebimento Período Final",
                "Comissão Prevista Vendedor", "Comissão Liberada Vendedor",
                "Comissão Prevista Gerente", "Comissão Liberada Gerente", "Status Período", "Status Comissão", "Observação",
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
                    "Valor Recebido Item Período": st.column_config.NumberColumn(format="R$ %.2f"),
                    "% Recebido Período": st.column_config.NumberColumn(format="%.2f%%"),
                    "Comissão Prevista Vendedor": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Comissão Liberada Vendedor": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Comissão Prevista Gerente": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Comissão Liberada Gerente": st.column_config.NumberColumn(format="R$ %.2f"),
                },
            )

    with tab2:
        exibir_inadimplencia(inad_acesso)

    with tab3:
        st.subheader("Resumo filtrado por vendedor / gerente")
        st.dataframe(resumo_filtrado(df), use_container_width=True, hide_index=True)

    st.download_button(
        "Baixar Excel calculado",
        data=excel_bytes,
        file_name="base_comissoes_calculadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
