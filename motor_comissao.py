"""
Motor de Comissão V2 - Portal Comercial Financeiro

Melhorias V2:
- Cruzamento Faturados x Recebido pela Nota Fiscal normalizada, removendo zeros à esquerda.
- Rateio correto do recebimento por NF: calcula % recebido sobre o total da NF e aplica proporcionalmente em cada item.
- Separação entre recebimento total e recebimento dentro do período de apuração.
- Comissão liberada calculada pelo período de recebimento informado, ideal para fechamento 20 a 20.
- Mantém status da comissão pelo recebimento total da NF.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import numpy as np
import pandas as pd

ARQUIVO_ENTRADA = "base.xlsx"
ARQUIVO_SAIDA = "base_comissoes_calculadas.xlsx"
UF_REFERENCIA_PADRAO = "SP"
TIPO_PRECO_PADRAO = "Venda Direta"
COLUNA_BASE_COMISSAO = "Valor Bruto"
PALAVRAS_MICROTECH = ["MICROTECH"]


def normalizar_texto(valor: Any) -> str:
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")
    txt = re.sub(r"\s+", " ", txt)
    return txt.upper()


def normalizar_produto(valor: Any) -> str:
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    if txt.endswith(".0"):
        txt = txt[:-2]
    return txt.upper()


def normalizar_nf(valor: Any) -> str:
    """Normaliza NF para cruzar 001879, 1879.0 e 000001879 como 1879."""
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    if txt.endswith(".0"):
        txt = txt[:-2]
    digitos = re.sub(r"\D", "", txt)
    if not digitos:
        return ""
    return digitos.lstrip("0") or "0"


def produto_base(codigo: str) -> str:
    codigo = normalizar_produto(codigo)
    return re.sub(r"(_RV|_TC|_VD|_CF)$", "", codigo)


def para_numero(valor: Any) -> float:
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float, np.number)):
        return float(valor)
    txt = str(valor).strip().replace("R$", "").replace("%", "").strip()
    if not txt:
        return 0.0
    if "," in txt:
        txt = txt.replace(".", "").replace(",", ".")
    try:
        num = float(txt)
        if "%" in str(valor) and num > 1:
            return num / 100
        return num
    except ValueError:
        return 0.0


def para_percentual(valor: Any) -> float:
    num = para_numero(valor)
    if num > 0.2:
        return num / 100
    return num


def para_data(valor: Any) -> pd.Timestamp | pd.NaT:
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, pd.Timestamp):
        return valor
    if isinstance(valor, (int, float, np.number)):
        # Datas do Excel/Protheus costumam vir como serial.
        try:
            return pd.to_datetime(float(valor), unit="D", origin="1899-12-30")
        except Exception:
            return pd.NaT
    return pd.to_datetime(valor, errors="coerce", dayfirst=True)


def encontrar_coluna(df: pd.DataFrame, opcoes: list[str]) -> Optional[str]:
    mapa = {normalizar_texto(c): c for c in df.columns}
    for opcao in opcoes:
        chave = normalizar_texto(opcao)
        if chave in mapa:
            return mapa[chave]
    return None


def carregar_regras(caminho: Path) -> Tuple[Dict[Tuple[str, str], Any], pd.DataFrame]:
    regras_raw = pd.read_excel(caminho, sheet_name="Regras", header=None)
    regras_fixas: Dict[Tuple[str, str], Any] = {}
    for _, row in regras_raw.iloc[1:].iterrows():
        perfil = normalizar_texto(row.iloc[0])
        finalidade = normalizar_texto(row.iloc[1])
        perc = row.iloc[2]
        if perfil and finalidade:
            if isinstance(perc, str) and normalizar_texto(perc) == "TABELA":
                regras_fixas[(perfil, finalidade)] = "TABELA"
            else:
                regras_fixas[(perfil, finalidade)] = para_percentual(perc)

    matriz = regras_raw.iloc[:, 7:11].copy()
    matriz.columns = ["TIPO", "VARIACAO", "COMISSAO_VENDEDOR", "COMISSAO_GERENTE"]
    matriz = matriz.iloc[2:].dropna(how="all")
    matriz["TIPO_NORM"] = matriz["TIPO"].map(normalizar_texto)
    matriz["VARIACAO_NORM"] = matriz["VARIACAO"].map(normalizar_texto)
    matriz["COMISSAO_VENDEDOR"] = matriz["COMISSAO_VENDEDOR"].map(para_percentual)
    matriz["COMISSAO_GERENTE"] = matriz["COMISSAO_GERENTE"].map(para_percentual)
    return regras_fixas, matriz


def definir_faixa_variacao(tipo_regra: str, variacao: Optional[float]) -> str:
    if variacao is None or pd.isna(variacao):
        return "SEM PRECO TABELA"
    tipo = normalizar_texto(tipo_regra)
    if tipo == "MICROTECH":
        if variacao <= -0.10:
            return "ATE -10%"
        if variacao >= 0.10:
            return "+10% OU SUPERIOR"
        return "PRECO DE TABELA"
    if tipo == "REPRESENTANTE":
        if variacao <= -0.15:
            return "ATE -15%"
        if variacao <= -0.10:
            return "ATE -10%"
        if variacao >= 0.10:
            return "+10% OU SUPERIOR"
        if variacao >= 0.05:
            return "5%"
        return "PRECO DE TABELA"
    if variacao <= -0.15:
        return "ATE -15%"
    if variacao <= -0.10:
        return "ATE -10%"
    if variacao >= 0.15:
        return "+15% OU SUPERIOR"
    if variacao >= 0.10:
        return "10%"
    return "PRECO DE TABELA"


def buscar_percentuais_matriz(matriz: pd.DataFrame, tipo_regra: str, faixa: str) -> Tuple[float, float]:
    tipo = normalizar_texto(tipo_regra)
    faixa_norm = normalizar_texto(faixa)
    achado = matriz.loc[(matriz["TIPO_NORM"] == tipo) & (matriz["VARIACAO_NORM"] == faixa_norm)]
    if achado.empty:
        return 0.0, 0.0
    return float(achado.iloc[0]["COMISSAO_VENDEDOR"]), float(achado.iloc[0]["COMISSAO_GERENTE"])


def carregar_tabela_precos(caminho: Path) -> pd.DataFrame:
    tabela = pd.read_excel(caminho, sheet_name="Tabela de Preços")
    tabela["PRODUTO_KEY"] = tabela["Produto"].map(normalizar_produto)
    tabela["PRODUTO_BASE_KEY"] = tabela["PRODUTO_KEY"].map(produto_base)
    tabela["TIPO_PRECO_NORM"] = tabela["TIPO_PRECO"].map(normalizar_texto) if "TIPO_PRECO" in tabela.columns else normalizar_texto(TIPO_PRECO_PADRAO)
    return tabela


def buscar_preco_tabela(tabela: pd.DataFrame, produto: Any, uf: str = UF_REFERENCIA_PADRAO) -> Tuple[float, str]:
    prod_key = normalizar_produto(produto)
    prod_base_key = produto_base(prod_key)
    tipo_norm = normalizar_texto(TIPO_PRECO_PADRAO)
    candidatos = tabela[(tabela["TIPO_PRECO_NORM"] == tipo_norm) & ((tabela["PRODUTO_KEY"] == prod_key) | (tabela["PRODUTO_BASE_KEY"] == prod_base_key))]
    if candidatos.empty:
        return 0.0, "Produto não encontrado na tabela"
    row = candidatos.iloc[0]
    coluna_preco = uf if uf in tabela.columns else UF_REFERENCIA_PADRAO
    preco = para_numero(row.get(coluna_preco, 0))
    origem = coluna_preco
    if preco <= 0 and "Consumidor Final" in tabela.columns:
        preco = para_numero(row.get("Consumidor Final", 0))
        origem = "Consumidor Final"
    if preco <= 0:
        return 0.0, "Preço zerado/não encontrado"
    return preco, origem


def calcular_comissoes(caminho: Path, data_inicio: Any = None, data_fim: Any = None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    faturados = pd.read_excel(caminho, sheet_name="Faturados")
    recebido = pd.read_excel(caminho, sheet_name="Recebido")
    classificacao = pd.read_excel(caminho, sheet_name="Classificação", header=None)
    regras_fixas, matriz = carregar_regras(caminho)
    tabela_precos = carregar_tabela_precos(caminho)

    col_nf_fat = encontrar_coluna(faturados, ["Nota Fiscal", "NF", "Numero"])
    col_nf_rec = encontrar_coluna(recebido, ["Nota Fiscal", "NF", "Numero"])
    col_valor_recebido = encontrar_coluna(recebido, ["Valor Pago", "Valor", "Valor "])
    col_data_rec = encontrar_coluna(recebido, ["Data", "Data ", "Data Recebimento", "DT Recebimento"])
    if not col_nf_fat or not col_nf_rec:
        raise ValueError("Não encontrei coluna de Nota Fiscal em Faturados e/ou Recebido.")
    if not col_valor_recebido:
        raise ValueError("Não encontrei coluna de Valor Pago na aba Recebido.")

    classificacao.columns = ["NOME", "PERFIL"]
    classificacao["NOME_NORM"] = classificacao["NOME"].map(normalizar_texto)
    mapa_perfil = dict(zip(classificacao["NOME_NORM"], classificacao["PERFIL"].map(normalizar_texto)))

    faturados["NF_KEY"] = faturados[col_nf_fat].map(normalizar_nf)
    faturados["VALOR_FATURADO_ITEM"] = faturados.apply(lambda r: para_numero(r.get(COLUNA_BASE_COMISSAO, r.get("Vlr.Total", 0))), axis=1)
    total_fat_nf = faturados.groupby("NF_KEY", as_index=False)["VALOR_FATURADO_ITEM"].sum().rename(columns={"VALOR_FATURADO_ITEM": "VALOR_FATURADO_NF_TOTAL"})

    recebido["NF_KEY"] = recebido[col_nf_rec].map(normalizar_nf)
    recebido["VALOR_RECEBIDO_NUM"] = recebido[col_valor_recebido].map(para_numero)
    recebido["DATA_RECEBIMENTO"] = recebido[col_data_rec].map(para_data) if col_data_rec else pd.NaT

    di = pd.to_datetime(data_inicio) if data_inicio is not None else None
    dfim = pd.to_datetime(data_fim) if data_fim is not None else None
    recebido_periodo = recebido.copy()
    if di is not None and col_data_rec:
        recebido_periodo = recebido_periodo[recebido_periodo["DATA_RECEBIMENTO"] >= di]
    if dfim is not None and col_data_rec:
        # inclui o dia final completo
        recebido_periodo = recebido_periodo[recebido_periodo["DATA_RECEBIMENTO"] < (dfim + pd.Timedelta(days=1))]

    receb_total = recebido.groupby("NF_KEY", as_index=False).agg(
        VALOR_RECEBIDO_TOTAL=("VALOR_RECEBIDO_NUM", "sum"),
        DATA_RECEBIMENTO_INICIAL=("DATA_RECEBIMENTO", "min"),
        DATA_RECEBIMENTO_FINAL=("DATA_RECEBIMENTO", "max"),
    )
    receb_periodo = recebido_periodo.groupby("NF_KEY", as_index=False).agg(
        VALOR_RECEBIDO_PERIODO=("VALOR_RECEBIDO_NUM", "sum"),
        DATA_RECEBIMENTO_PERIODO_INICIAL=("DATA_RECEBIMENTO", "min"),
        DATA_RECEBIMENTO_PERIODO_FINAL=("DATA_RECEBIMENTO", "max"),
    )

    faturados = faturados.merge(total_fat_nf, on="NF_KEY", how="left")
    faturados = faturados.merge(receb_total, on="NF_KEY", how="left")
    faturados = faturados.merge(receb_periodo, on="NF_KEY", how="left")
    for c in ["VALOR_RECEBIDO_TOTAL", "VALOR_RECEBIDO_PERIODO"]:
        faturados[c] = faturados[c].fillna(0)

    linhas_saida = []
    for _, row in faturados.iterrows():
        nf = row.get(col_nf_fat, "")
        cliente = row.get("Nome Cliente", row.get("Cliente", ""))
        produto = row.get("Produto", "")
        descricao = row.get("DESCRIÇÃO", row.get("Descrição", ""))
        vendedor = row.get("Nome", row.get("Vendedor 1", ""))
        gerente = row.get("GERENTE", "")
        finalidade = normalizar_texto(row.get("FINALIDADE", ""))
        linha = normalizar_texto(row.get("LINHA DE PRODUTO", ""))
        grupo = normalizar_texto(row.get("GRUPO", ""))
        classificacao_prod = normalizar_texto(row.get("CLASSIFICAÇÃO", ""))
        categoria = normalizar_texto(row.get("CATEGORIA", ""))

        valor_faturado = para_numero(row.get("VALOR_FATURADO_ITEM", 0))
        valor_nf_total = para_numero(row.get("VALOR_FATURADO_NF_TOTAL", valor_faturado))
        qtd = para_numero(row.get("Quantidade", 0))
        preco_unit = para_numero(row.get("Prc Unitario", 0))
        if preco_unit <= 0 and qtd > 0:
            preco_unit = valor_faturado / qtd

        valor_recebido_total_nf = para_numero(row.get("VALOR_RECEBIDO_TOTAL", 0))
        valor_recebido_periodo_nf = para_numero(row.get("VALOR_RECEBIDO_PERIODO", 0))
        perc_recebido_total = 0.0 if valor_nf_total <= 0 else min(valor_recebido_total_nf / valor_nf_total, 1.0)
        perc_recebido_periodo = 0.0 if valor_nf_total <= 0 else min(valor_recebido_periodo_nf / valor_nf_total, 1.0)
        valor_recebido_item_total = valor_faturado * perc_recebido_total
        valor_recebido_item_periodo = valor_faturado * perc_recebido_periodo

        operacao = "LOCAÇÃO" if "LOCACAO" in finalidade or "LOCACAO" in categoria or "LOCACAO" in normalizar_texto(row.get("SEGMENTO", "")) else "VENDA"
        perfil_vendedor = mapa_perfil.get(normalizar_texto(vendedor), "VENDEDOR")
        texto_micro = " ".join([linha, grupo, classificacao_prod, normalizar_texto(descricao), normalizar_texto(produto)])
        eh_microtech = any(p in texto_micro for p in PALAVRAS_MICROTECH)
        tipo_regra = "MICROTECH" if eh_microtech else ("REPRESENTANTE" if perfil_vendedor == "REPRESENTANTE" else "VENDEDOR")

        preco_tabela = 0.0
        origem_preco = "Não aplicável"
        variacao = np.nan
        faixa = "Não aplicável"
        perc_vendedor = 0.0
        perc_gerente = 0.0
        observacao = ""

        regra_vendedor = regras_fixas.get(("VENDEDOR", finalidade), 0.0)
        regra_gerente = regras_fixas.get(("GERENTE", finalidade), 0.0)

        if operacao == "LOCAÇÃO":
            perc_vendedor = 0.0 if regra_vendedor == "TABELA" else para_percentual(regra_vendedor)
            perc_gerente = 0.0 if regra_gerente == "TABELA" else para_percentual(regra_gerente)
            faixa = "Regra fixa locação"
        else:
            if regra_vendedor == "TABELA":
                preco_tabela, origem_preco = buscar_preco_tabela(tabela_precos, produto)
                if preco_tabela > 0:
                    variacao = (preco_unit - preco_tabela) / preco_tabela
                    faixa = definir_faixa_variacao(tipo_regra, variacao)
                    perc_vendedor, perc_gerente = buscar_percentuais_matriz(matriz, tipo_regra, faixa)
                    if perc_vendedor == 0 and perc_gerente == 0:
                        observacao = f"Faixa não encontrada na matriz: {tipo_regra} / {faixa}"
                else:
                    faixa = "Sem preço tabela"
                    observacao = origem_preco
            else:
                perc_vendedor = para_percentual(regra_vendedor)
                perc_gerente = para_percentual(regra_gerente)
                faixa = "Regra fixa venda"

        comissao_prev_vendedor = valor_faturado * perc_vendedor
        comissao_prev_gerente = valor_faturado * perc_gerente
        comissao_lib_vendedor_total = comissao_prev_vendedor * perc_recebido_total
        comissao_lib_gerente_total = comissao_prev_gerente * perc_recebido_total
        comissao_lib_vendedor_periodo = comissao_prev_vendedor * perc_recebido_periodo
        comissao_lib_gerente_periodo = comissao_prev_gerente * perc_recebido_periodo

        if valor_recebido_total_nf <= 0:
            status = "PREVISTA"
        elif perc_recebido_total < 0.9999:
            status = "PARCIAL"
        else:
            status = "LIBERADA"

        linhas_saida.append({
            "Nota Fiscal": nf,
            "NF Chave": row.get("NF_KEY", ""),
            "Cliente": cliente,
            "Produto": produto,
            "Descrição": descricao,
            "Finalidade": row.get("FINALIDADE", ""),
            "Operação": operacao,
            "Linha": row.get("LINHA DE PRODUTO", ""),
            "Grupo": row.get("GRUPO", ""),
            "Vendedor": vendedor,
            "Perfil Vendedor": perfil_vendedor,
            "Gerente": gerente,
            "Tipo Regra": tipo_regra,
            "Qtd": qtd,
            "Preço Praticado": preco_unit,
            "Preço Tabela": preco_tabela,
            "Origem Preço": origem_preco,
            "Variação %": variacao if not pd.isna(variacao) else np.nan,
            "Faixa": faixa,
            "% Comissão Vendedor": perc_vendedor,
            "% Comissão Gerente": perc_gerente,
            "Valor Faturado": valor_faturado,
            "Valor Faturado NF Total": valor_nf_total,
            "Valor Recebido NF Total": valor_recebido_total_nf,
            "Valor Recebido Item Total": valor_recebido_item_total,
            "% Recebido Total": perc_recebido_total,
            "Valor Recebido NF Período": valor_recebido_periodo_nf,
            "Valor Recebido Item Período": valor_recebido_item_periodo,
            "% Recebido Período": perc_recebido_periodo,
            "Data Recebimento Inicial": row.get("DATA_RECEBIMENTO_INICIAL", pd.NaT),
            "Data Recebimento Final": row.get("DATA_RECEBIMENTO_FINAL", pd.NaT),
            "Data Recebimento Período Inicial": row.get("DATA_RECEBIMENTO_PERIODO_INICIAL", pd.NaT),
            "Data Recebimento Período Final": row.get("DATA_RECEBIMENTO_PERIODO_FINAL", pd.NaT),
            "Comissão Prevista Vendedor": comissao_prev_vendedor,
            "Comissão Liberada Vendedor Total": comissao_lib_vendedor_total,
            "Comissão Liberada Vendedor": comissao_lib_vendedor_periodo,
            "Comissão Prevista Gerente": comissao_prev_gerente,
            "Comissão Liberada Gerente Total": comissao_lib_gerente_total,
            "Comissão Liberada Gerente": comissao_lib_gerente_periodo,
            "Status Comissão": status,
            "Observação": observacao,
        })

    saida = pd.DataFrame(linhas_saida)
    resumo = saida.groupby(["Vendedor", "Perfil Vendedor", "Gerente", "Status Comissão"], dropna=False).agg(
        Valor_Faturado=("Valor Faturado", "sum"),
        Valor_Recebido_Total=("Valor Recebido Item Total", "sum"),
        Valor_Recebido_Periodo=("Valor Recebido Item Período", "sum"),
        Comissao_Prevista_Vendedor=("Comissão Prevista Vendedor", "sum"),
        Comissao_Liberada_Vendedor_Periodo=("Comissão Liberada Vendedor", "sum"),
        Comissao_Prevista_Gerente=("Comissão Prevista Gerente", "sum"),
        Comissao_Liberada_Gerente_Periodo=("Comissão Liberada Gerente", "sum"),
        Qtde_Notas=("Nota Fiscal", "nunique"),
    ).reset_index()
    return saida, resumo


def salvar_excel(saida: pd.DataFrame, resumo: pd.DataFrame, caminho_saida: Path) -> None:
    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        saida.to_excel(writer, sheet_name="Comissoes_Calculadas", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#12324A", "border": 1})
        money_fmt = workbook.add_format({"num_format": 'R$ #,##0.00'})
        perc_fmt = workbook.add_format({"num_format": '0.00%'})
        date_fmt = workbook.add_format({"num_format": 'dd/mm/yyyy'})
        text_fmt = workbook.add_format({"text_wrap": False})
        formatos_moeda = {c for c in saida.columns if any(k in c for k in ["Valor", "Preço", "Comissão"])} | {c for c in resumo.columns if any(k in c or k.replace("ã","a") in c for k in ["Valor", "Comissao"])}
        formatos_perc = {"Variação %", "% Comissão Vendedor", "% Comissão Gerente", "% Recebido Total", "% Recebido Período"}
        for sheet_name, df in [("Comissoes_Calculadas", saida), ("Resumo", resumo)]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(df), len(df.columns) - 1)
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, header_fmt)
                largura = min(max(len(str(col_name)) + 3, 12), 38)
                fmt = text_fmt
                if col_name in formatos_moeda:
                    fmt = money_fmt
                    largura = max(largura, 16)
                elif col_name in formatos_perc:
                    fmt = perc_fmt
                    largura = max(largura, 14)
                elif "Data" in col_name:
                    fmt = date_fmt
                    largura = max(largura, 14)
                ws.set_column(col_idx, col_idx, largura, fmt)


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    caminho_entrada = base_dir / ARQUIVO_ENTRADA
    caminho_saida = base_dir / ARQUIVO_SAIDA
    if not caminho_entrada.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {caminho_entrada.resolve()}")
    saida, resumo = calcular_comissoes(caminho_entrada)
    salvar_excel(saida, resumo, caminho_saida)
    print("Arquivo gerado com sucesso:", caminho_saida.resolve())
    print("Linhas calculadas:", len(saida))
    print("Comissão liberada vendedor no período:", f"R$ {saida['Comissão Liberada Vendedor'].sum():,.2f}")


if __name__ == "__main__":
    main()
