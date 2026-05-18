"""
Motor de Comissão V1 - Portal Comercial Financeiro
Autor: Paula / ChatGPT

Objetivo:
- Ler a base Excel com as abas: Faturados, Recebido, Regras, Classificação e Tabela de Preços.
- Cruzar Faturados x Recebido pela Nota Fiscal.
- Confrontar preço praticado x tabela de preços somente para operações de venda.
- Aplicar regras para Vendedor, Representante, Gerente e Microtech.
- Gerar arquivo Excel com a aba Comissoes_Calculadas e Resumo.

Como usar:
1) Coloque este arquivo na mesma pasta da sua base Excel.
2) Ajuste o nome do arquivo em ARQUIVO_ENTRADA, se necessário.
3) Rode: python motor_comissao.py
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import numpy as np
import pandas as pd

# =========================
# CONFIGURAÇÕES PRINCIPAIS
# =========================
ARQUIVO_ENTRADA = "base.xlsx"
ARQUIVO_SAIDA = "base_comissoes_calculadas.xlsx"

# Como a tabela possui preços por UF e também Consumidor Final, usamos SP como padrão.
# Se futuramente a base de faturados tiver UF do cliente, podemos trocar para leitura dinâmica por UF.
UF_REFERENCIA_PADRAO = "SP"
TIPO_PRECO_PADRAO = "Venda Direta"

# Base financeira para cálculo da comissão.
# Na base do Protheus, Vlr.Total pode vir como valor total/contratual repetido em várias linhas;
# por isso usamos Valor Bruto como padrão. Se preferir, altere para "Vlr.Total".
COLUNA_BASE_COMISSAO = "Valor Bruto"

# Margem/regra especial Microtech já está parametrizada na matriz da aba Regras.
# O sistema identifica Microtech pela linha/grupo/classificação/descrição/produto.
PALAVRAS_MICROTECH = ["MICROTECH"]


# =========================
# FUNÇÕES AUXILIARES
# =========================
def normalizar_texto(valor: Any) -> str:
    """Padroniza texto para comparação: sem acento, maiúsculo e sem espaços duplicados."""
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")
    txt = re.sub(r"\s+", " ", txt)
    return txt.upper()


def normalizar_produto(valor: Any) -> str:
    """Padroniza código de produto preservando zeros à esquerda quando vier como texto."""
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    if txt.endswith(".0"):
        txt = txt[:-2]
    return txt.upper()


def produto_base(codigo: str) -> str:
    """Remove sufixos comuns para tentar uma segunda chave de busca de preço."""
    codigo = normalizar_produto(codigo)
    return re.sub(r"(_RV|_TC|_VD|_CF)$", "", codigo)


def para_numero(valor: Any) -> float:
    """Converte números em formatos pt-BR/Excel para float."""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float, np.number)):
        return float(valor)
    txt = str(valor).strip().replace("R$", "").replace("%", "").strip()
    if txt == "":
        return 0.0
    # Se vier como 1.234,56
    if "," in txt:
        txt = txt.replace(".", "").replace(",", ".")
    try:
        num = float(txt)
        # Se for percentual em texto tipo 6,0%, já removemos %, então vira 6.0; converter para 0.06
        if "%" in str(valor) and num > 1:
            return num / 100
        return num
    except ValueError:
        return 0.0



def para_percentual(valor: Any) -> float:
    """Converte percentual de comissão. Trata 1 como 1%, 6 como 6%, 0.06 como 6%."""
    num = para_numero(valor)
    if num > 0.2:
        return num / 100
    return num

def encontrar_coluna(df: pd.DataFrame, opcoes: list[str]) -> Optional[str]:
    """Encontra uma coluna mesmo com pequenas variações de nome."""
    mapa = {normalizar_texto(c): c for c in df.columns}
    for opcao in opcoes:
        chave = normalizar_texto(opcao)
        if chave in mapa:
            return mapa[chave]
    return None


def formatar_percentual(p: float) -> str:
    return f"{p:.2%}".replace(".", ",")


# =========================
# LEITURA DAS REGRAS
# =========================
def carregar_regras(caminho: Path) -> Tuple[Dict[Tuple[str, str], Any], pd.DataFrame]:
    """Carrega regras fixas da esquerda e matriz dinâmica da direita da aba Regras."""
    regras_raw = pd.read_excel(caminho, sheet_name="Regras", header=None)

    # Tabela esquerda: PERFIL | FINALIDADE | %
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

    # Matriz direita: colunas H:K, com cabeçalho na linha 2 do Excel = índice 1
    matriz = regras_raw.iloc[:, 7:11].copy()
    matriz.columns = ["TIPO", "VARIACAO", "COMISSAO_VENDEDOR", "COMISSAO_GERENTE"]
    matriz = matriz.iloc[2:].dropna(how="all")
    matriz["TIPO_NORM"] = matriz["TIPO"].map(normalizar_texto)
    matriz["VARIACAO_NORM"] = matriz["VARIACAO"].map(normalizar_texto)
    matriz["COMISSAO_VENDEDOR"] = matriz["COMISSAO_VENDEDOR"].map(para_percentual)
    matriz["COMISSAO_GERENTE"] = matriz["COMISSAO_GERENTE"].map(para_percentual)
    return regras_fixas, matriz


def definir_faixa_variacao(tipo_regra: str, variacao: Optional[float]) -> str:
    """Transforma variação numérica em faixa da matriz de regras."""
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

    # VENDEDOR padrão
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
    filtro = (matriz["TIPO_NORM"] == tipo) & (matriz["VARIACAO_NORM"] == faixa_norm)
    achado = matriz.loc[filtro]
    if achado.empty:
        return 0.0, 0.0
    return float(achado.iloc[0]["COMISSAO_VENDEDOR"]), float(achado.iloc[0]["COMISSAO_GERENTE"])


# =========================
# TABELA DE PREÇOS
# =========================
def carregar_tabela_precos(caminho: Path) -> pd.DataFrame:
    tabela = pd.read_excel(caminho, sheet_name="Tabela de Preços")
    tabela["PRODUTO_KEY"] = tabela["Produto"].map(normalizar_produto)
    tabela["PRODUTO_BASE_KEY"] = tabela["PRODUTO_KEY"].map(produto_base)
    if "TIPO_PRECO" in tabela.columns:
        tabela["TIPO_PRECO_NORM"] = tabela["TIPO_PRECO"].map(normalizar_texto)
    else:
        tabela["TIPO_PRECO_NORM"] = normalizar_texto(TIPO_PRECO_PADRAO)
    return tabela


def buscar_preco_tabela(tabela: pd.DataFrame, produto: Any, uf: str = UF_REFERENCIA_PADRAO) -> Tuple[float, str]:
    """Busca preço por Produto + Tipo Preço + UF. Faz fallback para Consumidor Final."""
    prod_key = normalizar_produto(produto)
    prod_base_key = produto_base(prod_key)
    tipo_norm = normalizar_texto(TIPO_PRECO_PADRAO)

    candidatos = tabela[
        (tabela["TIPO_PRECO_NORM"] == tipo_norm)
        & ((tabela["PRODUTO_KEY"] == prod_key) | (tabela["PRODUTO_BASE_KEY"] == prod_base_key))
    ]

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


# =========================
# MOTOR PRINCIPAL
# =========================
def calcular_comissoes(caminho: Path) -> Tuple[pd.DataFrame, pd.DataFrame]:
    faturados = pd.read_excel(caminho, sheet_name="Faturados")
    recebido = pd.read_excel(caminho, sheet_name="Recebido")
    classificacao = pd.read_excel(caminho, sheet_name="Classificação", header=None)
    regras_fixas, matriz = carregar_regras(caminho)
    tabela_precos = carregar_tabela_precos(caminho)

    # Padronização de colunas principais
    col_nf_fat = encontrar_coluna(faturados, ["Nota Fiscal", "NF", "Numero"])
    col_nf_rec = encontrar_coluna(recebido, ["Nota Fiscal", "NF", "Numero"])
    col_valor_recebido = encontrar_coluna(recebido, ["Valor Pago", "Valor", "Valor "])

    if not col_nf_fat or not col_nf_rec:
        raise ValueError("Não encontrei coluna de Nota Fiscal em Faturados e/ou Recebido.")
    if not col_valor_recebido:
        raise ValueError("Não encontrei coluna de Valor Pago na aba Recebido.")

    # Mapa vendedor/representante
    classificacao.columns = ["NOME", "PERFIL"]
    classificacao["NOME_NORM"] = classificacao["NOME"].map(normalizar_texto)
    mapa_perfil = dict(zip(classificacao["NOME_NORM"], classificacao["PERFIL"].map(normalizar_texto)))

    # Recebimentos por NF
    recebido["NF_KEY"] = recebido[col_nf_rec].map(normalizar_produto)
    recebido["VALOR_RECEBIDO_NUM"] = recebido[col_valor_recebido].map(para_numero)
    receb_por_nf = recebido.groupby("NF_KEY", as_index=False)["VALOR_RECEBIDO_NUM"].sum()

    # Preparar faturados
    faturados["NF_KEY"] = faturados[col_nf_fat].map(normalizar_produto)
    faturados = faturados.merge(receb_por_nf, on="NF_KEY", how="left")
    faturados["VALOR_RECEBIDO_NUM"] = faturados["VALOR_RECEBIDO_NUM"].fillna(0)

    linhas_saida = []

    for _, row in faturados.iterrows():
        nf = row.get(col_nf_fat, "")
        cliente = row.get("Nome Cliente", row.get("Cliente", ""))
        produto = row.get("Produto", "")
        descricao = row.get("DESCRIÇÃO", "")
        vendedor = row.get("Nome", row.get("Vendedor 1", ""))
        gerente = row.get("GERENTE", "")
        finalidade = normalizar_texto(row.get("FINALIDADE", ""))
        linha = normalizar_texto(row.get("LINHA DE PRODUTO", ""))
        grupo = normalizar_texto(row.get("GRUPO", ""))
        classificacao_prod = normalizar_texto(row.get("CLASSIFICAÇÃO", ""))
        categoria = normalizar_texto(row.get("CATEGORIA", ""))

        valor_faturado = para_numero(row.get(COLUNA_BASE_COMISSAO, row.get("Vlr.Total", 0)))
        qtd = para_numero(row.get("Quantidade", 0))
        preco_unit = para_numero(row.get("Prc Unitario", 0))
        if preco_unit <= 0 and qtd > 0:
            preco_unit = valor_faturado / qtd

        valor_recebido_nf = para_numero(row.get("VALOR_RECEBIDO_NUM", 0))
        perc_recebido = 0.0 if valor_faturado <= 0 else min(valor_recebido_nf / valor_faturado, 1.0)

        operacao = "LOCAÇÃO" if "LOCACAO" in finalidade or "LOCACAO" in categoria else "VENDA"

        vendedor_norm = normalizar_texto(vendedor)
        perfil_vendedor = mapa_perfil.get(vendedor_norm, "VENDEDOR")

        texto_micro = " ".join([linha, grupo, classificacao_prod, normalizar_texto(descricao), normalizar_texto(produto)])
        eh_microtech = any(p in texto_micro for p in PALAVRAS_MICROTECH)

        # Define tipo da matriz para vendas com preço em tabela
        if eh_microtech:
            tipo_regra = "MICROTECH"
        elif perfil_vendedor == "REPRESENTANTE":
            tipo_regra = "REPRESENTANTE"
        else:
            tipo_regra = "VENDEDOR"

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
            # Locação não tem variação por preço. Usa regra fixa da finalidade.
            perc_vendedor = 0.0 if regra_vendedor == "TABELA" else para_numero(regra_vendedor)
            perc_gerente = 0.0 if regra_gerente == "TABELA" else para_numero(regra_gerente)
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
                # Venda com regra fixa por finalidade, sem confronto com tabela.
                perc_vendedor = para_percentual(regra_vendedor)
                perc_gerente = para_percentual(regra_gerente)
                faixa = "Regra fixa venda"

        comissao_prev_vendedor = valor_faturado * perc_vendedor
        comissao_prev_gerente = valor_faturado * perc_gerente
        comissao_lib_vendedor = comissao_prev_vendedor * perc_recebido
        comissao_lib_gerente = comissao_prev_gerente * perc_recebido

        if valor_recebido_nf <= 0:
            status = "PREVISTA"
        elif perc_recebido < 0.9999:
            status = "PARCIAL"
        else:
            status = "LIBERADA"

        linhas_saida.append(
            {
                "Nota Fiscal": nf,
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
                "Valor Recebido NF": valor_recebido_nf,
                "% Recebido": perc_recebido,
                "Comissão Prevista Vendedor": comissao_prev_vendedor,
                "Comissão Liberada Vendedor": comissao_lib_vendedor,
                "Comissão Prevista Gerente": comissao_prev_gerente,
                "Comissão Liberada Gerente": comissao_lib_gerente,
                "Status Comissão": status,
                "Observação": observacao,
            }
        )

    saida = pd.DataFrame(linhas_saida)

    resumo = saida.groupby(["Vendedor", "Perfil Vendedor", "Gerente", "Status Comissão"], dropna=False).agg(
        Valor_Faturado=("Valor Faturado", "sum"),
        Valor_Recebido=("Valor Recebido NF", "sum"),
        Comissao_Prevista_Vendedor=("Comissão Prevista Vendedor", "sum"),
        Comissao_Liberada_Vendedor=("Comissão Liberada Vendedor", "sum"),
        Comissao_Prevista_Gerente=("Comissão Prevista Gerente", "sum"),
        Comissao_Liberada_Gerente=("Comissão Liberada Gerente", "sum"),
        Qtde_Notas=("Nota Fiscal", "nunique"),
    ).reset_index()

    return saida, resumo


def salvar_excel(saida: pd.DataFrame, resumo: pd.DataFrame, caminho_saida: Path) -> None:
    """Salva Excel com formatação leve e rápida."""
    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        saida.to_excel(writer, sheet_name="Comissoes_Calculadas", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1})
        money_fmt = workbook.add_format({"num_format": 'R$ #,##0.00'})
        perc_fmt = workbook.add_format({"num_format": '0.00%'})
        text_fmt = workbook.add_format({"text_wrap": False})

        formatos_moeda = {
            "Preço Praticado", "Preço Tabela", "Valor Faturado", "Valor Recebido NF",
            "Comissão Prevista Vendedor", "Comissão Liberada Vendedor",
            "Comissão Prevista Gerente", "Comissão Liberada Gerente",
            "Valor_Faturado", "Valor_Recebido", "Comissao_Prevista_Vendedor",
            "Comissao_Liberada_Vendedor", "Comissao_Prevista_Gerente", "Comissao_Liberada_Gerente"
        }
        formatos_perc = {"Variação %", "% Comissão Vendedor", "% Comissão Gerente", "% Recebido"}

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
    print("Comissão prevista vendedor:", f"R$ {saida['Comissão Prevista Vendedor'].sum():,.2f}")
    print("Comissão liberada vendedor:", f"R$ {saida['Comissão Liberada Vendedor'].sum():,.2f}")
    print("Comissão prevista gerente:", f"R$ {saida['Comissão Prevista Gerente'].sum():,.2f}")
    print("Comissão liberada gerente:", f"R$ {saida['Comissão Liberada Gerente'].sum():,.2f}")


if __name__ == "__main__":
    main()
