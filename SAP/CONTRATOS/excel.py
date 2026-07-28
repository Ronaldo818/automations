"""
=========================================================
LEITURA E VALIDAÇÃO DO EXCEL
=========================================================
"""

import pandas as pd

from config import (
    ARQUIVO_EXCEL,
    COLUNAS_OBRIGATORIAS
)

from util import (
    validar_colunas,
    limpar_texto
)


# ======================================================
# CARREGAR PLANILHA E TRATAR ERROS COMUNS
# ======================================================
def carregar_excel():
    try:
        # Lê a planilha já tratando células vazias (NaN) como texto em branco
        df = pd.read_excel(ARQUIVO_EXCEL)
        df = df.fillna("")
        
    except PermissionError:
        raise Exception(
            f"A planilha está aberta pelo usuário!\n"
            f"Por favor, feche o arquivo '{ARQUIVO_EXCEL}' e tente novamente."
        )
    except FileNotFoundError:
        raise Exception(
            f"Arquivo não encontrado!\n"
            f"Certifique-se de que a planilha existe no caminho: '{ARQUIVO_EXCEL}'."
        )
    except Exception as e:
        raise Exception(f"Erro inesperado ao ler a planilha: {str(e)}")

    faltando = validar_colunas(df, COLUNAS_OBRIGATORIAS)
    if faltando:
        raise Exception(
            "Colunas obrigatórias não encontradas na planilha:\n\n"
            + "\n".join(faltando)
        )

    if df.empty:
        raise Exception("A planilha está vazia. Insira os dados e tente novamente.")

    return df


# ======================================================
# AGRUPAR CONTRATOS
# ======================================================
def agrupar_contratos(df):
    return df.groupby("ID_CONTRATO", sort=False)


# ======================================================
# VALIDAÇÃO DO CABEÇALHO
# ======================================================
def validar_grupo(grupo):
    """
    Garante que os dados de cabeçalho sejam idênticos 
    para todas as linhas de um mesmo contrato.
    """
    campos = [
        "Fornecedor",
        "Tipo de contrato",
        "Organiz.compras",
        "Grupo de compradores",
        "Centro",
        "Fim da validade",
        "Condições de pagamento",
        "Incoterms",
        "Local Incoterms 1",
        "Moeda"
    ]

    for campo in campos:
        # Como já usamos fillna(""), não precisamos nos preocupar com NaN aqui
        valores = []
        for valor in grupo[campo]:
            texto = limpar_texto(valor)
            if texto != "":
                valores.append(texto)

        # Remove duplicatas
        valores = list(set(valores))

        if len(valores) > 1:
            id_contrato = grupo.iloc[0]['ID_CONTRATO']
            raise Exception(
                f"Inconsistência no Contrato '{id_contrato}':\n"
                f"O campo '{campo}' possui mais de um valor diferente na planilha "
                f"(Valores encontrados: {valores}). Todos os itens do mesmo contrato "
                f"devem ter o mesmo cabeçalho."
            )


# ======================================================
# VALIDAR TODOS OS CONTRATOS
# ======================================================
def validar_planilha(df):
    for _, grupo in agrupar_contratos(df):
        validar_grupo(grupo)


# ======================================================
# ACESSO AOS DADOS DO CONTRATO
# ======================================================
def obter_cabecalho(grupo):
    return grupo.iloc[0]

def obter_itens(grupo):
    return grupo.reset_index(drop=True)


# ======================================================
# ESTATÍSTICAS E RESUMO
# ======================================================
def quantidade_contratos(df):
    return df["ID_CONTRATO"].nunique()

def quantidade_itens(df):
    return len(df)

def resumo(df):
    return {
        "contratos": quantidade_contratos(df),
        "itens": quantidade_itens(df)
    }