"""
=========================================================
LEITURA E VALIDAÇÃO DO EXCEL
=========================================================
"""

import pandas as pd

from config import COLUNAS_OBRIGATORIAS
from util import validar_colunas, limpar_texto

# ======================================================
# CARREGAR PLANILHA E TRATAR ERROS COMUNS
# ======================================================
def carregar_excel(caminho_arquivo):
    """Agora recebe o caminho dinâmico escolhido pelo usuário."""
    try:
        df = pd.read_excel(caminho_arquivo)
        df = df.fillna("")
    except PermissionError:
        raise Exception(
            f"A planilha está aberta pelo usuário!\n"
            f"Por favor, feche o arquivo e tente novamente."
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
# VALIDAÇÃO DE CABEÇALHO (Consistência interna)
# ======================================================
def validar_grupo(grupo):
    campos = [
        "Fornecedor", "Tipo de contrato", "Organiz.compras",
        "Grupo de compradores", "Centro", "Fim da validade",
        "Condições de pagamento", "Incoterms", "Local Incoterms 1", "Moeda"
    ]

    for campo in campos:
        valores = []
        for valor in grupo[campo]:
            texto = limpar_texto(valor)
            if texto != "":
                valores.append(texto)

        valores = list(set(valores))
        if len(valores) > 1:
            id_contrato = grupo.iloc[0]['ID_CONTRATO']
            raise Exception(
                f"Inconsistência no Contrato '{id_contrato}':\n"
                f"O campo '{campo}' possui valores diferentes {valores} para o mesmo contrato."
            )

# ======================================================
# VALIDAÇÃO DE DADOS VAZIOS (Nova Trava)
# ======================================================
def validar_dados_preenchidos(df):
    """Impede que o robô inicie se houver campos vitais em branco."""
    
    # Colunas que não podem ficar vazias de jeito nenhum
    colunas_vitais = [
        "ID_CONTRATO", "Fornecedor", "Tipo de contrato", "Organiz.compras",
        "Grupo de compradores", "Fim da validade", "Condições de pagamento",
        "Material", "Qntde Prev", "Cód. Imposto"
    ]
    
    erros = []
    
    for indice, linha in df.iterrows():
        linha_excel = indice + 2 # +2 por causa do cabeçalho (linha 1) e índice base 0
        
        for col in colunas_vitais:
            if col in df.columns:
                valor = str(linha[col]).strip()
                if valor == "":
                    erros.append(f"Linha {linha_excel}: O campo '{col}' está vazio.")
                    
    if erros:
        # Mostra os 10 primeiros erros para não estourar a tela do usuário
        mensagem_erro = "Dados obrigatórios ausentes na planilha:\n\n" + "\n".join(erros[:10])
        if len(erros) > 10:
            mensagem_erro += f"\n...e mais {len(erros) - 10} erro(s)."
        raise Exception(mensagem_erro)

# ======================================================
# VALIDAR TODOS OS CONTRATOS
# ======================================================
def validar_planilha(df):
    validar_dados_preenchidos(df) # Trava 1: Células vazias
    for _, grupo in agrupar_contratos(df):
        validar_grupo(grupo)      # Trava 2: Divergência de cabeçalho

# ======================================================
# ACESSO AOS DADOS E RESUMO
# ======================================================
def obter_cabecalho(grupo): return grupo.iloc[0]
def obter_itens(grupo): return grupo.reset_index(drop=True)
def quantidade_contratos(df): return df["ID_CONTRATO"].nunique()
def quantidade_itens(df): return len(df)
def resumo(df):
    return {"contratos": quantidade_contratos(df), "itens": quantidade_itens(df)}