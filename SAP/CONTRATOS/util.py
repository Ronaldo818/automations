"""
=========================================================
FUNÇÕES AUXILIARES
=========================================================
Funções reutilizadas por toda a automação.
=========================================================
"""

import re
import pandas as pd
from datetime import datetime
from config import FORMATO_DATA_SAP


def limpar_texto(valor):
    """Remove espaços e trata células vazias."""
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def inteiro(valor):
    """Converte números do Excel (Ex: 1000.0 -> 1000)."""
    if pd.isna(valor):
        return ""
    try:
        return str(int(float(valor)))
    except ValueError:
        return str(valor).strip()


def decimal(valor, casas=2):
    """Formata valores monetários para o padrão do SAP (substitui . por ,)."""
    if pd.isna(valor):
        return ""
    try:
        numero = float(valor)
        texto = f"{numero:.{casas}f}"
        return texto.replace(".", ",")
    except ValueError:
        return str(valor)


def data_sap(valor):
    """Converte qualquer data para o padrão do SAP (dd.mm.aaaa)."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, datetime):
        return valor.strftime(FORMATO_DATA_SAP)
    try:
        data = pd.to_datetime(valor)
        return data.strftime(FORMATO_DATA_SAP)
    except Exception:
        return str(valor)


def possui_valor(valor):
    if pd.isna(valor):
        return False
    return str(valor).strip() != ""


def vazio(valor):
    return not possui_valor(valor)


def extrair_numero_contrato(status):
    """Extrai os 10 dígitos do número do contrato a partir da mensagem do SAP."""
    if not status:
        return ""
    match = re.search(r"\d{10}", status)
    if match:
        return match.group()
    return ""


def validar_colunas(df, colunas):
    return [coluna for coluna in colunas if coluna not in df.columns]


def iguais(a, b):
    return limpar_texto(a).upper() == limpar_texto(b).upper()