"""
=========================================================
CONFIGURAÇÕES DA AUTOMAÇÃO ME31K
=========================================================
Altere apenas este arquivo caso seja necessário mudar
caminhos de arquivos ou configurações gerais.
=========================================================
"""

from pathlib import Path

# ======================================================
# PASTAS E DIRETÓRIOS
# ======================================================
BASE_DIR = Path(__file__).resolve().parent
PASTA_PLANILHAS = BASE_DIR / "Planilhas"
PASTA_LOGS = PASTA_PLANILHAS / "Logs"

# Garante a criação das pastas fundamentais (redundância com o logger)
PASTA_LOGS.mkdir(parents=True, exist_ok=True)

# ======================================================
# ARQUIVOS
# ======================================================
ARQUIVO_EXCEL = PASTA_PLANILHAS / "Contratos.xlsx"
ARQUIVO_LOG_EXCEL = PASTA_LOGS / "Contratos_ME31K_Log.xlsx"
ARQUIVO_LOG_TXT = PASTA_LOGS / "Contratos_ME31K_Log.txt"

# ======================================================
# SAP
# ======================================================
SAP_TRANSACAO = "ME31K"
FORMATO_DATA_SAP = "%d.%m.%Y"

# ======================================================
# COLUNAS ESPERADAS NO EXCEL
# ======================================================
COLUNAS_OBRIGATORIAS = [
    "ID_CONTRATO",
    "Fornecedor",
    "Tipo de contrato",
    "Organiz.compras",
    "Grupo de compradores",
    "Centro",
    "Fim da validade",
    "Condições de pagamento",
    "Incoterms",
    "Local Incoterms 1",
    "Material",
    "Qntde Prev",
    "valor",
    "Por",
    "Classificação Contabil",
    "Cód. Imposto",
    "Moeda"
]

# ======================================================
# CONFIGURAÇÕES GERAIS E INTERFACE
# ======================================================
CONTINUAR_EM_CASO_DE_ERRO = True
SALVAR_LOG_TXT = True
SALVAR_LOG_EXCEL = True

TITULO = "Automação SAP - Criação de Contratos ME31K"
VERSAO = "1.0.0"