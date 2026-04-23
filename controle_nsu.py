import json
from zeep import Client, Transport
from zeep.plugins import HistoryPlugin
from requests import Session
from requests_pkcs12 import Pkcs12Adapter
from lxml import etree

# =========================
# CONFIGURAÇÕES
# =========================

WSDL_URL = "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"

CNPJ = "42816108000105"
UF = "MG"

CERT_PFX = r"C:\Users\ronaldo.gontijo\Downloads\7e0e25111057a8e4.pfx"
CERT_PASSWORD = "Pol964217"

NSU_CONTROLE_FILE = "nsu.json"

# =========================
# SOAP + CERTIFICADO
# =========================

session = Session()
session.mount(
    "https://",
    Pkcs12Adapter(
        pkcs12_filename=CERT_PFX,
        pkcs12_password=CERT_PASSWORD,
    ),
)

transport = Transport(session=session)
history = HistoryPlugin()

client = Client(
    wsdl=WSDL_URL,
    transport=transport,
    plugins=[history]
)

# =========================
# FUNÇÃO PRINCIPAL
# =========================

def alinhar_nsu():
    print("🔄 Consultando SEFAZ para alinhar NSU...")

    request_data = {
        "nfeDadosMsg": {
            "distDFeInt": {
                "@xmlns": "http://www.portalfiscal.inf.br/nfe",
                "tpAmb": "1",  # 1 = Produção | 2 = Homologação
                "cUFAutor": UF_COD(UF),
                "CNPJ": CNPJ,
                "distNSU": {
                    "ultNSU": "000000000000000"
                }
            }
        }
    }

    response = client.service.nfeDistDFeInteresse(**request_data)

    # =========================
    # LEITURA DO RETORNO
    # =========================

    ret = response["nfeDistDFeInteresseResult"]["retDistDFeInt"]

    cstat = ret["cStat"]
    motivo = ret["xMotivo"]

    print(f"✅ Retorno SEFAZ: {cstat} - {motivo}")

    if cstat not in ("138", "656"):
        raise Exception("Erro inesperado da SEFAZ")

    ult_nsu = ret["ultNSU"]
    max_nsu = ret["maxNSU"]

    print(f"📌 ultNSU SEFAZ: {ult_nsu}")
    print(f"📌 maxNSU SEFAZ: {max_nsu}")

    salvar_nsu(ult_nsu, max_nsu)

    print("✅ NSU alinhado e salvo com sucesso.")


# =========================
# FUNÇÕES AUXILIARES
# =========================

def salvar_nsu(ult_nsu, max_nsu):
    data = {
        "ultNSU": ult_nsu,
        "maxNSU": max_nsu
    }

    with open(NSU_CONTROLE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def UF_COD(uf):
    mapa = {
        "RO": "11", "AC": "12", "AM": "13", "RR": "14",
        "PA": "15", "AP": "16", "TO": "17",
        "MA": "21", "PI": "22", "CE": "23", "RN": "24",
        "PB": "25", "PE": "26", "AL": "27", "SE": "28",
        "BA": "29", "MG": "31", "ES": "32", "RJ": "33",
        "SP": "35", "PR": "41", "SC": "42", "RS": "43",
        "MS": "50", "MT": "51", "GO": "52", "DF": "53",
    }
    return mapa[uf]


# =========================
# EXECUÇÃO
# =========================

if __name__ == "__main__":
    alinhar_nsu()