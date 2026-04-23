import requests
from requests_pkcs12 import post
import xmltodict
import base64
import gzip
import os
import pandas as pd
import xml.etree.ElementTree as ET
import json
import html
import time

# ================= CONFIG =================
CNPJ = "42816108000105"
CERT_PATH = r"C:\Users\ronaldo.gontijo\Downloads\7e0e25111057a8e4.pfx"
CERT_PASSWORD = "Pol964217"

URL = "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"

ARQ_CONTROLE = "controle_nsu.json"
PASTA_XML = "xmls"

# ================= NSU =================
def carregar_nsu():
    if not os.path.exists(ARQ_CONTROLE):
        return "000000000000000"
    
    with open(ARQ_CONTROLE, "r") as f:
        return json.load(f)["ult_nsu"]

def salvar_nsu(nsu):
    with open(ARQ_CONTROLE, "w") as f:
        json.dump({"ult_nsu": nsu}, f)

# ================= XML SOAP =================
def montar_dist(ult_nsu):
    return f"""<distDFeInt xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.01">
<tpAmb>1</tpAmb>
<cUFAutor>31</cUFAutor>
<CNPJ>{CNPJ}</CNPJ>
<distNSU>
<ultNSU>{ult_nsu}</ultNSU>
</distNSU>
</distDFeInt>"""

def montar_xml(ult_nsu):
    dist = f'<distDFeInt xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.01"><tpAmb>1</tpAmb><cUFAutor>31</cUFAutor><CNPJ>{CNPJ}</CNPJ><distNSU><ultNSU>{ult_nsu}</ultNSU></distNSU></distDFeInt>'

    return f'''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
<soap:Body>
<nfeDistDFeInteresse xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe">
<nfeDadosMsg>{dist}</nfeDadosMsg>
</nfeDistDFeInteresse>
</soap:Body>
</soap:Envelope>'''

# ================= CONSULTA =================
def consultar_sefaz(ult_nsu):
    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe/nfeDistDFeInteresse"
    }

    response = post(
        URL,
        data=montar_xml(ult_nsu),
        headers=headers,
        pkcs12_filename=CERT_PATH,
        pkcs12_password=CERT_PASSWORD
    )

    return response.text

# ================= EXTRAÇÃO =================
import re

def extrair_docs(resposta):
    try:
        # limpa caracteres inválidos
        resposta_limpa = re.sub(r'[^\x09\x0A\x0D\x20-\x7F]+', '', resposta)

        dados = xmltodict.parse(resposta_limpa)

        retorno = dados['soap:Envelope']['soap:Body']['nfeDistDFeInteresseResponse']\
            ['nfeDistDFeInteresseResult']['retDistDFeInt']

        cstat = retorno.get("cStat")
        

        if cstat != "138":
            print(f"Consulta rejeitada: cStat {cstat} - {retorno.get('xMotivo')}")
            return retorno.get("ultNSU"), retorno.get("maxNSU"), []

        lote = retorno.get("loteDistDFeInt", {})
        docs = lote.get("docZip", [])

        if not isinstance(docs, list):
            docs = [docs]

        return retorno.get("ultNSU"), retorno.get("maxNSU"), docs

    except Exception as e:
        print("Erro ao interpretar XML:", e)
        print("Resposta bruta:")
        print(resposta[:1000])
        return None, None, []

# ================= SALVAR XML =================
def salvar_xmls(docs):
    os.makedirs(PASTA_XML, exist_ok=True)
    caminhos = []
 
    for doc in docs:
        conteudo = base64.b64decode(doc["#text"])
 
        try:
            xml = gzip.decompress(conteudo)
        except:
            xml = conteudo
 
        xml_str = xml.decode("utf-8", errors="ignore")
 
        # FILTRO PRINCIPAL
        if "<procNFe" not in xml_str:
            continue
 
        chave = doc.get("@NSU", str(len(caminhos)))
        caminho = os.path.join(PASTA_XML, f"{chave}.xml")
 
        with open(caminho, "wb") as f:
            f.write(xml)
 
        caminhos.append(caminho)
 
    return caminhos

# ================= PARSER =================
def ler_xml(caminho):
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
 
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()
 
        # garante que é NF válida
        if "procNFe" not in root.tag:
            return None
 
        return {
            "numero": root.find('.//nfe:nNF', ns).text,
            "cnpj_emitente": root.find('.//nfe:emit/nfe:CNPJ', ns).text,
            "razao_social": root.find('.//nfe:emit/nfe:xNome', ns).text,
            "valor": root.find('.//nfe:vNF', ns).text,
            "data": root.find('.//nfe:dhEmi', ns).text
        }
    except:
        return None

# ================= MAIN =================
def main():
    print("Iniciando consulta SEFAZ...")

    ult_nsu = carregar_nsu()

    todas_notas = []

    while True:
        resposta = consultar_sefaz(ult_nsu)

        novo_nsu, max_nsu, docs = extrair_docs(resposta)

        if not docs:
            print("Nenhum documento encontrado.")
            break

        caminhos = salvar_xmls(docs)

        for caminho in caminhos:
            dados = ler_xml(caminho)
            if dados:
                todas_notas.append(dados)

        if novo_nsu:
            ult_nsu = novo_nsu
            salvar_nsu(ult_nsu)

        print(f"NSU atualizado: {ult_nsu}")

        if ult_nsu == max_nsu:
            print("Tudo sincronizado.")
            break

        time.sleep(2)  # evita bloqueio SEFAZ

    if todas_notas:
        df = pd.DataFrame(todas_notas)
        df.to_excel("notas.xlsx", index=False)
        print("Planilha gerada!")
    else:
        print("Nenhuma nota válida encontrada.")

# ================= EXEC =================
if __name__ == "__main__":
    main()