import win32com.client
import pandas as pd
from datetime import datetime
import re

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\python_scripts\Planilhas\Imobilizados_AS01.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\Planilhas\Imobilizados_AS01_logs.xlsx"

# =========================
# SAP
# =========================
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# =========================
# PLANILHA
# =========================
df = pd.read_excel(ARQUIVO_ENTRADA)

log = []

# =========================
# FUNÇÕES
# =========================
def formatar_criterio(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return ""
    return str(int(float(valor))).zfill(2)

def limpar_valor(valor):
    if pd.isna(valor):
        return ""
    if isinstance(valor, float):
        return str(int(valor))
    return str(valor).strip()

# =========================
# LOOP
# =========================
for index, row in df.iterrows():

    try:
        classe = str(row['Classe'])
        descricao = str(row['Denominação'])
        serie = limpar_valor(row['Serie'])
        inventario = limpar_valor(row['Inventario'])
        centro_custo = str(row['Centro de custo'])
        centro = str(row['Centro'])

        criterio1 = formatar_criterio(row['Criterio_1'])
        criterio2 = formatar_criterio(row['Criterio_2'])

        ordem = limpar_valor(row['Ordem'])

        # NOVOS CAMPOS
        imob_origem = limpar_valor(row['Imob Origem'])
        origem_sub = limpar_valor(row['Origem Sub'])

        vida = str(int(row['Vida']))

        data_dep = pd.to_datetime(row['Depreciação']).strftime("%d.%m.%Y")
        data_fis = pd.to_datetime(row['Depreciação_Fiscal']).strftime("%d.%m.%Y")

        # =========================
        # INÍCIO
        # =========================
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nAS01"
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # TELA INICIAL
        # =========================
        session.findById("wnd[0]/usr/ctxtANLA-ANLKL").text = classe
        session.findById("wnd[0]/usr/ctxtANLA-BUKRS").text = "2000"
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 1
        # =========================
        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/"
            "ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-TXT50"
        ).text = descricao

        if serie:
            try:
                campo = session.findById(
                    "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/"
                    "ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-SERNR"
                )
                campo.text = serie
            except:
                pass

        if inventario:
            try:
                campo = session.findById(
                    "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/"
                    "ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-INVNR"
                )
                campo.text = inventario
            except:
                pass

        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 2
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select()

        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/"
            "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/ctxtANLZ-KOSTL"
        ).text = centro_custo

        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/"
            "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/ctxtANLZ-WERKS"
        ).text = centro

        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 3
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select()

        if criterio1:
            session.findById(
                "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/"
                "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1160/ctxtANLA-ORD41"
            ).text = criterio1

        if criterio2:
            session.findById(
                "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/"
                "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1160/ctxtANLA-ORD42"
            ).text = criterio2

        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 4 (ORIGEM)
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select()

        # AIBN1
        if imob_origem:
            try:
                session.findById(
                    "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/"
                    "ssubSUBSC:SAPLATAB:0202/subAREA1:SAPLAIST:1181/txtANLA-AIBN1"
                ).text = imob_origem
            except:
                pass

        # AIBN2
        if origem_sub:
            try:
                session.findById(
                    "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/"
                    "ssubSUBSC:SAPLATAB:0202/subAREA1:SAPLAIST:1181/txtANLA-AIBN2"
                ).text = origem_sub
            except:
                pass

        # EAUFN
        if ordem:
            session.findById(
                "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/"
                "ssubSUBSC:SAPLATAB:0202/subAREA2:SAPLAIST:1182/ctxtANLA-EAUFN"
            ).text = ordem

        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 8
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08").select()

        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/"
            "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/"
            "tblSAPLAISTTC_ANLB/txtANLB-NDJAR[4,0]"
        ).text = vida

        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/"
            "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/"
            "tblSAPLAISTTC_ANLB/ctxtANLB-AFABG[6,0]"
        ).text = data_dep

        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/"
            "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/"
            "tblSAPLAISTTC_ANLB/ctxtANLB-AFABG[6,2]"
        ).text = data_fis

        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # SALVAR
        # =========================
        session.findById("wnd[0]/tbar[0]/btn[11]").press()

        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        status = session.findById("wnd[0]/sbar").text

        match = re.search(r'\d+', status)
        imobilizado = match.group() if match else ""

        log.append({
            "linha": index + 2,
            "classe": classe,
            "descricao": descricao,
            "imobilizado": imobilizado,
            "status": "SUCESSO",
            "mensagem": status
        })

    except Exception as e:
        print(f"Erro na linha {index + 2}: {str(e)}")

        log.append({
            "linha": index + 2,
            "classe": row.get('Classe', ''),
            "descricao": row.get('Denominação', ''),
            "status": "ERRO",
            "mensagem": str(e)
        })

        break

# =========================
# LOG
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada com sucesso.")