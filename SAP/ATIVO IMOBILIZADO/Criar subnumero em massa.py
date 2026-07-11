import win32com.client
import pandas as pd
from datetime import datetime
import re

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\python_scripts\Planilhas\Imobilizados_AS11.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\Planilhas\Imobilizados_AS11_logs.xlsx"

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
        imobilizado = limpar_valor(row['Imobilizado'])
        descricao = str(row['Denominação'])

        criterio1 = formatar_criterio(row['Criterio_1'])
        criterio2 = formatar_criterio(row['Criterio_2'])

        vida = str(int(row['Vida']))

        # =========================
        # INÍCIO - AS11
        # =========================
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nAS11"
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # INFORMAR IMOBILIZADO
        # =========================
        session.findById("wnd[0]/usr/ctxtANLA-ANLN1").text = imobilizado
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 1 - DESCRIÇÃO
        # =========================
        if descricao:
            session.findById(
                "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/"
                "ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-TXT50"
            ).text = descricao

        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 3 - CRITÉRIOS
        # =========================
        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03"
        ).select()

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
        # ABA 8 - VIDA ÚTIL
        # =========================
        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08"
        ).select()

        if vida:
            session.findById(
                "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/"
                "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/"
                "tblSAPLAISTTC_ANLB/txtANLB-NDJAR[4,0]"
            ).text = vida

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

        log.append({
            "linha": index + 2,
            "imobilizado": imobilizado,
            "status": "SUCESSO",
            "mensagem": status
        })

    except Exception as e:
        print(f"Erro na linha {index + 2}: {str(e)}")

        log.append({
            "linha": index + 2,
            "imobilizado": row.get('Imobilizado', ''),
            "status": "ERRO",
            "mensagem": str(e)
        })

        break

# =========================
# LOG
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada com sucesso.")