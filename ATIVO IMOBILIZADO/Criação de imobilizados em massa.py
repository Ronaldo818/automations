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
# FUNÇÃO CRITÉRIO
# =========================
def formatar_criterio(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return ""
    return str(int(float(valor))).zfill(2)

def limpar_valor(valor):
    if pd.isna(valor):
        return ""
    
    # Se for número (float), remove .0
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

        ordem = str(row['Ordem'])

        vida = str(int(row['Vida']))

        data_dt = pd.to_datetime(row['Depreciação'])
        data_dep = data_dt.strftime("%d.%m.%Y")

        data_dt = pd.to_datetime(row['Depreciação_Fiscal'])
        data_fis = data_dt.strftime("%d.%m.%Y")

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
                campo.setFocus()
                campo.text = serie
                campo.caretPosition = len(serie)
            except:
                pass

        if inventario:
            try:
                campo = session.findById(
                    "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/"
                    "ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-INVNR"
                )
                campo.setFocus()
                campo.text = inventario
                campo.caretPosition = len(inventario)
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
        # ABA 3 (CRITÉRIOS)
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

        # =========================
        # CAPTURA IMOBILIZADO
        # =========================
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

        break  # evita loop infinito

# =========================
# LOG
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada com sucesso.")