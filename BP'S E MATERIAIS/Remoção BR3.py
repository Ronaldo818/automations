import pandas as pd
import win32com.client
import time
from datetime import datetime

# =============================
# CONFIG
# =============================
CAMINHO_ARQUIVO = r"C:\python_scripts\PLANILHAS\Fornecedores_BR3.xlsx"
CAMINHO_LOG = r"C:\python_scripts\PLANILHAS\Fornecedores_BR3_logs.xlsx"

# =============================
# SAP CONNECTION
# =============================
sap = win32com.client.GetObject("SAPGUI")
application = sap.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# =============================
# IDs (reaproveitados)
# =============================
CAMPO_PN = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2100/txtBUS_JOEL_SEARCH-PARTNER_NUMBER"

GRID_RESULTADO = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/cntlSCREEN_1080_CONTAINER/shellcont/shell"

ABA_TAX = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03"

TABELA_TAX = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA07P01:SAPLBUPA_BUTX_DIALOG:0100/tblSAPLBUPA_BUTX_DIALOGTCTRL_BPTAX"

BOTAO_DELETE = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA07P01:SAPLBUPA_BUTX_DIALOG:0100/btnBUPA_BUTX01-DELROW"

# =============================
# FUNÇÕES AUXILIARES
# =============================
def garantir_modo_edicao():
    try:
        campo = session.findById(f"{TABELA_TAX}/ctxtDFKKBPTAXNUM-TAXTYPE[0,0]")

        if campo.Changeable:
            return True
    except:
        pass

    # tenta entrar em edição
    session.findById("wnd[0]").sendVKey(6)
    time.sleep(1)

    try:
        campo = session.findById(f"{TABELA_TAX}/ctxtDFKKBPTAXNUM-TAXTYPE[0,0]")
        return campo.Changeable
    except:
        return False


def capturar_mensagem_sap():
    try:
        sbar = session.findById("wnd[0]/sbar")
        return sbar.text, sbar.MessageType
    except:
        return "", ""


def tratar_popup_sap():
    try:
        if session.Children.Count > 1:
            popup = session.findById("wnd[1]")
            msg = popup.findById("usr/txtMESSTXT1").text
            popup.findById("tbar[0]/btn[0]").press()
            return msg, "E"
    except:
        pass

    return "", ""


# =============================
# FUNÇÃO PRINCIPAL
# =============================
def remover_br3(lifnr):
    inicio = datetime.now()

    try:
        session.findById("wnd[0]").maximize()

        # abrir BP
        session.findById("wnd[0]/tbar[0]/okcd").text = "bp"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        # buscar
        session.findById(CAMPO_PN).text = lifnr
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        # entrar
        grid = session.findById(GRID_RESULTADO)
        grid.selectedRows = "0"
        grid.doubleClickCurrentCell()
        time.sleep(2)

        # aba TAX
        session.findById(ABA_TAX).select()
        time.sleep(1)

        # modo edição (CRÍTICO)
        if not garantir_modo_edicao():
            raise Exception("Não entrou em modo edição")

        encontrou = False

        # varredura
        for linha in range(0, 10):
            try:
                campo = session.findById(f"{TABELA_TAX}/ctxtDFKKBPTAXNUM-TAXTYPE[0,{linha}]")
                valor = campo.text.strip()

                if valor == "BR3":
                    session.findById(TABELA_TAX).getAbsoluteRow(linha).selected = True
                    session.findById(BOTAO_DELETE).press()
                    encontrou = True
                    break

            except:
                continue

        if not encontrou:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)

            return {
                "LIFNR": lifnr,
                "STATUS": "NAO_ENCONTRADO",
                "ERRO": "",
                "DATA_HORA": inicio.strftime("%Y-%m-%d %H:%M:%S")
            }

        # salvar
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(2)

        msg_popup, tipo_popup = tratar_popup_sap()
        msg_status, tipo_status = capturar_mensagem_sap()

        mensagem_final = msg_popup if msg_popup else msg_status

        # reset
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        return {
            "LIFNR": lifnr,
            "STATUS": "SUCESSO",
            "ERRO": mensagem_final,
            "DATA_HORA": inicio.strftime("%Y-%m-%d %H:%M:%S")
        }

    except Exception as e:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        return {
            "LIFNR": lifnr,
            "STATUS": "ERRO",
            "ERRO": str(e),
            "DATA_HORA": inicio.strftime("%Y-%m-%d %H:%M:%S")
        }


# =============================
# EXECUÇÃO
# =============================
df = pd.read_excel(CAMINHO_ARQUIVO)

logs = []

for _, row in df.iterrows():
    lifnr = str(row["LIFNR"]).zfill(10)
    print(f"Processando: {lifnr}")

    resultado = remover_br3(lifnr)
    logs.append(resultado)

pd.DataFrame(logs).to_excel(CAMINHO_LOG, index=False)

print("Finalizado!")