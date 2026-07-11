import pandas as pd
import win32com.client
import time
from datetime import datetime

# =============================
# CONFIG
# =============================
CAMINHO_ARQUIVO = r"C:\python_scripts\Planilhas\Fornecedores.xlsx"
CAMINHO_LOG = r"C:\python_scripts\Planilhas\Fornecedores_IRF_logs.xlsx"

CODIGO_ANTIGO = "FA"
CODIGO_NOVO = "YR"
IRF_NOVO = "R0"

# =============================
# SAP CONNECTION
# =============================
sap = win32com.client.GetObject("SAPGUI")
application = sap.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# =============================
# IDs
# =============================
CAMPO_PN = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2100/txtBUS_JOEL_SEARCH-PARTNER_NUMBER"

GRID_RESULTADO = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/cntlSCREEN_1080_CONTAINER/shellcont/shell"

CAMPO_ROLE = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/subSCREEN_1100_ROLE_AND_TIME_AREA:SAPLBUPA_DIALOG_JOEL:1110/cmbBUS_JOEL_MAIN-PARTNER_ROLE"

ABA_EMPRESA = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_05"

TABELA_IRF = "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_05/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLCVI_FS_UI_VENDOR_CC:0054/tblSAPLCVI_FS_UI_VENDOR_CCTCTRL_LFBW"

# =============================
# AUXILIARES
# =============================
def esta_em_edicao():
    try:
        campo = session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WITHT[0,0]")
        return campo.Changeable
    except:
        return False

def garantir_edicao():
    if esta_em_edicao():
        print("Já está em modo edição")
        return

    print("Entrando em modo edição")
    session.findById("wnd[0]/tbar[1]/btn[6]").press()
    time.sleep(1)

    # tratar popup
    try:
        if session.Children.Count > 1:
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(1)
    except:
        pass

# =============================
# FUNÇÃO PRINCIPAL
# =============================
def atualizar_irf(lifnr):
    try:
        session.findById("wnd[0]").maximize()

        session.findById("wnd[0]/tbar[0]/okcd").text = "bp"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        session.findById(CAMPO_PN).text = lifnr
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        grid = session.findById(GRID_RESULTADO)
        grid.selectedRows = "0"
        grid.doubleClickCurrentCell()
        time.sleep(2)

        # role
        campo_role = session.findById(CAMPO_ROLE)
        if campo_role.key != "FLVN00":
            campo_role.key = "FLVN00"
            session.findById("wnd[0]/tbar[1]/btn[26]").press()
            time.sleep(2)

        session.findById(ABA_EMPRESA).select()
        time.sleep(2)

        # 🔥 GARANTIR EDIÇÃO (SEM F6)
        garantir_edicao()

        inserido = False
        linha_vazia = None

        for linha in range(0, 10):
            try:
                campo = session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WITHT[0,{linha}]")
                valor = campo.text.strip().upper()

                print(f"Linha {linha}: {valor}")

                if valor == CODIGO_NOVO:
                    inserido = True
                    break

                if valor == CODIGO_ANTIGO:
                    campo.text = CODIGO_NOVO

                    session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WT_WITHCD[2,{linha}]").text = IRF_NOVO
                    session.findById(f"{TABELA_IRF}/chkCVIS_LFBW-WT_SUBJCT[3,{linha}]").selected = True

                    inserido = True
                    break

                if valor == "" and linha_vazia is None:
                    linha_vazia = linha

            except:
                continue

        if not inserido:
            linha = linha_vazia if linha_vazia is not None else 0

            session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WITHT[0,{linha}]").text = CODIGO_NOVO
            session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WT_WITHCD[2,{linha}]").text = IRF_NOVO
            session.findById(f"{TABELA_IRF}/chkCVIS_LFBW-WT_SUBJCT[3,{linha}]").selected = True

        time.sleep(1)

        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(2)

        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        return "OK"

    except Exception as e:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        return str(e)

# =============================
# EXECUÇÃO
# =============================
df = pd.read_excel(CAMINHO_ARQUIVO)

for _, row in df.iterrows():
    lifnr = str(row["LIFNR"]).zfill(10)
    print(f"\nProcessando {lifnr}")
    print(atualizar_irf(lifnr))

print("Finalizado")