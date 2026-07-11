import pandas as pd
import win32com.client
import time
from datetime import datetime

# =============================
# CONFIG
# =============================
CAMINHO_ARQUIVO = r"C:\Users\ronaldo.gontijo\Downloads\Fornecedores.xlsx"
CAMINHO_LOG = r"C:\Users\ronaldo.gontijo\Downloads\Fornecedores_IRF_logs.xlsx"
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

def garantir_modo_edicao():
    campo_id = f"{TABELA_IRF}/ctxtCVIS_LFBW-WITHT[0,0]"

    try:
        campo = session.findById(campo_id)

        # já está editável
        if campo.Changeable:
            return True

    except:
        pass

    # tenta F6
    session.findById("wnd[0]").sendVKey(6)
    time.sleep(1)

    try:
        campo = session.findById(campo_id)

        # se entrou em edição
        if campo.Changeable:
            return True

        # ⚠️ se NÃO entrou → NÃO tenta mais
        return False

    except:
        return False

def capturar_mensagem_sap():
    mensagem = ""
    tipo = ""

    try:
        sbar = session.findById("wnd[0]/sbar")
        mensagem = sbar.text
        tipo = sbar.MessageType
    except:
        pass

    return mensagem, tipo

def tratar_popup_sap():
    mensagem_popup = ""
    tipo = "I"

    try:
        if session.Children.Count > 1:
            popup = session.findById("wnd[1]")

            mensagem_popup = popup.findById("usr/txtMESSTXT1").text

            popup.findById("tbar[0]/btn[0]").press()
            time.sleep(1)

            # popup geralmente é erro ou aviso
            tipo = "E"

    except:
        pass

    return mensagem_popup, tipo


# =============================
# FUNÇÃO PRINCIPAL
# =============================
def atualizar_irf(lifnr):
    inicio = datetime.now()

    try:
        session.findById("wnd[0]").maximize()

        # abrir BP
        session.findById("wnd[0]/tbar[0]/okcd").text = "bp"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        # buscar PN
        session.findById(CAMPO_PN).text = lifnr
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        # entrar no BP (double click)
        grid = session.findById(GRID_RESULTADO)
        grid.selectedRows = "0"
        grid.doubleClickCurrentCell()
        time.sleep(2)
    
        # validar role
        campo_role = session.findById(CAMPO_ROLE)
        if campo_role.key != "FLVN00":
            campo_role.key = "FLVN00"
            session.findById("wnd[0]/tbar[1]/btn[26]").press()
            time.sleep(2)

        # aba empresa
        session.findById(ABA_EMPRESA).select()
        time.sleep(1)

        session.findById("wnd[0]/tbar[1]/btn[6]").press()
        time.sleep(1)
        if not garantir_modo_edicao():
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)

            return {
                "LIFNR": lifnr,
                "STATUS": "ERRO",
                "ERRO": "Não entrou em modo edição",
                "DATA_HORA": inicio.strftime("%Y-%m-%d %H:%M:%S")
            }

        # =============================
        # INSERÇÃO
        # =============================
        inserido = False
        linha_vazia = None

        # =============================
        # PASSO 1: VARREDURA SEGURA
        # =============================
        for linha in range(0, 10):
            try:
                campo_tipo = session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WITHT[0,{linha}]")
                valor = campo_tipo.text.strip()

                # já existe FA
                if valor == "FA":
                    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                    session.findById("wnd[0]").sendVKey(0)
                    time.sleep(1)

                    return {
                        "LIFNR": lifnr,
                        "STATUS": "JA_EXISTE",
                        "ERRO": "",
                        "DATA_HORA": inicio.strftime("%Y-%m-%d %H:%M:%S")
                    }

                # guarda primeira linha vazia
                if valor == "" and linha_vazia is None:
                    linha_vazia = linha

            except:
                continue

        # =============================
        # PASSO 2: USAR LINHA VAZIA
        # =============================
        if linha_vazia is not None:
            linha = linha_vazia

            session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WITHT[0,{linha}]").text = "FA"
            session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WT_WITHCD[2,{linha}]").text = "F0"
            session.findById(f"{TABELA_IRF}/chkCVIS_LFBW-WT_SUBJCT[3,{linha}]").selected = True

            inserido = True

        # =============================
        # PASSO 3: CRIAR NOVA LINHA
        # =============================
        if not inserido:
            try:
                # botão inserir linha (baseado no seu record)
                session.findById("wnd[0]/tbar[1]/btn[6]").press()
                time.sleep(1)

                linha = 0  # nova linha geralmente aparece no topo

                session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WITHT[0,{linha}]").text = "FA"
                session.findById(f"{TABELA_IRF}/ctxtCVIS_LFBW-WT_WITHCD[2,{linha}]").text = "F0"
                session.findById(f"{TABELA_IRF}/chkCVIS_LFBW-WT_SUBJCT[3,{linha}]").selected = True

                inserido = True

            except:
                raise Exception("Não foi possível inserir nova linha IRF")

        if not inserido:
            raise Exception("Nenhuma linha disponível e não foi possível criar nova")

        time.sleep(1)

        # salvar
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(2)
        
        msg_popup, tipo_popup = tratar_popup_sap()
        msg_status, tipo_status = capturar_mensagem_sap()

        mensagem_final = msg_popup if msg_popup else msg_status
        tipo_final = tipo_popup if msg_popup else tipo_status
        
        if tipo_final == "E":
            status = "ERRO"
        elif tipo_final == "W":
            status = "AVISO"
        else:
            status = "SUCESSO"

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
        # reset em erro
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

    resultado = atualizar_irf(lifnr)
    logs.append(resultado)

pd.DataFrame(logs).to_excel(CAMINHO_LOG, index=False)

print("Finalizado!")