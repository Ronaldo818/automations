import win32com.client
import pandas as pd
from datetime import datetime
import time

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\python_scripts\Planilhas\BPs.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\Planilhas\BPs_logs.xlsx"

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
df = pd.read_excel(ARQUIVO_ENTRADA, dtype={'BP': str})

log = []

# =========================
# FUNÇÕES
# =========================
def limpar_valor(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def formatar_bp(bp):
    bp = str(bp)
    if "." in bp:
        bp = bp.split(".")[0]
    return bp.zfill(10)

def salvar_log():
    pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

# =========================
# LOOP
# =========================
for index, row in df.iterrows():

    try:

        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        bp = formatar_bp(row["BP"])

        telefone = limpar_valor(row.get("TELEFONE"))
        telefone2 = limpar_valor(row.get("TELEFONE2"))
        email = limpar_valor(row.get("EMAIL"))

        # Ignora somente se não houver nenhum dado
        if not telefone and not telefone2 and not email:
            print(f"BP {bp} ignorado (sem dados)")
            continue

        print(f"\nProcessando BP {bp}...")

        # =========================
        # ABRIR BP
        # =========================
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nBP"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # =========================
        # BUSCAR BP
        # =========================
        campo_bp = session.findById(
            "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/"
            "subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/"
            "tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/"
            "ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/"
            "subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/"
            "subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2100/"
            "txtBUS_JOEL_SEARCH-PARTNER_NUMBER"
        )

        campo_bp.text = bp
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        grid = session.findById(
            "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/"
            "subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/"
            "tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/"
            "ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/"
            "subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/"
            "subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/"
            "ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/"
            "cntlSCREEN_1080_CONTAINER/shellcont/shell"
        )

        grid.selectedRows = "0"
        grid.doubleClickCurrentCell()
        time.sleep(1)

        # =========================
        # POPUP DE PERMISSÃO
        # =========================
        try:

            popup = session.findById("wnd[1]")

            popup.findById("tbar[0]/btn[0]").press()
            time.sleep(0.5)

            log.append({
                "linha": index + 2,
                "bp": bp,
                "telefone": telefone,
                "telefone2": telefone2,
                "email": email,
                "data_hora": data_hora,
                "status": "ERRO",
                "mensagem": "Sem permissão para alteração do BP"
            })

            salvar_log()

            print(f"BP {bp} sem permissão para alteração.")

            continue

        except:
            pass

        # =========================
        # MODO EDIÇÃO
        # =========================
        campo_tel = session.findById(
            "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/"
            "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
            "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
            "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
            "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/"
            "ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/"
            "ssubGENSUB:SAPLBUSS:7016/subA05P01:SAPLBUA0:0400/"
            "subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
            "txtSZA1_D0100-TEL_NUMBER"
        )

        if not campo_tel.Changeable:
            session.findById("wnd[0]/tbar[1]/btn[6]").press()
            time.sleep(1)

        alterou = False

        # =========================
        # TELEFONE / CELULAR
        # =========================
        if telefone:

            campo_tel.text = telefone

            celular = telefone2 if telefone2 else telefone

            session.findById(
                campo_tel.Id.replace("TEL_NUMBER", "MOB_NUMBER")
            ).text = celular

            alterou = True

        # =========================
        # EMAIL
        # =========================
        if email:

            session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/"
                "ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/"
                "ssubGENSUB:SAPLBUSS:7016/subA05P01:SAPLBUA0:0400/"
                "subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                "txtSZA1_D0100-SMTP_ADDR"
            ).text = email

            alterou = True

        # =========================
        # SALVAR
        # =========================
        if alterou:

            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)

            try:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                time.sleep(0.5)
            except:
                pass

            status = session.findById("wnd[0]/sbar").text

        else:
            status = "Nenhuma alteração necessária"

        log.append({
            "linha": index + 2,
            "bp": bp,
            "telefone": telefone,
            "telefone2": telefone2,
            "email": email,
            "data_hora": data_hora,
            "status": "SUCESSO",
            "mensagem": status
        })

        salvar_log()

        print(f"{bp} processado.")

    except Exception as e:

        print(f"Erro na linha {index + 2}: {str(e)}")

        log.append({
            "linha": index + 2,
            "bp": row.get("BP", ""),
            "telefone": row.get("TELEFONE", ""),
            "telefone2": row.get("TELEFONE2", ""),
            "email": row.get("EMAIL", ""),
            "data_hora": data_hora,
            "status": "ERRO",
            "mensagem": str(e)
        })

        salvar_log()
        break

print("Execução finalizada.")