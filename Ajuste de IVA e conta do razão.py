import win32com.client
import pandas as pd
from datetime import datetime

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\python_scripts\Planilhas\Pedidos.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\Planilhas\Pedidos_log.xlsx"

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

df["Item"] = df["Item"].apply(lambda x: str(int(x)).zfill(5))
df["Modo"] = df.get("Modo", "AMBOS").astype(str).str.upper().str.strip()

log = []

# =========================
# LOOP
# =========================
for index, row in df.iterrows():
    try:
        pedido = str(row["Pedidos"])
        item = str(row["Item"])
        conta = str(row.get("Nova Conta Razão", "")).strip()
        iva = str(row.get("Novo Código Imposto", "")).strip()
        modo = row["Modo"]

        # =========================
        # ENTRA NA ME22N
        # =========================
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME22N"
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABRE PEDIDO
        # =========================
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById(
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN"
        ).text = pedido
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # =========================
        # SELECIONA ITEM
        # =========================
        session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/"
            "subSUB3:SAPLMEVIEWS:1100/"
            "subSUB1:SAPLMEVIEWS:4002/"
            "btnDYN_4000-BUTTON"
        ).press()

        # =========================
        # IVA
        # =========================
        if modo in ("IVA", "AMBOS") and iva:

            session.findById(
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/"
                "subSUB3:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/"
                "subSUB1:SAPLMEGUI:1301/"
                "subSUB2:SAPLMEGUI:1303/"
                "tabsITEM_DETAIL/tabpTABIDT9"
            ).select()

            campo_iva = session.findById(
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
                "subSUB3:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/"
                "subSUB1:SAPLMEGUI:1301/"
                "subSUB2:SAPLMEGUI:1303/"
                "tabsITEM_DETAIL/tabpTABIDT9/"
                "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/"
                "ctxtMEPO1317-MWSKZ"
            )

            campo_iva.text = iva
            session.findById("wnd[0]").sendVKey(0)

        # =========================
        # CONTA RAZÃO
        # =========================
        if modo in ("RAZAO", "AMBOS") and conta:

            session.findById(
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
                "subSUB3:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/"
                "subSUB1:SAPLMEGUI:1301/"
                "subSUB2:SAPLMEGUI:1303/"
                "tabsITEM_DETAIL/tabpTABIDT16"
            ).select()

            campo_conta = session.findById(
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/"
                "subSUB3:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/"
                "subSUB1:SAPLMEGUI:1301/"
                "subSUB2:SAPLMEGUI:1303/"
                "tabsITEM_DETAIL/tabpTABIDT16/"
                "ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/"
                "subSUB2:SAPLMEACCTVI:0100/"
                "subSUB1:SAPLMEACCTVI:1100/"
                "ctxtMEACCT1100-SAKTO"
            )

            campo_conta.text = conta
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
            "pedido": pedido,
            "item": item,
            "modo": modo,
            "conta": conta,
            "iva": iva,
            "status": "SUCESSO",
            "mensagem": status
        })

    except Exception as e:
        print(f"Erro na linha {index + 2}: {str(e)}")

        log.append({
            "linha": index + 2,
            "pedido": pedido,
            "item": item,
            "status": "ERRO",
            "mensagem": str(e)
        })

        break

# =========================
# LOG
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada.")