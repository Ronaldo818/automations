import win32com.client
import pandas as pd
import time

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\python_scripts\Planilhas\Pedidos.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\Planilhas\Pedidos_logs2.xlsx"

# =========================
# FUNÇÕES AUXILIARES
# =========================
def wait_for_element(session, element_id, timeout=6):
    for _ in range(timeout * 10):
        try:
            return session.findById(element_id)
        except:
            time.sleep(0.1)
    raise Exception(f"Elemento não encontrado: {element_id}")

def acessar_aba(session, aba_id, tentativas=6):
    """
    Tenta acessar uma aba.
    Se não conseguir, tenta expandir o item e tenta novamente.
    """
    for _ in range(tentativas):
        try:
            session.findById(aba_id).select()
            return True
        except:
            # tenta expandir
            try:
                session.findById(
                    "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/"
                    "subSUB3:SAPLMEVIEWS:1100/"
                    "subSUB1:SAPLMEVIEWS:4002/"
                    "btnDYN_4000-BUTTON"
                ).press()
            except:
                pass

            time.sleep(0.5)

    raise Exception(f"Não conseguiu acessar aba")

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
    pedido = ""
    item = ""

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
        wait_for_element(session, "wnd[0]/tbar[1]/btn[17]").press()

        wait_for_element(
            session,
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN"
        ).text = pedido

        wait_for_element(session, "wnd[1]/tbar[0]/btn[0]").press()

        time.sleep(0.7)

        # =========================
        # IVA
        # =========================
        if modo in ("IVA", "AMBOS") and iva:

            acessar_aba(
                session,
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/"
                "subSUB3:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/"
                "subSUB1:SAPLMEGUI:1301/"
                "subSUB2:SAPLMEGUI:1303/"
                "tabsITEM_DETAIL/tabpTABIDT9"
            )

            time.sleep(0.4)

            campo_iva = wait_for_element(
                session,
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/"
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

            acessar_aba(
                session,
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/"
                "subSUB3:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/"
                "subSUB1:SAPLMEGUI:1301/"
                "subSUB2:SAPLMEGUI:1303/"
                "tabsITEM_DETAIL/tabpTABIDT16"
            )

            time.sleep(0.5)

            campo_conta = wait_for_element(
                session,
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
        wait_for_element(session, "wnd[0]/tbar[0]/btn[11]").press()

        try:
            wait_for_element(session, "wnd[1]/tbar[0]/btn[0]").press()
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

        continue

# =========================
# LOG
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada.")