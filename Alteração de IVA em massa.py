import os
import pandas as pd
from datetime import datetime
from pyrfc import (
    Connection,
    ABAPApplicationError,
    ABAPRuntimeError,
    CommunicationError,
    LogonError
)

# =========================
# CONFIGURAÇÕES
# =========================
EXCEL_PATH = r"C:\python_scripts\Planilhas\Pedidos - Copia.xlsx"
SHEET_NAME = 0
TESTRUN = False  # True = simulação | False = grava no SAP

SAP_CONN = dict(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.92",
    sysnr="00",
    client="300",
    lang="PT"
)

# =========================
# FUNÇÕES AUXILIARES
# =========================
def zfill_item(value):
    """Formata item para 5 dígitos (ex: 10 → 00010)"""
    return str(int(value)).zfill(5)


def show_return(ret):
    """Interpreta retorno do SAP"""
    messages = []
    error = False

    for r in ret or []:
        line = f"{r['TYPE']} - {r['ID']} {r['NUMBER']}: {r['MESSAGE']}"
        messages.append(line)

        if r["TYPE"] in ("E", "A"):
            error = True

    return error, messages


def get_tax(conn, po, item5):
    """Busca código de imposto atual"""
    print("VALOR:", po)
    print("TIPO:", type(po))
    print("REPR:", repr(po))
    print("TAMANHO:", len(str(po)))
    print("-" * 30)

    detail = conn.call(
        "BAPI_PO_GETDETAIL1",
        PURCHASEORDER=po
    )

    for item in detail.get("POITEM", []):
        if item["PO_ITEM"] == item5:
            return item.get("TAX_CODE")

    return None


# =========================
# LEITURA DO EXCEL
# =========================
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

# Limpa nomes das colunas
df.columns = [col.strip() for col in df.columns]

# Normalizações
df["PO_ITEM"] = df["Item"].apply(zfill_item)
df["Novo"] = df["Novo Código Imposto"].astype(str).str.upper().str.strip()
df["Pedidos"] = df["Pedidos"].astype(str).str.strip()

# =========================
# CONEXÃO SAP
# =========================
conn = Connection(**SAP_CONN)

results = []

# =========================
# PROCESSAMENTO
# =========================
for idx, row in df.iterrows():
    po = row["Pedidos"]
    item = row["PO_ITEM"]
    new_tax = row["Novo"]

    before = get_tax(conn, po, item)

    poitem = [{
        "PO_ITEM": item,
        "TAX_CODE": new_tax
    }]

    poitemx = [{
        "PO_ITEM": item,
        "PO_ITEMX": "X",
        "TAX_CODE": "X"
    }]

    try:
        params = {
            "PURCHASEORDER": po,
            "POITEM": poitem,
            "POITEMX": poitemx
        }

        if TESTRUN:
            params["TESTRUN"] = "X"

        response = conn.call("BAPI_PO_CHANGE", **params)

        has_error, messages = show_return(response.get("RETURN"))

        after = None

        if not TESTRUN and not has_error:
            conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
            after = get_tax(conn, po, item)

        results.append({
            "Pedido": po,
            "Item": item,
            "Antes": before,
            "Novo": new_tax,
            "Depois": after,
            "Status": "ERRO" if has_error else "OK",
            "Mensagens": " | ".join(messages)
        })

    except Exception as e:
        results.append({
            "Pedido": po,
            "Item": item,
            "Antes": before,
            "Novo": new_tax,
            "Depois": None,
            "Status": "EXCEPTION",
            "Mensagens": str(e)
        })

# =========================
# SAÍDA / RELATÓRIO
# =========================
output_df = pd.DataFrame(results)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

output_path = os.path.join(
    os.path.dirname(EXCEL_PATH),
    f"resultado_items_{timestamp}.csv"
)

output_df.to_csv(output_path, index=False, encoding="utf-8")

print("\nRelatório salvo em:", output_path)
print(output_df.head(20))