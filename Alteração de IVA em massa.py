import os
import pandas as pd
from datetime import datetime
from pyrfc import Connection, ABAPApplicationError, ABAPRuntimeError, CommunicationError, LogonError

# ====== CONFIG ======
EXCEL_PATH = r"C:\Users\ronaldo.gontijo\Downloads\Pedidos - Copia.xlsx"
SHEET_NAME = 0
TESTRUN = False   # teste → coloque False para gravar

SAP_CONN = dict(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.92",
    sysnr="00",
    client="300",
    lang="PT"
)

def zfill_item(v):
    return str(int(v)).zfill(5)

def show_return(ret):
    msgs = []
    error = False
    for r in ret or []:
        line = f"{r['TYPE']} - {r['ID']} {r['NUMBER']}: {r['MESSAGE']}"
        msgs.append(line)
        if r["TYPE"] in ("E", "A"):
            error = True
    return error, msgs

def get_tax(conn, po, item5):
    print("VALOR:", po)
    print("TIPO:", type(po))
    print("REPR:", repr(po))
    print("TAMANHO:", len(str(po)))
    print("-"*30)
    det = conn.call("BAPI_PO_GETDETAIL1", PURCHASEORDER=po)
    for it in det.get("POITEM", []):
        if it["PO_ITEM"] == item5:
            return it.get("TAX_CODE")
    return None

# ===== LEITURA =====
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df.columns = [c.strip() for c in df.columns]

df["PO_ITEM"] = df["Item"].apply(zfill_item)
df["Novo"]    = df["Novo Código Imposto"].astype(str).str.upper().str.strip()
df["Pedidos"] = df["Pedidos"].astype(str).str.strip()

conn = Connection(**SAP_CONN)
results = []

# ===== PROCESSA ITEM A ITEM =====
for idx, row in df.iterrows():
    po = row["Pedidos"]
    item = row["PO_ITEM"]
    new_tax = row["Novo"]

    before = get_tax(conn, po, item)

    poitem = [{"PO_ITEM": item, "TAX_CODE": new_tax}]
    poitemx = [{"PO_ITEM": item, "PO_ITEMX": "X", "TAX_CODE": "X"}]

    try:
        params = dict(
            PURCHASEORDER=po,
            POITEM=poitem,
            POITEMX=poitemx
        )
        if TESTRUN:
            params["TESTRUN"] = "X"

        resp = conn.call("BAPI_PO_CHANGE", **params)
        has_error, msgs = show_return(resp.get("RETURN"))

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
            "Mensagens": " | ".join(msgs)
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

# ===== RELATÓRIO =====
out = pd.DataFrame(results)
ts = datetime.now().strftime("%Y%m%d_%H%M%S")
out_path = os.path.join(os.path.dirname(EXCEL_PATH), f"resultado_items_{ts}.csv")
out.to_csv(out_path, index=False, encoding="utf-8")

print("\nRelatório salvo em:", out_path)
print(out.head(20))