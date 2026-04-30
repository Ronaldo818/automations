import os
import pandas as pd
from datetime import datetime
from pyrfc import Connection, ABAPApplicationError, ABAPRuntimeError, CommunicationError, LogonError

# ====== CONFIG ======
EXCEL_PATH = r"C:\Users\ronaldo.gontijo\Downloads\Pedidos.xlsx"
SHEET_NAME = 0
TESTRUN = False   # True = simulação | False = grava

SAP_CONN = dict(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.92",
    sysnr="00",
    client="300",
    lang="EN"  # recomendado EN
)

# ===== FUNÇÕES =====
def zfill_item(v):
    return str(int(v)).zfill(5)

def zfill_material(v):
    return str(v).strip().zfill(18)

def is_valid(val):
    return pd.notna(val) and str(val).strip() != "" and str(val).lower() != "nan"

def show_return(ret):
    msgs = []
    error = False
    for r in ret or []:
        line = f"{r['TYPE']} - {r['ID']} {r['NUMBER']}: {r['MESSAGE']}"
        msgs.append(line)
        if r["TYPE"] in ("E", "A"):
            error = True
    return error, msgs

# ===== BUSCA DADOS ATUAIS =====
def get_item_details(conn, po, item5):
    det = conn.call("BAPI_PO_GETDETAIL1", PURCHASEORDER=po)

    tax = None
    material = None
    cc = None

    for it in det.get("POITEM", []):
        if it["PO_ITEM"] == item5:
            tax = it.get("TAX_CODE")
            material = it.get("MATERIAL")

    for acc in det.get("POACCOUNT", []):
        if acc["PO_ITEM"] == item5:
            cc = acc.get("COSTCENTER")

    return tax, material, cc

# ===== LEITURA =====
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df.columns = [c.strip() for c in df.columns]

df["PO_ITEM"] = df["Item"].apply(zfill_item)
df["Pedidos"] = df["Pedidos"].fillna("").astype(str).str.strip()

df["Novo"] = df["Novo Código Imposto"].fillna("").astype(str).str.upper().str.strip()
df["Codigo_item"] = df["Codigo_item"].fillna("").astype(str).str.strip()
df["CC"] = df["CC"].fillna("").astype(str).str.strip()

# ===== CONEXÃO =====
conn = Connection(**SAP_CONN)
results = []

# ===== PROCESSAMENTO =====
for idx, row in df.iterrows():
    po = row["Pedidos"]
    item = row["PO_ITEM"]

    new_tax = row["Novo"]
    new_material = row["Codigo_item"]
    new_cc = row["CC"]

    before_tax, before_material, before_cc = get_item_details(conn, po, item)

    # ===== ESTRUTURAS DINÂMICAS =====
    poitem = [{"PO_ITEM": item}]
    poitemx = [{"PO_ITEM": item, "PO_ITEMX": "X"}]

    poaccount = []
    poaccountx = []

    alterou_algo = False

    # TAX
    if is_valid(new_tax):
        poitem[0]["TAX_CODE"] = new_tax
        poitemx[0]["TAX_CODE"] = "X"
        alterou_algo = True

    # MATERIAL
    if is_valid(new_material):
        material_formatado = zfill_material(new_material)
        poitem[0]["MATERIAL"] = material_formatado
        poitemx[0]["MATERIAL"] = "X"
        alterou_algo = True

    # CC
    if is_valid(new_cc):
        poaccount = [{
            "PO_ITEM": item,
            "SERIAL_NO": "01",
            "COSTCENTER": new_cc
        }]

        poaccountx = [{
            "PO_ITEM": item,
            "SERIAL_NO": "01",
            "PO_ITEMX": "X",
            "COSTCENTER": "X"
        }]

        alterou_algo = True

    # Se não tem nada pra alterar → pula
    if not alterou_algo:
        results.append({
            "Pedido": po,
            "Item": item,
            "Status": "IGNORADO",
            "Mensagens": "Linha sem dados para alteração"
        })
        continue

    try:
        params = dict(
            PURCHASEORDER=po,
            POITEM=poitem,
            POITEMX=poitemx
        )

        if poaccount:
            params["POACCOUNT"] = poaccount
            params["POACCOUNTX"] = poaccountx

        if TESTRUN:
            params["TESTRUN"] = "X"

        resp = conn.call("BAPI_PO_CHANGE", **params)
        has_error, msgs = show_return(resp.get("RETURN"))

        after_tax = None
        after_material = None
        after_cc = None

        if not TESTRUN and not has_error:
            conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
            after_tax, after_material, after_cc = get_item_details(conn, po, item)

        results.append({
            "Pedido": po,
            "Item": item,

            "Tax_Antes": before_tax,
            "Tax_Novo": new_tax,
            "Tax_Depois": after_tax,

            "Material_Antes": before_material,
            "Material_Novo": new_material,
            "Material_Depois": after_material,

            "CC_Antes": before_cc,
            "CC_Novo": new_cc,
            "CC_Depois": after_cc,

            "Status": "ERRO" if has_error else "OK",
            "Mensagens": " | ".join(msgs)
        })

    except Exception as e:
        results.append({
            "Pedido": po,
            "Item": item,
            "Status": "EXCEPTION",
            "Mensagens": str(e)
        })

# ===== RELATÓRIO =====
out = pd.DataFrame(results)
ts = datetime.now().strftime("%Y%m%d_%H%M%S")

out_path = os.path.join(
    os.path.dirname(EXCEL_PATH),
    f"resultado_items_{ts}.csv"
)

out.to_csv(out_path, index=False, encoding="utf-8")

print("\nRelatório salvo em:", out_path)
print(out.head(20))