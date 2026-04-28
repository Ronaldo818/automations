import os
import pandas as pd
from datetime import datetime
from pyrfc import Connection

# ===== CONFIG =====
EXCEL_PATH = r"C:\python_scripts\Planilhas\Pedidos.xlsx"
TESTRUN = False

SAP_CONN = dict(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.10",
    sysnr="00",
    client="310",
    lang="PT"
)

def zfill_item(v):
    return str(int(v)).zfill(5)

def bdc_field(fnam, fval):
    return {"FNAM": fnam, "FVAL": fval}

def bdc_dynpro(program, dynpro):
    return {
        "PROGRAM": program,
        "DYNPRO": dynpro,
        "DYNBEGIN": "X"
    }

# ===== MONTA BDC =====
def montar_bdc_me22n(po, item, nova_conta, novo_iva):
    bdc = []

    # Tela inicial
    bdc.append(bdc_dynpro("SAPLMEGUI", "0014"))
    bdc.append(bdc_field("BDC_OKCODE", "/00"))
    bdc.append(bdc_field("MEPO_SELECT-EBELN", po))

    # Entrar no item
    bdc.append(bdc_dynpro("SAPLMEGUI", "0014"))
    bdc.append(bdc_field("BDC_OKCODE", "=ME22N_ITEM"))
    bdc.append(bdc_field("MEPO1211-EBELP", item))

    # Aba imputação
    bdc.append(bdc_dynpro("SAPLMEGUI", "0014"))
    bdc.append(bdc_field("BDC_OKCODE", "=KONT"))

    # Alterar conta (SAKTO)
    if nova_conta:
        bdc.append(bdc_field("MEPO1211-SAKTO", nova_conta))

    # Alterar IVA
    if novo_iva:
        bdc.append(bdc_field("MEPO1211-MWSKZ", novo_iva))

    # Salvar
    bdc.append(bdc_field("BDC_OKCODE", "=BU"))

    return bdc

# ===== LEITURA =====
df = pd.read_excel(EXCEL_PATH)
df.columns = [c.strip() for c in df.columns]

df["PO_ITEM"] = df["Item"].apply(zfill_item)
df["Pedidos"] = df["Pedidos"].astype(str).str.strip()
df["NovaConta"] = df.get("Nova Conta Razão", "").astype(str).str.strip()
df["NovoIVA"] = df.get("Novo Código Imposto", "").astype(str).str.strip()

conn = Connection(**SAP_CONN)

results = []

# ===== PROCESSAMENTO =====
for _, row in df.iterrows():
    po = row["Pedidos"]
    item = row["PO_ITEM"]
    conta = row["NovaConta"]
    iva = row["NovoIVA"]

    try:
        bdcdata = montar_bdc_me22n(po, item, conta, iva)

        resp = conn.call(
            "RFC_CALL_TRANSACTION_USING",
            TCODE="ME22N",
            MODE="N",
            BT_DATA=bdcdata
        )

        results.append({
            "Pedido": po,
            "Item": item,
            "Conta": conta,
            "IVA": iva,
            "Status": "OK",
            "Mensagem": str(resp)
        })

    except Exception as e:
        results.append({
            "Pedido": po,
            "Item": item,
            "Conta": conta,
            "IVA": iva,
            "Status": "ERRO",
            "Mensagem": str(e)
        })

# ===== RELATÓRIO =====
out = pd.DataFrame(results)
ts = datetime.now().strftime("%Y%m%d_%H%M%S")
out_path = os.path.join(os.path.dirname(EXCEL_PATH), f"resultado_bdc_{ts}.csv")

out.to_csv(out_path, index=False, encoding="utf-8")

print("Finalizado:", out_path)