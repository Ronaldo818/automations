from pyrfc import Connection

# ===========================================================
# CONEXÃO
# ===========================================================

conn = Connection(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.10",
    sysnr="00",
    client="300",
    lang="EN"
)

pedido = "4500031022"
item = "00010"

# ===========================================================
# LEITURA DO PEDIDO
# ===========================================================

print("=" * 80)
print("LENDO PEDIDO")
print("=" * 80)

det = conn.call(
    "BAPI_PO_GETDETAIL1",
    PURCHASEORDER=pedido,
    ACCOUNT_ASSIGNMENT="X"
)

# ===========================================================
# POACCOUNT
# ===========================================================

conta = None

for acc in det["POACCOUNT"]:
    if acc["PO_ITEM"] == item:
        conta = acc
        break

if conta is None:
    raise Exception("Item não possui imputação.")

print("\n================ POACCOUNT ================\n")

for k, v in conta.items():
    print(f"{k:25} {v}")

serial = conta["SERIAL_NO"]

print("\nSERIAL_NO encontrado:", serial)

# ===========================================================
# POITEM
# ===========================================================

dados_item = None

for it in det["POITEM"]:
    if it["PO_ITEM"] == item:
        dados_item = it
        break

if dados_item:

    print("\n================ POITEM ================\n")

    campos = [
        "PO_ITEM",
        "SHORT_TEXT",
        "MATERIAL",
        "PLANT",
        "STGE_LOC",
        "QUANTITY",
        "NET_PRICE",
        "ACCTASSCAT",
        "PREQ_NAME",
        "PREQ_NO",
        "PREQ_ITEM",
        "MATL_GROUP",
        "ITEM_CAT",
        "DELETE_IND"
    ]

    for campo in campos:
        print(f"{campo:20} {dados_item.get(campo)}")

# ===========================================================
# TESTE DA BAPI
# ===========================================================

print("\n")
print("=" * 80)
print("SIMULANDO REMOÇÃO DO CENTRO DE CUSTO")
print("=" * 80)

params = {
    "PURCHASEORDER": pedido,

    "POITEM": [{
        "PO_ITEM": item
    }],

    "POITEMX": [{
        "PO_ITEM": item,
        "PO_ITEMX": "X"
    }],

    "POACCOUNT": [{
        "PO_ITEM": item,
        "SERIAL_NO": serial,
        "COSTCENTER": ""
    }],

    "POACCOUNTX": [{
        "PO_ITEM": item,
        "SERIAL_NO": serial,
        "PO_ITEMX": "X",
        "COSTCENTER": "X"
    }],

    "TESTRUN": "X"
}

ret = conn.call(
    "BAPI_PO_CHANGE",
    **params
)

print("\n================ RETURN ================\n")

for r in ret["RETURN"]:
    print(
        f"{r['TYPE']} | "
        f"{r['ID']} {r['NUMBER']} | "
        f"{r['MESSAGE']}"
    )