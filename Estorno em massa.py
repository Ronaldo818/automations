import os
import pandas as pd
from datetime import datetime
from pyrfc import Connection

# =========================================================
# CONFIGURAÇÕES
# =========================================================
EXCEL_PATH = r"C:\Users\ronaldo.gontijo\Downloads\Estorno_MIGO.xlsx"
SHEET_NAME = 0
TESTRUN = True   # True = simula | False = grava no SAP

SAP_CONN = dict(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.10",
    sysnr="00",
    client="310",
    lang="PT"
)

# =========================================================
def zfill_item(v):
    return str(int(v)).zfill(5)

def show_return(ret):
    msgs, error = [], False
    for r in ret or []:
        msgs.append(f"{r['TYPE']} - {r['ID']} {r['NUMBER']}: {r['MESSAGE']}")
        if r["TYPE"] in ("E", "A"):
            error = True
    return error, msgs

def buscar_itens_mseg(conn, doc, year):
    res = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE="MSEG",
        DELIMITER=";",
        FIELDS=[
            {"FIELDNAME": "ZEILE"},
            {"FIELDNAME": "MENGE"},
            {"FIELDNAME": "MEINS"}
        ],
        OPTIONS=[{"TEXT": f"MBLNR = '{doc}' AND MJAHR = '{year}'"}]
    )

    itens = []
    for r in res["DATA"]:
        f = r["WA"].split(";")
        itens.append({
            "ITEM": zfill_item(f[0]),
            "QTD": float(f[1]),
            "UM": f[2]
        })
    return itens

# =========================================================
def main():

    conn = Connection(**SAP_CONN)
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df.columns = [c.strip().upper() for c in df.columns]

    resultados = []

    for _, row in df.iterrows():
        doc = str(row["DOCUMENTO"])
        year = str(row["ANO"])
        manter = [zfill_item(x) for x in str(row["ITENS_MANTER"]).split(",")]

        try:
            itens = buscar_itens_mseg(conn, doc, year)
            estornar = [i for i in itens if i["ITEM"] not in manter]

            if not estornar:
                raise Exception("Nenhum item para estornar")

            header = {
                "PSTNG_DATE": datetime.now().strftime("%Y%m%d"),
                "DOC_DATE": datetime.now().strftime("%Y%m%d"),
                "HEADER_TXT": f"ESTORNO AUTO {doc}"
            }

            code = {"GM_CODE": "03"}  # Estorno por referência

            items = []
            for i in estornar:
                items.append({
                    "MOVE_TYPE": "102",
                    "MVT_IND": "B",           # <<< CAMPO CRÍTICO
                    "REF_DOC": doc,
                    "REF_DOC_IT": i["ITEM"],
                    "ENTRY_QNT": i["QTD"],
                    "ENTRY_UOM": i["UM"]
                })

            params = dict(
                GOODSMVT_HEADER=header,
                GOODSMVT_CODE=code,
                GOODSMVT_ITEM=items
            )

            if TESTRUN:
                params["TESTRUN"] = "X"

            resp = conn.call("BAPI_GOODSMVT_CREATE", **params)
            has_error, msgs = show_return(resp.get("RETURN"))

            status = "SIMULADO" if TESTRUN else "OK"
            matdoc = ""

            if has_error:
                status = "ERRO"
            elif not TESTRUN:
                conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
                matdoc = resp.get("MATERIALDOCUMENT")

            resultados.append({
                "DOCUMENTO": doc,
                "STATUS": status,
                "DOC_GERADO": matdoc,
                "ITENS_ESTORNADOS": ",".join([i["ITEM"] for i in estornar]),
                "MENSAGENS": " | ".join(msgs)
            })

        except Exception as e:
            resultados.append({
                "DOCUMENTO": doc,
                "STATUS": "EXCEPTION",
                "MENSAGENS": str(e)
            })

    out = pd.DataFrame(resultados)
    out_path = os.path.join(
        os.path.dirname(EXCEL_PATH),
        f"log_estorno_migo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    out.to_excel(out_path, index=False)
    print("Log:", out_path)

if __name__ == "__main__":
    main()