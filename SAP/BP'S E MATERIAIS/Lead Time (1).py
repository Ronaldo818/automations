import os
import pandas as pd
import time
from datetime import datetime
from pyrfc import Connection, ABAPApplicationError, ABAPRuntimeError

# ===== CONFIG =====
EXCEL_PATH = r"C:\Users\ronaldo.gontijo\Downloads\Lead Time.xlsx"
SHEET_NAME = 0
TESTRUN = False

SAP_CONN = dict(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.92",
    sysnr="00",
    client="300",
    lang="PT"
)

# =========================
# HELPERS
# =========================

def zfill_matnr(v):
    try:
        return str(int(float(v))).zfill(18)
    except:
        return str(v).strip()

def norm_werks(v):
    return str(v).strip().zfill(4)

def to_int_or_none(v):
    if pd.isna(v):
        return None
    try:
        return int(float(str(v).replace(",", ".")))
    except:
        return None

def show_return(ret):
    if not ret:
        return False, []

    if isinstance(ret, dict):
        ret = [ret]

    msgs = []
    error = False

    for r in ret:
        msg = r.get("MESSAGE", "")
        msgs.append(msg)
        if r.get("TYPE") in ("E", "A"):
            error = True

    return error, msgs

def progress_bar(current, total, start):
    percent = current / total
    bar_len = 25
    filled = int(bar_len * percent)
    bar = "█" * filled + "-" * (bar_len - filled)

    elapsed = time.time() - start
    eta = int((elapsed / current) * (total - current)) if current else 0

    print(
        f"\r[{bar}] {current}/{total} {percent*100:.1f}% ETA:{eta}s",
        end=""
    )

# =========================
# LEITURA MARC (MATERIAL + CENTRO)
# =========================

def get_marc_bulk_safe(conn, mat_werks_list):

    result = {}

    for matnr, werks in mat_werks_list:
        try:
            resp = conn.call(
                "RFC_READ_TABLE",
                QUERY_TABLE="MARC",
                DELIMITER="|",
                FIELDS=[
                    {"FIELDNAME": "PLIFZ"},
                    {"FIELDNAME": "WEBAZ"},
                    {"FIELDNAME": "DZEIT"},
                    {"FIELDNAME": "DISMM"}
                ],
                OPTIONS=[
                    {"TEXT": f"MATNR = '{matnr}'"},
                    {"TEXT": f"AND WERKS = '{werks}'"}
                ],
                ROWCOUNT=1
            )

            data = resp.get("DATA", [])

            if data:
                wa = data[0]["WA"].split("|")
                result[(matnr, werks)] = {
                    "PLIFZ": int(wa[0].strip()) if wa[0].strip() else None,
                    "WEBAZ": int(wa[1].strip()) if wa[1].strip() else None,
                    "DZEIT": int(wa[2].strip()) if wa[2].strip() else None,
                    "DISMM": wa[3].strip()
                }

        except Exception as e:
            print(f"ERRO MARC | MATNR={matnr} WERKS={werks} | {e}")

    return result

# =========================
# MAIN
# =========================

def main():

    print("=== Atualização Lead Time SAP ===")

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df.columns = [c.strip() for c in df.columns]

    df["MATNR18"] = df["Material"].apply(zfill_matnr)
    df["WERKS"] = df["Centro"].apply(norm_werks)
    df["PLIFZ_NEW"] = df["PLIFZ"].apply(to_int_or_none)

    conn = Connection(**SAP_CONN)

    # LISTA ÚNICA DE (MATNR, WERKS)
    mat_werks_list = list(
        set(zip(df["MATNR18"], df["WERKS"]))
    )

    print("Carregando MARC...")
    marc_data = get_marc_bulk_safe(conn, mat_werks_list)

    df["BEFORE"] = df.apply(
        lambda r: marc_data.get((r["MATNR18"], r["WERKS"])),
        axis=1
    )

    results = []
    start = time.time()
    total = len(df)

    for idx, row in df.iterrows():

        progress_bar(idx + 1, total, start)

        matnr = row["MATNR18"]
        werks = row["WERKS"]
        new_plifz = row["PLIFZ_NEW"]
        before = row["BEFORE"]

        now = datetime.now()

        base_log = {
            "Data": now.strftime("%Y-%m-%d"),
            "Hora": now.strftime("%H:%M:%S"),
            "Material": matnr,
            "Centro": werks,
            "PLIFZ_ANTES": before["PLIFZ"] if before else None,
            "PLIFZ_DEPOIS": new_plifz
        }

        # SEM MARC
        if before is None:
            base_log["Status"] = "SEM_MRP"
            results.append(base_log)
            continue

        # SEM MRP ATIVO
        if not before["DISMM"]:
            base_log["Status"] = "SEM_MRP_ATIVO"
            results.append(base_log)
            continue

        # SEM ALTERAÇÃO
        if before["PLIFZ"] == new_plifz:
            base_log["Status"] = "SEM_ALTERACAO"
            results.append(base_log)
            continue

        try:
            resp = conn.call(
                "BAPI_MATERIAL_SAVEDATA",
                HEADDATA={"MATERIAL": matnr},
                PLANTDATA={
                    "PLANT": werks,
                    "PLND_DELRY": str(new_plifz)
                },
                PLANTDATAX={
                    "PLANT": werks,
                    "PLND_DELRY": "X"
                }
            )

            has_error, msgs = show_return(resp.get("RETURN"))
            msg_text = " | ".join(msgs)

            base_log["Status"] = "ERRO" if has_error else "OK"
            base_log["Mensagem"] = msg_text

            results.append(base_log)

        except (ABAPApplicationError, ABAPRuntimeError):
            base_log["Status"] = "BLOQUEADO"
            results.append(base_log)

        except Exception:
            base_log["Status"] = "EXCEPTION"
            results.append(base_log)

    # COMMIT FINAL
    if not TESTRUN:
        print("\nCommit final...")
        conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")

    return results

# =========================
# EXECUÇÃO
# =========================

if __name__ == "__main__":

    results = main()
    df = pd.DataFrame(results)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    path = os.path.join(
        os.path.dirname(EXCEL_PATH),
        f"resultado_leadtime_{ts}.xlsx"
    )

    df.to_excel(path, index=False)
    print("\nArquivo Excel salvo:", path)