import os
import traceback
import pandas as pd
from datetime import datetime
from pyrfc import Connection, ABAPApplicationError, ABAPRuntimeError, CommunicationError, LogonError
 
# ====== CONFIG ======
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
    s = str(v).strip()
    if not s:
        return s
    if s.isdigit():
        return s.zfill(18)
    return s
 
 
def norm_werks(v):
    return str(v).strip().upper()
 
 
def to_int_or_none(v):
    if pd.isna(v) or v is None:
        return None
    try:
        return int(float(str(v).replace(",", ".")))
    except:
        return None
 
 
def sap_numc(v):
    """Converte número para formato aceito pelo SAP (NUMC/CHAR)."""
    if v is None:
        return None
    return str(int(v))
 
 
def show_return(ret):
 
    msgs = []
    error = False
 
    if not ret:
        return False, []
 
    # se for apenas uma estrutura
    if isinstance(ret, dict):
        ret = [ret]
 
    for r in ret:
 
        line = f"{r.get('TYPE','?')} - {r.get('ID','')} {r.get('NUMBER','')}: {r.get('MESSAGE','')}"
        msgs.append(line)
 
        if r.get("TYPE") in ("E", "A"):
            error = True
 
    return error, msgs
 
 
# =========================
# LER LEAD TIME ATUAL
# =========================
 
def get_lead_times(conn, matnr18, werks):
 
    try:
 
        fields = [
            {"FIELDNAME": "PLIFZ"},
            {"FIELDNAME": "WEBAZ"},
            {"FIELDNAME": "DZEIT"}
        ]
 
        options = [
            {"TEXT": f"MATNR = '{matnr18}' AND WERKS = '{werks}'"}
        ]
 
        resp = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="MARC",
            DELIMITER="|",
            FIELDS=fields,
            OPTIONS=options,
            ROWCOUNT=1
        )
 
        data = resp.get("DATA", [])
 
        if not data:
            return None
 
        wa = data[0]["WA"].split("|")
 
        return {
            "PLIFZ": int(wa[0].strip()) if wa[0].strip() else None,
            "WEBAZ": int(wa[1].strip()) if wa[1].strip() else None,
            "DZEIT": int(wa[2].strip()) if wa[2].strip() else None,
        }
 
    except Exception as e:
        print(f"[WARN] Falha ao ler MARC: {e}")
        return None
 
 
# =========================
# MAIN
# =========================
 
def main():
 
    print("=== Atualização de Lead Time ===")
 
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    except:
        print("Erro ao ler Excel")
        traceback.print_exc()
        return []
 
    df.columns = [c.strip() for c in df.columns]
 
    required_cols = {"Material", "Centro"}
 
    if not required_cols.issubset(df.columns):
        print("Colunas obrigatórias faltando")
        return []
 
    has_plifz = "PLIFZ" in df.columns
    has_webaz = "WEBAZ" in df.columns
    has_dzeit = "DZEIT" in df.columns
 
    df["MATNR18"] = df["Material"].apply(zfill_matnr)
    df["WERKS"] = df["Centro"].apply(norm_werks)
 
    df["PLIFZ_NEW"] = df["PLIFZ"].apply(to_int_or_none) if has_plifz else None
    df["WEBAZ_NEW"] = df["WEBAZ"].apply(to_int_or_none) if has_webaz else None
    df["DZEIT_NEW"] = df["DZEIT"].apply(to_int_or_none) if has_dzeit else None
 
    print("Abrindo conexão SAP...")
 
    conn = Connection(**SAP_CONN)
 
    print("Conectado.")
 
    results = []
 
    for idx, row in df.iterrows():
 
        matnr18 = row["MATNR18"]
        werks = row["WERKS"]
 
        new_plifz = row["PLIFZ_NEW"] if has_plifz else None
        new_webaz = row["WEBAZ_NEW"] if has_webaz else None
        new_dzeit = row["DZEIT_NEW"] if has_dzeit else None
 
        print(f"\n[{idx+1}/{len(df)}] {matnr18} / {werks}")
 
        before = get_lead_times(conn, matnr18, werks)
 
        print("Antes:", before)
 
        plantdata = {"PLANT": werks}
        plantdatax = {"PLANT": werks}
 
        any_field = False
 
        if new_plifz is not None:
 
            plantdata["PLND_DELRY"] = sap_numc(new_plifz)
            plantdatax["PLND_DELRY"] = "X"
 
            any_field = True
 
        if new_webaz is not None:
 
            plantdata["GR_PR_TIME"] = sap_numc(new_webaz)
            plantdatax["GR_PR_TIME"] = "X"
 
            any_field = True
 
        if new_dzeit is not None:
 
            plantdata["INHSEPRODT"] = sap_numc(new_dzeit)
            plantdatax["INHSEPRODT"] = "X"
 
            any_field = True
 
        if not any_field:
 
            print("Linha ignorada")
 
            results.append({
                "Material": matnr18,
                "Centro": werks,
                "Antes": before,
                "Depois": None,
                "Status": "IGNORADO"
            })
 
            continue
 
        try:
 
            headdata = {
                "MATERIAL": matnr18
            }
 
            params = dict(
                HEADDATA=headdata,
                PLANTDATA=plantdata,
                PLANTDATAX=plantdatax
            )
 
            if TESTRUN:
                params["TESTRUN"] = "X"
 
            print("Chamando BAPI...")
 
            resp = conn.call("BAPI_MATERIAL_SAVEDATA", **params)
 
            has_error, msgs = show_return(resp.get("RETURN"))
 
            print("Mensagens:", msgs)
 
            after = None
 
            if not has_error:
 
                if not TESTRUN:
 
                    conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
 
                    after = get_lead_times(conn, matnr18, werks)
 
                status = "TESTE" if TESTRUN else "OK"
 
            else:
 
                conn.call("BAPI_TRANSACTION_ROLLBACK")
 
                status = "ERRO"
 
            results.append({
                "Material": matnr18,
                "Centro": werks,
                "Antes": before,
                "Depois": after,
                "Status": status,
                "Mensagens": " | ".join(msgs)
            })
 
        except:
 
            traceback.print_exc()
 
            results.append({
                "Material": matnr18,
                "Centro": werks,
                "Antes": before,
                "Depois": None,
                "Status": "EXCEPTION"
            })
 
    return results
 
 
# =========================
# EXECUÇÃO
# =========================
 
if __name__ == "__main__":
 
    results = main()
 
    try:
 
        df = pd.DataFrame(results)
 
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
 
        out_path = os.path.join(
            os.path.dirname(EXCEL_PATH),
            f"resultado_leadtime_{ts}.csv"
        )
 
        df.to_csv(out_path, index=False, encoding="utf-8")
 
        print("\nRelatório salvo em:", out_path)
 
    except:
 
        traceback.print_exc()