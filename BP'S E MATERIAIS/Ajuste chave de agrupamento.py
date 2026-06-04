import os
import pandas as pd
from datetime import datetime
from pyrfc import Connection

# =========================
# CONFIGURAÇÕES
# =========================
EXCEL_PATH = r"C:\python_scripts\PLANILHAS\Clientes_zgrup.xlsx"
SHEET_NAME = 0
TESTRUN = True  # True = simulação

SAP_CONN = dict(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.10",
    sysnr="00",
    client="310",
    lang="PT"
)

# =========================
# FUNÇÃO RETORNO
# =========================
def show_return(ret):
    messages = []
    error = False

    for r in ret or []:
        line = f"{r['TYPE']} - {r['ID']} {r['NUMBER']}: {r['MESSAGE']}"
        messages.append(line)

        if r["TYPE"] in ("E", "A"):
            error = True

    return error, messages


# =========================
# LEITURA EXCEL
# =========================
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df.columns = [col.strip() for col in df.columns]

df["Cliente"] = df["Cliente"].astype(str).str.zfill(10)
df["Empresa"] = df["Empresa"].astype(str).str.zfill(4)
df["ZGRUP"] = df["ZGRUP"].astype(str).str.strip()

# =========================
# CONEXÃO SAP
# =========================
conn = Connection(**SAP_CONN)

results = []

# =========================
# PROCESSAMENTO
# =========================
for idx, row in df.iterrows():
    cliente = row["Cliente"]
    empresa = row["Empresa"]
    zgrup = row["ZGRUP"]

    try:
        companydata = [{
            "COMP_CODE": empresa,
            "ZGRUP": zgrup
        }]

        companydatax = [{
            "COMP_CODE": empresa,
            "COMP_CODEX": "X",
            "ZGRUP": "X"
        }]

        params = {
            "CUSTOMERNO": cliente,
            "COMPANYDATA": companydata,
            "COMPANYDATAX": companydatax
        }

        if TESTRUN:
            params["TESTRUN"] = "X"

        response = conn.call("BAPI_CUSTOMER_CHANGEFROMDATA1", **params)

        has_error, messages = show_return(response.get("RETURN"))

        if not TESTRUN and not has_error:
            conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")

        results.append({
            "Cliente": cliente,
            "Empresa": empresa,
            "Novo_ZGRUP": zgrup,
            "Status": "ERRO" if has_error else "OK",
            "Mensagens": " | ".join(messages)
        })

    except Exception as e:
        results.append({
            "Cliente": cliente,
            "Empresa": empresa,
            "Novo_ZGRUP": zgrup,
            "Status": "EXCEPTION",
            "Mensagens": str(e)
        })

# =========================
# RELATÓRIO
# =========================
output_df = pd.DataFrame(results)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

output_path = os.path.join(
    os.path.dirname(EXCEL_PATH),
    f"resultado_zgrup_{timestamp}.csv"
)

output_df.to_csv(output_path, index=False, encoding="utf-8")

print("\nRelatório salvo em:", output_path)
print(output_df.head(20))