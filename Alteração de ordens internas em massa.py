import win32com.client
import pandas as pd
import time

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\Users\ronaldo.gontijo\Downloads\Ordens internas.xlsx"
ARQUIVO_LOG = r"C:\Users\ronaldo.gontijo\Downloads\Ordens internas_logs.xlsx"

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
df.columns = df.columns.str.strip()

log = []

# =========================
# FUNÇÃO ATIVO
# =========================
def limpar_ativo(valor):
    if pd.isna(valor):
        return "", ""

    valor = str(valor).strip()

    if "-" in valor:
        ativo, sub = valor.split("-")
    else:
        ativo = valor
        sub = "0"

    return ativo.strip(), sub.strip()

# =========================
# FUNÇÃO LINHA INICIAL
# =========================
def encontrar_linha_vazia(session):
    for i in range(0, 200):
        try:
            campo = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtCOBRB-KONTY[0,{i}]"
            )
            if campo.text.strip() == "":
                return i
        except:
            return i
    return 0

# =========================
# SCROLL CONTROLADO
# =========================
def ajustar_scroll(session, linha_global):
    tabela = session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES")

    linhas_visiveis = 12
    pos_scroll = max(linha_global - linhas_visiveis + 1, 0)

    tabela.verticalScrollbar.position = pos_scroll
    time.sleep(0.3)

    return linha_global - pos_scroll

# =========================
# LOOP
# =========================
for ordem, grupo in df.groupby('ORDEM'):

    try:
        ordem = str(ordem).strip()

        # KO02
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nKO02"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtCOAS-AUFNR").text = ordem
        session.findById("wnd[0]").sendVKey(0)

        # Ir para regra
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        time.sleep(2)

        # 🔥 PEGA LINHA UMA VEZ SÓ
        linha_global = encontrar_linha_vazia(session)

        for index, row in grupo.iterrows():

            receptor = str(row['Receptor de apropriação']).strip()

            if receptor == "" or receptor.lower() == "nan":
                continue

            percentual = str(int(float(row['Percentual'])))
            coef = str(row['Coeficiente']).strip()

            ativo, sub = limpar_ativo(receptor)

            # 🔥 SCROLL CORRETO
            linha = ajustar_scroll(session, linha_global)

            # KONTY
            campo_konty = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtCOBRB-KONTY[0,{linha}]"
            )
            campo_konty.setFocus()
            campo_konty.text = "IMO"
            session.findById("wnd[0]").sendVKey(0)

            # ATIVO
            campo_ativo = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtDKOBR-EMPGE[1,{linha}]"
            )
            campo_ativo.setFocus()
            campo_ativo.text = f"{ativo}-{sub}"
            session.findById("wnd[0]").sendVKey(0)

            # %
            campo_percent = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/txtCOBRB-PROZS[3,{linha}]"
            )
            campo_percent.setFocus()
            campo_percent.text = percentual

            # coef
            campo_coef = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/txtCOBRB-AQZIF[4,{linha}]"
            )
            campo_coef.setFocus()
            campo_coef.text = coef

            session.findById("wnd[0]").sendVKey(0)

            # 🔥 AVANÇA LINHA MANUALMENTE
            linha_global += 1

        # SALVAR
        session.findById("wnd[0]/tbar[0]/btn[11]").press()

        status = session.findById("wnd[0]/sbar").text

        if "erro" in status.lower():
            raise Exception(status)

        for index, row in grupo.iterrows():
            log.append({
                "linha": index + 2,
                "ordem": ordem,
                "ativo": row['Receptor de apropriação'],
                "status": "SUCESSO",
                "mensagem": status
            })

    except Exception as e:
        print(f"Erro na ordem {ordem}: {str(e)}")

        for index, row in grupo.iterrows():
            log.append({
                "linha": index + 2,
                "ordem": ordem,
                "ativo": row.get('Receptor de apropriação', ''),
                "status": "ERRO",
                "mensagem": str(e)
            })

        break

# LOG
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada.")