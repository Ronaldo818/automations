import win32com.client
import pandas as pd
import time

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\python_scripts\Planilhas\Ordens internas.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\Planilhas\Ordens internas_logs.xlsx"

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
# FUNÇÕES
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


def formatar_coeficiente(valor):
    if pd.isna(valor):
        return ""

    try:
        # remove decimal (.0)
        valor = str(int(float(valor)))
    except:
        valor = str(valor).strip()

    # garante 10 dígitos (ajuste se necessário)
    valor = valor.zfill(10)

    # formato SAP: X.XXX.XXX.XXX
    return f"{valor[0]}.{valor[1:4]}.{valor[4:7]}.{valor[7:10]}"


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


def ajustar_scroll(session, linha_global):
    tabela = session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES")

    # pega quantidade real de linhas visíveis
    linhas_visiveis = tabela.visibleRowCount

    # calcula posição do scroll corretamente
    pos_scroll = max(linha_global - linhas_visiveis + 1, 0)

    tabela.verticalScrollbar.position = pos_scroll

    # espera até o SAP realmente atualizar
    time.sleep(0.5)

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

        linha_global = encontrar_linha_vazia(session)

        for index, row in grupo.iterrows():

            receptor = str(row['Receptor de apropriação']).strip()

            if receptor == "" or receptor.lower() == "nan":
                continue

            percentual = str(int(float(row['Percentual'])))
            coef = formatar_coeficiente(row['Coeficiente'])

            ativo, sub = limpar_ativo(receptor)

            linha = ajustar_scroll(session, linha_global)

            # =========================
            # KONTY
            # =========================
            campo_konty = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtCOBRB-KONTY[0,{linha}]"
            )
            campo_konty.setFocus()
            campo_konty.text = "IMO"
            session.findById("wnd[0]").sendVKey(0)

            # =========================
            # ATIVO
            # =========================
            campo_ativo = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtDKOBR-EMPGE[1,{linha}]"
            )
            campo_ativo.setFocus()
            campo_ativo.text = f"{ativo}-{sub}"
            session.findById("wnd[0]").sendVKey(0)

            # =========================
            # %
            # =========================
            campo_percent = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/txtCOBRB-PROZS[3,{linha}]"
            )
            campo_percent.setFocus()
            campo_percent.text = percentual

            # =========================
            # COEFICIENTE
            # =========================
            campo_coef = session.findById(
                f"wnd[0]/usr/tblSAPLKOBSTC_RULES/txtCOBRB-AQZIF[4,{linha}]"
            )
            campo_coef.setFocus()

            # limpa antes (evita erro de bloqueio)
            campo_coef.text = ""

            # pequena pausa ajuda SAP não travar
            time.sleep(0.1)

            campo_coef.text = coef
            session.findById("wnd[0]").sendVKey(0)

            linha_global += 1

        # =========================
        # SALVAR
        # =========================
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

# =========================
# LOG
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada.")