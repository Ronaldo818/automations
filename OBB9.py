import pandas as pd
import win32com.client
import time

CAMINHO_EXCEL = r"C:\Users\ronaldo.gontijo\Downloads\condicoes_pagamento.xlsx"
CAMINHO_SAIDA = r"C:\Users\ronaldo.gontijo\Downloads\condicoes_parceladas_resultados.xlsx"
ABA = "PARCELADAS"
REQUEST = "DS4K9A05KJ"

# ==============================
# UTIL
# ==============================
def extrair_diretas(texto):
    if pd.isna(texto):
        return []
    return [x.strip() for x in texto.split(",")]

def gerar_parcelas(diretas):
    qtd = len(diretas)
    if qtd == 0:
        return []

    percentual_base = round(100 / qtd, 3)

    parcelas = []
    soma = 0

    for i, base in enumerate(diretas):
        if i < qtd - 1:
            percentual = percentual_base
            soma += percentual
        else:
            percentual = round(100 - soma, 3)

        parcelas.append({
            "parcela": i + 1,
            "percentual": percentual,
            "base": base
        })

    return parcelas

def formatar_percentual(valor):
    return f"{valor:.3f}".replace(".", ",")

# ==============================
# SAP
# ==============================
def conectar_sap():
    sap = win32com.client.GetObject("SAPGUI")
    application = sap.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session

def entrar_condicao(session, zterm):
    # entra na OBB9
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nOBB9"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    # busca
    session.findById("wnd[0]/usr/btnVIM_POSI_PUSH").press()
    time.sleep(1)

    session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = zterm
    session.findById("wnd[1]").sendVKey(0)
    time.sleep(1)

def criar_nova(session):
    # 🔥 ESSENCIAL — entrar em modo criação
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    time.sleep(1)

# ==============================
# PREENCHIMENTO
# ==============================
def preencher_parcelas(session, zterm, parcelas):
    tabela = "wnd[0]/usr/tblSAPL0F30TCTRL_V_T052S"

    for i, p in enumerate(parcelas):

        session.findById(f"{tabela}/ctxtV_T052S-ZTERM[0,{i}]").text = zterm
        session.findById(f"{tabela}/txtV_T052S-RATNR[1,{i}]").text = str(p["parcela"])
        session.findById(f"{tabela}/txtV_T052S-RATPZ[2,{i}]").text = formatar_percentual(p["percentual"])

        campo = session.findById(f"{tabela}/ctxtV_T052S-RATZT[3,{i}]")
        campo.text = p["base"]
        campo.setFocus()
        campo.caretPosition = len(p["base"])

        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.3)

# ==============================
# PROCESSO
# ==============================
def processar():

    df = pd.read_excel(CAMINHO_EXCEL, sheet_name=ABA)
    session = conectar_sap()

    for index, row in df.iterrows():

        zterm = str(row["Condicao"]).strip()
        diretas = extrair_diretas(row["Diretas Relacionadas"])

        print(f"\nProcessando {zterm}...")

        try:
            entrar_condicao(session, zterm)

            # 🔥 GARANTE QUE ESTÁ NA TELA CERTA
            criar_nova(session)

            parcelas = gerar_parcelas(diretas)

            preencher_parcelas(session, zterm, parcelas)

            # salvar
            session.findById("wnd[0]/tbar[0]/btn[11]").press()

            # request
            try:
                session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = REQUEST
                session.findById("wnd[1]").sendVKey(0)
            except:
                pass

            status = "CRIADO"

        except Exception as e:
            status = f"ERRO: {str(e)}"

        df.at[index, "STATUS"] = status

    df.to_excel(CAMINHO_SAIDA, index=False)
    print("\nFinalizado!")

# ==============================
# EXECUÇÃO
# ==============================
if __name__ == "__main__":
    processar()