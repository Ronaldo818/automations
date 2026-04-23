import pandas as pd
import win32com.client
import time
import re

CAMINHO_EXCEL = r"C:\Users\ronaldo.gontijo\Downloads\condicoes_pagamento.xlsx"
CAMINHO_SAIDA = r"C:\Users\ronaldo.gontijo\Downloads\condicoes_pagamento_resultados.xlsx"
ABA = "OBB8_DIRETAS"
REQUEST = "DS4K9A05KJ"

# ==============================
# CONTROLE DE EXECUÇÃO
# ==============================
PREENCHER_VTEXT = False
LIMPAR_VTEXT = True
MARCAR_XDEBI = False

# ==============================
# UTIL
# ==============================
def extrair_dias(texto):
    if pd.isna(texto):
        return []
    return [int(n) for n in re.findall(r'\d+', str(texto))]

# ==============================
# SAP
# ==============================
def conectar_sap():
    sap = win32com.client.GetObject("SAPGUI")
    application = sap.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session

def abrir_obb8(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nOBB8"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

# ==============================
# REQUEST
# ==============================
def tratar_request(session):
    try:
        time.sleep(1)

        if session.Children.Count > 1:
            popup = session.findById("wnd[1]")
            campo = popup.findById("usr/ctxtKO008-TRKORR")

            campo.text = ""
            campo.text = REQUEST

            popup.sendVKey(0)

            print(f"Request {REQUEST} aplicada!")

    except Exception as e:
        print(f"Erro ao tratar request: {e}")

# ==============================
# BUSCA
# ==============================
def buscar_condicao(session, zterm):
    session.findById("wnd[0]/usr/btnVIM_POSI_PUSH").press()
    time.sleep(1)

    session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = zterm
    session.findById("wnd[1]").sendVKey(0)
    time.sleep(1)

    session.findById("wnd[0]").sendVKey(2)
    time.sleep(1)

# ==============================
# VALIDAÇÃO SEGURA
# ==============================
def condicao_encontrada(session, zterm_procurado):
    try:
        zterm_tela = session.findById("wnd[0]/usr/txtV_T052-ZTERM").text.strip()
        print(f"Esperado: {zterm_procurado} | Tela: {zterm_tela}")
        return zterm_tela == zterm_procurado
    except:
        return False

# ==============================
# PREENCHIMENTO DINÂMICO
# ==============================
def preencher_campos(session, descricao, tipo, dias, metodo, texto_sd):

    # ==========================
    # LOCALIZA CAMPO VTEXT
    # ==========================
    try:
        campo = session.findById("wnd[0]/usr/txtTVZBT-VTEXT")
    except:
        campo = session.findById("wnd[0]/usr/sub:SAPL0F30:0020/txtTVZBT-VTEXT")

    # ==========================
    # LIMPAR (MESMO CICLO)
    # ==========================
    if LIMPAR_VTEXT:
        print("Limpando VTEXT (mesmo ciclo)...")

        session.findById("wnd[0]/usr/chkR052-XDEBI").selected = False
        session.findById("wnd[0]/usr/chkR052-XKRED").selected = True

        campo.setFocus()
        campo.text = ""

        # ENTER só depois de tudo
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

    # ==========================
    # PREENCHER (MESMO CICLO)
    # ==========================
    else:
        print("Preenchendo VTEXT (mesmo ciclo)...")

        session.findById("wnd[0]/usr/chkR052-XDEBI").selected = MARCAR_XDEBI
        session.findById("wnd[0]/usr/chkR052-XKRED").selected = True

        if PREENCHER_VTEXT and texto_sd:
            campo.setFocus()
            campo.text = texto_sd[:50]

        # ENTER depois de tudo
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

    # ==========================
    # RESTANTE DOS CAMPOS
    # ==========================
    session.findById("wnd[0]/usr/radR052-XBLDA").select()
    session.findById("wnd[0]/usr/txtV_T052-TEXT1").text = descricao
    session.findById("wnd[0]/usr/ctxtV_T052-ZLSCH").text = metodo

    # ==========================
    # DIRETA
    # ==========================
    if tipo in ["Z", "D"]:
        session.findById("wnd[0]/usr/chkV_T052-XCHPM").selected = True
        session.findById("wnd[0]/usr/chkV_T052-XSPLT").selected = False
        session.findById("wnd[0]/usr/txtV_T052-ZTAG1").text = str(dias[0])

    # ==========================
    # PARCELADA
    # ==========================
    else:
        session.findById("wnd[0]/usr/chkV_T052-XCHPM").selected = True
        session.findById("wnd[0]/usr/chkV_T052-XSPLT").selected = True
        session.findById("wnd[0]/usr/txtV_T052-ZTAG1").text = ""

    # ENTER final
    session.findById("wnd[0]").sendVKey(0)

    # ==========================
    # DIRETA
    # ==========================
    if tipo in ["Z", "D"]:
        session.findById("wnd[0]/usr/chkV_T052-XCHPM").selected = True
        session.findById("wnd[0]/usr/chkV_T052-XSPLT").selected = False
        session.findById("wnd[0]/usr/txtV_T052-ZTAG1").text = str(dias[0])

    # ==========================
    # PARCELADA
    # ==========================
    else:
        session.findById("wnd[0]/usr/chkV_T052-XCHPM").selected = True
        session.findById("wnd[0]/usr/chkV_T052-XSPLT").selected = True
        session.findById("wnd[0]/usr/txtV_T052-ZTAG1").text = ""

    # ENTER final (garantia)
    session.findById("wnd[0]").sendVKey(0)

# ==============================
# CRIAR
# ==============================
def criar_condicao(session, zterm, descricao, tipo, dias, metodo, texto_sd):
    try:
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
        time.sleep(1)

        session.findById("wnd[0]/usr/txtV_T052-ZTERM").text = zterm
        session.findById("wnd[0]").sendVKey(0)

        preencher_campos(session, descricao, tipo, dias, metodo, texto_sd)

        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        tratar_request(session)

        return "CRIADO"

    except Exception as e:
        return f"ERRO CRIAR: {str(e)}"

# ==============================
# ATUALIZAR
# ==============================
def atualizar_condicao(session, zterm, descricao, tipo, dias, metodo, texto_sd):
    try:
        if not condicao_encontrada(session, zterm):
            print("SEGURANÇA: ZTERM não confere, abortando update")
            return "ERRO SEGURANÇA"

        preencher_campos(session, descricao, tipo, dias, metodo, texto_sd)

        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        tratar_request(session)

        return "ATUALIZADO"

    except Exception as e:
        return f"ERRO UPDATE: {str(e)}"

# ==============================
# PROCESSO PRINCIPAL
# ==============================
def processar():
    df = pd.read_excel(CAMINHO_EXCEL, sheet_name=ABA)
    session = conectar_sap()

    for index, row in df.iterrows():

        zterm = str(row["ZTERM"]).strip()
        descricao = str(row["DESCRICAO"]).strip()
        metodo = str(row["ZLSCH"]).strip()
        texto_sd = str(row["Texto_SD"]).strip()
        dias = extrair_dias(row["DIAS"])
        tipo = zterm[0]

        print(f"\nProcessando {zterm}...")

        try:
            abrir_obb8(session)
            buscar_condicao(session, zterm)

            if condicao_encontrada(session, zterm):
                status = atualizar_condicao(session, zterm, descricao, tipo, dias, metodo, texto_sd)
            else:
                status = criar_condicao(session, zterm, descricao, tipo, dias, metodo, texto_sd)

        except Exception as e:
            status = f"ERRO GERAL: {str(e)}"

        df.at[index, "STATUS"] = status

    df.to_excel(CAMINHO_SAIDA, index=False)
    print("\nProcesso finalizado!")

# ==============================
# EXECUÇÃO
# ==============================
if __name__ == "__main__":
    processar()