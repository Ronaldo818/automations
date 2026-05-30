import win32com.client
import pandas as pd

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\python_scripts\Planilhas\Imobilizados_AS02.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\Planilhas\Imobilizados_AS02_logs.xlsx"

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

log = []

# =========================
# LOOP
# =========================
for index, row in df.iterrows():

    try:
        # =========================
        # EXTRAI IMOBILIZADO + SUBNUMERO
        # =========================
        valor = str(row['Imobilizado']).strip()

        try:
            imobilizado, subnumero = valor.split("-")
        except:
            imobilizado = valor
            subnumero = "0"

        empresa = str(row.get('Empresa', '2000'))
        ordem = str(row['Ordem'])

        # =========================
        # ENTRA NA AS02
        # =========================
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nAS02"
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # TELA INICIAL
        # =========================
        session.findById("wnd[0]/usr/ctxtANLA-ANLN1").text = imobilizado
        session.findById("wnd[0]/usr/ctxtANLA-ANLN2").text = subnumero
        session.findById("wnd[0]/usr/ctxtANLA-BUKRS").text = empresa
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # VALIDA ERRO DE ENTRADA
        # =========================
        tipo_msg = session.findById("wnd[0]/sbar").MessageType
        mensagem = session.findById("wnd[0]/sbar").text

        if tipo_msg == "E":
            raise Exception(mensagem)

        # =========================
        # ABA ORIGEM
        # =========================
        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04"
        ).select()

        campo_ordem = session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/"
            "ssubSUBSC:SAPLATAB:0202/subAREA2:SAPLAIST:1182/ctxtANLA-EAUFN"
        )

        campo_ordem.text = ordem

        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # SALVAR
        # =========================
        session.findById("wnd[0]/tbar[0]/btn[11]").press()

        # popup (se aparecer)
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        tipo_msg = session.findById("wnd[0]/sbar").MessageType
        mensagem = session.findById("wnd[0]/sbar").text

        if tipo_msg == "E":
            raise Exception(mensagem)

        # =========================
        # LOG SUCESSO
        # =========================
        log.append({
            "linha": index + 2,
            "imobilizado": imobilizado,
            "subnumero": subnumero,
            "ordem": ordem,
            "status": "SUCESSO",
            "mensagem": mensagem
        })

    except Exception as e:

        erro_msg = str(e)

        # =========================
        # CANCELA ALTERAÇÃO
        # =========================
        try:
            session.findById("wnd[0]/tbar[0]/btn[12]").press()
        except:
            pass

        # fecha popup se existir
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        # =========================
        # LOG ERRO
        # =========================
        log.append({
            "linha": index + 2,
            "imobilizado": row.get('Imobilizado', ''),
            "ordem": row.get('Ordem', ''),
            "status": "ERRO",
            "mensagem": erro_msg
        })

        continue

# =========================
# SALVA LOG
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada com sucesso.")