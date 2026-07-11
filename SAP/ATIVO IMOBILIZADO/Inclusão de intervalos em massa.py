import win32com.client
import pandas as pd

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\python_scripts\PLANILHAS\Imobilizados_AS02 - Tranf.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\PLANILHAS\Imobilizados_AS02 - Tranf_logs.xlsx"

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

        valor = str(row["Imobilizado"]).strip()

        try:
            imobilizado, subnumero = valor.split("-")
        except:
            imobilizado = valor
            subnumero = "0"

        empresa = str(row.get("Empresa", "2000")).strip()
        data = str(row["Data"]).strip()
        centro = str(row["Centro"]).strip()
        centro_custo = str(row["CentroCusto"]).strip()

        # =========================
        # AS02
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

        tipo_msg = session.findById("wnd[0]/sbar").MessageType
        mensagem = session.findById("wnd[0]/sbar").text

        if tipo_msg == "E":
            raise Exception(mensagem)

        # =========================
        # ABA DEPENDENTE DO TEMPO
        # =========================
        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02"
        ).select()

        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/"
            "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/btnTIME"
        ).press()

        # Novo intervalo
        session.findById("wnd[0]/usr/btn%#AUTOTEXT002").press()

        # Data
        session.findById("wnd[1]/usr/ctxtANLZ-ADATU").text = data
        session.findById("wnd[1]").sendVKey(0)

        # Valida linha criada
        data_tela = session.findById(
            "wnd[0]/usr/tblSAPLAISTTIME/ctxtANLZ-ADATU[0,0]"
        ).text.strip()

        if data_tela != data:
            raise Exception(
                f"Intervalo incorreto. Esperado {data} e encontrado {data_tela}"
            )

        # Limpa campos
        session.findById(
            "wnd[0]/usr/tblSAPLAISTTIME/ctxtANLZ-KOSTL[1,0]"
        ).text = ""

        session.findById(
            "wnd[0]/usr/tblSAPLAISTTIME/ctxtANLZ-WERKS[3,0]"
        ).text = ""

        session.findById(
            "wnd[0]/usr/tblSAPLAISTTIME/ctxtANLZ-PRCTR[9,0]"
        ).text = ""

        session.findById(
            "wnd[0]/usr/tblSAPLAISTTIME/ctxtANLZ-SEGMENT[10,0]"
        ).text = ""

        # Preenche novos dados
        session.findById(
            "wnd[0]/usr/tblSAPLAISTTIME/ctxtANLZ-KOSTL[1,0]"
        ).text = centro_custo

        session.findById(
            "wnd[0]/usr/tblSAPLAISTTIME/ctxtANLZ-WERKS[3,0]"
        ).text = centro

        # SAP recalcula PRCTR e SEGMENT
        session.findById("wnd[0]").sendVKey(0)

        tipo_msg = session.findById("wnd[0]/sbar").MessageType
        mensagem = session.findById("wnd[0]/sbar").text

        if tipo_msg == "E":
            raise Exception(mensagem)

        # Salvar
        session.findById("wnd[0]/tbar[0]/btn[11]").press()

        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        tipo_msg = session.findById("wnd[0]/sbar").MessageType
        mensagem = session.findById("wnd[0]/sbar").text

        if tipo_msg == "E":
            raise Exception(mensagem)

        log.append({
            "linha": index + 2,
            "imobilizado": imobilizado,
            "subnumero": subnumero,
            "empresa": empresa,
            "data": data,
            "centro": centro,
            "centro_custo": centro_custo,
            "status": "SUCESSO",
            "mensagem": mensagem
        })

    except Exception as e:

        try:
            session.findById("wnd[0]/tbar[0]/btn[12]").press()
        except:
            pass

        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        log.append({
            "linha": index + 2,
            "imobilizado": valor,
            "empresa": empresa if 'empresa' in locals() else "",
            "data": data if 'data' in locals() else "",
            "centro": centro if 'centro' in locals() else "",
            "centro_custo": centro_custo if 'centro_custo' in locals() else "",
            "status": "ERRO",
            "mensagem": str(e)
        })

    finally:
        pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada com sucesso.")