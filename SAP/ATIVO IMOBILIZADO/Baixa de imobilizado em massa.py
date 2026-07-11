import win32com.client
import pandas as pd
import time

ARQUIVO_ENTRADA = r"C:\python_scripts\PLANILHAS\Baixa_imobilizado.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\PLANILHAS\Baixa_imobilizado_logs.xlsx"

# =========================
# CONEXÃO SAP
# =========================
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# =========================
# FUNÇÕES AUXILIARES
# =========================
def esperar(seg=0.5):
    time.sleep(seg)

def limpar_status(data_doc, data_ref):
    for _ in range(5):
        status = session.findById("wnd[0]/sbar").text.lower()

        if "data do documento" in status:
            session.findById(
                "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/"
                "tabsTABSTRIP100/tabpTAB01/"
                "ssubSUBSC:SAPLATAB:0202/"
                "subAREA2:SAPLAMDPS2I:1105/"
                "subSUBSCREEN1:SAPLAMDPS2I:0200/"
                "ctxtRAIFP1-BLDAT"
            ).text = data_doc

        elif "data de referência" in status:
            session.findById(
                "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/"
                "tabsTABSTRIP100/tabpTAB01/"
                "ssubSUBSC:SAPLATAB:0202/"
                "subAREA2:SAPLAMDPS2I:1105/"
                "subSUBSCREEN3:SAPLAMDPS2I:0202/"
                "ctxtRAIFP1-BZDAT"
            ).text = data_ref

        else:
            break

        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)

def preencher_campo(caminho, valor):
    campo = session.findById(caminho)
    campo.setFocus()
    campo.text = valor

def preencher_bldat(data_doc):
    try:
        preencher_campo(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/"
            "tabsTABSTRIP100/tabpTAB01/"
            "ssubSUBSC:SAPLATAB:0202/"
            "subAREA2:SAPLAMDPS2I:1105/"
            "subSUBSCREEN1:SAPLAMDPS2I:0200/"
            "ctxtRAIFP1-BLDAT",
            data_doc
        )
        return True
    except:
        return False

def preencher_bzdat(data_ref):
    try:
        preencher_campo(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/"
            "tabsTABSTRIP100/tabpTAB01/"
            "ssubSUBSC:SAPLATAB:0202/"
            "subAREA2:SAPLAMDPS2I:1105/"
            "subSUBSCREEN3:SAPLAMDPS2I:0202/"
            "ctxtRAIFP1-BZDAT",
            data_ref
        )
        return True
    except:
        return False

# =========================
# LER PLANILHA
# =========================
df = pd.read_excel(ARQUIVO_ENTRADA)
df.columns = df.columns.str.strip()

log = []

# =========================
# LOOP
# =========================
for index, row in df.iterrows():

    try:
        print(f"Processando linha {index + 2}")

        ativo = format(int(row['Imobilizado']), 'd')
        subnumero = str(row['Subnº'])

        data_doc = pd.to_datetime(row['Data do documento']).strftime("%d.%m.%Y")
        data_lanc = pd.to_datetime(row['Data de lançamento']).strftime("%d.%m.%Y")
        data_ref = pd.to_datetime(row['Data de referência']).strftime("%d.%m.%Y")

        texto = str(row['Texto'])
        mes = str(int(row['Período contábil']))
        referencia = str(row['Referência'])
        atribuicao = str(row['Atribuição'])
        tipo = str(row.get('Referência.1', ''))
        nota = str(row.get('Nota', ''))

        # =========================
        # INICIAR TRANSAÇÃO
        # =========================
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nABAVN"
        session.findById("wnd[0]").sendVKey(0)
        esperar(1)

        try:
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "2000"
            session.findById("wnd[1]").sendVKey(0)
        except:
            pass

        # =========================
        # ATIVO
        # =========================
        preencher_campo("wnd[0]/usr/subOBJECT:SAPLAMDPS2I:0300/ctxtRAIFP2-ANLN1", ativo)
        preencher_campo("wnd[0]/usr/subOBJECT:SAPLAMDPS2I:0300/ctxtRAIFP2-ANLN2", subnumero)

        session.findById("wnd[0]").sendVKey(0)
        esperar()
        limpar_status(data_doc, data_ref)

        # =========================
        # DATAS (NÃO MEXER — JÁ FUNCIONA)
        # =========================
        preencher_bldat(data_doc)
        session.findById("wnd[0]").sendVKey(0)
        limpar_status(data_doc, data_ref)

        preencher_bzdat(data_ref)
        session.findById("wnd[0]").sendVKey(0)
        limpar_status(data_doc, data_ref)

        preencher_campo(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/"
            "tabsTABSTRIP100/tabpTAB01/"
            "ssubSUBSC:SAPLATAB:0202/"
            "subAREA2:SAPLAMDPS2I:1105/"
            "subSUBSCREEN2:SAPLAMDPS2I:0201/"
            "ctxtRAIFP1-BUDAT",
            data_lanc
        )

        session.findById("wnd[0]").sendVKey(0)
        limpar_status(data_doc, data_ref)

        # =========================
        # TEXTO
        # =========================
        preencher_campo(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/"
            "tabsTABSTRIP100/tabpTAB01/"
            "ssubSUBSC:SAPLATAB:0202/"
            "subAREA2:SAPLAMDPS2I:1105/"
            "subSUBSCREEN4:SAPLAMDPS2I:0206/"
            "txtRAIFP2-SGTXT",
            texto
        )

        # =========================
        # ABA 2
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select()
        esperar()

        preencher_campo(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/"
            "ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAMDPS2I:1000/"
            "subSUBSCREEN1:SAPLAMDPS2I:0203/txtRAIFP2-MONAT",
            mes
        )

        preencher_campo(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/"
            "ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAMDPS2I:1002/"
            "subSUBSCREEN1:SAPLAMDPS2I:0207/txtRAIFP1-XBLNR",
            referencia
        )

        preencher_campo(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/"
            "ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAMDPS2I:1002/"
            "subSUBSCREEN2:SAPLAMDPS2I:0208/txtRAIFP2-ZUONR",
            atribuicao
        )

        # =========================
        # ABA 3
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select()
        esperar()

        if "anterior" in tipo.lower():
            session.findById(
                "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/"
                "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDPS2I:1004/"
                "subSUBSCREEN1:SAPLAMDPS2I:0401/radRAIFP2-XAALT"
            ).select()

        # =========================
        # ABA 4
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select()
        esperar()

        session.findById(
            "wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/"
            "ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDS:0600/cntlEDITOR/shell"
        ).text = nota + "\n"

        # =========================
        # SALVAR
        # =========================
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        esperar(1)

        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        status = session.findById("wnd[0]/sbar").text

        log.append({
            "linha": index + 2,
            "ativo": ativo,
            "status": "SUCESSO",
            "mensagem": status
        })

    except Exception as e:
        log.append({
            "linha": index + 2,
            "ativo": row.get('Imobilizado', ''),
            "status": "ERRO",
            "mensagem": str(e)
        })

# =========================
# LOG FINAL
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada.")