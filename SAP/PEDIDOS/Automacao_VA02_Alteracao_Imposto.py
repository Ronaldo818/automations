import win32com.client
import pandas as pd
import time
from datetime import datetime

# ==========================================================
# CONFIG
# ==========================================================
ARQUIVO_ENTRADA = r"C:\python_scripts\Planilhas\VA02_Impostos.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\Planilhas\VA02_Impostos_Log.xlsx"

# ==========================================================
# SAP
# ==========================================================
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# ==========================================================
# PLANILHA
# ==========================================================
df = pd.read_excel(ARQUIVO_ENTRADA)

df["Pedido"] = df["Pedido"].astype(str).str.replace(".0", "", regex=False)
df["Item"] = df["Item"].astype(str).str.replace(".0", "", regex=False)
df["Codigo_Imposto"] = df["Codigo_Imposto"].astype(str).str.strip()

log = []

CAMPO_PEDIDO = "wnd[0]/usr/ctxtVBAK-VBELN"
CAMPO_POSNR = "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR"
CAMPO_IMPOSTO = ("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\05/"
                 "ssubSUBSCREEN_BODY:SAPMV45A:4470/"
                 "ctxtVBAP-J_1B_TAX_SITUATION")
BOTAO_PROXIMO = "wnd[0]/usr/btnBT_ITPP"

def status_bar():
    return session.findById("wnd[0]/sbar").text.strip()

def tipo_status():
    try:
        return session.findById("wnd[0]/sbar").messageType
    except:
        return ""

def confirmar_popups():

    while True:

        try:

            if session.Children.Count > 1:

                session.findById(
                    "wnd[1]/tbar[0]/btn[0]"
                ).press()

                time.sleep(0.2)

            else:
                break

        except:
            break  

def verificar_erro():
    if tipo_status() == "E":
        raise Exception(status_bar())

def fechar_popup():
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

def abrir_va02(pedido):
    session.findById("wnd[0]/tbar[0]/okcd").text="/nVA02"
    session.findById("wnd[0]").sendVKey(0)
    session.findById(CAMPO_PEDIDO).text=str(pedido)
    session.findById("wnd[0]").sendVKey(0)
    verificar_erro()

def abrir_primeiro_item():
    campo=("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/"
           "ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/"
           "tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]")
    session.findById(campo).setFocus()
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\05").select()

def item_atual():
    return session.findById(CAMPO_POSNR).text.strip()

def alterar_codigo(codigo):
    c=session.findById(CAMPO_IMPOSTO)
    c.text=str(codigo)
    session.findById("wnd[0]").sendVKey(0)
    confirmar_popups()
    verificar_erro()

def ir_proximo():
    ant=item_atual()
    session.findById(BOTAO_PROXIMO).press()
    ini=time.time()
    while time.time()-ini<5:
        if item_atual()!=ant:
            return item_atual()
        time.sleep(0.2)
    raise Exception("Não foi possível avançar para o próximo item.")

def localizar_item(item_desejado):
    atual=item_atual()
    if atual==str(item_desejado):
        return
    for _ in range(500):
        atual=ir_proximo()
        if atual==str(item_desejado):
            return
    raise Exception(f"Item {item_desejado} não encontrado.")

def salvar():
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    confirmar_popups()
    verificar_erro()
    return status_bar()

for pedido,grupo in df.groupby("Pedido"):
    grupo=grupo.sort_values("Item", key=lambda s:s.astype(int))
    try:
        abrir_va02(pedido)
        abrir_primeiro_item()
        for _,row in grupo.iterrows():
            localizar_item(row["Item"])
            alterar_codigo(row["Codigo_Imposto"])
            log.append({
                "linha":row.name+2,
                "pedido":pedido,
                "item":row["Item"],
                "codigo_imposto":row["Codigo_Imposto"],
                "status":"SUCESSO",
                "mensagem":"Item alterado",
                "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            })
        msg=salvar()
        for l in log:
            if l["pedido"]==pedido and l["status"]=="SUCESSO":
                l["mensagem"]=msg
    except Exception as e:
        for _,row in grupo.iterrows():
            log.append({
                "linha":row.name+2,
                "pedido":pedido,
                "item":row["Item"],
                "codigo_imposto":row["Codigo_Imposto"],
                "status":"ERRO",
                "mensagem":str(e),
                "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            })
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text="/n"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass

pd.DataFrame(log).to_excel(ARQUIVO_LOG,index=False)
print("Execução finalizada.")
