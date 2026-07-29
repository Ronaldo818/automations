import time
import logging
import pandas as pd
import win32com.client

ARQUIVO_ENTRADA=r"C:\python_scripts\PLANILHAS\Vinculo_Ativos.xlsx"
ARQUIVO_LOG_EXCEL=r"C:\python_scripts\PLANILHAS\Vinculo_Ativos_logs.xlsx"
ARQUIVO_LOG_TXT=r"C:\python_scripts\PLANILHAS\Vinculo_Ativos_logs.txt"

logging.basicConfig(filename=ARQUIVO_LOG_TXT,level=logging.INFO,format="%(asctime)s - %(levelname)s - %(message)s",encoding="utf-8")

def conectar():
    sap=win32com.client.GetObject("SAPGUI")
    app=sap.GetScriptingEngine
    return app.Children(0).Children(0)

def aguardar(session):
    while session.Busy:
        time.sleep(0.2)

def abrir_transacao(session):
    session.findById("wnd[0]/tbar[0]/okcd").text="/nZFI_VINCULA_ATIVO"
    session.findById("wnd[0]").sendVKey(0)
    aguardar(session)

def preencher_ativo(session,a1,a2):
    session.findById("wnd[0]/usr/ctxtP_ANLN1").text=a1
    session.findById("wnd[0]/usr/ctxtP_ANLN2").text=a2
    session.findById("wnd[0]").sendVKey(0)
    aguardar(session)

def preencher(session,row):
    session.findById("wnd[0]/usr/txtP_MAT").text=str(row["Matrícula"]).strip()
    session.findById("wnd[0]").sendVKey(0)
    aguardar(session)
    session.findById("wnd[0]/usr/txtP_PARTNR").text=str(row["Responsável"]).strip()
    session.findById("wnd[0]/usr/txtP_SETOR").text = (
    str(row["Setor"]).strip()[:20]
    )   
    session.findById("wnd[0]/usr/ctxtP_DATA").text=str(row["Data Vínculo"]).strip()
    session.findById("wnd[0]/usr/ctxtP_KOSTL").text=str(row["Centro de custo do responsável"]).strip()

def executar(session):
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    aguardar(session)
    sb=session.findById("wnd[0]/sbar")
    return sb.MessageType,sb.Text

def separar(v):
    p=str(v).strip().split("-")
    return p[0], p[1] if len(p)>1 else "0"

def main():
    ini=time.time()
    df=pd.read_excel(ARQUIVO_ENTRADA)
    session=conectar()
    log=[]
    for idx,row in df.iterrows():
        ativo=str(row["Ativo Imobilizado"]).strip()
        try:
            abrir_transacao(session)
            a1,a2=separar(ativo)
            preencher_ativo(session,a1,a2)
            preencher(session,row)
            tipo,msg=executar(session)
            status={"E":"ERRO","W":"AVISO"}.get(tipo,"SUCESSO")
            log.append({"Linha":idx+2,"Ativo":ativo,"Status":status,"Mensagem":msg})
            logging.info("%s %s %s",ativo,status,msg)
        except Exception as e:
            log.append({"Linha":idx+2,"Ativo":ativo,"Status":"ERRO","Mensagem":str(e)})
            logging.exception("Erro")
    pd.DataFrame(log).to_excel(ARQUIVO_LOG_EXCEL,index=False)
    print(f"Finalizado em {time.time()-ini:.1f}s")

if __name__=="__main__":
    main()
