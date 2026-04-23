import win32com.client
import pandas as pd
from datetime import datetime

# =========================
# CONFIGURAÇÕES
# =========================
ARQUIVO_ENTRADA = r"C:\Users\ronaldo.gontijo\Downloads\Imobilizados.xlsx"
ARQUIVO_LOG = r"C:\Users\ronaldo.gontijo\Downloads\Imobilizados_logs.xlsx"

# =========================
# CONEXÃO SAP
# =========================
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# =========================
# LER PLANILHA
# =========================
df = pd.read_excel(ARQUIVO_ENTRADA)

log = []

# =========================
# LOOP
# =========================
for index, row in df.iterrows():

    try:
        origem = str(row['origem'])
        sub_origem = str(row['sub_origem'])
        destino = str(row['destino'])
        sub_destino = str(row['sub_destino'])
        valor = str(row['valor']).replace(".", ",")
        ano_exercicio = int(row['ano_exercicio'])
        ano_atual = datetime.now().year
        
        data_dt = pd.to_datetime(row['data'])
        data = data_dt.strftime("%d.%m.%Y")

        mes = int(row['periodo'])
        if mes < 1 or mes > 12:
            raise Exception(f"Período inválido: {mes}")

        texto_cab = str(row['texto_cabecalho'])
        texto_longo = str(row['texto_longo'])

        # =========================
        # INÍCIO LIMPO
        # =========================
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nABUMN"
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ORIGEM
        # =========================
        session.findById("wnd[0]/usr/subOBJECT:SAPLAMDPS2I:0300/ctxtRAIFP2-ANLN1").text = origem
        session.findById("wnd[0]/usr/subOBJECT:SAPLAMDPS2I:0300/ctxtRAIFP2-ANLN2").text = sub_origem
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # DATAS
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPLAMDPS2I:1111/subSUBSCREEN1:SAPLAMDPS2I:0200/ctxtRAIFP1-BLDAT").text = data
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPLAMDPS2I:1111/subSUBSCREEN3:SAPLAMDPS2I:0202/ctxtRAIFP1-BZDAT").text = data
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPLAMDPS2I:1111/subSUBSCREEN2:SAPLAMDPS2I:0201/ctxtRAIFP1-BUDAT").text = data
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # TEXTO CABEÇALHO
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPLAMDPS2I:1111/subSUBSCREEN4:SAPLAMDPS2I:0206/txtRAIFP2-SGTXT").text = texto_cab

        # =========================
        # DESTINO (CORRETO)
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAMDPS2I:0320/ctxtRAIFP3-ANLN1").text = destino
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAMDPS2I:0320/ctxtRAIFP3-ANLN2").text = sub_destino
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 2 - PERÍODO
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select()

        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAMDPS2I:1000/subSUBSCREEN1:SAPLAMDPS2I:0203/txtRAIFP2-MONAT").text = str(mes)
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # ABA 3 - VALOR
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select()

        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDPS2I:1005/subSUBSCREEN1:SAPLAMDPS2I:0401/txtRAIFP2-ANBTR").text = valor
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # DEFINIR TIPO (XANEU / XAALT)
        # =========================
        if ano_exercicio == ano_atual:
            session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDPS2I:1005/subSUBSCREEN1:SAPLAMDPS2I:0401/radRAIFP2-XANEU").select()
        else:
            session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDPS2I:1005/subSUBSCREEN1:SAPLAMDPS2I:0401/radRAIFP2-XAALT").select()

        # =========================
        # ABA 4 - TEXTO LONGO
        # =========================
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select()

        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDS:0600/cntlEDITOR/shell").text = texto_longo + "\n"

        # =========================
        # SALVAR
        # =========================
        session.findById("wnd[0]/tbar[0]/btn[11]").press()

        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        status = session.findById("wnd[0]/sbar").text

        log.append({
            "linha": index + 2,
            "origem": origem,
            "sequencia": sub_origem,
            "destino": destino,
            "sub_destino": sub_destino,
            "valor": valor,
            "tipo_ref": "NOVO" if ano_exercicio == ano_atual else "ANTIGO",
            "status": "SUCESSO",
            "mensagem": status
        })

    except Exception as e:

        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass

        log.append({
            "linha": index + 2,
            "origem": row.get('origem', ''),
            "sequencia": row.get('sub_origem', ''),
            "destino": row.get('destino', ''),
            "sub_destino": row.get('sub_destino', ''),
            "valor": row.get('valor', ''),
            "tipo_ref": "NOVO" if ano_exercicio == ano_atual else "ANTIGO",
            "status": "ERRO",
            "mensagem": str(e)
        })

# =========================
# LOG FINAL
# =========================
pd.DataFrame(log).to_excel(ARQUIVO_LOG, index=False)

print("Execução finalizada com sucesso.")