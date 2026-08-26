import win32com.client
import pandas as pd
from datetime import datetime
import time

# ============================================================
# CONFIG
# ============================================================

ARQUIVO_ENTRADA = r"C:\python_scripts\PLANILHAS\BPs Data.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\PLANILHAS\BPs Data_logs.xlsx"

TIMEOUT_SAP = 15


# ============================================================
# CONEXÃO SAP
# ============================================================

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

session.findById("wnd[0]").maximize()


# ============================================================
# PLANILHA
# ============================================================

df = pd.read_excel(
    ARQUIVO_ENTRADA,
    dtype={"BP": str}
)

log = []


# ============================================================
# FUNÇÕES
# ============================================================

def limpar_valor(valor):

    if pd.isna(valor):
        return ""

    return str(valor).strip()


def formatar_bp(bp):

    bp = str(bp).strip()

    if "." in bp:
        bp = bp.split(".")[0]

    return bp.zfill(10)


def salvar_log():

    pd.DataFrame(log).to_excel(
        ARQUIVO_LOG,
        index=False
    )


def registrar_log(
    linha,
    bp,
    valor,
    status,
    mensagem
):

    log.append({
        "linha": linha,
        "bp": bp,
        "valor": valor,
        "data_hora": datetime.now().strftime(
            "%d/%m/%Y %H:%M:%S"
        ),
        "status": status,
        "mensagem": mensagem
    })

    salvar_log()


def esperar_sap(timeout=TIMEOUT_SAP):

    """
    Aguarda o SAP terminar o processamento.
    """

    inicio = time.time()

    while time.time() - inicio < timeout:

        try:

            if not session.Busy:
                return

        except:

            pass

        time.sleep(0.1)

    raise Exception(
        "SAP demorou demais para responder."
    )


def esperar_elemento(
    id_elemento,
    timeout=TIMEOUT_SAP
):

    """
    Aguarda o elemento realmente aparecer.
    """

    inicio = time.time()

    while time.time() - inicio < timeout:

        try:

            elemento = session.findById(
                id_elemento
            )

            return elemento

        except:

            time.sleep(0.1)

    raise Exception(
        f"Elemento não encontrado:\n{id_elemento}"
    )


def pressionar(id_elemento):

    elemento = esperar_elemento(
        id_elemento
    )

    elemento.press()

    esperar_sap()


# ============================================================
# IDS SAP
# ============================================================

CAMPO_BUSCA_BP = (
    "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/"
    "subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/"
    "tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/"
    "ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/"
    "subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/"
    "subSCREEN_3200_SEARCH_FIELDS_AREA:"
    "SAPLBUPA_DIALOG_SEARCH:2100/"
    "txtBUS_JOEL_SEARCH-PARTNER_NUMBER"
)


GRID_RESULTADO = (
    "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/"
    "subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/"
    "tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/"
    "ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/"
    "subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/"
    "subSCREEN_3200_RESULT_AREA:"
    "SAPLBUPA_DIALOG_JOEL:1060/"
    "ssubSCREEN_1060_RESULT_AREA:"
    "SAPLBUPA_DIALOG_JOEL:1080/"
    "cntlSCREEN_1080_CONTAINER/"
    "shellcont/shell"
)


CAMPO_VALID_TO = (
    "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/"
    "subSCREEN_1010_RIGHT_AREA:"
    "SAPLBUPA_DIALOG_JOEL:1000/"
    "ssubSCREEN_1000_WORKAREA_AREA:"
    "SAPLBUPA_DIALOG_JOEL:1100/"
    "ssubSCREEN_1100_MAIN_AREA:"
    "SAPLBUPA_DIALOG_JOEL:1101/"
    "tabsGS_SCREEN_1100_TABSTRIP/"
    "tabpSCREEN_1100_TAB_01/"
    "ssubSCREEN_1100_TABSTRIP_AREA:"
    "SAPLBUSS:0028/"
    "ssubGENSUB:SAPLBUSS:7016/"
    "subA05P02:SAPLBUA0:0310/"
    "ctxtBUS000FLDS-ADDR_VALID_TO"
)


BOTAO_EDITAR = (
    "wnd[0]/tbar[1]/btn[6]"
)


BOTAO_SALVAR = (
    "wnd[0]/tbar[0]/btn[11]"
)


BOTAO_VOLTAR = (
    "wnd[0]/tbar[0]/btn[3]"
)


# ============================================================
# LOOP
# ============================================================

for index, row in df.iterrows():

    linha_excel = index + 2

    bp = ""
    valor = ""

    inicio_bp = time.time()

    try:

        # ====================================================
        # DADOS
        # ====================================================

        bp = formatar_bp(
            row["BP"]
        )

        valor = limpar_valor(
            row["Valor"]
        )

        print()
        print("=" * 65)
        print(f"Processando BP: {bp}")
        print(f"Valor: {valor}")
        print("=" * 65)


        # ====================================================
        # VALIDAÇÃO
        # ====================================================

        if not bp:

            registrar_log(
                linha_excel,
                bp,
                valor,
                "ERRO",
                "BP não informado"
            )

            continue


        if not valor:

            registrar_log(
                linha_excel,
                bp,
                valor,
                "ERRO",
                "Valor não informado"
            )

            continue


        # ====================================================
        # ENTRAR NOVAMENTE NO BP
        # ====================================================

        session.findById(
            "wnd[0]/tbar[0]/okcd"
        ).text = "/nBP"

        session.findById(
            "wnd[0]"
        ).sendVKey(0)

        esperar_sap()


        # ====================================================
        # AGUARDAR TELA DE PESQUISA
        # ====================================================

        campo_bp = esperar_elemento(
            CAMPO_BUSCA_BP
        )


        # ====================================================
        # PESQUISAR BP
        # ====================================================

        campo_bp.text = bp

        session.findById(
            "wnd[0]"
        ).sendVKey(0)

        esperar_sap()


        # ====================================================
        # AGUARDAR RESULTADO
        # ====================================================

        grid = esperar_elemento(
            GRID_RESULTADO
        )


        # ====================================================
        # VERIFICAR RESULTADO
        # ====================================================

        try:

            quantidade_linhas = grid.RowCount

        except:

            quantidade_linhas = 0


        if quantidade_linhas == 0:

            registrar_log(
                linha_excel,
                bp,
                valor,
                "ERRO",
                "BP não encontrado"
            )

            print(
                f"BP {bp}: não encontrado."
            )

            continue


        # ====================================================
        # SELECIONAR BP
        # ====================================================

        grid.currentCellColumn = "DESCRIPTION"
        grid.selectedRows = "0"

        grid.doubleClickCurrentCell()

        esperar_sap()


        # ====================================================
        # POPUP DE PERMISSÃO
        # ====================================================

        try:

            popup = session.findById(
                "wnd[1]"
            )

            popup.findById(
                "tbar[0]/btn[0]"
            ).press()

            esperar_sap()

            registrar_log(
                linha_excel,
                bp,
                valor,
                "ERRO",
                "Sem permissão para alteração do BP"
            )

            print(
                f"BP {bp}: sem permissão."
            )

            continue

        except:

            pass


        # ====================================================
        # CAMPO DE VALIDADE
        # ====================================================

        campo_valid_to = esperar_elemento(
            CAMPO_VALID_TO
        )


        # ====================================================
        # ENTRAR EM MODO DE EDIÇÃO
        # ====================================================

        if not campo_valid_to.Changeable:

            pressionar(
                BOTAO_EDITAR
            )

            campo_valid_to = esperar_elemento(
                CAMPO_VALID_TO
            )


        # ====================================================
        # ALTERAR VALOR
        # ====================================================

        campo_valid_to.text = valor


        # ====================================================
        # SALVAR
        # ====================================================

        pressionar(
            BOTAO_SALVAR
        )


        # ====================================================
        # POPUP APÓS SALVAR
        # ====================================================

        try:

            popup = session.findById(
                "wnd[1]"
            )

            popup.findById(
                "tbar[0]/btn[0]"
            ).press()

            esperar_sap()

        except:

            pass


        # ====================================================
        # STATUS SAP
        # ====================================================

        try:

            status_sap = session.findById(
                "wnd[0]/sbar"
            ).text

        except:

            status_sap = ""


        # ====================================================
        # TEMPO
        # ====================================================

        tempo = time.time() - inicio_bp


        # ====================================================
        # LOG
        # ====================================================

        registrar_log(
            linha_excel,
            bp,
            valor,
            "SUCESSO",
            status_sap if status_sap else
            "BP alterado com sucesso"
        )


        print(
            f"BP {bp} processado com sucesso."
        )

        print(
            f"Tempo: {tempo:.2f} segundos"
        )


        # ====================================================
        # VOLTAR
        # ====================================================

        pressionar(
            BOTAO_VOLTAR
        )


        print(
            f"BP {bp}: retornou."
        )


    except Exception as e:

        mensagem = str(e)

        tempo = time.time() - inicio_bp

        print()
        print(
            f"ERRO no BP {bp}: {mensagem}"
        )

        registrar_log(
            linha_excel,
            bp,
            valor,
            "ERRO",
            mensagem
        )


        # ====================================================
        # TENTAR VOLTAR
        # ====================================================

        try:

            session.findById(
                BOTAO_VOLTAR
            ).press()

            esperar_sap()

        except:

            pass


        # ====================================================
        # NÃO INTERROMPE O LOTE
        # ====================================================

        continue


# ============================================================
# FINALIZAÇÃO
# ============================================================

print()
print("=" * 65)
print("EXECUÇÃO FINALIZADA")
print("=" * 65)

print(
    f"Total de registros: {len(df)}"
)

print(
    f"Log salvo em:\n{ARQUIVO_LOG}"
)