import win32com.client
import pandas as pd
import time
import os
from datetime import datetime


# ============================================================
# CONFIGURAÇÕES
# ============================================================

ARQUIVO_ENTRADA = r"C:\python_scripts\PLANILHAS\Excluir_XML.xlsx"
ARQUIVO_LOG = r"C:\python_scripts\PLANILHAS\Excluir_XML_logs.xlsx"

NOME_COLUNA_CHAVE = "Chave"

TEMPO_ENTRE_ETAPAS = 0.5
TEMPO_APOS_EXECUTAR = 1.0
TEMPO_APOS_CONFIRMACAO = 1.0


# ============================================================
# CONEXÃO SAP
# ============================================================

print("Conectando ao SAP GUI...")

try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    print("Conectado ao SAP com sucesso.")

except Exception as e:
    print("ERRO ao conectar ao SAP:")
    print(e)
    raise


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def esperar(seg=TEMPO_ENTRE_ETAPAS):
    time.sleep(seg)


def obter_status_sap():
    """
    Retorna exatamente a mensagem exibida na barra inferior
    do SAP.
    """

    try:
        return session.findById("wnd[0]/sbar").text.strip()

    except Exception as e:
        print(f"Não foi possível ler a barra de status: {e}")
        return ""


def verificar_tela_principal():
    """
    Verifica se a janela principal do SAP está disponível.
    """

    try:
        session.findById("wnd[0]")
        return True

    except:
        return False


def pressionar_enter():
    session.findById("wnd[0]").sendVKey(0)


def voltar_tela_anterior():
    """
    Volta uma tela utilizando o botão Voltar do SAP.
    """

    try:

        session.findById(
            "wnd[0]/tbar[0]/btn[3]"
        ).press()

        esperar()

        return True

    except Exception as e:

        print(f"Erro ao voltar: {e}")

        return False


# ============================================================
# ENTRAR NA SE38 E ABRIR O PROGRAMA
# ============================================================

def entrar_programa():

    print("Entrando na SE38...")

    # /n garante que estamos iniciando a transação novamente
    session.findById(
        "wnd[0]/tbar[0]/okcd"
    ).text = "/nSE38"

    pressionar_enter()

    esperar(1)

    print("Informando programa...")

    session.findById(
        "wnd[0]/usr/ctxtRS38M-PROGRAMM"
    ).text = "EDOC_BR_DELETE_EDOCUMENT"

    pressionar_enter()

    esperar(0.5)

    print("Executando programa...")

    session.findById(
        "wnd[0]/tbar[1]/btn[8]"
    ).press()

    esperar(1)


# ============================================================
# PROCESSAR UMA CHAVE
# ============================================================

def processar_chave(chave):

    try:

        # ----------------------------------------------------
        # PREENCHER CHAVE
        # ----------------------------------------------------

        print(f"Informando chave: {chave}")

        campo_chave = session.findById(
            "wnd[0]/usr/txtP_KEY"
        )

        campo_chave.setFocus()

        campo_chave.text = chave

        campo_chave.caretPosition = len(chave)

        # ----------------------------------------------------
        # EXECUTAR
        # ----------------------------------------------------

        print("Executando...")

        session.findById(
            "wnd[0]/tbar[1]/btn[8]"
        ).press()

        esperar(TEMPO_APOS_EXECUTAR)

        # ----------------------------------------------------
        # CONFIRMAÇÃO
        # ----------------------------------------------------

        print("Verificando confirmação...")

        try:

            session.findById(
                "wnd[1]/usr/btnBUTTON_1"
            ).press()

            print("Confirmação realizada.")

        except:

            print("Nenhum popup de confirmação encontrado.")

        # ----------------------------------------------------
        # AGUARDAR PROCESSAMENTO
        # ----------------------------------------------------

        esperar(TEMPO_APOS_CONFIRMACAO)

        # ----------------------------------------------------
        # LER BARRA DE STATUS DO SAP
        # ----------------------------------------------------

        mensagem_sap = obter_status_sap()

        print()
        print("Mensagem do SAP:")
        print(mensagem_sap)
        print()

        return mensagem_sap

    except Exception as e:

        print(f"Erro ao processar chave: {e}")

        return f"ERRO PYTHON: {str(e)}"


# ============================================================
# LER PLANILHA
# ============================================================

print()
print("Lendo planilha...")

if not os.path.exists(ARQUIVO_ENTRADA):

    raise FileNotFoundError(
        f"Arquivo não encontrado:\n{ARQUIVO_ENTRADA}"
    )


df = pd.read_excel(
    ARQUIVO_ENTRADA,
    dtype=str
)

df.columns = df.columns.str.strip()


# ============================================================
# VALIDAR COLUNA
# ============================================================

if NOME_COLUNA_CHAVE not in df.columns:

    raise ValueError(
        f"A coluna '{NOME_COLUNA_CHAVE}' não foi encontrada.\n"
        f"Colunas disponíveis: {list(df.columns)}"
    )


# ============================================================
# NORMALIZAR CHAVES
# ============================================================

df[NOME_COLUNA_CHAVE] = (
    df[NOME_COLUNA_CHAVE]
    .fillna("")
    .astype(str)
    .str.strip()
)


# Remove espaços que eventualmente estejam dentro da chave
df[NOME_COLUNA_CHAVE] = (
    df[NOME_COLUNA_CHAVE]
    .str.replace(" ", "", regex=False)
)


# ============================================================
# LOG
# ============================================================

log = []


# ============================================================
# TOTAL
# ============================================================

total = len(df)

print()
print("=" * 70)
print(f"TOTAL DE CHAVES: {total}")
print("=" * 70)


# ============================================================
# LOOP
# ============================================================

for index, row in df.iterrows():

    linha = index + 2

    chave = row[NOME_COLUNA_CHAVE]

    print()
    print("=" * 70)
    print(f"PROCESSANDO {index + 1}/{total}")
    print(f"LINHA: {linha}")
    print(f"CHAVE: {chave}")
    print("=" * 70)

    inicio = datetime.now()

    # ========================================================
    # VALIDAR CHAVE
    # ========================================================

    if not chave:

        print("Chave vazia. Pulando...")

        log.append({
            "linha": linha,
            "chave": chave,
            "mensagem_sap": "",
            "status": "ERRO",
            "detalhe": "Chave vazia",
            "data_hora": inicio.strftime("%d/%m/%Y %H:%M:%S")
        })

        continue


    # ========================================================
    # VALIDAR TAMANHO
    # ========================================================

    if len(chave) != 44:

        print(
            f"Chave inválida: possui {len(chave)} caracteres."
        )

        log.append({
            "linha": linha,
            "chave": chave,
            "mensagem_sap": "",
            "status": "ERRO",
            "detalhe": (
                f"Chave possui {len(chave)} caracteres. "
                f"Esperado: 44."
            ),
            "data_hora": inicio.strftime("%d/%m/%Y %H:%M:%S")
        })

        continue


    # ========================================================
    # VERIFICAR SAP
    # ========================================================

    if not verificar_tela_principal():

        print("Janela principal do SAP não encontrada.")

        log.append({
            "linha": linha,
            "chave": chave,
            "mensagem_sap": "",
            "status": "ERRO",
            "detalhe": "Janela principal do SAP não encontrada.",
            "data_hora": inicio.strftime("%d/%m/%Y %H:%M:%S")
        })

        break


    # ========================================================
    # ENTRAR NO PROGRAMA
    # ========================================================

    try:

        entrar_programa()

    except Exception as e:

        print(f"Erro ao abrir o programa: {e}")

        log.append({
            "linha": linha,
            "chave": chave,
            "mensagem_sap": "",
            "status": "ERRO",
            "detalhe": f"Erro ao abrir programa: {str(e)}",
            "data_hora": inicio.strftime("%d/%m/%Y %H:%M:%S")
        })

        break


    # ========================================================
    # PROCESSAR CHAVE
    # ========================================================

    mensagem_sap = processar_chave(chave)


    # ========================================================
    # DEFINIR STATUS
    # ========================================================

    if mensagem_sap:

        status = "PROCESSADO"
        detalhe = ""

    else:

        status = "SEM MENSAGEM"
        detalhe = (
            "Nenhuma mensagem foi encontrada "
            "na barra de status do SAP."
        )


    # ========================================================
    # ADICIONAR LOG
    # ========================================================

    log.append({
        "linha": linha,
        "chave": chave,
        "mensagem_sap": mensagem_sap,
        "status": status,
        "detalhe": detalhe,
        "data_hora": inicio.strftime("%d/%m/%Y %H:%M:%S")
    })


    # ========================================================
    # MOSTRAR RESULTADO
    # ========================================================

    print()
    print("RESULTADO")
    print("-" * 70)
    print(f"Status:       {status}")
    print(f"Mensagem SAP: {mensagem_sap}")
    print("-" * 70)


    # ========================================================
    # VOLTAR
    # ========================================================

    print("Voltando para a tela anterior...")

    voltar_tela_anterior()

    esperar(0.5)


# ============================================================
# SALVAR LOG
# ============================================================

print()
print("=" * 70)
print("SALVANDO LOG")
print("=" * 70)

df_log = pd.DataFrame(log)

df_log.to_excel(
    ARQUIVO_LOG,
    index=False
)


# ============================================================
# FINAL
# ============================================================

print()
print("=" * 70)
print("EXECUÇÃO FINALIZADA")
print("=" * 70)

print(f"Total processado: {len(log)}")
print(f"Log: {ARQUIVO_LOG}")
print()