"""
SAP Fiori - Supplier Clear Open Items (Reinicializar saídas de pagamento)
Automação via Playwright

DEPENDÊNCIAS:
    pip install playwright pandas openpyxl
    playwright install chromium

PLANILHA — colunas esperadas:
    Fornecedor | Lançamento contábil | Tipo | Valor de crédito/débito |
    Data de vencimento | Referência | Conta do razão | Forma de pagamento |
    Atribuição | Centro de lucro | Texto do item | Status

    Status: ABERTO → busca o documento e clica Compensar antes de lançar
            COMPENSADO → entra no BP e lança diretamente

MODO_TESTE = True  → preenche tudo mas clica em Simular (não posta de verdade)
MODO_TESTE = False → clica em Lançar
"""

from __future__ import annotations

import sys
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout

# ============================================================
# CONFIG
# ============================================================
EXCEL_PATH  = r"C:\python_scripts\PLANILHAS\Partidas.xlsx"
OUTPUT_PATH = r"C:\python_scripts\PLANILHAS\Partidas_logs.xlsx"

URLQAS_HOME = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Shell-home"
URLPRD_HOME = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Shell-home"

URLQAS_APP  = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Supplier-clearOpenItems&/clearing/true/{empresa}/{bp}/undefined/undefined/undefined/undefined"
URLPRD_APP  = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Supplier-clearOpenItems&/clearing/true/{empresa}/{bp}/undefined/undefined/undefined/undefined"

PROFILE_DIR_QAS = str(Path.cwd() / "pw_profile_qas")
PROFILE_DIR_PRD = str(Path.cwd() / "pw_profile_prd")

# True  → clica em Simular (seguro para validar)
# False → clica em Lançar (executa de verdade)
MODO_TESTE = False

EMPRESA_PADRAO = "2000"

TIMEOUT = 60_000   # ms — aguarda elementos
SLOW_MO = 400      # ms — pausa entre ações

# Tempo de pausa após Simular para validação visual (ms)
DELAY_SIMULACAO = 60_000

# ============================================================
# SELETORES
# ============================================================
SEL_SHELL_HEADER = "#shell-header"
SEL_USER         = "#USERNAME_FIELD-inner"
SEL_PASS         = "#PASSWORD_FIELD-inner"
SEL_LOGIN_BTN    = "#LOGIN_LINK"

SEL_BTN_COMPENSAR_PARTIDAS = "bdi#application-Supplier-clearOpenItems-component---PaymentListView--buttonAssignOpenItems-BDI-content"

SEL_INPUT_EMPRESA = "#openItemAssSelCompanyCode-inner"
SEL_INPUT_BP      = "#openItemAssSelAccountID-inner"
SEL_BTN_OK_POPUP  = "bdi#openItemsAssSelDialogOKButton-BDI-content"

# Tipo de documento contábil (preenchido com "KP" ao entrar no BP)
SEL_TIPO_DOC = "#application-Supplier-clearOpenItems-component---ClearingView--acctgDocTypeInput-input-inner"

SEL_CAMPO_BUSCA         = "#application-Supplier-clearOpenItems-component---ClearingView--openItemSearch-I"
SEL_BTN_COMPENSAR_LINHA = "bdi:has-text('Compensar')"

SEL_ABA_RAZAO  = "#application-Supplier-clearOpenItems-component---ClearingView--tabChargeOffDiff-text"

SEL_CONTA_GL   = "#application-Supplier-clearOpenItems-component---ClearingView--glItemGLAccount-application-Supplier-clearOpenItems-component---ClearingView--glItems-0-input-inner"
SEL_VALOR_DEB  = "#application-Supplier-clearOpenItems-component---ClearingView--glItemDebitAmount-application-Supplier-clearOpenItems-component---ClearingView--glItems-0-input-inner"
SEL_VALOR_CRED = "#application-Supplier-clearOpenItems-component---ClearingView--glItemCreditAmount-application-Supplier-clearOpenItems-component---ClearingView--glItems-0-input-inner"

SEL_EXPAND_GL    = "#application-Supplier-clearOpenItems-component---ClearingView--itemPanel-application-Supplier-clearOpenItems-component---ClearingView--glItems-0-expandButton-img"
SEL_ATRIBUICAO = "input[aria-labelledby*='codingBlockDetailsGroupElement_DocumentItemText-label']"
SEL_CENTRO_LUCRO = "input[aria-labelledby*='codingBlockGroupElement_ProfitCenter-label']"

SEL_ABA_ON_ACCOUNT   = "#application-Supplier-clearOpenItems-component---ClearingView--tabOnAccount-text"
SEL_ON_ACCOUNT_VALOR = "#application-Supplier-clearOpenItems-component---ClearingView--onAccountItemCreditAmount-application-Supplier-clearOpenItems-component---ClearingView--onAccountItems-0-input-inner"

SEL_EXPAND_ON_ACCOUNT = "#application-Supplier-clearOpenItems-component---ClearingView--APARItemPanel-application-Supplier-clearOpenItems-component---ClearingView--onAccountItems-0-expandButton-img"
SEL_REFERENCIA        = "#application-Supplier-clearOpenItems-component---ClearingView--OnAccountInputAssignmentReference-application-Supplier-clearOpenItems-component---ClearingView--onAccountItems-0-input-inner"
SEL_VENCIMENTO        = "#application-Supplier-clearOpenItems-component---ClearingView--OnAccountInputDueCalculationBaseDate-application-Supplier-clearOpenItems-component---ClearingView--onAccountItems-0-datePicker-inner"

SEL_BTN_SIMULAR = "bdi#application-Supplier-clearOpenItems-component---ClearingView--simulateMenuButton-internalSplitBtn-textButton-BDI-content"
SEL_BTN_LANCAR = "#application-Supplier-clearOpenItems-component---ClearingView--buttonPost"
SEL_BTN_LANCAR_FINAL = "bdi#application-AccountingDocument-manage-component---singleGLDocumentDisplay--btnPost-BDI-content"

SEL_BTN_OK_CONFIRMACAO = "bdi:has-text('OK')"

SEL_TOAST = ".sapMMessageToast"

# IDs das colunas da grid para identificar onde o valor foi encontrado
COLID_ATRIBUICAO = "application-Supplier-clearOpenItems-component---ClearingView--assignment"
COLID_REFERENCIA = "application-Supplier-clearOpenItems-component---ClearingView--documentReference"

# ============================================================
# HELPERS
# ============================================================
def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalizar_valor(valor_str: str) -> str:
    """
    Converte para formato SAP PT-BR (vírgula decimal, sem separador de milhar).
    'R$ 51,00' → '51,00' | '1.234,56' → '1234,56' | '51.00' → '51,00'
    """
    v = str(valor_str).strip()
    v = re.sub(r"[^\d\.,\-]", "", v)
    if not v or v in ("-", ",", "."):
        return "0,00"
    if "," in v and "." in v:
        v = v.replace(".", "")
    elif "." in v:
        v = v.replace(".", ",")
    parts = v.split(",")
    inteiro = parts[0].lstrip("0") or "0"
    dec = (parts[1] + "00")[:2] if len(parts) > 1 else "00"
    return f"{inteiro},{dec}"


def wait_busy_settle(page: Page, timeout=60_000):
    """Aguarda sumir todos os indicadores de carregamento UI5."""
    page.wait_for_function(
        """() => {
            const els = Array.from(document.querySelectorAll('.sapUiLocalBusyIndicator'));
            if (els.length === 0) return true;
            const visible = el => {
                const s = window.getComputedStyle(el);
                return s.display !== 'none' && s.visibility !== 'hidden'
                       && s.opacity !== '0' && el.offsetParent !== null;
            };
            return els.every(el => !visible(el));
        }""",
        timeout=timeout
    )


def wait_shell(page: Page):
    page.wait_for_load_state("domcontentloaded")
    page.locator(SEL_SHELL_HEADER).wait_for(state="visible", timeout=60_000)


def login_se_necessario(page: Page, user: str, pwd: str):
    if page.locator(SEL_USER).count() == 0:
        return
    page.fill(SEL_USER, user)
    page.fill(SEL_PASS, pwd)
    page.click(SEL_LOGIN_BTN)
    wait_shell(page)


def preencher_campo(page: Page, seletor: str, valor: str, pressionar_enter=False):
    campo = page.locator(seletor).first
    campo.wait_for(state="visible", timeout=TIMEOUT)
    campo.click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    campo.type(str(valor))
    if pressionar_enter:
        page.keyboard.press("Enter")
        page.wait_for_timeout(300)


def clicar_elemento(page: Page, seletor: str, timeout=TIMEOUT):
    el = page.locator(seletor).first
    el.wait_for(state="visible", timeout=timeout)
    el.scroll_into_view_if_needed()
    el.click()
    page.wait_for_timeout(500)


def capturar_resultado(page: Page) -> str:
    """Tenta capturar mensagem de retorno do SAP após Simular ou Lançar."""
    try:
        t = page.locator(SEL_TOAST)
        if t.count() > 0:
            return t.first.inner_text(timeout=3000).strip()
    except:
        pass
    try:
        dialog = page.locator("[role='alertdialog'], .sapMMessageBox, .sapMDialog")
        if dialog.count() > 0:
            return dialog.first.inner_text(timeout=2000).strip()
    except:
        pass
    return "sem mensagem capturada"


def preencher_tipo_documento(page: Page):
    try:
        campo = page.locator(SEL_TIPO_DOC).first
        campo.wait_for(state="visible", timeout=TIMEOUT)
        campo.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        campo.type("KP")
        page.keyboard.press("Enter")
        page.wait_for_timeout(800)
        wait_busy_settle(page)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)
        wait_busy_settle(page)
    except Exception as e:
        raise Exception(f"Falha ao preencher Tipo de Documento com 'KP': {e}")


def texto_celula_por_colid(linha, colid: str) -> str:
    """Retorna o texto de uma célula específica da linha pelo data-sap-ui-colid."""
    try:
        cel = linha.locator(f"td[data-sap-ui-colid='{colid}']").first
        return cel.inner_text(timeout=500).strip()
    except:
        return ""


def pick_env():
    while True:
        resp = input("Ambiente (QAS/PRD): ").strip().upper()
        if resp == "QAS":
            return "QAS", URLQAS_HOME, URLQAS_APP, PROFILE_DIR_QAS
        if resp == "PRD":
            confirm = input("Digite CONFIRMAR PRD para continuar: ").strip().upper()
            if confirm != "CONFIRMAR PRD":
                sys.exit(0)
            return "PRD", URLPRD_HOME, URLPRD_APP, PROFILE_DIR_PRD

def clicar_ok_se_existir(page: Page, timeout=5000):
    """
    Clica no botão OK de popup de confirmação, se aparecer.
    Não quebra o fluxo caso não exista.
    """
    try:
        btn_ok = page.locator(SEL_BTN_OK_CONFIRMACAO)
        if btn_ok.count() > 0:
            btn_ok.first.wait_for(state="visible", timeout=timeout)
            btn_ok.first.click()
            page.wait_for_timeout(500)
            wait_busy_settle(page)
    except:
        pass

# ============================================================
# FLUXO PRINCIPAL DE LANÇAMENTO
# ============================================================
def preencher_lancamento(page: Page, row: dict) -> tuple[str, dict]:
    """
    Executa os passos comuns após entrar na tela de compensação.
    Retorna (mensagem_sap, detalhes_log).
    """
    tipo         = str(row["Tipo"]).strip().upper()
    valor        = normalizar_valor(row["Valor de crédito/débito"])
    conta_gl     = str(row["Conta do razão"]).strip()
    atribuicao   = str(row["Atribuição"]).strip()
    centro_lucro = str(row["Centro de lucro"]).strip()
    referencia   = str(row["Referência"]).strip()
    vencimento   = str(row["Data de vencimento"]).strip()

    detalhes = {
        "tipo_lancamento":            "Débito" if tipo == "DEVOLUÇÃO" else "Crédito",
        "conta_gl":                   conta_gl,
        "valor_preenchido":           valor,
        "atribuicao":                 atribuicao,
        "centro_lucro":               centro_lucro,
        "referencia":                 referencia,
        "vencimento":                 vencimento,
        "valor_on_account_calculado": "não capturado",
        "mensagem_sap":               "",
        "modo":                       "SIMULADO" if MODO_TESTE else "LANÇADO",
        "linha_grid_selecionada":     "",
        "criterio_match":             "",
    }

    # ── 1. Aba "Lançar em conta do Razão" ────────────────────
    clicar_elemento(page, SEL_ABA_RAZAO)
    wait_busy_settle(page)

    # ── 2. Conta GL ──────────────────────────────────────────
    preencher_campo(page, SEL_CONTA_GL, conta_gl, pressionar_enter=True)
    wait_busy_settle(page)

    # ── 3. Valor ─────────────────────────────────────────────
    if tipo == "DEVOLUÇÃO":
        preencher_campo(page, SEL_VALOR_DEB, valor, pressionar_enter=True)
    else:
        preencher_campo(page, SEL_VALOR_CRED, valor, pressionar_enter=True)
    wait_busy_settle(page)

    # ── 4. Expande painel GL ──────────────────────────────────
    clicar_elemento(page, SEL_EXPAND_GL)
    wait_busy_settle(page)

    # ── 5. Atribuição + Centro de lucro ───────────────────────
    preencher_campo(page, SEL_ATRIBUICAO, atribuicao)
    preencher_campo(page, SEL_CENTRO_LUCRO, centro_lucro, pressionar_enter=True)
    wait_busy_settle(page)

    # ── 6. Aba "Lançar em conta (valor BRL)" ─────────────────
    clicar_elemento(page, SEL_ABA_ON_ACCOUNT)
    wait_busy_settle(page)

    # ── 7. Captura valor calculado pelo SAP + dblclick ────────
    try:
        val_calculado = page.locator(SEL_ON_ACCOUNT_VALOR).first.input_value()
        detalhes["valor_on_account_calculado"] = val_calculado
    except:
        pass

    campo_valor = page.locator(SEL_ON_ACCOUNT_VALOR).first
    campo_valor.wait_for(state="visible", timeout=TIMEOUT)
    campo_valor.dblclick()
    page.wait_for_timeout(400)
    wait_busy_settle(page)

    # ── 8. Expande painel "Lançar em conta" ───────────────────
    clicar_elemento(page, SEL_EXPAND_ON_ACCOUNT)
    wait_busy_settle(page)

    # ── 9. Referência + Vencimento ────────────────────────────
    preencher_campo(page, SEL_REFERENCIA, referencia)
    preencher_campo(page, SEL_VENCIMENTO, vencimento, pressionar_enter=True)
    wait_busy_settle(page)

    # ── 10. Simular ou Lançar ─────────────────────────────────
    if MODO_TESTE:
        clicar_elemento(page, SEL_BTN_SIMULAR)
        wait_busy_settle(page)
        clicar_ok_se_existir(page)
        page.wait_for_timeout(DELAY_SIMULACAO)

    else:
        # 1️⃣ Sempre simula primeiro
        clicar_elemento(page, SEL_BTN_SIMULAR)
        wait_busy_settle(page)
        clicar_ok_se_existir(page)

        # 2️⃣ Aguarda tela do documento (onde aparece o novo botão Lançar)
        page.locator(SEL_BTN_LANCAR_FINAL).first.wait_for(state="visible", timeout=TIMEOUT)

        # 3️⃣ Clica no Lançar final
        clicar_elemento(page, SEL_BTN_LANCAR_FINAL)
        wait_busy_settle(page)
        clicar_ok_se_existir(page)

    msg = capturar_resultado(page)
    detalhes["mensagem_sap"] = msg
    return msg, detalhes


# ============================================================
# FLUXO: PARTIDA EM ABERTO
# ============================================================
def processar_aberto(page: Page, row: dict, app_url_template: str) -> tuple[str, dict]:
    empresa    = str(row.get("Empresa", EMPRESA_PADRAO)).strip()
    bp         = str(row["Fornecedor"]).strip()
    doc_sap    = str(row["Lançamento contábil"]).strip()
    referencia = str(row["Referência"]).strip()  # valor único usado para buscar na grid

    url = app_url_template.format(empresa=empresa, bp=bp)
    page.goto(url, wait_until="domcontentloaded", timeout=60_000)
    wait_busy_settle(page)

    # ── Preenche Tipo de Documento com "KP" ──────────────────
    preencher_tipo_documento(page)

    # ── Pesquisa o documento ──────────────────────────────────
    page.locator(SEL_CAMPO_BUSCA).wait_for(state="visible", timeout=TIMEOUT)
    preencher_campo(page, SEL_CAMPO_BUSCA, doc_sap, pressionar_enter=True)
    wait_busy_settle(page)

    linhas = page.locator("tr[data-sap-ui-rowindex]")
    linhas.first.wait_for(state="visible", timeout=TIMEOUT)

    btn_compensar   = None
    texto_linha_log = "não capturado"
    criterio_match  = "não encontrado"

    for i in range(linhas.count()):
        linha = linhas.nth(i)
        texto = linha.inner_text()

        # Linha precisa conter o número do documento
        if doc_sap not in texto:
            continue

        # 1ª tentativa: valor da coluna Referência da planilha
        #               confrontado com a coluna Atribuição da grid
        cel_atribuicao = texto_celula_por_colid(linha, COLID_ATRIBUICAO)
        if referencia and referencia in cel_atribuicao:
            btn_compensar   = linha.locator("bdi:has-text('Compensar')")
            texto_linha_log = texto.replace("\n", " | ").strip()
            criterio_match  = f"atribuição (grid) = '{cel_atribuicao}'"
            break

        # 2ª tentativa: valor da coluna Referência da planilha
        #               confrontado com a coluna Referência da grid
        cel_referencia = texto_celula_por_colid(linha, COLID_REFERENCIA)
        if referencia and referencia in cel_referencia:
            btn_compensar   = linha.locator("bdi:has-text('Compensar')")
            texto_linha_log = texto.replace("\n", " | ").strip()
            criterio_match  = f"referência (grid) = '{cel_referencia}'"
            break

    if btn_compensar is None or btn_compensar.count() == 0:
        raise Exception(
            f"Valor '{referencia}' não encontrado nas colunas Atribuição nem Referência "
            f"da grid para o documento {doc_sap} | BP {bp}"
        )

    btn_compensar.scroll_into_view_if_needed()
    btn_compensar.click()
    page.wait_for_timeout(600)
    wait_busy_settle(page)

    msg, detalhes = preencher_lancamento(page, row)
    detalhes["linha_grid_selecionada"] = texto_linha_log
    detalhes["criterio_match"]         = criterio_match
    return msg, detalhes


# ============================================================
# FLUXO: PARTIDA COMPENSADA (nova partida direta)
# ============================================================
def processar_compensado(page: Page, row: dict, app_url_template: str) -> tuple[str, dict]:
    empresa = str(row.get("Empresa", EMPRESA_PADRAO)).strip()
    bp      = str(row["Fornecedor"]).strip()

    url = app_url_template.format(empresa=empresa, bp=bp)
    page.goto(url, wait_until="domcontentloaded", timeout=60_000)
    wait_busy_settle(page)

    # ── Preenche Tipo de Documento com "KP" ──────────────────
    preencher_tipo_documento(page)

    try:
        page.locator(SEL_ABA_RAZAO).wait_for(state="visible", timeout=TIMEOUT)
    except PWTimeout:
        page.locator(SEL_CAMPO_BUSCA).wait_for(state="visible", timeout=TIMEOUT)

    msg, detalhes = preencher_lancamento(page, row)
    detalhes["linha_grid_selecionada"] = "N/A — partida compensada, lançamento direto"
    detalhes["criterio_match"]         = "N/A"
    return msg, detalhes


# ============================================================
# MAIN
# ============================================================
def main():
    xlsx = Path(EXCEL_PATH)
    if not xlsx.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {xlsx}")

    df = pd.read_excel(xlsx, dtype=str).fillna("")
    print(f"Total de linhas: {len(df)}")

    for col in ["Resultado", "Mensagem SAP", "Modo", "Conta GL usada",
                "Tipo lançamento", "Valor preenchido", "Valor on account calculado",
                "Referência usada", "Vencimento usado",
                "Linha grid selecionada", "Critério match", "Data/Hora"]:
        if col not in df.columns:
            df[col] = ""

    if input("Deseja continuar? (S/N): ").strip().upper() != "S":
        return

    env, home_url, app_url_template, profile_dir = pick_env()

    sap_user, sap_pass = "", ""
    if env == "QAS":
        sap_user = input("Usuário SAP: ").strip()
        sap_pass = input("Senha SAP: ").strip()
    else:
        print("PRD — SSO ativo, autentique manualmente se solicitado.")

    out = Path(OUTPUT_PATH)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            channel="chrome",
            slow_mo=SLOW_MO,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
        )

        page = context.new_page()
        page.set_viewport_size({"width": 1920, "height": 1080})

        page.goto(home_url, wait_until="domcontentloaded", timeout=60_000)
        if env == "QAS":
            login_se_necessario(page, sap_user, sap_pass)
        wait_shell(page)

        for idx, row in df.iterrows():
            fornecedor = str(row["Fornecedor"]).strip()
            doc_sap    = str(row.get("Lançamento contábil", "")).strip()
            status_doc = str(row["Status"]).strip().upper()
            tipo       = str(row["Tipo"]).strip()

            print(f"\n[Linha {idx+2}] Fornecedor {fornecedor} | Doc {doc_sap} | {status_doc} | {tipo}")

            resultado = ""
            msg       = ""
            detalhes  = {}

            try:
                if status_doc == "ABERTO":
                    msg, detalhes = processar_aberto(page, row, app_url_template)
                    resultado = "SUCESSO" + (" (SIMULADO)" if MODO_TESTE else "")

                elif status_doc == "COMPENSADO":
                    msg, detalhes = processar_compensado(page, row, app_url_template)
                    resultado = "SUCESSO" + (" (SIMULADO)" if MODO_TESTE else "")

                else:
                    msg       = f"Status desconhecido: '{status_doc}' — esperado ABERTO ou COMPENSADO"
                    resultado = "IGNORADO"

                print(f"  → {resultado}: {msg}")

            except Exception as e:
                resultado = "ERRO"
                msg       = str(e)[:500]
                detalhes  = {}
                print(f"  ✗ ERRO: {msg}")
                try:
                    page.goto(home_url, wait_until="domcontentloaded", timeout=30_000)
                    wait_shell(page)
                except:
                    pass

            # ── Grava log detalhado ───────────────────────────
            df.at[idx, "Resultado"]                  = resultado
            df.at[idx, "Mensagem SAP"]               = detalhes.get("mensagem_sap", msg)
            df.at[idx, "Modo"]                       = detalhes.get("modo", "SIMULADO" if MODO_TESTE else "LANÇADO")
            df.at[idx, "Conta GL usada"]             = detalhes.get("conta_gl", "")
            df.at[idx, "Tipo lançamento"]            = detalhes.get("tipo_lancamento", "")
            df.at[idx, "Valor preenchido"]           = detalhes.get("valor_preenchido", "")
            df.at[idx, "Valor on account calculado"] = detalhes.get("valor_on_account_calculado", "")
            df.at[idx, "Referência usada"]           = detalhes.get("referencia", "")
            df.at[idx, "Vencimento usado"]           = detalhes.get("vencimento", "")
            df.at[idx, "Linha grid selecionada"]     = detalhes.get("linha_grid_selecionada", "")
            df.at[idx, "Critério match"]             = detalhes.get("criterio_match", "")
            df.at[idx, "Data/Hora"]                  = now_str()
            df.to_excel(out, index=False)

        context.close()

    print(f"\nFinalizado. Log salvo em: {out}")


if __name__ == "__main__":
    main()