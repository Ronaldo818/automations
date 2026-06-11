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
EXCEL_PATH  = r"C:\Users\junio\OneDrive\Área de Trabalho\Documentos\Scripts Github\automations\Planilhas\Partidas.xlsx"
OUTPUT_PATH = r"C:\Users\junio\OneDrive\Área de Trabalho\Documentos\Scripts Github\automations\Planilhas\Partidas_logs.xlsx"

URLQAS_HOME = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Shell-home"
URLPRD_HOME = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Shell-home"

# URL base da transação (sem empresa/BP — montamos o deep link por linha)
URLQAS_APP  = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Supplier-clearOpenItems&/clearing/true/{empresa}/{bp}/undefined/undefined/undefined/undefined"
URLPRD_APP  = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Supplier-clearOpenItems&/clearing/true/{empresa}/{bp}/undefined/undefined/undefined/undefined"

PROFILE_DIR_QAS = str(Path.cwd() / "pw_profile_qas")
PROFILE_DIR_PRD = str(Path.cwd() / "pw_profile_prd")

# True  → clica em Simular (seguro para validar)
# False → clica em Lançar (executa de verdade)
MODO_TESTE = True

EMPRESA_PADRAO = "2000"

TIMEOUT = 60_000   # ms — aguarda elementos
SLOW_MO = 400      # ms — pausa entre ações

# Tempo de pausa após Simular para validação visual (ms)
# 5000 = 5 segundos — ajuste conforme necessário
DELAY_SIMULACAO = 1000000000_000

# ============================================================
# SELETORES
# ============================================================
SEL_SHELL_HEADER = "#shell-header"
SEL_USER         = "#USERNAME_FIELD-inner"
SEL_PASS         = "#PASSWORD_FIELD-inner"
SEL_LOGIN_BTN    = "#LOGIN_LINK"

# Tela inicial — botão "Compensar partidas em aberto"
SEL_BTN_COMPENSAR_PARTIDAS = "bdi#application-Supplier-clearOpenItems-component---PaymentListView--buttonAssignOpenItems-BDI-content"

# Popup de seleção empresa/BP
SEL_INPUT_EMPRESA   = "#openItemAssSelCompanyCode-inner"
SEL_INPUT_BP        = "#openItemAssSelAccountID-inner"
SEL_BTN_OK_POPUP    = "bdi#openItemsAssSelDialogOKButton-BDI-content"

# Tela de partidas — campo de busca
SEL_CAMPO_BUSCA     = "#application-Supplier-clearOpenItems-component---ClearingView--openItemSearch-I"

# Botão "Compensar" em cada linha da grid (o __clone muda — usamos o texto)
SEL_BTN_COMPENSAR_LINHA = "bdi:has-text('Compensar')"

# Aba "Lançar em conta do Razão"
SEL_ABA_RAZAO = "#application-Supplier-clearOpenItems-component---ClearingView--tabChargeOffDiff-text"

# Campos GL
SEL_CONTA_GL    = "#application-Supplier-clearOpenItems-component---ClearingView--glItemGLAccount-application-Supplier-clearOpenItems-component---ClearingView--glItems-0-input-inner"
SEL_VALOR_DEB   = "#application-Supplier-clearOpenItems-component---ClearingView--glItemDebitAmount-application-Supplier-clearOpenItems-component---ClearingView--glItems-0-input-inner"
SEL_VALOR_CRED  = "#application-Supplier-clearOpenItems-component---ClearingView--glItemCreditAmount-application-Supplier-clearOpenItems-component---ClearingView--glItems-0-input-inner"

# Botão expandir painel GL (ícone de seta)
SEL_EXPAND_GL   = "#application-Supplier-clearOpenItems-component---ClearingView--itemPanel-application-Supplier-clearOpenItems-component---ClearingView--glItems-0-expandButton-img"

# Campos dentro do painel GL expandido
SEL_ATRIBUICAO = "input[aria-labelledby*='codingBlockDetailsGroupElement_AssignmentReference-label']"
SEL_CENTRO_LUCRO = "input[aria-labelledby*='codingBlockGroupElement_ProfitCenter-label']"

# Aba "Lançar em conta (valor BRL)"
SEL_ABA_ON_ACCOUNT  = "#application-Supplier-clearOpenItems-component---ClearingView--tabOnAccount-text"

# Campo de valor na aba "Lançar em conta"
SEL_ON_ACCOUNT_VALOR = "#application-Supplier-clearOpenItems-component---ClearingView--onAccountItemCreditAmount-application-Supplier-clearOpenItems-component---ClearingView--onAccountItems-0-input-inner"

# Botão expandir painel "Lançar em conta"
SEL_EXPAND_ON_ACCOUNT = "#application-Supplier-clearOpenItems-component---ClearingView--APARItemPanel-application-Supplier-clearOpenItems-component---ClearingView--onAccountItems-0-expandButton-img"

# Campos dentro do painel "Lançar em conta" expandido
SEL_REFERENCIA      = "#application-Supplier-clearOpenItems-component---ClearingView--OnAccountInputAssignmentReference-application-Supplier-clearOpenItems-component---ClearingView--onAccountItems-0-input-inner"
SEL_VENCIMENTO      = "#application-Supplier-clearOpenItems-component---ClearingView--OnAccountInputDueCalculationBaseDate-application-Supplier-clearOpenItems-component---ClearingView--onAccountItems-0-datePicker-inner"

# Botões finais
SEL_BTN_SIMULAR = "bdi#application-Supplier-clearOpenItems-component---ClearingView--simulateMenuButton-internalSplitBtn-textButton-BDI-content"
SEL_BTN_LANCAR  = "bdi#application-Supplier-clearOpenItems-component---ClearingView--buttonPost-BDI-content"

# Toast de confirmação
SEL_TOAST = ".sapMMessageToast"

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
        v = v.replace(".", "")                      # remove milhar
    elif "." in v:
        v = v.replace(".", ",")                     # ponto decimal → vírgula
    parts = v.split(",")
    inteiro = parts[0].lstrip("0") or "0"
    dec = (parts[1] + "00")[:2] if len(parts) > 1 else "00"
    return f"{inteiro},{dec}"


def wait_busy_settle(page: Page, timeout=120_000):
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
    page.locator(SEL_SHELL_HEADER).wait_for(state="visible", timeout=120_000)


def login_se_necessario(page: Page, user: str, pwd: str):
    if page.locator(SEL_USER).count() == 0:
        return
    page.fill(SEL_USER, user)
    page.fill(SEL_PASS, pwd)
    page.click(SEL_LOGIN_BTN)
    wait_shell(page)


def preencher_campo(page: Page, seletor: str, valor: str, pressionar_enter=False):
    """Limpa e preenche um campo de input."""
    campo = page.locator(seletor)
    campo.wait_for(state="visible", timeout=TIMEOUT)
    campo.click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    campo.type(str(valor))
    if pressionar_enter:
        page.keyboard.press("Enter")
        page.wait_for_timeout(300)


def clicar_elemento(page: Page, seletor: str, timeout=TIMEOUT):
    """Aguarda o elemento ficar visível e clica."""
    el = page.locator(seletor).first
    el.wait_for(state="visible", timeout=timeout)
    el.scroll_into_view_if_needed()
    el.click()
    page.wait_for_timeout(500)


def try_get_toast(page: Page) -> str:
    try:
        t = page.locator(SEL_TOAST)
        if t.count() > 0:
            return t.first.inner_text(timeout=3000).strip()
    except:
        pass
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


# ============================================================
# FLUXO PRINCIPAL DE LANÇAMENTO
# (compartilhado entre ABERTO e COMPENSADO a partir da aba GL)
# ============================================================
def preencher_lancamento(page: Page, row: dict):
    """
    Executa os passos comuns após entrar na tela de compensação:
    - Aba GL → conta + valor
    - Expande painel → atribuição + centro de lucro
    - Aba "Lançar em conta" → atualiza valor
    - Expande painel → referência + vencimento
    - Simula ou Lança
    """
    tipo         = str(row["Tipo"]).strip().upper()          # FATURA | DEVOLUÇÃO
    valor        = normalizar_valor(row["Valor de crédito/débito"])
    conta_gl     = str(row["Conta do razão"]).strip()
    atribuicao   = str(row["Atribuição"]).strip()
    centro_lucro = str(row["Centro de lucro"]).strip()
    referencia   = str(row["Referência"]).strip()
    vencimento   = str(row["Data de vencimento"]).strip()

    # ── 1. Aba "Lançar em conta do Razão" ────────────────────
    clicar_elemento(page, SEL_ABA_RAZAO)
    wait_busy_settle(page)

    # ── 2. Conta GL ──────────────────────────────────────────
    preencher_campo(page, SEL_CONTA_GL, conta_gl, pressionar_enter=True)
    wait_busy_settle(page)

    # ── 3. Valor (débito ou crédito conforme tipo) ────────────
    if tipo == "DEVOLUÇÃO":
        preencher_campo(page, SEL_VALOR_DEB, valor, pressionar_enter=True)
    else:  # FATURA
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

    # ── 7. Clica duas vezes no campo de valor para atualizar ──
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
        page.wait_for_timeout(DELAY_SIMULACAO)  # pausa para validar visualmente
    else:
        clicar_elemento(page, SEL_BTN_LANCAR)
        wait_busy_settle(page)

    toast = try_get_toast(page)
    return toast or "Ação executada (sem toast capturado)"


# ============================================================
# FLUXO: PARTIDA EM ABERTO
# ============================================================
def processar_aberto(page: Page, row: dict, app_url_template: str):
    empresa  = str(row.get("Empresa", EMPRESA_PADRAO)).strip()
    bp       = str(row["Fornecedor"]).strip()
    doc_sap  = str(row["Lançamento contábil"]).strip()

    # Navega direto pelo deep link (empresa + BP já na URL)
    url = app_url_template.format(empresa=empresa, bp=bp)
    page.goto(url, wait_until="domcontentloaded", timeout=120_000)
    wait_busy_settle(page)

    # Aguarda campo de busca da grid
    page.locator(SEL_CAMPO_BUSCA).wait_for(state="visible", timeout=TIMEOUT)

    # Pesquisa o número do documento
    preencher_campo(page, SEL_CAMPO_BUSCA, doc_sap, pressionar_enter=True)
    wait_busy_settle(page)

    # Localiza e clica no botão "Compensar" da linha correspondente
    # A grid pode ter múltiplas linhas — busca a linha que contém o número do doc
    linhas = page.locator("tr[data-sap-ui-rowindex]")
    linhas.first.wait_for(state="visible", timeout=TIMEOUT)

    btn_compensar = None
    for i in range(linhas.count()):
        linha = linhas.nth(i)
        if doc_sap in linha.inner_text():
            btn_compensar = linha.locator("bdi:has-text('Compensar')")
            break

    if btn_compensar is None or btn_compensar.count() == 0:
        raise Exception(f"Documento {doc_sap} não encontrado na grid para o BP {bp}.")

    btn_compensar.scroll_into_view_if_needed()
    btn_compensar.click()
    page.wait_for_timeout(600)
    wait_busy_settle(page)

    # A partir daqui o fluxo é idêntico ao COMPENSADO
    return preencher_lancamento(page, row)


# ============================================================
# FLUXO: PARTIDA COMPENSADA (nova partida direta)
# ============================================================
def processar_compensado(page: Page, row: dict, app_url_template: str):
    empresa = str(row.get("Empresa", EMPRESA_PADRAO)).strip()
    bp      = str(row["Fornecedor"]).strip()

    # Navega direto pelo deep link
    url = app_url_template.format(empresa=empresa, bp=bp)
    page.goto(url, wait_until="domcontentloaded", timeout=120_000)
    wait_busy_settle(page)

    # Aguarda a tela carregar (campo de busca ou aba GL visível)
    try:
        page.locator(SEL_ABA_RAZAO).wait_for(state="visible", timeout=TIMEOUT)
    except PWTimeout:
        # Fallback: aguarda qualquer elemento principal da tela
        page.locator(SEL_CAMPO_BUSCA).wait_for(state="visible", timeout=TIMEOUT)

    # Vai direto para o lançamento (sem buscar documento nem clicar Compensar)
    return preencher_lancamento(page, row)


# ============================================================
# MAIN
# ============================================================
def main():
    xlsx = Path(EXCEL_PATH)
    if not xlsx.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {xlsx}")

    df = pd.read_excel(xlsx, dtype=str).fillna("")
    print(f"Total de linhas: {len(df)}")

    for col in ["Status", "Resultado", "Mensagem", "Data/Hora"]:
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

        # Login
        page.goto(home_url, wait_until="domcontentloaded", timeout=120_000)
        if env == "QAS":
            login_se_necessario(page, sap_user, sap_pass)
        wait_shell(page)

        # Loop de linhas
        for idx, row in df.iterrows():
            fornecedor = str(row["Fornecedor"]).strip()
            doc_sap    = str(row.get("Lançamento contábil", "")).strip()
            status_doc = str(row["Status"]).strip().upper()
            tipo       = str(row["Tipo"]).strip()

            print(f"\n[Linha {idx+2}] Fornecedor {fornecedor} | Doc {doc_sap} | {status_doc} | {tipo}")

            try:
                if status_doc == "ABERTO":
                    msg = processar_aberto(page, row, app_url_template)
                    resultado = "SUCESSO" + (" (SIMULADO)" if MODO_TESTE else "")

                elif status_doc == "COMPENSADO":
                    msg = processar_compensado(page, row, app_url_template)
                    resultado = "SUCESSO" + (" (SIMULADO)" if MODO_TESTE else "")

                else:
                    msg = f"Status desconhecido: '{status_doc}' — esperado ABERTO ou COMPENSADO"
                    resultado = "IGNORADO"

                print(f"  → {resultado}: {msg}")

            except Exception as e:
                resultado = "ERRO"
                msg = str(e)[:500]
                print(f"  ✗ ERRO: {msg}")

                # Tenta voltar para uma tela limpa antes da próxima linha
                try:
                    page.goto(home_url, wait_until="domcontentloaded", timeout=30_000)
                    wait_shell(page)
                except:
                    pass

            df.at[idx, "Resultado"]  = resultado
            df.at[idx, "Mensagem"]   = msg
            df.at[idx, "Data/Hora"]  = now_str()
            df.to_excel(out, index=False)  # salva o log a cada linha processada

        context.close()

    print(f"\nFinalizado. Log salvo em: {out}")


if __name__ == "__main__":
    main()