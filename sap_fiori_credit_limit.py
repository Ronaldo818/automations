from __future__ import annotations

import sys
import re
from datetime import datetime
from pathlib import Path
from typing import Tuple

import pandas as pd
from playwright.sync_api import sync_playwright


# ==========
# INPUTS
# ==========
EXCEL_PATH = r"c:\Users\ronaldo.gontijo\Downloads\Itens (71).xlsx"

URLPRD_HOME = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Shell-home"
URLQAS_HOME = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=310&sap-language=PT#Shell-home"

PRD_APP = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#BusinessPartner-manageCreditAccounts&/"
QAS_APP = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=310&sap-language=PT#BusinessPartner-manageCreditAccounts&/"

PROFILE_DIR_PRD = str(Path.cwd() / "pw_profile_prd")
PROFILE_DIR_QAS = str(Path.cwd() / "pw_profile_qas")

SEL_SHELL_HEADER = "#shell-header"

SEL_USER = "#USERNAME_FIELD-inner"
SEL_PASS = "#PASSWORD_FIELD-inner"
SEL_LOGIN_BTN = "#LOGIN_LINK"

SEL_BP_FILTER = "[id='fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ListReport.view.ListReport::CrdtMBusinessPartner--listReportFilter-filterItemControl_BASIC-BusinessPartner-inner']"
SEL_GO_BTN = "[id='fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ListReport.view.ListReport::CrdtMBusinessPartner--listReportFilter-btnGo']"

SEL_EDIT_BTN = "[id='fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ObjectPage.view.Details::CrdtMBusinessPartner--edit']"
SEL_CREDIT_LIMIT_INPUT = "[id='fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ObjectPage.view.Details::Account--_FieldGroup_Limit::CreditLimitAmount::Field-input-inner']"

SEL_APLICAR_BTN = "[id='fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ObjectPage.view.Details::Account--footerObjectPageBackTo']"
SEL_SAVE_BTN = "[id='fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ObjectPage.view.Details::CrdtMBusinessPartner--activate']"

SEL_TOAST = ".sapMMessageToast"
SHEET_NAME = 0


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_number_text(s: str) -> str:
    """
    Normaliza para o padrão que o SAP (PT-BR) aceita no input:
    - separador decimal: vírgula
    - sempre 2 casas decimais
    - sem separador de milhar (ou seja, "4246,50")
    Aceita entradas tipo:
      "4.246,50" "4246,5" "4246.5" "4246" "R$ 4.246,50"
    """
    if s is None:
        return ""

    s = str(s).strip().replace("\u00A0", " ").strip()
    s = re.sub(r"[^\d\.,\-]", "", s)  # mantém só dígitos, ponto, vírgula, sinal

    if s == "" or s == "-" or s == "," or s == ".":
        return ""

    # Caso 1: tem vírgula e ponto -> assume ponto milhar e vírgula decimal (pt-BR)
    if "," in s and "." in s:
        s = s.replace(".", "")  # remove milhar
        # garante só 1 vírgula (decimal)
        parts = s.split(",")
        s = parts[0] + "," + "".join(parts[1:])

    # Caso 2: só ponto -> assume ponto decimal e converte pra vírgula
    elif "." in s and "," not in s:
        parts = s.split(".")
        s = parts[0] + "," + "".join(parts[1:])  # primeira ocorrência vira decimal

    # Caso 3: só vírgula -> já está ok (só garantir 1 vírgula)
    elif "," in s:
        parts = s.split(",")
        s = parts[0] + "," + "".join(parts[1:])

    # Agora: garantir 2 casas decimais
    if "," in s:
        inteiro, dec = s.split(",", 1)
        dec = re.sub(r"\D", "", dec)
        dec = (dec + "00")[:2]
        s = f"{inteiro},{dec}"
    else:
        # inteiro puro
        s = f"{s},00"

    # remove zeros à esquerda do inteiro (mantendo "0")
    neg = s.startswith("-")
    if neg:
        s = s[1:]
    inteiro, dec = s.split(",", 1)
    inteiro = inteiro.lstrip("0") or "0"
    s = f"{inteiro},{dec}"
    return f"-{s}" if neg else s


def object_page_link(app_base: str, bp: str):
    return f"{app_base}CrdtMBusinessPartner(BusinessPartner='{bp}',IsActiveEntity=true)"


def deep_link_segmento(app_base: str, bp: str, segmento: str):
    return (
        f"{app_base}"
        f"CrdtMBusinessPartner(BusinessPartner='{bp}',IsActiveEntity=false)"
        f"/to_CreditMgmtAccountTP(BusinessPartner='{bp}',CreditSegment='{segmento}',IsActiveEntity=false)/"
    )


def pick_env() -> Tuple[str, str, str, str]:
    while True:
        resp = input("Ambiente (QAS/PRD): ").strip().upper()
        if resp == "QAS":
            return "QAS", URLQAS_HOME, QAS_APP, PROFILE_DIR_QAS
        if resp == "PRD":
            confirm = input("CONFIRMAR PRD: ").strip().upper()
            if confirm != "CONFIRMAR PRD":
                sys.exit(0)
            return "PRD", URLPRD_HOME, PRD_APP, PROFILE_DIR_PRD


def wait_shell(page):
    page.wait_for_load_state("domcontentloaded")
    page.locator(SEL_SHELL_HEADER).wait_for(state="visible", timeout=120000)


def login_qas_if_needed(page, user, pwd):
    if page.locator(SEL_USER).count() == 0:
        return
    page.fill(SEL_USER, user)
    page.fill(SEL_PASS, pwd)
    page.click(SEL_LOGIN_BTN)
    wait_shell(page)


def try_get_toast(page):
    try:
        t = page.locator(SEL_TOAST)
        if t.count() > 0:
            return t.first.inner_text(timeout=1500).strip()
    except:
        pass
    return ""


# ==========
# AJUSTES NOVOS (determinismo pós-gravar)
# ==========
def wait_busy_settle(page, timeout=120000):
    page.wait_for_function(
        """
        () => {
          const els = Array.from(document.querySelectorAll('.sapUiLocalBusyIndicator'));
          if (els.length === 0) return true;

          const isVisible = (el) => {
            const cs = window.getComputedStyle(el);
            if (!cs) return false;
            if (cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0') return false;
            if (el.offsetParent === null) return false;
            return true;
          };

          return els.every(el => !isVisible(el));
        }
        """,
        timeout=timeout
    )


def force_back_to_listreport(page, app_url: str, attempts: int = 5):
    last_err = None
    for _ in range(attempts):
        try:
            page.evaluate("u => window.location.replace(u);", app_url)
            page.wait_for_load_state("domcontentloaded", timeout=120000)
            page.locator(SEL_BP_FILTER).wait_for(state="visible", timeout=30000)
            return
        except Exception as e:
            last_err = e

        try:
            page.goto(app_url, wait_until="domcontentloaded", timeout=120000)
            page.locator(SEL_BP_FILTER).wait_for(state="visible", timeout=30000)
            return
        except Exception as e:
            last_err = e

        try:
            page.reload(wait_until="domcontentloaded", timeout=120000)
            page.locator(SEL_BP_FILTER).wait_for(state="visible", timeout=30000)
            return
        except Exception as e:
            last_err = e

        page.wait_for_timeout(800)

    raise RuntimeError(f"Não conseguiu voltar para a tela de filtro após {attempts} tentativas. Último erro: {last_err}")


# ==========
# UI5/DOM helpers
# ==========
def wait_ui5_core(page, timeout=120000):
    page.wait_for_function(
        "() => window.sap && sap.ui && sap.ui.getCore && sap.ui.getCore()",
        timeout=timeout
    )


def commit_field_change(page):
    page.keyboard.press("Enter")
    page.wait_for_timeout(100)
    page.keyboard.press("Tab")
    page.wait_for_timeout(150)


def ui5_set_input_value_from_dom_inner(page, dom_inner_id: str, value: str, timeout=120000):
    wait_ui5_core(page, timeout=timeout)
    page.evaluate(
        """
        ({domInnerId, val}) => {
          const core = sap.ui.getCore();
          const ctlId = domInnerId.endsWith("-inner") ? domInnerId.slice(0, -6) : domInnerId;
          const c = core.byId(ctlId);
          if (!c) throw new Error("UI5 Input control não encontrado: " + ctlId);

          if (typeof c.setValue === "function") c.setValue(val);
          if (typeof c.fireLiveChange === "function") c.fireLiveChange({ value: val, newValue: val });
          if (typeof c.fireChange === "function") c.fireChange({ value: val, newValue: val });
          if (typeof c.fireSubmit === "function") c.fireSubmit({ value: val });

          core.applyChanges();
        }
        """,
        {"domInnerId": dom_inner_id, "val": value}
    )


def robust_press_button(page, button_id: str, timeout=120000):
    wait_ui5_core(page, timeout=timeout)

    pressed = page.evaluate(
        """
        (id) => {
          const c = sap.ui.getCore().byId(id);
          if (!c) return { ok:false, step:"no-control" };

          const enabled = (typeof c.getEnabled === "function") ? c.getEnabled() : true;

          try {
            if (typeof c.firePress === "function") { c.firePress(); return { ok:true, step:"firePress", enabled }; }
          } catch (e) {}

          try {
            if (typeof c.fireTap === "function") { c.fireTap(); return { ok:true, step:"fireTap", enabled }; }
          } catch (e) {}

          return { ok:false, step:"no-fire", enabled };
        }
        """,
        button_id
    )

    if pressed and pressed.get("ok"):
        return

    page.evaluate(
        """
        (id) => {
          const el = document.getElementById(id);
          if (!el) throw new Error("DOM element do botão não encontrado: " + id);

          const opts = { bubbles: true, cancelable: true, view: window };
          el.dispatchEvent(new MouseEvent("mouseover", opts));
          el.dispatchEvent(new MouseEvent("mousedown", opts));
          el.dispatchEvent(new MouseEvent("mouseup", opts));
          el.dispatchEvent(new MouseEvent("click", opts));
        }
        """,
        button_id
    )


# ==========
# NOVO: salvar com retry + trata popups comuns
# ==========
def close_possible_dialogs(page):
    # fecha/confirmar popups comuns SAPUI5 (sem quebrar se não existir)
    for txt in ["OK", "Ok", "Fechar", "Close", "Sim", "Yes", "Continuar", "Continue"]:
        try:
            btn = page.locator(f"button:has-text('{txt}')")
            if btn.count() > 0 and btn.first.is_visible():
                btn.first.click(timeout=1000)
                page.wait_for_timeout(300)
        except Exception:
            pass


def ui5_wait_button_enabled(page, button_id: str, timeout=15000):
    # espera o UI5 dizer que está enabled (se tiver getEnabled)
    wait_ui5_core(page, timeout=timeout)
    page.wait_for_function(
        """
        (id) => {
          const c = sap.ui.getCore().byId(id);
          if (!c) return false;
          if (typeof c.getEnabled === "function") return c.getEnabled();
          return true;
        }
        """,
        button_id,
        timeout=timeout
    )


def save_with_retry(page, app_url: str, attempts: int = 5):
    last_err = None
    for _ in range(attempts):
        try:
            close_possible_dialogs(page)
            wait_busy_settle(page, timeout=120000)

            # tenta esperar habilitar um pouco (sem travar infinito)
            try:
                ui5_wait_button_enabled(
                    page,
                    "fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ObjectPage.view.Details::CrdtMBusinessPartner--activate",
                    timeout=15000
                )
            except Exception:
                pass

            robust_press_button(
                page,
                "fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ObjectPage.view.Details::CrdtMBusinessPartner--activate",
                timeout=120000
            )

            page.wait_for_timeout(600)
            wait_busy_settle(page, timeout=120000)
            close_possible_dialogs(page)

            # se conseguiu voltar pro filtro, ótimo
            force_back_to_listreport(page, app_url, attempts=3)
            return

        except Exception as e:
            last_err = e
            # dá uma respirada e tenta de novo
            page.wait_for_timeout(800)

    raise RuntimeError(f"Falha ao GRAVAR após {attempts} tentativas. Último erro: {last_err}")


def main():

    xlsx = Path(EXCEL_PATH)
    if not xlsx.exists():
        raise FileNotFoundError(xlsx)

    df = pd.read_excel(xlsx, sheet_name=SHEET_NAME, dtype=str).fillna("")
    print(f"Total de linhas encontradas: {len(df)}")

    if input("Deseja continuar? (S/N): ").upper() != "S":
        return

    for col in ["Limite_Anterior", "Status", "Mensagem", "Data/Hora"]:
        if col not in df.columns:
            df[col] = ""

    env, home_url, app_url, profile_dir = pick_env()

    sap_user = ""
    sap_pass = ""
    if env == "QAS":
        sap_user = input("User SAP: ")
        sap_pass = input("Senha SAP: ")
    else:
        print("PRD SSO - autentique manualmente se necessário.")

    out = xlsx.with_name(xlsx.stem + "_resultado.xlsx")
    df.to_excel(out, index=False)

    with sync_playwright() as p:

        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            channel="chrome",
            slow_mo=400,
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
            ],
        )

        context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        """)

        page = context.new_page()
        page.set_viewport_size({"width": 1920, "height": 1080})

        page.goto(home_url)
        if env == "QAS":
            login_qas_if_needed(page, sap_user, sap_pass)

        wait_shell(page)
        page.goto(app_url)

        for idx, row in df.iterrows():

            bp = row["Cliente"].strip()
            segmento = row["Segmento"].strip()
            novo_limite_raw = row["Novo_Limite"].strip()

            try:
                bp_filter = page.locator(SEL_BP_FILTER)
                bp_filter.wait_for(state="visible", timeout=120000)

                bp_filter.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                bp_filter.type(bp)
                page.keyboard.press("Enter")

                page.locator(SEL_GO_BTN).click()
                page.wait_for_timeout(1500)

                page.goto(object_page_link(app_url, bp), timeout=120000)

                page.wait_for_function("""
                () => {
                    const btn = document.querySelector("[id*='--edit']");
                    return btn && !btn.disabled;
                }
                """, timeout=120000)

                page.locator(SEL_EDIT_BTN).click()
                page.wait_for_timeout(2500)

                segment_hash = deep_link_segmento(app_url, bp, segmento).split("#")[1]
                page.evaluate("(h) => { window.location.hash = h; }", segment_hash)

                page.wait_for_function("""
                () => {
                    return document.querySelector("[id*='CreditLimitAmount']");
                }
                """, timeout=120000)

                limit_input = page.locator(SEL_CREDIT_LIMIT_INPUT)
                limit_input.wait_for(state="visible", timeout=120000)

                df.at[idx, "Limite_Anterior"] = limit_input.input_value()

                novo_limite = normalize_number_text(novo_limite_raw)

                limit_input.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                limit_input.type(novo_limite)
                page.keyboard.press("Enter")

                dom_inner_id = "fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ObjectPage.view.Details::Account--_FieldGroup_Limit::CreditLimitAmount::Field-input-inner"
                ui5_set_input_value_from_dom_inner(page, dom_inner_id, novo_limite, timeout=120000)
                commit_field_change(page)

                # APLICAR
                robust_press_button(
                    page,
                    "fin.fscm.cr.creditaccounts.manage::sap.suite.ui.generic.template.ObjectPage.view.Details::Account--footerObjectPageBackTo",
                    timeout=120000
                )
                page.wait_for_timeout(400)
                wait_busy_settle(page, timeout=120000)

                # GRAVAR com retry + dialogs
                save_with_retry(page, app_url, attempts=5)

                df.at[idx, "Status"] = "OK"
                df.at[idx, "Mensagem"] = try_get_toast(page)
                df.at[idx, "Data/Hora"] = now_str()
                df.at[idx, "Usuário"] = sap_user
                df.to_excel(out, index=False)

            except Exception as e:
                df.at[idx, "Status"] = "ERRO"
                df.at[idx, "Mensagem"] = str(e)[:400]
                df.at[idx, "Data/Hora"] = now_str()
                df.at[idx, "Usuário"] = sap_user
                df.to_excel(out, index=False)

        context.close()

    print("Finalizado:", out)

if __name__ == "__main__":
    main()
