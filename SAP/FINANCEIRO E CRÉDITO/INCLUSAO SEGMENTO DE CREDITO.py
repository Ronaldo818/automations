from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PwTimeoutError


# ========= CONFIG =========
EXCEL_PATH = r"C:\Users\ronaldo.gontijo\Downloads\segmentos.xlsx"

URL = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=310&sap-language=PT#BusinessPartner-manageCreditAccounts&/"
PROFILE_DIR = str(Path.cwd() / "pw_profile")

SEL_BP_FILTER = "[id*='BusinessPartner-inner']"
SEL_GO_BTN = "[id*='btnGo']"
SEL_EDIT_BTN = "[id*='edit']"
SEL_CREDIT_LIMIT_INPUT = "[id*='CreditLimitAmount']"


# ========= UTILS =========
def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def wait_busy(page):
    page.wait_for_function("""
    () => {
        const els = document.querySelectorAll('.sapUiLocalBusyIndicator');
        return [...els].every(e => e.style.display === 'none');
    }
    """)


def deep_link_segment(app_url, bp, seg):
    return f"{app_url}CrdtMBusinessPartner(BusinessPartner='{bp}',IsActiveEntity=false)/to_CreditMgmtAccountTP(BusinessPartner='{bp}',CreditSegment='{seg}',IsActiveEntity=false)/"


def normalize(v):
    v = str(v).replace(".", "").replace(",", ".")
    return f"{float(v):.2f}".replace(".", ",")


# ========= SEGMENT =========
def segment_exists(page, seg):
    linhas = page.locator("[role='row']")
    return linhas.filter(has_text=seg).count() > 0


def create_segment(page, seg):

    btn = page.locator("bdi:has-text('Criar')").first
    btn.click()

    dlg = page.locator(".sapMDialog").last
    dlg.wait_for()

    inp = dlg.locator("input").first
    inp.fill(seg)
    page.keyboard.press("Enter")

    dlg.locator("button").filter(has_text="OK").click()

    page.wait_for_timeout(1000)
    wait_busy(page)


# ========= MAIN =========
def main():

    df = pd.read_excel(EXCEL_PATH).fillna("")
    df["Status"] = ""
    df["Mensagem"] = ""

    with sync_playwright() as p:

        ctx = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False
        )

        page = ctx.new_page()
        page.goto(URL)

        input("Faça login e pressione ENTER...")

        for i, row in df.iterrows():

            bp = str(row["Cliente"]).strip()
            seg = str(row["Segmento"]).strip()
            limite = normalize(row["Limite"])

            try:
                # filtro
                f = page.locator(SEL_BP_FILTER)
                f.fill(bp)
                page.keyboard.press("Enter")
                page.locator(SEL_GO_BTN).click()
                page.wait_for_timeout(1000)

                # entra BP
                page.click(f"text={bp}")
                page.wait_for_timeout(2000)

                # edit
                try:
                    page.locator(SEL_EDIT_BTN).click()
                except PwTimeoutError:
                    df.at[i, "Status"] = "PULADO"
                    continue

                wait_busy(page)

                # verifica segmento
                if segment_exists(page, seg):

                    df.at[i, "Status"] = "OK"
                    df.at[i, "Mensagem"] = "Já existe"
                    page.go_back()
                    continue

                # cria segmento
                create_segment(page, seg)

                # entra deep link
                page.goto(deep_link_segment(URL, bp, seg))
                wait_busy(page)

                # define limite
                inp = page.locator(SEL_CREDIT_LIMIT_INPUT)
                inp.wait_for()

                inp.fill(limite)
                page.keyboard.press("Enter")

                wait_busy(page)

                # salvar
                page.locator("button:has-text('Salvar')").click()
                wait_busy(page)

                df.at[i, "Status"] = "OK"
                df.at[i, "Mensagem"] = "Criado"

                page.goto(URL)

            except Exception as e:
                df.at[i, "Status"] = "ERRO"
                df.at[i, "Mensagem"] = str(e)
                page.goto(URL)

        df.to_excel("resultado.xlsx", index=False)


if __name__ == "__main__":
    main()