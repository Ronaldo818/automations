from __future__ import annotations

import sys
import re
from datetime import datetime
from pathlib import Path
from typing import Tuple

import pandas as pd
from playwright.sync_api import sync_playwright


# ========== CONFIG ==========
EXCEL_PATH = r"C:\Users\ronaldo.gontijo\Downloads\Pedidos.xlsx"

URLPRD_HOME = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#Shell-home"
URLQAS_HOME = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=310&sap-language=PT#Shell-home"

PRD_APP = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#SalesDocument-change?sap-ui-tech-hint=GUI"
QAS_APP = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=310&sap-language=PT#SalesDocument-change?sap-ui-tech-hint=GUI"

PROFILE_DIR_PRD = str(Path.cwd() / "pw_profile_prd")
PROFILE_DIR_QAS = str(Path.cwd() / "pw_profile_qas")

SEL_SHELL_HEADER = "#shell-header"
SEL_USER = "#USERNAME_FIELD-inner"
SEL_PASS = "#PASSWORD_FIELD-inner"
SEL_LOGIN_BTN = "#LOGIN_LINK"


# ========== AUX ==========
def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


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


# ========== FRAME WEBGUI ==========
def get_webgui_frame(page):
    for _ in range(60):
        for f in page.frames:
            url = (f.url or "").lower()
            if "webgui" in url or "its" in url:
                return f
        page.wait_for_timeout(300)
    raise Exception("VA02 não carregou — iframe WebGUI não encontrado.")


# ========== MENSAGEM SAP ==========
def ler_mensagem_sap(frame) -> str:
    try:
        msg = frame.locator('//*[@id="wnd[0]/sbar_msg-txt"]')
        if msg.count() == 0:
            return ""
        texto = msg.get_attribute("title") or msg.inner_text()
        return texto.strip() if texto else ""
    except:
        return ""


def aguardar_mensagem_nova(frame, mensagem_anterior, pedido, timeout_ms=6000) -> str:
    inicio = datetime.now()

    while (datetime.now() - inicio).total_seconds() * 1000 < timeout_ms:
        msg = ler_mensagem_sap(frame)

        if not msg or msg == mensagem_anterior:
            frame.page.wait_for_timeout(200)
            continue

        msg_low = msg.lower()

        if (
            str(pedido) in msg
            or "não pode ser eliminado" in msg_low
            or "não existe no banco de dados" in msg_low
            or "foi arquivado" in msg_low
        ):
            return msg

        frame.page.wait_for_timeout(200)

    return ""


# ========== TELA DE PESQUISA ==========
def garantir_tela_pesquisa(page):
    frame = get_webgui_frame(page)
    try:
        campo = frame.locator("input[title='Documento de vendas']")
        if campo.count() > 0 and campo.get_attribute("readonly") is not None:
            page.keyboard.press("F3")
            page.wait_for_timeout(800)
            try:
                btn_nao = frame.locator("div[accesskey='N']")
                if btn_nao.count() > 0:
                    btn_nao.first.click()
                    page.wait_for_timeout(800)
            except:
                pass
    except:
        pass


# ========== PROCESSAR EXCLUSÃO ==========
def processar_exclusao_pedido(page, pedido):
    frame = get_webgui_frame(page)
    campo = frame.locator("input[title='Documento de vendas']")
    campo.wait_for(state="visible", timeout=20000)

    # Garantir campo editável
    for _ in range(20):
        if campo.get_attribute("readonly") is None:
            break
        page.wait_for_timeout(300)

    # Preencher pedido
    frame.fill("input[title='Documento de vendas']", str(pedido))
    page.wait_for_timeout(300)

    # ENTER
    frame.get_by_role("button", name=re.compile("Avançar|Entrada", re.I)).click()
    page.wait_for_timeout(600)

    # >>> NOVO: capturar erro ainda na pesquisa <<<
    mensagem_pesquisa = ler_mensagem_sap(frame)
    if mensagem_pesquisa:
        msg_low = mensagem_pesquisa.lower()
        if (
            "não existe no banco de dados" in msg_low
            or "foi arquivado" in msg_low
        ):
            return mensagem_pesquisa  # ✅ já eliminado / inexistente

    # Ignorar avisos iniciais (ex: considerar docs subsequentes)
    mensagem_base = mensagem_pesquisa

    # Verificar se documento abriu
    for _ in range(10):
        if campo.get_attribute("readonly") is not None:
            break
        page.wait_for_timeout(200)

    if campo.get_attribute("readonly") is None:
        return "Documento não abriu"

    page.wait_for_timeout(800)

    # Menu → Eliminar
    frame.get_by_role("button", name=re.compile("Menu", re.I)).click()
    page.wait_for_timeout(300)
    frame.get_by_role("menuitem", name="Documento de vendas").click()
    page.wait_for_timeout(200)
    frame.get_by_role("menuitem", name="Eliminar").click()
    page.wait_for_timeout(400)

    try:
        frame.locator("div[title='Sim']").wait_for(timeout=5000)
        frame.locator("div[title='Sim']").click()
    except:
        return "Popup Sim não apareceu"

    # Aguardar mensagem final
    mensagem = aguardar_mensagem_nova(frame, mensagem_base, pedido)

    if mensagem:
        garantir_tela_pesquisa(page)
        return mensagem

    garantir_tela_pesquisa(page)
    return "Timeout aguardando resposta do SAP"


# ========== MAIN ==========
def main():
    xlsx = Path(EXCEL_PATH)
    if not xlsx.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {xlsx}")

    df = pd.read_excel(xlsx, dtype=str).fillna("")

    total = len(df)
    print("\n" + "=" * 60)
    print(f"Arquivo carregado : {xlsx.name}")
    print(f"Total de pedidos  : {total}")
    print("=" * 60)

    if total == 0:
        return

    if input("Deseja continuar com essa quantidade? (S/N): ").strip().upper() != "S":
        return

    for col in ["Status", "Mensagem", "Data/Hora"]:
        if col not in df.columns:
            df[col] = ""

    env, home_url, app_url, profile_dir = pick_env()

    sap_user = sap_pass = ""
    if env == "QAS":
        sap_user = input("User SAP: ").strip()
        sap_pass = input("Senha SAP: ").strip()
    else:
        print("PRD — SSO ativo.")

    out = xlsx.with_name(xlsx.stem + "_resultado.xlsx")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            channel="chrome",
            slow_mo=200,
            args=["--start-maximized"]
        )

        page = context.new_page()
        page.goto(home_url)

        if env == "QAS":
            login_qas_if_needed(page, sap_user, sap_pass)

        wait_shell(page)
        page.goto(app_url, timeout=120000)

        for idx, row in df.iterrows():
            pedido = row.get("Pedido", "").strip()
            garantir_tela_pesquisa(page)

            try:
                resultado = processar_exclusao_pedido(page, pedido)
                msg_low = resultado.lower()

                if (
                    "foi eliminado" in msg_low
                    or "não existe no banco de dados" in msg_low
                    or "foi arquivado" in msg_low
                ):
                    df.at[idx, "Status"] = "OK"
                else:
                    df.at[idx, "Status"] = "ERRO"

                df.at[idx, "Mensagem"] = resultado
                df.at[idx, "Data/Hora"] = now_str()

            except Exception as e:
                df.at[idx, "Status"] = "ERRO"
                df.at[idx, "Mensagem"] = str(e)
                df.at[idx, "Data/Hora"] = now_str()

            df.to_excel(out, index=False)

        context.close()

    print("Processo finalizado. Resultado:", out)


if __name__ == "__main__":
    main()