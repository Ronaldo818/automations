from __future__ import annotations

import os
import sys
import threading
import win32event
import win32api
import winerror
import sys
from datetime import datetime
from playwright.sync_api import sync_playwright
import tkinter as tk
from tkinter import scrolledtext

mutex = win32event.CreateMutex(None, False, "EnvioNotasAgrosysMutex")

if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    print("Já existe uma instância em execução.")
    sys.exit(0)

EXECUTANDO = True

timestamp = datetime.now().strftime("%Y%m%d_%H%M")
LOG_PATH = f"C:\\python_scripts\\logs\\envio_notas_{timestamp}.log"
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

def salvar_log_arquivo(texto):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(texto + "\n")

class App:

    def __init__(self, root):
        self.root = root
        self.root.title("Envio de Notas")
        self.root.geometry("700x500")

        self.status = tk.Label(root, text="Status: Aguardando", font=("Arial", 12))
        self.status.pack()

        self.nota = tk.Label(root, text="Nota atual: -", font=("Arial", 12))
        self.nota.pack()

        self.tabela = tk.Label(root, text="Sucesso: 0 | Erros: 0", font=("Arial", 12, "bold"))
        self.tabela.pack()

        self.btn_parar = tk.Button(
            root,
            text="ENCERRAR",
            bg="red",
            fg="white",
            font=("Arial", 12, "bold"),
            command=self.parar_execucao
        )
        self.btn_parar.pack(pady=5)

        self.log = scrolledtext.ScrolledText(root, height=20)
        self.log.pack(fill="both", expand=True)

    def atualizar_status(self, texto):
        self.status.config(text=f"Status: {texto}")
        self.root.update()

    def atualizar_nota(self, nota):
        self.nota.config(text=f"Nota atual: {nota}")
        self.root.update()

    def atualizar_contador(self, enviados, erros):
        self.tabela.config(text=f"Sucesso: {enviados} | Erros: {erros}")
        self.root.update()

    def escrever_log(self, texto):
        self.log.insert(tk.END, texto + "\n")
        self.log.see(tk.END)
        self.root.update()

    def parar_execucao(self):
        global EXECUTANDO
        EXECUTANDO = False
        self.atualizar_status("Encerrando...")
        self.escrever_log("Encerramento solicitado pelo usuário")
        try:
            self.root.quit()
            self.root.destroy()
        except Exception:
            pass



def finalizar(browser=None, page=None, root=None):
    global EXECUTANDO
    EXECUTANDO = False

    try:
        if page:
            page.close()
    except Exception:
        pass

    try:
        if browser:
            browser.close()
    except Exception:
        pass

    try:
        if root:
            root.quit()
            root.destroy()
    except Exception:
        pass


def esperar_mudar_linha(frame, page, linha_antiga):
    for _ in range(10):
        if not EXECUTANDO:
            return
        try:
            nova = frame.locator("input.radio_sel").first.locator("xpath=ancestor::tr").inner_text()
            if nova != linha_antiga:
                return
        except:
            return
        page.wait_for_timeout(300)

def esperar_tabela_atualizar(frame, texto_antes):
    for _ in range(20):
        try:
            texto_depois = frame.locator("#vtabela").inner_text()
            if texto_depois != texto_antes:
                return
        except:
            return
        frame.page.wait_for_timeout(300)

def executar_envio(usuario, senha, data_ini, data_fim, app=None):

    global EXECUTANDO

    inicio_execucao = datetime.now()

    enviados = 0
    erros = 0

    notas_processadas_execucao = set()

    def log(msg, tipo="INFO"):
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        linha = f"[{data_hora}] [{tipo}] {msg}"

        print(linha)
        salvar_log_arquivo(linha)

        if app:
            app.escrever_log(linha)

    if app:
        app.atualizar_status("Rodando")

    log("Iniciando execução")

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        page.goto("https://sistema.avivar.com.br/webpro/webpad/acesso")

        page.fill("input[name='vusuario']", usuario)
        page.fill("input[name='vsenha']", senha)
        page.click("button.BtLogin")

        page.wait_for_load_state("networkidle")

        with page.expect_navigation():
            page.select_option("#vobj-modulo", "22776")

        with page.expect_navigation():
            page.select_option("#vobj-unidade", "2")

        page.click("#ui-id-1")
        page.click("#ui-id-28")
        page.click("#ui-id-29")

        frame = page.frame(name="frameprog")

        frame.wait_for_selector("#vpar-dt-ini")

        frame.fill("#vpar-dt-ini", data_ini)
        frame.fill("#vpar-dt-fim", data_fim)

        with frame.expect_navigation():
            frame.select_option("#wnfs-status", "1")

        with frame.expect_navigation():
            frame.click("#vpad-btpesq\\.x")

        frame.wait_for_selector("input.radio_sel")

        primeira_linha = frame.locator("input.radio_sel").first.locator("xpath=ancestor::tr").inner_text()
        total_primeira = frame.locator("input.radio_sel").count()

        modo_primeira = True

        if total_primeira >= 20:

            btn_ult = frame.locator("input[name='vpad-btult']")

            if btn_ult.count() > 0:

                btn_ult.click()

                esperar_mudar_linha(frame, page, primeira_linha)

                nova = frame.locator("input.radio_sel").first.locator("xpath=ancestor::tr").inner_text()

                if nova != primeira_linha:
                    modo_primeira = False
                    log("Trabalhando da última aba")

        while True:

            if not EXECUTANDO:
                finalizar(browser=browser, page=page, root=app.root if app else None)
                return

            frame.wait_for_selector("input.radio_sel")

            radios = frame.locator("input.radio_sel")
            total = radios.count()

            encontrou = False

            for i in reversed(range(total)):

                if not EXECUTANDO:
                    finalizar(browser=browser, page=page, root=app.root if app else None)
                    return

                linha = radios.nth(i).locator("xpath=ancestor::tr")

                numero = linha.locator("td").nth(3).inner_text().strip()
                cliente = linha.locator("td").nth(15).inner_text().strip()

                img = linha.locator("img[alt]")

                if img.count() == 0:
                    continue

                status = img.first.get_attribute("alt")

                if status == "Gerada":
                    continue

                if status != "Autorizada":
                    log(f"Nota: {numero} | Status: {status}", "IGNORADO")
                    continue

                chave = numero

                if chave in notas_processadas_execucao:
                    continue

                notas_processadas_execucao.add(chave)

                if app:
                    app.atualizar_nota(numero)

                encontrou = True

                texto_antes = frame.locator("#vtabela").inner_text()

                radios.nth(i).click()

                log(f"Nota: {numero} | Cliente: {cliente}", "ENVIO")

                frame.click("input[name='btenvia.x']")

                frame.wait_for_selector("div.box_mensagens", timeout=10000)

                mensagem = frame.locator("div.box_mensagens").first.inner_text()

                if "Processo de integração foi iniciado" in mensagem:
                    enviados += 1
                    resultado = "SUCESSO"
                else:
                    erros += 1
                    resultado = "ERRO"

                mensagem_limpa = mensagem.replace("\n", " ").strip()

                if app:
                    app.atualizar_contador(enviados, erros)

                log(
                    f"Nota: {numero} | Cliente: {cliente} | Msg: {mensagem_limpa}",
                    resultado
                )

                esperar_tabela_atualizar(frame, texto_antes)

                if not modo_primeira:

                    btn_ult = frame.locator("input[name='vpad-btult']")

                    if btn_ult.count() > 0:

                        atual = frame.locator("input.radio_sel").first.locator("xpath=ancestor::tr").inner_text()

                        btn_ult.click()

                        esperar_mudar_linha(frame, page, atual)

                break

            if not encontrou:

                if modo_primeira:
                    break

                btn_ant = frame.locator("input[name='vpad-btant']")

                if btn_ant.count() == 0:
                    break

                atual = frame.locator("input.radio_sel").first.locator("xpath=ancestor::tr").inner_text()

                btn_ant.click()

                esperar_mudar_linha(frame, page, atual)

        tempo_total = str(datetime.now() - inicio_execucao).split(".")[0]

        log(
            f"Sucesso: {enviados} | Erros: {erros} | Tempo: {tempo_total}",
            "RESUMO"
        )

        finalizar(browser=browser, page=page)

    if app:
        app.atualizar_status("Finalizado")
        app.escrever_log("Processo concluído")
        app.root.after(1500, lambda: finalizar(root=app.root))

def iniciar(usuario, senha):

    hoje = datetime.now()

    data_ini = hoje.replace(day=1).strftime("%d/%m/%Y")
    data_fim = hoje.strftime("%d/%m/%Y")

    root = tk.Tk()
    app = App(root)

    def rodar():
        executar_envio(usuario, senha, data_ini, data_fim, app)

    threading.Thread(target=rodar, daemon=True).start()

    root.mainloop()

if __name__ == "__main__":

    if len(sys.argv) >= 3:
        usuario = sys.argv[1]
        senha = sys.argv[2]
    else:
        usuario = "SEU_USUARIO"
        senha = "SUA_SENHA"

    iniciar(usuario, senha)