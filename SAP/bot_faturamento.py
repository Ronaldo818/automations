import customtkinter as ctk
import threading
import time
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ============================================================
# CONFIGURAÇÕES E SELETORES SAP FIORI
# ============================================================
URLQAS = "https://s4qas.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=310&sap-language=PT#ZSD_COCKPIT_FAT_AGROSYS-display"
URLPRD = "https://s4prd.sap.avivar.com.br/sap/bc/ui2/flp?sap-client=300&sap-language=PT#ZSD_COCKPIT_FAT_AGROSYS-display"

# Seletores
SEL_INPUT_OC = "input[id*='FilterField::Carga-inner-inner']"
SEL_INPUT_DATA = "input[id*='FilterField::DataRemessa-inner-inner']"
SEL_BTN_INICIAR_BUSCA = "bdi:has-text('Iniciar')"
SEL_CHECKBOX_TODOS = "div[id*='LineItem-innerTable-sa-CbBg']"
SEL_BTN_FATURAR = "bdi:has-text('Faturar')"

# Tabela e Status
SEL_LINHAS_TABELA = "tbody tr.sapMListTblRow"
SEL_CELULA_STATUS_FAT = "td[data-sap-ui-column*='StatusFaturamento-innerColumn'] .sapMObjStatusText"
SEL_COLUNA_DOCNUM = "th[id*='DocNum-innerColumn']"
SEL_MENU_CRESCENTE = "text='Crescente'" # Texto do menu do SAP UI5 em português

# Impressão
SEL_BTN_IMPRIMIR_NFE = "bdi:has-text('Imprimir NF-e')"
SEL_INPUT_IMPRESSORA = "input[id*='printer-inner-inner']"
SEL_BTN_CONFIRMAR_IMPRESSAO = "bdi[id*='Imprimir::Action::Ok-BDI-content']"
SEL_TOAST = ".sapMMessageToast"


class FaturamentoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Robô de Faturamento - Avivar")
        self.geometry("500x550")
        ctk.set_appearance_mode("dark")
        
        # --- UI Elements ---
        self.lbl_title = ctk.CTkLabel(self, text="Faturamento de Ordem de Carga", font=("Arial", 18, "bold"))
        self.lbl_title.pack(pady=(20, 10))

        self.env_var = ctk.StringVar(value="PRD")
        self.radio_qas = ctk.CTkRadioButton(self, text="QAS (Homologação)", variable=self.env_var, value="QAS")
        self.radio_qas.pack(pady=5)
        self.radio_prd = ctk.CTkRadioButton(self, text="PRD (Produção)", variable=self.env_var, value="PRD")
        self.radio_prd.pack(pady=5)

        self.inp_oc = ctk.CTkEntry(self, placeholder_text="Ordem de Carga (ex: 41230/1)", width=300)
        self.inp_oc.pack(pady=10)

        self.inp_data = ctk.CTkEntry(self, placeholder_text="Data Remessa (Vazio = Hoje)", width=300)
        self.inp_data.pack(pady=10)

        self.btn_iniciar = ctk.CTkButton(self, text="Iniciar Automação", command=self.iniciar_thread, width=300, fg_color="green", hover_color="darkgreen")
        self.btn_iniciar.pack(pady=20)

        self.txt_log = ctk.CTkTextbox(self, width=450, height=200, state="disabled")
        self.txt_log.pack(pady=10)
        
    def log(self, mensagem):
        """Atualiza a caixa de texto na interface"""
        agora = datetime.now().strftime("%H:%M:%S")
        msg_formatada = f"[{agora}] {mensagem}\n"
        
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg_formatada)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")
        self.update()
        
    def executar_com_retentativas(self, acao_func, nome_acao, max_tentativas=3):
        """Executa uma função e tenta novamente em caso de falha."""
        for tentativa in range(1, max_tentativas + 1):
            try:
                return acao_func() # Tenta rodar o código
            except Exception as e:
                self.log(f"Aviso: Falha na etapa '{nome_acao}' (Tentativa {tentativa}/{max_tentativas}). Erro: {str(e)}")
                if tentativa == max_tentativas:
                    self.log(f"ERRO FATAL: Limite de retentativas atingido para '{nome_acao}'.")
                    raise Exception(f"Falha definitiva em {nome_acao}: {str(e)}")
                time.sleep(3)

    def iniciar_thread(self):
        """Dispara a automação em background para não travar a tela"""
        oc = self.inp_oc.get().strip()
        data_remessa = self.inp_data.get().strip()
        ambiente = self.env_var.get()

        # Validação da OC
        if not (oc.endswith("/1") or oc.endswith("/13")):
            self.log("ERRO: A OC deve terminar com '/1' ou '/13'.")
            return
            
        # Tratamento da Data
        if not data_remessa:
            data_remessa = datetime.now().strftime("%d.%m.%Y")
            
        # Define Impressora baseada na OC
        impressora = "ILOGM003" if oc.endswith("/1") else "ILOGU001"

        self.btn_iniciar.configure(state="disabled", text="Rodando...")
        self.log(f"Iniciando para OC {oc} no ambiente {ambiente}...")
        self.log(f"Impressora mapeada: {impressora} | Data: {data_remessa}")

        # Inicia a thread do Playwright
        threading.Thread(target=self.rodar_automacao, args=(oc, data_remessa, ambiente, impressora), daemon=True).start()

    def wait_busy(self, page):
        """Espera a tela do SAP Fiori terminar de carregar (flores sumirem)"""
        try:
            page.wait_for_function("""() => {
                const els = Array.from(document.querySelectorAll('.sapUiLocalBusyIndicator'));
                return els.every(el => window.getComputedStyle(el).display === 'none' || window.getComputedStyle(el).visibility === 'hidden');
            }""", timeout=60000)
            time.sleep(1) # Pausa extra de segurança
        except:
            pass

    def rolar_tabela_ate_final(self, page):
        """Rola a tabela para forçar o carregamento de todas as remessas"""
        self.log("Rolando tabela para carregar todos os itens...")
        page.locator(SEL_LINHAS_TABELA).first.wait_for(state="visible", timeout=30000)
        
        qtd_anterior = 0
        tentativas_sem_mudar = 0
        
        while tentativas_sem_mudar < 3:
            page.keyboard.press("PageDown")
            time.sleep(1)
            qtd_atual = page.locator(SEL_LINHAS_TABELA).count()
            
            if qtd_atual == qtd_anterior:
                tentativas_sem_mudar += 1
            else:
                tentativas_sem_mudar = 0
                qtd_anterior = qtd_atual
                
        self.log(f"Tabela carregada. Total de linhas: {qtd_atual}")
        return qtd_atual

    def rodar_automacao(self, oc, data, ambiente, impressora):
        url = URLQAS if ambiente == "QAS" else URLPRD
        
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, args=["--start-maximized"])
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
            page = context.new_page()

            try:
                # 1. Acessa o SAP
                self.log("Abrindo navegador... Faça login se necessário.")
                page.goto(url, timeout=90000)
                
                # Aguarda campo OC estar visível (Indica que logou)
                page.locator(SEL_INPUT_OC).wait_for(state="visible", timeout=120000)
                self.log("SAP Carregado. Preenchendo dados...")

                # 2. Pesquisa de Remessas com Retentativas
                def buscar_remessas():
                    page.fill(SEL_INPUT_OC, oc)
                    page.fill(SEL_INPUT_DATA, data)
                    page.keyboard.press("Enter")
                    page.click(SEL_BTN_INICIAR_BUSCA)
                    self.wait_busy(page)
                    total = self.rolar_tabela_ate_final(page)
                    if total == 0:
                        raise ValueError("Tabela retornou vazia.")
                    return total

                try:
                    total_linhas = self.executar_com_retentativas(buscar_remessas, "Consulta da OC", max_tentativas=3)
                except Exception as e:
                    self.log("OC não encontrada ou nenhuma remessa localizada. Processo encerrado.")
                    return # Interrompe a execução aqui conforme regra de negócio

                # 3. Faturamento com Retentativas
                def acao_faturar():
                    page.click(SEL_CHECKBOX_TODOS)
                    time.sleep(1)
                    page.click(SEL_BTN_FATURAR)
                    self.wait_busy(page)
                
                self.log("Iniciando Faturamento...")
                self.executar_com_retentativas(acao_faturar, "Faturamento", max_tentativas=3)

                # 4. Loop de Polling (Aguardando SEFAZ - Status 3) com Timeout
                self.log("Aguardando autorização da SEFAZ (Status 3)...")
                autorizados = 0
                tentativas_sefaz = 0
                max_tentativas_sefaz = 30 # Limite configurável (aprox. 3 a 4 min)
                
                while autorizados < total_linhas:
                    if tentativas_sefaz >= max_tentativas_sefaz:
                        raise Exception("Timeout no retorno da SEFAZ. Processo interrompido.")
                        
                    page.click(SEL_BTN_INICIAR_BUSCA)
                    self.wait_busy(page)
                    self.rolar_tabela_ate_final(page)
                    
                    status_elements = page.locator(SEL_CELULA_STATUS_FAT).all_inner_texts()
                    autorizados = status_elements.count("3")
                    self.log(f"SEFAZ: {autorizados}/{total_linhas} autorizados...")
                    
                    if autorizados < total_linhas:
                        time.sleep(5) # Espera 5s antes de checar de novo
                        tentativas_sefaz += 1
                
                self.log("Todos os documentos faturados! Preparando impressão...")

                # 5. Ordenação Crescente
                page.click(SEL_COLUNA_DOCNUM)
                time.sleep(1)
                self.executar_com_retentativas(
                    lambda: page.locator(SEL_MENU_CRESCENTE).click(),
                    "Clique no menu de ordenação"
                )
                self.wait_busy(page)

                # 6. Desmarca Seleção Geral (Checkbox Mestre)
                page.click(SEL_CHECKBOX_TODOS)
                time.sleep(1)

                # 7. Loop de Impressão (30 em 30) com Retentativas
                checkboxes = page.locator("div[id*='-selectMulti-CbBg']").all()
                lotes = [checkboxes[i:i + 30] for i in range(0, len(checkboxes), 30)]
                
                for idx, lote in enumerate(lotes):
                    self.log(f"Imprimindo lote {idx+1}/{len(lotes)} ({len(lote)} notas)...")
                    
                    def acao_imprimir_lote():
                        # Marca o lote
                        for cb in lote:
                            cb.scroll_into_view_if_needed()
                            cb.click()
                        
                        # Abre Modal de Impressão
                        page.click(SEL_BTN_IMPRIMIR_NFE)
                        page.locator(SEL_INPUT_IMPRESSORA).wait_for(state="visible")
                        
                        # Preenche Impressora e Confirma
                        page.fill(SEL_INPUT_IMPRESSORA, impressora)
                        time.sleep(1)
                        page.click(SEL_BTN_CONFIRMAR_IMPRESSAO)
                        self.wait_busy(page)
                        
                        # Aguarda Toast sumir para evitar atropelo
                        try:
                            page.locator(SEL_TOAST).first.wait_for(state="visible", timeout=10000)
                            page.locator(SEL_TOAST).first.wait_for(state="hidden", timeout=15000)
                        except:
                            pass # Se não capturou o toast, apenas segue
                        
                        # Desmarca o lote
                        for cb in lote:
                            cb.scroll_into_view_if_needed()
                            cb.click()

                    # Executa a impressão do lote encapsulada nas retentativas
                    self.executar_com_retentativas(acao_imprimir_lote, f"Impressão do Lote {idx+1}", max_tentativas=3)

                self.log(f"PROCESSO CONCLUÍDO COM SUCESSO! OC: {oc}")
                
            except Exception as e:
                self.log(f"ERRO: {str(e)}")
            finally:
                self.log("Fechando navegador...")
                browser.close()
                self.btn_iniciar.configure(state="normal", text="Iniciar Automação")


if __name__ == "__main__":
    app = FaturamentoApp()
    app.mainloop()