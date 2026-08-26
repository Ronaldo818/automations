"""
=========================================================
ROTINAS SAP GUI - ME31K
=========================================================
"""

import time
import win32com.client
from util import (
    limpar_texto,
    inteiro,
    decimal,
    data_sap,
    possui_valor,
    extrair_numero_contrato
)
import pythoncom

# =============================================================================
# MAPEAMENTO DE IDs (Centralizado)
# =============================================================================

# Comandos
TX_CMD = "wnd[0]/tbar[0]/okcd"

# Cabeçalho
ID_FORNECEDOR = "wnd[0]/usr/ctxtEKKO-LIFNR"
ID_TIPO_CONTRATO = "wnd[0]/usr/ctxtRM06E-EVART"
ID_ORG_COMPRAS = "wnd[0]/usr/ctxtEKKO-EKORG"
ID_GRUPO_COMPRADORES = "wnd[0]/usr/ctxtEKKO-EKGRP"
ID_CENTRO = "wnd[0]/usr/ctxtRM06E-WERKS"
ID_FIM_VALIDADE = "wnd[0]/usr/ctxtEKKO-KDATE"
ID_PAGAMENTO = "wnd[0]/usr/ctxtEKKO-ZTERM"
ID_DIAS_PAGAMENTO = "wnd[0]/usr/ctxtEKKO-ZBD1T"
ID_INCOTERMS = "wnd[0]/usr/ctxtMMPUR_INCOTERMS_CONTRACT-INCO1"
ID_INCOTERMS_LOCAL = "wnd[0]/usr/txtMMPUR_INCOTERMS_CONTRACT-INCO2_L"

# Itens (Usando funções lambda/def para lidar com as linhas dinâmicas da tabela)
TABELA = "wnd[0]/usr/tblSAPMM06ETC_0220"
def ID_MATERIAL(linha): return f"{TABELA}/ctxtEKPO-EMATN[3,{linha}]"
def ID_QUANTIDADE(linha): return f"{TABELA}/txtEKPO-KTMNG[5,{linha}]"
def ID_PRECO(linha): return f"{TABELA}/txtEKPO-NETPR[7,{linha}]"
def ID_KNTTP(linha): return f"{TABELA}/ctxtEKPO-KNTTP[2,{linha}]"
def ID_CENTRO_ITEM(linha): return f"{TABELA}/ctxtEKPO-WERKS[12,{linha}]"

# Imposto
ID_MWSKZ = "wnd[0]/usr/ctxtEKPO-MWSKZ"

# Botões e Status
BTN_SINTESE = "wnd[0]/tbar[1]/btn[5]"
BTN_CLASSIFICACAO = "wnd[0]/tbar[1]/btn[38]"
BTN_SALVAR = "wnd[0]/tbar[0]/btn[11]"
BARRA_STATUS = "wnd[0]/sbar"


class SAPContrato:
    def __init__(self):
        self.session = None
        self.historico_mensagens = []

    # =====================================================
    # CONECTAR
    # =====================================================
    def conectar(self):
            pythoncom.CoInitialize() 
            
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            self.session = connection.Children(0)
            self.session.findById("wnd[0]").maximize()

    # =====================================================
    # MÉTODOS GENÉRICOS DE INTERAÇÃO (Evita quebra por lentidão)
    # =====================================================
    def find(self, id_objeto, timeout=10):
        """Tenta localizar o objeto. Se não achar, tenta novamente até o timeout."""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Único lugar do código inteiro onde usamos findById!
                return self.session.findById(id_objeto)
            except Exception:
                time.sleep(0.5)
        raise Exception(f"Elemento não carregou a tempo no SAP: {id_objeto}")

    def escrever(self, id_objeto, texto):
        elemento = self.find(id_objeto)
        elemento.text = str(texto)

    def pressionar(self, id_objeto):
        elemento = self.find(id_objeto)
        elemento.press()

    def ler(self, id_objeto):
        elemento = self.find(id_objeto)
        return elemento.text

    def enter(self):
        self.session.findById("wnd[0]").sendVKey(0)
    
    def duplo_clique(self, id_objeto):
        """Foca no elemento e envia o comando de duplo clique (F2)"""
        elemento = self.find(id_objeto)
        elemento.setFocus()
        self.session.findById("wnd[0]").sendVKey(2)

    # =====================================================
    # ABRIR ME31K
    # =====================================================
    def abrir_me31k(self):
        self.escrever(TX_CMD, "/nME31K")
        self.enter()

    # =====================================================
    # CABEÇALHO
    # =====================================================
    def preencher_cabecalho(self, dados):
            self.escrever(ID_FORNECEDOR, inteiro(dados["Fornecedor"]))
            self.escrever(ID_TIPO_CONTRATO, limpar_texto(dados["Tipo de contrato"]))
            self.escrever(ID_ORG_COMPRAS, limpar_texto(dados["Organiz.compras"]))
            self.escrever(ID_GRUPO_COMPRADORES, limpar_texto(dados["Grupo de compradores"]))

            #if possui_valor(dados["Centro"]):
            #    self.escrever(ID_CENTRO, limpar_texto(dados["Centro"]))
            #else:
            #    self.escrever(ID_CENTRO, "")

            self.pressionar(BTN_SINTESE)

            if self.status():
                self.enter()

            # Preenche a data para liberar o bloqueio do SAP
            self.escrever(ID_FIM_VALIDADE, data_sap(dados["Fim da validade"]))
            
            try:
                self.session.findById(ID_DIAS_PAGAMENTO).text = ""
            except Exception:
                pass 
            
            self.escrever(ID_PAGAMENTO, limpar_texto(dados["Condições de pagamento"]))
            self.escrever(ID_INCOTERMS, limpar_texto(dados["Incoterms"]))
            self.escrever(ID_INCOTERMS_LOCAL, limpar_texto(dados["Local Incoterms 1"]))

            # Dispara a validação cruzada do SAP (que também já avança de tela)
            self.enter()
            
            mensagem = self.status()
            if "06754" in mensagem or "prestações" in mensagem.lower():
                self.enter()
            elif mensagem:
                self.enter()

    # =====================================================
    # ITEM
    # =====================================================
    def preencher_item(self, linha, item):
            # ---------------------------------------------------------
            # MOTOR DE PAGINAÇÃO (Baseado no script KO02)
            # ---------------------------------------------------------
            tabela = self.session.findById(TABELA)
            linhas_visiveis = tabela.visibleRowCount
            
            pos_scroll = max(linha - linhas_visiveis + 1, 0)
            
            if pos_scroll > 0:
                try:
                    tabela.verticalScrollbar.position = pos_scroll
                    import time
                    time.sleep(0.3)
                except Exception:
                    pass 
                    
            try:
                topo_atual = tabela.verticalScrollbar.position
            except Exception:
                topo_atual = 0 
                
            linha_visivel = linha - topo_atual
            
            if linha_visivel >= linhas_visiveis:
                linha_visivel = linhas_visiveis - 1
                
            # ---------------------------------------------------------
            # PREENCHIMENTO DOS DADOS BÁSICOS
            # ---------------------------------------------------------
            self.escrever(ID_MATERIAL(linha_visivel), inteiro(item["Material"]))
            self.escrever(ID_QUANTIDADE(linha_visivel), decimal(item["Qntde Prev"], 0))
            self.escrever(ID_PRECO(linha_visivel), decimal(item["valor"]))
            
            if possui_valor(item.get("Centro")):
                self.escrever(ID_CENTRO_ITEM(linha_visivel), limpar_texto(item["Centro"]))
                
            # ---------------------------------------------------------
            # A SUA REGRA DE OURO (KNTTP)
            # ---------------------------------------------------------
            if possui_valor(item["Classificação Contabil"]):
                self.escrever(ID_KNTTP(linha_visivel), limpar_texto(item["Classificação Contabil"]))
                self.enter()
                
                if self.status():
                    self.enter()
                    if self.status(): 
                        self.enter()
                        
                try:
                    self.session.findById(ID_KNTTP(linha_visivel)) 
                    self.duplo_clique(ID_KNTTP(linha_visivel))     
                except Exception:
                    pass 

            else:
                self.duplo_clique(ID_KNTTP(linha_visivel))
                
                if self.status():
                    self.enter()
                    if self.status():
                        self.enter()

            # ---------------------------------------------------------
            # FINALIZA O ITEM E TRATA AVISOS (IMPOSTO / PREÇO EFETIVO)
            # ---------------------------------------------------------
            self.preencher_imposto(item["Cód. Imposto"])
            
            # CENÁRIO 2: Captura a mensagem do Preço Efetivo (06207)
            msg_imposto = self.status()
            if "preço efetivo" in msg_imposto.lower():
                self.enter() # Dá um enter extra para limpar o aviso amarelo
                
            self.voltar_sintese()

    # =====================================================
    # IMPOSTO
    # =====================================================
    def preencher_imposto(self, codigo):
        self.escrever(ID_MWSKZ, limpar_texto(codigo))

    # =====================================================
    # SÍNTESE
    # =====================================================
    def voltar_sintese(self):
        self.pressionar(BTN_SINTESE)

    # =====================================================
    # NOVO ITEM
    # =====================================================
    def novo_item(self):
        pass

    # =====================================================
    # SALVAR
    # =====================================================
    def salvar(self):
        self.pressionar(BTN_SALVAR)

    # =====================================================
    # STATUS
    # =====================================================
    def status(self):
            try:
                sbar = self.session.findById("wnd[0]/sbar")
                texto = sbar.text.strip()
                tipo_mensagem = sbar.messageType 
                
                # Grava no histórico para a planilha de auditoria
                if texto and (not self.historico_mensagens or self.historico_mensagens[-1] != texto):
                    self.historico_mensagens.append(texto)
                    
                # ---------------------------------------------------------
                # CENÁRIO 1: Erro de Centro (M3351) - Aborta Imediatamente
                # ---------------------------------------------------------
                if "não está atualizado no centro" in texto.lower():
                    raise Exception("O item precisa ser ajustado no centro de custo.")
                    
                # ---------------------------------------------------------
                # CENÁRIO 2: Tratamento de Erros Vermelhos padrão
                # ---------------------------------------------------------
                if tipo_mensagem in ['E', 'A']:
                    texto_lower = texto.lower()
                    
                    # EXCEÇÃO (Bypass): Erros vermelhos temporários (O robô vai preencher no próximo passo)
                    # Adicionamos "imposto" para que ele não trave no meio do caminho!
                    if "data" in texto_lower or "validade" in texto_lower or "imposto" in texto_lower:
                        return texto 
                    
                    # Qualquer outro erro vermelho que não esteja na exceção, aborta o contrato
                    raise Exception(f"Bloqueio SAP: {texto}")
                    
                return texto
                
            except Exception as e:
                # Repassa a exceção customizada para o main.py capturar
                if "ajustado no centro" in str(e) or "Bloqueio SAP" in str(e):
                    raise e 
                return ""

    # =====================================================
    # CONTRATO
    # =====================================================
    def numero_contrato(self):
        return extrair_numero_contrato(self.status())