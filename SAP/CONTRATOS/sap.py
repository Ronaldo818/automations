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

            if possui_valor(dados["Centro"]):
                self.escrever(ID_CENTRO, limpar_texto(dados["Centro"]))
            else:
                self.escrever(ID_CENTRO, "")

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
            # 1. Preenchemos os campos básicos da tabela SEM enviar o ENTER
            self.escrever(ID_MATERIAL(linha), inteiro(item["Material"]))
            self.escrever(ID_QUANTIDADE(linha), decimal(item["Qntde Prev"], 0))
            self.escrever(ID_PRECO(linha), decimal(item["valor"]))
            
            # 2. Tratamos o KNTTP e executamos a navegação de tela
            if possui_valor(item["Classificação Contabil"]):
                # Se tem KNTTP, preenche na célula e usa o ENTER para validar a linha
                self.escrever(ID_KNTTP(linha), limpar_texto(item["Classificação Contabil"]))
                self.enter()
                
                # Limpa avisos amarelos (ex: "Preço efetivo é X" ou "Data de remessa")
                if self.status():
                    self.enter()
                    if self.status(): # Duplo check de segurança
                        self.enter()
            else:
                # Se NÃO tem KNTTP, usamos o duplo clique na célula como nosso "avançar"
                self.duplo_clique(ID_KNTTP(linha))
                
                # Limpa avisos amarelos que possam surgir após o duplo clique
                if self.status():
                    self.enter()
                    if self.status():
                        self.enter()

            # 3. Neste ponto, o robô já está na tela de detalhes (onde fica o imposto)
            self.preencher_imposto(item["Cód. Imposto"])
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
            texto = self.ler(BARRA_STATUS).strip()
            
            # Se houver texto e ele não for idêntico ao último capturado, nós o registramos
            if texto and (not self.historico_mensagens or self.historico_mensagens[-1] != texto):
                self.historico_mensagens.append(texto)
                
            return texto

    # =====================================================
    # CONTRATO
    # =====================================================
    def numero_contrato(self):
        return extrair_numero_contrato(self.status())