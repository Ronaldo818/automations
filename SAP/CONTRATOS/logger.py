"""
=========================================================
LOGGER DA AUTOMAÇÃO
=========================================================
"""

import os
import getpass
from datetime import datetime
from pathlib import Path

import pandas as pd

from config import (
    ARQUIVO_LOG_EXCEL,
    ARQUIVO_LOG_TXT,
    SALVAR_LOG_EXCEL,
    SALVAR_LOG_TXT,
    SAP_TRANSACAO
)


class Logger:
    def __init__(self):
        self.logs = []
        self.inicio = datetime.now()
        self.usuario = getpass.getuser()
        self.maquina = os.environ.get('COMPUTERNAME', 'Desconhecida')

    # ==================================================
    def adicionar(
            self,
            dados_linha,
            contrato_sap,
            status,
            detalhes,
            historico_mensagens
    ):
        """Registra a linha de execução e todas as colunas de input."""
        # 1. Cria o registro base com os dados de auditoria e resposta do SAP
        log_entry = {
            "Data/Hora Execução": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Usuário OS": self.usuario,
            "Máquina": self.maquina,
            "Transação": SAP_TRANSACAO,
            "Contrato Gerado SAP": contrato_sap,
            "Status Execução": status,
            "Mensagem Erro/Falha": detalhes,
            "Histórico de Avisos SAP": " | ".join(historico_mensagens)
        }
        
        # 2. Despeja TODAS as colunas do Excel original dinamicamente neste registro
        for coluna, valor in dados_linha.items():
            log_entry[f"Input_{coluna}"] = valor

        # 3. Salva na memória do robô
        self.logs.append(log_entry)

    # ==================================================
    def _garantir_pasta(self, caminho_arquivo):
        """Garante que a pasta de destino do log exista."""
        pasta = Path(caminho_arquivo).parent
        pasta.mkdir(parents=True, exist_ok=True)

    # ==================================================
    def salvar_excel(self):
        """Salva o histórico em um arquivo Excel."""
        if not SALVAR_LOG_EXCEL or not self.logs:
            return

        self._garantir_pasta(ARQUIVO_LOG_EXCEL)
        df = pd.DataFrame(self.logs)
        caminho = Path(ARQUIVO_LOG_EXCEL)
        
        try:
            df.to_excel(caminho, index=False)
        except PermissionError:
            # Salva-vidas caso o arquivo esteja aberto pelo usuário
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_alternativo = f"{caminho.stem}_{timestamp}{caminho.suffix}"
            df.to_excel(caminho.with_name(nome_alternativo), index=False)

    # ==================================================
    def salvar_txt(self):
        """Salva um resumo estruturado em formato TXT."""
        if not SALVAR_LOG_TXT or not self.logs:
            return

        self._garantir_pasta(ARQUIVO_LOG_TXT)
        caminho = Path(ARQUIVO_LOG_TXT)

        # Cabeçalho do arquivo de texto
        linhas_txt = [
            "========== RELATÓRIO DE AUDITORIA ==========\n\n",
            f"Transação : {SAP_TRANSACAO}\n",
            f"Operador  : {self.usuario} (Máquina: {self.maquina})\n",
            f"Início    : {self.inicio.strftime('%d/%m/%Y %H:%M:%S')}\n",
            f"Fim       : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n",
            "============================================\n\n"
        ]
        
        # Insere as informações de cada item executado
        for log in self.logs:
            id_contrato = log.get("Input_ID_CONTRATO", "-")
            material = log.get("Input_Material", "-")
            
            linhas_txt.append(
                f"[{log['Status Execução']}] "
                f"Contrato: {id_contrato} | "
                f"Material: {material} | "
                f"SAP: {log['Contrato Gerado SAP']} | "
                f"Detalhes: {log['Mensagem Erro/Falha']} | "
                f"Avisos SAP: {log['Histórico de Avisos SAP']}\n"
            )

        try:
            with open(caminho, "w", encoding="utf-8") as arquivo:
                arquivo.writelines(linhas_txt)
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_alternativo = f"{caminho.stem}_{timestamp}{caminho.suffix}"
            with open(caminho.with_name(nome_alternativo), "w", encoding="utf-8") as arquivo:
                arquivo.writelines(linhas_txt)

    # ==================================================
    def salvar(self):
        """Dispara a gravação final nos arquivos."""
        self.salvar_excel()
        self.salvar_txt()