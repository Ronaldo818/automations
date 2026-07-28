"""
=========================================================
LOGGER DA AUTOMAÇÃO
=========================================================
"""

from datetime import datetime
from pathlib import Path

import pandas as pd

from config import (
    ARQUIVO_LOG_EXCEL,
    ARQUIVO_LOG_TXT,
    SALVAR_LOG_EXCEL,
    SALVAR_LOG_TXT
)


class Logger:
    def __init__(self):
        self.logs = []
        self.inicio = datetime.now()

    # ==================================================
    def adicionar(
            self,
            id_contrato,
            linha_excel,
            fornecedor,
            material,
            contrato_sap,
            status,
            mensagem
    ):
        self.logs.append({
            "Data/Hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "ID_CONTRATO": id_contrato,
            "Linha Excel": linha_excel,
            "Fornecedor": fornecedor,
            "Material": material,
            "Contrato SAP": contrato_sap,
            "Status": status,
            "Mensagem": mensagem
        })

    # ==================================================
    def _garantir_pasta(self, caminho_arquivo):
        """Garante que a pasta de destino do log exista."""
        pasta = Path(caminho_arquivo).parent
        pasta.mkdir(parents=True, exist_ok=True)

    # ==================================================
    def salvar_excel(self):
        if not SALVAR_LOG_EXCEL:
            return
        if not self.logs:
            return

        self._garantir_pasta(ARQUIVO_LOG_EXCEL)
        df = pd.DataFrame(self.logs)

        caminho = Path(ARQUIVO_LOG_EXCEL)
        
        try:
            df.to_excel(caminho, index=False)
        except PermissionError:
            # SALVA-VIDAS: Se o usuário estiver com o log aberto, não perde a execução!
            # Salva um novo arquivo com timestamp.
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_alternativo = f"{caminho.stem}_{timestamp}{caminho.suffix}"
            caminho_alternativo = caminho.with_name(nome_alternativo)
            
            df.to_excel(caminho_alternativo, index=False)

    # ==================================================
    def salvar_txt(self):
        if not SALVAR_LOG_TXT:
            return
        if not self.logs:
            return

        self._garantir_pasta(ARQUIVO_LOG_TXT)
        caminho = Path(ARQUIVO_LOG_TXT)

        # Prepara o conteúdo do TXT
        linhas_txt = [
            "========== EXECUÇÃO ==========\n\n",
            f"Início : {self.inicio.strftime('%d/%m/%Y %H:%M:%S')}\n",
            f"Fim    : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
        ]
        
        for log in self.logs:
            linhas_txt.append(
                f"[{log['Status']}] "
                f"Contrato: {log['ID_CONTRATO']} | "
                f"Material: {log['Material']} | "
                f"SAP: {log['Contrato SAP']} | "
                f"{log['Mensagem']}\n"
            )

        try:
            with open(caminho, "w", encoding="utf-8") as arquivo:
                arquivo.writelines(linhas_txt)
        except PermissionError:
            # Mesmo fallback de segurança para o TXT
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_alternativo = f"{caminho.stem}_{timestamp}{caminho.suffix}"
            caminho_alternativo = caminho.with_name(nome_alternativo)
            
            with open(caminho_alternativo, "w", encoding="utf-8") as arquivo:
                arquivo.writelines(linhas_txt)

    # ==================================================
    def salvar(self):
        self.salvar_excel()
        self.salvar_txt()