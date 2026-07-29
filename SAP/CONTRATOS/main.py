"""
=========================================================
AUTOMAÇÃO SAP ME31K
CRIAÇÃO DE CONTRATOS
=========================================================
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from excel import (
    carregar_excel,
    validar_planilha,
    agrupar_contratos,
    obter_cabecalho,
    obter_itens,
    resumo
)
from sap import SAPContrato
from logger import Logger
from interface import Interface


def executar_processo(tela, logger, sap, df, caminho_arquivo):
    """
    Função executada em background (Thread).
    Agora ela já recebe o DataFrame (df) pronto e validado pela tela inicial.
    """
    try:
        info = resumo(df)
        tela.escrever(f"Arquivo carregado: {caminho_arquivo}")
        tela.escrever(f"Contratos encontrados: {info['contratos']}")
        tela.escrever(f"Itens encontrados: {info['itens']}\n")
        
        tela.atualizar_status("Conectando ao SAP...")
        sap.conectar()
        
        contratos = agrupar_contratos(df)
        total = info["contratos"]

        for indice, (id_contrato, grupo) in enumerate(contratos):
            
            # Trava principal: aborta se a janela for fechada
            if getattr(tela, 'parar', False):
                tela.escrever("Processo interrompido pelo usuário.")
                break

            tela.progresso(indice + 1, total)
            tela.atualizar_status("Criando contrato...")
            tela.atualizar_contrato(id_contrato)
            tela.escrever("\n" + "=" * 60)
            tela.escrever(f"Processando Contrato {id_contrato}")

            cabecalho = obter_cabecalho(grupo)
            itens = obter_itens(grupo)
            sap.historico_mensagens = []

            try:
                sap.abrir_me31k()
                sap.preencher_cabecalho(cabecalho)

                # Loop de Itens do contrato
                for linha, item in itens.iterrows():
                    # Trava secundária
                    if getattr(tela, 'parar', False):
                        raise Exception("Execução abortada brutalmente pelo usuário.")
                    
                    tela.atualizar_item(linha + 1)
                    tela.escrever(f"Inserindo material {item['Material']}")
                    sap.preencher_item(linha, item)
                    
                    if linha < len(itens) - 1:
                        sap.novo_item()

                sap.salvar()
                numero_sap = sap.numero_contrato()
                mensagens_sap = sap.historico_mensagens.copy()
                
                tela.atualizar_sap(numero_sap)
                tela.escrever(f"Sucesso! Contrato SAP gerado: {numero_sap}")
                tela.registrar_sucesso() 

                # Log Completo - SUCESSO
                for linha, item in itens.iterrows():
                    logger.adicionar(
                        dados_linha=item.to_dict(),
                        contrato_sap=numero_sap,
                        status="SUCESSO",
                        detalhes="Contrato criado com sucesso",
                        historico_mensagens=mensagens_sap
                    )

            except Exception as erro_contrato:
                mensagens_sap = sap.historico_mensagens.copy()
                erro_msg = str(erro_contrato)
                
                tela.escrever(f"FALHA no contrato {id_contrato}: {erro_msg}")
                tela.registrar_erro()

                # Log Completo - ERRO
                for linha, item in itens.iterrows():
                    logger.adicionar(
                        dados_linha=item.to_dict(),
                        contrato_sap="",
                        status="ERRO",
                        detalhes=erro_msg,
                        historico_mensagens=mensagens_sap
                    )

        logger.salvar()
        tela.atualizar_status("Processo finalizado.")
        tela.escrever("\n" + "=" * 60 + "\nProcesso concluído.")

    except Exception as erro_geral:
        tela.atualizar_status("Erro Crítico")
        tela.escrever(f"\nERRO FATAL: {str(erro_geral)}")


def main():
    root_oculta = tk.Tk()
    root_oculta.withdraw() 
    
    # 1. Abre a caixa de seleção de arquivo
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a Planilha de Contratos",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    
    if not caminho_arquivo:
        return # Encerra se o usuário cancelar
        
    # ========================================================
    # 2. NOVA ETAPA: VALIDAÇÃO E CONFIRMAÇÃO DO ARQUIVO
    # ========================================================
    try:
        # Carrega a planilha imediatamente para verificar erros
        df = carregar_excel(caminho_arquivo)
        validar_planilha(df) 
        info = resumo(df)
        
        # Monta a estrutura de texto para a caixa de confirmação
        contratos_agrupados = agrupar_contratos(df)
        detalhes = []
        for id_contrato, grupo in contratos_agrupados:
            detalhes.append(f"Contrato {id_contrato}: {len(grupo)} item(ns)")
            
        texto_resumo = f"Foram encontrados {info['contratos']} contratos e {info['itens']} itens na planilha.\n\n"
        
        # Junta no máximo 15 contratos para a janela não ficar gigante fora da tela
        texto_resumo += "\n".join(detalhes[:15]) 
        
        if info['contratos'] > 15:
            texto_resumo += f"\n...e mais {info['contratos'] - 15} contratos ocultos."
            
        texto_resumo += "\n\nDeseja iniciar o processamento?"
        
        # Exibe o Pop-Up com Botão Sim/Não
        resposta = messagebox.askyesno("Resumo da Planilha", texto_resumo)
        
        if not resposta:
            return # Encerra se o usuário clicar em "Não"
            
    except Exception as e:
        # Se as validações do excel.py estourarem (ex: coluna faltando ou arquivo aberto)
        # nós usamos o messagebox.showerror para avisar o usuário de forma amigável
        messagebox.showerror("Erro na Planilha", str(e))
        return
    # ========================================================

    # 3. Tudo Certo! Inicia a Interface e o Processamento
    tela = Interface()
    logger = Logger()
    sap = SAPContrato()

    # Passamos o 'df' (já validado) para a Thread para não precisar ler do disco de novo
    thread_automacao = threading.Thread(
        target=executar_processo, 
        args=(tela, logger, sap, df, caminho_arquivo)
    )
    thread_automacao.daemon = True 
    thread_automacao.start()

    tela.iniciar()


if __name__ == "__main__":
    main()