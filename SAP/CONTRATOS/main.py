"""
=========================================================
AUTOMAÇÃO SAP ME31K
CRIAÇÃO DE CONTRATOS
=========================================================
"""

import threading
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


def executar_processo(tela, logger, sap):
    """
    Função que roda em background (Thread) para não congelar a interface.
    """
    try:
        tela.atualizar_status("Lendo planilha...")
        df = carregar_excel()
        validar_planilha(df)
        info = resumo(df)

        tela.escrever(f"Contratos encontrados: {info['contratos']}")
        tela.escrever(f"Itens encontrados: {info['itens']}\n")
        
        tela.atualizar_status("Conectando ao SAP...")
        sap.conectar()
        
        contratos = agrupar_contratos(df)
        total = info["contratos"]

        for indice, (id_contrato, grupo) in enumerate(contratos):
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

            # =========================================================
            # NOVO: Try/Except isolado POR CONTRATO
            # =========================================================
            try:
                sap.abrir_me31k()
                sap.preencher_cabecalho(cabecalho)

                for linha, item in itens.iterrows():
                    tela.atualizar_item(linha + 1)
                    tela.escrever(f"Inserindo material {item['Material']}")
                    
                    sap.preencher_item(linha, item)
                    
                    if linha < len(itens) - 1:
                        sap.novo_item()

                sap.salvar()
                numero_sap = sap.numero_contrato()
                tela.atualizar_sap(numero_sap)
                tela.escrever(f"Sucesso! Contrato SAP gerado: {numero_sap}")

                # Log de Sucesso para cada item deste contrato
                for linha, item in itens.iterrows():
                    logger.adicionar(
                        id_contrato=id_contrato,
                        linha_excel=linha + 2,
                        fornecedor=item["Fornecedor"],
                        material=item["Material"],
                        contrato_sap=numero_sap,
                        status="SUCESSO",
                        mensagem="Contrato criado"
                    )

            except Exception as erro_contrato:
                # Se este contrato falhar, loga o erro, avisa na tela e CONTINUA o loop
                erro_msg = str(erro_contrato)
                tela.escrever(f"FALHA no contrato {id_contrato}: {erro_msg}")
                
                for linha, item in itens.iterrows():
                    logger.adicionar(
                        id_contrato=id_contrato,
                        linha_excel=linha + 2,
                        fornecedor=item["Fornecedor"],
                        material=item["Material"],
                        contrato_sap="",
                        status="ERRO",
                        mensagem=erro_msg
                    )

        # Após terminar o loop de contratos, salva o log final
        logger.salvar()
        tela.atualizar_status("Processo finalizado.")
        tela.escrever("\n" + "=" * 60 + "\nProcesso concluído.")

    except Exception as erro_geral:
        # Pega erros críticos antes do loop (ex: falha ao ler Excel, falha ao conectar no SAP)
        tela.atualizar_status("Erro Crítico")
        tela.escrever(f"\nERRO FATAL: {str(erro_geral)}")


def main():
    tela = Interface()
    logger = Logger()
    sap = SAPContrato()

    # Cria e inicia a Thread de execução do processo
    thread_automacao = threading.Thread(target=executar_processo, args=(tela, logger, sap))
    thread_automacao.daemon = True # Garante que a thread morra se o programa for fechado
    thread_automacao.start()

    # Inicia o mainloop do Tkinter (Esta linha trava a execução principal para manter a UI viva)
    tela.iniciar()

if __name__ == "__main__":
    main()