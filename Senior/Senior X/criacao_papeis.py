import requests
import pandas as pd
from datetime import datetime

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\Users\junio\OneDrive\Área de Trabalho\Documentos\Scripts Github\automations\Planilhas\Roles_input.xlsx"
ARQUIVO_LOG = r"C:\Users\junio\OneDrive\Área de Trabalho\Documentos\Scripts Github\automations\Planilhas\Roles_input_logs.xlsx"
TOKEN_SENIOR = "V2pXGPEaHvONvpMBkj0bwYpvotxm89Mv"

# =========================
# API & HEADERS
# =========================
URL = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/platform/authorization/actions/createRole"
HEADERS = {
    "Authorization": f"Bearer {TOKEN_SENIOR}",
    "Content-Type": "application/json"
}

def main():
    print("Iniciando leitura da planilha de entrada...")
    
    # 1. Carrega os dados de entrada
    try:
        planilha = pd.read_excel(ARQUIVO_ENTRADA).fillna("")
    except Exception as e:
        print(f"Erro ao ler o arquivo de entrada: {e}")
        return

    # Lista vazia que vai armazenar os resultados para o Log
    dados_log = []

    # 2. Loop pelas linhas
    for index, linha in planilha.iterrows():
        # ATENÇÃO: Certifique-se de que os nomes das colunas na sua planilha são estes mesmos
        nome_role = linha['Nome']
        descricao_role = linha['Descricao']
        
        payload = {
            "name": nome_role,
            "description": descricao_role
        }
        
        print(f"[{index + 1}/{len(planilha)}] Criando Role: {nome_role}...")
        
        # 3. Disparo para a API
        try:
            resposta = requests.post(URL, json=payload, headers=HEADERS)
            status_code = resposta.status_code
            
            if status_code == 200:
                mensagem_retorno = "Sucesso"
            else:
                # Se der erro (ex: role já existe), pega a mensagem da Senior
                mensagem_retorno = resposta.text 
                
        except Exception as e:
            # Caso caia a internet ou a API não responda
            status_code = "Falha de Conexão"
            mensagem_retorno = str(e)
            
        # 4. Registra o resultado na lista de log
        dados_log.append({
            "Nome_Role": nome_role,
            "Status_HTTP": status_code,
            "Mensagem": mensagem_retorno,
            "Data_Hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        })

    # 5. Salva o Log no novo arquivo Excel
    print("\nProcessamento concluído! Gerando arquivo de log...")
    df_log = pd.DataFrame(dados_log)
    
    try:
        df_log.to_excel(ARQUIVO_LOG, index=False)
        print(f"Arquivo de log salvo com sucesso em: {ARQUIVO_LOG}")
    except Exception as e:
        print(f"Erro ao salvar arquivo de log (feche a planilha se estiver aberta): {e}")

# Executa o script
if __name__ == "__main__":
    main()