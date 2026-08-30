import requests
import pandas as pd
from datetime import datetime
import math # Nova biblioteca nativa para ajudar no cálculo dos lotes

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\Users\junio\OneDrive\Área de Trabalho\Documentos\Scripts Github\automations\Planilhas\Vinculo_recursos - Copia.xlsx"
ARQUIVO_LOG = r"C:\Users\junio\OneDrive\Área de Trabalho\Documentos\Scripts Github\automations\Planilhas\Vinculo_recursos_logs.xlsx"
TOKEN_SENIOR = "V2pXGPEaHvONvpMBkj0bwYpvotxm89Mv"

# Define quantos recursos vão em cada payload (100 é um número muito seguro para a Senior)
TAMANHO_LOTE = 100 

URL = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/platform/authorization/actions/savePermissions"

HEADERS = {
    "Authorization": f"Bearer {TOKEN_SENIOR}",
    "Content-Type": "application/json"
}

def main():
    print("Iniciando leitura da planilha de vínculos...")
    
    try:
        planilha = pd.read_excel(ARQUIVO_ENTRADA, dtype=str).fillna("")
    except Exception as e:
        print(f"Erro ao ler o arquivo de entrada: {e}")
        return

    dados_log = []
    grupos_roles = planilha.groupby("Role")
    total_roles = len(grupos_roles)
    contador_role = 1

    for nome_role, dados_grupo in grupos_roles:
        print(f"[{contador_role}/{total_roles}] Processando Role: {nome_role}...")
        
        lista_to_grant_completa = []
        detalhes_recursos_completa = []
        
        # 1. Monta a lista com TODOS os recursos da role
        for index, linha in dados_grupo.iterrows():
            recurso = linha["Recurso"]
            acao = linha["Acao"]
            
            if recurso:
                lista_to_grant_completa.append({
                    "resource": recurso,
                    "action": acao
                })
                detalhes_recursos_completa.append(f"{recurso} ({acao})")
        
        total_recursos = len(lista_to_grant_completa)
        total_lotes = math.ceil(total_recursos / TAMANHO_LOTE)
        
        # 2. O FATIAMENTO (Chunking): Loop que pula de 100 em 100
        for i in range(0, total_recursos, TAMANHO_LOTE):
            numero_lote = (i // TAMANHO_LOTE) + 1
            
            # Pega apenas a fatia (chunk) da vez
            chunk_to_grant = lista_to_grant_completa[i : i + TAMANHO_LOTE]
            chunk_detalhes = detalhes_recursos_completa[i : i + TAMANHO_LOTE]
            
            print(f"   -> Enviando Lote {numero_lote}/{total_lotes} ({len(chunk_to_grant)} recursos)...")
            
            payload = {
                "roles": [nome_role],
                "toGrant": chunk_to_grant
            }
            
            try:
                resposta = requests.post(URL, json=payload, headers=HEADERS)
                status_code = resposta.status_code
                
                if status_code == 200:
                    mensagem_retorno = "Sucesso"
                else:
                    mensagem_retorno = resposta.text 
                    
            except Exception as e:
                status_code = "Falha de Conexão"
                mensagem_retorno = str(e)
                
            # Registra no log especificando qual lote é
            dados_log.append({
                "Nome_Role": f"{nome_role} (Lote {numero_lote}/{total_lotes})",
                "Qtd_Recursos_Enviados": len(chunk_to_grant),
                "Recursos_Detalhados": " | ".join(chunk_detalhes), 
                "Status_HTTP": status_code,
                "Mensagem": mensagem_retorno,
                "Data_Hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            })
            
        contador_role += 1

    print("\nProcessamento concluído! Gerando arquivo de log...")
    df_log = pd.DataFrame(dados_log)
    
    try:
        df_log.to_excel(ARQUIVO_LOG, index=False)
        print(f"Arquivo de log salvo com sucesso em: {ARQUIVO_LOG}")
    except Exception as e:
        print(f"Erro ao salvar arquivo de log: {e}")

if __name__ == "__main__":
    main()