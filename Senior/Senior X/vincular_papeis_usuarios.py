import requests
import pandas as pd
from datetime import datetime
import math

# =========================
# CONFIG
# =========================
ARQUIVO_ENTRADA = r"C:\Users\junio\OneDrive\Área de Trabalho\Documentos\Scripts Github\automations\Planilhas\Vinculo_usuarios.xlsx"
ARQUIVO_LOG = r"C:\Users\junio\OneDrive\Área de Trabalho\Documentos\Scripts Github\automations\Planilhas\Vinculo_usuarios_logs.xlsx"
TOKEN_SENIOR = "siDaL5h6vs5Mfv10Gr2s2LIr5apxNOgC"

TAMANHO_LOTE = 100 

URL = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/platform/authorization/actions/reassignUsers"

HEADERS = {
    "Authorization": f"Bearer {TOKEN_SENIOR}",
    "Content-Type": "application/json"
}

def main():
    print("Iniciando leitura da planilha de usuários...")
    
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
        
        lista_usuarios_completa = []
        
        for index, linha in dados_grupo.iterrows():
            usuario = linha["Usuario"]
            
            if usuario:
                lista_usuarios_completa.append(usuario)
        
        total_usuarios = len(lista_usuarios_completa)
        
        if total_usuarios == 0:
            contador_role += 1
            continue
            
        total_lotes = math.ceil(total_usuarios / TAMANHO_LOTE)
        
        for i in range(0, total_usuarios, TAMANHO_LOTE):
            numero_lote = (i // TAMANHO_LOTE) + 1
            
            chunk_usuarios = lista_usuarios_completa[i : i + TAMANHO_LOTE]
            
            print(f"   -> Enviando Lote {numero_lote}/{total_lotes} ({len(chunk_usuarios)} usuários)...")
            
            payload = {
                "roles": [nome_role],
                "toAssign": chunk_usuarios,
                "toUnassign": []
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
                
            # Registra no log
            dados_log.append({
                "Nome_Role": f"{nome_role} (Lote {numero_lote}/{total_lotes})",
                "Qtd_Usuarios_Enviados": len(chunk_usuarios),
                "Usuarios_Detalhados": " | ".join(chunk_usuarios), 
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