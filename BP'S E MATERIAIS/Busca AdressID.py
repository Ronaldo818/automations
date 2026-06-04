from pyrfc import Connection
import pandas as pd

# =========================
# CONFIG SAP
# =========================
conn = Connection(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.10",
    sysnr="00",
    client="310",
    lang="PT"
)

# =========================
# FUNÇÕES AUXILIARES
# =========================

def obter_addrnumber(bp):
    resultado = conn.call(
        'RFC_READ_TABLE',
        QUERY_TABLE='BUT020',
        DELIMITER=';',
        OPTIONS=[{'TEXT': f"PARTNER = '{bp}'"}],
        FIELDS=[{'FIELDNAME': 'ADDRNUMBER'}]
    )

    dados = resultado.get('DATA', [])
    if dados:
        return dados[0]['WA'].strip()
    return None


def obter_guid(addrnumber):
    resultado = conn.call(
        'RFC_READ_TABLE',
        QUERY_TABLE='ADRC',
        DELIMITER=';',
        OPTIONS=[{'TEXT': f"ADDRNUMBER = '{addrnumber}'"}],
        FIELDS=[{'FIELDNAME': 'ADDR_GUID'}]
    )

    dados = resultado.get('DATA', [])
    if dados:
        return dados[0]['WA'].strip()
    return None

# =========================
# CONFIGURAÇÕES
# =========================
EXCEL_PATH = r"C:\python_scripts\Planilhas\BPs.xlsx"

# =========================
# LER EXCEL
# =========================
df = pd.read_excel(EXCEL_PATH)

# =========================
# PROCESSAMENTO
# =========================
for _, row in df.iterrows():

    bp = str(row['BP']).zfill(10)

    telefone_novo = str(row['TELEFONE']).strip() if pd.notna(row['TELEFONE']) else ''
    email_novo = str(row['EMAIL']).strip() if pd.notna(row['EMAIL']) else ''

    if not telefone_novo and not email_novo:
        print(f'BP {bp} ignorado (sem dados)')
        continue

    print(f'\nProcessando BP {bp}...')

    # =========================
    # 1. BUSCAR ENDEREÇO
    # =========================
    addrnumber = obter_addrnumber(bp)

    if not addrnumber:
        print('ADDRESS não encontrado')
        continue

    guid = obter_guid(addrnumber)

    if not guid:
        print('GUID não encontrado')
        continue

    # =========================
    # 2. BUSCAR DADOS ATUAIS
    # =========================
    detalhe = conn.call(
        'BAPI_BUPA_ADDRESS_GETDETAIL',
        BUSINESSPARTNER=bp,
        ADDRESSGUID=guid
    )

    # =========================
    # 3. TELEFONE
    # =========================
    telefone_update = []

    if telefone_novo:
        for tel in detalhe.get('TELEPHONE', []):
            telefone_update.append({
                'CONSNUMBER': tel['CONSNUMBER'],
                'TEL_NO': telefone_novo,
                'R_3_USER': tel.get('R_3_USER', ''),
                'MOBILE': tel.get('MOBILE', ''),
                'VALID_FROM': tel.get('VALID_FROM', '20250101'),
                'VALID_TO': tel.get('VALID_TO', '99991231')
            })

    # =========================
    # 4. EMAIL
    # =========================
    email_update = []

    if email_novo:
        for mail in detalhe.get('E_MAIL', []):
            email_update.append({
                'CONSNUMBER': mail['CONSNUMBER'],
                'E_MAIL': email_novo,
                'STD_NO': mail.get('STD_NO', 'X'),
                'VALID_FROM': mail.get('VALID_FROM', '20250101'),
                'VALID_TO': mail.get('VALID_TO', '99991231')
            })

    # =========================
    # 5. MONTAR PARAMETROS
    # =========================
    params = {
        'BUSINESSPARTNER': bp,
        'ADDRESSGUID': guid
    }

    if telefone_update:
        params['TELEPHONE'] = telefone_update

    if email_update:
        params['E_MAIL'] = email_update

    # =========================
    # 6. EXECUTAR
    # =========================
    try:
        retorno = conn.call('BAPI_BUPA_ADDRESS_CHANGE', **params)

        # Verificar retorno SAP
        if 'RETURN' in retorno:
            for msg in retorno['RETURN']:
                if msg['TYPE'] in ('E', 'A'):
                    print(f"❌ Erro SAP: {msg['MESSAGE']}")
                    raise Exception(msg['MESSAGE'])

        conn.call('BAPI_TRANSACTION_COMMIT')

        print('✅ Atualizado com sucesso')

    except Exception as e:
        print(f'❌ Erro: {e}')