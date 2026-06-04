from pyrfc import Connection
import pandas as pd

conn = Connection(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.10",
    sysnr="00",
    client="310",
    lang="PT"
)

# =========================
# CONFIGURAÇÕES
# =========================
EXCEL_PATH = r"C:\python_scripts\Planilhas\BPs.xlsx"

# =========================
# LEITURA DO EXCEL
# =========================
df = pd.read_excel(EXCEL_PATH)

for _, row in df.iterrows():

    bp = str(row['BP'])
    address_id = str(row['ADDRESS_ID'])

    telefone_novo = str(row['TELEFONE']) if pd.notna(row['TELEFONE']) else ''
    email_novo = str(row['EMAIL']) if pd.notna(row['EMAIL']) else ''

    print(f'Processando BP {bp}...')

    # =========================
    # BUSCAR DADOS ATUAIS
    # =========================
    detalhe = conn.call(
        'BAPI_BUPA_ADDRESS_GETDETAIL',
        BUSINESSPARTNER=bp,
        ADDRESSGUID=address_id
    )

    # =========================
    # TELEFONE
    # =========================
    telefone_update = []

    if telefone_novo:  # só entra se tiver valor
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
    # EMAIL
    # =========================
    email_update = []

    if email_novo:  # só entra se tiver valor
        for mail in detalhe.get('E_MAIL', []):
            email_update.append({
                'CONSNUMBER': mail['CONSNUMBER'],
                'E_MAIL': email_novo,
                'STD_NO': mail.get('STD_NO', 'X'),
                'VALID_FROM': mail.get('VALID_FROM', '20250101'),
                'VALID_TO': mail.get('VALID_TO', '99991231')
            })

    # =========================
    # CHAMAR BAPI (SÓ COM O QUE EXISTE)
    # =========================
    parametros = {
        'BUSINESSPARTNER': bp,
        'ADDRESSGUID': address_id
    }

    if telefone_update:
        parametros['TELEPHONE'] = telefone_update

    if email_update:
        parametros['E_MAIL'] = email_update

    # Se não tiver nada pra alterar, pula
    if not telefone_update and not email_update:
        print(f'BP {bp} ignorado (sem dados)')
        continue

    conn.call('BAPI_BUPA_ADDRESS_CHANGE', **parametros)
    conn.call('BAPI_TRANSACTION_COMMIT')

    print(f'BP {bp} atualizado com sucesso')