
from pyrfc import Connection

try:
    conn = Connection(
        user='S-SDKRFC',
        passwd='RFC@2026sdk&&15',
        ashost='10.200.2.73',
        sysnr='00',
        client='100',
        lang='PT'
    )

    result = conn.call(
        'RFC_READ_TABLE',
        QUERY_TABLE='T000',
        ROWCOUNT=1,
        DELIMITER=';'
    )

    print("✅ Conexão RFC OK")
    print(result['DATA'])

except Exception as e:
    print("❌ Erro na conexão RFC")
    print(str(e))
