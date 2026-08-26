# Ambiente de homologação QAS 300

from pyrfc import Connection

conn = Connection(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.10",
    sysnr="00",
    client="310",
    lang="EN"
)

info = conn.call("RFC_SYSTEM_INFO")

for k, v in info.items():
    print(k, "=", v)


from pyrfc import Connection



# Ambiente de produção PRD 300

from pyrfc import Connection

conn = Connection(
    user="S-SDKRFC",
    passwd="RFC@2026sdk&&15",
    ashost="10.200.3.92",
    sysnr="00",
    client="310",
    lang="EN"
)

info = conn.call("RFC_SYSTEM_INFO")

for k, v in info.items():
    print(k, "=", v)    
