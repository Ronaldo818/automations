import win32com.client
import pythoncom


def validar_sap():
    try:
        pythoncom.CoInitialize()

        print("Tentando conectar ao SAP GUI...")

        sap_gui = win32com.client.GetObject("SAPGUI")

        if not sap_gui:
            print("SAP GUI não encontrado.")
            return

        # GetScriptingEngine é PROPRIEDADE (sem parênteses)
        application = sap_gui.GetScriptingEngine

        if not application:
            print("Scripting NÃO está habilitado no SAP.")
            return

        print("SAP GUI encontrado e scripting ativo!")

        if application.Children.Count == 0:
            print("Nenhuma conexão ativa encontrada.")
            return

        print(f"Conexões encontradas: {application.Children.Count}")

        for i in range(application.Children.Count):
            connection = application.Children(i)
            print(f"\n📡 Conexão {i + 1}")

            try:
                print(f"   - Descrição: {connection.Description}")
            except:
                pass

            # USAR Sessions (mais confiável)
            if connection.Sessions.Count == 0:
                print("   ⚠ Nenhuma sessão ativa")
                continue

            print(f"   - Sessões: {connection.Sessions.Count}")

            for j in range(connection.Sessions.Count):
                session = connection.Sessions(j)

                print(f"      Sessão {j + 1} encontrada")
                print(f"         - Usuário: {session.Info.User}")
                print(f"         - Sistema: {session.Info.SystemName}")
                print(f"         - Transação: {session.Info.Transaction}")

        print("\n Validação concluída com sucesso!")

    except Exception as e:
        print(f"Erro ao conectar no SAP GUI: {e}")


if __name__ == "__main__":
    validar_sap()