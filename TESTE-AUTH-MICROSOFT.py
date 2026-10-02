import os
import sys
import json
import msal
import requests


# ============================================================
# CONFIGURAÇÃO
# ============================================================

CLIENT_ID = os.environ.get("MS_CLIENT_ID")

if not CLIENT_ID:
    raise RuntimeError(
        "❌ Secret MS_CLIENT_ID não encontrado nas variáveis de ambiente."
    )

# O aplicativo foi criado somente para contas Microsoft pessoais.
AUTHORITY = "https://login.microsoftonline.com/consumers"

# Permissões delegadas configuradas no aplicativo.
# offline_access é tratado pelo fluxo de autenticação/token.
SCOPES = [
    "User.Read",
    "Files.ReadWrite",
    "Calendars.ReadWrite",
]


# ============================================================
# AUTENTICAÇÃO MICROSOFT
# ============================================================

def autenticar():
    print("🔐 Iniciando autenticação Microsoft...")
    print()

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
    )

    # Inicia o Device Code Flow.
    flow = app.initiate_device_flow(
        scopes=SCOPES
    )

    if "user_code" not in flow:
        print("❌ Não foi possível iniciar o Device Code Flow.")
        print(json.dumps(flow, indent=2))
        sys.exit(1)

    print("=" * 60)
    print("📱 AUTORIZAÇÃO NECESSÁRIA")
    print("=" * 60)
    print()
    print(flow.get("message"))
    print()
    print("=" * 60)
    print()

    # O GitHub Actions ficará aguardando enquanto você
    # autoriza o aplicativo pelo navegador.
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print("❌ Falha ao obter token Microsoft.")
        print()

        print(
            "Erro:",
            result.get("error")
        )

        print(
            "Descrição:",
            result.get("error_description")
        )

        sys.exit(1)

    print("✅ Token Microsoft obtido com sucesso.")
    print()

    return result["access_token"]


# ============================================================
# TESTE MICROSOFT GRAPH
# ============================================================

def testar_graph(access_token):
    print("🔗 Testando Microsoft Graph...")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }

    response = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers=headers,
        timeout=30,
    )

    if not response.ok:
        print(
            f"❌ Microsoft Graph retornou HTTP "
            f"{response.status_code}"
        )

        print(response.text)
        sys.exit(1)

    user = response.json()

    print()
    print("🎉 MICROSOFT GRAPH CONECTADO!")
    print()

    print(
        "👤 Nome:",
        user.get("displayName", "Não informado")
    )

    # Não é erro se algum destes campos não vier preenchido.
    conta = (
        user.get("mail")
        or user.get("userPrincipalName")
        or "Não informada"
    )

    print(
        "📧 Conta:",
        conta
    )

    print()
    print("=" * 60)
    print("✅ TESTE DE AUTENTICAÇÃO CONCLUÍDO")
    print("=" * 60)


# ============================================================
# EXECUÇÃO
# ============================================================

def main():
    try:
        access_token = autenticar()
        testar_graph(access_token)

    except requests.RequestException as error:
        print()
        print("❌ Erro de comunicação com Microsoft Graph:")
        print(str(error))
        sys.exit(1)

    except Exception as error:
        print()
        print("❌ Erro inesperado:")
        print(str(error))
        sys.exit(1)


if __name__ == "__main__":
    main()
