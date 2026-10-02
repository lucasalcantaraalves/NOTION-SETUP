import json
import os
import sys
from urllib.parse import quote

import msal
import requests


# ============================================================
# CONFIGURAÇÃO
# ============================================================

CLIENT_ID = os.environ.get("MS_CLIENT_ID")

if not CLIENT_ID:
    raise RuntimeError(
        "Secret MS_CLIENT_ID não encontrado nas variáveis de ambiente."
    )

AUTHORITY = "https://login.microsoftonline.com/consumers"

SCOPES = [
    "User.Read",
    "Files.ReadWrite",
    "Calendars.ReadWrite",
]

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

EXCEL_FILE_NAME = "Second Brain - Dados.xlsx"


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def encerrar_com_erro(mensagem, detalhes=None):
    print()
    print(f"ERRO: {mensagem}")

    if detalhes:
        print()
        print(detalhes)

    sys.exit(1)


def chamar_graph(
    method,
    endpoint,
    access_token,
    params=None,
    json_body=None,
):
    url = f"{GRAPH_BASE_URL}{endpoint}"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }

    response = requests.request(
        method=method,
        url=url,
        headers=headers,
        params=params,
        json=json_body,
        timeout=30,
    )

    if not response.ok:
        encerrar_com_erro(
            f"Microsoft Graph retornou HTTP {response.status_code}.",
            response.text,
        )

    if response.status_code == 204:
        return None

    return response.json()


# ============================================================
# AUTENTICAÇÃO MICROSOFT
# ============================================================

def autenticar():
    print("Iniciando autenticação Microsoft...")
    print()

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
    )

    flow = app.initiate_device_flow(
        scopes=SCOPES,
    )

    if "user_code" not in flow:
        encerrar_com_erro(
            "Não foi possível iniciar o Device Code Flow.",
            json.dumps(flow, indent=2),
        )

    print("=" * 60)
    print("AUTORIZAÇÃO NECESSÁRIA")
    print("=" * 60)
    print()
    print(flow.get("message"))
    print()
    print("=" * 60)
    print()

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        encerrar_com_erro(
            "Não foi possível obter o token Microsoft.",
            json.dumps(
                {
                    "error": result.get("error"),
                    "error_description": result.get(
                        "error_description"
                    ),
                },
                indent=2,
            ),
        )

    print("Token Microsoft obtido com sucesso.")

    return result["access_token"]


# ============================================================
# TESTE DA CONTA
# ============================================================

def identificar_usuario(access_token):
    print()
    print("Consultando a conta Microsoft...")

    user = chamar_graph(
        method="GET",
        endpoint="/me",
        access_token=access_token,
    )

    conta = (
        user.get("mail")
        or user.get("userPrincipalName")
        or "Não informada"
    )

    print(f"Conta conectada: {conta}")


# ============================================================
# LOCALIZAÇÃO DO EXCEL NO ONEDRIVE
# ============================================================

def localizar_excel(access_token):
    print()
    print("=" * 60)
    print("PROCURANDO O EXCEL NO ONEDRIVE")
    print("=" * 60)
    print()
    print(f"Arquivo esperado: {EXCEL_FILE_NAME}")
    print()

    # A busca é recursiva e pode encontrar o arquivo mesmo que
    # ele esteja dentro de uma pasta do OneDrive.
    nome_codificado = quote(
        EXCEL_FILE_NAME,
        safe="",
    )

    resultado = chamar_graph(
        method="GET",
        endpoint=f"/me/drive/root/search(q='{nome_codificado}')",
        access_token=access_token,
        params={
            "$select": (
                "id,name,size,webUrl,file,parentReference,"
                "lastModifiedDateTime"
            ),
            "$top": "100",
        },
    )

    itens = resultado.get("value", [])

    # A pesquisa do Graph pode retornar aproximações.
    # Por isso, filtramos pelo nome exato.
    correspondencias = [
        item
        for item in itens
        if item.get("name", "").casefold()
        == EXCEL_FILE_NAME.casefold()
        and item.get("file") is not None
    ]

    if not correspondencias:
        print("Arquivo não encontrado.")
        print()
        print("Resultados aproximados encontrados:")

        if not itens:
            print("- Nenhum resultado retornado pelo OneDrive.")
        else:
            for item in itens[:10]:
                print(f"- {item.get('name', 'Sem nome')}")

        sys.exit(1)

    if len(correspondencias) > 1:
        print(
            "ATENÇÃO: mais de um arquivo com o mesmo nome "
            "foi encontrado."
        )
        print()

        for numero, item in enumerate(
            correspondencias,
            start=1,
        ):
            parent = item.get("parentReference", {})
            caminho = parent.get("path", "Caminho não informado")

            print(f"{numero}. {item.get('name')}")
            print(f"   ID: {item.get('id')}")
            print(f"   Caminho: {caminho}")
            print()

        encerrar_com_erro(
            "Existem arquivos duplicados. "
            "Mantenha apenas a cópia oficial ou defina "
            "um caminho fixo."
        )

    arquivo = correspondencias[0]
    parent = arquivo.get("parentReference", {})

    caminho = parent.get(
        "path",
        "Caminho não informado",
    )

    print("EXCEL ENCONTRADO!")
    print()
    print(f"Nome: {arquivo.get('name')}")
    print(f"ID do arquivo: {arquivo.get('id')}")
    print(f"ID do drive: {parent.get('driveId', 'Não informado')}")
    print(f"Caminho: {caminho}")
    print(f"Tamanho: {arquivo.get('size', 'Não informado')} bytes")
    print(
        "Última modificação: "
        f"{arquivo.get('lastModifiedDateTime', 'Não informada')}"
    )
    print()
    print("=" * 60)
    print("TESTE DO ONEDRIVE CONCLUÍDO")
    print("=" * 60)

    return arquivo


# ============================================================
# EXECUÇÃO
# ============================================================

def main():
    try:
        access_token = autenticar()

        identificar_usuario(
            access_token=access_token,
        )

        localizar_excel(
            access_token=access_token,
        )

    except requests.RequestException as error:
        encerrar_com_erro(
            "Falha de comunicação com a Microsoft.",
            str(error),
        )

    except Exception as error:
        encerrar_com_erro(
            "Erro inesperado durante o teste.",
            str(error),
        )


if __name__ == "__main__":
    main()
