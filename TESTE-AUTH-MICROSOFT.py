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

REQUEST_TIMEOUT = 30


# ============================================================
# EXCEÇÕES
# ============================================================

class GraphError(Exception):
    def __init__(
        self,
        status_code,
        method,
        url,
        response_text="",
    ):
        self.status_code = status_code
        self.method = method
        self.url = url
        self.response_text = response_text

        mensagem = (
            f"Microsoft Graph retornou HTTP {status_code} "
            f"em {method} {url}"
        )

        super().__init__(mensagem)


class GraphNotFoundError(GraphError):
    pass


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def encerrar_com_erro(mensagem, detalhes=None):
    print()
    print("=" * 60)
    print("ERRO")
    print("=" * 60)
    print()
    print(mensagem)

    if detalhes:
        print()
        print(detalhes)

    sys.exit(1)


def montar_url_graph(endpoint):
    if endpoint.startswith("https://"):
        return endpoint

    if not endpoint.startswith("/"):
        endpoint = f"/{endpoint}"

    return f"{GRAPH_BASE_URL}{endpoint}"


def chamar_graph(
    method,
    endpoint,
    access_token,
    params=None,
    json_body=None,
):
    url = montar_url_graph(endpoint)

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
        timeout=REQUEST_TIMEOUT,
    )

    if response.status_code == 404:
        raise GraphNotFoundError(
            status_code=response.status_code,
            method=method,
            url=url,
            response_text=response.text,
        )

    if not response.ok:
        raise GraphError(
            status_code=response.status_code,
            method=method,
            url=url,
            response_text=response.text,
        )

    if response.status_code == 204:
        return None

    if not response.content:
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
            json.dumps(
                flow,
                indent=2,
                ensure_ascii=False,
            ),
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
        detalhes = {
            "error": result.get("error"),
            "error_description": result.get(
                "error_description"
            ),
            "correlation_id": result.get(
                "correlation_id"
            ),
        }

        encerrar_com_erro(
            "Não foi possível obter o token Microsoft.",
            json.dumps(
                detalhes,
                indent=2,
                ensure_ascii=False,
            ),
        )

    print("Token Microsoft obtido com sucesso.")

    return result["access_token"]


# ============================================================
# IDENTIFICAÇÃO DA CONTA
# ============================================================

def identificar_usuario(access_token):
    print()
    print("Consultando a conta Microsoft...")

    usuario = chamar_graph(
        method="GET",
        endpoint="/me",
        access_token=access_token,
        params={
            "$select": (
                "id,displayName,mail,userPrincipalName"
            )
        },
    )

    conta = (
        usuario.get("mail")
        or usuario.get("userPrincipalName")
        or "Não informada"
    )

    print(
        f"Nome da conta: "
        f"{usuario.get('displayName', 'Não informado')}"
    )
    print(f"Conta conectada: {conta}")

    return usuario


# ============================================================
# DIAGNÓSTICO DO ONEDRIVE
# ============================================================

def obter_drive(access_token):
    print()
    print("Consultando o OneDrive da conta...")

    drive = chamar_graph(
        method="GET",
        endpoint="/me/drive",
        access_token=access_token,
        params={
            "$select": (
                "id,driveType,name,webUrl,owner"
            )
        },
    )

    print("OneDrive acessado com sucesso.")
    print(f"ID do drive: {drive.get('id')}")
    print(
        f"Tipo do drive: "
        f"{drive.get('driveType', 'Não informado')}"
    )

    if drive.get("name"):
        print(f"Nome do drive: {drive.get('name')}")

    return drive


# ============================================================
# PAGINAÇÃO DO MICROSOFT GRAPH
# ============================================================

def obter_todas_paginas(
    access_token,
    endpoint,
    params=None,
):
    itens = []
    proxima_url = endpoint
    parametros = params

    while proxima_url:
        resultado = chamar_graph(
            method="GET",
            endpoint=proxima_url,
            access_token=access_token,
            params=parametros,
        )

        # O nextLink já possui os parâmetros.
        parametros = None

        itens.extend(
            resultado.get("value", [])
        )

        proxima_url = resultado.get("@odata.nextLink")

    return itens


# ============================================================
# ESTRATÉGIA 1: ACESSO DIRETO NA RAIZ
# ============================================================

def procurar_diretamente_na_raiz(access_token):
    print("Tentativa 1: procurando diretamente na raiz...")

    nome_codificado = quote(
        EXCEL_FILE_NAME,
        safe="",
    )

    endpoint = (
        f"/me/drive/root:/{nome_codificado}"
    )

    try:
        arquivo = chamar_graph(
            method="GET",
            endpoint=endpoint,
            access_token=access_token,
            params={
                "$select": (
                    "id,name,size,webUrl,file,folder,"
                    "parentReference,lastModifiedDateTime"
                )
            },
        )

    except GraphNotFoundError:
        print("O arquivo não está diretamente na raiz.")
        return []

    if (
        arquivo.get("file") is not None
        and arquivo.get("name", "").casefold()
        == EXCEL_FILE_NAME.casefold()
    ):
        print("Arquivo encontrado diretamente na raiz.")
        return [arquivo]

    return []


# ============================================================
# ESTRATÉGIA 2: PESQUISA DO ONEDRIVE
# ============================================================

def procurar_com_pesquisa(access_token):
    print()
    print("Tentativa 2: usando a pesquisa do OneDrive...")

    # Aspas simples precisam ser duplicadas em expressões OData.
    nome_pesquisa = EXCEL_FILE_NAME.replace(
        "'",
        "''",
    )

    endpoint = (
        f"/me/drive/root/"
        f"search(q='{nome_pesquisa}')"
    )

    itens = obter_todas_paginas(
        access_token=access_token,
        endpoint=endpoint,
        params={
            "$select": (
                "id,name,size,webUrl,file,folder,"
                "parentReference,lastModifiedDateTime"
            ),
            "$top": "200",
        },
    )

    print(
        f"Resultados retornados pela pesquisa: "
        f"{len(itens)}"
    )

    correspondencias = [
        item
        for item in itens
        if (
            item.get("file") is not None
            and item.get("name", "").casefold()
            == EXCEL_FILE_NAME.casefold()
        )
    ]

    return correspondencias


# ============================================================
# ESTRATÉGIA 3: VARREDURA RECURSIVA
# ============================================================

def procurar_recursivamente(
    access_token,
    folder_id=None,
    caminho_atual="/",
    pastas_visitadas=None,
):
    if pastas_visitadas is None:
        pastas_visitadas = set()

    if folder_id:
        if folder_id in pastas_visitadas:
            return []

        pastas_visitadas.add(folder_id)

        endpoint = (
            f"/me/drive/items/{folder_id}/children"
        )
    else:
        endpoint = "/me/drive/root/children"

    encontrados = []

    itens = obter_todas_paginas(
        access_token=access_token,
        endpoint=endpoint,
        params={
            "$select": (
                "id,name,size,webUrl,file,folder,"
                "parentReference,lastModifiedDateTime"
            ),
            "$top": "200",
        },
    )

    for item in itens:
        nome = item.get(
            "name",
            "Item sem nome",
        )

        caminho_item = (
            f"{caminho_atual.rstrip('/')}/{nome}"
        )

        tipo = (
            "Pasta"
            if item.get("folder") is not None
            else "Arquivo"
        )

        print(
            f"- {tipo}: {caminho_item}"
        )

        if (
            item.get("file") is not None
            and nome.casefold()
            == EXCEL_FILE_NAME.casefold()
        ):
            encontrados.append(item)

        if item.get("folder") is not None:
            encontrados.extend(
                procurar_recursivamente(
                    access_token=access_token,
                    folder_id=item.get("id"),
                    caminho_atual=caminho_item,
                    pastas_visitadas=pastas_visitadas,
                )
            )

    return encontrados


# ============================================================
# REMOÇÃO DE DUPLICATAS
# ============================================================

def remover_duplicatas_por_id(itens):
    itens_unicos = {}

    for item in itens:
        item_id = item.get("id")

        if item_id:
            itens_unicos[item_id] = item

    return list(itens_unicos.values())


# ============================================================
# EXIBIÇÃO DO RESULTADO
# ============================================================

def apresentar_arquivo(arquivo):
    parent = arquivo.get(
        "parentReference",
        {},
    )

    print()
    print("=" * 60)
    print("EXCEL ENCONTRADO!")
    print("=" * 60)
    print()
    print(f"Nome: {arquivo.get('name')}")
    print(f"ID do arquivo: {arquivo.get('id')}")
    print(
        f"ID do drive: "
        f"{parent.get('driveId', 'Não informado')}"
    )
    print(
        f"Caminho: "
        f"{parent.get('path', 'Não informado')}"
    )
    print(
        f"Tamanho: "
        f"{arquivo.get('size', 'Não informado')} bytes"
    )
    print(
        "Última modificação: "
        f"{arquivo.get('lastModifiedDateTime', 'Não informada')}"
    )

    if arquivo.get("webUrl"):
        print("Link do arquivo disponível no Microsoft Graph.")

    print()
    print("=" * 60)
    print("TESTE DO ONEDRIVE CONCLUÍDO")
    print("=" * 60)


# ============================================================
# LOCALIZAÇÃO DO EXCEL
# ============================================================

def localizar_excel(access_token):
    print()
    print("=" * 60)
    print("PROCURANDO O EXCEL NO ONEDRIVE")
    print("=" * 60)
    print()
    print(f"Arquivo esperado: {EXCEL_FILE_NAME}")
    print()

    correspondencias = []

    # Estratégia 1
    correspondencias.extend(
        procurar_diretamente_na_raiz(
            access_token=access_token,
        )
    )

    # Estratégia 2
    if not correspondencias:
        correspondencias.extend(
            procurar_com_pesquisa(
                access_token=access_token,
            )
        )

    # Estratégia 3
    if not correspondencias:
        print()
        print(
            "Tentativa 3: percorrendo as pastas "
            "do OneDrive..."
        )
        print()

        correspondencias.extend(
            procurar_recursivamente(
                access_token=access_token,
            )
        )

    correspondencias = remover_duplicatas_por_id(
        correspondencias
    )

    if not correspondencias:
        encerrar_com_erro(
            f"O arquivo {EXCEL_FILE_NAME} não foi encontrado "
            "no OneDrive retornado pelo Microsoft Graph.",
            (
                "A autenticação funcionou e o drive pôde ser "
                "consultado, mas nenhuma cópia com o nome exato "
                "foi localizada. Confira no OneDrive se o arquivo "
                "está em 'Meus arquivos' da mesma conta utilizada "
                "na autorização."
            ),
        )

    if len(correspondencias) > 1:
        print()
        print("=" * 60)
        print("ARQUIVOS DUPLICADOS ENCONTRADOS")
        print("=" * 60)
        print()

        for numero, item in enumerate(
            correspondencias,
            start=1,
        ):
            parent = item.get(
                "parentReference",
                {},
            )

            print(f"{numero}. {item.get('name')}")
            print(f"   ID: {item.get('id')}")
            print(
                f"   Caminho: "
                f"{parent.get('path', 'Não informado')}"
            )
            print(
                "   Última modificação: "
                f"{item.get('lastModifiedDateTime', 'Não informada')}"
            )
            print()

        encerrar_com_erro(
            "Mais de um arquivo com o nome oficial foi encontrado.",
            (
                "Mantenha apenas uma cópia oficial ou defina "
                "posteriormente um caminho fixo para a automação."
            ),
        )

    arquivo = correspondencias[0]

    apresentar_arquivo(
        arquivo=arquivo,
    )

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

        obter_drive(
            access_token=access_token,
        )

        localizar_excel(
            access_token=access_token,
        )

    except GraphNotFoundError as error:
        encerrar_com_erro(
            "Um recurso esperado não foi encontrado "
            "no Microsoft Graph.",
            (
                f"Método: {error.method}\n"
                f"URL: {error.url}\n"
                f"Resposta: {error.response_text}"
            ),
        )

    except GraphError as error:
        encerrar_com_erro(
            "Falha ao consultar o Microsoft Graph.",
            (
                f"HTTP: {error.status_code}\n"
                f"Método: {error.method}\n"
                f"URL: {error.url}\n"
                f"Resposta: {error.response_text}"
            ),
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
