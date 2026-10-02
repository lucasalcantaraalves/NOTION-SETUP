import base64
import json
import os
import sys
from pathlib import Path
from urllib.parse import quote

import msal
import requests
from cryptography.fernet import Fernet, InvalidToken


# ============================================================
# CONFIGURACAO
# ============================================================

CLIENT_ID = os.environ.get("MS_CLIENT_ID")
CACHE_KEY_TEXT = os.environ.get("MS_CACHE_KEY")

if not CLIENT_ID:
    raise RuntimeError("Secret MS_CLIENT_ID nao encontrado.")

if not CACHE_KEY_TEXT:
    raise RuntimeError("Secret MS_CACHE_KEY nao encontrado.")

AUTHORITY = "https://login.microsoftonline.com/consumers"

SCOPES = [
    "User.Read",
    "Files.ReadWrite",
    "Calendars.ReadWrite",
]

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

REQUEST_TIMEOUT = 60

EXCEL_PATH = (
    "/Microsoft Copilot Chat Files/"
    "Copilot Notebook Uploads/"
    "Second Brain - Dados.xlsx"
)

TABLE_NAME = "tbAgenda"
TEST_ID = "TEST-0001"

CACHE_ENCRYPTED_PATH = Path(".auth/msal_cache.bin")


# ============================================================
# ERROS E HTTP
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

        super().__init__(
            f"Microsoft Graph retornou HTTP "
            f"{status_code} em {method} {url}"
        )


def encerrar_com_erro(
    mensagem,
    detalhes=None,
):

    print()
    print("=" * 70)
    print("ERRO")
    print("=" * 70)
    print(mensagem)

    if detalhes:
        print()
        print(detalhes)

    sys.exit(1)


def graph_request(
    method,
    endpoint,
    access_token,
    params=None,
    json_body=None,
    extra_headers=None,
):

    if endpoint.startswith("https://"):
        url = endpoint
    else:
        url = f"{GRAPH_BASE_URL}{endpoint}"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }

    if extra_headers:
        headers.update(extra_headers)

    response = requests.request(
        method=method,
        url=url,
        headers=headers,
        params=params,
        json=json_body,
        timeout=REQUEST_TIMEOUT,
    )

    if not response.ok:
        raise GraphError(
            response.status_code,
            method,
            url,
            response.text,
        )

    if (
        response.status_code == 204
        or not response.content
    ):
        return None

    return response.json()


# ============================================================
# CACHE MSAL CRIPTOGRAFADO
# ============================================================

def normalizar_chave_fernet(valor):

    import hashlib

    raw = valor.encode("utf-8")

    try:
        Fernet(raw)
        return raw

    except Exception:

        digest = hashlib.sha256(
            raw
        ).digest()

        return base64.urlsafe_b64encode(
            digest
        )


def carregar_cache():

    CACHE_ENCRYPTED_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    cache = msal.SerializableTokenCache()

    if not CACHE_ENCRYPTED_PATH.exists():

        print(
            "Cache de autenticacao ainda nao existe."
        )

        return cache

    fernet = Fernet(
        normalizar_chave_fernet(
            CACHE_KEY_TEXT
        )
    )

    try:

        encrypted = (
            CACHE_ENCRYPTED_PATH.read_bytes()
        )

        serialized = (
            fernet.decrypt(
                encrypted
            )
            .decode("utf-8")
        )

        cache.deserialize(
            serialized
        )

        print(
            "Cache de autenticacao restaurado "
            "e descriptografado."
        )

        return cache

    except InvalidToken:

        encerrar_com_erro(
            "Nao foi possivel descriptografar "
            "o cache MSAL.",
            "Verifique se o secret MS_CACHE_KEY "
            "continua com o mesmo valor.",
        )

    except Exception as error:

        encerrar_com_erro(
            "Falha ao carregar o cache MSAL.",
            str(error),
        )


def salvar_cache(cache):

    CACHE_ENCRYPTED_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    serialized = (
        cache.serialize()
        .encode("utf-8")
    )

    fernet = Fernet(
        normalizar_chave_fernet(
            CACHE_KEY_TEXT
        )
    )

    encrypted = fernet.encrypt(
        serialized
    )

    CACHE_ENCRYPTED_PATH.write_bytes(
        encrypted
    )

    print(
        f"Cache criptografado salvo em "
        f"{CACHE_ENCRYPTED_PATH}."
    )


# ============================================================
# AUTENTICACAO
# ============================================================

def autenticar():

    cache = carregar_cache()

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    accounts = app.get_accounts()

    result = None

    if accounts:

        print(
            "Tentando autenticacao silenciosa..."
        )

        result = app.acquire_token_silent(
            scopes=SCOPES,
            account=accounts[0],
        )

    if (
        result
        and "access_token" in result
    ):

        print(
            "Autenticacao silenciosa "
            "concluida com sucesso."
        )

        salvar_cache(cache)

        return result["access_token"]

    print(
        "Autenticacao silenciosa indisponivel."
    )

    print(
        "Iniciando Device Code Flow "
        "como fallback..."
    )

    print()

    flow = app.initiate_device_flow(
        scopes=SCOPES
    )

    if "user_code" not in flow:

        encerrar_com_erro(
            "Nao foi possivel iniciar "
            "o Device Code Flow.",
            json.dumps(
                flow,
                indent=2,
                ensure_ascii=False,
            ),
        )

    print("=" * 70)
    print("AUTORIZACAO NECESSARIA")
    print("=" * 70)
    print(flow.get("message"))
    print("=" * 70)
    print()

    result = (
        app.acquire_token_by_device_flow(
            flow
        )
    )

    if "access_token" not in result:

        encerrar_com_erro(
            "Nao foi possivel obter "
            "o token Microsoft.",
            json.dumps(
                {
                    "error":
                        result.get("error"),

                    "error_description":
                        result.get(
                            "error_description"
                        ),

                    "correlation_id":
                        result.get(
                            "correlation_id"
                        ),
                },
                indent=2,
                ensure_ascii=False,
            ),
        )

    salvar_cache(cache)

    print(
        "Device Code Flow concluido "
        "e cache criado."
    )

    return result["access_token"]


# ============================================================
# CONTA MICROSOFT
# ============================================================

def identificar_usuario(
    access_token
):

    usuario = graph_request(
        "GET",
        "/me",
        access_token,
        params={
            "$select":
                "id,displayName,"
                "mail,userPrincipalName"
        },
    )

    conta = (
        usuario.get("mail")
        or usuario.get(
            "userPrincipalName"
        )
    )

    print(
        f"Conta conectada: {conta}"
    )


# ============================================================
# LOCALIZAR EXCEL
# ============================================================

def localizar_excel_por_caminho(
    access_token
):

    caminho_codificado = quote(
        EXCEL_PATH,
        safe="/",
    )

    endpoint = (
        f"/me/drive/root:"
        f"{caminho_codificado}"
    )

    print()
    print(
        "Localizando o Excel "
        "pelo caminho fixo..."
    )

    arquivo = graph_request(
        "GET",
        endpoint,
        access_token,
        params={
            "$select": (
                "id,name,size,"
                "parentReference,"
                "lastModifiedDateTime,"
                "webUrl"
            )
        },
    )

    print(
        f"Excel encontrado: "
        f"{arquivo.get('name')}"
    )

    print(
        f"Item ID atual: "
        f"{arquivo.get('id')}"
    )

    print(
        "Ultima modificacao: "
        f"{arquivo.get(
            'lastModifiedDateTime',
            'nao informada'
        )}"
    )

    return arquivo


# ============================================================
# SESSAO DO WORKBOOK
# ============================================================

def criar_sessao_workbook(
    access_token,
    item_id,
):

    endpoint = (
        f"/me/drive/items/"
        f"{item_id}/"
        f"workbook/createSession"
    )

    resultado = graph_request(
        "POST",
        endpoint,
        access_token,
        json_body={
            "persistChanges": False
        },
    )

    session_id = resultado.get(
        "id"
    )

    if not session_id:

        encerrar_com_erro(
            "O Graph nao retornou "
            "o ID da sessao do workbook."
        )

    return session_id


# ============================================================
# COLUNAS DA tbAgenda
# ============================================================

def obter_colunas_tbagenda(
    access_token,
    item_id,
    session_id,
):

    endpoint = (
        f"/me/drive/items/"
        f"{item_id}/workbook/"
        f"tables/{TABLE_NAME}/"
        f"headerRowRange"
    )

    resultado = graph_request(
        "GET",
        endpoint,
        access_token,
        extra_headers={
            "workbook-session-id":
                session_id
        },
    )

    values = resultado.get(
        "values",
        [],
    )

    if (
        not values
        or not values[0]
    ):

        encerrar_com_erro(
            f"A tabela {TABLE_NAME} "
            "nao retornou cabecalhos."
        )

    colunas = values[0]

    mapa = {
        nome: indice
        for indice, nome
        in enumerate(colunas)
    }

    print()
    print("=" * 70)

    print(
        f"COLUNAS DA {TABLE_NAME}"
    )

    print("=" * 70)

    for indice, nome in enumerate(
        colunas
    ):

        print(
            f"Indice {indice:02d} | "
            f"Coluna Excel "
            f"{indice + 1:02d} | "
            f"{nome}"
        )

    return colunas, mapa


# ============================================================
# LINHAS DA tbAgenda
# ============================================================

def obter_todas_linhas_tbagenda(
    access_token,
    item_id,
    session_id,
):

    endpoint = (
        f"/me/drive/items/"
        f"{item_id}/workbook/"
        f"tables/{TABLE_NAME}/rows"
    )

    linhas = []

    proxima_url = endpoint

    params = {
        "$top": "200"
    }

    while proxima_url:

        resultado = graph_request(
            "GET",
            proxima_url,
            access_token,
            params=params,
            extra_headers={
                "workbook-session-id":
                    session_id
            },
        )

        params = None

        linhas.extend(
            resultado.get(
                "value",
                [],
            )
        )

        proxima_url = (
            resultado.get(
                "@odata.nextLink"
            )
        )

    return linhas


# ============================================================
# LOCALIZAR TEST-0001
# ============================================================

def analisar_linhas(
    colunas,
    mapa_colunas,
    linhas,
):

    print()
    print("=" * 70)

    print(
        f"LEITURA DA {TABLE_NAME}"
    )

    print("=" * 70)

    print(
        f"Quantidade de linhas "
        f"de dados: {len(linhas)}"
    )

    if not linhas:

        print(
            "A tabela existe e foi lida, "
            "mas ainda nao possui "
            "linhas de dados."
        )

        return

    indice_id = mapa_colunas.get(
        "ID"
    )

    if indice_id is None:

        encerrar_com_erro(
            "A coluna ID nao existe "
            "na tbAgenda."
        )

    encontrada = None

    for posicao, linha in enumerate(
        linhas
    ):

        values_matrix = linha.get(
            "values",
            [],
        )

        values = (
            values_matrix[0]
            if values_matrix
            else []
        )

        registro = {
            coluna:
                values[indice]
                if indice < len(values)
                else None

            for indice, coluna
            in enumerate(colunas)
        }

        print()

        print(
            f"Linha logica {posicao} | "
            f"Indice Graph "
            f"{linha.get(
                'index',
                'nao informado'
            )} | "
            f"ID: "
            f"{registro.get('ID')} | "
            f"Nome: "
            f"{registro.get('Nome')}"
        )

        if (
            str(
                registro.get(
                    "ID",
                    "",
                )
            )
            .strip()
            .casefold()
            == TEST_ID.casefold()
        ):

            encontrada = {
                "posicao":
                    posicao,

                "graph_index":
                    linha.get("index"),

                "registro":
                    registro,
            }

    print()

    if encontrada:

        print(
            f"{TEST_ID} encontrado "
            "com sucesso!"
        )

        print(
            "Posicao na resposta: "
            f"{encontrada['posicao']}"
        )

        print(
            "Indice da linha no Graph: "
            f"{encontrada[
                'graph_index'
            ]}"
        )

        print(
            json.dumps(
                encontrada[
                    "registro"
                ],
                indent=2,
                ensure_ascii=False,
            )
        )

    else:

        print(
            f"{TEST_ID} nao foi "
            "encontrado entre "
            "as linhas da tabela."
        )


# ============================================================
# EXECUCAO
# ============================================================

def main():

    try:

        access_token = autenticar()

        identificar_usuario(
            access_token
        )

        arquivo = (
            localizar_excel_por_caminho(
                access_token
            )
        )

        item_id = arquivo["id"]

        session_id = (
            criar_sessao_workbook(
                access_token,
                item_id,
            )
        )

        print(
            "Sessao temporaria "
            "do workbook criada."
        )

        colunas, mapa_colunas = (
            obter_colunas_tbagenda(
                access_token,
                item_id,
                session_id,
            )
        )

        linhas = (
            obter_todas_linhas_tbagenda(
                access_token,
                item_id,
                session_id,
            )
        )

        analisar_linhas(
            colunas,
            mapa_colunas,
            linhas,
        )

        print()
        print("=" * 70)
        print("TESTE CONCLUIDO")
        print("=" * 70)

        print(
            "Autenticacao: OK"
        )

        print(
            "Localizacao do Excel: OK"
        )

        print(
            f"Leitura da "
            f"{TABLE_NAME}: OK"
        )

    except GraphError as error:

        encerrar_com_erro(
            "Falha ao consultar "
            "o Microsoft Graph.",
            (
                f"HTTP: "
                f"{error.status_code}\n"

                f"Metodo: "
                f"{error.method}\n"

                f"URL: "
                f"{error.url}\n"

                f"Resposta: "
                f"{error.response_text}"
            ),
        )

    except requests.RequestException as error:

        encerrar_com_erro(
            "Falha de comunicacao "
            "com a Microsoft.",
            str(error),
        )

    except Exception as error:

        encerrar_com_erro(
            "Erro inesperado.",
            repr(error),
        )


if __name__ == "__main__":
    main()
