import base64
import hashlib
import json
import os
import sys
import uuid
from datetime import datetime, timedelta
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
    raise RuntimeError(
        "Secret MS_CLIENT_ID nao encontrado."
    )

if not CACHE_KEY_TEXT:
    raise RuntimeError(
        "Secret MS_CACHE_KEY nao encontrado."
    )

AUTHORITY = (
    "https://login.microsoftonline.com/consumers"
)

SCOPES = [
    "User.Read",
    "Files.ReadWrite",
    "Calendars.ReadWrite",
]

GRAPH_BASE_URL = (
    "https://graph.microsoft.com/v1.0"
)

REQUEST_TIMEOUT = 60

EXCEL_PATH = (
    "/Microsoft Copilot Chat Files/"
    "Copilot Notebook Uploads/"
    "Second Brain - Dados.xlsx"
)

TABLE_NAME = "tbAgenda"
TARGET_ID = "TEST-0001"

OUTLOOK_TIME_ZONE = (
    "E. South America Standard Time"
)

CACHE_ENCRYPTED_PATH = Path(
    ".auth/msal_cache.bin"
)


# ============================================================
# ERROS
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


# ============================================================
# MICROSOFT GRAPH
# ============================================================

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
        "Authorization": (
            f"Bearer {access_token}"
        ),
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
            status_code=response.status_code,
            method=method,
            url=url,
            response_text=response.text,
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
            "Cache de autenticacao ainda "
            "nao existe."
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
            fernet.decrypt(encrypted)
            .decode("utf-8")
        )

        cache.deserialize(serialized)

        print(
            "Cache de autenticacao restaurado "
            "e descriptografado."
        )

        return cache

    except InvalidToken:
        encerrar_com_erro(
            "Nao foi possivel descriptografar "
            "o cache MSAL.",
            (
                "Verifique se o secret "
                "MS_CACHE_KEY continua "
                "com o mesmo valor."
            ),
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
                    "error": result.get(
                        "error"
                    ),
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

def identificar_usuario(access_token):
    usuario = graph_request(
        method="GET",
        endpoint="/me",
        access_token=access_token,
        params={
            "$select": (
                "id,displayName,"
                "mail,userPrincipalName"
            )
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
# EXCEL NO ONEDRIVE
# ============================================================

def localizar_excel_por_caminho(
    access_token,
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
        method="GET",
        endpoint=endpoint,
        access_token=access_token,
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
        f"{item_id}/workbook/createSession"
    )

    resultado = graph_request(
        method="POST",
        endpoint=endpoint,
        access_token=access_token,
        json_body={
            "persistChanges": True
        },
    )

    session_id = resultado.get("id")

    if not session_id:
        encerrar_com_erro(
            "O Graph nao retornou "
            "o ID da sessao do workbook."
        )

    print(
        "Sessao persistente "
        "do workbook criada."
    )

    return session_id


def fechar_sessao_workbook(
    access_token,
    item_id,
    session_id,
):
    endpoint = (
        f"/me/drive/items/"
        f"{item_id}/workbook/closeSession"
    )

    graph_request(
        method="POST",
        endpoint=endpoint,
        access_token=access_token,
        extra_headers={
            "workbook-session-id":
                session_id
        },
    )

    print(
        "Sessao do workbook encerrada."
    )


# ============================================================
# LEITURA DA tbAgenda
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
        method="GET",
        endpoint=endpoint,
        access_token=access_token,
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

    return colunas, mapa


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
            method="GET",
            endpoint=proxima_url,
            access_token=access_token,
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

        proxima_url = resultado.get(
            "@odata.nextLink"
        )

    return linhas


def localizar_registro(
    colunas,
    linhas,
    target_id,
):
    for posicao, linha in enumerate(
        linhas
    ):
        matriz = linha.get(
            "values",
            [],
        )

        valores = (
            matriz[0]
            if matriz
            else []
        )

        registro = {
            coluna:
                valores[indice]
                if indice < len(valores)
                else None

            for indice, coluna
            in enumerate(colunas)
        }

        registro_id = str(
            registro.get(
                "ID",
                "",
            )
        ).strip()

        if (
            registro_id.casefold()
            == target_id.casefold()
        ):
            return {
                "posicao": posicao,
                "graph_index":
                    linha.get(
                        "index",
                        posicao,
                    ),
                "valores": valores,
                "registro": registro,
            }

    return None


# ============================================================
# DATA E HORA DO EXCEL
# ============================================================

def excel_serial_para_datetime(
    data_serial,
    hora_serial=0,
):
    if data_serial in (
        None,
        "",
    ):
        raise ValueError(
            "Data do Excel nao informada."
        )

    if hora_serial in (
        None,
        "",
    ):
        hora_serial = 0

    base = datetime(
        1899,
        12,
        30,
    )

    return base + timedelta(
        days=(
            float(data_serial)
            + float(hora_serial)
        )
    )


def valor_sim(valor):
    return (
        str(valor or "")
        .strip()
        .casefold()
        in {
            "sim",
            "s",
            "yes",
            "true",
            "1",
        }
    )


# ============================================================
# MONTAGEM DO EVENTO
# ============================================================

def montar_evento_outlook(registro):
    inicio = excel_serial_para_datetime(
        registro.get("Data Início"),
        registro.get("Hora Início"),
    )

    data_fim = (
        registro.get("Data Fim")
        or registro.get("Data Início")
    )

    hora_fim = (
        registro.get("Hora Fim")
        or registro.get("Hora Início")
    )

    fim = excel_serial_para_datetime(
        data_fim,
        hora_fim,
    )

    if fim <= inicio:
        fim = inicio + timedelta(
            hours=1
        )

    dia_inteiro = valor_sim(
        registro.get("Dia Inteiro")
    )

    lembrete = valor_sim(
        registro.get("Lembrete")
    )

    observacoes = str(
        registro.get("Observações")
        or ""
    ).strip()

    referencia = str(
        registro.get("Referência")
        or ""
    ).strip()

    area = str(
        registro.get("Área")
        or ""
    ).strip()

    tipo = str(
        registro.get("Tipo")
        or ""
    ).strip()

    identificador = str(
        registro.get("ID")
        or ""
    ).strip()

    corpo = [
        (
            f"Second Brain ID: "
            f"{identificador}"
        ),
        f"Area: {area}",
        f"Tipo: {tipo}",
    ]

    if referencia:
        corpo.append(
            f"Referencia: {referencia}"
        )

    if observacoes:
        corpo.append("")
        corpo.append(observacoes)

    transaction_id = str(
        uuid.uuid5(
            uuid.NAMESPACE_URL,
            (
                f"second-brain:"
                f"{identificador}"
            ),
        )
    )

    evento = {
        "subject": str(
            registro.get("Nome")
            or identificador
        ),
        "body": {
            "contentType": "text",
            "content": "\n".join(
                corpo
            ),
        },
        "start": {
            "dateTime": inicio.strftime(
                "%Y-%m-%dT%H:%M:%S"
            ),
            "timeZone":
                OUTLOOK_TIME_ZONE,
        },
        "end": {
            "dateTime": fim.strftime(
                "%Y-%m-%dT%H:%M:%S"
            ),
            "timeZone":
                OUTLOOK_TIME_ZONE,
        },
        "isAllDay": dia_inteiro,
        "isReminderOn": lembrete,
        "transactionId":
            transaction_id,
    }

    if lembrete:
        evento[
            "reminderMinutesBeforeStart"
        ] = 30

    return evento


# ============================================================
# OUTLOOK
# ============================================================

def obter_evento_outlook(
    access_token,
    event_id,
):
    event_id_codificado = quote(
        str(event_id),
        safe="",
    )

    endpoint = (
        f"/me/events/"
        f"{event_id_codificado}"
    )

    try:
        return graph_request(
            method="GET",
            endpoint=endpoint,
            access_token=access_token,
            params={
                "$select": (
                    "id,subject,start,end,"
                    "isCancelled,webLink"
                )
            },
        )

    except GraphError as error:
        if error.status_code == 404:
            return None

        raise


def criar_evento_outlook(
    access_token,
    registro,
):
    evento = montar_evento_outlook(
        registro
    )

    print()
    print(
        "Criando evento no Outlook..."
    )

    criado = graph_request(
        method="POST",
        endpoint="/me/events",
        access_token=access_token,
        json_body=evento,
    )

    event_id = criado.get("id")

    if not event_id:
        encerrar_com_erro(
            "O Outlook criou o evento, "
            "mas nao retornou o ID."
        )

    print(
        "Evento criado com sucesso."
    )

    print(
        f"Outlook Event ID: {event_id}"
    )

    return criado


# ============================================================
# GRAVACAO DO ID NO EXCEL
# ============================================================

def gravar_outlook_id_na_linha(
    access_token,
    item_id,
    session_id,
    graph_index,
    colunas,
    valores_atuais,
    event_id,
):
    if (
        "Outlook Event ID"
        not in colunas
    ):
        encerrar_com_erro(
            "A coluna Outlook Event ID "
            "nao existe na tbAgenda."
        )

    valores_novos = list(
        valores_atuais
    )

    while (
        len(valores_novos)
        < len(colunas)
    ):
        valores_novos.append("")

    indice_event_id = colunas.index(
        "Outlook Event ID"
    )

    valores_novos[
        indice_event_id
    ] = event_id

    endpoint = (
        f"/me/drive/items/"
        f"{item_id}/workbook/"
        f"tables/{TABLE_NAME}/rows/"
        f"itemAt(index={graph_index})/"
        f"range"
    )

    graph_request(
        method="PATCH",
        endpoint=endpoint,
        access_token=access_token,
        json_body={
            "values": [
                valores_novos
            ]
        },
        extra_headers={
            "workbook-session-id":
                session_id
        },
    )

    print(
        "Outlook Event ID "
        "gravado no Excel."
    )


def confirmar_id_gravado(
    access_token,
    item_id,
    session_id,
    colunas,
):
    linhas = (
        obter_todas_linhas_tbagenda(
            access_token,
            item_id,
            session_id,
        )
    )

    encontrado = localizar_registro(
        colunas,
        linhas,
        TARGET_ID,
    )

    if not encontrado:
        encerrar_com_erro(
            f"{TARGET_ID} desapareceu "
            "da tbAgenda apos a gravacao."
        )

    event_id = str(
        encontrado["registro"].get(
            "Outlook Event ID"
        )
        or ""
    ).strip()

    if not event_id:
        encerrar_com_erro(
            "A gravacao terminou, "
            "mas o Outlook Event ID "
            "continua vazio."
        )

    print(
        "Leitura de confirmacao: "
        "Outlook Event ID esta "
        "preenchido no Excel."
    )

    return event_id


# ============================================================
# SINCRONIZACAO DO TESTE
# ============================================================

def sincronizar_teste(
    access_token,
    item_id,
    session_id,
):
    colunas, _ = (
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

    encontrado = localizar_registro(
        colunas,
        linhas,
        TARGET_ID,
    )

    if not encontrado:
        encerrar_com_erro(
            f"O registro {TARGET_ID} "
            "nao foi encontrado "
            "na tbAgenda."
        )

    registro = encontrado["registro"]

    print()
    print("=" * 70)
    print(
        "SINCRONIZACAO EXCEL -> OUTLOOK"
    )
    print("=" * 70)

    print(
        f"ID: {registro.get('ID')}"
    )

    print(
        f"Nome: {registro.get('Nome')}"
    )

    print(
        f"Status: "
        f"{registro.get('Status')}"
    )

    status = str(
        registro.get("Status")
        or ""
    ).strip().casefold()

    if status != "ativo":
        print(
            "Registro nao esta Ativo. "
            "Nenhum evento sera criado."
        )

        return "IGNORADO"

    event_id_existente = str(
        registro.get(
            "Outlook Event ID"
        )
        or ""
    ).strip()

    # --------------------------------------------------------
    # EVENTO JA POSSUI ID
    # --------------------------------------------------------

    if event_id_existente:
        print(
            "Outlook Event ID "
            "ja preenchido:"
        )

        print(
            event_id_existente
        )

        evento = obter_evento_outlook(
            access_token,
            event_id_existente,
        )

        if evento:
            print(
                "Evento confirmado "
                "no Outlook."
            )

            print(
                "Assunto no Outlook: "
                f"{evento.get('subject')}"
            )

            print(
                "Nenhum novo evento "
                "foi criado."
            )

            return "JA_EXISTIA"

        encerrar_com_erro(
            (
                "O Excel possui um "
                "Outlook Event ID, mas "
                "o evento nao foi encontrado."
            ),
            (
                "Para evitar duplicacao, "
                "o script nao criou outro "
                "evento automaticamente."
            ),
        )

    # --------------------------------------------------------
    # CRIAR EVENTO
    # --------------------------------------------------------

    criado = criar_evento_outlook(
        access_token,
        registro,
    )

    event_id = criado["id"]

    # --------------------------------------------------------
    # GRAVAR ID NO EXCEL
    # --------------------------------------------------------

    gravar_outlook_id_na_linha(
        access_token=access_token,
        item_id=item_id,
        session_id=session_id,
        graph_index=(
            encontrado["graph_index"]
        ),
        colunas=colunas,
        valores_atuais=(
            encontrado["valores"]
        ),
        event_id=event_id,
    )

    # --------------------------------------------------------
    # CONFIRMAR GRAVACAO
    # --------------------------------------------------------

    event_id_confirmado = (
        confirmar_id_gravado(
            access_token,
            item_id,
            session_id,
            colunas,
        )
    )

    # --------------------------------------------------------
    # CONFIRMAR EVENTO NO OUTLOOK
    # --------------------------------------------------------

    evento_confirmado = (
        obter_evento_outlook(
            access_token,
            event_id_confirmado,
        )
    )

    if not evento_confirmado:
        encerrar_com_erro(
            "O ID foi gravado, "
            "mas o evento nao foi "
            "confirmado no Outlook."
        )

    print(
        "Evento confirmado no Outlook "
        "depois da gravacao no Excel."
    )

    return "CRIADO"


# ============================================================
# EXECUCAO
# ============================================================

def main():
    access_token = None
    item_id = None
    session_id = None

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

        resultado = sincronizar_teste(
            access_token,
            item_id,
            session_id,
        )

        fechar_sessao_workbook(
            access_token,
            item_id,
            session_id,
        )

        session_id = None

        print()
        print("=" * 70)
        print("TESTE CONCLUIDO")
        print("=" * 70)

        print(
            "Autenticacao silenciosa: OK"
        )

        print(
            "Leitura da tbAgenda: OK"
        )

        if resultado == "CRIADO":
            print(
                "Criacao do evento "
                "no Outlook: OK"
            )

            print(
                "Gravacao do Outlook "
                "Event ID no Excel: OK"
            )

            print(
                "Agora execute novamente "
                "para testar a protecao "
                "contra duplicacao."
            )

        elif resultado == "JA_EXISTIA":
            print(
                "Confirmacao do evento "
                "existente: OK"
            )

            print(
                "Protecao contra "
                "duplicacao: OK"
            )

        else:
            print(
                "Registro ignorado "
                "por regra de status."
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

    finally:
        if (
            access_token
            and item_id
            and session_id
        ):
            try:
                fechar_sessao_workbook(
                    access_token,
                    item_id,
                    session_id,
                )

            except Exception:
                pass


if __name__ == "__main__":
    main()
