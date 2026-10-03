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

OUTLOOK_TIME_ZONE = (
    "E. South America Standard Time"
)

CACHE_PATH = Path(
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


def falhar(
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
    token,
    params=None,
    json_body=None,
    extra_headers=None,
):
    if endpoint.startswith("https://"):
        url = endpoint
    else:
        url = f"{GRAPH_BASE_URL}{endpoint}"

    headers = {
        "Authorization": f"Bearer {token}",
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

def chave_fernet(valor):
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
    CACHE_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    cache = msal.SerializableTokenCache()

    if not CACHE_PATH.exists():
        print(
            "Cache de autenticacao ainda "
            "nao existe."
        )

        return cache

    try:
        fernet = Fernet(
            chave_fernet(
                CACHE_KEY_TEXT
            )
        )

        encrypted = (
            CACHE_PATH.read_bytes()
        )

        serialized = (
            fernet.decrypt(encrypted)
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
        falhar(
            "Nao foi possivel descriptografar "
            "o cache MSAL.",
            (
                "Confira se o secret "
                "MS_CACHE_KEY continua "
                "com o mesmo valor."
            ),
        )

    except Exception as error:
        falhar(
            "Falha ao carregar "
            "o cache MSAL.",
            str(error),
        )


def salvar_cache(cache):
    CACHE_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    serialized = (
        cache.serialize()
        .encode("utf-8")
    )

    fernet = Fernet(
        chave_fernet(
            CACHE_KEY_TEXT
        )
    )

    encrypted = fernet.encrypt(
        serialized
    )

    CACHE_PATH.write_bytes(
        encrypted
    )

    print(
        f"Cache criptografado salvo em "
        f"{CACHE_PATH}."
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

    result = None
    contas = app.get_accounts()

    if contas:
        print(
            "Tentando autenticacao silenciosa..."
        )

        result = app.acquire_token_silent(
            scopes=SCOPES,
            account=contas[0],
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
        "Autenticacao silenciosa "
        "indisponivel."
    )

    print(
        "Iniciando Device Code Flow "
        "como fallback..."
    )

    flow = app.initiate_device_flow(
        scopes=SCOPES
    )

    if "user_code" not in flow:
        falhar(
            "Nao foi possivel iniciar "
            "o Device Code Flow.",
            json.dumps(
                flow,
                indent=2,
                ensure_ascii=False,
            ),
        )

    print()
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
        falhar(
            "Nao foi possivel obter "
            "o token Microsoft.",
            json.dumps(
                result,
                indent=2,
                ensure_ascii=False,
            ),
        )

    salvar_cache(cache)

    return result["access_token"]


# ============================================================
# CONTA MICROSOFT
# ============================================================

def identificar_usuario(token):
    usuario = graph_request(
        method="GET",
        endpoint="/me",
        token=token,
        params={
            "$select": (
                "displayName,"
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
# LOCALIZACAO DO EXCEL
# ============================================================

def localizar_excel(token):
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
        token=token,
        params={
            "$select": (
                "id,name,"
                "lastModifiedDateTime"
            )
        },
    )

    print(
        f"Excel encontrado: "
        f"{arquivo.get('name')}"
    )

    print(
        f"Item ID: "
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

def criar_sessao(
    token,
    item_id,
):
    resultado = graph_request(
        method="POST",
        endpoint=(
            f"/me/drive/items/"
            f"{item_id}/workbook/"
            f"createSession"
        ),
        token=token,
        json_body={
            "persistChanges": True
        },
    )

    session_id = resultado.get("id")

    if not session_id:
        falhar(
            "O Graph nao retornou "
            "o ID da sessao."
        )

    print(
        "Sessao persistente "
        "do workbook criada."
    )

    return session_id


def fechar_sessao(
    token,
    item_id,
    session_id,
):
    graph_request(
        method="POST",
        endpoint=(
            f"/me/drive/items/"
            f"{item_id}/workbook/"
            f"closeSession"
        ),
        token=token,
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

def obter_cabecalhos(
    token,
    item_id,
    session_id,
):
    resultado = graph_request(
        method="GET",
        endpoint=(
            f"/me/drive/items/"
            f"{item_id}/workbook/"
            f"tables/{TABLE_NAME}/"
            f"headerRowRange"
        ),
        token=token,
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
        falhar(
            f"A tabela {TABLE_NAME} "
            "nao retornou cabecalhos."
        )

    return values[0]


def obter_linhas(
    token,
    item_id,
    session_id,
):
    endpoint = (
        f"/me/drive/items/"
        f"{item_id}/workbook/"
        f"tables/{TABLE_NAME}/rows"
    )

    acumulado = []

    params = {
        "$top": "200"
    }

    while endpoint:
        resultado = graph_request(
            method="GET",
            endpoint=endpoint,
            token=token,
            params=params,
            extra_headers={
                "workbook-session-id":
                    session_id
            },
        )

        params = None

        acumulado.extend(
            resultado.get(
                "value",
                [],
            )
        )

        endpoint = resultado.get(
            "@odata.nextLink"
        )

    return acumulado


def obter_registros_ativos(
    colunas,
    rows,
):
    ativos = []

    for posicao, row in enumerate(
        rows
    ):
        matriz = row.get(
            "values",
            [],
        )

        valores = (
            matriz[0]
            if matriz
            else []
        )

        registro = {
            coluna: (
                valores[indice]
                if indice < len(valores)
                else ""
            )
            for indice, coluna
            in enumerate(colunas)
        }

        status = str(
            registro.get("Status")
            or ""
        ).strip().casefold()

        if status == "ativo":
            ativos.append(
                {
                    "posicao":
                        posicao,

                    "graph_index":
                        row.get(
                            "index",
                            posicao,
                        ),

                    "valores":
                        valores,

                    "registro":
                        registro,
                }
            )

    return ativos


# ============================================================
# DATA E HORA DO EXCEL
# ============================================================

def excel_datetime(
    data_serial,
    hora_serial=0,
):
    if data_serial in (
        None,
        "",
    ):
        raise ValueError(
            "Data de inicio "
            "nao informada."
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
# RECORRENCIA
# ============================================================

def montar_recorrencia(
    registro,
    inicio,
):
    valor = str(
        registro.get("Recorrência")
        or "Nenhuma"
    ).strip().casefold()

    if valor in {
        "",
        "nenhuma",
        "nao",
        "não",
    }:
        return None

    dias_semana = [
        "monday",
        "tuesday",
        "wednesday",
        "thursday",
        "friday",
        "saturday",
        "sunday",
    ]

    recurrence_range = {
        "type": "noEnd",
        "startDate":
            inicio.strftime(
                "%Y-%m-%d"
            ),
    }

    if valor in {
        "diaria",
        "diária",
        "daily",
    }:
        pattern = {
            "type": "daily",
            "interval": 1,
        }

    elif valor in {
        "semanal",
        "weekly",
    }:
        pattern = {
            "type": "weekly",
            "interval": 1,
            "daysOfWeek": [
                dias_semana[
                    inicio.weekday()
                ]
            ],
            "firstDayOfWeek":
                "monday",
        }

    elif valor in {
        "mensal",
        "monthly",
    }:
        pattern = {
            "type":
                "absoluteMonthly",
            "interval": 1,
            "dayOfMonth":
                inicio.day,
        }

    elif valor in {
        "anual",
        "annual",
        "yearly",
    }:
        pattern = {
            "type":
                "absoluteYearly",
            "interval": 1,
            "dayOfMonth":
                inicio.day,
            "month":
                inicio.month,
        }

    else:
        raise ValueError(
            "Recorrencia nao reconhecida: "
            f"{registro.get('Recorrência')}"
        )

    return {
        "pattern": pattern,
        "range": recurrence_range,
    }


# ============================================================
# MONTAGEM DO EVENTO
# ============================================================

def montar_evento(registro):
    dia_inteiro = valor_sim(
        registro.get("Dia Inteiro")
    )

    if dia_inteiro:
        inicio = excel_datetime(
            registro.get("Data Início"),
            0,
        )

        data_fim = (
            registro.get("Data Fim")
            or registro.get("Data Início")
        )

        fim_inclusivo = excel_datetime(
            data_fim,
            0,
        )

        # Para eventos de dia inteiro,
        # o fim do Outlook e exclusivo.
        fim = fim_inclusivo + timedelta(
            days=1
        )

    else:
        inicio = excel_datetime(
            registro.get("Data Início"),
            registro.get("Hora Início"),
        )

        fim = excel_datetime(
            (
                registro.get("Data Fim")
                or registro.get(
                    "Data Início"
                )
            ),
            (
                registro.get("Hora Fim")
                or registro.get(
                    "Hora Início"
                )
            ),
        )

        if fim <= inicio:
            fim = inicio + timedelta(
                hours=1
            )

    identificador = str(
        registro.get("ID")
        or ""
    ).strip()

    corpo = [
        (
            f"Second Brain ID: "
            f"{identificador}"
        ),
        (
            f"Area: "
            f"{registro.get('Área') or ''}"
        ),
        (
            f"Tipo: "
            f"{registro.get('Tipo') or ''}"
        ),
    ]

    if registro.get("Referência"):
        corpo.append(
            "Referencia: "
            f"{registro.get('Referência')}"
        )

    if registro.get("Observações"):
        corpo.append("")
        corpo.append(
            str(
                registro.get(
                    "Observações"
                )
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
            "dateTime":
                inicio.strftime(
                    "%Y-%m-%dT%H:%M:%S"
                ),
            "timeZone":
                OUTLOOK_TIME_ZONE,
        },

        "end": {
            "dateTime":
                fim.strftime(
                    "%Y-%m-%dT%H:%M:%S"
                ),
            "timeZone":
                OUTLOOK_TIME_ZONE,
        },

        "isAllDay":
            dia_inteiro,

        "isReminderOn":
            valor_sim(
                registro.get(
                    "Lembrete"
                )
            ),

        "transactionId": str(
            uuid.uuid5(
                uuid.NAMESPACE_URL,
                (
                    f"second-brain:"
                    f"{identificador}"
                ),
            )
        ),
    }

    recorrencia = montar_recorrencia(
        registro,
        inicio,
    )

    if recorrencia:
        evento["recurrence"] = (
            recorrencia
        )

    if evento["isReminderOn"]:
        evento[
            "reminderMinutesBeforeStart"
        ] = 30

    return evento


# ============================================================
# OUTLOOK
# ============================================================

def buscar_evento(
    token,
    event_id,
):
    event_id_codificado = quote(
        str(event_id),
        safe="",
    )

    try:
        return graph_request(
            method="GET",
            endpoint=(
                f"/me/events/"
                f"{event_id_codificado}"
            ),
            token=token,
            params={
                "$select": (
                    "id,subject,start,end,"
                    "isAllDay,recurrence,"
                    "isReminderOn"
                )
            },
        )

    except GraphError as error:
        if error.status_code == 404:
            return None

        raise


def criar_evento(
    token,
    registro,
):
    payload = montar_evento(
        registro
    )

    return graph_request(
        method="POST",
        endpoint="/me/events",
        token=token,
        json_body=payload,
    )


def atualizar_evento(
    token,
    event_id,
    registro,
):
    payload = montar_evento(
        registro
    )

    # transactionId e utilizado
    # somente na criacao.
    payload.pop(
        "transactionId",
        None,
    )

    event_id_codificado = quote(
        str(event_id),
        safe="",
    )

    return graph_request(
        method="PATCH",
        endpoint=(
            f"/me/events/"
            f"{event_id_codificado}"
        ),
        token=token,
        json_body=payload,
    )


# ============================================================
# GRAVACAO DO EVENT ID NO EXCEL
# ============================================================

def gravar_event_id(
    token,
    item_id,
    session_id,
    graph_index,
    valores,
    colunas,
    event_id,
):
    if (
        "Outlook Event ID"
        not in colunas
    ):
        raise ValueError(
            "A coluna Outlook Event ID "
            "nao existe."
        )

    novos_valores = list(
        valores
    )

    while (
        len(novos_valores)
        < len(colunas)
    ):
        novos_valores.append("")

    indice_event_id = colunas.index(
        "Outlook Event ID"
    )

    novos_valores[
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
        token=token,
        json_body={
            "values": [
                novos_valores
            ]
        },
        extra_headers={
            "workbook-session-id":
                session_id
        },
    )


# ============================================================
# SINCRONIZACAO
# ============================================================

def sincronizar(
    token,
    item_id,
    session_id,
):
    colunas = obter_cabecalhos(
        token,
        item_id,
        session_id,
    )

    rows = obter_linhas(
        token,
        item_id,
        session_id,
    )

    ativos = obter_registros_ativos(
        colunas,
        rows,
    )

    print()
    print("=" * 70)
    print(
        "SINCRONIZACAO EXCEL -> OUTLOOK"
    )
    print("=" * 70)

    print(
        f"Linhas ativas encontradas: "
        f"{len(ativos)}"
    )

    resumo = {
        "criados": 0,
        "atualizados": 0,
        "erros": 0,
    }

    for item in ativos:
        registro = item["registro"]

        registro_id = str(
            registro.get("ID")
            or ""
        ).strip()

        nome = str(
            registro.get("Nome")
            or ""
        ).strip()

        event_id = str(
            registro.get(
                "Outlook Event ID"
            )
            or ""
        ).strip()

        print()
        print("-" * 70)

        print(
            f"Processando "
            f"{registro_id} | {nome}"
        )

        try:
            # ------------------------------------------------
            # ATUALIZAR EVENTO EXISTENTE
            # ------------------------------------------------

            if event_id:
                existente = buscar_evento(
                    token,
                    event_id,
                )

                if not existente:
                    raise RuntimeError(
                        "Outlook Event ID "
                        "preenchido, mas o evento "
                        "nao foi encontrado."
                    )

                atualizar_evento(
                    token,
                    event_id,
                    registro,
                )

                resumo[
                    "atualizados"
                ] += 1

                print(
                    "Evento existente "
                    "atualizado no Outlook."
                )

            # ------------------------------------------------
            # CRIAR EVENTO NOVO
            # ------------------------------------------------

            else:
                criado = criar_evento(
                    token,
                    registro,
                )

                event_id = criado.get(
                    "id"
                )

                if not event_id:
                    raise RuntimeError(
                        "O Outlook nao retornou "
                        "o ID do evento criado."
                    )

                gravar_event_id(
                    token=token,
                    item_id=item_id,
                    session_id=session_id,
                    graph_index=(
                        item["graph_index"]
                    ),
                    valores=(
                        item["valores"]
                    ),
                    colunas=colunas,
                    event_id=event_id,
                )

                resumo[
                    "criados"
                ] += 1

                print(
                    "Evento criado e "
                    "Outlook Event ID "
                    "gravado no Excel."
                )

            # ------------------------------------------------
            # CONFIRMACAO
            # ------------------------------------------------

            confirmado = buscar_evento(
                token,
                event_id,
            )

            if not confirmado:
                raise RuntimeError(
                    "Evento nao confirmado "
                    "depois da sincronizacao."
                )

            print(
                "Confirmado no Outlook: "
                f"{confirmado.get('subject')}"
            )

            print(
                "Dia inteiro: "
                f"{confirmado.get('isAllDay')}"
            )

            recorrencia_confirmada = (
                confirmado.get(
                    "recurrence"
                )
            )

            print(
                "Recorrencia: "
                f"{json.dumps(
                    recorrencia_confirmada,
                    ensure_ascii=False,
                )}"
            )

        except Exception as error:
            resumo["erros"] += 1

            print(
                f"ERRO em {registro_id}: "
                f"{error}"
            )

    # ========================================================
    # RESUMO
    # ========================================================

    print()
    print("=" * 70)
    print("RESUMO DA SINCRONIZACAO")
    print("=" * 70)

    print(
        f"Ativos processados: "
        f"{len(ativos)}"
    )

    print(
        f"Eventos criados: "
        f"{resumo['criados']}"
    )

    print(
        f"Eventos atualizados: "
        f"{resumo['atualizados']}"
    )

    print(
        f"Erros: "
        f"{resumo['erros']}"
    )

    if resumo["erros"]:
        raise RuntimeError(
            "Uma ou mais linhas falharam. "
            "Veja os logs acima."
        )


# ============================================================
# EXECUCAO
# ============================================================

def main():
    token = None
    item_id = None
    session_id = None

    try:
        token = autenticar()

        identificar_usuario(
            token
        )

        arquivo = localizar_excel(
            token
        )

        item_id = arquivo["id"]

        session_id = criar_sessao(
            token,
            item_id,
        )

        sincronizar(
            token,
            item_id,
            session_id,
        )

        fechar_sessao(
            token,
            item_id,
            session_id,
        )

        session_id = None

        print()
        print(
            "SINCRONIZACAO CONCLUIDA "
            "COM SUCESSO."
        )

    except GraphError as error:
        falhar(
            "Falha no Microsoft Graph.",
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
        falhar(
            "Falha de comunicacao "
            "com a Microsoft.",
            str(error),
        )

    except Exception as error:
        falhar(
            "Erro inesperado.",
            repr(error),
        )

    finally:
        if (
            token
            and item_id
            and session_id
        ):
            try:
                fechar_sessao(
                    token,
                    item_id,
                    session_id,
                )

            except Exception:
                pass


if __name__ == "__main__":
    main()
