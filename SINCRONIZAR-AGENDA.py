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

from categorias_outlook import (
    categoria_da_area,
    categorias_da_area,
    garantir_categorias,
)


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
OUTLOOK_TIME_ZONE = "E. South America Standard Time"
CACHE_PATH = Path(".auth/msal_cache.bin")

POLITICA_VENCIMENTO = (180, 90, 30, 7)

STATUS_ATIVOS = {
    "ativo",
}

STATUS_REMOVER = {
    "inativo",
    "cancelado",
    "cancelada",
}


class GraphError(Exception):
    def __init__(
        self,
        status_code,
        method,
        url,
        response_text="",
        request_id="",
        client_request_id="",
        payload=None,
    ):
        self.status_code = status_code
        self.method = method
        self.url = url
        self.response_text = response_text
        self.request_id = request_id
        self.client_request_id = client_request_id
        self.payload = payload

        mensagem = (
            f"Microsoft Graph retornou HTTP {status_code} "
            f"em {method} {url}"
        )

        if response_text:
            mensagem += (
                "\n\n"
                "Resposta do Microsoft Graph:\n"
                f"{formatar_resposta_graph(response_text)}"
            )

        if request_id:
            mensagem += f"\nRequest ID: {request_id}"

        if client_request_id:
            mensagem += (
                f"\nClient Request ID: {client_request_id}"
            )

        super().__init__(mensagem)


def formatar_resposta_graph(response_text):
    if response_text in (None, ""):
        return "(sem corpo de resposta)"

    try:
        resposta_json = json.loads(response_text)

        return json.dumps(
            resposta_json,
            indent=2,
            ensure_ascii=False,
        )

    except (json.JSONDecodeError, TypeError):
        return str(response_text)


def formatar_payload_para_log(payload):
    if payload is None:
        return "(sem payload)"

    try:
        return json.dumps(
            payload,
            indent=2,
            ensure_ascii=False,
            default=str,
        )

    except (TypeError, ValueError):
        return repr(payload)


def imprimir_erro_graph(registro_id, error):
    print()
    print("!" * 70)
    print(f"ERRO MICROSOFT GRAPH em {registro_id}")
    print("!" * 70)
    print(f"HTTP: {error.status_code}")
    print(f"Metodo: {error.method}")
    print(f"URL: {error.url}")

    if error.request_id:
        print(f"Request ID: {error.request_id}")

    if error.client_request_id:
        print(
            "Client Request ID: "
            f"{error.client_request_id}"
        )

    print()
    print("Resposta do Microsoft Graph:")
    print(formatar_resposta_graph(error.response_text))

    if error.payload is not None:
        print()
        print("Payload enviado:")
        print(formatar_payload_para_log(error.payload))

    print("!" * 70)
    print()


def imprimir_erro_comunicacao(registro_id, error):
    print()
    print("!" * 70)
    print(f"ERRO DE COMUNICACAO em {registro_id}")
    print("!" * 70)
    print(f"Tipo: {type(error).__name__}")
    print(f"Detalhes: {repr(error)}")
    print("!" * 70)
    print()


def imprimir_erro_inesperado(registro_id, error):
    print()
    print("!" * 70)
    print(f"ERRO INESPERADO em {registro_id}")
    print("!" * 70)
    print(f"Tipo: {type(error).__name__}")
    print(f"Detalhes: {repr(error)}")
    print("!" * 70)
    print()


def falhar(mensagem, detalhes=None):
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
    token,
    params=None,
    json_body=None,
    extra_headers=None,
    aceitar_404=False,
):
    url = (
        endpoint
        if endpoint.startswith("https://")
        else f"{GRAPH_BASE_URL}{endpoint}"
    )

    client_request_id = str(uuid.uuid4())

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "client-request-id": client_request_id,
        "return-client-request-id": "true",
    }

    if json_body is not None:
        headers["Content-Type"] = "application/json"

    if extra_headers:
        headers.update(extra_headers)

    try:
        response = requests.request(
            method=method,
            url=url,
            headers=headers,
            params=params,
            json=json_body,
            timeout=REQUEST_TIMEOUT,
        )

    except requests.RequestException:
        raise

    if aceitar_404 and response.status_code == 404:
        return None

    if not response.ok:
        request_id = (
            response.headers.get("request-id")
            or response.headers.get("x-ms-request-id")
            or ""
        )

        returned_client_request_id = (
            response.headers.get("client-request-id")
            or client_request_id
        )

        raise GraphError(
            status_code=response.status_code,
            method=method,
            url=url,
            response_text=response.text,
            request_id=request_id,
            client_request_id=returned_client_request_id,
            payload=json_body,
        )

    if response.status_code == 204:
        return None

    if not response.content:
        return None

    try:
        return response.json()

    except requests.JSONDecodeError as error:
        raise RuntimeError(
            "O Microsoft Graph retornou uma resposta "
            "que nao e um JSON valido.\n"
            f"Metodo: {method}\n"
            f"URL: {url}\n"
            f"Resposta: {response.text}"
        ) from error


def chave_fernet(valor):
    raw = valor.encode("utf-8")

    try:
        Fernet(raw)
        return raw

    except Exception:
        return base64.urlsafe_b64encode(
            hashlib.sha256(raw).digest()
        )


def carregar_cache():
    CACHE_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    cache = msal.SerializableTokenCache()

    if not CACHE_PATH.exists():
        print("Cache de autenticacao ainda nao existe.")
        return cache

    try:
        encrypted = CACHE_PATH.read_bytes()

        serialized = (
            Fernet(chave_fernet(CACHE_KEY_TEXT))
            .decrypt(encrypted)
            .decode("utf-8")
        )

        cache.deserialize(serialized)

        print(
            "Cache de autenticacao restaurado "
            "e descriptografado."
        )

        return cache

    except InvalidToken:
        falhar(
            "Nao foi possivel descriptografar o cache MSAL.",
            "Confira o secret MS_CACHE_KEY.",
        )

    except Exception as error:
        falhar(
            "Falha ao carregar o cache MSAL.",
            str(error),
        )


def salvar_cache(cache):
    CACHE_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    serialized = cache.serialize().encode("utf-8")

    encrypted = (
        Fernet(chave_fernet(CACHE_KEY_TEXT))
        .encrypt(serialized)
    )

    CACHE_PATH.write_bytes(encrypted)

    print(
        f"Cache criptografado salvo em {CACHE_PATH}."
    )


def autenticar():
    cache = carregar_cache()

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    result = None
    contas = app.get_accounts()

    if contas:
        print("Tentando autenticacao silenciosa...")

        result = app.acquire_token_silent(
            SCOPES,
            account=contas[0],
        )

    if result and "access_token" in result:
        print(
            "Autenticacao silenciosa concluida "
            "com sucesso."
        )

        salvar_cache(cache)

        return result["access_token"]

    print("Autenticacao silenciosa indisponivel.")
    print("Iniciando Device Code Flow como fallback...")

    flow = app.initiate_device_flow(
        scopes=SCOPES,
    )

    if "user_code" not in flow:
        falhar(
            "Nao foi possivel iniciar o Device Code Flow.",
            json.dumps(
                flow,
                indent=2,
                ensure_ascii=False,
            ),
        )

    print(flow.get("message"))

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        falhar(
            "Nao foi possivel obter o token Microsoft.",
            json.dumps(
                result,
                indent=2,
                ensure_ascii=False,
            ),
        )

    salvar_cache(cache)

    return result["access_token"]


def identificar_usuario(token):
    usuario = graph_request(
        "GET",
        "/me",
        token,
        params={
            "$select": (
                "displayName,"
                "mail,"
                "userPrincipalName"
            )
        },
    )

    conta = (
        usuario.get("mail")
        or usuario.get("userPrincipalName")
    )

    print(f"Conta conectada: {conta}")


def localizar_excel(token):
    caminho = quote(
        EXCEL_PATH,
        safe="/",
    )

    print("Localizando o Excel pelo caminho fixo...")

    arquivo = graph_request(
        "GET",
        f"/me/drive/root:{caminho}",
        token,
        params={
            "$select": (
                "id,"
                "name,"
                "lastModifiedDateTime"
            )
        },
    )

    print(f"Excel encontrado: {arquivo.get('name')}")
    print(f"Item ID: {arquivo.get('id')}")

    return arquivo


def criar_sessao(token, item_id):
    resultado = graph_request(
        "POST",
        (
            f"/me/drive/items/{item_id}"
            "/workbook/createSession"
        ),
        token,
        json_body={
            "persistChanges": True,
        },
    )

    session_id = resultado.get("id")

    if not session_id:
        raise RuntimeError(
            "O Graph nao retornou o ID da sessao."
        )

    print("Sessao persistente do workbook criada.")

    return session_id


def fechar_sessao(
    token,
    item_id,
    session_id,
):
    graph_request(
        "POST",
        (
            f"/me/drive/items/{item_id}"
            "/workbook/closeSession"
        ),
        token,
        extra_headers={
            "workbook-session-id": session_id,
        },
    )

    print("Sessao do workbook encerrada.")


def obter_cabecalhos(
    token,
    item_id,
    session_id,
):
    resultado = graph_request(
        "GET",
        (
            f"/me/drive/items/{item_id}"
            f"/workbook/tables/{TABLE_NAME}"
            "/headerRowRange"
        ),
        token,
        extra_headers={
            "workbook-session-id": session_id,
        },
    )

    values = resultado.get("values", [])

    if not values or not values[0\]:
        raise RuntimeError(
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
        f"/me/drive/items/{item_id}"
        f"/workbook/tables/{TABLE_NAME}/rows"
    )

    acumulado = []

    params = {
        "$top": "200",
    }

    while endpoint:
        resultado = graph_request(
            "GET",
            endpoint,
            token,
            params=params,
            extra_headers={
                "workbook-session-id": session_id,
            },
        )

        params = None

        acumulado.extend(
            resultado.get("value", [])
        )

        endpoint = resultado.get("@odata.nextLink")

    return acumulado


def obter_registros(colunas, rows):
    registros = []

    for posicao, row in enumerate(rows):
        matriz = row.get("values", [])
        valores = matriz[0] if matriz else []

        registro = {
            coluna: (
                valores[indice]
                if indice < len(valores)
                else ""
            )
            for indice, coluna in enumerate(colunas)
        }

        possui_conteudo = any(
            str(valor or "").strip()
            for valor in valores
        )

        if possui_conteudo:
            registros.append(
                {
                    "graph_index": row.get(
                        "index",
                        posicao,
                    ),
                    "valores": valores,
                    "registro": registro,
                }
            )

    return registros


def excel_datetime(
    data_serial,
    hora_serial=0,
):
    if data_serial in (None, ""):
        raise ValueError(
            "Data de inicio nao informada."
        )

    if hora_serial in (None, ""):
        hora_serial = 0

    return datetime(1899, 12, 30) + timedelta(
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


def status_normalizado(registro):
    return (
        str(registro.get("Status") or "")
        .strip()
        .casefold()
    )


def usa_politica_vencimento(registro):
    politica = (
        str(
            registro.get(
                "Política de Aviso"
            )
            or ""
        )
        .strip()
        .replace(" ", "")
    )

    return (
        valor_sim(registro.get("Lembrete"))
        and politica == "180/90/30/7"
    )


def interpretar_ids(valor):
    texto = str(valor or "").strip()

    if not texto:
        return {}

    if texto.startswith("{"):
        try:
            dados = json.loads(texto)

            if isinstance(dados, dict):
                return {
                    str(chave): str(event_id)
                    for chave, event_id in dados.items()
                    if str(event_id).strip()
                }

        except json.JSONDecodeError:
            pass

    return {
        "principal": texto,
    }


def serializar_ids(ids):
    limpos = {
        str(chave): str(event_id)
        for chave, event_id in ids.items()
        if str(event_id).strip()
    }

    if not limpos:
        return ""

    if set(limpos) == {"principal"}:
        return limpos["principal"]

    return json.dumps(
        limpos,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def montar_recorrencia(registro, inicio):
    valor = (
        str(
            registro.get("Recorrência")
            or "Nenhuma"
        )
        .strip()
        .casefold()
    )

    if valor in {
        "",
        "nenhuma",
        "nao",
        "não",
    }:
        return None

    recurrence_range = {
        "type": "noEnd",
        "startDate": inicio.strftime(
            "%Y-%m-%d"
        ),
    }

    dias = [
        "monday",
        "tuesday",
        "wednesday",
        "thursday",
        "friday",
        "saturday",
        "sunday",
    ]

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
                dias[inicio.weekday()]
            ],
            "firstDayOfWeek": "monday",
        }

    elif valor in {
        "mensal",
        "monthly",
    }:
        pattern = {
            "type": "absoluteMonthly",
            "interval": 1,
            "dayOfMonth": inicio.day,
        }

    elif valor in {
        "anual",
        "annual",
        "yearly",
    }:
        pattern = {
            "type": "absoluteYearly",
            "interval": 1,
            "dayOfMonth": inicio.day,
            "month": inicio.month,
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


def corpo_evento(
    registro,
    complemento=None,
):
    linhas = [
        (
            "Second Brain ID: "
            f"{str(registro.get('ID') or '').strip()}"
        ),
        f"Area: {registro.get('Área') or ''}",
        f"Tipo: {registro.get('Tipo') or ''}",
    ]

    categoria = categoria_da_area(
        registro.get("Área")
    )

    if categoria:
        linhas.append(
            f"Categoria: {categoria}"
        )

    if complemento:
        linhas.append(complemento)

    if registro.get("Referência"):
        linhas.append(
            "Referencia: "
            f"{registro.get('Referência')}"
        )

    if registro.get("Observações"):
        linhas.extend(
            [
                "",
                str(
                    registro.get("Observações")
                ),
            ]
        )

    return "\n".join(linhas)


def montar_evento_normal(registro):
    dia_inteiro = valor_sim(
        registro.get("Dia Inteiro")
    )

    if dia_inteiro:
        inicio = excel_datetime(
            registro.get("Data Início"),
            0,
        )

        fim = (
            excel_datetime(
                registro.get("Data Fim")
                or registro.get("Data Início"),
                0,
            )
            + timedelta(days=1)
        )

    else:
        inicio = excel_datetime(
            registro.get("Data Início"),
            registro.get("Hora Início"),
        )

        fim = excel_datetime(
            registro.get("Data Fim")
            or registro.get("Data Início"),
            registro.get("Hora Fim")
            or registro.get("Hora Início"),
        )

        if fim <= inicio:
            fim = inicio + timedelta(hours=1)

    identificador = str(
        registro.get("ID") or ""
    ).strip()

    evento = {
        "subject": str(
            registro.get("Nome")
            or identificador
        ),
        "body": {
            "contentType": "text",
            "content": corpo_evento(registro),
        },
        "start": {
            "dateTime": inicio.strftime(
                "%Y-%m-%dT%H:%M:%S"
            ),
            "timeZone": OUTLOOK_TIME_ZONE,
        },
        "end": {
            "dateTime": fim.strftime(
                "%Y-%m-%dT%H:%M:%S"
            ),
            "timeZone": OUTLOOK_TIME_ZONE,
        },
        "isAllDay": dia_inteiro,
        "isReminderOn": valor_sim(
            registro.get("Lembrete")
        ),
        "categories": categorias_da_area(
            registro.get("Área")
        ),
        "transactionId": str(
            uuid.uuid5(
                uuid.NAMESPACE_URL,
                (
                    "second-brain:"
                    f"{identificador}:principal"
                ),
            )
        ),
    }

    recorrencia = montar_recorrencia(
        registro,
        inicio,
    )

    if recorrencia:
        evento["recurrence"] = recorrencia

    if evento["isReminderOn"\]:
        evento["reminderMinutesBeforeStart"] = 30

    return evento


def montar_evento_politica(
    registro,
    chave,
):
    vencimento = excel_datetime(
        registro.get("Data Início"),
        0,
    )

    nome = str(
        registro.get("Nome")
        or registro.get("ID")
        or "Vencimento"
    )

    identificador = str(
        registro.get("ID")
        or ""
    ).strip()

    if chave == "vencimento":
        data_evento = vencimento

        assunto = (
            f"⚫ VENCE HOJE · {nome}"
        )

        complemento = (
            "Politica de aviso: vencimento."
        )

    else:
        dias = int(chave)

        data_evento = (
            vencimento
            - timedelta(days=dias)
        )

        rotulos = {
            180: "🔵 INFORMATIVO",
            90: "🟡 ATENCAO",
            30: "🟠 IMPORTANTE",
            7: "🔴 URGENTE",
        }

        assunto = (
            f"{rotulos[dias]} · "
            f"{nome} vence em {dias} dias"
        )

        complemento = (
            "Politica de aviso: "
            f"{dias} dias antes do vencimento."
        )

    fim = data_evento + timedelta(days=1)

    return {
        "subject": assunto,
        "body": {
            "contentType": "text",
            "content": corpo_evento(
                registro,
                complemento,
            ),
        },
        "start": {
            "dateTime": data_evento.strftime(
                "%Y-%m-%dT00:00:00"
            ),
            "timeZone": OUTLOOK_TIME_ZONE,
        },
        "end": {
            "dateTime": fim.strftime(
                "%Y-%m-%dT00:00:00"
            ),
            "timeZone": OUTLOOK_TIME_ZONE,
        },
        "isAllDay": True,
        "isReminderOn": True,
        "reminderMinutesBeforeStart": 0,
        "categories": categorias_da_area(
            registro.get("Área")
        ),
        "transactionId": str(
            uuid.uuid5(
                uuid.NAMESPACE_URL,
                (
                    "second-brain:"
                    f"{identificador}:"
                    f"aviso:{chave}"
                ),
            )
        ),
    }


def buscar_evento(
    token,
    event_id,
):
    event_id_codificado = quote(
        str(event_id),
        safe="",
    )

    return graph_request(
        "GET",
        f"/me/events/{event_id_codificado}",
        token,
        params={
            "$select": (
                "id,"
                "subject,"
                "start,"
                "end,"
                "isAllDay,"
                "recurrence,"
                "isReminderOn,"
                "categories"
            )
        },
        aceitar_404=True,
    )


def criar_evento_payload(
    token,
    payload,
):
    return graph_request(
        "POST",
        "/me/events",
        token,
        json_body=payload,
    )


def atualizar_evento_payload(
    token,
    event_id,
    payload,
):
    atualizacao = dict(payload)

    atualizacao.pop(
        "transactionId",
        None,
    )

    event_id_codificado = quote(
        str(event_id),
        safe="",
    )

    return graph_request(
        "PATCH",
        f"/me/events/{event_id_codificado}",
        token,
        json_body=atualizacao,
    )


def excluir_evento(
    token,
    event_id,
):
    event_id_codificado = quote(
        str(event_id),
        safe="",
    )

    return graph_request(
        "DELETE",
        f"/me/events/{event_id_codificado}",
        token,
        aceitar_404=True,
    )


def gravar_valor_coluna(
    token,
    item_id,
    session_id,
    graph_index,
    valores,
    colunas,
    coluna,
    valor,
):
    if coluna not in colunas:
        raise ValueError(
            f"A coluna {coluna} nao existe."
        )

    novos = list(valores)

    while len(novos) < len(colunas):
        novos.append("")

    novos[colunas.index(coluna)] = valor

    graph_request(
        "PATCH",
        (
            f"/me/drive/items/{item_id}"
            f"/workbook/tables/{TABLE_NAME}"
            f"/rows/itemAt(index={graph_index})"
            "/range"
        ),
        token,
        json_body={
            "values": [novos],
        },
        extra_headers={
            "workbook-session-id": session_id,
        },
    )

    return novos


def processar_remocao(
    token,
    item_id,
    session_id,
    item,
    colunas,
    resumo,
):
    ids = interpretar_ids(
        item["registro"].get(
            "Outlook Event ID"
        )
    )

    if not ids:
        resumo["ignorados"] += 1

        print(
            "Registro sem Outlook Event ID. "
            "Nada para excluir."
        )

        return

    for chave, event_id in ids.items():
        existente = buscar_evento(
            token,
            event_id,
        )

        if existente:
            excluir_evento(
                token,
                event_id,
            )

            print(
                "Evento excluido: "
                f"{chave} | "
                f"{existente.get('subject')}"
            )

        else:
            print(
                "Evento ja nao existe "
                f"no Outlook: {chave}"
            )

        resumo["excluidos"] += 1

    item["valores"] = gravar_valor_coluna(
        token,
        item_id,
        session_id,
        item["graph_index"],
        item["valores"],
        colunas,
        "Outlook Event ID",
        "",
    )


def processar_evento_normal(
    token,
    item_id,
    session_id,
    item,
    colunas,
    resumo,
):
    registro = item["registro"]

    event_id = interpretar_ids(
        registro.get("Outlook Event ID")
    ).get("principal")

    payload = montar_evento_normal(
        registro
    )

    if event_id:
        existente = buscar_evento(
            token,
            event_id,
        )

        if not existente:
            raise RuntimeError(
                "Outlook Event ID preenchido, "
                "mas o evento nao foi encontrado."
            )

        atualizar_evento_payload(
            token,
            event_id,
            payload,
        )

        resumo["atualizados"] += 1

        print(
            "Evento existente atualizado "
            "no Outlook."
        )

    else:
        criado = criar_evento_payload(
            token,
            payload,
        )

        event_id = criado.get("id")

        if not event_id:
            raise RuntimeError(
                "O Outlook nao retornou "
                "o ID do evento criado."
            )

        item["valores"] = gravar_valor_coluna(
            token,
            item_id,
            session_id,
            item["graph_index"],
            item["valores"],
            colunas,
            "Outlook Event ID",
            event_id,
        )

        resumo["criados"] += 1

        print(
            "Evento criado e Outlook Event ID "
            "gravado no Excel."
        )

    confirmado = buscar_evento(
        token,
        event_id,
    )

    if not confirmado:
        raise RuntimeError(
            "Evento nao confirmado "
            "depois da sincronizacao."
        )

    categoria = categoria_da_area(
        registro.get("Área")
    )

    categorias_confirmadas = (
        confirmado.get("categories")
        or []
    )

    if (
        categoria
        and categoria
        not in
