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


CLIENT_ID = os.environ.get("MS_CLIENT_ID")
CACHE_KEY_TEXT = os.environ.get("MS_CACHE_KEY")

if not CLIENT_ID:
    raise RuntimeError("Secret MS_CLIENT_ID nao encontrado.")
if not CACHE_KEY_TEXT:
    raise RuntimeError("Secret MS_CACHE_KEY nao encontrado.")

AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["User.Read", "Files.ReadWrite", "Calendars.ReadWrite"]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
REQUEST_TIMEOUT = 60
EXCEL_PATH = "/Microsoft Copilot Chat Files/Copilot Notebook Uploads/Second Brain - Dados.xlsx"
TABLE_NAME = "tbAgenda"
OUTLOOK_TIME_ZONE = "E. South America Standard Time"
CACHE_PATH = Path(".auth/msal_cache.bin")
POLITICA_VENCIMENTO = (180, 90, 30, 7)
STATUS_ATIVOS = {"ativo"}
STATUS_REMOVER = {"inativo", "cancelado", "cancelada"}


class GraphError(Exception):
    def __init__(self, status_code, method, url, response_text=""):
        self.status_code = status_code
        self.method = method
        self.url = url
        self.response_text = response_text
        super().__init__(f"Microsoft Graph retornou HTTP {status_code} em {method} {url}")


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


def graph_request(method, endpoint, token, params=None, json_body=None, extra_headers=None, aceitar_404=False):
    url = endpoint if endpoint.startswith("https://") else f"{GRAPH_BASE_URL}{endpoint}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
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
    if aceitar_404 and response.status_code == 404:
        return None
    if not response.ok:
        raise GraphError(response.status_code, method, url, response.text)
    if response.status_code == 204 or not response.content:
        return None
    return response.json()


def chave_fernet(valor):
    raw = valor.encode("utf-8")
    try:
        Fernet(raw)
        return raw
    except Exception:
        return base64.urlsafe_b64encode(hashlib.sha256(raw).digest())


def carregar_cache():
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    cache = msal.SerializableTokenCache()
    if not CACHE_PATH.exists():
        print("Cache de autenticacao ainda nao existe.")
        return cache
    try:
        fernet = Fernet(chave_fernet(CACHE_KEY_TEXT))
        serialized = fernet.decrypt(CACHE_PATH.read_bytes()).decode("utf-8")
        cache.deserialize(serialized)
        print("Cache de autenticacao restaurado e descriptografado.")
        return cache
    except InvalidToken:
        falhar("Nao foi possivel descriptografar o cache MSAL.", "Confira o secret MS_CACHE_KEY.")
    except Exception as error:
        falhar("Falha ao carregar o cache MSAL.", str(error))


def salvar_cache(cache):
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    encrypted = Fernet(chave_fernet(CACHE_KEY_TEXT)).encrypt(cache.serialize().encode("utf-8"))
    CACHE_PATH.write_bytes(encrypted)
    print(f"Cache criptografado salvo em {CACHE_PATH}.")


def autenticar():
    cache = carregar_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    result = None
    contas = app.get_accounts()
    if contas:
        print("Tentando autenticacao silenciosa...")
        result = app.acquire_token_silent(SCOPES, account=contas[0])
    if result and "access_token" in result:
        print("Autenticacao silenciosa concluida com sucesso.")
        salvar_cache(cache)
        return result["access_token"]
    print("Autenticacao silenciosa indisponivel.")
    print("Iniciando Device Code Flow como fallback...")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        falhar("Nao foi possivel iniciar o Device Code Flow.", json.dumps(flow, indent=2, ensure_ascii=False))
    print()
    print("=" * 70)
    print("AUTORIZACAO NECESSARIA")
    print("=" * 70)
    print(flow.get("message"))
    print("=" * 70)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        falhar("Nao foi possivel obter o token Microsoft.", json.dumps(result, indent=2, ensure_ascii=False))
    salvar_cache(cache)
    return result["access_token"]


def identificar_usuario(token):
    usuario = graph_request("GET", "/me", token, params={"$select": "displayName,mail,userPrincipalName"})
    conta = usuario.get("mail") or usuario.get("userPrincipalName")
    print(f"Conta conectada: {conta}")


def localizar_excel(token):
    caminho_codificado = quote(EXCEL_PATH, safe="/")
    print()
    print("Localizando o Excel pelo caminho fixo...")
    arquivo = graph_request(
        "GET",
        f"/me/drive/root:{caminho_codificado}",
        token,
        params={"$select": "id,name,lastModifiedDateTime"},
    )
    print(f"Excel encontrado: {arquivo.get('name')}")
    print(f"Item ID: {arquivo.get('id')}")
    print(f"Ultima modificacao: {arquivo.get('lastModifiedDateTime', 'nao informada')}")
    return arquivo


def criar_sessao(token, item_id):
    resultado = graph_request(
        "POST",
        f"/me/drive/items/{item_id}/workbook/createSession",
        token,
        json_body={"persistChanges": True},
    )
    session_id = resultado.get("id")
    if not session_id:
        falhar("O Graph nao retornou o ID da sessao.")
    print("Sessao persistente do workbook criada.")
    return session_id


def fechar_sessao(token, item_id, session_id):
    graph_request(
        "POST",
        f"/me/drive/items/{item_id}/workbook/closeSession",
        token,
        extra_headers={"workbook-session-id": session_id},
    )
    print("Sessao do workbook encerrada.")


def obter_cabecalhos(token, item_id, session_id):
    resultado = graph_request(
        "GET",
        f"/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/headerRowRange",
        token,
        extra_headers={"workbook-session-id": session_id},
    )
    values = resultado.get("values", [])
    if not values or not values[0]:
        falhar(f"A tabela {TABLE_NAME} nao retornou cabecalhos.")
    return values[0]


def obter_linhas(token, item_id, session_id):
    endpoint = f"/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/rows"
    acumulado = []
    params = {"$top": "200"}
    while endpoint:
        resultado = graph_request(
            "GET",
            endpoint,
            token,
            params=params,
            extra_headers={"workbook-session-id": session_id},
        )
        params = None
        acumulado.extend(resultado.get("value", []))
        endpoint = resultado.get("@odata.nextLink")
    return acumulado


def obter_registros(colunas, rows):
    registros = []
    for posicao, row in enumerate(rows):
        matriz = row.get("values", [])
        valores = matriz[0] if matriz else []
        registro = {
            coluna: valores[indice] if indice < len(valores) else ""
            for indice, coluna in enumerate(colunas)
        }
        if any(str(v or "").strip() for v in valores):
            registros.append(
                {
                    "posicao": posicao,
                    "graph_index": row.get("index", posicao),
                    "valores": valores,
                    "registro": registro,
                }
            )
    return registros


def excel_datetime(data_serial, hora_serial=0):
    if data_serial in (None, ""):
        raise ValueError("Data de inicio nao informada.")
    if hora_serial in (None, ""):
        hora_serial = 0
    return datetime(1899, 12, 30) + timedelta(days=float(data_serial) + float(hora_serial))


def valor_sim(valor):
    return str(valor or "").strip().casefold() in {"sim", "s", "yes", "true", "1"}


def status_normalizado(registro):
    return str(registro.get("Status") or "").strip().casefold()


def usa_politica_vencimento(registro):
    politica = str(registro.get("Política de Aviso") or "").strip().replace(" ", "")
    return valor_sim(registro.get("Lembrete")) and politica == "180/90/30/7"


def interpretar_ids(valor):
    texto = str(valor or "").strip()
    if not texto:
        return {}
    if texto.startswith("{"):
        try:
            dados = json.loads(texto)
            if isinstance(dados, dict):
                return {str(k): str(v) for k, v in dados.items() if str(v).strip()}
        except json.JSONDecodeError:
            pass
    return {"principal": texto}


def serializar_ids(ids):
    ids_limpos = {str(k): str(v) for k, v in ids.items() if str(v).strip()}
    if not ids_limpos:
        return ""
    if set(ids_limpos) == {"principal"}:
        return ids_limpos["principal"]
    return json.dumps(ids_limpos, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def montar_recorrencia(registro, inicio):
    valor = str(registro.get("Recorrência") or "Nenhuma").strip().casefold()
    if valor in {"", "nenhuma", "nao", "não"}:
        return None
    dias = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    recurrence_range = {"type": "noEnd", "startDate": inicio.strftime("%Y-%m-%d")}
    if valor in {"diaria", "diária", "daily"}:
        pattern = {"type": "daily", "interval": 1}
    elif valor in {"semanal", "weekly"}:
        pattern = {
            "type": "weekly",
            "interval": 1,
            "daysOfWeek": [dias[inicio.weekday()]],
            "firstDayOfWeek": "monday",
        }
    elif valor in {"mensal", "monthly"}:
        pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": inicio.day}
    elif valor in {"anual", "annual", "yearly"}:
        pattern = {
            "type": "absoluteYearly",
            "interval": 1,
            "dayOfMonth": inicio.day,
            "month": inicio.month,
        }
    else:
        raise ValueError(f"Recorrencia nao reconhecida: {registro.get('Recorrência')}")
    return {"pattern": pattern, "range": recurrence_range}


def corpo_evento(registro, complemento=None):
    linhas = [
        f"Second Brain ID: {str(registro.get('ID') or '').strip()}",
        f"Area: {registro.get('Área') or ''}",
        f"Tipo: {registro.get('Tipo') or ''}",
    ]
    if complemento:
        linhas.append(complemento)
    if registro.get("Referência"):
        linhas.append(f"Referencia: {registro.get('Referência')}")
    if registro.get("Observações"):
        linhas.extend(["", str(registro.get("Observações"))])
    return "\n".join(linhas)


def montar_evento_normal(registro):
    dia_inteiro = valor_sim(registro.get("Dia Inteiro"))
    if dia_inteiro:
        inicio = excel_datetime(registro.get("Data Início"), 0)
        fim_inclusivo = excel_datetime(registro.get("Data Fim") or registro.get("Data Início"), 0)
        fim = fim_inclusivo + timedelta(days=1)
    else:
        inicio = excel_datetime(registro.get("Data Início"), registro.get("Hora Início"))
        fim = excel_datetime(
            registro.get("Data Fim") or registro.get("Data Início"),
            registro.get("Hora Fim") or registro.get("Hora Início"),
        )
        if fim <= inicio:
            fim = inicio + timedelta(hours=1)
    identificador = str(registro.get("ID") or "").strip()
    evento = {
        "subject": str(registro.get("Nome") or identificador),
        "body": {"contentType": "text", "content": corpo_evento(registro)},
        "start": {"dateTime": inicio.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": OUTLOOK_TIME_ZONE},
        "end": {"dateTime": fim.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": OUTLOOK_TIME_ZONE},
        "isAllDay": dia_inteiro,
        "isReminderOn": valor_sim(registro.get("Lembrete")),
        "transactionId": str(uuid.uuid5(uuid.NAMESPACE_URL, f"second-brain:{identificador}:principal")),
    }
    recorrencia = montar_recorrencia(registro, inicio)
    if recorrencia:
        evento["recurrence"] = recorrencia
    if evento["isReminderOn"]:
        evento["reminderMinutesBeforeStart"] = 30
    return evento


def montar_evento_politica(registro, chave):
    vencimento = excel_datetime(registro.get("Data Início"), 0)
    nome = str(registro.get("Nome") or registro.get("ID") or "Vencimento")
    identificador = str(registro.get("ID") or "").strip()
    if chave == "vencimento":
        data_evento = vencimento
        assunto = f"⚫ VENCE HOJE · {nome}"
        complemento = "Politica de aviso: vencimento."
    else:
        dias = int(chave)
        data_evento = vencimento - timedelta(days=dias)
        rotulos = {180: "🔵 INFORMATIVO", 90: "🟡 ATENCAO", 30: "🟠 IMPORTANTE", 7: "🔴 URGENTE"}
        assunto = f"{rotulos[dias]} · {nome} vence em {dias} dias"
        complemento = f"Politica de aviso: {dias} dias antes do vencimento."
    fim = data_evento + timedelta(days=1)
    return {
        "subject": assunto,
        "body": {"contentType": "text", "content": corpo_evento(registro, complemento)},
        "start": {"dateTime": data_evento.strftime("%Y-%m-%dT00:00:00"), "timeZone": OUTLOOK_TIME_ZONE},
        "end": {"dateTime": fim.strftime("%Y-%m-%dT00:00:00"), "timeZone": OUTLOOK_TIME_ZONE},
        "isAllDay": True,
        "isReminderOn": True,
        "reminderMinutesBeforeStart": 0,
        "transactionId": str(uuid.uuid5(uuid.NAMESPACE_URL, f"second-brain:{identificador}:aviso:{chave}")),
    }


def buscar_evento(token, event_id):
    return graph_request(
        "GET",
        f"/me/events/{quote(str(event_id), safe='')}",
        token,
        params={"$select": "id,subject,start,end,isAllDay,recurrence,isReminderOn"},
        aceitar_404=True,
    )


def criar_evento_payload(token, payload):
    return graph_request("POST", "/me/events", token, json_body=payload)


def atualizar_evento_payload(token, event_id, payload):
    payload = dict(payload)
    payload.pop("transactionId", None)
    return graph_request("PATCH", f"/me/events/{quote(str(event_id), safe='')}", token, json_body=payload)


def excluir_evento(token, event_id):
    return graph_request(
        "DELETE",
        f"/me/events/{quote(str(event_id), safe='')}",
        token,
        aceitar_404=True,
    )


def gravar_valor_coluna(token, item_id, session_id, graph_index, valores, colunas, coluna, valor):
    if coluna not in colunas:
        raise ValueError(f"A coluna {coluna} nao existe.")
    novos_valores = list(valores)
    while len(novos_valores) < len(colunas):
        novos_valores.append("")
    novos_valores[colunas.index(coluna)] = valor
    graph_request(
        "PATCH",
        f"/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/rows/itemAt(index={graph_index})/range",
        token,
        json_body={"values": [novos_valores]},
        extra_headers={"workbook-session-id": session_id},
    )
    return novos_valores


def processar_remocao(token, item_id, session_id, item, colunas, resumo):
    registro = item["registro"]
    ids = interpretar_ids(registro.get("Outlook Event ID"))
    if not ids:
        print("Registro sem Outlook Event ID. Nada para excluir.")
        resumo["ignorados"] += 1
        return
    excluidos = 0
    for chave, event_id in ids.items():
        existente = buscar_evento(token, event_id)
        if existente:
            excluir_evento(token, event_id)
            print(f"Evento excluido: {chave} | {existente.get('subject')}")
        else:
            print(f"Evento ja nao existe no Outlook: {chave}")
        excluidos += 1
    item["valores"] = gravar_valor_coluna(
        token, item_id, session_id, item["graph_index"], item["valores"], colunas, "Outlook Event ID", ""
    )
    resumo["excluidos"] += excluidos
    print("Outlook Event ID limpo no Excel.")


def processar_evento_normal(token, item_id, session_id, item, colunas, resumo):
    registro = item["registro"]
    ids = interpretar_ids(registro.get("Outlook Event ID"))
    event_id = ids.get("principal")
    payload = montar_evento_normal(registro)
    if event_id:
        existente = buscar_evento(token, event_id)
        if not existente:
            raise RuntimeError("Outlook Event ID preenchido, mas o evento nao foi encontrado.")
        atualizar_evento_payload(token, event_id, payload)
        resumo["atualizados"] += 1
        print("Evento existente atualizado no Outlook.")
    else:
        criado = criar_evento_payload(token, payload)
        event_id = criado.get("id")
        if not event_id:
            raise RuntimeError("O Outlook nao retornou o ID do evento criado.")
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
        print("Evento criado e Outlook Event ID gravado no Excel.")
    confirmado = buscar_evento(token, event_id)
    if not confirmado:
        raise RuntimeError("Evento nao confirmado depois da sincronizacao.")
    print(f"Confirmado no Outlook: {confirmado.get('subject')}")


def processar_politica_vencimento(token, item_id, session_id, item, colunas, resumo):
    registro = item["registro"]
    ids_atuais = interpretar_ids(registro.get("Outlook Event ID"))
    ids_finais = {}
    chaves = [str(dias) for dias in POLITICA_VENCIMENTO] + ["vencimento"]
    for chave in chaves:
        payload = montar_evento_politica(registro, chave)
        event_id = ids_atuais.get(chave)
        if event_id:
            existente = buscar_evento(token, event_id)
            if existente:
                atualizar_evento_payload(token, event_id, payload)
                resumo["avisos_atualizados"] += 1
                print(f"Aviso atualizado: {chave}")
            else:
                criado = criar_evento_payload(token, payload)
                event_id = criado.get("id")
                resumo["avisos_criados"] += 1
                print(f"Aviso recriado porque o ID anterior nao existia: {chave}")
        else:
            criado = criar_evento_payload(token, payload)
            event_id = criado.get("id")
            resumo["avisos_criados"] += 1
            print(f"Aviso criado: {chave}")
        if not event_id:
            raise RuntimeError(f"O Outlook nao retornou ID para o aviso {chave}.")
        ids_finais[chave] = event_id
    valor_ids = serializar_ids(ids_finais)
    item["valores"] = gravar_valor_coluna(
        token,
        item_id,
        session_id,
        item["graph_index"],
        item["valores"],
        colunas,
        "Outlook Event ID",
        valor_ids,
    )
    print("IDs dos avisos 180/90/30/7 e vencimento gravados no Excel.")


def sincronizar(token, item_id, session_id):
    colunas = obter_cabecalhos(token, item_id, session_id)
    rows = obter_linhas(token, item_id, session_id)
    registros = obter_registros(colunas, rows)
    print()
    print("=" * 70)
    print("SINCRONIZACAO EXCEL -> OUTLOOK")
    print("=" * 70)
    print(f"Linhas encontradas na tbAgenda: {len(registros)}")
    resumo = {
        "criados": 0,
        "atualizados": 0,
        "excluidos": 0,
        "avisos_criados": 0,
        "avisos_atualizados": 0,
        "ignorados": 0,
        "erros": 0,
    }
    for item in registros:
        registro = item["registro"]
        registro_id = str(registro.get("ID") or "").strip()
        nome = str(registro.get("Nome") or "").strip()
        status = status_normalizado(registro)
        print()
        print("-" * 70)
        print(f"Processando {registro_id} | {nome} | Status: {registro.get('Status')}")
        try:
            if status in STATUS_REMOVER:
                processar_remocao(token, item_id, session_id, item, colunas, resumo)
            elif status in STATUS_ATIVOS:
                if usa_politica_vencimento(registro):
                    processar_politica_vencimento(token, item_id, session_id, item, colunas, resumo)
                else:
                    processar_evento_normal(token, item_id, session_id, item, colunas, resumo)
            else:
                resumo["ignorados"] += 1
                print("Status nao processado. Linha ignorada.")
        except Exception as error:
            resumo["erros"] += 1
            print(f"ERRO em {registro_id}: {error}")
    print()
    print("=" * 70)
    print("RESUMO DA SINCRONIZACAO")
    print("=" * 70)
    print(f"Linhas processadas: {len(registros)}")
    print(f"Eventos normais criados: {resumo['criados']}")
    print(f"Eventos normais atualizados: {resumo['atualizados']}")
    print(f"Eventos excluidos/ja ausentes: {resumo['excluidos']}")
    print(f"Avisos criados: {resumo['avisos_criados']}")
    print(f"Avisos atualizados: {resumo['avisos_atualizados']}")
    print(f"Linhas ignoradas: {resumo['ignorados']}")
    print(f"Erros: {resumo['erros']}")
    if resumo["erros"]:
        raise RuntimeError("Uma ou mais linhas falharam. Veja os logs acima.")


def main():
    token = None
    item_id = None
    session_id = None
    try:
        token = autenticar()
        identificar_usuario(token)
        arquivo = localizar_excel(token)
        item_id = arquivo["id"]
        session_id = criar_sessao(token, item_id)
        sincronizar(token, item_id, session_id)
        fechar_sessao(token, item_id, session_id)
        session_id = None
        print()
        print("SINCRONIZACAO CONCLUIDA COM SUCESSO.")
    except GraphError as error:
        falhar(
            "Falha no Microsoft Graph.",
            f"HTTP: {error.status_code}\nMetodo: {error.method}\nURL: {error.url}\nResposta: {error.response_text}",
        )
    except requests.RequestException as error:
        falhar("Falha de comunicacao com a Microsoft.", str(error))
    except Exception as error:
        falhar("Erro inesperado.", repr(error))
    finally:
        if token and item_id and session_id:
            try:
                fechar_sessao(token, item_id, session_id)
            except Exception:
                pass


if __name__ == "__main__":
    main()
