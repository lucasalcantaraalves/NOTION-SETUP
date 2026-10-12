import base64
import hashlib
import html
import json
import os
import re
import sys
import time
import unicodedata
import uuid
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path
from urllib.parse import quote

import msal
import requests
from cryptography.fernet import Fernet, InvalidToken

from categorias_outlook import (
    CATEGORIAS_AREA,
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
SCOPES = ["User.Read", "Files.ReadWrite", "Calendars.ReadWrite"]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
REQUEST_TIMEOUT = 60
MAX_GRAPH_ATTEMPTS = 5
RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}
EXCEL_PATH = "/Microsoft Copilot Chat Files/Copilot Notebook Uploads/Second Brain - Dados.xlsx"
TABLE_NAME = "tbAgenda"
OUTLOOK_TIME_ZONE = "E. South America Standard Time"
CACHE_PATH = Path(os.environ.get("MS_CACHE_FILE", ".auth/msal_cache.bin"))
POLITICA_VENCIMENTO = (180, 90, 30, 7)
STATUS_ATIVOS = {"ativo"}
STATUS_REMOVER = {"inativo", "cancelado", "cancelada"}

# Inferencia conservadora: frases no titulo valem mais do que frases no corpo.
REGRAS_INFERENCIA_AREA = {
    "malkuth": ["corte de cabelo", "cabeleireiro", "barbeiro", "compras", "mercado", "tarefas domesticas", "manutencao da casa", "transporte", "detran", "renovar cnh", "licenciamento", "passaporte"],
    "yesod": ["cinema", "filme", "anime", "show", "concerto", "videogame", "video game", "jogo", "festival", "exposicao", "museu", "entretenimento"],
    "hod": ["aula", "curso", "estudar", "estudo", "frances", "idioma", "prova", "aprendizado", "palestra", "workshop", "comunicacao"],
    "netzach": ["aniversario", "encontro com amigos", "encontro amoroso", "namoro", "casamento", "visitar familia", "reuniao familiar", "amigos", "confraternizacao", "date"],
    "tiphereth": ["desenvolvimento pessoal", "reflexao", "meditacao", "proposito", "espiritualidade", "metas pessoais", "autoconhecimento"],
    "geburah": ["academia", "treino", "musculacao", "fisioterapia", "consulta medica", "medico", "exame medico", "dentista", "cirurgia", "vacina", "saude", "nutricionista", "depilacao"],
    "chesed": ["pagamento", "conta bancaria", "banco", "investimento", "orcamento", "imposto", "fatura", "financas", "planejamento financeiro", "patrimonio", "crlv"],
    "binah": ["reuniao de trabalho", "entrevista de emprego", "entrevista profissional", "trabalho", "emprego", "carreira", "certificacao profissional", "documentacao profissional", "projeto profissional", "vencimento de documento"],
}

class GraphError(Exception):
    def __init__(self, status_code, method, url, response_text="", request_id="", client_request_id="", payload=None):
        self.status_code = status_code
        self.method = method
        self.url = url
        self.response_text = response_text
        self.request_id = request_id
        self.client_request_id = client_request_id
        self.payload = payload
        super().__init__(f"Microsoft Graph retornou HTTP {status_code} em {method} {url}\n{response_text}")


def formatar_resposta_graph(texto):
    if not texto:
        return "(sem corpo de resposta)"
    try:
        return json.dumps(json.loads(texto), indent=2, ensure_ascii=False)
    except (json.JSONDecodeError, TypeError):
        return str(texto)


def falhar(mensagem, detalhes=None):
    print("\n" + "=" * 70 + "\nERRO\n" + "=" * 70)
    print(mensagem)
    if detalhes:
        print("\n" + str(detalhes))
    sys.exit(1)


def graph_request(method, endpoint, token, params=None, json_body=None, extra_headers=None, aceitar_404=False):
    url = endpoint if endpoint.startswith("https://") else f"{GRAPH_BASE_URL}{endpoint}"
    last_response = None
    for tentativa in range(MAX_GRAPH_ATTEMPTS):
        client_request_id = str(uuid.uuid4())
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json", "client-request-id": client_request_id, "return-client-request-id": "true"}
        if json_body is not None:
            headers["Content-Type"] = "application/json"
        if extra_headers:
            headers.update(extra_headers)
        try:
            response = requests.request(method, url, headers=headers, params=params, json=json_body, timeout=REQUEST_TIMEOUT)
            last_response = response
        except requests.RequestException:
            if tentativa == MAX_GRAPH_ATTEMPTS - 1:
                raise
            espera = min(2 ** tentativa, 30)
            print(f"Falha de comunicacao com Graph; nova tentativa em {espera}s.")
            time.sleep(espera)
            continue
        if aceitar_404 and response.status_code == 404:
            return None
        if response.ok:
            if response.status_code == 204 or not response.content:
                return None
            try:
                return response.json()
            except requests.JSONDecodeError as error:
                raise RuntimeError(f"Graph retornou JSON invalido em {method} {url}: {response.text}") from error
        if response.status_code in RETRYABLE_STATUS_CODES and tentativa < MAX_GRAPH_ATTEMPTS - 1:
            retry_after = response.headers.get("Retry-After", "")
            try:
                espera = max(1, min(int(retry_after), 120)) if retry_after else min(2 ** tentativa, 30)
            except ValueError:
                espera = min(2 ** tentativa, 30)
            print(f"Graph HTTP {response.status_code}; nova tentativa em {espera}s.")
            time.sleep(espera)
            continue
        raise GraphError(response.status_code, method, url, response.text, response.headers.get("request-id", ""), response.headers.get("client-request-id", client_request_id), json_body)
    if last_response is not None:
        raise RuntimeError("Microsoft Graph permaneceu indisponivel apos varias tentativas.")
    raise RuntimeError("Microsoft Graph nao retornou resposta.")


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
        cache.deserialize(Fernet(chave_fernet(CACHE_KEY_TEXT)).decrypt(CACHE_PATH.read_bytes()).decode("utf-8"))
        print("Cache de autenticacao restaurado e descriptografado.")
        return cache
    except InvalidToken:
        falhar("Nao foi possivel descriptografar o cache MSAL.", "Confira o secret MS_CACHE_KEY.")
    except Exception as error:
        falhar("Falha ao carregar o cache MSAL.", repr(error))


def salvar_cache(cache):
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    dados = cache.serialize().encode("utf-8")
    CACHE_PATH.write_bytes(Fernet(chave_fernet(CACHE_KEY_TEXT)).encrypt(dados))
    print(f"Cache criptografado salvo em {CACHE_PATH}.")


def autenticar():
    cache = carregar_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    contas = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=contas[0]) if contas else None
    if result and "access_token" in result:
        print("Autenticacao silenciosa concluida com sucesso.")
        salvar_cache(cache)
        return result["access_token"]
    print("Autenticacao silenciosa indisponivel; iniciando Device Code Flow.")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        falhar("Nao foi possivel iniciar Device Code Flow.", json.dumps(flow, indent=2, ensure_ascii=False))
    print(flow.get("message"))
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        falhar("Nao foi possivel obter token Microsoft.", json.dumps(result, indent=2, ensure_ascii=False))
    salvar_cache(cache)
    return result["access_token"]


def identificar_usuario(token):
    usuario = graph_request("GET", "/me", token, params={"$select": "displayName,mail,userPrincipalName"})
    print(f"Conta conectada: {usuario.get('mail') or usuario.get('userPrincipalName')}")


def localizar_excel(token):
    caminho = quote(EXCEL_PATH, safe="/")
    print("Localizando o Excel pelo caminho fixo...")
    arquivo = graph_request("GET", f"/me/drive/root:{caminho}", token, params={"$select": "id,name,lastModifiedDateTime"})
    print(f"Excel encontrado: {arquivo.get('name')} | Item ID: {arquivo.get('id')}")
    return arquivo


def criar_sessao(token, item_id):
    resultado = graph_request("POST", f"/me/drive/items/{item_id}/workbook/createSession", token, json_body={"persistChanges": True})
    if not resultado.get("id"):
        raise RuntimeError("Graph nao retornou ID da sessao do workbook.")
    print("Sessao persistente do workbook criada.")
    return resultado["id"]


def fechar_sessao(token, item_id, session_id):
    graph_request("POST", f"/me/drive/items/{item_id}/workbook/closeSession", token, extra_headers={"workbook-session-id": session_id})
    print("Sessao do workbook encerrada.")


def obter_cabecalhos(token, item_id, session_id):
    resultado = graph_request("GET", f"/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/headerRowRange", token, extra_headers={"workbook-session-id": session_id})
    values = resultado.get("values", [])
    if not values or not values[0]:
        raise RuntimeError(f"A tabela {TABLE_NAME} nao retornou cabecalhos.")
    return values[0]


def obter_linhas(token, item_id, session_id):
    endpoint = f"/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/rows"
    params = {"$top": "200"}
    acumulado = []
    while endpoint:
        resultado = graph_request("GET", endpoint, token, params=params, extra_headers={"workbook-session-id": session_id})
        params = None
        acumulado.extend(resultado.get("value", []))
        endpoint = resultado.get("@odata.nextLink")
    return acumulado


def chave_coluna(valor):
    return normalizar_texto(valor).replace("_", " ").replace("-", " ").replace("/", " ").replace("(", " ").replace(")", " ").strip()


ALIASES_COLUNAS = {
    "ID": {"id", "identificador", "second brain id"},
    "Nome": {"nome", "titulo", "título", "assunto", "evento"},
    "Status": {"status", "situacao", "situação"},
    "Área": {"area", "área", "area do second brain", "area second brain"},
    "Tipo": {"tipo", "tipo de evento"},
    "Data Início": {
        "data inicio", "data de inicio", "data inicial", "inicio", "inicio data",
        "data hora inicio", "data hora de inicio", "data e hora inicio",
        "data e hora de inicio", "data de inicio do evento", "inicio do evento",
        "data do evento", "data evento", "data compromisso", "data do compromisso",
        "data de vencimento", "vencimento", "validade", "data validade",
        "data de validade", "data vencimento", "vencimento em"
    },
    "Data Fim": {"data fim", "data de fim", "data final", "termino", "data termino", "data de termino"},
    "Hora Início": {"hora inicio", "hora de inicio", "horario inicio", "horario de inicio"},
    "Hora Fim": {"hora fim", "hora de fim", "horario fim", "horario de fim"},
    "Dia Inteiro": {"dia inteiro", "evento de dia inteiro"},
    "Lembrete": {"lembrete", "tem lembrete", "ativar lembrete"},
    "Política de Aviso": {"politica de aviso", "política de aviso", "politica aviso", "avisos de vencimento"},
    "Recorrência": {"recorrencia", "recorrência", "repeticao", "repetição"},
    "Referência": {"referencia", "referência", "referencia externa"},
    "Observações": {"observacoes", "observações", "notas", "descricao", "descrição"},
    "Outlook Event ID": {"outlook event id", "outlook id", "id do evento outlook", "id evento outlook"},
}


def obter_registros(colunas, rows):
    # Mantém os nomes originais para gravar no Excel, mas cria também chaves
    # canônicas para ler cabeçalhos com pequenas variações de acento/formato.
    mapa_alias = {}
    for canonica, aliases in ALIASES_COLUNAS.items():
        for alias in aliases | {canonica}:
            mapa_alias[chave_coluna(alias)] = canonica
    registros = []
    for posicao, row in enumerate(rows):
        matriz = row.get("values", [])
        valores = matriz[0] if matriz else []
        registro = {coluna: valores[i] if i < len(valores) else "" for i, coluna in enumerate(colunas)}
        for i, coluna in enumerate(colunas):
            canonica = mapa_alias.get(chave_coluna(coluna))
            if canonica and canonica not in registro:
                registro[canonica] = valores[i] if i < len(valores) else ""
            elif canonica and (registro.get(canonica) in (None, "") or coluna == canonica):
                registro[canonica] = valores[i] if i < len(valores) else ""
        if any(str(valor or "").strip() for valor in valores):
            registros.append({"graph_index": row.get("index", posicao), "valores": valores, "registro": registro})
    return registros


def excel_datetime(data_serial, hora_serial=0):
    if data_serial in (None, ""):
        raise ValueError("Data de inicio nao informada; confira [CABECALHOS TBAGENDA] e os valores da linha.")
    if hora_serial in (None, ""):
        hora_serial = 0

    # O Excel via Graph normalmente retorna datas como números seriais, mas
    # também aceitamos datas textuais caso a coluna esteja formatada como texto.
    if isinstance(data_serial, str):
        texto = data_serial.strip()
        for formato in ("%d/%m/%Y", "%d/%m/%Y %H:%M", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%m/%d/%Y"):
            try:
                data = datetime.strptime(texto, formato)
                break
            except ValueError:
                data = None
        if data is not None:
            if isinstance(hora_serial, str) and hora_serial.strip():
                hora_txt = hora_serial.strip()
                for formato_hora in ("%H:%M", "%H:%M:%S"):
                    try:
                        hora = datetime.strptime(hora_txt, formato_hora).time()
                        return data.replace(hour=hora.hour, minute=hora.minute, second=hora.second)
                    except ValueError:
                        pass
            return data
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
    limpos = {str(k): str(v) for k, v in ids.items() if str(v).strip()}
    if not limpos:
        return ""
    if set(limpos) == {"principal"}:
        return limpos["principal"]
    return json.dumps(limpos, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def normalizar_texto(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    return " ".join(texto.casefold().strip().split())


def extrair_dia_semana(valor, inicio):
    texto = normalizar_texto(valor)
    mapa = {"segunda": "monday", "segunda-feira": "monday", "monday": "monday", "terca": "tuesday", "terca-feira": "tuesday", "tuesday": "tuesday", "quarta": "wednesday", "quarta-feira": "wednesday", "wednesday": "wednesday", "quinta": "thursday", "quinta-feira": "thursday", "thursday": "thursday", "sexta": "friday", "sexta-feira": "friday", "friday": "friday", "sabado": "saturday", "saturday": "saturday", "domingo": "sunday", "sunday": "sunday"}
    for nome, graph_day in mapa.items():
        if nome in texto:
            return graph_day
    return ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"][inicio.weekday()]


def montar_recorrencia(registro, inicio):
    original = str(registro.get("Recorrência") or "Nenhuma").strip()
    valor = normalizar_texto(original)
    if valor in {"", "nenhuma", "nao", "no"}:
        return None
    recurrence_range = {"type": "noEnd", "startDate": inicio.strftime("%Y-%m-%d")}
    if valor == "diaria" or valor == "daily" or valor.startswith("diaria "):
        pattern = {"type": "daily", "interval": 1}
    elif valor in {"semanal", "weekly"} or valor.startswith("semanal") or valor.startswith("weekly"):
        pattern = {"type": "weekly", "interval": 1, "daysOfWeek": [extrair_dia_semana(original, inicio)], "firstDayOfWeek": "monday"}
    elif valor in {"mensal", "monthly"} or valor.startswith("mensal "):
        pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": inicio.day}
    elif valor in {"anual", "annual", "yearly"} or valor.startswith("anual "):
        pattern = {"type": "absoluteYearly", "interval": 1, "dayOfMonth": inicio.day, "month": inicio.month}
    else:
        raise ValueError(f"Recorrencia nao reconhecida: {original}")
    return {"pattern": pattern, "range": recurrence_range}


def corpo_evento(registro, complemento=None):
    linhas = [f"Second Brain ID: {str(registro.get('ID') or '').strip()}", f"Área: {registro.get('Área') or ''}", f"Tipo: {registro.get('Tipo') or ''}"]
    categoria = categoria_da_area(registro.get("Área"))
    if categoria:
        linhas.append(f"Categoria: {categoria}")
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
        fim = excel_datetime(registro.get("Data Fim") or registro.get("Data Início"), 0) + timedelta(days=1)
    else:
        inicio = excel_datetime(registro.get("Data Início"), registro.get("Hora Início"))
        fim = excel_datetime(registro.get("Data Fim") or registro.get("Data Início"), registro.get("Hora Fim") or registro.get("Hora Início"))
        if fim <= inicio:
            fim = inicio + timedelta(hours=1)
    identificador = str(registro.get("ID") or "").strip()
    evento = {"subject": str(registro.get("Nome") or identificador), "body": {"contentType": "text", "content": corpo_evento(registro)}, "start": {"dateTime": inicio.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": OUTLOOK_TIME_ZONE}, "end": {"dateTime": fim.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": OUTLOOK_TIME_ZONE}, "isAllDay": dia_inteiro, "isReminderOn": valor_sim(registro.get("Lembrete")), "categories": categorias_da_area(registro.get("Área")), "transactionId": str(uuid.uuid5(uuid.NAMESPACE_URL, f"second-brain:{identificador}:principal"))}
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
    return {"subject": assunto, "body": {"contentType": "text", "content": corpo_evento(registro, complemento)}, "start": {"dateTime": data_evento.strftime("%Y-%m-%dT00:00:00"), "timeZone": OUTLOOK_TIME_ZONE}, "end": {"dateTime": fim.strftime("%Y-%m-%dT00:00:00"), "timeZone": OUTLOOK_TIME_ZONE}, "isAllDay": True, "isReminderOn": True, "reminderMinutesBeforeStart": 0, "categories": categorias_da_area(registro.get("Área")), "transactionId": str(uuid.uuid5(uuid.NAMESPACE_URL, f"second-brain:{identificador}:aviso:{chave}"))}


def endpoint_evento(event_id, calendar_id=None):
    event_id = quote(str(event_id), safe="")
    if calendar_id:
        return f"/me/calendars/{quote(str(calendar_id), safe='')}/events/{event_id}"
    return f"/me/events/{event_id}"


def buscar_evento(token, event_id, calendar_id=None):
    return graph_request("GET", endpoint_evento(event_id, calendar_id), token, params={"$select": "id,subject,body,start,end,isAllDay,recurrence,isReminderOn,categories"}, aceitar_404=True)


def criar_evento_payload(token, payload):
    return graph_request("POST", "/me/events", token, json_body=payload)


def atualizar_evento_payload(token, event_id, payload, calendar_id=None):
    atualizacao = dict(payload)
    atualizacao.pop("transactionId", None)
    return graph_request("PATCH", endpoint_evento(event_id, calendar_id), token, json_body=atualizacao)


def excluir_evento(token, event_id):
    return graph_request("DELETE", endpoint_evento(event_id), token, aceitar_404=True)


def extrair_evento_existente_de_erro(error):
    if error.status_code != 400:
        return None
    try:
        dados = json.loads(error.response_text).get("error") or {}
    except (json.JSONDecodeError, TypeError):
        return None
    return dados.get("@event.existingEventId") if dados.get("code") == "ErrorDuplicateTransactionId" else None


def criar_ou_recuperar_evento(token, payload):
    try:
        criado = criar_evento_payload(token, payload)
        if not criado.get("id"):
            raise RuntimeError("Outlook nao retornou ID do evento criado.")
        return criado["id"], True
    except GraphError as error:
        event_id = extrair_evento_existente_de_erro(error)
        if not event_id:
            raise
        existente = buscar_evento(token, event_id)
        if not existente:
            raise RuntimeError("Graph informou evento existente, mas ele nao foi encontrado.")
        atualizar_evento_payload(token, event_id, payload)
        print(f"TransactionId existente recuperado e atualizado: {event_id}")
        return event_id, False


def gravar_valor_coluna(token, item_id, session_id, graph_index, valores, colunas, coluna, valor):
    if coluna not in colunas:
        raise ValueError(f"A coluna {coluna} nao existe.")
    novos = list(valores)
    while len(novos) < len(colunas):
        novos.append("")
    novos[colunas.index(coluna)] = valor
    graph_request("PATCH", f"/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/rows/itemAt(index={graph_index})/range", token, json_body={"values": [novos]}, extra_headers={"workbook-session-id": session_id})
    return novos


def limpar_html(texto):
    texto = html.unescape(str(texto or ""))
    texto = re.sub(r"(?i)<\s*br\s*/?\s*>|</\s*(?:p|div|li|tr)\s*>", "\n", texto)
    return re.sub(r"<[^>]+>", " ", texto)


def identificar_area_na_descricao(descricao):
    texto = limpar_html(descricao)
    for linha in texto.splitlines():
        match = re.match(r"^\s*(?:área|area)\s*[:\-–]\s*(.*?)\s*$", linha, flags=re.IGNORECASE)
        if not match:
            continue
        valor = normalizar_texto(match.group(1))
        for chave in CATEGORIAS_AREA:
            if re.search(rf"(?<![a-z0-9]){re.escape(normalizar_texto(chave))}(?![a-z0-9])", valor):
                return chave
        return None
    return None


def inferir_area_por_contexto(evento):
    assunto = normalizar_texto(evento.get("subject") or "")
    corpo = normalizar_texto(limpar_html((evento.get("body") or {}).get("content") or ""))
    contexto = f"{assunto} {corpo}"
    pontuacoes = {}
    for area, frases in REGRAS_INFERENCIA_AREA.items():
        pontos = 0
        for frase in frases:
            frase = normalizar_texto(frase)
            if frase and frase in assunto:
                pontos += 3
            elif frase and frase in contexto:
                pontos += 1
        if pontos:
            pontuacoes[area] = pontos
    if not pontuacoes:
        return None
    ordenadas = sorted(pontuacoes.items(), key=lambda item: item[1], reverse=True)
    melhor_area, pontos = ordenadas[0]
    segundo = ordenadas[1][1] if len(ordenadas) > 1 else 0
    return melhor_area if pontos >= 3 and pontos > segundo else None


def resolver_area_evento(evento):
    area = identificar_area_na_descricao((evento.get("body") or {}).get("content", ""))
    if area:
        return area, "descricao"
    area = inferir_area_por_contexto(evento)
    if area:
        return area, "contexto"
    return None, None


def sincronizar_categoria_pela_descricao(token, evento, calendar_id=None):
    event_id = evento.get("id")
    assunto = evento.get("subject") or "(sem assunto)"
    if not event_id:
        print(f"[EVENTO IGNORADO] {assunto}: sem ID.")
        return "erro"
    area, origem = resolver_area_evento(evento)
    if not area:
        print(f"[AREA AMBIGUA] {assunto}. Categoria preservada.")
        return "ambiguo"
    categoria_correta = CATEGORIAS_AREA[area]["display_name"]
    nomes_area = {cfg["display_name"] for cfg in CATEGORIAS_AREA.values()}
    atuais = evento.get("categories") or []
    novas = [categoria for categoria in atuais if categoria not in nomes_area]
    novas.append(categoria_correta)
    if set(novas) == set(atuais):
        print(f"[CATEGORIA OK] {assunto} -> {area} (origem: {origem})")
        return "ok"
    graph_request("PATCH", endpoint_evento(event_id, calendar_id), token, json_body={"categories": novas})
    confirmado = buscar_evento(token, event_id, calendar_id)
    if not confirmado:
        raise RuntimeError(f"Evento nao encontrado apos corrigir categoria: {assunto}")
    if categoria_correta not in (confirmado.get("categories") or []):
        raise RuntimeError(f"Categoria nao confirmada no Outlook: {assunto}")
    print(f"[CATEGORIA CORRIGIDA E CONFIRMADA] {assunto} -> {categoria_correta} (origem: {origem})")
    return "corrigido"


def processar_remocao(token, item_id, session_id, item, colunas, resumo):
    ids = interpretar_ids(item["registro"].get("Outlook Event ID"))
    if not ids:
        resumo["ignorados"] += 1
        print("Registro sem Outlook Event ID. Nada para excluir.")
        return
    for chave, event_id in ids.items():
        existente = buscar_evento(token, event_id)
        if existente:
            excluir_evento(token, event_id)
            print(f"Evento excluido: {chave} | {existente.get('subject')}")
        else:
            print(f"Evento ja nao existe no Outlook: {chave}")
        resumo["excluidos"] += 1
    item["valores"] = gravar_valor_coluna(token, item_id, session_id, item["graph_index"], item["valores"], colunas, "Outlook Event ID", "")


def processar_evento_normal(token, item_id, session_id, item, colunas, resumo):
    registro = item["registro"]
    event_id = interpretar_ids(registro.get("Outlook Event ID")).get("principal")
    payload = montar_evento_normal(registro)
    if event_id:
        existente = buscar_evento(token, event_id)
        if not existente:
            raise RuntimeError("Outlook Event ID preenchido, mas evento nao encontrado.")
        atualizar_evento_payload(token, event_id, payload)
        resumo["atualizados"] += 1
        print("Evento existente atualizado no Outlook.")
    else:
        event_id, foi_criado = criar_ou_recuperar_evento(token, payload)
        item["valores"] = gravar_valor_coluna(token, item_id, session_id, item["graph_index"], item["valores"], colunas, "Outlook Event ID", event_id)
        resumo["criados" if foi_criado else "atualizados"] += 1
        print("Evento criado e Outlook Event ID gravado no Excel." if foi_criado else "Evento recuperado e Outlook Event ID gravado no Excel.")
    confirmado = buscar_evento(token, event_id)
    if not confirmado:
        raise RuntimeError("Evento nao confirmado depois da sincronizacao.")
    sincronizar_categoria_pela_descricao(token, confirmado)
    print(f"Confirmado no Outlook: {confirmado.get('subject')}")


def processar_politica_vencimento(token, item_id, session_id, item, colunas, resumo):
    registro = item["registro"]
    ids_atuais = interpretar_ids(registro.get("Outlook Event ID"))
    ids_finais = {}
    for chave in [str(dias) for dias in POLITICA_VENCIMENTO] + ["vencimento"]:
        payload = montar_evento_politica(registro, chave)
        event_id = ids_atuais.get(chave)
        existente = buscar_evento(token, event_id) if event_id else None
        if event_id and existente:
            atualizar_evento_payload(token, event_id, payload)
            resumo["avisos_atualizados"] += 1
            print(f"Aviso atualizado: {chave}")
        else:
            event_id, criado = criar_ou_recuperar_evento(token, payload)
            resumo["avisos_criados" if criado else "avisos_atualizados"] += 1
            print(f"Aviso {'criado' if criado else 'recuperado'}: {chave}")
        confirmado = buscar_evento(token, event_id)
        if not confirmado:
            raise RuntimeError(f"Aviso nao confirmado: {chave}")
        sincronizar_categoria_pela_descricao(token, confirmado)
        ids_finais[chave] = event_id
    item["valores"] = gravar_valor_coluna(token, item_id, session_id, item["graph_index"], item["valores"], colunas, "Outlook Event ID", serializar_ids(ids_finais))


def sincronizar(token, item_id, session_id):
    colunas = obter_cabecalhos(token, item_id, session_id)
    print("[CABECALHOS TBAGENDA] " + json.dumps(colunas, ensure_ascii=False))
    registros = obter_registros(colunas, obter_linhas(token, item_id, session_id))
    chaves_encontradas = {chave_coluna(coluna) for coluna in colunas}
    aliases_data = {chave_coluna(alias) for alias in ALIASES_COLUNAS["Data Início"] | {"Data Início"}}
    if not chaves_encontradas.intersection(aliases_data):
        print("[ATENCAO] Nenhuma coluna reconhecida como Data Início. As linhas que dependem dessa data serão registradas como erro; a varredura do Outlook continuará.")
    resumo = {"criados": 0, "atualizados": 0, "excluidos": 0, "avisos_criados": 0, "avisos_atualizados": 0, "ignorados": 0, "erros": 0}
    print(f"Linhas encontradas na tbAgenda: {len(registros)}")
    for item in registros:
        registro = item["registro"]
        registro_id = str(registro.get("ID") or "").strip()
        print(f"Processando {registro_id} | {registro.get('Nome')} | Status: {registro.get('Status')}")
        try:
            status = status_normalizado(registro)
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
        except GraphError as error:
            resumo["erros"] += 1
            print(f"[ERRO GRAPH] {registro_id}: HTTP {error.status_code} | {error.method} {error.url}\n{formatar_resposta_graph(error.response_text)}")
        except requests.RequestException as error:
            resumo["erros"] += 1
            print(f"[ERRO COMUNICACAO] {registro_id}: {repr(error)}")
        except Exception as error:
            resumo["erros"] += 1
            print(f"[ERRO INESPERADO] {registro_id}: {type(error).__name__}: {error}")
    print("[RESUMO SINCRONIZACAO EXCEL] " + json.dumps(resumo, ensure_ascii=False))
    # Não interrompe a rotina aqui: a varredura independente do Outlook deve
    # acontecer mesmo se uma ou mais linhas do Excel falharem.
    return resumo


def listar_calendarios(token):
    endpoint = "/me/calendars"
    params = {"$select": "id,name,canEdit", "$top": "100"}
    calendarios = []
    while endpoint:
        resultado = graph_request("GET", endpoint, token, params=params)
        params = None
        calendarios.extend(resultado.get("value", []))
        endpoint = resultado.get("@odata.nextLink")
    return calendarios


def janela_varredura_outlook():
    """Retorna a janela de hoje até três meses corridos à frente, no fuso de Brasília."""
    fuso = ZoneInfo("America/Sao_Paulo")
    agora = datetime.now(fuso)
    inicio = agora.replace(hour=0, minute=0, second=0, microsecond=0)
    ano = inicio.year
    mes = inicio.month + 3
    if mes > 12:
        ano += (mes - 1) // 12
        mes = ((mes - 1) % 12) + 1
    # Se o mês de destino não tiver o dia de hoje (ex.: 31), usa o último dia do mês.
    import calendar as _calendar
    dia = min(inicio.day, _calendar.monthrange(ano, mes)[1])
    fim = inicio.replace(year=ano, month=mes, day=dia)
    return inicio.isoformat(), fim.isoformat()


def listar_eventos_calendario(token, calendar_id, inicio, fim):
    # calendarView limita a busca temporal e inclui ocorrências de eventos recorrentes.
    endpoint = f"/me/calendars/{quote(str(calendar_id), safe='')}/calendarView"
    params = {
        "startDateTime": inicio,
        "endDateTime": fim,
        "$select": "id,subject,body,categories,type,seriesMasterId,start,end,isAllDay",
        "$top": "100",
    }
    eventos = []
    while endpoint:
        resultado = graph_request("GET", endpoint, token, params=params)
        params = None
        eventos.extend(resultado.get("value", []))
        endpoint = resultado.get("@odata.nextLink")
    return eventos


def classificar_eventos_de_todos_calendarios(token):
    """Classifica eventos na janela temporal, sem tentar alterar calendarios somente-leitura."""
    calendarios = listar_calendarios(token)
    inicio, fim = janela_varredura_outlook()
    resumo = {
        "calendarios": len(calendarios),
        "calendarios_somente_leitura": 0,
        "eventos_examinados": 0,
        "eventos_ignorados_somente_leitura": 0,
        "categorias_corrigidas": 0,
        "ja_corretos": 0,
        "ambiguos": 0,
        "erros": 0,
    }
    amostras_erros = []
    print(f"[VARREDURA OUTLOOK] {len(calendarios)} calendarios; janela {inicio} ate {fim}.")

    for calendario in calendarios:
        calendar_id = calendario.get("id")
        nome = calendario.get("name") or "(sem nome)"
        if not calendar_id:
            continue

        # O calendário padrão de aniversários é gerenciado pelo Outlook e é somente-leitura.
        nome_normalizado = normalizar_texto(nome)
        somente_leitura = calendario.get("canEdit") is False or "calendario de aniversarios" in nome_normalizado
        if somente_leitura:
            resumo["calendarios_somente_leitura"] += 1
            try:
                eventos = listar_eventos_calendario(token, calendar_id, inicio, fim)
                resumo["eventos_examinados"] += len(eventos)
                resumo["eventos_ignorados_somente_leitura"] += len(eventos)
                print(f"[CALENDARIO SOMENTE LEITURA] {nome}: {len(eventos)} eventos examinados; nenhuma alteração tentada.")
            except GraphError as error:
                resumo["erros"] += 1
                if len(amostras_erros) < 3:
                    amostras_erros.append(f"{nome}: HTTP {error.status_code}")
            continue

        try:
            eventos = listar_eventos_calendario(token, calendar_id, inicio, fim)
        except GraphError as error:
            resumo["erros"] += 1
            if len(amostras_erros) < 3:
                amostras_erros.append(f"Calendario {nome}: HTTP {error.status_code}")
            print(f"[ERRO CALENDARIO] {nome}: HTTP {error.status_code}")
            continue

        print(f"[CALENDARIO] {nome}: {len(eventos)} eventos.")
        for evento in eventos:
            resumo["eventos_examinados"] += 1
            assunto = evento.get("subject") or "(sem assunto)"
            try:
                resultado = sincronizar_categoria_pela_descricao(token, evento, calendar_id=calendar_id)
                if resultado == "corrigido":
                    resumo["categorias_corrigidas"] += 1
                elif resultado == "ambiguo":
                    resumo["ambiguos"] += 1
                elif resultado == "ok":
                    resumo["ja_corretos"] += 1
            except GraphError as error:
                # Se o Graph indicar calendário somente-leitura, para de tentar modificar
                # os demais eventos desse calendário e contabiliza-os como ignorados.
                texto_erro = (error.response_text or "").lower()
                if "read-only calendars can't be modified" in texto_erro or "read only calendar" in texto_erro:
                    resumo["calendarios_somente_leitura"] += 1
                    restantes = len(eventos) - eventos.index(evento) - 1
                    resumo["eventos_ignorados_somente_leitura"] += restantes + 1
                    resumo["eventos_examinados"] += restantes
                    print(f"[CALENDARIO SOMENTE LEITURA] {nome}: alterações interrompidas; {restantes + 1} eventos sem modificação.")
                    break
                resumo["erros"] += 1
                if len(amostras_erros) < 3:
                    amostras_erros.append(f"{assunto}: HTTP {error.status_code}")
                print(f"[ERRO EVENTO] {assunto}: HTTP {error.status_code}")
            except Exception as error:
                resumo["erros"] += 1
                if len(amostras_erros) < 3:
                    amostras_erros.append(f"{assunto}: {type(error).__name__}")
                print(f"[ERRO EVENTO] {assunto}: {type(error).__name__}")

    if amostras_erros:
        resumo["amostras_erros"] = amostras_erros
    print("[RESUMO VARREDURA OUTLOOK] " + json.dumps(resumo, ensure_ascii=False))
    if resumo["erros"]:
        raise RuntimeError("A varredura do Outlook encontrou erros reais; consulte as amostras no resumo.")


def main():
    token = None
    item_id = None
    session_id = None
    try:
        print("[VERSAO SCRIPT] SINCRONIZAR-AGENDA.py | varredura Outlook limitada a hoje + 3 meses | 2026-10-12-v4")
        token = autenticar()
        identificar_usuario(token)
        garantir_categorias(token)
        arquivo = localizar_excel(token)
        item_id = arquivo["id"]
        session_id = criar_sessao(token, item_id)
        resumo_excel = sincronizar(token, item_id, session_id)
        # Rotina independente: examina eventos criados diretamente no Outlook,
        # mesmo quando alguma linha do Excel falhou.
        classificar_eventos_de_todos_calendarios(token)
        fechar_sessao(token, item_id, session_id)
        session_id = None
        if resumo_excel.get("erros"):
            raise RuntimeError("A varredura do Outlook foi executada, mas uma ou mais linhas da tbAgenda falharam. Consulte os logs e os cabeçalhos impressos.")
        print("SINCRONIZACAO E VARREDURA CONCLUIDAS COM SUCESSO.")
    except GraphError as error:
        falhar("Falha no Microsoft Graph.", f"HTTP: {error.status_code}\nMetodo: {error.method}\nURL: {error.url}\nRequest ID: {error.request_id}\nResposta:\n{formatar_resposta_graph(error.response_text)}")
    except requests.RequestException as error:
        falhar("Falha de comunicacao com a Microsoft.", repr(error))
    except Exception as error:
        falhar("Erro inesperado.", f"Tipo: {type(error).__name__}\nDetalhes: {repr(error)}")
    finally:
        if token and item_id and session_id:
            try:
                fechar_sessao(token, item_id, session_id)
            except Exception:
                pass


if __name__ == "__main__":
    main()
