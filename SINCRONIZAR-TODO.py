#!/usr/bin/env python3
"""
SINCRONIZAR-TODO.py

Sincroniza com o Microsoft To Do:
  1. tbTodo, mantendo o comportamento existente.
  2. Somente os registros da tbListas cuja coluna Lista seja "Wishlist".

O Inventario e quaisquer outras listas da tbListas sao ignorados.

A Wishlist nao precisa de uma coluna de ID externo. O vinculo e localizado pelo
marcador "Second Brain ID: <ID>" gravado no corpo da tarefa no Microsoft To Do.

Variaveis obrigatorias:
  MS_CLIENT_ID
  MS_CACHE_KEY

Variaveis opcionais:
  MS_CACHE_FILE          Padrao: .msal_cache.enc
  MS_TENANT_ID           Padrao: consumers
  SB_EXCEL_PATH          Caminho do Excel no OneDrive
  SB_TODO_TABLE          Padrao: tbTodo
  SB_LISTS_TABLE         Padrao: tbListas
  SB_WISHLIST_NAME       Padrao: Wishlist
  SB_FORCE_LOGIN         Use 1 para forcar Device Code Flow
  SB_TIME_ZONE           Padrao: America/New_York
  SB_RECREATE_MISSING    Use 1 para recriar tarefa da tbTodo cujo vinculo sumiu

Dependencias:
  pip install msal requests cryptography
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import sys
import time
import unicodedata
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

import msal
import requests
from cryptography.fernet import Fernet, InvalidToken

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
CLIENT_ID = os.getenv("MS_CLIENT_ID", "").strip()
CACHE_KEY = os.getenv("MS_CACHE_KEY", "").strip()
TENANT_ID = os.getenv("MS_TENANT_ID", "consumers").strip() or "consumers"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.ReadWrite", "Tasks.ReadWrite"]

CACHE_FILE = os.getenv("MS_CACHE_FILE", ".msal_cache.enc").strip() or ".msal_cache.enc"
FORCE_LOGIN = os.getenv("SB_FORCE_LOGIN", "").strip().lower() in {"1", "true", "sim", "yes"}
RECREATE_MISSING = os.getenv("SB_RECREATE_MISSING", "").strip().lower() in {"1", "true", "sim", "yes"}
TIME_ZONE = os.getenv("SB_TIME_ZONE", "America/New_York").strip() or "America/New_York"
EXCEL_PATH = os.getenv(
    "SB_EXCEL_PATH",
    "/Microsoft Copilot Chat Files/Copilot Notebook Uploads/Second Brain - Dados.xlsx",
).strip()
TODO_TABLE = os.getenv("SB_TODO_TABLE", "tbTodo").strip() or "tbTodo"
LISTS_TABLE = os.getenv("SB_LISTS_TABLE", "tbListas").strip() or "tbListas"
WISHLIST_NAME = os.getenv("SB_WISHLIST_NAME", "Wishlist").strip() or "Wishlist"

REQUEST_TIMEOUT = 60
MAX_RETRIES = 5
TODO_HEADERS = [
    "ID", "Tarefa", "Lista", "Área", "Status", "Data", "Prioridade",
    "Observações", "Última atualização", "Destino Externo ID",
]
LIST_HEADERS = [
    "ID", "Lista", "Item", "Área", "Status", "Data", "Valor",
    "Detalhes", "Referência", "Última atualização",
]


@dataclass(frozen=True)
class ExternalRef:
    list_id: str
    task_id: str

    def serialize(self) -> str:
        return json.dumps(
            {"listId": self.list_id, "taskId": self.task_id},
            ensure_ascii=False,
            separators=(",", ":"),
        )


@dataclass
class CacheState:
    cache: msal.SerializableTokenCache
    force_save: bool = False


def log(message: str) -> None:
    print(message, flush=True)


def normalize_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize_status(value: Any) -> str:
    text = unicodedata.normalize("NFKD", normalize_text(value))
    text = text.encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"\s+", " ", text).strip()


def normalize_priority(value: Any) -> str:
    text = normalize_status(value)
    if text in {"alta", "high", "urgente"}:
        return "high"
    if text in {"baixa", "low"}:
        return "low"
    return "normal"


def normalize_key(value: str) -> bytes:
    raw = value.encode("utf-8")
    try:
        decoded = base64.urlsafe_b64decode(raw)
        if len(decoded) == 32:
            return raw
    except Exception:
        pass
    return base64.urlsafe_b64encode(hashlib.sha256(raw).digest())


def valid_cache_json(serialized: str) -> bool:
    try:
        payload = json.loads(serialized)
    except (TypeError, ValueError, json.JSONDecodeError):
        return False
    if not isinstance(payload, dict):
        return False
    sections = {"AccessToken", "RefreshToken", "IdToken", "Account", "AppMetadata", "TokenType"}
    return not payload or bool(sections.intersection(payload))


def validate_environment() -> None:
    missing = [name for name, value in (("MS_CLIENT_ID", CLIENT_ID), ("MS_CACHE_KEY", CACHE_KEY)) if not value]
    if missing:
        raise RuntimeError("Defina as variaveis obrigatorias: " + ", ".join(missing))


def load_cache() -> CacheState:
    cache = msal.SerializableTokenCache()
    path = Path(CACHE_FILE)
    if not path.exists():
        return CacheState(cache)
    data = path.read_bytes()
    if not data:
        return CacheState(cache, True)
    try:
        serialized = Fernet(normalize_key(CACHE_KEY)).decrypt(data).decode("utf-8")
        force_save = False
    except (InvalidToken, UnicodeDecodeError) as encrypted_error:
        try:
            serialized = data.decode("utf-8")
        except UnicodeDecodeError as exc:
            raise RuntimeError("Nao foi possivel abrir o cache MSAL.") from exc
        if not valid_cache_json(serialized):
            raise RuntimeError("MS_CACHE_KEY incorreta ou cache MSAL corrompido.") from encrypted_error
        force_save = True
        log("AVISO: cache MSAL em texto puro detectado; sera criptografado.")
    if serialized.strip():
        cache.deserialize(serialized)
    return CacheState(cache, force_save)


def save_cache(state: CacheState, force: bool = False) -> None:
    if not force and not state.force_save and not state.cache.has_state_changed:
        return
    path = Path(CACHE_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    encrypted = Fernet(normalize_key(CACHE_KEY)).encrypt(state.cache.serialize().encode("utf-8"))
    temp = path.with_name(path.name + ".tmp")
    temp.write_bytes(encrypted)
    try:
        os.chmod(temp, 0o600)
    except OSError:
        pass
    temp.replace(path)
    state.force_save = False
    log(f"Cache MSAL criptografado salvo em: {path}")


def acquire_token(state: CacheState) -> str:
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=state.cache)
    result: dict[str, Any] | None = None
    if not FORCE_LOGIN:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Falha ao iniciar Device Code Flow: {flow}")
        log(flow.get("message", "Abra a pagina indicada e informe o codigo."))
        result = app.acquire_token_by_device_flow(flow)
    if not result or "access_token" not in result:
        detail = (result or {}).get("error_description") or (result or {}).get("error") or "erro desconhecido"
        raise RuntimeError(f"Falha na autenticacao Microsoft: {detail}")
    save_cache(state)
    return result["access_token"]


class GraphClient:
    def __init__(self, token: str) -> None:
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        })

    def request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        if url.startswith("/"):
            url = GRAPH_ROOT + url
        elif not url.startswith("http"):
            url = GRAPH_ROOT + "/" + url
        response = None
        for attempt in range(MAX_RETRIES):
            try:
                response = self.session.request(method, url, timeout=REQUEST_TIMEOUT, **kwargs)
            except requests.RequestException as exc:
                if attempt == MAX_RETRIES - 1:
                    raise RuntimeError(f"Falha de rede: {exc}") from exc
                wait = min(2**attempt, 30)
                time.sleep(wait)
                continue
            if response.status_code not in {429, 500, 502, 503, 504}:
                return response
            try:
                wait = max(1, int(response.headers.get("Retry-After", "")))
            except ValueError:
                wait = min(2**attempt, 30)
            log(f"Graph retornou {response.status_code}; nova tentativa em {wait}s...")
            time.sleep(wait)
        if response is None:
            raise RuntimeError("Microsoft Graph nao retornou resposta.")
        return response

    def json(self, method: str, url: str, expected: tuple[int, ...] = (200,), **kwargs: Any) -> dict[str, Any]:
        response = self.request(method, url, **kwargs)
        if response.status_code not in expected:
            try:
                detail = response.json()
            except ValueError:
                detail = response.text
            raise RuntimeError(f"Microsoft Graph retornou {response.status_code}: {detail}")
        if response.status_code == 204 or not response.content:
            return {}
        return response.json()

    def paged_values(self, url: str) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        while url:
            payload = self.json("GET", url)
            result.extend(payload.get("value", []))
            url = payload.get("@odata.nextLink")
        return result


def encode_onedrive_path(path: str) -> str:
    return "/".join(quote(part, safe="") for part in path.strip("/").split("/") if part)


def get_workbook_item(graph: GraphClient) -> dict[str, Any]:
    return graph.json("GET", f"/me/drive/root:/{encode_onedrive_path(EXCEL_PATH)}")


def create_workbook_session(graph: GraphClient, item_id: str) -> str:
    data = graph.json("POST", f"/me/drive/items/{item_id}/workbook/createSession", expected=(200, 201), json={"persistChanges": True})
    session_id = normalize_text(data.get("id"))
    if not session_id:
        raise RuntimeError("Microsoft Graph nao retornou o ID da sessao do workbook.")
    return session_id


def close_workbook_session(graph: GraphClient, item_id: str, session_id: str) -> None:
    graph.json("POST", f"/me/drive/items/{item_id}/workbook/closeSession", expected=(200, 204), headers={"workbook-session-id": session_id})


def load_table(graph: GraphClient, item_id: str, session_id: str, table_name: str, expected_headers: list[str]) -> tuple[list[str], list[dict[str, Any]]]:
    headers = {"workbook-session-id": session_id}
    table = quote(table_name, safe="")
    base = f"/me/drive/items/{item_id}/workbook/tables/{table}"
    header_payload = graph.json("GET", base + "/headerRowRange", headers=headers)
    rows = header_payload.get("values", [])
    if not rows:
        raise RuntimeError(f"A tabela {table_name} nao possui cabecalho.")
    table_headers = [normalize_text(value) for value in rows[0]]
    missing = [name for name in expected_headers if name not in table_headers]
    if missing:
        raise RuntimeError(f"Colunas ausentes em {table_name}: {', '.join(missing)}")
    data = graph.json("GET", base + "/rows", headers=headers).get("value", [])
    if not isinstance(data, list):
        raise RuntimeError(f"Resposta invalida ao ler {table_name}.")
    return table_headers, data


def update_excel_row(graph: GraphClient, item_id: str, session_id: str, table_name: str, row_position: int, values: list[Any]) -> None:
    table = quote(table_name, safe="")
    graph.json(
        "PATCH",
        f"/me/drive/items/{item_id}/workbook/tables/{table}/rows/itemAt(index={row_position})/range",
        expected=(200,),
        headers={"workbook-session-id": session_id},
        json={"values": [values]},
    )


def excel_value_to_date(value: Any) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        raise ValueError(f"Data invalida: {value}")
    if isinstance(value, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    text = normalize_text(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise ValueError(f"Data invalida: {text}") from exc


def graph_date_time(day: date) -> dict[str, str]:
    return {"dateTime": f"{day.isoformat()}T09:00:00.0000000", "timeZone": TIME_ZONE}


def parse_external_ref(value: Any) -> ExternalRef | str | None:
    text = normalize_text(value)
    if not text:
        return None
    if text.startswith("{"):
        payload = json.loads(text)
        list_id = normalize_text(payload.get("listId"))
        task_id = normalize_text(payload.get("taskId"))
        if not list_id or not task_id:
            raise ValueError("Destino Externo ID deve conter listId e taskId.")
        return ExternalRef(list_id, task_id)
    return text


def get_task_lists(graph: GraphClient) -> tuple[dict[str, str], dict[str, str]]:
    by_name: dict[str, str] = {}
    by_id: dict[str, str] = {}
    for item in graph.paged_values("/me/todo/lists?$top=100"):
        list_id = normalize_text(item.get("id"))
        name = normalize_text(item.get("displayName"))
        if list_id:
            by_id[list_id] = name
        if list_id and name:
            by_name[name.casefold()] = list_id
    return by_name, by_id


def get_or_create_task_list(graph: GraphClient, by_name: dict[str, str], by_id: dict[str, str], name: str) -> str:
    name = name or "Tarefas"
    if name.casefold() in by_name:
        return by_name[name.casefold()]
    created = graph.json("POST", "/me/todo/lists", expected=(200, 201), json={"displayName": name})
    list_id = normalize_text(created.get("id"))
    if not list_id:
        raise RuntimeError("Lista criada sem ID.")
    by_name[name.casefold()] = list_id
    by_id[list_id] = name
    log(f"Lista criada no Microsoft To Do: {name}")
    return list_id


def get_task(graph: GraphClient, list_id: str, task_id: str) -> dict[str, Any] | None:
    response = graph.request("GET", f"/me/todo/lists/{quote(list_id, safe='')}/tasks/{quote(task_id, safe='')}")
    if response.status_code == 404:
        return None
    if response.status_code != 200:
        raise RuntimeError(f"Falha ao consultar tarefa externa: {response.status_code} {response.text}")
    return response.json()


def create_task(graph: GraphClient, list_id: str, payload: dict[str, Any]) -> ExternalRef:
    created = graph.json("POST", f"/me/todo/lists/{quote(list_id, safe='')}/tasks", expected=(200, 201), json=payload)
    task_id = normalize_text(created.get("id"))
    if not task_id:
        raise RuntimeError("Tarefa criada sem ID.")
    return ExternalRef(list_id, task_id)


def update_task(graph: GraphClient, ref: ExternalRef, payload: dict[str, Any]) -> None:
    graph.json("PATCH", f"/me/todo/lists/{quote(ref.list_id, safe='')}/tasks/{quote(ref.task_id, safe='')}", expected=(200,), json=payload)


def delete_task(graph: GraphClient, ref: ExternalRef) -> None:
    graph.json("DELETE", f"/me/todo/lists/{quote(ref.list_id, safe='')}/tasks/{quote(ref.task_id, safe='')}", expected=(200, 204))


def locate_task_across_lists(graph: GraphClient, task_id: str, list_ids: list[str]) -> ExternalRef | None:
    for list_id in list_ids:
        if get_task(graph, list_id, task_id) is not None:
            return ExternalRef(list_id, task_id)
    return None


def resolve_external_ref(graph: GraphClient, raw_ref: ExternalRef | str | None, list_ids: list[str]) -> ExternalRef | None:
    if raw_ref is None:
        return None
    if isinstance(raw_ref, ExternalRef):
        if get_task(graph, raw_ref.list_id, raw_ref.task_id) is not None:
            return raw_ref
        return locate_task_across_lists(graph, raw_ref.task_id, [x for x in list_ids if x != raw_ref.list_id])
    return locate_task_across_lists(graph, raw_ref, list_ids)


def todo_payload(record: dict[str, Any], complete: bool) -> dict[str, Any]:
    body = [f"Second Brain ID: {normalize_text(record.get('ID'))}"]
    if normalize_text(record.get("Área")):
        body.append(f"Área: {normalize_text(record.get('Área'))}")
    if normalize_text(record.get("Observações")):
        body.extend(["", normalize_text(record.get("Observações"))])
    due = excel_value_to_date(record.get("Data"))
    return {
        "title": normalize_text(record.get("Tarefa")),
        "importance": normalize_priority(record.get("Prioridade")),
        "status": "completed" if complete else "notStarted",
        "body": {"content": "\n".join(body), "contentType": "text"},
        "dueDateTime": graph_date_time(due) if due else None,
    }


def process_todo_record(graph: GraphClient, by_name: dict[str, str], by_id: dict[str, str], record: dict[str, Any]) -> tuple[str, ExternalRef | None]:
    local_id = normalize_text(record.get("ID"))
    title = normalize_text(record.get("Tarefa"))
    if not local_id and not title:
        return "ignorar", None
    if not local_id or not title:
        raise ValueError("Registro da tbTodo sem ID ou titulo.")
    list_id = get_or_create_task_list(graph, by_name, by_id, normalize_text(record.get("Lista")) or "Tarefas")
    raw_ref = parse_external_ref(record.get("Destino Externo ID"))
    resolved = resolve_external_ref(graph, raw_ref, list(by_id))
    status = normalize_status(record.get("Status")) or "pendente"
    delete_statuses = {"cancelado", "cancelada", "inativo", "inativa", "arquivado", "arquivada"}
    complete_statuses = {"concluido", "concluida", "completed", "feito", "feita"}
    if status in delete_statuses:
        if resolved:
            delete_task(graph, resolved)
            return "excluida", None
        if raw_ref is not None:
            raise RuntimeError(f"{local_id}: vinculo preenchido, mas tarefa nao localizada.")
        return "sem alteracao", None
    payload = todo_payload(record, status in complete_statuses)
    if raw_ref is not None and resolved is None and not RECREATE_MISSING:
        raise RuntimeError(f"{local_id}: tarefa vinculada nao localizada. Use SB_RECREATE_MISSING=1 para recriar.")
    if resolved is None:
        return ("recriada" if raw_ref is not None else "criada"), create_task(graph, list_id, payload)
    if resolved.list_id != list_id:
        new_ref = create_task(graph, list_id, payload)
        try:
            delete_task(graph, resolved)
        except Exception:
            delete_task(graph, new_ref)
            raise
        return "movida", new_ref
    update_task(graph, resolved, payload)
    return "atualizada", resolved


def task_marker(task: dict[str, Any]) -> str:
    content = normalize_text((task.get("body") or {}).get("content"))
    match = re.search(r"(?mi)^Second Brain ID:\s*(\S+)\s*$", content)
    return match.group(1).strip() if match else ""


def wishlist_tasks_by_local_id(graph: GraphClient, list_id: str) -> dict[str, ExternalRef]:
    result: dict[str, ExternalRef] = {}
    url = f"/me/todo/lists/{quote(list_id, safe='')}/tasks?$top=100"
    for task in graph.paged_values(url):
        marker = task_marker(task)
        task_id = normalize_text(task.get("id"))
        if marker and task_id:
            if marker.casefold() in result:
                raise RuntimeError(f"Wishlist possui tarefas duplicadas para o ID {marker}.")
            result[marker.casefold()] = ExternalRef(list_id, task_id)
    return result


def wishlist_payload(record: dict[str, Any]) -> dict[str, Any]:
    local_id = normalize_text(record.get("ID"))
    lines = [f"Second Brain ID: {local_id}", "Origem: tbListas / Wishlist"]
    area = normalize_text(record.get("Área"))
    value = normalize_text(record.get("Valor"))
    details = normalize_text(record.get("Detalhes"))
    reference = normalize_text(record.get("Referência"))
    if area:
        lines.append(f"Área: {area}")
    if value:
        lines.append(f"Valor: {value}")
    if details:
        lines.extend(["", details])
    if reference:
        lines.extend(["", f"Referência: {reference}"])
    due = excel_value_to_date(record.get("Data"))
    return {
        "title": normalize_text(record.get("Item")),
        "importance": "normal",
        "status": "notStarted",
        "body": {"content": "\n".join(lines), "contentType": "text"},
        "dueDateTime": graph_date_time(due) if due else None,
    }


def sync_wishlist(graph: GraphClient, by_name: dict[str, str], by_id: dict[str, str], headers: list[str], rows: list[dict[str, Any]]) -> Counter:
    counters: Counter = Counter()
    wishlist_list_id = get_or_create_task_list(graph, by_name, by_id, WISHLIST_NAME)
    existing = wishlist_tasks_by_local_id(graph, wishlist_list_id)
    seen_ids: set[str] = set()
    delete_statuses = {"cancelado", "cancelada", "inativo", "inativa", "arquivado", "arquivada", "removido", "removida"}

    for fallback, row_info in enumerate(rows):
        values = list((row_info.get("values") or [[]])[0])
        values += [""] * max(0, len(headers) - len(values))
        record = {header: values[index] for index, header in enumerate(headers)}
        if normalize_text(record.get("Lista")).casefold() != WISHLIST_NAME.casefold():
            counters["ignorada"] += 1
            continue
        local_id = normalize_text(record.get("ID"))
        title = normalize_text(record.get("Item"))
        if not local_id or not title:
            raise ValueError(f"tbListas linha {fallback + 1}: Wishlist sem ID ou Item.")
        key = local_id.casefold()
        if key in seen_ids:
            raise RuntimeError(f"ID duplicado na Wishlist: {local_id}")
        seen_ids.add(key)
        ref = existing.get(key)
        status = normalize_status(record.get("Status"))
        if status in delete_statuses:
            if ref:
                delete_task(graph, ref)
                counters["excluida"] += 1
                log(f"Wishlist {local_id}: excluida.")
            else:
                counters["sem alteracao"] += 1
            continue
        payload = wishlist_payload(record)
        if ref:
            update_task(graph, ref, payload)
            counters["atualizada"] += 1
            log(f"Wishlist {local_id}: atualizada.")
        else:
            create_task(graph, wishlist_list_id, payload)
            counters["criada"] += 1
            log(f"Wishlist {local_id}: criada.")
    return counters


def validate_todo_unique(headers: list[str], rows: list[dict[str, Any]]) -> None:
    ids, refs = [], []
    for row_info in rows:
        values = list((row_info.get("values") or [[]])[0])
        values += [""] * max(0, len(headers) - len(values))
        record = {header: values[index] for index, header in enumerate(headers)}
        local_id = normalize_text(record.get("ID"))
        if local_id:
            ids.append(local_id.casefold())
        ref = parse_external_ref(record.get("Destino Externo ID"))
        if isinstance(ref, ExternalRef):
            refs.append(f"{ref.list_id}:{ref.task_id}")
        elif isinstance(ref, str):
            refs.append(ref)
    duplicate_ids = [key for key, count in Counter(ids).items() if count > 1]
    duplicate_refs = [key for key, count in Counter(refs).items() if count > 1]
    if duplicate_ids or duplicate_refs:
        raise RuntimeError(f"Duplicacoes na tbTodo. IDs={duplicate_ids}; externos={duplicate_refs}")


def sync_todo(graph: GraphClient, item_id: str, session_id: str, by_name: dict[str, str], by_id: dict[str, str], headers: list[str], rows: list[dict[str, Any]]) -> Counter:
    counters: Counter = Counter()
    validate_todo_unique(headers, rows)
    indexes = {name: headers.index(name) for name in headers}
    for fallback, row_info in enumerate(rows):
        row_position = int(row_info.get("index", fallback))
        values = list((row_info.get("values") or [[]])[0])
        values += [""] * max(0, len(headers) - len(values))
        values = values[:len(headers)]
        record = {header: values[index] for index, header in enumerate(headers)}
        local_id = normalize_text(record.get("ID")) or f"linha {fallback + 1}"
        try:
            action, ref = process_todo_record(graph, by_name, by_id, record)
            counters[action] += 1
            if action in {"ignorar", "sem alteracao"}:
                continue
            values[indexes["Destino Externo ID"]] = ref.serialize() if ref else ""
            values[indexes["Última atualização"]] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            update_excel_row(graph, item_id, session_id, TODO_TABLE, row_position, values)
            log(f"{local_id}: {action}.")
        except Exception as exc:
            counters["erro"] += 1
            log(f"ERRO em {local_id}: {exc}")
    return counters


def format_summary(name: str, counters: Counter) -> str:
    values = ", ".join(f"{key}={value}" for key, value in counters.items() if value)
    return f"{name}: {values or 'nenhuma alteracao'}"


def main() -> int:
    validate_environment()
    cache_state = load_cache()
    token = acquire_token(cache_state)
    graph = GraphClient(token)
    workbook = get_workbook_item(graph)
    item_id = normalize_text(workbook.get("id"))
    if not item_id:
        raise RuntimeError("Microsoft Graph nao retornou o ID do Excel.")
    log(f"Arquivo localizado: {workbook.get('name', 'Second Brain - Dados.xlsx')}")
    session_id = create_workbook_session(graph, item_id)
    processing_error: BaseException | None = None
    try:
        todo_headers, todo_rows = load_table(graph, item_id, session_id, TODO_TABLE, TODO_HEADERS)
        list_headers, list_rows = load_table(graph, item_id, session_id, LISTS_TABLE, LIST_HEADERS)
        by_name, by_id = get_task_lists(graph)
        todo_result = sync_todo(graph, item_id, session_id, by_name, by_id, todo_headers, todo_rows)
        wishlist_result = sync_wishlist(graph, by_name, by_id, list_headers, list_rows)
        log(format_summary("Resumo tbTodo", todo_result))
        log(format_summary("Resumo Wishlist", wishlist_result))
        return 1 if todo_result["erro"] or wishlist_result["erro"] else 0
    except BaseException as exc:
        processing_error = exc
        raise
    finally:
        try:
            close_workbook_session(graph, item_id, session_id)
            log("Sessao do workbook encerrada.")
        except Exception as close_error:
            if processing_error is None:
                raise
            log(f"AVISO: falha ao encerrar a sessao: {close_error}")
        finally:
            save_cache(cache_state)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("Execucao interrompida.", file=sys.stderr)
        sys.exit(130)
    except Exception as exc:
        print(f"ERRO FATAL: {exc}", file=sys.stderr)
        sys.exit(1)
