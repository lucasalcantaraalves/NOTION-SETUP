#!/usr/bin/env python3
"""
SINCRONIZAR-TODO.py

Sincroniza a tabela tbTodo do arquivo "Second Brain - Dados.xlsx"
com o Microsoft To Do por meio do Microsoft Graph.

Variáveis de ambiente obrigatórias:
  MS_CLIENT_ID

Variáveis opcionais:
  MS_CACHE_KEY       Chave usada para criptografar o cache MSAL.
  MS_CACHE_FILE      Caminho do cache. Padrao: .msal_cache.enc
  MS_TENANT_ID       Padrao: consumers
  SB_EXCEL_PATH      Caminho no OneDrive.
  SB_TODO_TABLE      Padrao: tbTodo
  SB_FORCE_LOGIN     Use 1 para forcar Device Code Flow e renovar consentimento.

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
from datetime import date, datetime, timedelta, timezone
from typing import Any
from urllib.parse import quote

import msal
import requests
from cryptography.fernet import Fernet, InvalidToken

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
CLIENT_ID = os.getenv("MS_CLIENT_ID", "").strip()
TENANT_ID = os.getenv("MS_TENANT_ID", "consumers").strip() or "consumers"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.ReadWrite", "Tasks.ReadWrite"]

CACHE_FILE = os.getenv("MS_CACHE_FILE", ".msal_cache.enc").strip() or ".msal_cache.enc"
CACHE_KEY = os.getenv("MS_CACHE_KEY", "").strip()
FORCE_LOGIN = os.getenv("SB_FORCE_LOGIN", "").strip().lower() in {"1", "true", "sim", "yes"}

EXCEL_PATH = os.getenv(
    "SB_EXCEL_PATH",
    "/Microsoft Copilot Chat Files/Copilot Notebook Uploads/Second Brain - Dados.xlsx",
).strip()
TABLE_NAME = os.getenv("SB_TODO_TABLE", "tbTodo").strip() or "tbTodo"

REQUEST_TIMEOUT = 60
MAX_RETRIES = 5
EXPECTED_HEADERS = [
    "ID",
    "Tarefa",
    "Lista",
    "Área",
    "Status",
    "Data",
    "Prioridade",
    "Observações",
    "Última atualização",
    "Destino Externo ID",
]


def log(message: str) -> None:
    print(message, flush=True)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_key(value: str) -> bytes:
    raw = value.encode("utf-8")
    try:
        decoded = base64.urlsafe_b64decode(raw)
        if len(decoded) == 32:
            return raw
    except Exception:
        pass
    return base64.urlsafe_b64encode(hashlib.sha256(raw).digest())


def load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    path = Path(CACHE_FILE)
    if not path.exists():
        return cache

    data = path.read_bytes()
    serialized = ""

    if CACHE_KEY:
        try:
            serialized = Fernet(normalize_key(CACHE_KEY)).decrypt(data).decode("utf-8")
        except InvalidToken:
            try:
                serialized = data.decode("utf-8")
                log("AVISO: cache existente estava em texto puro; sera criptografado ao salvar.")
            except UnicodeDecodeError as exc:
                raise RuntimeError("Nao foi possivel descriptografar o cache MSAL.") from exc
    else:
        try:
            serialized = data.decode("utf-8")
        except UnicodeDecodeError as exc:
            raise RuntimeError("MS_CACHE_KEY e obrigatoria para ler este cache criptografado.") from exc

    if serialized.strip():
        cache.deserialize(serialized)
    return cache


def save_cache(cache: msal.SerializableTokenCache, force: bool = False) -> None:
    if not force and not cache.has_state_changed:
        return

    serialized = cache.serialize().encode("utf-8")
    output = serialized
    if CACHE_KEY:
        output = Fernet(normalize_key(CACHE_KEY)).encrypt(serialized)

    path = Path(CACHE_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(output)
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass
    log(f"Cache MSAL salvo em: {path}")


def acquire_token(cache: msal.SerializableTokenCache) -> str:
    if not CLIENT_ID:
        raise RuntimeError("Defina a variavel de ambiente MS_CLIENT_ID.")

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    result: dict[str, Any] | None = None

    if not FORCE_LOGIN:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result or "access_token" not in result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Falha ao iniciar Device Code Flow: {json.dumps(flow, ensure_ascii=False)}")
        log(flow.get("message", "Abra a pagina indicada e informe o codigo exibido."))
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        description = result.get("error_description") or result.get("error") or "erro desconhecido"
        raise RuntimeError(f"Falha na autenticacao Microsoft: {description}")

    save_cache(cache)
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

        for attempt in range(MAX_RETRIES):
            response = self.session.request(method, url, timeout=REQUEST_TIMEOUT, **kwargs)
            if response.status_code not in {429, 500, 502, 503, 504}:
                return response
            wait = int(response.headers.get("Retry-After", min(2 ** attempt, 30)))
            log(f"Graph temporariamente indisponivel ({response.status_code}). Nova tentativa em {wait}s...")
            time.sleep(wait)
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
        values: list[dict[str, Any]] = []
        next_url: str | None = url
        while next_url:
            payload = self.json("GET", next_url)
            values.extend(payload.get("value", []))
            next_url = payload.get("@odata.nextLink")
        return values


def encode_onedrive_path(path: str) -> str:
    parts = [quote(part, safe="") for part in path.strip("/").split("/") if part]
    return "/".join(parts)


def get_workbook_item(graph: GraphClient) -> dict[str, Any]:
    encoded = encode_onedrive_path(EXCEL_PATH)
    return graph.json("GET", f"/me/drive/root:/{encoded}")


def create_workbook_session(graph: GraphClient, item_id: str) -> str:
    data = graph.json(
        "POST",
        f"/me/drive/items/{item_id}/workbook/createSession",
        expected=(201,),
        json={"persistChanges": True},
    )
    return data["id"]


def close_workbook_session(graph: GraphClient, item_id: str, session_id: str) -> None:
    headers = {"workbook-session-id": session_id}
    graph.json(
        "POST",
        f"/me/drive/items/{item_id}/workbook/closeSession",
        expected=(204,),
        headers=headers,
    )


def workbook_headers(session_id: str) -> dict[str, str]:
    return {"workbook-session-id": session_id}


def load_todo_table(graph: GraphClient, item_id: str, session_id: str) -> tuple[list[str], list[list[Any]]]:
    headers = workbook_headers(session_id)
    base = f"/me/drive/items/{item_id}/workbook/tables/{quote(TABLE_NAME, safe='')}"
    header_payload = graph.json("GET", base + "/headerRowRange", headers=headers)
    body_payload = graph.json("GET", base + "/dataBodyRange", headers=headers)

    header_rows = header_payload.get("values", [])
    if not header_rows:
        raise RuntimeError(f"A tabela {TABLE_NAME} nao possui cabecalho.")
    table_headers = [normalize_text(v) for v in header_rows[0]]

    missing = [name for name in EXPECTED_HEADERS if name not in table_headers]
    if missing:
        raise RuntimeError(f"Colunas ausentes em {TABLE_NAME}: {', '.join(missing)}")

    return table_headers, body_payload.get("values", [])


def update_excel_row(graph: GraphClient, item_id: str, session_id: str, row_index: int, values: list[Any]) -> None:
    headers = workbook_headers(session_id)
    url = (
        f"/me/drive/items/{item_id}/workbook/tables/{quote(TABLE_NAME, safe='')}"
        f"/rows/itemAt(index={row_index})/range"
    )
    graph.json("PATCH", url, headers=headers, json={"values": [values]})


def excel_value_to_date(value: Any) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    text = normalize_text(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise ValueError(f"Data invalida: {text}") from exc


def graph_date_time(day: date) -> dict[str, str]:
    return {"dateTime": f"{day.isoformat()}T09:00:00.0000000", "timeZone": "America/New_York"}


def normalize_status(value: Any) -> str:
    text = normalize_text(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"\s+", " ", text)


def normalize_priority(value: Any) -> str:
    text = normalize_status(value)
    if text in {"alta", "high", "urgente"}:
        return "high"
    if text in {"baixa", "low"}:
        return "low"
    return "normal"


def get_task_lists(graph: GraphClient) -> dict[str, str]:
    lists = graph.paged_values("/me/todo/lists?$top=100")
    return {normalize_text(item.get("displayName")).casefold(): item["id"] for item in lists}


def get_or_create_task_list(graph: GraphClient, cache: dict[str, str], name: str) -> str:
    clean_name = name or "Tarefas"
    key = clean_name.casefold()
    if key in cache:
        return cache[key]
    created = graph.json(
        "POST",
        "/me/todo/lists",
        expected=(201,),
        json={"displayName": clean_name},
    )
    cache[key] = created["id"]
    log(f"Lista criada no Microsoft To Do: {clean_name}")
    return created["id"]


def build_task_payload(record: dict[str, Any], complete: bool = False) -> dict[str, Any]:
    task_id = normalize_text(record.get("ID"))
    area = normalize_text(record.get("Área"))
    notes = normalize_text(record.get("Observações"))
    body_lines = [f"Second Brain ID: {task_id}"]
    if area:
        body_lines.append(f"Área: {area}")
    if notes:
        body_lines.extend(["", notes])

    payload: dict[str, Any] = {
        "title": normalize_text(record.get("Tarefa")),
        "importance": normalize_priority(record.get("Prioridade")),
        "status": "completed" if complete else "notStarted",
        "body": {"content": "\n".join(body_lines), "contentType": "text"},
    }
    due = excel_value_to_date(record.get("Data"))
    payload["dueDateTime"] = graph_date_time(due) if due else None
    return payload


def task_exists(graph: GraphClient, list_id: str, task_id: str) -> bool:
    response = graph.request("GET", f"/me/todo/lists/{quote(list_id, safe='')}/tasks/{quote(task_id, safe='')}")
    if response.status_code == 404:
        return False
    if response.status_code != 200:
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise RuntimeError(f"Falha ao consultar tarefa externa ({response.status_code}): {detail}")
    return True


def create_task(graph: GraphClient, list_id: str, payload: dict[str, Any]) -> dict[str, Any]:
    return graph.json(
        "POST",
        f"/me/todo/lists/{quote(list_id, safe='')}/tasks",
        expected=(201,),
        json=payload,
    )


def update_task(graph: GraphClient, list_id: str, task_id: str, payload: dict[str, Any]) -> None:
    graph.json(
        "PATCH",
        f"/me/todo/lists/{quote(list_id, safe='')}/tasks/{quote(task_id, safe='')}",
        json=payload,
    )


def delete_task(graph: GraphClient, list_id: str, task_id: str) -> None:
    graph.json(
        "DELETE",
        f"/me/todo/lists/{quote(list_id, safe='')}/tasks/{quote(task_id, safe='')}",
        expected=(204,),
    )


def process_row(
    graph: GraphClient,
    list_cache: dict[str, str],
    record: dict[str, Any],
) -> tuple[str, str | None]:
    local_id = normalize_text(record.get("ID"))
    title = normalize_text(record.get("Tarefa"))
    list_name = normalize_text(record.get("Lista")) or "Tarefas"
    status = normalize_status(record.get("Status")) or "pendente"
    external_id = normalize_text(record.get("Destino Externo ID"))

    if not local_id and not title:
        return "ignorar", None
    if not local_id:
        raise ValueError("Registro sem ID.")
    if not title:
        raise ValueError(f"{local_id}: tarefa sem titulo.")

    list_id = get_or_create_task_list(graph, list_cache, list_name)
    complete_statuses = {"concluido", "concluida", "completed", "feito", "feita"}
    delete_statuses = {"cancelado", "cancelada", "inativo", "inativa", "arquivado", "arquivada"}

    if status in delete_statuses:
        if external_id and task_exists(graph, list_id, external_id):
            delete_task(graph, list_id, external_id)
            return "excluida", ""
        return "sem alteracao", ""

    complete = status in complete_statuses
    payload = build_task_payload(record, complete=complete)

    if external_id and task_exists(graph, list_id, external_id):
        update_task(graph, list_id, external_id, payload)
        return "atualizada", external_id

    created = create_task(graph, list_id, payload)
    return "criada", created["id"]


def main() -> int:
    cache = load_cache()
    token = acquire_token(cache)
    graph = GraphClient(token)

    workbook = get_workbook_item(graph)
    item_id = workbook["id"]
    log(f"Arquivo localizado: {workbook.get('name', 'Second Brain - Dados.xlsx')}")

    session_id = create_workbook_session(graph, item_id)
    counters = {"criada": 0, "atualizada": 0, "excluida": 0, "sem alteracao": 0, "ignorar": 0, "erro": 0}

    try:
        headers, rows = load_todo_table(graph, item_id, session_id)
        if not rows:
            log(f"A tabela {TABLE_NAME} esta vazia. Nada para sincronizar.")
            return 0

        indexes = {name: headers.index(name) for name in headers}
        list_cache = get_task_lists(graph)

        for row_index, original_row in enumerate(rows):
            row = list(original_row) + [""] * max(0, len(headers) - len(original_row))
            record = {header: row[index] for index, header in enumerate(headers)}
            local_id = normalize_text(record.get("ID")) or f"linha {row_index + 1}"

            try:
                action, external_id = process_row(graph, list_cache, record)
                counters[action] += 1
                if action == "ignorar":
                    continue

                changed = False
                if external_id is not None and normalize_text(row[indexes["Destino Externo ID"]]) != external_id:
                    row[indexes["Destino Externo ID"]] = external_id
                    changed = True

                now_value = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                row[indexes["Última atualização"]] = now_value
                changed = True

                if changed:
                    update_excel_row(graph, item_id, session_id, row_index, row[: len(headers)])

                log(f"{local_id}: {action}.")
            except Exception as exc:
                counters["erro"] += 1
                log(f"ERRO em {local_id}: {exc}")

        log(
            "Resumo: "
            + ", ".join(f"{key}={value}" for key, value in counters.items() if value)
        )
        return 1 if counters["erro"] else 0
    finally:
        try:
            close_workbook_session(graph, item_id, session_id)
            log("Sessao do workbook encerrada.")
        finally:
            save_cache(cache)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("Execucao interrompida.", file=sys.stderr)
        sys.exit(130)
    except Exception as exc:
        print(f"ERRO FATAL: {exc}", file=sys.stderr)
        sys.exit(1)
