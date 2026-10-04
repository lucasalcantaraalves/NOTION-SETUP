#!/usr/bin/env python3
"""
SINCRONIZAR-TODO.py

Sincroniza a tabela tbTodo do arquivo "Second Brain - Dados.xlsx"
com o Microsoft To Do por meio do Microsoft Graph.

Fluxo unidirecional:
  tbTodo -> Microsoft To Do

Variaveis obrigatorias:
  MS_CLIENT_ID
  MS_CACHE_KEY

Variaveis opcionais:
  MS_CACHE_FILE       Padrao: .msal_cache.enc
  MS_TENANT_ID        Padrao: consumers
  SB_EXCEL_PATH       Caminho do Excel no OneDrive
  SB_TODO_TABLE       Padrao: tbTodo
  SB_FORCE_LOGIN      Use 1 para forcar Device Code Flow
  SB_TIME_ZONE        Padrao: America/New_York
  SB_RECREATE_MISSING Use 1 para recriar tarefa cujo vinculo externo sumiu

Dependencias:
  pip install msal requests cryptography

Observacao sobre Destino Externo ID:
  Novos vinculos sao salvos como JSON compacto contendo listId e taskId.
  IDs antigos contendo somente taskId continuam sendo aceitos e sao migrados.
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


def is_valid_msal_cache_json(serialized: str) -> bool:
    try:
        payload = json.loads(serialized)
    except (TypeError, ValueError, json.JSONDecodeError):
        return False
    if not isinstance(payload, dict):
        return False
    known_sections = {
        "AccessToken",
        "RefreshToken",
        "IdToken",
        "Account",
        "AppMetadata",
        "TokenType",
    }
    return not payload or bool(known_sections.intersection(payload))


def validate_environment() -> None:
    missing = []
    if not CLIENT_ID:
        missing.append("MS_CLIENT_ID")
    if not CACHE_KEY:
        missing.append("MS_CACHE_KEY")
    if missing:
        raise RuntimeError("Defina as variaveis obrigatorias: " + ", ".join(missing))


def load_cache() -> CacheState:
    cache = msal.SerializableTokenCache()
    path = Path(CACHE_FILE)
    if not path.exists():
        return CacheState(cache=cache)

    data = path.read_bytes()
    if not data:
        return CacheState(cache=cache, force_save=True)

    try:
        serialized = Fernet(normalize_key(CACHE_KEY)).decrypt(data).decode("utf-8")
    except (InvalidToken, UnicodeDecodeError) as encrypted_error:
        try:
            plain_text = data.decode("utf-8")
        except UnicodeDecodeError as exc:
            raise RuntimeError(
                "Nao foi possivel abrir o cache MSAL. Verifique MS_CACHE_KEY ou remova o cache para autenticar novamente."
            ) from exc

        if not is_valid_msal_cache_json(plain_text):
            raise RuntimeError(
                "O cache MSAL nao pode ser descriptografado. MS_CACHE_KEY pode estar incorreta ou o arquivo pode estar corrompido."
            ) from encrypted_error

        serialized = plain_text
        log("AVISO: cache MSAL em texto puro detectado. Ele sera migrado para formato criptografado.")
        force_save = True
    else:
        force_save = False

    if serialized.strip():
        try:
            cache.deserialize(serialized)
        except Exception as exc:
            raise RuntimeError("O conteudo do cache MSAL e invalido ou esta corrompido.") from exc

    return CacheState(cache=cache, force_save=force_save)


def save_cache(state: CacheState, force: bool = False) -> None:
    cache = state.cache
    if not force and not state.force_save and not cache.has_state_changed:
        return

    path = Path(CACHE_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = cache.serialize().encode("utf-8")
    encrypted = Fernet(normalize_key(CACHE_KEY)).encrypt(serialized)

    temp_path = path.with_name(path.name + ".tmp")
    temp_path.write_bytes(encrypted)
    try:
        os.chmod(temp_path, 0o600)
    except OSError:
        pass
    temp_path.replace(path)
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
            raise RuntimeError(f"Falha ao iniciar Device Code Flow: {json.dumps(flow, ensure_ascii=False)}")
        log(flow.get("message", "Abra a pagina indicada e informe o codigo exibido."))
        result = app.acquire_token_by_device_flow(flow)

    if not result or "access_token" not in result:
        description = (result or {}).get("error_description") or (result or {}).get("error") or "erro desconhecido"
        raise RuntimeError(f"Falha na autenticacao Microsoft: {description}")

    save_cache(state)
    return result["access_token"]


class GraphClient:
    def __init__(self, token: str) -> None:
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Content-Type": "application/json",
            }
        )

    def request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        if url.startswith("/"):
            url = GRAPH_ROOT + url
        elif not url.startswith("http"):
            url = GRAPH_ROOT + "/" + url

        response: requests.Response | None = None
        for attempt in range(MAX_RETRIES):
            try:
                response = self.session.request(method, url, timeout=REQUEST_TIMEOUT, **kwargs)
            except requests.RequestException as exc:
                if attempt == MAX_RETRIES - 1:
                    raise RuntimeError(f"Falha de rede ao acessar Microsoft Graph: {exc}") from exc
                wait = min(2**attempt, 30)
                log(f"Falha de rede. Nova tentativa em {wait}s...")
                time.sleep(wait)
                continue

            if response.status_code not in {429, 500, 502, 503, 504}:
                return response

            retry_after = response.headers.get("Retry-After", "")
            try:
                wait = max(1, int(retry_after))
            except ValueError:
                wait = min(2**attempt, 30)
            log(f"Graph temporariamente indisponivel ({response.status_code}). Nova tentativa em {wait}s...")
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
        expected=(200, 201),
        json={"persistChanges": True},
    )
    session_id = normalize_text(data.get("id"))
    if not session_id:
        raise RuntimeError("Microsoft Graph nao retornou o ID da sessao do workbook.")
    return session_id


def close_workbook_session(graph: GraphClient, item_id: str, session_id: str) -> None:
    graph.json(
        "POST",
        f"/me/drive/items/{item_id}/workbook/closeSession",
        expected=(200, 204),
        headers={"workbook-session-id": session_id},
    )


def workbook_headers(session_id: str) -> dict[str, str]:
    return {"workbook-session-id": session_id}


def load_todo_table(
    graph: GraphClient,
    item_id: str,
    session_id: str,
) -> tuple[list[str], list[dict[str, Any]]]:
    headers = workbook_headers(session_id)
    table = quote(TABLE_NAME, safe="")
    base = f"/me/drive/items/{item_id}/workbook/tables/{table}"

    header_payload = graph.json("GET", base + "/headerRowRange", headers=headers)
    header_rows = header_payload.get("values", [])
    if not header_rows:
        raise RuntimeError(f"A tabela {TABLE_NAME} nao possui cabecalho.")

    table_headers = [normalize_text(value) for value in header_rows[0]]
    missing = [name for name in EXPECTED_HEADERS if name not in table_headers]
    if missing:
        raise RuntimeError(f"Colunas ausentes em {TABLE_NAME}: {', '.join(missing)}")

    row_payload = graph.json("GET", base + "/rows", headers=headers)
    rows = row_payload.get("value", [])
    if not isinstance(rows, list):
        raise RuntimeError(f"Resposta invalida ao ler as linhas da tabela {TABLE_NAME}.")
    return table_headers, rows


def update_excel_row(
    graph: GraphClient,
    item_id: str,
    session_id: str,
    row_position: int,
    values: list[Any],
) -> None:
    table = quote(TABLE_NAME, safe="")
    url = (
        f"/me/drive/items/{item_id}/workbook/tables/{table}"
        f"/rows/itemAt(index={row_position})/range"
    )
    graph.json(
        "PATCH",
        url,
        expected=(200,),
        headers=workbook_headers(session_id),
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
            continue
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise ValueError(f"Data invalida: {text}") from exc


def graph_date_time(day: date) -> dict[str, str]:
    return {"dateTime": f"{day.isoformat()}T09:00:00.0000000", "timeZone": TIME_ZONE}


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


def parse_external_ref(value: Any) -> ExternalRef | str | None:
    text = normalize_text(value)
    if not text:
        return None
    if text.startswith("{"):
        try:
            payload = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError("Destino Externo ID contem JSON invalido.") from exc
        list_id = normalize_text(payload.get("listId"))
        task_id = normalize_text(payload.get("taskId"))
        if not list_id or not task_id:
            raise ValueError("Destino Externo ID deve conter listId e taskId.")
        return ExternalRef(list_id=list_id, task_id=task_id)
    return text


def validate_unique_rows(headers: list[str], rows: list[dict[str, Any]]) -> None:
    id_index = headers.index("ID")
    external_index = headers.index("Destino Externo ID")
    local_ids: list[str] = []
    external_keys: list[str] = []

    for row_info in rows:
        values = list(row_info.get("values") or [[]])[0]
        values += [""] * max(0, len(headers) - len(values))
        local_id = normalize_text(values[id_index])
        if local_id:
            local_ids.append(local_id.casefold())
        external = parse_external_ref(values[external_index])
        if isinstance(external, ExternalRef):
            external_keys.append(f"{external.list_id}:{external.task_id}")
        elif isinstance(external, str):
            external_keys.append(external)

    duplicate_ids = sorted(key for key, count in Counter(local_ids).items() if count > 1)
    duplicate_external = sorted(key for key, count in Counter(external_keys).items() if count > 1)
    problems = []
    if duplicate_ids:
        problems.append("IDs locais duplicados: " + ", ".join(duplicate_ids))
    if duplicate_external:
        problems.append("Destinos externos duplicados: " + ", ".join(duplicate_external))
    if problems:
        raise RuntimeError(" | ".join(problems))


def get_task_lists(graph: GraphClient) -> tuple[dict[str, str], dict[str, str]]:
    lists = graph.paged_values("/me/todo/lists?$top=100")
    by_name: dict[str, str] = {}
    by_id: dict[str, str] = {}
    for item in lists:
        list_id = normalize_text(item.get("id"))
        name = normalize_text(item.get("displayName"))
        if list_id:
            by_id[list_id] = name
        if list_id and name:
            by_name[name.casefold()] = list_id
    return by_name, by_id


def get_or_create_task_list(
    graph: GraphClient,
    by_name: dict[str, str],
    by_id: dict[str, str],
    name: str,
) -> str:
    clean_name = name or "Tarefas"
    key = clean_name.casefold()
    if key in by_name:
        return by_name[key]

    created = graph.json(
        "POST",
        "/me/todo/lists",
        expected=(200, 201),
        json={"displayName": clean_name},
    )
    list_id = normalize_text(created.get("id"))
    if not list_id:
        raise RuntimeError("Microsoft Graph criou a lista, mas nao retornou seu ID.")
    by_name[key] = list_id
    by_id[list_id] = clean_name
    log(f"Lista criada no Microsoft To Do: {clean_name}")
    return list_id


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


def get_task(graph: GraphClient, list_id: str, task_id: str) -> dict[str, Any] | None:
    response = graph.request(
        "GET",
        f"/me/todo/lists/{quote(list_id, safe='')}/tasks/{quote(task_id, safe='')}",
    )
    if response.status_code == 404:
        return None
    if response.status_code != 200:
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise RuntimeError(f"Falha ao consultar tarefa externa ({response.status_code}): {detail}")
    return response.json()


def locate_task_across_lists(
    graph: GraphClient,
    task_id: str,
    list_ids: list[str],
) -> ExternalRef | None:
    for list_id in list_ids:
        if get_task(graph, list_id, task_id) is not None:
            return ExternalRef(list_id=list_id, task_id=task_id)
    return None


def create_task(graph: GraphClient, list_id: str, payload: dict[str, Any]) -> ExternalRef:
    created = graph.json(
        "POST",
        f"/me/todo/lists/{quote(list_id, safe='')}/tasks",
        expected=(200, 201),
        json=payload,
    )
    task_id = normalize_text(created.get("id"))
    if not task_id:
        raise RuntimeError("Microsoft Graph criou a tarefa, mas nao retornou seu ID.")
    return ExternalRef(list_id=list_id, task_id=task_id)


def update_task(graph: GraphClient, ref: ExternalRef, payload: dict[str, Any]) -> None:
    graph.json(
        "PATCH",
        f"/me/todo/lists/{quote(ref.list_id, safe='')}/tasks/{quote(ref.task_id, safe='')}",
        expected=(200,),
        json=payload,
    )


def delete_task(graph: GraphClient, ref: ExternalRef) -> None:
    graph.json(
        "DELETE",
        f"/me/todo/lists/{quote(ref.list_id, safe='')}/tasks/{quote(ref.task_id, safe='')}",
        expected=(200, 204),
    )


def move_task_between_lists(
    graph: GraphClient,
    current_ref: ExternalRef,
    target_list_id: str,
    payload: dict[str, Any],
) -> ExternalRef:
    new_ref = create_task(graph, target_list_id, payload)
    try:
        delete_task(graph, current_ref)
    except Exception:
        try:
            delete_task(graph, new_ref)
        except Exception as rollback_error:
            log(f"AVISO: falha ao desfazer a nova tarefa apos erro de movimentacao: {rollback_error}")
        raise
    return new_ref


def resolve_external_ref(
    graph: GraphClient,
    raw_ref: ExternalRef | str | None,
    known_list_ids: list[str],
) -> ExternalRef | None:
    if raw_ref is None:
        return None
    if isinstance(raw_ref, ExternalRef):
        if get_task(graph, raw_ref.list_id, raw_ref.task_id) is not None:
            return raw_ref
        other_lists = [list_id for list_id in known_list_ids if list_id != raw_ref.list_id]
        return locate_task_across_lists(graph, raw_ref.task_id, other_lists)
    return locate_task_across_lists(graph, raw_ref, known_list_ids)


def process_row(
    graph: GraphClient,
    by_name: dict[str, str],
    by_id: dict[str, str],
    record: dict[str, Any],
) -> tuple[str, ExternalRef | None]:
    local_id = normalize_text(record.get("ID"))
    title = normalize_text(record.get("Tarefa"))
    list_name = normalize_text(record.get("Lista")) or "Tarefas"
    status = normalize_status(record.get("Status")) or "pendente"
    raw_ref = parse_external_ref(record.get("Destino Externo ID"))

    if not local_id and not title:
        return "ignorar", None
    if not local_id:
        raise ValueError("Registro sem ID.")
    if not title:
        raise ValueError(f"{local_id}: tarefa sem titulo.")

    target_list_id = get_or_create_task_list(graph, by_name, by_id, list_name)
    known_list_ids = list(by_id.keys())
    resolved_ref = resolve_external_ref(graph, raw_ref, known_list_ids)

    complete_statuses = {"concluido", "concluida", "completed", "feito", "feita"}
    delete_statuses = {"cancelado", "cancelada", "inativo", "inativa", "arquivado", "arquivada"}

    if status in delete_statuses:
        if raw_ref is not None and resolved_ref is None:
            raise RuntimeError(
                f"{local_id}: Destino Externo ID preenchido, mas a tarefa nao foi localizada. Vinculo preservado para investigacao."
            )
        if resolved_ref is not None:
            delete_task(graph, resolved_ref)
            return "excluida", None
        return "sem alteracao", None

    complete = status in complete_statuses
    payload = build_task_payload(record, complete=complete)

    if raw_ref is not None and resolved_ref is None:
        if not RECREATE_MISSING:
            raise RuntimeError(
                f"{local_id}: Destino Externo ID preenchido, mas a tarefa nao foi localizada. "
                "Defina SB_RECREATE_MISSING=1 somente se desejar recria-la conscientemente."
            )
        created = create_task(graph, target_list_id, payload)
        return "recriada", created

    if resolved_ref is None:
        created = create_task(graph, target_list_id, payload)
        return "criada", created

    if resolved_ref.list_id != target_list_id:
        moved = move_task_between_lists(graph, resolved_ref, target_list_id, payload)
        return "movida", moved

    update_task(graph, resolved_ref, payload)
    return "atualizada", resolved_ref


def main() -> int:
    validate_environment()
    cache_state = load_cache()
    token = acquire_token(cache_state)
    graph = GraphClient(token)

    workbook = get_workbook_item(graph)
    item_id = normalize_text(workbook.get("id"))
    if not item_id:
        raise RuntimeError("Microsoft Graph nao retornou o ID do arquivo Excel.")
    log(f"Arquivo localizado: {workbook.get('name', 'Second Brain - Dados.xlsx')}")

    session_id = create_workbook_session(graph, item_id)
    processing_error: BaseException | None = None
    counters = {
        "criada": 0,
        "recriada": 0,
        "atualizada": 0,
        "movida": 0,
        "excluida": 0,
        "sem alteracao": 0,
        "ignorar": 0,
        "erro": 0,
    }

    try:
        headers, row_infos = load_todo_table(graph, item_id, session_id)
        if not row_infos:
            log(f"A tabela {TABLE_NAME} esta vazia. Nada para sincronizar.")
            return 0

        validate_unique_rows(headers, row_infos)
        indexes = {name: headers.index(name) for name in headers}
        by_name, by_id = get_task_lists(graph)

        for fallback_position, row_info in enumerate(row_infos):
            row_position = row_info.get("index", fallback_position)
            raw_values = row_info.get("values") or [[]]
            row = list(raw_values[0] if raw_values else [])
            row += [""] * max(0, len(headers) - len(row))
            row = row[: len(headers)]
            record = {header: row[index] for index, header in enumerate(headers)}
            local_id = normalize_text(record.get("ID")) or f"linha {fallback_position + 1}"

            try:
                action, external_ref = process_row(graph, by_name, by_id, record)
                counters[action] += 1
                if action == "ignorar":
                    continue

                if action == "sem alteracao":
                    log(f"{local_id}: sem alteracao.")
                    continue

                destination_value = external_ref.serialize() if external_ref else ""
                row[indexes["Destino Externo ID"]] = destination_value
                row[indexes["Última atualização"]] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                update_excel_row(graph, item_id, session_id, int(row_position), row)
                log(f"{local_id}: {action}.")
            except Exception as exc:
                counters["erro"] += 1
                log(f"ERRO em {local_id}: {exc}")

        summary = ", ".join(f"{key}={value}" for key, value in counters.items() if value)
        log("Resumo: " + (summary or "nenhuma alteracao"))
        return 1 if counters["erro"] else 0
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
            log(f"AVISO: falha ao encerrar a sessao do workbook: {close_error}")
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
