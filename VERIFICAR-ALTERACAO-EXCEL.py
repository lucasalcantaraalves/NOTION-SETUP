#!/usr/bin/env python3
"""Verifica se o Excel oficial do Second Brain mudou no OneDrive."""

from __future__ import annotations

import base64
import hashlib
import json
import os
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote

import msal
import requests
from cryptography.fernet import Fernet, InvalidToken

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
CLIENT_ID = os.getenv("MS_CLIENT_ID", "").strip()
TENANT_ID = os.getenv("MS_TENANT_ID", "consumers").strip() or "consumers"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.ReadWrite"]

CACHE_FILE = Path(os.getenv("MS_CACHE_FILE", ".auth/msal_cache.bin"))
CACHE_KEY = os.getenv("MS_CACHE_KEY", "").strip()
STATE_DIR = Path(os.getenv("SB_STATE_CACHE_DIR", ".state"))
STATE_FILE = STATE_DIR / "excel-state.json"
EXCEL_PATH = os.getenv(
    "SB_EXCEL_PATH",
    "/Microsoft Copilot Chat Files/Copilot Notebook Uploads/Second Brain - Dados.xlsx",
).strip()
FORCE_FULL = os.getenv("SB_FORCE_FULL", "false").strip().lower() in {
    "1", "true", "yes", "sim"
}
TIMEOUT = 60


def log(message: str) -> None:
    print(message, flush=True)


def github_output(name: str, value: str) -> None:
    output_file = os.getenv("GITHUB_OUTPUT", "").strip()
    if not output_file:
        log(f"OUTPUT {name}={value}")
        return
    with open(output_file, "a", encoding="utf-8") as handle:
        handle.write(f"{name}={value}\n")


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
    if not CLIENT_ID:
        raise RuntimeError("MS_CLIENT_ID nao configurado.")
    if not CACHE_KEY:
        raise RuntimeError("MS_CACHE_KEY nao configurado.")

    cache = msal.SerializableTokenCache()
    if not CACHE_FILE.exists():
        raise RuntimeError(f"Cache MSAL nao encontrado: {CACHE_FILE}")

    encrypted = CACHE_FILE.read_bytes()
    try:
        serialized = Fernet(normalize_key(CACHE_KEY)).decrypt(encrypted).decode("utf-8")
    except InvalidToken as exc:
        raise RuntimeError("Nao foi possivel descriptografar o cache MSAL.") from exc

    cache.deserialize(serialized)
    return cache


def acquire_token(cache: msal.SerializableTokenCache) -> str:
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )
    accounts = app.get_accounts()
    if not accounts:
        raise RuntimeError("Nenhuma conta encontrada no cache MSAL.")

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        description = "autenticacao silenciosa indisponivel"
        if result:
            description = result.get("error_description") or result.get("error") or description
        raise RuntimeError(f"Falha ao adquirir token: {description}")

    if cache.has_state_changed:
        serialized = cache.serialize().encode("utf-8")
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        CACHE_FILE.write_bytes(Fernet(normalize_key(CACHE_KEY)).encrypt(serialized))
        try:
            os.chmod(CACHE_FILE, 0o600)
        except OSError:
            pass

    return result["access_token"]


def encode_onedrive_path(path: str) -> str:
    return "/".join(
        quote(part, safe="")
        for part in path.strip("/").split("/")
        if part
    )


def get_excel_metadata(token: str) -> dict[str, Any]:
    encoded = encode_onedrive_path(EXCEL_PATH)
    url = (
        f"{GRAPH_ROOT}/me/drive/root:/{encoded}"
        "?$select=id,name,eTag,cTag,lastModifiedDateTime,size"
    )
    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
        timeout=TIMEOUT,
    )
    if response.status_code != 200:
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise RuntimeError(f"Microsoft Graph retornou {response.status_code}: {detail}")
    return response.json()


def load_previous_state() -> dict[str, Any]:
    if not STATE_FILE.exists():
        return {}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def save_state(metadata: dict[str, Any]) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    state = {
        "id": metadata.get("id", ""),
        "name": metadata.get("name", ""),
        "eTag": metadata.get("eTag", ""),
        "cTag": metadata.get("cTag", ""),
        "lastModifiedDateTime": metadata.get("lastModifiedDateTime", ""),
        "size": metadata.get("size", 0),
    }
    STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> int:
    cache = load_cache()
    token = acquire_token(cache)
    metadata = get_excel_metadata(token)
    previous = load_previous_state()

    current_etag = str(metadata.get("eTag", ""))
    previous_etag = str(previous.get("eTag", ""))
    modified_at = str(metadata.get("lastModifiedDateTime", ""))

    changed = bool(current_etag) and current_etag != previous_etag
    first_check = not previous_etag
    execute = FORCE_FULL or changed or first_check

    if FORCE_FULL:
        reason = "execucao completa manual ou de seguranca"
    elif first_check:
        reason = "primeira verificacao sem estado anterior"
    elif changed:
        reason = "alteracao detectada no Excel"
    else:
        reason = "Excel sem alteracao desde a ultima verificacao"

    etag_hash = hashlib.sha256(current_etag.encode("utf-8")).hexdigest()[:16]

    log(f"Arquivo localizado: {metadata.get('name', '')}")
    log(f"Ultima modificacao: {modified_at}")
    log(f"Resultado: {reason}.")

    github_output("executar", "true" if execute else "false")
    github_output("etag", current_etag.replace("\n", ""))
    github_output("etag-hash", etag_hash)
    github_output("modificado_em", modified_at.replace("\n", ""))

    save_state(metadata)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ERRO FATAL: {exc}", file=sys.stderr, flush=True)
        sys.exit(1)
