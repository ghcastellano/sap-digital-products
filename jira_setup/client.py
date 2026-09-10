"""Thin wrapper over Jira Cloud REST API · Basic auth · idempotent helpers."""
import json
import time
from base64 import b64encode
from pathlib import Path
from urllib import request
from urllib.error import HTTPError, URLError

from config import (
    CHECKPOINT_DIR,
    JIRA_API_TOKEN,
    JIRA_BASE_URL,
    JIRA_EMAIL,
)

# Auth header (Basic email:token)
_token = b64encode(f"{JIRA_EMAIL}:{JIRA_API_TOKEN}".encode()).decode()
HEADERS = {
    "Authorization": f"Basic {_token}",
    "Accept": "application/json",
    "Content-Type": "application/json",
}


def _req(method: str, path: str, body=None, query=None, raw=False):
    """Perform a JSON HTTP request. Path is appended to JIRA_BASE_URL."""
    url = f"{JIRA_BASE_URL}{path}"
    if query:
        from urllib.parse import urlencode
        url += "?" + urlencode(query)
    data = json.dumps(body).encode() if body is not None else None
    req = request.Request(url, data=data, method=method, headers=HEADERS)
    try:
        with request.urlopen(req, timeout=30) as resp:
            text = resp.read().decode()
            if raw:
                return resp.status, text
            return resp.status, (json.loads(text) if text else None)
    except HTTPError as e:
        body_text = e.read().decode(errors="replace") if e.fp else ""
        try:
            err_json = json.loads(body_text)
        except Exception:
            err_json = {"raw": body_text}
        return e.code, err_json
    except URLError as e:
        return 0, {"error": str(e)}


def get(path, query=None):
    return _req("GET", path, query=query)


def post(path, body=None):
    return _req("POST", path, body=body)


def put(path, body=None):
    return _req("PUT", path, body=body)


def delete(path):
    return _req("DELETE", path)


# === Checkpointing for idempotency ===

def checkpoint_save(name: str, payload: dict):
    """Persist a checkpoint after a successful operation."""
    p = CHECKPOINT_DIR / f"{name}.json"
    p.write_text(json.dumps(payload, indent=2, ensure_ascii=False))


def checkpoint_load(name: str) -> dict | None:
    p = CHECKPOINT_DIR / f"{name}.json"
    if p.exists():
        return json.loads(p.read_text())
    return None


# === Helpers ===

def whoami():
    """Return user info via /rest/api/3/myself · raises if auth fails."""
    code, body = get("/rest/api/3/myself")
    return code, body


def find_field_by_name(name: str):
    """Search custom fields by display name. Returns dict or None."""
    code, body = get("/rest/api/3/field")
    if code >= 300 or not isinstance(body, list):
        return None
    for f in body:
        if f.get("name", "").lower() == name.lower():
            return f
    return None


def find_issue_type_by_name(name: str):
    code, body = get("/rest/api/3/issuetype")
    if code >= 300 or not isinstance(body, list):
        return None
    for it in body:
        if it.get("name", "").lower() == name.lower():
            return it
    return None


def find_project_by_key(key: str):
    code, body = get(f"/rest/api/3/project/{key}")
    if code == 200:
        return body
    return None


def safe_print(label: str, code: int, body):
    """Pretty-print API response."""
    if code < 300:
        print(f"  ✓ {label}  ·  {code}")
    else:
        msg = body if isinstance(body, str) else json.dumps(body, ensure_ascii=False)[:300]
        print(f"  ✗ {label}  ·  {code}  ·  {msg}")
