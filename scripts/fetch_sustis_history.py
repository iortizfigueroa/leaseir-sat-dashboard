"""
Pull HISTORICO de sustituciones de Jira (2026): activas + cerradas.

Misma idea que fetch_sustis.py pero sin filtrar por estado: trae TODAS las
sub-tareas de sustitución creadas o cerradas en 2026. Captura el changelog
del estado para poder calcular el tiempo total que duró la sustitución
(fecha_envio → fecha_devolucion).

Auth: JIRA_BASE_URL, JIRA_USER, JIRA_TOKEN env vars.

Output: cache/sustis_historico.json con shape:
{
  "_meta": {...},
  "items": [
    {
      "key": "LEAS-XXXX",
      "subtask_status": "Devuelto",     # estado final
      "cliente": "...",
      "loc": "...",
      "consola_susti": "...",
      "manipulo_susti": null,
      "fecha_envio": "2026-03-01T...",
      "fecha_devolucion": "2026-04-15T...",  # primera transición a estado terminal
      "dias": 45,                              # diferencia entre envio y devolucion
      "parent_key": "LEAS-...",
      "parent_status": "...",
      "is_active": false,                      # True si la sustitucion sigue activa
      "modelo": "MHR"
    }, ...
  ]
}
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from requests.auth import HTTPBasicAuth


# Estados que indican que la sustitución sigue activa (no devuelta)
ACTIVE_STATES = {"Solicitado", "Equipo enviado", "En préstamo", "En prestamo"}

# Estados terminales (devolución/cierre). Importante incluir variantes EXACTAS de Jira:
# - "Equipo devuelto" es el nombre real cuando la máquina vuelve al SAT.
# - "Finalizada"/"Cancelado"/"Closed" son los terminales globales.
TERMINAL_STATES = {
    "Equipo devuelto", "Devuelto",
    "Cerrado", "Finalizada", "Cancelado", "Closed",
    "Finalizado técnico externo", "Resuelto",
}


def session_from_env():
    base = os.environ["JIRA_BASE_URL"].rstrip("/")
    user = os.environ["JIRA_USER"]
    token = os.environ["JIRA_TOKEN"]
    s = requests.Session()
    s.auth = HTTPBasicAuth(user, token)
    s.headers["Accept"] = "application/json"
    return base, s


# Sub-tareas creadas en 2026 o cerradas en 2026 (cualquier estado)
JQL_SUSTIS_HIST = (
    'type = "Sub-task (máquina sustitución)" '
    'AND (created >= "2026-01-01" OR resolutiondate >= "2026-01-01" OR status in (Solicitado, "Equipo enviado", "En préstamo")) '
    'ORDER BY created DESC'
)


def search_all(base, sess, jql, fields, max_results=100):
    issues = []
    next_token = None
    while True:
        body = {"jql": jql, "fields": fields, "maxResults": max_results}
        if next_token:
            body["nextPageToken"] = next_token
        r = sess.post(f"{base}/rest/api/3/search/jql", json=body, timeout=30)
        r.raise_for_status()
        data = r.json()
        issues.extend(data.get("issues", []))
        if data.get("isLast", True):
            break
        next_token = data.get("nextPageToken")
        if not next_token:
            break
    return issues


def fetch_issue_with_changelog(base, sess, key, fields):
    """Fetch ticket con changelog para extraer transitions de estado."""
    r = sess.get(f"{base}/rest/api/3/issue/{key}",
                 params={"expand": "changelog", "fields": ",".join(fields)},
                 timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_issue(base, sess, key, fields):
    r = sess.get(f"{base}/rest/api/3/issue/{key}",
                 params={"fields": ",".join(fields)}, timeout=30)
    r.raise_for_status()
    return r.json()


def opt(v):
    if isinstance(v, dict):
        return v.get("value")
    return v


def parse_iso(s):
    if not s:
        return None
    s2 = str(s).replace("Z", "+00:00")
    s2 = re.sub(r"([+\-]\d{2})(\d{2})$", r"\1:\2", s2)
    try:
        return datetime.fromisoformat(s2)
    except ValueError:
        return None


def extract_fecha_devolucion(issue):
    """Recorre el changelog del sub-task y devuelve la primera transición a un
    estado terminal (Devuelto/Cerrado/Finalizada/Cancelado)."""
    histories = (issue.get("changelog") or {}).get("histories", [])
    histories_sorted = sorted(histories, key=lambda h: h.get("created", ""))
    for h in histories_sorted:
        for item in h.get("items", []) or []:
            if item.get("field") == "status":
                to_state = item.get("toString") or ""
                if to_state in TERMINAL_STATES:
                    return parse_iso(h.get("created"))
    return None


def infer_modelo(consola, manipulo):
    for s in [consola or "", manipulo or ""]:
        s = s.strip().upper()
        if s.startswith("SUSTMHR"): return "MHR"
        if s.startswith("SUSTDUAL"): return "MHR Dual"
        if s.startswith("SUSTQUAD"): return "MHR Quad"
        if s.startswith("SUSTHR"): return "HR"
        if s.startswith("C") and len(s) > 1 and s[1:].isdigit(): return "XCell HR"
        if s.startswith("H") and len(s) > 1 and s[1:].isdigit(): return "XCell HR"
        if s.startswith("MHP"): return "MHR"
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", required=True)
    args = ap.parse_args()

    base, sess = session_from_env()
    print(f"[sustis-hist] JQL: {JQL_SUSTIS_HIST}")

    sub_fields = ["summary", "status", "resolutiondate",
                  "customfield_10199", "customfield_10200",
                  "customfield_10211", "customfield_10138", "customfield_10549",
                  "parent", "created"]
    subs = search_all(base, sess, JQL_SUSTIS_HIST, sub_fields)
    print(f"[sustis-hist] {len(subs)} sub-tasks (activas + cerradas 2026)")

    # Fetch changelog de cada sub-task para sacar fecha_devolucion
    items = []
    parent_keys_to_fetch = set()
    for i, s in enumerate(subs):
        key = s.get("key")
        try:
            full = fetch_issue_with_changelog(base, sess, key, sub_fields)
        except Exception as e:
            print(f"  [warn] fetch {key} con changelog falló: {e}", file=sys.stderr)
            full = s
        time.sleep(0.05)
        f = full.get("fields", {})
        parent = f.get("parent", {}) or {}
        pk = parent.get("key")
        if pk:
            parent_keys_to_fetch.add(pk)
        consola_susti = (f.get("customfield_10199") or "").strip() if f.get("customfield_10199") else None
        manipulo_susti = (f.get("customfield_10200") or "").strip() if f.get("customfield_10200") else None
        cur_status = (f.get("status") or {}).get("name", "")
        is_active = cur_status not in TERMINAL_STATES
        fecha_envio = parse_iso(f.get("customfield_10549"))
        fecha_dev = extract_fecha_devolucion(full)
        dias = None
        if fecha_envio:
            end = fecha_dev or datetime.now(timezone.utc)
            if not fecha_envio.tzinfo: fecha_envio = fecha_envio.replace(tzinfo=timezone.utc)
            if not end.tzinfo: end = end.replace(tzinfo=timezone.utc)
            dias = (end - fecha_envio).days
        items.append({
            "key": key,
            "subtask_status": cur_status,
            "cliente": (f.get("customfield_10211") or "").strip(),
            "loc": (f.get("customfield_10138") or "").strip(),
            "consola_susti": consola_susti,
            "manipulo_susti": manipulo_susti,
            "fecha_envio": fecha_envio.isoformat() if fecha_envio else None,
            "fecha_devolucion": fecha_dev.isoformat() if fecha_dev else None,
            "dias": dias if dias is not None and dias >= 0 else None,
            "is_active": is_active,
            "parent_key": pk,
            "parent_status": "",  # se rellena abajo
            "parent_cliente": "",
            "consola_averiada": "",
            "manipulo_averiado": "",
            "modelo": infer_modelo(consola_susti, manipulo_susti),
        })
        if (i + 1) % 20 == 0:
            print(f"  [sustis-hist] {i+1}/{len(subs)} sub-tasks procesadas")

    # Fetch parents (no necesitamos changelog)
    print(f"[sustis-hist] {len(parent_keys_to_fetch)} parents a fetch")
    parent_data = {}
    for k in sorted(parent_keys_to_fetch):
        try:
            issue = fetch_issue(base, sess, k,
                                ["summary", "status",
                                 "customfield_10171", "customfield_10150",
                                 "customfield_10211", "customfield_10138"])
            parent_data[k] = issue
        except Exception as e:
            print(f"  [warn] fetch parent {k} falló: {e}", file=sys.stderr)
        time.sleep(0.05)

    for it in items:
        pk = it["parent_key"]
        if not pk: continue
        pdata = parent_data.get(pk, {}).get("fields", {})
        it["parent_status"] = (pdata.get("status") or {}).get("name", "")
        it["parent_cliente"] = (pdata.get("customfield_10211") or "").strip()
        it["consola_averiada"] = (pdata.get("customfield_10171") or "").strip()
        it["manipulo_averiado"] = (pdata.get("customfield_10150") or "").strip()

    payload = {
        "_meta": {
            "last_fetched_at": datetime.now(timezone.utc).isoformat(),
            "schema_version": 1,
            "count": len(items),
        },
        "items": items,
    }
    p = Path(args.cache)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[sustis-hist] Wrote {len(items)} items to {args.cache}")
    n_active = sum(1 for x in items if x["is_active"])
    n_closed = len(items) - n_active
    n_with_dias = sum(1 for x in items if x["dias"] is not None)
    print(f"[sustis-hist] Activas: {n_active} · Cerradas: {n_closed} · Con días: {n_with_dias}")


if __name__ == "__main__":
    main()
