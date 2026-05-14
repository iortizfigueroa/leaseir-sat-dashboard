"""
Pull de sustituciones activas de Jira:
- Sub-tasks tipo "Sub-task (máquina sustitución)" con consola/HP prestado y status NO cerrado
- Para cada sub-task, fetch del ticket parent (cf10171 consola averiada, cf10150 HP averiado, status)

Auth: JIRA_BASE_URL, JIRA_USER, JIRA_TOKEN env vars (mismas que fetch_jira.py).

Output: cache/sustis_activas.json con shape:
{
  "_meta": {...},
  "items": [
    {
      "key": "LEAS-6214",
      "subtask_status": "Equipo enviado",
      "cliente": "...",
      "loc": "...",
      "consola_susti": "SUSTMHR033",
      "manipulo_susti": null,
      "fecha_envio": "2026-05-07T14:14:00.000+0200",
      "parent_key": "LEAS-6213",
      "parent_status": "Pendiente recogida",
      "consola_averiada": "...",  // del parent
      "manipulo_averiado": "...", // del parent
      "modelo": "MHR"             // inferido del prefijo del serial
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


def session_from_env():
    base = os.environ["JIRA_BASE_URL"].rstrip("/")
    user = os.environ["JIRA_USER"]
    token = os.environ["JIRA_TOKEN"]
    s = requests.Session()
    s.auth = HTTPBasicAuth(user, token)
    s.headers["Accept"] = "application/json"
    return base, s


JQL_SUSTIS = (
    'type = "Sub-task (máquina sustitución)" '
    'AND status NOT IN (Closed, Finalizada, Cancelado) '
    'ORDER BY cf[10549] DESC'
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


def fetch_issue(base, sess, key, fields):
    r = sess.get(f"{base}/rest/api/3/issue/{key}",
                 params={"fields": ",".join(fields)}, timeout=30)
    r.raise_for_status()
    return r.json()


def opt(v):
    if isinstance(v, dict):
        return v.get("value")
    return v


def infer_modelo(consola, manipulo):
    """Infiere el modelo a partir del prefijo del serial prestado."""
    for s in [consola or "", manipulo or ""]:
        s = s.strip().upper()
        if s.startswith("SUSTMHR"): return "MHR"
        if s.startswith("SUSTDUAL"): return "MHR Dual"
        if s.startswith("SUSTQUAD"): return "MHR Quad"
        if s.startswith("SUSTHR"): return "HR"
        if s.startswith("C") and len(s) > 1 and s[1:].isdigit():
            return "XCell HR"
        if s.startswith("H") and len(s) > 1 and s[1:].isdigit():
            return "XCell HR"
        if s.startswith("MHP"): return "MHR"
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", required=True)
    args = ap.parse_args()

    base, sess = session_from_env()
    print(f"[sustis] JQL: {JQL_SUSTIS}")

    # 1) pull sub-tasks
    sub_fields = ["summary", "status",
                  "customfield_10199", "customfield_10200",
                  "customfield_10211", "customfield_10138", "customfield_10549",
                  "parent"]
    subs = search_all(base, sess, JQL_SUSTIS, sub_fields)
    print(f"[sustis] {len(subs)} sub-tasks activas encontradas")

    # 2) extraer parents únicos
    parent_keys = sorted({i.get("fields", {}).get("parent", {}).get("key")
                          for i in subs
                          if i.get("fields", {}).get("parent", {}).get("key")})
    print(f"[sustis] {len(parent_keys)} parents únicos a fetch")

    # 3) fetch parents (uno a uno — son pocos)
    parent_data = {}
    for k in parent_keys:
        try:
            issue = fetch_issue(base, sess, k,
                                ["summary", "status",
                                 "customfield_10171", "customfield_10150",
                                 "customfield_10211", "customfield_10138"])
            parent_data[k] = issue
        except Exception as e:
            print(f"  [warn] fetch parent {k} falló: {e}", file=sys.stderr)
        time.sleep(0.05)

    # 4) construir items
    items = []
    for s in subs:
        f = s.get("fields", {})
        parent = f.get("parent", {}) or {}
        pk = parent.get("key")
        pdata = parent_data.get(pk, {}).get("fields", {})
        consola_susti = (f.get("customfield_10199") or "").strip() if f.get("customfield_10199") else None
        manipulo_susti = (f.get("customfield_10200") or "").strip() if f.get("customfield_10200") else None
        items.append({
            "key": s.get("key"),
            "subtask_status": (f.get("status") or {}).get("name", ""),
            "cliente": (f.get("customfield_10211") or "").strip(),
            "loc": (f.get("customfield_10138") or "").strip(),
            "consola_susti": consola_susti,
            "manipulo_susti": manipulo_susti,
            "fecha_envio": f.get("customfield_10549"),
            "parent_key": pk,
            "parent_status": (pdata.get("status") or parent.get("fields", {}).get("status") or {}).get("name", ""),
            "parent_cliente": (pdata.get("customfield_10211") or "").strip(),
            "consola_averiada": (pdata.get("customfield_10171") or "").strip(),
            "manipulo_averiado": (pdata.get("customfield_10150") or "").strip(),
            "modelo": infer_modelo(consola_susti, manipulo_susti),
        })

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
    print(f"[sustis] Wrote {len(items)} items to {args.cache}")

    # Resumen
    from collections import Counter
    by_modelo = Counter(i["modelo"] or "?" for i in items)
    print(f"[sustis] Por modelo:")
    for m, n in by_modelo.most_common():
        print(f"  {m}: {n}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
