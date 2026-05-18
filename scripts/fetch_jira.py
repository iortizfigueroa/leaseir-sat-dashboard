"""
Fetch and cache the status changelog + extended fields for LEAS-Task tickets
that have been "open" at any moment during the lookback window.

Auth: expects JIRA_BASE_URL, JIRA_USER, JIRA_TOKEN env vars.

Usage:
    python fetch_jira.py --cache cache/jira_status_timeline.json --mode seed
    python fetch_jira.py --cache cache/jira_status_timeline.json --mode update
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

OPEN_STATUSES = [
    "Abierto", "Creado", "Devuelto a cliente", "En preparación presupuesto",
    "En cola taller", "En préstamo", "En reparación", "Enviado a técnico externo",
    "Equipo devuelto", "Equipo enviado", "Esperando inicio reparación",
    "Esperando respuesta cliente a presupuesto", "Formulario en curso",
    "Formulario enviado a calidad", "Gestionado transporte", "Inspección de salida",
    "Investigación", "Material enviado", "Pendiente agendar llamada",
    "Pendiente asignar técnico", "Pendiente confirmación presupuesto",
    "Pendiente definir servicio externo", "Pendiente recogida",
    "Presupuesto aceptado", "Presupuesto enviado",
    "Presupuesto preparado pendiente de enviar", "Queja creada",
    "Recepcionado SAT", "Reportado", "Solicitado",
]

# Custom field IDs for LEAS project (see jira_field_mapping.md memory)
FIELDS = [
    "summary", "created", "status", "assignee",
    "customfield_10211",  # Cliente
    "customfield_10138",  # Localización averiado
    "customfield_10140",  # Tipo de avería
    "customfield_10208",  # Bloqueante?
    "customfield_10198",  # Máquina de sustitución
    "customfield_10225",  # Fecha de venta
    "customfield_10171",  # Consola serial
    "customfield_10150",  # Handpiece serial
    "customfield_10182",  # Máquina en garantía
    "customfield_10210",  # Descripción avería
    "customfield_10133",  # Número de disparos entrada
    "customfield_10247",  # Técnico taller
    "customfield_10301",  # Importe presupuesto
    "customfield_10143",  # Nombre técnico externo
    "customfield_10615",  # Resumen avería (multi-select)
    "customfield_10815",  # Cambios, observaciones y mejoras (multi-select)
    "customfield_11354",  # Localización en Google Maps (URL)
    "customfield_10128",  # Forma de resolución (select) → mapea a gestión
    "customfield_10183",  # ¿Tiene contrato de mantenimiento? (select)
    "comment",
]


def map_gestion(forma_resolucion):
    """Mapea el campo 'Forma de resolución' de Jira a la categoría de Gestión:
    None/vacío → 'Inicio', 'Gestión online' → 'Online',
    'Reparación en taller' → 'Interna', 'Técnico externo' → 'Externa'.
    """
    if not forma_resolucion:
        return "Inicio"
    v = str(forma_resolucion).lower()
    if "online" in v: return "Online"
    if "externo" in v or "externa" in v: return "Externa"
    if "taller" in v or "interna" in v: return "Interna"
    return "Inicio"


def _q(s):
    return f'"{s}"' if any(c in s for c in " \t") else s


def jql_open_during(since, until):
    statuses = ", ".join(_q(s) for s in OPEN_STATUSES)
    return (f'project = LEAS AND issuetype = Task AND status was in ({statuses}) '
            f'DURING ("{since}", "{until}")')


def jql_updated_since(since):
    return f'project = LEAS AND issuetype = Task AND updated >= "{since}"'


def session_from_env():
    base = os.environ["JIRA_BASE_URL"].rstrip("/")
    user = os.environ["JIRA_USER"]
    token = os.environ["JIRA_TOKEN"]
    s = requests.Session()
    s.auth = HTTPBasicAuth(user, token)
    s.headers["Accept"] = "application/json"
    return base, s


def search_keys(base, sess, jql):
    keys = []
    next_token = None
    while True:
        body = {"jql": jql, "fields": ["summary"], "maxResults": 100}
        if next_token:
            body["nextPageToken"] = next_token
        r = sess.post(f"{base}/rest/api/3/search/jql", json=body, timeout=30)
        r.raise_for_status()
        data = r.json()
        keys.extend(i["key"] for i in data.get("issues", []))
        if data.get("isLast", True):
            break
        next_token = data.get("nextPageToken")
        if not next_token:
            break
    return keys


def fetch_issue(base, sess, key):
    r = sess.get(
        f"{base}/rest/api/3/issue/{key}",
        params={"expand": "changelog", "fields": ",".join(FIELDS)},
        timeout=30,
    )
    r.raise_for_status()
    return r.json()


def opt(v):
    if isinstance(v, dict):
        return v.get("value")
    return v


def adf_text(body):
    """Crude ADF -> plain text."""
    if not body:
        return ""
    if isinstance(body, str):
        return body
    out = []

    def walk(node):
        if isinstance(node, dict):
            t = node.get("type")
            if t == "text":
                out.append(node.get("text", ""))
            elif t == "hardBreak":
                out.append("\n")
            for c in (node.get("content") or []):
                walk(c)
            if t in ("paragraph", "heading", "listItem"):
                out.append(" ")
    walk(body)
    return re.sub(r"\s+", " ", "".join(out)).strip()


def latest_comment(fields):
    c = fields.get("comment")
    if not c or not isinstance(c, dict):
        return ""
    comments = c.get("comments") or []
    if not comments:
        return ""
    latest = max(comments, key=lambda x: x.get("updated") or x.get("created") or "")
    text = adf_text(latest.get("body"))
    author = (latest.get("author") or {}).get("displayName", "")
    when = (latest.get("updated") or latest.get("created") or "")[:10]
    head = f"{when} {author}".strip()
    return (head + " — " + text) if (head and text) else (text or head)


def extract(issue):
    key = issue["key"]
    f = issue["fields"]
    transitions = []
    for h in sorted((issue.get("changelog") or {}).get("histories", []),
                    key=lambda h: h.get("created", "")):
        for item in h.get("items", []):
            if item.get("field") == "status":
                transitions.append([h.get("created"),
                                    item.get("fromString") or "",
                                    item.get("toString") or ""])
    importe = f.get("customfield_10301")
    return {
        "key": key,
        "summary": f.get("summary", ""),
        "cliente": (f.get("customfield_10211") or "").strip(),
        "loc": (f.get("customfield_10138") or "").strip(),
        "created": f.get("created"),
        "current_status": (f.get("status") or {}).get("name", ""),
        "tipo": opt(f.get("customfield_10140")) or "",
        "bloq": opt(f.get("customfield_10208")) or "",
        "susti": opt(f.get("customfield_10198")) or "",
        "fventa": f.get("customfield_10225") or "",
        "consola": (f.get("customfield_10171") or "").strip(),
        "hp": (f.get("customfield_10150") or "").strip(),
        "garantia": opt(f.get("customfield_10182")) or "",
        "descripcion": (f.get("customfield_10210") or "").strip(),
        "disparos": f.get("customfield_10133") if f.get("customfield_10133") is not None else "",
        "tec_taller": (f.get("customfield_10247") or {}).get("displayName", "") if isinstance(f.get("customfield_10247"), dict) else (f.get("customfield_10247") or ""),
        "asignado": (f.get("assignee") or {}).get("displayName", "") if isinstance(f.get("assignee"), dict) else "",
        "importe": importe if importe is not None else "",
        "tec_externo": opt(f.get("customfield_10143")) or "",
        "motivo": [opt(x) for x in (f.get("customfield_10615") or []) if opt(x)],
        "cambios": [opt(x) for x in (f.get("customfield_10815") or []) if opt(x)],
        "gmap_url": (f.get("customfield_11354") or "").strip(),
        "forma_resolucion": opt(f.get("customfield_10128")) or "",
        "gestion": map_gestion(opt(f.get("customfield_10128"))),
        "mantenimiento": opt(f.get("customfield_10183")) or "",
        "ult_comentario": latest_comment(f),
        "transitions": transitions,
        "fetched_at": datetime.now(timezone.utc).isoformat(),
    }


def load_cache(path):
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {"_meta": {"last_fetched_at": None, "schema_version": 2}, "tickets": {}}


def save_cache(path, cache):
    cache.setdefault("_meta", {})
    cache["_meta"]["last_fetched_at"] = datetime.now(timezone.utc).isoformat()
    cache["_meta"]["schema_version"] = 2
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", required=True)
    ap.add_argument("--since", default="2025-10-30")
    ap.add_argument("--until", default=None)
    ap.add_argument("--mode", choices=("seed", "update"), default="update")
    args = ap.parse_args()

    base, sess = session_from_env()
    cache = load_cache(Path(args.cache))

    if args.mode == "seed":
        until = args.until or datetime.now(timezone.utc).strftime("%Y-%m-%d")
        jql = jql_open_during(args.since, until)
        print(f"[seed] {jql}")
        keys = search_keys(base, sess, jql)
    else:
        last = (cache.get("_meta") or {}).get("last_fetched_at")
        if not last:
            print("No previous cache - switching to seed mode.", file=sys.stderr)
            until = args.until or datetime.now(timezone.utc).strftime("%Y-%m-%d")
            keys = search_keys(base, sess, jql_open_during(args.since, until))
        else:
            since_date = last[:10]
            jql = jql_updated_since(since_date)
            print(f"[update] {jql}")
            keys = search_keys(base, sess, jql)
            # Además, refresca todos los tickets que en cache están con estado abierto.
            # Esto garantiza assignee y otros campos siempre frescos para los tickets abiertos
            # aunque Jira no los haya marcado como "updated" recientemente.
            open_set = set(OPEN_STATUSES)
            open_in_cache = [k for k, t in (cache.get("tickets") or {}).items()
                             if t.get("current_status") in open_set]
            extra = [k for k in open_in_cache if k not in set(keys)]
            print(f"[update] {len(extra)} tickets abiertos extra a refrescar (assignee+campos)")
            keys = list(keys) + extra
            # Backfill: refresca tickets que no tienen campos nuevos (gestion / forma_resolucion /
            # mantenimiento / gmap_url). Pasa una sola vez tras añadir un campo nuevo al schema.
            already = set(keys)
            backfill = [k for k, t in (cache.get("tickets") or {}).items()
                        if k not in already and (
                            "gestion" not in t or
                            "forma_resolucion" not in t or
                            "mantenimiento" not in t
                        )]
            if backfill:
                print(f"[update] backfill: {len(backfill)} tickets sin campos nuevos a refrescar")
                keys = list(keys) + backfill

    print(f"Fetching {len(keys)} tickets...")
    for i, k in enumerate(keys, 1):
        issue = fetch_issue(base, sess, k)
        cache["tickets"][k] = extract(issue)
        if i % 20 == 0:
            print(f"  {i}/{len(keys)}")
        time.sleep(0.05)

    save_cache(Path(args.cache), cache)
    print(f"Saved cache with {len(cache['tickets'])} tickets to {args.cache}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
