"""geocode.py — Resuelve direcciones de tickets y sustis a lat/lon usando
Nominatim (OpenStreetMap) y cachea resultados en cache/geocodes.json.

Idempotente: solo geocodifica las direcciones nuevas o las que fallaron.
Rate limit 1 req/segundo (política Nominatim).

Estrategias para mejorar matches:
  1. Restringir a Europa (countrycodes).
  2. Limpiar nombres comerciales conocidos (Sin Vello, Elha, etc.) y extraer
     la parte que parezca ciudad/dirección.
  3. Reintentar fallidas si el query cambia respecto a la última vez.

Uso:
    python scripts/geocode.py            # geocodifica pendientes + reintenta fallidas con query nuevo
    python scripts/geocode.py --max 25   # solo 25 (para no agotar timeout)
    python scripts/geocode.py --reset-failed  # vacía las fallidas para reintentar todas
"""
from __future__ import annotations
import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote_plus
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

ROOT = Path(__file__).resolve().parent.parent
CACHE_DIR = ROOT / "cache"
GEO_PATH = CACHE_DIR / "geocodes.json"

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
USER_AGENT = "leaseir-sat-dashboard/1.0 (https://github.com/iortizfigueroa/leaseir-sat-dashboard)"

# Países donde operan los clientes de Leaseir
COUNTRY_CODES = "es,it,pt,fr,de,uk,gb,be,nl,ch,at,ie,lu,dk,se,no,fi,pl,cz"

OPEN_STATUSES = {
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
}

# Nombres comerciales que NO sirven como dirección (a quitar al principio)
COMMERCIAL_PREFIXES = [
    "sin vello", "sinvello",
    "elha",
    "epil point", "epilpoint", "epil",
    "dermasana",
    "smart duck", "smartduck",
    "laser factory", "laserfactory",
    "centri unico", "centro unico", "centros unico",
    "beauty cool",
    "haiku larios", "haiku",
    "ces depilacion", "ces depilación", "ces",
    "mel clinics",
    "estètica àngels", "estetica angels",
    "samrt duck", "samrt",  # typos
    "quora",
    "cleanbody",
    "depilacio laser profesional",
    "depilación laser profesional",
]


def load_cache():
    if GEO_PATH.exists():
        try:
            return json.loads(GEO_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"_meta": {"version": 1, "updated_at": None}, "entries": {}}


def save_cache(cache):
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache["_meta"]["updated_at"] = datetime.now(timezone.utc).isoformat()
    GEO_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def clean_address(s):
    if not s:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def addr_key(cliente, loc):
    loc = clean_address(loc)
    cliente = clean_address(cliente)
    if loc and len(loc) > 8:
        return loc.lower()
    if cliente:
        return cliente.lower()
    return ""


def build_query(cliente, loc):
    """Genera la mejor query para Nominatim. Si la loc tiene calle/CP, usarla
    tal cual. Si no, intentar extraer ciudad del cliente quitando el nombre
    comercial conocido."""
    loc = clean_address(loc)
    cliente = clean_address(cliente)

    # Si tenemos loc con pinta de dirección postal (largo + tiene número o CP)
    if loc and len(loc) > 10 and (re.search(r"\d", loc) or "," in loc):
        return loc, "loc-direct"

    # Si loc es solo una ciudad (corta sin número)
    if loc and len(loc) >= 4 and not re.search(r"\d{4,}", loc):
        # Probablemente solo el nombre de ciudad
        return loc, "loc-city"

    # Solo cliente — limpiar nombre comercial
    base = cliente.lower()
    for prefix in COMMERCIAL_PREFIXES:
        if base.startswith(prefix):
            rest = cliente[len(prefix):].strip(" -,")
            if rest:
                # rest probablemente es ciudad. Añadir España como default.
                return rest, "cliente-clean"
    # Quitar códigos tipo MHP40546, C03291 al final
    cleaned = re.sub(r"[CHMP]\d{4,}", "", cliente).strip(" -,")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or cliente, "cliente-raw"


def nominatim_search(query, retries=2):
    if not query:
        return None
    url = (f"{NOMINATIM_URL}?q={quote_plus(query)}"
           f"&format=json&limit=1&addressdetails=0"
           f"&countrycodes={COUNTRY_CODES}")
    req = Request(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "es,en,it,pt,fr"})
    for attempt in range(retries + 1):
        try:
            with urlopen(req, timeout=15) as r:
                data = json.loads(r.read().decode("utf-8"))
            if data:
                d = data[0]
                return {
                    "lat": float(d["lat"]),
                    "lon": float(d["lon"]),
                    "display_name": d.get("display_name", ""),
                }
            return None
        except (URLError, HTTPError, json.JSONDecodeError, ValueError) as e:
            if attempt < retries:
                time.sleep(2)
                continue
            print(f"  [warn] Nominatim falló para {query!r}: {e}", file=sys.stderr)
            return None


def collect_addresses():
    addresses = {}
    tickets_cache = CACHE_DIR / "jira_status_timeline.json"
    if tickets_cache.exists():
        cache = json.loads(tickets_cache.read_text(encoding="utf-8"))
        for k, t in cache.get("tickets", {}).items():
            st = t.get("current_status", "")
            if st not in OPEN_STATUSES:
                continue
            cliente = t.get("cliente", "")
            loc = t.get("loc", "")
            key = addr_key(cliente, loc)
            if not key:
                continue
            if key not in addresses:
                q, strat = build_query(cliente, loc)
                addresses[key] = {"query": q, "strategy": strat, "cliente": cliente, "loc": loc}
    sustis_cache = CACHE_DIR / "sustis_activas.json"
    if sustis_cache.exists():
        s = json.loads(sustis_cache.read_text(encoding="utf-8"))
        for it in s.get("items", []) or []:
            cliente = it.get("cliente", "") or it.get("parent_cliente", "")
            loc = it.get("loc", "")
            key = addr_key(cliente, loc)
            if not key:
                continue
            if key not in addresses:
                q, strat = build_query(cliente, loc)
                addresses[key] = {"query": q, "strategy": strat, "cliente": cliente, "loc": loc}
    return addresses


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--max", type=int, default=0, help="Max nuevas a geocodificar")
    ap.add_argument("--reset-failed", action="store_true", help="Vaciar entradas fallidas para reintentar")
    ap.add_argument("--rebuild-bad", action="store_true", help="Reintentar las que no están en países OK (lat lejos de Europa)")
    args = ap.parse_args()

    print("[geocode] Cargando cache de geocodes...")
    cache = load_cache()
    entries = cache["entries"]

    if args.reset_failed:
        before = len(entries)
        entries = {k: v for k, v in entries.items() if not v.get("failed")}
        cache["entries"] = entries
        save_cache(cache)
        print(f"[geocode] reset-failed: {before - len(entries)} entradas borradas")

    if args.rebuild_bad:
        # Quitar entries cuyos lat/lon caen fuera de Europa (lat 25-72, lon -25 to 45)
        before = len(entries)
        bad_keys = []
        for k, v in entries.items():
            lat = v.get("lat"); lon = v.get("lon")
            if lat is None or lon is None: continue
            if not (25 <= lat <= 72 and -25 <= lon <= 45):
                bad_keys.append(k)
        for k in bad_keys:
            del entries[k]
        cache["entries"] = entries
        save_cache(cache)
        print(f"[geocode] rebuild-bad: {len(bad_keys)} entradas fuera de Europa borradas")

    print(f"[geocode] {len(entries)} direcciones ya cacheadas (OK + failed)")

    print("[geocode] Recopilando direcciones únicas de tickets y sustis...")
    addresses = collect_addresses()
    print(f"[geocode] {len(addresses)} direcciones únicas detectadas")

    # Pendientes = no en cache OR (failed AND query distinta)
    pending = []
    for k, v in addresses.items():
        if k not in entries:
            pending.append((k, v))
        else:
            cur = entries[k]
            if cur.get("failed") and cur.get("query") != v["query"]:
                pending.append((k, v))
    if args.max > 0:
        pending = pending[:args.max]
    print(f"[geocode] {len(pending)} nuevas/reintentos a geocodificar")

    n_new, n_failed = 0, 0
    for i, (key, info) in enumerate(pending, 1):
        q = info["query"]
        print(f"[geocode] ({i}/{len(pending)}) [{info['strategy']}] {q[:60]}...", flush=True)
        res = nominatim_search(q)
        now_iso = datetime.now(timezone.utc).isoformat()
        if res:
            entries[key] = {
                "lat": res["lat"], "lon": res["lon"],
                "display_name": res["display_name"],
                "query": q, "strategy": info["strategy"],
                "cliente": info["cliente"], "loc": info["loc"],
                "geocoded_at": now_iso,
            }
            n_new += 1
        else:
            entries[key] = {
                "lat": None, "lon": None, "display_name": None,
                "query": q, "strategy": info["strategy"],
                "cliente": info["cliente"], "loc": info["loc"],
                "geocoded_at": now_iso, "failed": True,
            }
            n_failed += 1
        if i % 10 == 0:
            save_cache(cache)
        time.sleep(1.1)

    save_cache(cache)
    total = len(entries)
    geo_ok = sum(1 for v in entries.values() if v.get("lat") is not None)
    print(f"[geocode] Done. Total cache: {total} · OK: {geo_ok} · Failed (run): {n_failed} · Nuevos OK: {n_new}")


if __name__ == "__main__":
    main()
