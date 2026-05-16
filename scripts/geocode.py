"""geocode.py — Resuelve direcciones de tickets y sustis a lat/lon con Nominatim.
Idempotente. Rate limit 1 req/s. Limpia ruido (Dirección:, Teléfono:, etc.)."""
from __future__ import annotations
import json, re, sys, time
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

COMMERCIAL_PREFIXES = [
    "sin vello!", "sin vello", "sinvello!", "sinvello",
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
    "samrt duck", "samrt",
    "quora", "cleanbody",
    "depilacio laser profesional", "depilación laser profesional",
    "dglow",
]

# Prefijos basura típicos en la columna Localización
NOISE_PREFIXES = [
    "dirección:", "direccion:", "dirección :", "direccion :",
    "se encuentra en:", "se encuentra en :",
    "ubicacion:", "ubicación:",
]
# Patrones a quitar dentro del texto
NOISE_PATTERNS = [
    r"tel[eé]fono\s*:.*$",
    r"\btlf\.?\s*\d.*$",
    r"\btel\.?\s*\d.*$",
    r",?\s*bajos?\b",
    r",?\s*pasaje\s+interior",
    r",?\s*pasaje\s+interno",
    r",?\s*local\s+l?\d+[a-z\-\s\d]*",
    r",?\s*piso\s+\d+\w*",
    r",?\s*c\.c\.\s+[^,]+",
    r",?\s*entresuelo",
    r",?\s*ent\.\s*\d+[\sªº]*",
    r"\(\s*\d+\s*\)",
]


def load_cache():
    if GEO_PATH.exists():
        try: return json.loads(GEO_PATH.read_text(encoding="utf-8"))
        except: pass
    return {"_meta": {"version": 1, "updated_at": None}, "entries": {}}


def save_cache(cache):
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache["_meta"]["updated_at"] = datetime.now(timezone.utc).isoformat()
    GEO_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def clean_address(s):
    if not s: return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def addr_key(cliente, loc):
    loc, cliente = clean_address(loc), clean_address(cliente)
    if loc and len(loc) > 8: return loc.lower()
    if cliente: return cliente.lower()
    return ""


def aggressive_clean(text):
    """Cleanup agresivo de la string para Nominatim."""
    if not text: return ""
    t = text.strip()
    # Quitar prefijos basura
    low = t.lower()
    for px in NOISE_PREFIXES:
        if low.startswith(px):
            t = t[len(px):].strip()
            low = t.lower()
            break
    # Quitar : inicial
    while t.startswith(":") or t.startswith("-"):
        t = t[1:].strip()
    # Aplicar patrones
    for pat in NOISE_PATTERNS:
        t = re.sub(pat, "", t, flags=re.IGNORECASE)
    # Quitar códigos tipo MHP40546, C03291 al final/medio
    t = re.sub(r"\b[CHMP]\d{4,}\b", "", t)
    # Quitar trailing comas/espacios/guiones
    t = re.sub(r"[,\-\s]+$", "", t)
    # Compactar espacios
    t = re.sub(r"\s+", " ", t).strip()
    return t


def build_query(cliente, loc):
    loc, cliente = clean_address(loc), clean_address(cliente)
    # 1. Si loc parece dirección (largo y con número o coma)
    if loc and len(loc) > 10 and (re.search(r"\d", loc) or "," in loc):
        cleaned = aggressive_clean(loc)
        return cleaned or loc, "loc-direct-clean"
    # 2. Si loc es solo ciudad
    if loc and len(loc) >= 4 and not re.search(r"\d{4,}", loc):
        return aggressive_clean(loc) or loc, "loc-city"
    # 3. Solo cliente, quitar nombre comercial
    base = cliente.lower()
    for prefix in COMMERCIAL_PREFIXES:
        if base.startswith(prefix):
            rest = cliente[len(prefix):].strip(" -,!")
            if rest:
                return aggressive_clean(rest) or rest, "cliente-clean"
    # 4. Fallback
    cleaned = re.sub(r"[CHMP]\d{4,}", "", cliente).strip(" -,")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or cliente, "cliente-raw"


def nominatim_search(query, retries=2):
    if not query: return None
    url = (f"{NOMINATIM_URL}?q={quote_plus(query)}"
           f"&format=json&limit=1&addressdetails=0&countrycodes={COUNTRY_CODES}")
    req = Request(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "es,en,it,pt,fr"})
    for attempt in range(retries + 1):
        try:
            with urlopen(req, timeout=15) as r:
                data = json.loads(r.read().decode("utf-8"))
            if data:
                d = data[0]
                return {"lat": float(d["lat"]), "lon": float(d["lon"]),
                        "display_name": d.get("display_name", "")}
            return None
        except (URLError, HTTPError, json.JSONDecodeError, ValueError) as e:
            if attempt < retries:
                time.sleep(2); continue
            print(f"  [warn] Nominatim falló para {query!r}: {e}", file=sys.stderr)
            return None


def collect_addresses():
    addresses = {}
    for cache_name, key_field in [("jira_status_timeline.json", "tickets"),
                                    ("sustis_activas.json", "items")]:
        path = CACHE_DIR / cache_name
        if not path.exists(): continue
        d = json.loads(path.read_text(encoding="utf-8"))
        if key_field == "tickets":
            for k, t in d.get("tickets", {}).items():
                if t.get("current_status") not in OPEN_STATUSES: continue
                cliente = t.get("cliente", "")
                loc = t.get("loc", "")
                key = addr_key(cliente, loc)
                if key and key not in addresses:
                    q, strat = build_query(cliente, loc)
                    addresses[key] = {"query": q, "strategy": strat, "cliente": cliente, "loc": loc}
        else:
            for it in d.get("items", []) or []:
                cliente = it.get("cliente", "") or it.get("parent_cliente", "")
                loc = it.get("loc", "")
                key = addr_key(cliente, loc)
                if key and key not in addresses:
                    q, strat = build_query(cliente, loc)
                    addresses[key] = {"query": q, "strategy": strat, "cliente": cliente, "loc": loc}
    return addresses


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--max", type=int, default=0)
    ap.add_argument("--reset-failed", action="store_true")
    ap.add_argument("--rebuild-bad", action="store_true")
    args = ap.parse_args()
    print("[geocode] Cargando cache...")
    cache = load_cache()
    entries = cache["entries"]
    if args.reset_failed:
        before = len(entries)
        entries = {k: v for k, v in entries.items() if not v.get("failed")}
        cache["entries"] = entries
        save_cache(cache)
        print(f"[geocode] reset-failed: {before - len(entries)} borradas")
    if args.rebuild_bad:
        bad_keys = []
        for k, v in entries.items():
            lat, lon = v.get("lat"), v.get("lon")
            if lat is None or lon is None: continue
            if not (25 <= lat <= 72 and -25 <= lon <= 45):
                bad_keys.append(k)
        for k in bad_keys: del entries[k]
        cache["entries"] = entries
        save_cache(cache)
        print(f"[geocode] rebuild-bad: {len(bad_keys)} fuera de Europa borradas")
    print(f"[geocode] {len(entries)} cacheadas")
    addresses = collect_addresses()
    print(f"[geocode] {len(addresses)} direcciones únicas detectadas")
    pending = []
    for k, v in addresses.items():
        if k not in entries:
            pending.append((k, v))
        else:
            cur = entries[k]
            if cur.get("failed") and cur.get("query") != v["query"]:
                pending.append((k, v))
    if args.max > 0: pending = pending[:args.max]
    print(f"[geocode] {len(pending)} pendientes")
    n_new, n_failed = 0, 0
    for i, (key, info) in enumerate(pending, 1):
        q = info["query"]
        print(f"[geocode] ({i}/{len(pending)}) [{info['strategy']}] {q[:60]}...", flush=True)
        res = nominatim_search(q)
        now_iso = datetime.now(timezone.utc).isoformat()
        if res:
            entries[key] = {"lat": res["lat"], "lon": res["lon"],
                            "display_name": res["display_name"], "query": q,
                            "strategy": info["strategy"], "cliente": info["cliente"],
                            "loc": info["loc"], "geocoded_at": now_iso}
            n_new += 1
        else:
            entries[key] = {"lat": None, "lon": None, "display_name": None,
                            "query": q, "strategy": info["strategy"],
                            "cliente": info["cliente"], "loc": info["loc"],
                            "geocoded_at": now_iso, "failed": True}
            n_failed += 1
        if i % 10 == 0: save_cache(cache)
        time.sleep(1.1)
    save_cache(cache)
    total = len(entries)
    geo_ok = sum(1 for v in entries.values() if v.get("lat") is not None)
    print(f"[geocode] Done. Total: {total} · OK: {geo_ok} · Failed (run): {n_failed} · Nuevos OK: {n_new}")


if __name__ == "__main__":
    main()
