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


def _extract_city_from_cliente(cliente):
    """Quita el prefijo comercial de un cliente y devuelve la parte ciudad/dirección."""
    base = cliente.lower()
    for prefix in COMMERCIAL_PREFIXES:
        if base.startswith(prefix):
            rest = cliente[len(prefix):].strip(" -,!")
            if rest:
                return rest
    # Quitar códigos tipo C00510 que no aportan
    cleaned = re.sub(r"\b[CHMP]\d{4,}\b", "", cliente).strip(" -,")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned


def build_query(cliente, loc):
    loc, cliente = clean_address(loc), clean_address(cliente)
    # 1. loc parece dirección postal completa (largo + con número o coma)
    if loc and len(loc) > 15 and (re.search(r"\d", loc) or "," in loc):
        return aggressive_clean(loc) or loc, "loc-direct-clean"
    # 2. loc cortito y ambiguo (ej "sebastian", "fuenlabrada"): combinar con
    #    el cliente para evitar matches absurdos en otros países
    if loc and len(loc) < 20 and not re.search(r"\d{4,}", loc):
        # Combinar loc + cliente limpio
        city_from_cliente = _extract_city_from_cliente(cliente)
        if city_from_cliente and city_from_cliente.lower() != loc.lower():
            combined = f"{city_from_cliente}"
            # Si city_from_cliente ya menciona el loc, no duplicar
            if loc.lower() not in city_from_cliente.lower():
                combined = f"{city_from_cliente}, {loc}"
            return aggressive_clean(combined) or combined, "loc-with-cliente"
        return aggressive_clean(loc) or loc, "loc-city"
    # 3. loc largo sin número (raro) → tal cual
    if loc and len(loc) >= 4:
        return aggressive_clean(loc) or loc, "loc-long"
    # 4. Sin loc, solo cliente → quitar nombre comercial
    cleaned = _extract_city_from_cliente(cliente)
    return aggressive_clean(cleaned) or cleaned or cliente, "cliente-clean"


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


def preprocess_gmap_url(url):
    """Maneja casos especiales antes de hacer fetch HTTP:
    - google.com/url?...&url=DESTINO  → extrae el parámetro url y lo decodifica
    Devuelve el URL listo para resolver (puede ser relativo, en cuyo caso se prefija)."""
    if not url: return ""
    s = str(url).strip()
    if "google." in s and "/url?" in s:
        try:
            from urllib.parse import urlparse, parse_qs, unquote
            parsed = urlparse(s)
            qs = parse_qs(parsed.query)
            target = qs.get("url", [""])[0]
            if target:
                s = unquote(target)
                if s.startswith("/"):
                    s = "https://www.google.com" + s
        except Exception:
            pass
    return s


def http_resolve(url, timeout=10):
    """Hace GET siguiendo redirects y devuelve el URL final. None si falla."""
    if not url: return None
    s = str(url).strip()
    try:
        req = Request(s, headers={"User-Agent": "Mozilla/5.0 " + USER_AGENT,
                                    "Accept-Language": "en,es;q=0.9"})
        with urlopen(req, timeout=timeout) as r:
            return r.geturl()
    except (URLError, HTTPError, ValueError) as e:
        print(f"  [warn] http_resolve({s[:60]}): {e}", file=sys.stderr)
        return None


def http_fetch_html(url, timeout=10, max_bytes=200_000):
    """Descarga el HTML (limitado a max_bytes) de un URL. Devuelve string o ""."""
    if not url: return ""
    try:
        req = Request(str(url).strip(),
                      headers={"User-Agent": "Mozilla/5.0 " + USER_AGENT,
                               "Accept-Language": "en,es;q=0.9"})
        with urlopen(req, timeout=timeout) as r:
            data = r.read(max_bytes)
            return data.decode("utf-8", errors="ignore")
    except (URLError, HTTPError, ValueError) as e:
        print(f"  [warn] http_fetch_html({str(url)[:60]}): {e}", file=sys.stderr)
        return ""


# Coords servidas por la pagina de consentimiento de cookies UE (Gijón) — NO son del lugar real.
# Si extraemos eso, devolvemos None para que caiga al fallback Nominatim con la direccion de texto.
_GMAP_CONSENT_COORDS = {(43.50639385, -5.6786944)}

def _extract_latlng_from_html(html):
    """Busca coords en el HTML de Google Maps. Soporta múltiples patrones."""
    if not html: return None
    # Pattern 1: center=lat%2Clon (URL-encoded comma)
    m = re.search(r"center=(-?\d+\.\d+)%2C(-?\d+\.\d+)", html)
    if m:
        try: return float(m.group(1)), float(m.group(2))
        except ValueError: pass
    # Pattern 2: center=lat,lon (raw comma)
    m = re.search(r"center=(-?\d+\.\d+),(-?\d+\.\d+)", html)
    if m:
        try: return float(m.group(1)), float(m.group(2))
        except ValueError: pass
    # Pattern 3: !3d<lat>!4d<lon>
    m = re.search(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)", html)
    if m:
        try: return float(m.group(1)), float(m.group(2))
        except ValueError: pass
    # Pattern 4: @lat,lon
    m = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", html)
    if m:
        try: return float(m.group(1)), float(m.group(2))
        except ValueError: pass
    return None


def resolve_gmap_short_url(url, timeout=10):
    """Resuelve short links de Google Maps (maps.app.goo.gl / goo.gl/maps /
    share.google) siguiendo el HTTP redirect. Devuelve el URL final o el original
    si no es un short link."""
    if not url: return None
    s = str(url).strip()
    is_short = any(d in s for d in ("maps.app.goo.gl", "goo.gl/maps", "share.google"))
    if not is_short: return s
    resolved = http_resolve(s, timeout=timeout)
    return resolved if resolved else None


def _extract_latlng_patterns(s):
    """Aplica los regex de coords sobre el URL ya resuelto.
    PRIORIDAD: !3d/!4d (coords del PLACE/centro) > q=/ll=/destination= > @ (vista de cámara)
    El @ es la posición de la cámara del mapa, NO del centro — solo usarlo como último recurso."""
    # 1. !3d<lat>!4d<lon> (coords reales del place — máxima prioridad)
    m = re.search(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)", s)
    if m:
        try: return float(m.group(1)), float(m.group(2))
        except ValueError: pass
    # 2. q=lat,lon / ll=lat,lon / destination=lat,lon (param explícito)
    m = re.search(r"[?&](?:q|ll|destination)=(-?\d+\.\d+),(-?\d+\.\d+)", s)
    if m:
        try: return float(m.group(1)), float(m.group(2))
        except ValueError: pass
    # 3. @lat,lon — solo si no hay nada mejor (es la posición de cámara, no del centro)
    m = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", s)
    if m:
        try: return float(m.group(1)), float(m.group(2))
        except ValueError: pass
    return None


def _is_consent_dummy(c):
    if not c: return False
    try:
        return (round(float(c[0]), 6), round(float(c[1]), 6)) in {(43.506394, -5.678694)}
    except Exception:
        return False

def extract_latlng_from_gmap_url(url):
    """Extrae (lat, lon) de un URL de Google Maps.
    Maneja:
    - URLs directos con @lat,lon o !3d!4d
    - Short links (maps.app.goo.gl, goo.gl/maps, share.google) → resuelve redirect
    - URLs de redirect de Google (google.com/url?...&url=DESTINO)
    - URLs con Place ID hex (!1s0xHEX) sin coords → fetch para obtener canónica
    Devuelve None si no se puede parsear.
    """
    if not url: return None
    # Paso 1: si es un redirect google.com/url?, extraer el url destino
    s = preprocess_gmap_url(url)
    if not s: return None
    # Paso 2: si es un short link, resolver redirect HTTP
    if any(d in s for d in ("maps.app.goo.gl", "goo.gl/maps", "share.google")):
        resolved = resolve_gmap_short_url(s)
        if not resolved: return None
        s = resolved
    # Paso 3: intentar extraer coords del URL actual
    coords = _extract_latlng_patterns(s)
    if coords: return coords
    # Paso 4: si tiene Place ID hex (!1s0x...) o feature ID (?cid=) pero no coords,
    # hacer fetch para que Google redirija a la URL canónica con @lat,lon o !3d!4d
    if "/maps/place" in s and ("!1s0x" in s or "?cid=" in s or "?ftid=" in s or "data=" in s):
        resolved = http_resolve(s)
        if resolved and resolved != s:
            coords = _extract_latlng_patterns(resolved)
            if coords: return coords
    # Paso 5 (fallback final): descargar HTML y buscar coords en el body
    # Sirve para URLs como /maps?ftid=, share.google con kgmid, búsquedas Google, etc.
    if "google." in s:
        html = http_fetch_html(s)
        coords = _extract_latlng_from_html(html)
        if coords: return coords
        # Intento extra: si el HTML menciona @lat,lon en algún anchor con /place/, fetch ese
        m = re.search(r"https?://[^\"']*google\.com/maps[^\"']*@(-?\d+\.\d+),(-?\d+\.\d+)", html or "")
        if m:
            try: return float(m.group(1)), float(m.group(2))
            except ValueError: pass
    return None


_orig_extract_latlng_from_gmap_url = extract_latlng_from_gmap_url
# Coords dummy adicionales que Google sirve en su HTML de consent/landing:
#  - (37.0625, -95.677068)  → centro geografico USA (consent con gl=us)
#  - (37.09024, -95.712891) → variante centro USA
#  - (43.50639385, -5.6786944) → Gijón (consent default UE) — ya en _GMAP_CONSENT_COORDS
_GMAP_DUMMY_COORDS = {
    (43.506394, -5.678694),
    (37.0625, -95.677068),
    (37.09024, -95.712891),
    (39.8283, -98.5795),
}

def _is_dummy_coord(c):
    if not c: return False
    try:
        key = (round(float(c[0]), 6), round(float(c[1]), 6))
    except Exception:
        return False
    return key in _GMAP_DUMMY_COORDS

def extract_latlng_from_gmap_url(url):  # type: ignore[no-redef]
    c = _orig_extract_latlng_from_gmap_url(url)
    if _is_consent_dummy(c) or _is_dummy_coord(c):
        return None
    return c


def collect_addresses():
    addresses = {}
    # Pre-pass: tickets con gmap_url tienen prioridad — recorrer tickets primero y
    # capturar gmap_url por addr_key (cualquier ticket abierto en ese centro lo usa).
    gmap_by_key = {}  # addr_key -> gmap_url
    for cache_name in ["jira_status_timeline.json"]:
        path = CACHE_DIR / cache_name
        if not path.exists(): continue
        d = json.loads(path.read_text(encoding="utf-8"))
        for k, t in d.get("tickets", {}).items():
            # Permitimos cualquier ticket (no solo abiertos) ya que gmap_url
            # es semánticamente atributo del CENTRO no del ticket
            cliente = t.get("cliente", "")
            loc = t.get("loc", "")
            gmap_url = (t.get("gmap_url") or "").strip()
            if not gmap_url: continue
            # Validación barata (sin fetch HTTP): debe parecer un URL de Google Maps
            if not (gmap_url.startswith("http") and
                    ("google" in gmap_url or "goo.gl" in gmap_url)):
                continue
            key = addr_key(cliente, loc)
            if key and key not in gmap_by_key:
                gmap_by_key[key] = gmap_url

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
                    if key in gmap_by_key:
                        addresses[key] = {"query": gmap_by_key[key], "strategy": "gmap-url",
                                          "cliente": cliente, "loc": loc,
                                          "gmap_url": gmap_by_key[key]}
                    else:
                        q, strat = build_query(cliente, loc)
                        addresses[key] = {"query": q, "strategy": strat,
                                          "cliente": cliente, "loc": loc}
        else:
            for it in d.get("items", []) or []:
                cliente = it.get("cliente", "") or it.get("parent_cliente", "")
                loc = it.get("loc", "")
                key = addr_key(cliente, loc)
                if key and key not in addresses:
                    if key in gmap_by_key:
                        addresses[key] = {"query": gmap_by_key[key], "strategy": "gmap-url",
                                          "cliente": cliente, "loc": loc,
                                          "gmap_url": gmap_by_key[key]}
                    else:
                        q, strat = build_query(cliente, loc)
                        addresses[key] = {"query": q, "strategy": strat,
                                          "cliente": cliente, "loc": loc}
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
        # gmap-url: SIEMPRE refresca para que cambios en Jira se reflejen rápido
        if v.get("strategy") == "gmap-url":
            cur = entries.get(k, {})
            if (cur.get("strategy") != "gmap-url" or
                cur.get("query") != v["query"]):
                pending.append((k, v))
            continue
        if k not in entries:
            pending.append((k, v))
        else:
            cur = entries[k]
            if cur.get("failed") and cur.get("query") != v["query"]:
                pending.append((k, v))
            # Si la entry actual NO es gmap-url pero AHORA hay opciones, ya se
            # gestiona arriba (strategy gmap-url tiene prioridad).
    if args.max > 0: pending = pending[:args.max]
    print(f"[geocode] {len(pending)} pendientes")
    n_new, n_failed, n_gmap = 0, 0, 0
    for i, (key, info) in enumerate(pending, 1):
        q = info["query"]
        now_iso = datetime.now(timezone.utc).isoformat()
        # gmap-url: extraer lat/lon del URL sin llamar Nominatim
        if info.get("strategy") == "gmap-url":
            coords = extract_latlng_from_gmap_url(info.get("gmap_url") or q)
            if coords:
                lat, lon = coords
                entries[key] = {"lat": lat, "lon": lon,
                                "display_name": "Google Maps URL del centro",
                                "query": info["gmap_url"] or q,
                                "strategy": "gmap-url", "cliente": info["cliente"],
                                "loc": info["loc"], "geocoded_at": now_iso}
                n_gmap += 1
                print(f"[geocode] ({i}/{len(pending)}) [gmap-url] {info['cliente'][:30]} → {lat:.4f},{lon:.4f}")
                if i % 10 == 0: save_cache(cache)
                continue
            else:
                print(f"[geocode] ({i}/{len(pending)}) [gmap-url FAILED] {info['cliente'][:30]} — caer a Nominatim", file=sys.stderr)
                # Fallback: usar build_query estándar
                q, _ = build_query(info["cliente"], info["loc"])
        print(f"[geocode] ({i}/{len(pending)}) [{info['strategy']}] {q[:60]}...", flush=True)
        res = nominatim_search(q)
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
    print(f"[geocode] Done. Total: {total} · OK: {geo_ok} · Failed: {n_failed} · Nominatim nuevos: {n_new} · gmap-url: {n_gmap}")


if __name__ == "__main__":
    main()
