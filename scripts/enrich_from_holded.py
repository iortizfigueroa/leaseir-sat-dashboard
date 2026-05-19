"""enrich_from_holded.py — completa el campo `presupuesto` de tickets Jira con
datos de Holded.

Para cada ticket que NO tiene presupuesto en attachments (campo `presupuesto`
vacío en cache), busca en los estimates de Holded uno cuya descripción contenga
"LEAS-XXXX". Si encuentra, actualiza el campo `presupuesto` con el docNumber.

Cobertura típica: gana 1-5 tickets adicionales sobre los ~99% que ya cubre el
método de attachments (caso de técnico que subió a Holded pero no al ticket).

Requiere variable de entorno HOLDED_API_KEY. Si no está, el script sale sin
hacer nada (no fail, no es bloqueante).

Uso: python scripts/enrich_from_holded.py --cache cache/jira_status_timeline.json
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError


HOLDED_BASE = "https://api.holded.com/api/invoicing/v1/documents/estimate"
LEAS_PAT = re.compile(r"LEAS-(\d+)", re.IGNORECASE)
PPTO_CODE_PAT = re.compile(r"E2[0-9]\d{4}", re.IGNORECASE)


def fetch_all_estimates(api_key, timeout=30):
    """Devuelve una lista con todos los estimates de Holded (paginado)."""
    out = []
    for page in range(1, 11):  # límite de seguridad: 10 páginas * 500 = 5000
        url = (f"{HOLDED_BASE}?starttmp=1700000000&endtmp=2000000000&page={page}")
        req = Request(url, headers={"key": api_key, "Accept": "application/json"})
        try:
            with urlopen(req, timeout=timeout) as r:
                data = json.loads(r.read().decode("utf-8"))
        except (URLError, HTTPError, ValueError) as e:
            print(f"  [warn] page {page} falló: {e}", file=sys.stderr)
            break
        if not isinstance(data, list) or not data:
            break
        out.extend(data)
        if len(data) < 500:
            break
    return out


def build_leas_to_ppto(estimates):
    """Devuelve {LEAS-XXXX: 'E2YYYYY'} indexando estimates por mención de LEAS en su desc."""
    by_leas = {}
    for e in estimates:
        desc = str(e.get("desc", ""))
        doc = e.get("docNumber") or ""
        if not doc or not PPTO_CODE_PAT.match(doc):
            continue
        for m in LEAS_PAT.findall(desc):
            leas = f"LEAS-{m}"
            # Mantener el primer match (los más recientes salen primero en la API)
            if leas not in by_leas:
                by_leas[leas] = doc.upper()
    return by_leas


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", required=True)
    args = ap.parse_args()

    api_key = os.environ.get("HOLDED_API_KEY", "").strip()
    if not api_key:
        print("[enrich-holded] HOLDED_API_KEY no definida, skip enrichment.")
        return 0

    cache_path = Path(args.cache)
    if not cache_path.exists():
        print(f"[enrich-holded] cache {cache_path} no existe, skip.")
        return 0

    cache = json.loads(cache_path.read_text(encoding="utf-8"))
    tickets = cache.get("tickets", {}) or {}

    # Tickets sin presupuesto detectado por attachments
    sin_ppto = [k for k, t in tickets.items() if not (t.get("presupuesto") or "").strip()]
    if not sin_ppto:
        print("[enrich-holded] no hay tickets sin presupuesto, skip.")
        return 0

    print(f"[enrich-holded] {len(sin_ppto)} tickets sin presupuesto. Consultando Holded...")
    estimates = fetch_all_estimates(api_key)
    print(f"[enrich-holded] {len(estimates)} estimates descargados.")
    by_leas = build_leas_to_ppto(estimates)
    print(f"[enrich-holded] {len(by_leas)} LEAS mapeados a presupuesto en Holded.")

    enriched = 0
    for k in sin_ppto:
        ppto = by_leas.get(k)
        if ppto:
            tickets[k]["presupuesto"] = ppto
            enriched += 1

    if enriched:
        cache["tickets"] = tickets
        cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[enrich-holded] {enriched} tickets enriquecidos con presupuesto de Holded.")
    else:
        print("[enrich-holded] ningún ticket coincidía con un estimate Holded.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
