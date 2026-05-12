"""
Pull de Airtable Pedidos (Leaseir) → ventas nuevas (devices entregados/enviados)
agregadas por (cadena, mes).

Auth: AIRTABLE_API_KEY env var (PAT con data.records:read en workspace Leaseir).

Output: cache/airtable_pedidos.json con shape:
{
  "_meta": {"last_fetched_at": "...", "schema_version": 1},
  "ventas_by_chain_month": {
    "Elha":      {"2026-01": 5, "2026-02": 3, ...},
    "Sin Vello": {"2026-01": 2, ...},
    ...
  },
  "records": [ {customer, centro, devices, fecha, chain}, ... ]
}
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests

BASE_ID = "app9U5sz7YS8y9Oit"  # Leaseir
TABLE_ID = "tblKDAT56Z9tX9zvd"  # Pedidos

FLD_CUSTOMER = "fld8eiM2bLOjMVdag"
FLD_CENTRO = "fld61lhh4FIi5zG3d"
FLD_DEVICES = "fldZV2MuSwpB9gVQs"
FLD_STATUS = "fldK7fWGQ2sDUYmBb"
FLD_TYPE = "fld6fZunKNDlwIiWc"
FLD_REPORTING = "fld0Pajx1ow8CmFSM"
FLD_FECHA_REP = "fld2it0UXrAxZoCVu"

# Selectors válidos
SEL_ENTREGADO = "selDZR7ZBPnzYWD6I"   # Entregado a Cliente
SEL_EN_VUELO = "sel51wkk1eSQe4Qv3"    # Enviado a Cliente (en vuelo)
SEL_NEW = "selFfAESn1BuH0PHl"         # Sale of new device
SEL_COMPETITOR = "selR2HAmRlYC0OahI"  # Competitors device


def chain_of(customer, centro):
    c = (customer or "").lower()
    l = (centro or "").lower()
    if "elha" in c: return "Elha"
    if "sin vello" in c or "sinvello" in c: return "Sin Vello"
    if "dermasana" in c: return "Dermasana"
    if "smart duck" in c or "smartduck" in c: return "Smart Duck"
    if "epil point" in c or "epilpoint" in c: return "Epil Point"
    if "laser factory" in c or "laserfactory" in c: return "Laser Factory"
    if "beauty cool" in c or "centro unico" in c or "centri unico" in c or "centros unico" in c:
        ital = ("italia", "italy", "roma", "milano", "napoli", "torino", "bologna",
                "firenze", "ostia", "orio", "giugliano", "andria", "castellammare",
                "piacenza", "novara", "aprilia", "euroma")
        if any(t in l for t in ital) or any(t in c for t in ital):
            return "Unico Italia"
        return "Otros"
    return "Otros"


def session_from_env():
    key = os.environ.get("AIRTABLE_API_KEY")
    if not key:
        print("FATAL: AIRTABLE_API_KEY env var no definida", file=sys.stderr)
        sys.exit(2)
    s = requests.Session()
    s.headers["Authorization"] = f"Bearer {key}"
    s.headers["Accept"] = "application/json"
    return s


def fetch_pedidos(sess, year_start):
    """Pull todos los pedidos con Status entregado/en-vuelo, Type new/competitor,
    Reporting=true, Fecha >= year_start."""
    url = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_ID}"
    fields = [FLD_CUSTOMER, FLD_CENTRO, FLD_DEVICES, FLD_FECHA_REP]
    filter_formula = (
        f"AND("
        f"OR({{Status}}='Entregado a Cliente', {{Status}}='Enviado a Cliente (en vuelo)'),"
        f"OR({{Type of request}}='Sale of new device', {{Type of request}}='Competitors device (e.g. Cocoon or Opphalo)'),"
        f"{{Reporting}}=1,"
        f"IS_AFTER({{Fecha Definitiva Reporting}}, '{year_start}')"
        f")"
    )
    params = {"pageSize": 100, "filterByFormula": filter_formula}
    for f in fields:
        params.setdefault("fields[]", [])
        params["fields[]"] = params.get("fields[]", []) + [f]
    records = []
    offset = None
    while True:
        if offset:
            params["offset"] = offset
        r = sess.get(url, params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break
    return records


def aggregate(records):
    ventas = {}
    flat = []
    for rec in records:
        f = rec.get("fields", {})
        customer = f.get("Customer", "") or ""
        centro = f.get("Centro", "") or ""
        devices = f.get("Number of devices", 0) or 0
        fecha = f.get("Fecha Definitiva Reporting", "") or ""
        if not fecha or not devices:
            continue
        month = fecha[:7]  # YYYY-MM
        chain = chain_of(customer, centro)
        ventas.setdefault(chain, {})
        ventas[chain][month] = ventas[chain].get(month, 0) + int(devices)
        flat.append({
            "customer": customer, "centro": centro, "devices": int(devices),
            "fecha": fecha, "chain": chain,
        })
    return ventas, flat


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", required=True)
    ap.add_argument("--year", type=int, default=None)
    args = ap.parse_args()

    today = datetime.now(timezone.utc)
    year = args.year or today.year
    year_start = f"{year}-01-01"

    sess = session_from_env()
    print(f"[airtable] Fetching pedidos since {year_start}...")
    records = fetch_pedidos(sess, year_start)
    print(f"  {len(records)} pedidos recibidos")

    ventas, flat = aggregate(records)
    payload = {
        "_meta": {
            "last_fetched_at": datetime.now(timezone.utc).isoformat(),
            "schema_version": 1,
            "year": year,
        },
        "ventas_by_chain_month": ventas,
        "records": flat,
    }
    p = Path(args.cache)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[airtable] Saved {args.cache}")

    # Resumen
    for ch, months in sorted(ventas.items()):
        total = sum(months.values())
        print(f"  {ch}: {total} devices ({', '.join(f'{m}={n}' for m, n in sorted(months.items()))})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
