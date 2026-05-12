"""
Pull de Airtable Inmovilizado (Leaseir) → equipos en estado SAT o Backups for customers.
Auth: AIRTABLE_API_KEY env var (mismo PAT que fetch_airtable.py).

Output: cache/airtable_inmovilizado.json
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests

BASE_ID = "app9U5sz7YS8y9Oit"   # Leaseir
TABLE_ID = "tblSbQH0GaKghVODz"  # Inmovilizado

# Campos clave
FLD_ID = "fldcYjOBSG7HCGQZB"            # ID (serial)
FLD_TYPE_ASSET = "fldemw9wbX80CHEqR"    # Type of Asset (Console / Handpiece)
FLD_CONSOLE = "fldWM2Yz93iePFhcO"       # Console (modelo: MHR, AHR, etc)
FLD_SPOT = "fldghW6Mvgmjr15k1"          # Spot Size (Quad, Dual, SRF, Single)
FLD_ACTIVITY = "fldl0Nfiky8GiKuDm"      # Current Activity
FLD_CUSTOMER = "fldJVSJi4x69CxYA5"      # Customer
FLD_COMMENTS = "fld0eQ5WNIhbAq1SN"      # Comments
FLD_RENOMBRADA = "fldUEroKktAIMdVGX"    # Renombrada (alias)

SEL_SAT = "selbYwGkNg8EyxNba"
SEL_BACKUPS = "selY5ybVUIT0WCnZf"


def session_from_env():
    key = os.environ.get("AIRTABLE_API_KEY")
    if not key:
        print("FATAL: AIRTABLE_API_KEY env var no definida", file=sys.stderr)
        sys.exit(2)
    s = requests.Session()
    s.headers["Authorization"] = f"Bearer {key}"
    s.headers["Accept"] = "application/json"
    return s


def fetch_all(sess):
    """Pull todos los registros con Current Activity ∈ {SAT, Backups for customers}."""
    url = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_ID}"
    # Filtro usando formula (string en el sel value)
    formula = "OR({Current Activity}='SAT',{Current Activity}='Backups for customers')"
    params = {"pageSize": 100, "filterByFormula": formula}
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


def normalize_record(rec):
    f = rec.get("fields", {})
    return {
        "id_rec": rec.get("id"),
        "serial": (f.get("ID") or "").strip(),
        "type_asset": f.get("Type of Asset", ""),
        "console_model": f.get("Console", ""),
        "spot_size": f.get("Spot Size", ""),
        "activity": f.get("Current Activity", ""),
        "customer": f.get("Customer", ""),
        "comments": f.get("Comments", "") or "",
        "renombrada": f.get("Renombrada", "") or "",
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", required=True)
    args = ap.parse_args()

    sess = session_from_env()
    print("[airtable-inmov] Fetching Inmovilizado (SAT + Backups)...")
    records = fetch_all(sess)
    print(f"  {len(records)} registros recibidos")

    items = [normalize_record(r) for r in records]
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
    print(f"[airtable-inmov] Saved {len(items)} items to {args.cache}")

    from collections import Counter
    by_act = Counter(i["activity"] for i in items)
    by_type = Counter(i["type_asset"] for i in items)
    print(f"  Por Activity: {dict(by_act)}")
    print(f"  Por Type: {dict(by_type)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
