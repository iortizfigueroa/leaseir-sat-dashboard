"""enrich_invoices_from_holded.py — para cada ticket con presupuesto (E2xxxxx),
busca la factura asociada en Holded y rellena en el cache:
  - factura            : docNumber de la invoice (e.g. "SInv26-202600945")
  - factura_importe    : total (float, €)
  - factura_cobrado    : suma de pagos (float, €)
  - factura_pendiente  : total - cobrado
  - factura_estado     : "Cobrado" | "Vencido" | "Pendiente" | ""
  - factura_fecha_venc : timestamp UNIX del vencimiento (int o None)

Estrategia de cruce:
  Holded enlaza un estimate (E26xxxx) aceptado con su factura. En la API la
  factura suele exponer ese vínculo en `from.docNumber`. Como fallback, busca
  el código de presupuesto en `notes`, `desc` o `description` de la factura.

Si un mismo presupuesto tiene varias facturas, nos quedamos con la más reciente
(criterio: campo `date` de la invoice, mayor = más nueva).

Requiere variable de entorno HOLDED_API_KEY. Si no está, sale sin hacer nada
(no es bloqueante; el workflow lleva `continue-on-error: true`).

Uso: python scripts/enrich_invoices_from_holded.py --cache cache/jira_status_timeline.json
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError


HOLDED_BASE = "https://api.holded.com/api/invoicing/v1/documents"
# Códigos de presupuesto Holded para Leaseir: E2YYxxx, E2YYxxxx, etc.
PPTO_PAT = re.compile(r"E2[0-9]\d{3,6}", re.IGNORECASE)


def fetch_all_invoices(api_key, timeout=30):
    """Devuelve lista completa de invoices Holded paginadas."""
    out = []
    for page in range(1, 21):  # límite seguridad: 20 pages * 500 = 10.000
        url = f"{HOLDED_BASE}/invoice?starttmp=1700000000&endtmp=2000000000&page={page}"
        req = Request(url, headers={"key": api_key, "Accept": "application/json"})
        try:
            with urlopen(req, timeout=timeout) as r:
                data = json.loads(r.read().decode("utf-8"))
        except (URLError, HTTPError, ValueError) as e:
            print(f"  [warn] invoice page {page} falló: {e}", file=sys.stderr)
            break
        if not isinstance(data, list) or not data:
            break
        out.extend(data)
        if len(data) < 500:
            break
    return out


def _safe_str(v):
    return v if isinstance(v, str) else (str(v) if v is not None else "")


def build_estimate_to_invoice(invoices):
    """Indexa {estimate_docNumber_uppercase: invoice} probando varios campos.

    Prefer matches por `from.docNumber` (vínculo directo de Holded).
    Si la misma estimate apunta a varias invoices, conservamos la de fecha
    más reciente.
    """
    idx = {}
    for inv in invoices:
        candidates = []
        # 1) from.docNumber — convención Holded para documentos derivados
        frm = inv.get("from")
        if isinstance(frm, dict):
            dn = _safe_str(frm.get("docNumber")).strip()
            if dn and PPTO_PAT.match(dn):
                candidates.append(dn)
        elif isinstance(frm, list):
            for entry in frm:
                if isinstance(entry, dict):
                    dn = _safe_str(entry.get("docNumber")).strip()
                    if dn and PPTO_PAT.match(dn):
                        candidates.append(dn)
        # 2) campos forecastedFrom / fromDocNumber (otras variantes)
        for fld in ("forecastedFrom", "fromDocNumber", "originalDocNumber"):
            v = _safe_str(inv.get(fld)).strip()
            if v and PPTO_PAT.match(v):
                candidates.append(v)
        # 3) fallback: códigos E2xxx en notes / desc / description
        if not candidates:
            for fld in ("notes", "desc", "description"):
                v = _safe_str(inv.get(fld))
                if v:
                    for m in PPTO_PAT.findall(v):
                        candidates.append(m)
        # Pick más reciente
        inv_date = inv.get("date") or 0
        try:
            inv_date = int(inv_date)
        except (ValueError, TypeError):
            inv_date = 0
        for k in candidates:
            k_up = k.upper()
            prev = idx.get(k_up)
            if not prev:
                idx[k_up] = inv
                continue
            prev_date = prev.get("date") or 0
            try:
                prev_date = int(prev_date)
            except (ValueError, TypeError):
                prev_date = 0
            if inv_date > prev_date:
                idx[k_up] = inv
    return idx


def compute_invoice_state(total, paid, due_ts):
    """Devuelve (estado, pending). Si pending <= 1 céntimo => Cobrado."""
    pending = round((total or 0) - (paid or 0), 2)
    if pending <= 0.01:
        return "Cobrado", 0.0
    now = datetime.now(timezone.utc).timestamp()
    if due_ts and now > due_ts:
        return "Vencido", pending
    return "Pendiente", pending


def _f(v):
    try:
        return float(v or 0)
    except (ValueError, TypeError):
        return 0.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", required=True)
    args = ap.parse_args()

    api_key = os.environ.get("HOLDED_API_KEY", "").strip()
    if not api_key:
        print("[enrich-invoices] HOLDED_API_KEY no definida, skip.")
        return 0

    cache_path = Path(args.cache)
    if not cache_path.exists():
        print(f"[enrich-invoices] cache {cache_path} no existe, skip.")
        return 0

    cache = json.loads(cache_path.read_text(encoding="utf-8"))
    tickets = cache.get("tickets", {}) or {}

    with_ppto = [(k, t) for k, t in tickets.items() if (t.get("presupuesto") or "").strip()]
    if not with_ppto:
        print("[enrich-invoices] no hay tickets con presupuesto, skip.")
        return 0

    print(f"[enrich-invoices] {len(with_ppto)} tickets con presupuesto. Consultando Holded invoices...")
    invoices = fetch_all_invoices(api_key)
    print(f"[enrich-invoices] {len(invoices)} invoices descargadas.")
    if invoices:
        sample_keys = sorted(list(invoices[0].keys()))[:25]
        print(f"[enrich-invoices] DEBUG primera invoice (claves): {sample_keys}")

    idx = build_estimate_to_invoice(invoices)
    print(f"[enrich-invoices] {len(idx)} estimates con factura mapeada.")

    enriched = 0
    no_match = 0
    for k, t in with_ppto:
        ppto = (t.get("presupuesto") or "").strip().upper()
        inv = idx.get(ppto)
        if not inv:
            # Limpia campos antiguos por si quedaron de runs previos
            t["factura"] = ""
            t["factura_importe"] = None
            t["factura_cobrado"] = None
            t["factura_pendiente"] = None
            t["factura_estado"] = ""
            t["factura_fecha_venc"] = None
            no_match += 1
            continue
        total = _f(inv.get("total"))
        paid = _f(inv.get("paymentsTotal") or inv.get("paid") or 0)
        # Vencimiento: prueba varios campos (Holded usa `dueDate` en unixtime)
        due_raw = inv.get("dueDate") or inv.get("duedate") or inv.get("expirationDate") or None
        try:
            due_ts = int(due_raw) if due_raw else None
        except (ValueError, TypeError):
            due_ts = None
        estado, pending = compute_invoice_state(total, paid, due_ts)
        t["factura"] = _safe_str(inv.get("docNumber")).strip()
        t["factura_importe"] = round(total, 2)
        t["factura_cobrado"] = round(paid, 2)
        t["factura_pendiente"] = round(pending, 2)
        t["factura_estado"] = estado
        t["factura_fecha_venc"] = due_ts
        enriched += 1

    cache["tickets"] = tickets
    cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[enrich-invoices] {enriched} tickets con factura encontrada, {no_match} sin factura.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
