"""enrich_invoices_from_holded.py - cruce ppto-factura via from.id en Holded.

Para cada ticket con presupuesto (E2xxxxx) busca la factura asociada en Holded
y rellena en el cache:
  - factura            : docNumber de la invoice (SInv26-202600750)
  - factura_importe    : total (float EUR)
  - factura_cobrado    : paymentsTotal (float EUR)
  - factura_pendiente  : total - paymentsTotal
  - factura_estado     : "Cobrado" | "Vencido" | "Pendiente" | ""
  - factura_fecha_venc : timestamp UNIX (int) o None

Cruce: Holded enlaza por ID interno (invoice.from.id apunta a estimate.id).
Pasos: 1) descargar estimates indexando id -> docNumber, 2) descargar invoices
y para cada una resolver invoice.from.id al docNumber del estimate origen,
3) fallback a texto E2xxx en notes/desc si no hay from.

Multi-factura: usamos la mas reciente (campo date).
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
PPTO_PAT = re.compile(r"E2[0-9]\d{3,6}", re.IGNORECASE)


def fetch_all(doc_type, api_key, timeout=30):
    out = []
    for page in range(1, 21):
        url = f"{HOLDED_BASE}/{doc_type}?starttmp=1700000000&endtmp=2000000000&page={page}"
        req = Request(url, headers={"key": api_key, "Accept": "application/json"})
        try:
            with urlopen(req, timeout=timeout) as r:
                data = json.loads(r.read().decode("utf-8"))
        except (URLError, HTTPError, ValueError) as e:
            print(f"  [warn] {doc_type} page {page} fallo: {e}", file=sys.stderr)
            break
        if not isinstance(data, list) or not data:
            break
        out.extend(data)
        if len(data) < 500:
            break
    return out


def _safe_str(v):
    return v if isinstance(v, str) else (str(v) if v is not None else "")


def _extract_ppto_code(raw):
    """Extrae el codigo E2xxxxx de un campo presupuesto crudo.

    El campo puede ser un codigo limpio ('E260398'), el filename de un
    attachment ('E260398 LEASEIR Brull.pdf'), o algo con ruido. Devuelve el
    codigo upper-case o '' si no encuentra match.
    """
    if not raw:
        return ""
    s = str(raw).strip()
    m = PPTO_PAT.search(s)
    return m.group(0).upper() if m else ""


def _to_int(v):
    try:
        return int(v) if v else 0
    except (ValueError, TypeError):
        return 0


def build_estimate_to_invoice(estimates, invoices):
    est_id_to_docnum = {}
    for e in estimates:
        eid = e.get("id")
        dn = _safe_str(e.get("docNumber")).strip()
        if eid and dn:
            est_id_to_docnum[eid] = dn

    idx = {}
    n_by_id = 0
    n_by_text = 0

    for inv in invoices:
        inv_date = _to_int(inv.get("date"))
        candidates = []

        frm = inv.get("from")
        if isinstance(frm, dict):
            ftype = _safe_str(frm.get("docType")).lower()
            fid = _safe_str(frm.get("id")).strip()
            if ftype == "estimate" and fid:
                est_dn = est_id_to_docnum.get(fid)
                if est_dn:
                    candidates.append(("id", est_dn))
        elif isinstance(frm, list):
            for entry in frm:
                if isinstance(entry, dict) and _safe_str(entry.get("docType")).lower() == "estimate":
                    fid = _safe_str(entry.get("id")).strip()
                    est_dn = est_id_to_docnum.get(fid)
                    if est_dn:
                        candidates.append(("id", est_dn))

        if not candidates:
            for fld in ("notes", "desc", "description"):
                v = _safe_str(inv.get(fld))
                if v:
                    for m in PPTO_PAT.findall(v):
                        candidates.append(("text", m))

        for source, k in candidates:
            k_up = k.upper()
            prev = idx.get(k_up)
            if not prev:
                idx[k_up] = inv
                if source == "id":
                    n_by_id += 1
                else:
                    n_by_text += 1
            else:
                if inv_date > _to_int(prev.get("date")):
                    idx[k_up] = inv

    print(f"[enrich-invoices] cruces: {n_by_id} por from.id, {n_by_text} por texto/desc")
    return idx


def compute_invoice_state(total, paid, due_ts):
    total = total or 0
    paid = paid or 0
    pending = round(total - paid, 2)
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

    # Tickets con presupuesto + codigo E2xxxxx extraido
    with_ppto = []
    n_dirty = 0
    for k, t in tickets.items():
        ppto_raw = (t.get("presupuesto") or "").strip()
        if not ppto_raw:
            continue
        code = _extract_ppto_code(ppto_raw)
        if not code:
            n_dirty += 1
            continue
        with_ppto.append((k, t, code))
    if not with_ppto:
        print("[enrich-invoices] no hay tickets con presupuesto, skip.")
        return 0
    if n_dirty:
        print(f"[enrich-invoices] {n_dirty} tickets con presupuesto pero sin codigo E2xxx valido.")

    print(f"[enrich-invoices] {len(with_ppto)} tickets con presupuesto. Descargando Holded...")
    estimates = fetch_all("estimate", api_key)
    print(f"[enrich-invoices] {len(estimates)} estimates descargados.")
    invoices = fetch_all("invoice", api_key)
    print(f"[enrich-invoices] {len(invoices)} invoices descargadas.")

    idx = build_estimate_to_invoice(estimates, invoices)
    print(f"[enrich-invoices] {len(idx)} estimates con factura mapeada.")

    enriched = 0
    no_match = 0
    for k, t, ppto in with_ppto:
        inv = idx.get(ppto)
        if not inv:
            t["factura"] = ""
            t["factura_importe"] = None
            t["factura_cobrado"] = None
            t["factura_pendiente"] = None
            t["factura_estado"] = ""
            t["factura_fecha_venc"] = None
            no_match += 1
            continue
        total = _f(inv.get("total"))
        paid = _f(inv.get("paymentsTotal"))
        due_raw = inv.get("dueDate") or inv.get("duedate") or None
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
    print(f"[enrich-invoices] {enriched} tickets con factura, {no_match} sin factura.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
