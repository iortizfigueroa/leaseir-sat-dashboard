"""diag_holded_invoice.py — diagnóstico puntual del cruce estimate↔invoice en Holded.

Busca la factura SInv26-202600750 y el estimate E260629, imprime todos los campos
relevantes (especialmente el campo `from` y similares) para entender cómo Holded
expone el vínculo entre ambos documentos.

Uso (desde la carpeta sat-dashboard, con HOLDED_API_KEY en entorno):

    $env:HOLDED_API_KEY = "tu_clave"
    python scripts/diag_holded_invoice.py

(opcional: pasar otros docNumbers como argumentos)
"""
import json
import os
import sys
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError


HOLDED_BASE = "https://api.holded.com/api/invoicing/v1/documents"


def fetch_all(doc_type, api_key, timeout=30):
    out = []
    for page in range(1, 21):
        url = f"{HOLDED_BASE}/{doc_type}?starttmp=1700000000&endtmp=2000000000&page={page}"
        req = Request(url, headers={"key": api_key, "Accept": "application/json"})
        try:
            with urlopen(req, timeout=timeout) as r:
                data = json.loads(r.read().decode("utf-8"))
        except (URLError, HTTPError, ValueError) as e:
            print(f"  [warn] {doc_type} page {page} fallo: {e}")
            break
        if not isinstance(data, list) or not data:
            break
        out.extend(data)
        if len(data) < 500:
            break
    return out


def fetch_detail(doc_type, doc_id, api_key, timeout=30):
    url = f"{HOLDED_BASE}/{doc_type}/{doc_id}"
    req = Request(url, headers={"key": api_key, "Accept": "application/json"})
    try:
        with urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except (URLError, HTTPError, ValueError) as e:
        print(f"  [warn] {doc_type}/{doc_id} fallo: {e}")
        return None


def main():
    api_key = os.environ.get("HOLDED_API_KEY", "").strip()
    if not api_key:
        print("ERROR: HOLDED_API_KEY no definida")
        return 1

    estimate_docnumber = "E260629"
    invoice_docnumber = "SInv26-202600750"
    if len(sys.argv) > 1:
        estimate_docnumber = sys.argv[1]
    if len(sys.argv) > 2:
        invoice_docnumber = sys.argv[2]

    print(f"=== Buscando estimate {estimate_docnumber} y invoice {invoice_docnumber} ===")
    print()
    print("Descargando estimates...")
    estimates = fetch_all("estimate", api_key)
    print(f"  {len(estimates)} estimates descargados")
    print("Descargando invoices...")
    invoices = fetch_all("invoice", api_key)
    print(f"  {len(invoices)} invoices descargados")
    print()

    # Encontrar el estimate y la invoice de interes
    est = next((e for e in estimates if (e.get("docNumber") or "").upper() == estimate_docnumber.upper()), None)
    inv = next((i for i in invoices if (i.get("docNumber") or "").upper() == invoice_docnumber.upper()), None)

    if est:
        print(f"=== Estimate {estimate_docnumber} ENCONTRADO ===")
        print(f"  id: {est.get('id')}")
        print(f"  status: {est.get('status')}")
        print(f"  total: {est.get('total')}")
        print(f"  date: {est.get('date')}")
        print(f"  desc[:100]: {(est.get('desc') or '')[:100]}")
        print(f"  Claves disponibles: {sorted(est.keys())}")
        # Buscar campos que apunten a invoice
        print()
        print("  Campos potencialmente apuntando a invoice:")
        for k, v in est.items():
            if "invoice" in k.lower() or "convert" in k.lower() or "to" == k.lower() or "linked" in k.lower():
                print(f"    {k}: {v!r}")
        # Pedir detalle
        print()
        print("  Pidiendo detalle del estimate...")
        detail = fetch_detail("estimate", est.get("id"), api_key)
        if detail:
            print(f"    detail claves: {sorted(detail.keys())}")
            for k, v in detail.items():
                if "invoice" in k.lower() or "convert" in k.lower() or "to" == k.lower() or "linked" in k.lower() or "status" in k.lower():
                    print(f"    detail.{k}: {v!r}")
    else:
        print(f"!!! Estimate {estimate_docnumber} NO encontrado")
    print()

    if inv:
        print(f"=== Invoice {invoice_docnumber} ENCONTRADA ===")
        print(f"  id: {inv.get('id')}")
        print(f"  status: {inv.get('status')}")
        print(f"  total: {inv.get('total')}")
        print(f"  paymentsTotal: {inv.get('paymentsTotal')}")
        print(f"  paid: {inv.get('paid')}")
        print(f"  dueDate: {inv.get('dueDate')}")
        print(f"  date: {inv.get('date')}")
        print(f"  Claves disponibles: {sorted(inv.keys())}")
        # Mostrar campos que pueden apuntar a estimate
        print()
        print("  Campos potencialmente apuntando a estimate (from/source/etc):")
        for k, v in inv.items():
            if "from" in k.lower() or "source" in k.lower() or "estimate" in k.lower() or "parent" in k.lower() or "original" in k.lower():
                print(f"    {k}: {v!r}")
        # Pedir detalle
        print()
        print("  Pidiendo detalle de la invoice...")
        detail = fetch_detail("invoice", inv.get("id"), api_key)
        if detail:
            print(f"    detail claves: {sorted(detail.keys())}")
            for k, v in detail.items():
                if "from" in k.lower() or "source" in k.lower() or "estimate" in k.lower() or "parent" in k.lower() or "original" in k.lower() or "linked" in k.lower():
                    print(f"    detail.{k}: {v!r}")
    else:
        print(f"!!! Invoice {invoice_docnumber} NO encontrada en la lista paginada")

    print()
    print("=== FIN ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
