"""
Genera el HTML de la pestaña "Por cadena" (Fase 2).

Dependencias:
  - cache/jira_status_timeline.json (cache de fetch_jira.py)
  - data/serial_year.json (mapping serial→año)
  - cache/airtable_pedidos.json (ventas nuevas, opcional)

Exporta: build_chains_html(cache, airtable, serial_year) → string HTML
para reemplazar __CHAINS_HTML__ en el template.
"""
from __future__ import annotations

import re
from datetime import date, datetime, timezone

# Re-uso constantes de build_report.py
try:
    from build_report import (
        OPEN_STATUSES, FUNNEL_TAG, chain_of, parse_iso, is_open,
        last_status_change_dt, days_since, days_bucket,
        COLOR_GREEN, COLOR_YELLOW, COLOR_RED, FUNNEL_COLOR,
    )
except ImportError:
    # Fallback: importar individualmente si hay problemas
    pass


CHAIN_ORDER = ["Elha", "Sin Vello", "Dermasana", "Smart Duck",
               "Epil Point", "Laser Factory", "Unico Italia", "Otros"]

# Parque base a diciembre 2025 (último cierre confirmado)
PARQUE_DEC25 = {
    "Elha": 455, "Sin Vello": 267, "Dermasana": 49, "Smart Duck": 12,
    "Epil Point": 97, "Laser Factory": 52, "Unico Italia": 36, "Otros": 0,
}

MONTH_LABELS_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
                   "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

MOTIVO_BUCKETS = ["Diodo", "Umbilical", "Puntera", "Placa",
                  "Pantalla", "Gatillo", "Buffer", "Otros", "No resuelto aún"]

COSTE_BUCKETS = ["Sin coste", "< 1.000 €", "1.000 - 6.000 €", "> 6.000 €"]
SUSTI_BUCKETS = ["Sí entregada", "Sin sustitución"]
YEAR_BUCKETS = ["2026", "2025", "2024", "2023", "2022", "<2022", "Sin compra"]


def html_escape(s):
    if s is None:
        return ""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def add_serial_variants(s, variants):
    if s is None:
        return
    s = str(s).strip()
    if not s or s.lower() in ('0', '0.0', 'no aparece', 'none', 'null', ''):
        return
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    variants.add(s)
    variants.add(s.lower())
    variants.add(s.upper())
    m = re.match(r'^([CH])(\d+)$', s, re.IGNORECASE)
    if m:
        digits = m.group(2).lstrip('0')
        if digits:
            variants.add(digits)
        if len(m.group(2)) < 5:
            variants.add(m.group(1).upper() + m.group(2).zfill(5))
    if s.isdigit():
        padded = s.zfill(5)
        variants.update([s.lstrip('0'), 'C' + padded, 'H' + padded])


def serial_candidates(ticket):
    variants = set()
    for fld in ['consola', 'hp']:
        v = (ticket.get(fld) or '').strip()
        if v:
            add_serial_variants(v, variants)
    cliente = ticket.get('cliente', '') or ''
    for m in re.finditer(r'\b([CHMcheh]+\d{3,5}|\d{4,5})\b', cliente):
        s = m.group(1)
        m2 = re.match(r'^MHP?(\d+)$', s, re.IGNORECASE)
        if m2:
            add_serial_variants(m2.group(1), variants)
        else:
            add_serial_variants(s, variants)
    return variants


def lookup_year(serial_year, ticket, max_known=None):
    """Devuelve el año (int) o None si no se encuentra.
    Si no hay match y el serial es mayor que max_known[prefijo], asume año actual."""
    cands = serial_candidates(ticket)
    yr = None
    for s in cands:
        if s in serial_year:
            y = serial_year[s]
            if yr is None or y > yr:
                yr = y
    if yr is not None:
        return yr
    # Heurística: serial avanzado no mapeado → año actual
    if max_known:
        cur_year = datetime.now().year
        for c in cands:
            m = re.match(r'^([CH])(\d+)$', c, re.IGNORECASE)
            if m:
                prefix = m.group(1).upper()
                num = int(m.group(2))
                max_n = max_known.get(prefix, 0)
                if num > max_n:
                    return cur_year
    return None


def bucket_year(y):
    if y is None:
        return "Sin compra"
    if y >= 2026:
        return "2026"
    if y in (2025, 2024, 2023, 2022):
        return str(y)
    return "<2022"


def bucket_motivo(motivos, status):
    """Clasifica la lista de motivos en 1 bucket. Si status no es cerrado y no hay motivo, 'No resuelto aún'."""
    cerrados = ("Finalizada", "Resuelto", "Cancelado", "Finalizado técnico externo")
    if not motivos:
        if status in cerrados:
            return "Otros"
        return "No resuelto aún"
    txt = " ".join(motivos).lower()
    if "diodo" in txt: return "Diodo"
    if "umbilical" in txt or "reparación de umbilical" in txt: return "Umbilical"
    if "puntera" in txt or "zafiro" in txt: return "Puntera"
    if "placa" in txt: return "Placa"
    if "pantalla" in txt: return "Pantalla"
    if "gatillo" in txt: return "Gatillo"
    if "buffer" in txt: return "Buffer"
    return "Otros"


def bucket_coste(importe):
    try:
        v = float(importe) if importe not in ("", None) else 0
    except (TypeError, ValueError):
        v = 0
    if v == 0:
        return "Sin coste"
    if v < 1000:
        return "< 1.000 €"
    if v <= 6000:
        return "1.000 - 6.000 €"
    return "> 6.000 €"


def chain_of_ticket(t):
    return chain_of(t.get("cliente", ""), t.get("loc", ""))


def compute_max_known_per_prefix(serial_year):
    """Para la heurística 2026: máximo número por prefijo C/H en el mapping."""
    max_n = {"C": 0, "H": 0}
    for k in serial_year.keys():
        m = re.match(r'^([CH])(\d+)$', k)
        if m:
            p = m.group(1)
            n = int(m.group(2))
            if n > max_n.get(p, 0):
                max_n[p] = n
    return max_n


def parque_for_chain(chain, ventas, month_idx, year):
    """parque acumulado para una cadena al final del mes month_idx (1-12)."""
    base = PARQUE_DEC25.get(chain, 0)
    extra = 0
    if ventas and chain in ventas:
        for m in range(1, month_idx + 1):
            key = f"{year}-{m:02d}"
            extra += ventas[chain].get(key, 0)
    return base + extra


def build_chain_pane(chain, tickets_chain, ventas, serial_year, max_known,
                     year, num_months, today_total_open):
    """Construye el HTML para el panel de una cadena."""
    monthly = {m: {"nuevas": 0, "bloq": 0, "motivo": {}, "compra": {},
                   "coste": {}, "susti_si": 0, "susti_no": 0}
               for m in range(1, num_months + 1)}
    abiertas_hoy = 0
    for t in tickets_chain:
        created = parse_iso(t.get("created"))
        if not created or created.year != year:
            # Para 2026 only; los de años pasados no cuentan en YTD
            pass
        if is_open(t.get("current_status")):
            abiertas_hoy += 1
        if not created:
            continue
        m = created.month
        if m > num_months:
            continue
        if created.year != year:
            continue
        b = monthly[m]
        b["nuevas"] += 1
        if t.get("bloq") == "Sí":
            b["bloq"] += 1
        mot = bucket_motivo(t.get("motivo") or [], t.get("current_status") or "")
        b["motivo"][mot] = b["motivo"].get(mot, 0) + 1
        y = lookup_year(serial_year, t, max_known)
        yb = bucket_year(y)
        b["compra"][yb] = b["compra"].get(yb, 0) + 1
        cb = bucket_coste(t.get("importe"))
        b["coste"][cb] = b["coste"].get(cb, 0) + 1
        if t.get("susti") == "Sí":
            b["susti_si"] += 1
        else:
            b["susti_no"] += 1

    # YTD totals
    ytd_nuevas = sum(monthly[m]["nuevas"] for m in range(1, num_months + 1))
    ytd_bloq = sum(monthly[m]["bloq"] for m in range(1, num_months + 1))
    parque_ult = parque_for_chain(chain, ventas, num_months, year)
    tasa_ytd = (ytd_bloq / parque_ult / num_months) if (parque_ult and num_months) else 0
    bloq_ano = tasa_ytd * 12

    # === Tabla Evolución mensual ===
    headers = "".join(f'<th>{MONTH_LABELS_ES[m-1]}</th>' for m in range(1, num_months + 1))
    out = []
    out.append('<p class="chain-section-title">Evolución mensual</p>')
    out.append('<table class="evol-table"><thead><tr><th>Métrica</th>')
    out.append(headers)
    out.append('<th class="ytd">YTD</th></tr></thead><tbody>')

    def row(label, values, ytd, klass=""):
        cells = "".join(f'<td>{v}</td>' for v in values)
        return f'<tr class="{klass}"><td class="lbl">{label}</td>{cells}<td>{ytd}</td></tr>'

    nu_vals = [monthly[m]["nuevas"] or "-" for m in range(1, num_months + 1)]
    bl_vals = [monthly[m]["bloq"] or "-" for m in range(1, num_months + 1)]
    parque_vals = [parque_for_chain(chain, ventas, m, year) for m in range(1, num_months + 1)]
    tasa_vals = []
    bloq_ano_vals = []
    for m in range(1, num_months + 1):
        p_m = parque_vals[m - 1]
        b_m = monthly[m]["bloq"]
        t_m = (b_m / p_m) if p_m else 0
        tasa_vals.append(f"{t_m*100:.2f}%" if t_m else "-")
        bloq_ano_vals.append(f"{t_m*12:.2f}" if t_m else "-")

    out.append(row("Nuevas incidencias", nu_vals, ytd_nuevas, "zebra"))
    out.append(row("Bloqueantes", bl_vals, ytd_bloq))
    out.append(row("Total parque", parque_vals, parque_ult, "zebra"))
    out.append(row("Tasa bloqueantes", tasa_vals, f"{tasa_ytd*100:.2f}%"))
    out.append(row("Bloq/año/equipo", bloq_ano_vals, f"{bloq_ano:.2f}", "zebra"))
    out.append(row("Abiertas hoy", ["-"] * num_months, abiertas_hoy, "abiertas"))
    out.append('</tbody></table>')

    # === Tablas mensuales: motivo / compra / coste / sustitución ===
    def breakdown_table(title, buckets, extract_dict, ytd_dict):
        h = [f'<p class="chain-section-title">{title}</p>',
             '<table class="breakdown-table"><thead><tr><th>',
             buckets[0] if title == "Motivo avería" else title,
             '</th>', headers, '<th class="ytd">YTD</th></tr></thead><tbody>']
        # Replace generic header label
        return None  # see below

    def make_breakdown(title, buckets):
        rows = []
        ytd_totals = {b: 0 for b in buckets}
        for b in buckets:
            ytd_totals[b] = sum(monthly[m].get(key_for_bucket(title), {}).get(b, 0)
                                for m in range(1, num_months + 1))
        return rows, ytd_totals

    def key_for_bucket(title):
        return {"Motivo avería": "motivo", "Fecha de compra del equipo": "compra",
                "Coste": "coste"}.get(title, "motivo")

    for title, buckets in [
        ("Motivo avería", MOTIVO_BUCKETS),
        ("Fecha de compra del equipo", YEAR_BUCKETS),
        ("Coste", COSTE_BUCKETS),
    ]:
        key = key_for_bucket(title)
        out.append(f'<p class="chain-section-title">{title}</p>')
        out.append('<table class="breakdown-table"><thead><tr><th>')
        out.append({"motivo": "Motivo", "compra": "Año compra", "coste": "Rango"}[key])
        out.append('</th>')
        out.append(headers)
        out.append('<th class="ytd">YTD</th></tr></thead><tbody>')
        for i, b in enumerate(buckets):
            vals = [monthly[m][key].get(b, 0) for m in range(1, num_months + 1)]
            ytd = sum(vals)
            if ytd == 0:
                continue
            vals_str = [v if v else "-" for v in vals]
            cells = "".join(f'<td>{v}</td>' for v in vals_str)
            klass = "zebra" if i % 2 else ""
            out.append(f'<tr class="{klass}"><td class="lbl">{b}</td>{cells}<td>{ytd}</td></tr>')
        out.append('</tbody></table>')

    # Sustitución
    out.append('<p class="chain-section-title">Sustitución entregada</p>')
    out.append('<table class="breakdown-table"><thead><tr><th>Sustitución</th>')
    out.append(headers)
    out.append('<th class="ytd">YTD</th></tr></thead><tbody>')
    si_vals = [monthly[m]["susti_si"] or "-" for m in range(1, num_months + 1)]
    no_vals = [monthly[m]["susti_no"] or "-" for m in range(1, num_months + 1)]
    si_ytd = sum(monthly[m]["susti_si"] for m in range(1, num_months + 1))
    no_ytd = sum(monthly[m]["susti_no"] for m in range(1, num_months + 1))
    out.append('<tr><td class="lbl">Sí entregada</td>' + "".join(f'<td>{v}</td>' for v in si_vals) + f'<td>{si_ytd}</td></tr>')
    out.append('<tr class="zebra"><td class="lbl">Sin sustitución</td>' + "".join(f'<td>{v}</td>' for v in no_vals) + f'<td>{no_ytd}</td></tr>')
    out.append('</tbody></table>')

    # Tabla simple de incidencias abiertas
    open_rows = [t for t in tickets_chain if is_open(t.get("current_status"))]
    open_rows.sort(key=lambda r: -(days_since(last_status_change_dt(r)) or 0))
    JIRA_URL = "https://leaseir.atlassian.net/browse/"
    out.append(f'<p class="chain-section-title">Incidencias abiertas hoy ({len(open_rows)})</p>')
    out.append('<div class="scroller"><table style="font-size:11px;min-width:1100px"><thead><tr>')
    for h in ["Ticket", "Cliente / Centro", "Estado", "Gestión", "Días", "Localización",
              "Creada", "Tipo", "Bloq", "Susti", "Fecha venta", "Consola", "HP", "Garantía"]:
        out.append(f'<th>{h}</th>')
    out.append('</tr></thead><tbody>')
    for r in open_rows:
        st = r.get("current_status", "")
        ft = FUNNEL_TAG.get(st, "?")
        ft_color = FUNNEL_COLOR.get(ft, "#888")
        days = days_since(last_status_change_dt(r))
        days_color = COLOR_GREEN if (days is not None and days < 5) else (
            COLOR_YELLOW if (days is not None and days <= 15) else COLOR_RED)
        created_dt = parse_iso(r.get("created"))
        created_s = created_dt.strftime("%d/%m/%Y") if created_dt else ""
        fventa_dt = parse_iso(r.get("fventa") or "")
        fventa_s = fventa_dt.strftime("%d/%m/%Y") if fventa_dt else (r.get("fventa") or "")
        link = f'<a href="{JIRA_URL}{r["key"]}" target="_blank" rel="noopener" style="color:#2a59c4;text-decoration:none">{r["key"]}</a>'
        out.append(
            '<tr>'
            f'<td style="text-align:left">{link}</td>'
            f'<td class="d-trunc" style="text-align:left" title="{html_escape(r.get("cliente",""))}">{html_escape(r.get("cliente",""))}</td>'
            f'<td style="text-align:left">{html_escape(st)}</td>'
            f'<td style="color:{ft_color};font-weight:500">{ft}</td>'
            f'<td style="color:{days_color};font-weight:500">{days if days is not None else ""}</td>'
            f'<td class="d-trunc" style="text-align:left" title="{html_escape(r.get("loc",""))}">{html_escape(r.get("loc",""))}</td>'
            f'<td>{created_s}</td>'
            f'<td style="text-align:left">{html_escape(r.get("tipo",""))}</td>'
            f'<td>{html_escape(r.get("bloq",""))}</td>'
            f'<td>{html_escape(r.get("susti",""))}</td>'
            f'<td>{fventa_s}</td>'
            f'<td>{html_escape(r.get("consola",""))}</td>'
            f'<td>{html_escape(r.get("hp",""))}</td>'
            f'<td>{html_escape(r.get("garantia",""))}</td>'
            '</tr>'
        )
    out.append('</tbody></table></div>')

    return "".join(out), {
        "nuevas": ytd_nuevas, "bloq": ytd_bloq, "parque": parque_ult,
        "tasa": tasa_ytd, "bloq_ano": bloq_ano, "abiertas": abiertas_hoy,
    }


def build_resumen_pane(stats):
    """Tabla resumen agregada (Resumen tab)."""
    out = ['<table class="breakdown-table" style="margin-top:8px"><thead><tr>',
           '<th>Cadena</th><th>Nuevas</th><th>Bloqueantes</th><th>Parque</th>',
           '<th>Tasa</th><th>Bloq/año/eq</th><th>Abiertas hoy</th>',
           '</tr></thead><tbody>']
    totals = {"nuevas": 0, "bloq": 0, "parque": 0, "abiertas": 0}
    for i, ch in enumerate(CHAIN_ORDER):
        s = stats.get(ch, {})
        if not s and ch == "Otros":
            continue
        nu, bl = s.get("nuevas", 0), s.get("bloq", 0)
        p = s.get("parque", 0)
        ta = s.get("tasa", 0)
        ba = s.get("bloq_ano", 0)
        ab = s.get("abiertas", 0)
        totals["nuevas"] += nu
        totals["bloq"] += bl
        totals["parque"] += p
        totals["abiertas"] += ab
        klass = "zebra" if i % 2 else ""
        out.append(
            f'<tr class="{klass}"><td class="lbl">{ch}</td>'
            f'<td>{nu}</td><td>{bl}</td><td>{p}</td>'
            f'<td>{ta*100:.2f}%</td><td>{ba:.2f}</td><td>{ab}</td></tr>'
        )
    # Total
    num_chains = max(1, sum(1 for ch in CHAIN_ORDER if stats.get(ch, {}).get("parque")))
    tasa_global = (totals["bloq"] / totals["parque"]) if totals["parque"] else 0
    out.append(
        f'<tr style="background:#fff4d6;font-weight:500"><td class="lbl">Total</td>'
        f'<td>{totals["nuevas"]}</td><td>{totals["bloq"]}</td><td>{totals["parque"]}</td>'
        f'<td>{tasa_global*100:.2f}%</td><td>—</td><td>{totals["abiertas"]}</td></tr>'
    )
    out.append('</tbody></table>')
    return "".join(out)


def build_chains_html(cache, airtable, serial_year):
    """Devuelve el HTML completo para __CHAINS_HTML__."""
    today = date.today()
    year = today.year
    num_months = today.month
    max_known = compute_max_known_per_prefix(serial_year)

    # Group tickets by chain
    tickets = cache.get("tickets", {})
    by_chain = {ch: [] for ch in CHAIN_ORDER}
    for k, t in tickets.items():
        ch = chain_of_ticket(t)
        if ch not in by_chain:
            by_chain[ch] = []
        by_chain[ch].append(t)

    ventas = (airtable or {}).get("ventas_by_chain_month", {})

    chains_to_show = [ch for ch in CHAIN_ORDER if ch != "Otros"]

    # Tabs2 navigation
    tabs = ['<div class="tabs2">']
    tabs.append('<button type="button" data-chain="resumen" class="active">Resumen</button>')
    for ch in chains_to_show:
        tabs.append(f'<button type="button" data-chain="{html_escape(ch)}">{html_escape(ch)}</button>')
    tabs.append('</div>')

    # Build panes
    stats_per_chain = {}
    panes = []
    for ch in chains_to_show:
        pane_html, s = build_chain_pane(ch, by_chain.get(ch, []), ventas,
                                        serial_year, max_known, year, num_months,
                                        sum(1 for t in by_chain.get(ch, []) if is_open(t.get("current_status"))))
        stats_per_chain[ch] = s
        meta = (f'{s["nuevas"]} nuevas YTD · {s["bloq"]} bloqueantes · '
                f'{s["abiertas"]} abiertas hoy')
        panes.append(
            f'<div class="chain-pane" data-chain="{html_escape(ch)}">'
            f'<h3>{html_escape(ch)}</h3>'
            f'<p class="chain-meta">{meta}</p>'
            f'{pane_html}</div>'
        )

    # Resumen pane (first)
    resumen_html = (
        '<div class="chain-pane active" data-chain="resumen">'
        '<h3>Resumen por cadena</h3>'
        '<p class="chain-meta">YTD · 5 métricas estándar (mismas del email a Ignacio) + abiertas hoy</p>'
        + build_resumen_pane(stats_per_chain)
        + '</div>'
    )

    return "".join(tabs) + resumen_html + "".join(panes)
