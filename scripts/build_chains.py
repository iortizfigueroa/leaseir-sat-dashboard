"""
Genera el HTML de la pestaña "Por cadena" (Fase 2 iter B).

Cambios vs iter A:
- Tabla detalle muestra TODAS las incidencias del año (no solo abiertas)
- Cada celda numérica de las tablas mensuales lleva class="filter-cell" y data-filter
  JSON para que JS filtre la tabla detalle al click.
"""
from __future__ import annotations

import json as _json
import re
from datetime import date, datetime, timezone

from build_report import (
    OPEN_STATUSES, FUNNEL_TAG, chain_of, parse_iso, is_open,
    last_status_change_dt, days_since, days_bucket,
    COLOR_GREEN, COLOR_YELLOW, COLOR_RED, FUNNEL_COLOR,
)


CHAIN_ORDER = ["Elha", "Sin Vello", "Dermasana", "Smart Duck",
               "Epil Point", "Laser Factory", "Unico Italia", "Otros"]

PARQUE_DEC25 = {
    "Elha": 455, "Sin Vello": 267, "Dermasana": 49, "Smart Duck": 12,
    "Epil Point": 97, "Laser Factory": 52, "Unico Italia": 36, "Otros": 0,
}

MONTH_LABELS_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
                   "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

MOTIVO_BUCKETS = ["Diodo", "Umbilical", "Puntera", "Placa",
                  "Pantalla", "Gatillo", "Buffer", "Otros", "No resuelto aún"]
COSTE_BUCKETS = ["Sin coste", "< 1.000 €", "1.000 - 6.000 €", "> 6.000 €"]
YEAR_BUCKETS = ["2026", "2025", "2024", "2023", "2022", "<2022", "Sin compra"]


def html_escape(s):
    if s is None:
        return ""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def jdata(d):
    """Serializa dict a JSON para data-filter attr (con escape HTML)."""
    return html_escape(_json.dumps(d, ensure_ascii=False))


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
    cands = serial_candidates(ticket)
    yr = None
    for s in cands:
        if s in serial_year:
            y = serial_year[s]
            if yr is None or y > yr:
                yr = y
    if yr is not None:
        return yr
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
    cerrados = ("Finalizada", "Resuelto", "Cancelado", "Finalizado técnico externo")
    if not motivos:
        if status in cerrados:
            return "Otros"
        return "No resuelto aún"
    txt = " ".join(motivos).lower()
    if "diodo" in txt: return "Diodo"
    if "umbilical" in txt: return "Umbilical"
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
    base = PARQUE_DEC25.get(chain, 0)
    extra = 0
    if ventas and chain in ventas:
        for m in range(1, month_idx + 1):
            key = f"{year}-{m:02d}"
            extra += ventas[chain].get(key, 0)
    return base + extra


def enrich(t, serial_year, max_known, year):
    """Calcula campos derivados de un ticket para el detalle: month, bloq, susti, motivo, year_compra, coste, isopen."""
    created = parse_iso(t.get("created"))
    month = created.month if created and created.year == year else None
    y = lookup_year(serial_year, t, max_known)
    yb = bucket_year(y)
    return {
        "month": month,
        "bloq": t.get("bloq") or "No",
        "susti": t.get("susti") or "No",
        "motivo": bucket_motivo(t.get("motivo") or [], t.get("current_status") or ""),
        "yearcompra": yb,
        "coste": bucket_coste(t.get("importe")),
        "isopen": 1 if is_open(t.get("current_status")) else 0,
    }


def build_chain_pane(chain, tickets_chain, ventas, serial_year, max_known,
                     year, num_months):
    """HTML para el pane de una cadena."""
    # Enrich y filtrar a YTD
    enriched = []
    for t in tickets_chain:
        e = enrich(t, serial_year, max_known, year)
        if e["month"] is not None:
            enriched.append((t, e))

    # Stats mensuales
    monthly = {m: {"nuevas": 0, "bloq": 0, "motivo": {}, "yearcompra": {},
                   "coste": {}, "susti_si": 0, "susti_no": 0}
               for m in range(1, num_months + 1)}
    abiertas_hoy = sum(1 for t in tickets_chain if is_open(t.get("current_status")))
    for t, e in enriched:
        m = e["month"]
        if m > num_months:
            continue
        b = monthly[m]
        b["nuevas"] += 1
        if e["bloq"] == "Sí":
            b["bloq"] += 1
        b["motivo"][e["motivo"]] = b["motivo"].get(e["motivo"], 0) + 1
        b["yearcompra"][e["yearcompra"]] = b["yearcompra"].get(e["yearcompra"], 0) + 1
        b["coste"][e["coste"]] = b["coste"].get(e["coste"], 0) + 1
        if e["susti"] == "Sí":
            b["susti_si"] += 1
        else:
            b["susti_no"] += 1

    ytd_nuevas = sum(monthly[m]["nuevas"] for m in range(1, num_months + 1))
    ytd_bloq = sum(monthly[m]["bloq"] for m in range(1, num_months + 1))
    parque_ult = parque_for_chain(chain, ventas, num_months, year)
    tasa_ytd = (ytd_bloq / parque_ult / num_months) if (parque_ult and num_months) else 0
    bloq_ano = tasa_ytd * 12

    headers = "".join(f'<th>{MONTH_LABELS_ES[m-1]}</th>' for m in range(1, num_months + 1))
    out = []

    # === Evolución mensual ===
    out.append('<p class="chain-section-title">Evolución mensual <span style="color:var(--grey);font-weight:400">— click en cualquier número para filtrar la tabla de incidencias abajo</span></p>')
    out.append('<table class="evol-table"><thead><tr><th>Métrica</th>')
    out.append(headers)
    out.append('<th class="ytd">YTD</th></tr></thead><tbody>')

    def cell_filter(val, fdict, klass=""):
        if not val or val == "-":
            return f'<td>-</td>'
        return f'<td class="filter-cell {klass}" data-filter=\'{jdata(fdict)}\' title="Click para filtrar">{val}</td>'

    # Nuevas
    cells = [cell_filter(monthly[m]["nuevas"] or 0, {"month": m})
             for m in range(1, num_months + 1)]
    out.append('<tr class="zebra"><td class="lbl">Nuevas incidencias</td>' +
               "".join(cells) +
               f'<td class="filter-cell" data-filter=\'{jdata({})}\' title="Click: todas las del año">{ytd_nuevas}</td></tr>')

    # Bloqueantes
    cells = [cell_filter(monthly[m]["bloq"] or 0, {"month": m, "bloq": "Sí"})
             for m in range(1, num_months + 1)]
    out.append('<tr><td class="lbl">Bloqueantes</td>' +
               "".join(cells) +
               f'<td class="filter-cell" data-filter=\'{jdata({"bloq":"Sí"})}\'>{ytd_bloq}</td></tr>')

    # Parque (no clicable)
    parque_vals = [parque_for_chain(chain, ventas, m, year) for m in range(1, num_months + 1)]
    out.append('<tr class="zebra"><td class="lbl">Total parque</td>' +
               "".join(f'<td>{v}</td>' for v in parque_vals) +
               f'<td>{parque_ult}</td></tr>')

    # Tasa (no clicable)
    tasa_cells = []
    for m in range(1, num_months + 1):
        p_m = parque_vals[m - 1]
        b_m = monthly[m]["bloq"]
        t_m = (b_m / p_m) if p_m else 0
        tasa_cells.append(f'<td>{t_m*100:.2f}%</td>' if t_m else '<td>-</td>')
    out.append('<tr><td class="lbl">Tasa bloqueantes</td>' +
               "".join(tasa_cells) +
               f'<td>{tasa_ytd*100:.2f}%</td></tr>')

    # Bloq/año/eq (no clicable)
    ba_cells = []
    for m in range(1, num_months + 1):
        p_m = parque_vals[m - 1]
        b_m = monthly[m]["bloq"]
        t_m = (b_m / p_m) if p_m else 0
        ba_cells.append(f'<td>{t_m*12:.2f}</td>' if t_m else '<td>-</td>')
    out.append('<tr class="zebra"><td class="lbl">Bloq/año/equipo</td>' +
               "".join(ba_cells) +
               f'<td>{bloq_ano:.2f}</td></tr>')

    # Abiertas hoy (clicable, filtra isopen)
    out.append('<tr class="abiertas"><td class="lbl">Abiertas hoy</td>' +
               ('<td>-</td>' * num_months) +
               f'<td class="filter-cell" data-filter=\'{jdata({"isopen":1})}\' title="Click: solo incidencias abiertas hoy">{abiertas_hoy}</td></tr>')
    out.append('</tbody></table>')

    # === Breakdown tables: motivo / compra / coste / susti ===
    def breakdown(title, key_field, buckets, label_col):
        out.append(f'<p class="chain-section-title">{title}</p>')
        out.append(f'<table class="breakdown-table"><thead><tr><th>{label_col}</th>')
        out.append(headers)
        out.append('<th class="ytd">YTD</th></tr></thead><tbody>')
        idx = 0
        for b in buckets:
            month_vals = [monthly[m][key_field].get(b, 0) for m in range(1, num_months + 1)]
            ytd_v = sum(month_vals)
            if ytd_v == 0:
                continue
            cells = []
            for i, v in enumerate(month_vals):
                if v == 0:
                    cells.append('<td>-</td>')
                else:
                    cells.append(f'<td class="filter-cell" data-filter=\'{jdata({"month": i+1, key_field: b})}\'>{v}</td>')
            klass = 'zebra' if idx % 2 else ''
            out.append(f'<tr class="{klass}"><td class="lbl">{html_escape(b)}</td>' +
                       "".join(cells) +
                       f'<td class="filter-cell" data-filter=\'{jdata({key_field: b})}\'>{ytd_v}</td></tr>')
            idx += 1
        out.append('</tbody></table>')

    breakdown("Motivo avería", "motivo", MOTIVO_BUCKETS, "Motivo")
    breakdown("Fecha de compra del equipo", "yearcompra", YEAR_BUCKETS, "Año compra")
    breakdown("Coste", "coste", COSTE_BUCKETS, "Rango")

    # Sustitución (manual, con bloq Sí/No)
    out.append('<p class="chain-section-title">Sustitución entregada</p>')
    out.append('<table class="breakdown-table"><thead><tr><th>Sustitución</th>')
    out.append(headers)
    out.append('<th class="ytd">YTD</th></tr></thead><tbody>')
    si_cells, no_cells = [], []
    si_ytd, no_ytd = 0, 0
    for m in range(1, num_months + 1):
        s = monthly[m]["susti_si"]
        n = monthly[m]["susti_no"]
        si_ytd += s
        no_ytd += n
        si_cells.append(f'<td class="filter-cell" data-filter=\'{jdata({"month": m, "susti": "Sí"})}\'>{s}</td>' if s else '<td>-</td>')
        no_cells.append(f'<td class="filter-cell" data-filter=\'{jdata({"month": m, "susti": "No"})}\'>{n}</td>' if n else '<td>-</td>')
    out.append('<tr><td class="lbl">Sí entregada</td>' + "".join(si_cells) +
               f'<td class="filter-cell" data-filter=\'{jdata({"susti":"Sí"})}\'>{si_ytd}</td></tr>')
    out.append('<tr class="zebra"><td class="lbl">Sin sustitución</td>' + "".join(no_cells) +
               f'<td class="filter-cell" data-filter=\'{jdata({"susti":"No"})}\'>{no_ytd}</td></tr>')
    out.append('</tbody></table>')

    # === Tabla detalle YTD (TODAS las incidencias del año) ===
    enriched.sort(key=lambda x: (x[1]["month"] or 99, -(days_since(last_status_change_dt(x[0])) or 0)))
    JIRA_URL = "https://leaseir.atlassian.net/browse/"
    out.append(f'<p class="chain-section-title">Incidencias del año ({len(enriched)}) <span style="color:var(--grey);font-weight:400">— <span class="active-filter-info">sin filtro</span> · <button class="clear-chain-filter" type="button" style="font-size:11px;padding:2px 8px;margin-left:6px;border:1px solid var(--line);border-radius:4px;background:white;cursor:pointer;display:none">✕ Limpiar filtro</button></span></p>')
    out.append('<div class="scroller"><table class="ticket-table" style="font-size:11px;min-width:1200px"><thead><tr>')
    for h in ["Ticket", "Cliente / Centro", "Estado", "Gestión", "Días", "Localización",
              "Creada", "Tipo", "Bloq", "Susti", "Motivo", "Año compra", "Coste", "Fecha venta", "Consola", "HP", "Garantía"]:
        out.append(f'<th>{h}</th>')
    out.append('</tr></thead><tbody>')

    for t, e in enriched:
        st = t.get("current_status", "")
        ft = FUNNEL_TAG.get(st, "?")
        ft_color = FUNNEL_COLOR.get(ft, "#888")
        opens = is_open(st)
        days = days_since(last_status_change_dt(t)) if opens else None
        days_color = COLOR_GREEN if (days is not None and days < 5) else (
            COLOR_YELLOW if (days is not None and days <= 15) else COLOR_RED)
        days_html = f'<span style="color:{days_color};font-weight:500">{days}</span>' if days is not None else '<span style="color:var(--grey)">cerrada</span>'
        created_dt = parse_iso(t.get("created"))
        created_s = created_dt.strftime("%d/%m/%Y") if created_dt else ""
        fventa_dt = parse_iso(t.get("fventa") or "")
        fventa_s = fventa_dt.strftime("%d/%m/%Y") if fventa_dt else (t.get("fventa") or "")
        link = f'<a href="{JIRA_URL}{t["key"]}" target="_blank" rel="noopener" style="color:#2a59c4;text-decoration:none">{t["key"]}</a>'
        out.append(
            f'<tr class="ticket-row" '
            f'data-month="{e["month"]}" '
            f'data-bloq="{html_escape(e["bloq"])}" '
            f'data-susti="{html_escape(e["susti"])}" '
            f'data-motivo="{html_escape(e["motivo"])}" '
            f'data-yearcompra="{html_escape(e["yearcompra"])}" '
            f'data-coste="{html_escape(e["coste"])}" '
            f'data-isopen="{e["isopen"]}">'
            f'<td style="text-align:left">{link}</td>'
            f'<td class="d-trunc" style="text-align:left" title="{html_escape(t.get("cliente",""))}">{html_escape(t.get("cliente",""))}</td>'
            f'<td style="text-align:left">{html_escape(st)}</td>'
            f'<td style="color:{ft_color};font-weight:500">{ft}</td>'
            f'<td>{days_html}</td>'
            f'<td class="d-trunc" style="text-align:left" title="{html_escape(t.get("loc",""))}">{html_escape(t.get("loc",""))}</td>'
            f'<td>{created_s}</td>'
            f'<td style="text-align:left">{html_escape(t.get("tipo",""))}</td>'
            f'<td>{html_escape(e["bloq"])}</td>'
            f'<td>{html_escape(e["susti"])}</td>'
            f'<td>{html_escape(e["motivo"])}</td>'
            f'<td>{html_escape(e["yearcompra"])}</td>'
            f'<td>{html_escape(e["coste"])}</td>'
            f'<td>{fventa_s}</td>'
            f'<td>{html_escape(t.get("consola",""))}</td>'
            f'<td>{html_escape(t.get("hp",""))}</td>'
            f'<td>{html_escape(t.get("garantia",""))}</td>'
            '</tr>'
        )
    out.append('</tbody></table></div>')

    return "".join(out), {
        "nuevas": ytd_nuevas, "bloq": ytd_bloq, "parque": parque_ult,
        "tasa": tasa_ytd, "bloq_ano": bloq_ano, "abiertas": abiertas_hoy,
    }


def build_resumen_pane(stats):
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
    tasa_global = (totals["bloq"] / totals["parque"]) if totals["parque"] else 0
    out.append(
        f'<tr style="background:#fff4d6;font-weight:500"><td class="lbl">Total</td>'
        f'<td>{totals["nuevas"]}</td><td>{totals["bloq"]}</td><td>{totals["parque"]}</td>'
        f'<td>{tasa_global*100:.2f}%</td><td>—</td><td>{totals["abiertas"]}</td></tr>'
    )
    out.append('</tbody></table>')
    return "".join(out)


def build_chains_html(cache, airtable, serial_year):
    today = date.today()
    year = today.year
    num_months = today.month
    max_known = compute_max_known_per_prefix(serial_year)

    tickets = cache.get("tickets", {})
    by_chain = {ch: [] for ch in CHAIN_ORDER}
    for k, t in tickets.items():
        ch = chain_of_ticket(t)
        if ch not in by_chain:
            by_chain[ch] = []
        by_chain[ch].append(t)

    ventas = (airtable or {}).get("ventas_by_chain_month", {})
    chains_to_show = [ch for ch in CHAIN_ORDER if ch != "Otros"]

    tabs = ['<div class="tabs2">']
    tabs.append('<button type="button" data-chain="resumen" class="active">Resumen</button>')
    for ch in chains_to_show:
        tabs.append(f'<button type="button" data-chain="{html_escape(ch)}">{html_escape(ch)}</button>')
    tabs.append('</div>')

    stats_per_chain = {}
    panes = []
    for ch in chains_to_show:
        pane_html, s = build_chain_pane(ch, by_chain.get(ch, []), ventas,
                                        serial_year, max_known, year, num_months)
        stats_per_chain[ch] = s
        meta = (f'{s["nuevas"]} nuevas YTD · {s["bloq"]} bloqueantes · '
                f'{s["abiertas"]} abiertas hoy')
        panes.append(
            f'<div class="chain-pane" data-chain="{html_escape(ch)}">'
            f'<h3>{html_escape(ch)}</h3>'
            f'<p class="chain-meta">{meta}</p>'
            f'{pane_html}</div>'
        )

    resumen_html = (
        '<div class="chain-pane active" data-chain="resumen">'
        '<h3>Resumen por cadena</h3>'
        '<p class="chain-meta">YTD · 5 métricas estándar (mismas del email a Ignacio) + abiertas hoy</p>'
        + build_resumen_pane(stats_per_chain)
        + '</div>'
    )

    return "".join(tabs) + resumen_html + "".join(panes)
