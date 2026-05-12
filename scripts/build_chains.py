"""build_chains.py — Fase 2 iter E2: Inmovilizado SAT con 4 secciones + colores."""
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

# Colores del brief
INMOV_RED = "#FFD1D1"     # solo en Jira
INMOV_YELLOW = "#FFF2A8"  # duplicado
INMOV_GREEN = "#D9F2D6"   # disponible (solo en AT con SAT)
INMOV_HDR_DARK = "#1F3B5E"
INMOV_HDR_BLUE = "#2E5496"


def html_escape(s):
    if s is None:
        return ""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def jdata(d):
    return html_escape(_json.dumps(d, ensure_ascii=False))


def add_serial_variants(s, variants):
    if s is None:
        return
    s = str(s).strip()
    if not s or s.lower() in ('0', '0.0', 'no aparece', 'none', 'null', ''):
        return
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    variants.add(s); variants.add(s.lower()); variants.add(s.upper())
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
    if y is None: return "Sin compra"
    if y >= 2026: return "2026"
    if y in (2025, 2024, 2023, 2022): return str(y)
    return "<2022"


def bucket_motivo(motivos, status):
    cerrados = ("Finalizada", "Resuelto", "Cancelado", "Finalizado técnico externo")
    if not motivos:
        if status in cerrados: return "Otros"
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
    if v == 0: return "Sin coste"
    if v < 1000: return "< 1.000 €"
    if v <= 6000: return "1.000 - 6.000 €"
    return "> 6.000 €"


def chain_of_ticket(t):
    return chain_of(t.get("cliente", ""), t.get("loc", ""))


def compute_max_known_per_prefix(serial_year):
    max_n = {"C": 0, "H": 0}
    for k in serial_year.keys():
        m = re.match(r'^([CH])(\d+)$', k)
        if m:
            p = m.group(1); n = int(m.group(2))
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


# ========================
# CRUCE INMOVILIZADO SAT
# ========================

def normalize_serial(s):
    """Normaliza serial para matching: upper, strip, sin sufijo BIS."""
    if s is None: return ""
    s = str(s).strip().upper()
    s = re.sub(r'\s*BIS\s*$', '', s)
    # Strip .0 si viene como float string
    if s.endswith('.0') and s[:-2].replace('-', '').isdigit():
        s = s[:-2]
    return s


def get_serial_aliases(inmov_item):
    """Devuelve set de aliases para un equipo de Airtable (ID + Renombrada)."""
    aliases = set()
    sid = normalize_serial(inmov_item.get("serial", ""))
    if sid:
        aliases.add(sid)
    renombrada = inmov_item.get("renombrada", "") or ""
    for a in re.split(r'[,;\n/]', renombrada):
        a = normalize_serial(a)
        if a and a not in ("IDEAL", "CU", "—", "BAJA", "BIS"):
            aliases.add(a)
    return aliases


def section_for_inmov(item):
    """A/B/C/D según Type of Asset + Console + Spot Size."""
    spot = (item.get("spot_size", "") or "").upper()
    type_a = (item.get("type_asset", "") or "").lower()
    console = (item.get("console_model", "") or "").upper()
    if "SRF" in spot:
        return "D"
    if "console" in type_a or "consola" in type_a:
        if console == "AHR":
            return "C"
        return "A"
    if "handpiece" in type_a or "manípulo" in type_a or "manipulo" in type_a:
        return "B"
    return "A"


def section_for_solo_jira(serial):
    """Cuando solo_jira (sin match en Airtable), infiere sección por prefijo del serial."""
    s = (serial or "").upper().strip()
    if not s: return "A"
    if s.startswith("SUSTSRF") or "SRF" in s: return "D"
    if s.startswith("SUSTMHR"): return "A"
    if s.startswith("SUSTDUAL") or s.startswith("SUSTQUAD") or s.startswith("SUSTSINGLE"): return "B"
    if re.match(r'^C\d', s): return "A"
    if re.match(r'^H\d', s): return "B"
    if re.match(r'^9\d{4}$', s): return "A"
    if re.match(r'^(30|40)\d{3}$', s): return "B"
    if s == "DEMO": return "A"
    return "A"


def chain_of_inmov(customer):
    """Mapea Customer de Airtable Inmovilizado a una cadena."""
    return chain_of(customer or "", "")


def chain_of_sustis(s):
    """Mapea sustitución (Jira) a cadena."""
    return chain_of(s.get("cliente", "") or s.get("parent_cliente", ""), s.get("loc", ""))


def build_inmovilizado_section(chain, sustis_items, inmov_items):
    """Genera el HTML de la tabla Inmovilizado SAT para una cadena.
    Devuelve "" si no hay nada que mostrar."""
    # Filtrar sustis e inmov por cadena
    sustis_chain = [s for s in sustis_items if chain_of_sustis(s) == chain]
    inmov_chain = [i for i in inmov_items if chain_of_inmov(i.get("customer", "")) == chain]

    # Construir lookup de Airtable por serial alias
    alias_to_inmov = {}
    for it in inmov_chain:
        for a in get_serial_aliases(it):
            alias_to_inmov[a] = it
    # También global (no filtrado por cadena) para resolver matches que crucen
    alias_to_inmov_global = {}
    for it in inmov_items:
        for a in get_serial_aliases(it):
            alias_to_inmov_global.setdefault(a, it)

    # Por cada sub-task Jira, generar fila por serial prestado (consola y/o manípulo)
    rows = []  # cada fila: dict con todos los campos + kind + color + section
    matched_inmov_ids = set()
    now_utc = datetime.now(timezone.utc)

    def make_row_from_sustis(s, serial, role):
        n_serial = normalize_serial(serial)
        if not n_serial:
            return None
        inmov_match = alias_to_inmov_global.get(n_serial)
        is_match = inmov_match is not None
        if is_match:
            matched_inmov_ids.add(inmov_match.get("id_rec"))
            modelo = inmov_match.get("console_model", "") or inmov_match.get("type_asset", "")
            activity = inmov_match.get("activity", "")
            section = section_for_inmov(inmov_match)
        else:
            modelo = ""
            activity = ""
            section = section_for_solo_jira(serial)
        fe_dt = parse_iso(s.get("fecha_envio"))
        dias = (now_utc - fe_dt).days if fe_dt else None
        return {
            "num_serie": serial,
            "modelo": modelo,
            "cliente": s.get("cliente", ""),
            "subtask_key": s.get("key", ""),
            "parent_key": s.get("parent_key", ""),
            "parent_status": s.get("parent_status", ""),
            "consola_avr": s.get("consola_averiada", ""),
            "hp_avr": s.get("manipulo_averiado", ""),
            "subtask_status": s.get("subtask_status", ""),
            "loc": s.get("loc", ""),
            "fecha_envio": s.get("fecha_envio"),
            "fecha_envio_fmt": fe_dt.strftime("%d/%m/%Y %H:%M") if fe_dt else (s.get("fecha_envio") or ""),
            "dias": dias,
            "activity": activity,
            "kind": "union" if is_match else "solo_jira",
            "section": section,
        }

    for s in sustis_chain:
        for serial, role in [(s.get("consola_susti"), "consola"), (s.get("manipulo_susti"), "manipulo")]:
            r = make_row_from_sustis(s, serial, role)
            if r:
                rows.append(r)

    # Detectar duplicados (mismo serial en múltiples filas)
    serial_counts = {}
    for r in rows:
        n = normalize_serial(r["num_serie"])
        serial_counts.setdefault(n, []).append(r)
    for serials, group in serial_counts.items():
        if len(group) > 1:
            # marcar todas menos la más reciente como duplicado amarillo
            group.sort(key=lambda r: r["dias"] if r["dias"] is not None else -1)
            # El de menos días (más reciente) es el actual; los demás son antiguos
            for r in group[:-1]:
                r["kind"] = "duplicado"

    # Añadir equipos Airtable SAT que NO están en ninguna sub-task (disponibles)
    for it in inmov_chain:
        if it.get("activity") != "SAT": continue
        if it.get("id_rec") in matched_inmov_ids: continue
        section = section_for_inmov(it)
        rows.append({
            "num_serie": it.get("serial", ""),
            "modelo": it.get("console_model", "") or it.get("type_asset", ""),
            "cliente": "—",
            "subtask_key": "", "parent_key": "", "parent_status": "",
            "consola_avr": "", "hp_avr": "",
            "subtask_status": "", "loc": "",
            "fecha_envio": None, "fecha_envio_fmt": "",
            "dias": None,
            "activity": "SAT",
            "kind": "solo_at",
            "section": section,
        })

    if not rows:
        return ""

    # Orden dentro de cada sección
    def sort_key(r):
        is_avail = (r["kind"] == "solo_at")
        days = r["dias"] if isinstance(r["dias"], int) else -1
        return (1 if is_avail else 0, -days, r["num_serie"] or "")
    rows.sort(key=sort_key)

    # Agrupar por sección
    by_section = {"A": [], "B": [], "C": [], "D": []}
    for r in rows:
        by_section[r["section"]].append(r)

    section_titles = {
        "A": "A — Consolas (MHR y otras, excl. AHR y SRF)",
        "B": "B — Manípulos (Quad / Dual / Single, excl. SRF)",
        "C": "C — Consolas AHR",
        "D": "D — Tecnología SRF (consolas + manípulos)",
    }
    color_map = {
        "solo_jira": INMOV_RED,
        "duplicado": INMOV_YELLOW,
        "solo_at": INMOV_GREEN,
    }

    JIRA_URL = "https://leaseir.atlassian.net/browse/"
    out = []
    today = date.today().isoformat()
    out.append(f'<p class="chain-section-title">Inmovilizado SAT — {today} <span style="color:var(--grey);font-weight:400">— {len(rows)} equipos en cadena {html_escape(chain)}</span></p>')
    out.append('<p class="legend">Leyenda: '
               f'<span class="pill" style="background:{INMOV_RED};padding:2px 8px;border-radius:3px">Solo en Jira</span> · '
               f'<span class="pill" style="background:{INMOV_YELLOW};padding:2px 8px;border-radius:3px">Duplicado antiguo</span> · '
               f'<span class="pill" style="background:{INMOV_GREEN};padding:2px 8px;border-radius:3px">Disponible (SAT sin ticket activo)</span> · '
               'normal: en ambos</p>')
    out.append('<div class="scroller"><table style="font-size:11px;min-width:1700px">')

    headers = ["Num. Serie", "Modelo", "Cliente",
               "Incidencia (sub-task)", "Incidencia principal", "Estado principal",
               "Consola principal (avería)", "HP principal (avería)",
               "Estado Jira (sub-task)", "Localización equipo averiado",
               "Fecha y hora de envío", "Días con sustitución",
               "Current Activity (Airtable)"]
    NCOLS = len(headers)

    for sec_key in ["A", "B", "C", "D"]:
        rows_s = by_section[sec_key]
        if not rows_s:
            continue
        out.append(f'<tr><th colspan="{NCOLS}" style="background:{INMOV_HDR_DARK};color:white;text-align:left;padding:6px 10px">{section_titles[sec_key]}</th></tr>')
        out.append('<tr>')
        for h in headers:
            out.append(f'<th style="background:{INMOV_HDR_BLUE};color:white;padding:6px 10px">{h}</th>')
        out.append('</tr>')
        for r in rows_s:
            bg = color_map.get(r["kind"], "")
            style = f' style="background:{bg}"' if bg else ""
            sub_link = f'<a href="{JIRA_URL}{r["subtask_key"]}" target="_blank" rel="noopener" style="color:#2a59c4;text-decoration:none">{r["subtask_key"]}</a>' if r["subtask_key"] else "—"
            par_link = f'<a href="{JIRA_URL}{r["parent_key"]}" target="_blank" rel="noopener" style="color:#2a59c4;text-decoration:none">{r["parent_key"]}</a>' if r["parent_key"] else "—"
            dias_v = r["dias"] if r["dias"] is not None else "—"
            out.append(f'<tr{style}>')
            out.append(f'<td style="text-align:center"><b>{html_escape(r["num_serie"])}</b></td>')
            out.append(f'<td>{html_escape(r["modelo"])}</td>')
            out.append(f'<td style="text-align:left">{html_escape(r["cliente"])}</td>')
            out.append(f'<td style="text-align:center">{sub_link}</td>')
            out.append(f'<td style="text-align:center">{par_link}</td>')
            out.append(f'<td>{html_escape(r["parent_status"])}</td>')
            out.append(f'<td style="text-align:center">{html_escape(r["consola_avr"])}</td>')
            out.append(f'<td style="text-align:center">{html_escape(r["hp_avr"])}</td>')
            out.append(f'<td>{html_escape(r["subtask_status"])}</td>')
            out.append(f'<td class="d-trunc" style="text-align:left" title="{html_escape(r["loc"])}">{html_escape(r["loc"])}</td>')
            out.append(f'<td>{html_escape(r["fecha_envio_fmt"])}</td>')
            out.append(f'<td style="text-align:center">{dias_v}</td>')
            out.append(f'<td>{html_escape(r["activity"])}</td>')
            out.append('</tr>')

    out.append('</table></div>')
    return "".join(out)


def build_chain_pane(chain, tickets_chain, ventas, serial_year, max_known,
                     year, num_months, sustis_items=None, inmov_items=None):
    enriched = []
    for t in tickets_chain:
        e = enrich(t, serial_year, max_known, year)
        if e["month"] is not None:
            enriched.append((t, e))

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
        if e["bloq"] == "Sí": b["bloq"] += 1
        b["motivo"][e["motivo"]] = b["motivo"].get(e["motivo"], 0) + 1
        b["yearcompra"][e["yearcompra"]] = b["yearcompra"].get(e["yearcompra"], 0) + 1
        b["coste"][e["coste"]] = b["coste"].get(e["coste"], 0) + 1
        if e["susti"] == "Sí": b["susti_si"] += 1
        else: b["susti_no"] += 1

    ytd_nuevas = sum(monthly[m]["nuevas"] for m in range(1, num_months + 1))
    ytd_bloq = sum(monthly[m]["bloq"] for m in range(1, num_months + 1))
    parque_ult = parque_for_chain(chain, ventas, num_months, year)
    tasa_ytd = (ytd_bloq / parque_ult / num_months) if (parque_ult and num_months) else 0
    bloq_ano = tasa_ytd * 12

    headers = "".join(f'<th>{MONTH_LABELS_ES[m-1]}</th>' for m in range(1, num_months + 1))
    out = []

    out.append('<p class="chain-section-title">Evolución mensual <span style="color:var(--grey);font-weight:400">— click en cualquier número para filtrar la tabla de incidencias abajo</span></p>')
    out.append('<table class="evol-table"><thead><tr><th>Métrica</th>' + headers + '<th class="ytd">YTD</th></tr></thead><tbody>')

    def cell_filter(val, fdict):
        if not val: return '<td>-</td>'
        return f'<td class="filter-cell" data-filter=\'{jdata(fdict)}\' title="Click para filtrar">{val}</td>'

    cells = [cell_filter(monthly[m]["nuevas"], {"month": m}) for m in range(1, num_months + 1)]
    out.append('<tr class="zebra"><td class="lbl">Nuevas incidencias</td>' + "".join(cells) +
               f'<td class="filter-cell" data-filter=\'{jdata({})}\'>{ytd_nuevas}</td></tr>')

    cells = [cell_filter(monthly[m]["bloq"], {"month": m, "bloq": "Sí"}) for m in range(1, num_months + 1)]
    out.append('<tr><td class="lbl">Bloqueantes</td>' + "".join(cells) +
               f'<td class="filter-cell" data-filter=\'{jdata({"bloq":"Sí"})}\'>{ytd_bloq}</td></tr>')

    parque_vals = [parque_for_chain(chain, ventas, m, year) for m in range(1, num_months + 1)]
    out.append('<tr class="zebra"><td class="lbl">Total parque</td>' +
               "".join(f'<td>{v}</td>' for v in parque_vals) + f'<td>{parque_ult}</td></tr>')

    tasa_cells = []
    for m in range(1, num_months + 1):
        p_m = parque_vals[m - 1]; b_m = monthly[m]["bloq"]
        t_m = (b_m / p_m) if p_m else 0
        tasa_cells.append(f'<td>{t_m*100:.2f}%</td>' if t_m else '<td>-</td>')
    out.append('<tr><td class="lbl">Tasa bloqueantes</td>' + "".join(tasa_cells) +
               f'<td>{tasa_ytd*100:.2f}%</td></tr>')

    ba_cells = []
    for m in range(1, num_months + 1):
        p_m = parque_vals[m - 1]; b_m = monthly[m]["bloq"]
        t_m = (b_m / p_m) if p_m else 0
        ba_cells.append(f'<td>{t_m*12:.2f}</td>' if t_m else '<td>-</td>')
    out.append('<tr class="zebra"><td class="lbl">Bloq/año/equipo</td>' + "".join(ba_cells) +
               f'<td>{bloq_ano:.2f}</td></tr>')

    out.append('<tr class="abiertas"><td class="lbl">Abiertas hoy</td>' +
               ('<td>-</td>' * num_months) +
               f'<td class="filter-cell" data-filter=\'{jdata({"isopen":1})}\'>{abiertas_hoy}</td></tr>')
    out.append('</tbody></table>')

    def breakdown(title, key_field, buckets, label_col):
        out.append(f'<p class="chain-section-title">{title}</p>')
        out.append(f'<table class="breakdown-table"><thead><tr><th>{label_col}</th>' + headers + '<th class="ytd">YTD</th></tr></thead><tbody>')
        idx = 0
        for b in buckets:
            mvals = [monthly[m][key_field].get(b, 0) for m in range(1, num_months + 1)]
            ytd_v = sum(mvals)
            if ytd_v == 0: continue
            cs = []
            for i, v in enumerate(mvals):
                if v == 0: cs.append('<td>-</td>')
                else: cs.append(f'<td class="filter-cell" data-filter=\'{jdata({"month": i+1, key_field: b})}\'>{v}</td>')
            klass = 'zebra' if idx % 2 else ''
            out.append(f'<tr class="{klass}"><td class="lbl">{html_escape(b)}</td>' + "".join(cs) +
                       f'<td class="filter-cell" data-filter=\'{jdata({key_field: b})}\'>{ytd_v}</td></tr>')
            idx += 1
        out.append('</tbody></table>')

    breakdown("Motivo avería", "motivo", MOTIVO_BUCKETS, "Motivo")
    breakdown("Fecha de compra del equipo", "yearcompra", YEAR_BUCKETS, "Año compra")
    breakdown("Coste", "coste", COSTE_BUCKETS, "Rango")

    out.append('<p class="chain-section-title">Sustitución entregada</p>')
    out.append('<table class="breakdown-table"><thead><tr><th>Sustitución</th>' + headers + '<th class="ytd">YTD</th></tr></thead><tbody>')
    si_cells, no_cells = [], []
    si_ytd, no_ytd = 0, 0
    for m in range(1, num_months + 1):
        s = monthly[m]["susti_si"]; n = monthly[m]["susti_no"]
        si_ytd += s; no_ytd += n
        si_cells.append(f'<td class="filter-cell" data-filter=\'{jdata({"month": m, "susti": "Sí"})}\'>{s}</td>' if s else '<td>-</td>')
        no_cells.append(f'<td class="filter-cell" data-filter=\'{jdata({"month": m, "susti": "No"})}\'>{n}</td>' if n else '<td>-</td>')
    out.append('<tr><td class="lbl">Sí entregada</td>' + "".join(si_cells) +
               f'<td class="filter-cell" data-filter=\'{jdata({"susti":"Sí"})}\'>{si_ytd}</td></tr>')
    out.append('<tr class="zebra"><td class="lbl">Sin sustitución</td>' + "".join(no_cells) +
               f'<td class="filter-cell" data-filter=\'{jdata({"susti":"No"})}\'>{no_ytd}</td></tr>')
    out.append('</tbody></table>')

    # === Tabla Inmovilizado SAT con 4 secciones + colores ===
    inmov_html = build_inmovilizado_section(chain, sustis_items or [], inmov_items or [])
    if inmov_html:
        out.append(inmov_html)

    enriched.sort(key=lambda x: (x[1]["month"] or 99, -(days_since(last_status_change_dt(x[0])) or 0)))
    JIRA_URL = "https://leaseir.atlassian.net/browse/"
    out.append(f'<p class="chain-section-title">Incidencias del año ({len(enriched)}) <span style="color:var(--grey);font-weight:400">— <span class="active-filter-info">sin filtro</span> · <button class="clear-chain-filter" type="button" style="font-size:11px;padding:2px 8px;margin-left:6px;border:1px solid var(--line);border-radius:4px;background:white;cursor:pointer;display:none">✕ Limpiar filtro</button></span></p>')

    states_in_chain = sorted({t.get("current_status", "") for t, _ in enriched if t.get("current_status")})
    state_checkboxes = "".join(f'<label><input type="checkbox" class="cf-estado-cb" data-status="{html_escape(s)}"> {html_escape(s)}</label>' for s in states_in_chain)
    out.append('<div class="chain-toolbar">')
    out.append('<input class="cf-search" type="text" placeholder="Buscar texto...">')
    out.append('<div class="multi-wrap"><button type="button" class="multi-btn cf-estado-btn">Estado: todos &#9662;</button>')
    out.append(f'<div class="multi-pop cf-estado-pop"><label class="all-row"><input type="checkbox" class="cf-estado-all" checked> Todos</label><div class="cf-estado-list">{state_checkboxes}</div></div></div>')
    out.append('<select class="cf-gestion"><option value="">Gestión: toda</option><option value="Inicio">Inicio</option><option value="Interna">Interna</option><option value="Online">Online</option><option value="Externa">Externa</option></select>')
    out.append('<select class="cf-bloq"><option value="">Bloq: todos</option><option value="Sí">Sí</option><option value="No">No</option></select>')
    out.append('<select class="cf-susti"><option value="">Susti: todas</option><option value="Sí">Sí</option><option value="No">No</option></select>')
    out.append('<select class="cf-garantia"><option value="">Garantía: todas</option><option value="Sí">Sí</option><option value="No">No</option></select>')
    out.append('<select class="cf-dias"><option value="">Días: todos</option><option value="g">verde &lt;5d</option><option value="y">amarillo 5-15d</option><option value="r">rojo &gt;15d</option><option value="x">cerradas</option></select>')
    out.append('<button class="cf-clear" type="button">✕ Limpiar todo</button>')
    out.append('</div>')

    out.append('<div class="scroller"><table class="ticket-table" style="font-size:11px;min-width:1300px"><thead><tr>')
    hdrs = [("Ticket","text"),("Cliente / Centro","text"),("Estado","text"),("Gestión","text"),
            ("Días","num"),("Localización","text"),("Creada","date"),("Tipo avería","text"),
            ("Bloq","text"),("Susti","text"),("Fecha venta","date"),("Consola","text"),
            ("HP","text"),("Garantía","text"),("Motivo","text"),("Año compra","yearcompra"),
            ("Coste","coste")]
    for i, (h, dtype) in enumerate(hdrs):
        out.append(f'<th class="sortable" data-col="{i}" data-type="{dtype}">{h}</th>')
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
            f'data-isopen="{e["isopen"]}" '
            f'data-estado="{html_escape(st)}" '
            f'data-gestion="{ft}" '
            f'data-garantia="{html_escape(t.get("garantia","") or "No")}" '
            f'data-days="{days if days is not None else ""}">'
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
            f'<td>{fventa_s}</td>'
            f'<td>{html_escape(t.get("consola",""))}</td>'
            f'<td>{html_escape(t.get("hp",""))}</td>'
            f'<td>{html_escape(t.get("garantia","") or "No")}</td>'
            f'<td>{html_escape(e["motivo"])}</td>'
            f'<td>{html_escape(e["yearcompra"])}</td>'
            f'<td>{html_escape(e["coste"])}</td>'
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
        p = s.get("parque", 0); ta = s.get("tasa", 0); ba = s.get("bloq_ano", 0); ab = s.get("abiertas", 0)
        totals["nuevas"] += nu; totals["bloq"] += bl; totals["parque"] += p; totals["abiertas"] += ab
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


def build_chains_html(cache, airtable, serial_year, sustis=None, inmov=None):
    today = date.today()
    year = today.year
    num_months = today.month
    max_known = compute_max_known_per_prefix(serial_year)

    tickets = cache.get("tickets", {})
    by_chain = {ch: [] for ch in CHAIN_ORDER}
    for k, t in tickets.items():
        ch = chain_of_ticket(t)
        if ch not in by_chain: by_chain[ch] = []
        by_chain[ch].append(t)

    ventas = (airtable or {}).get("ventas_by_chain_month", {})
    chains_to_show = [ch for ch in CHAIN_ORDER if ch != "Otros"]

    sustis_items = (sustis or {}).get("items", []) if sustis else []
    inmov_items = (inmov or {}).get("items", []) if inmov else []

    tabs = ['<div class="tabs2">']
    tabs.append('<button type="button" data-chain="resumen" class="active">Resumen</button>')
    for ch in chains_to_show:
        tabs.append(f'<button type="button" data-chain="{html_escape(ch)}">{html_escape(ch)}</button>')
    tabs.append('</div>')

    stats_per_chain = {}
    panes = []
    for ch in chains_to_show:
        pane_html, s = build_chain_pane(ch, by_chain.get(ch, []), ventas,
                                        serial_year, max_known, year, num_months,
                                        sustis_items=sustis_items, inmov_items=inmov_items)
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
