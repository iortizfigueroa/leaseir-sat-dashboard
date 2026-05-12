"""Build report (Excel + HTML) from cached timelines + extended fields."""
from __future__ import annotations

import argparse
import json
import re
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


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

FUNNEL_ORDER = [
    "Abierto", "Recepcionado SAT", "Pendiente recogida", "Gestionado transporte",
    "Pendiente asignar técnico", "En cola taller", "En preparación presupuesto",
    "Presupuesto preparado pendiente de enviar", "Pendiente confirmación presupuesto",
    "Esperando inicio reparación", "En reparación", "Inspección de salida",
    "Devuelto a cliente", "Pendiente agendar llamada", "Pendiente definir servicio externo",
    "Esperando respuesta cliente a presupuesto", "Enviado a técnico externo",
]

FUNNEL_TAG = {
    "Abierto": "Inicio", "Recepcionado SAT": "Interna", "Pendiente recogida": "Interna",
    "Gestionado transporte": "Interna", "Pendiente asignar técnico": "Interna",
    "En cola taller": "Interna", "En preparación presupuesto": "Interna",
    "Presupuesto preparado pendiente de enviar": "Interna",
    "Pendiente confirmación presupuesto": "Interna",
    "Esperando inicio reparación": "Interna", "En reparación": "Interna",
    "Inspección de salida": "Interna", "Devuelto a cliente": "Interna",
    "Pendiente agendar llamada": "Online",
    "Pendiente definir servicio externo": "Externa",
    "Esperando respuesta cliente a presupuesto": "Externa",
    "Enviado a técnico externo": "Externa",
}

CHAIN_ORDER = ["Elha", "Sin Vello", "Dermasana", "Smart Duck",
               "Epil Point", "Laser Factory", "Unico Italia", "Otros"]

FUNNEL_COLOR = {"Inicio": "#9aa5b1", "Interna": "#1f6feb", "Online": "#cb6f0a", "Externa": "#a23b72"}
COLOR_GREEN = "#1f8a4c"
COLOR_YELLOW = "#b8860b"
COLOR_RED = "#c0392b"


def chain_of(cliente, loc):
    c = (cliente or "").lower()
    l = (loc or "").lower()
    if "elha" in c: return "Elha"
    if "sin vello" in c or "sinvello" in c: return "Sin Vello"
    if "dermasana" in c: return "Dermasana"
    if "smart duck" in c or "smartduck" in c: return "Smart Duck"
    if "epil point" in c or "epilpoint" in c: return "Epil Point"
    if "laser factory" in c or "laserfactory" in c or "laser fcatory" in c: return "Laser Factory"
    if any(t in c for t in ("centro unico", "centri unico", "centros unico", "beauty cool")):
        ital = ("italia", "italy", "roma", "milano", "napoli", "torino", "bologna",
                "firenze", "ostia", "orio", "giugliano", "andria", "castellammare",
                "piacenza", "novara", "aprilia", "euroma")
        if any(ct in l for ct in ital):
            return "Unico Italia"
        return "Otros"
    return "Otros"


def parse_iso(s):
    if not s:
        return None
    s2 = s.replace("Z", "+00:00")
    s2 = re.sub(r"([+\-]\d{2})(\d{2})$", r"\1:\2", s2)
    try:
        return datetime.fromisoformat(s2)
    except ValueError:
        return None


def status_at(ticket, cutoff_dt):
    created = parse_iso(ticket.get("created"))
    if not created or created > cutoff_dt:
        return None
    transitions = ticket.get("transitions", [])
    if not transitions:
        return ticket.get("current_status")
    current = transitions[0][1]
    for ts, from_s, to_s in transitions:
        when = parse_iso(ts)
        if when and when <= cutoff_dt:
            current = to_s
        else:
            break
    return current


def last_status_change_dt(ticket):
    transitions = ticket.get("transitions", [])
    if transitions:
        last = transitions[-1]
        d = parse_iso(last[0])
        if d:
            return d
    return parse_iso(ticket.get("created"))


def is_open(s):
    return s in OPEN_STATUSES


def compute_cutoffs(today):
    """Devuelve hasta 21 cortes únicos: 6 fin-de-mes + 15 días laborables.
    Si un fin-de-mes cae en uno de los 15 días laborables, se conserva una sola vez."""
    cutoffs = []
    seen = set()
    y, m = today.year, today.month
    monthly = []
    for _ in range(6):
        m -= 1
        if m == 0:
            m, y = 12, y - 1
        nm, ny = (m + 1, y) if m < 12 else (1, y + 1)
        monthly.append(date(ny, nm, 1) - timedelta(days=1))
    monthly.sort()
    for d in monthly:
        lbl = d.strftime("%Y-%m-%d")
        if lbl in seen:
            continue
        seen.add(lbl)
        cutoffs.append((lbl, d))
    bd, d = [], today
    while len(bd) < 15:
        if d.weekday() < 5:
            bd.append(d)
        d -= timedelta(days=1)
    bd.sort()
    for x in bd:
        lbl = x.strftime("%Y-%m-%d")
        if lbl in seen:
            continue
        seen.add(lbl)
        cutoffs.append((lbl, x))
    return cutoffs


def replay_opens(cache, cutoffs):
    opens = {lbl: [] for lbl, _ in cutoffs}
    for key, t in cache.get("tickets", {}).items():
        chain = chain_of(t.get("cliente", ""), t.get("loc", ""))
        for lbl, d in cutoffs:
            cdt = datetime(d.year, d.month, d.day, 23, 59, 59, tzinfo=timezone.utc)
            s = status_at(t, cdt)
            if is_open(s):
                opens[lbl].append((key, chain, s))
    return opens


def fmt_short(lbl):
    try:
        dt = datetime.fromisoformat(lbl)
        months = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
        return f"{dt.day:02d} {months[dt.month-1]}"
    except Exception:
        return lbl


def days_since(dt):
    if not dt:
        return None
    now = datetime.now(timezone.utc)
    return (now - dt).days


def days_bucket(d):
    if d is None:
        return "y"
    if d < 5:
        return "g"
    if d <= 15:
        return "y"
    return "r"


FONT_X = "Arial"
HDR_BLUE = "0B3D91"
ZEBRA = "F5F7FB"
TOTAL_BG = "FFF4D6"
FUNNEL_C_BG = "F0EAFB"
thin = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def hdr_cell(c, t):
    c.value = t
    c.font = Font(name=FONT_X, bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", start_color=HDR_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER


def lbl(c, t, bold=False, fill=None):
    c.value = t
    c.font = Font(name=FONT_X, bold=bold, size=10)
    if fill:
        c.fill = PatternFill("solid", start_color=fill)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = BORDER


def num(c, v, bold=False, fill=None):
    c.value = v
    c.font = Font(name=FONT_X, bold=bold, size=10)
    if fill:
        c.fill = PatternFill("solid", start_color=fill)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDER


def build_excel(cache, out_path):
    today = date.today()
    cutoffs = compute_cutoffs(today)
    opens = replay_opens(cache, cutoffs)

    wb = Workbook()
    ws = wb.active
    ws.title = "Evol_Estado"
    ws["A1"] = f"Evolución abiertas SAT por estado — {today.isoformat()}"
    ws["A1"].font = Font(name=FONT_X, bold=True, size=13, color=HDR_BLUE)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cutoffs) + 2)

    hdr_cell(ws.cell(row=3, column=1), "Funnel")
    hdr_cell(ws.cell(row=3, column=2), "Estado")
    for j, (l, _) in enumerate(cutoffs, 3):
        hdr_cell(ws.cell(row=3, column=j), l)

    present = set()
    for l, _ in cutoffs:
        for _, _, s in opens[l]:
            if s:
                present.add(s)
    states = [s for s in FUNNEL_ORDER if s in present] + sorted(s for s in present if s not in FUNNEL_ORDER)

    row = 4
    for st in states:
        ft = FUNNEL_TAG.get(st, "?")
        f = FUNNEL_C_BG if ft == "C" else (ZEBRA if (row % 2 == 0) else None)
        lbl(ws.cell(row=row, column=1), ft, fill=f)
        lbl(ws.cell(row=row, column=2), st, fill=f)
        for j, (l, _) in enumerate(cutoffs, 3):
            c = sum(1 for _, _, s in opens[l] if s == st)
            num(ws.cell(row=row, column=j), c if c else "", fill=f)
        row += 1
    lbl(ws.cell(row=row, column=1), "", bold=True, fill=TOTAL_BG)
    lbl(ws.cell(row=row, column=2), "Total abiertas", bold=True, fill=TOTAL_BG)
    for j, (l, _) in enumerate(cutoffs, 3):
        num(ws.cell(row=row, column=j), len(opens[l]), bold=True, fill=TOTAL_BG)

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 38
    for j in range(len(cutoffs)):
        ws.column_dimensions[get_column_letter(3 + j)].width = 12
    ws.freeze_panes = "C4"

    ws2 = wb.create_sheet("Evol_Cadena")
    ws2["A1"] = f"Evolución abiertas SAT por cadena — {today.isoformat()}"
    ws2["A1"].font = Font(name=FONT_X, bold=True, size=13, color=HDR_BLUE)
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cutoffs) + 1)
    hdr_cell(ws2.cell(row=3, column=1), "Cadena")
    for j, (l, _) in enumerate(cutoffs, 2):
        hdr_cell(ws2.cell(row=3, column=j), l)
    row = 4
    for ch in CHAIN_ORDER:
        f = "EAEAEA" if ch == "Otros" else (ZEBRA if (row % 2 == 0) else None)
        lbl(ws2.cell(row=row, column=1), ch, fill=f, bold=(ch == "Otros"))
        for j, (l, _) in enumerate(cutoffs, 2):
            c = sum(1 for _, x, _ in opens[l] if x == ch)
            num(ws2.cell(row=row, column=j), c if c else "", fill=f)
        row += 1
    lbl(ws2.cell(row=row, column=1), "Total", bold=True, fill=TOTAL_BG)
    for j, (l, _) in enumerate(cutoffs, 2):
        num(ws2.cell(row=row, column=j), len(opens[l]), bold=True, fill=TOTAL_BG)
    ws2.column_dimensions["A"].width = 18
    for j in range(len(cutoffs)):
        ws2.column_dimensions[get_column_letter(2 + j)].width = 12
    ws2.freeze_panes = "B4"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def html_escape(s):
    if s is None:
        return ""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def build_abiertas_hoy_section(cache):
    rows_data = []
    for key, t in cache.get("tickets", {}).items():
        if not is_open(t.get("current_status")):
            continue
        chain = chain_of(t.get("cliente", ""), t.get("loc", ""))
        days = days_since(last_status_change_dt(t))
        rows_data.append((chain, t.get("current_status"), days))

    grid = {}
    for ch, st, d in rows_data:
        k = (st, ch)
        if k not in grid:
            grid[k] = {"g": 0, "y": 0, "r": 0, "total": 0}
        grid[k][days_bucket(d)] += 1
        grid[k]["total"] += 1

    present = sorted({k[0] for k in grid.keys()})
    states_order = [s for s in FUNNEL_ORDER if s in present]
    states_order += sorted(s for s in present if s not in FUNNEL_ORDER)

    def cell_rich(d):
        if d["total"] == 0:
            return '<span style="color:#c0d0e0">·</span>'
        return (f'<span style="font-weight:500">{d["total"]}</span> '
                f'<span style="font-size:11px;color:#7d8590">(</span>'
                f'<span style="color:{COLOR_GREEN};font-weight:500">{d["g"]}</span>'
                f'<span style="font-size:11px;color:#7d8590">/</span>'
                f'<span style="color:{COLOR_YELLOW};font-weight:500">{d["y"]}</span>'
                f'<span style="font-size:11px;color:#7d8590">/</span>'
                f'<span style="color:{COLOR_RED};font-weight:500">{d["r"]}</span>'
                f'<span style="font-size:11px;color:#7d8590">)</span>')

    parts = ['<tr><th class="sticky-l1">Funnel</th><th class="sticky-l2">Estado</th>']
    for ch in CHAIN_ORDER:
        parts.append(f'<th>{ch}</th>')
    parts.append('<th>Total</th></tr>')
    header = "".join(parts)

    body_parts = []
    chain_tot = {ch: {"g": 0, "y": 0, "r": 0, "total": 0} for ch in CHAIN_ORDER}
    grand = {"g": 0, "y": 0, "r": 0, "total": 0}
    for st in states_order:
        ft = FUNNEL_TAG.get(st, "?")
        color = FUNNEL_COLOR.get(ft, "#888")
        row_tot = {"g": 0, "y": 0, "r": 0, "total": 0}
        cells = [f'<tr><td class="lbl-funnel" style="color:{color}">{ft}</td>',
                 f'<td class="lbl-name">{html_escape(st)}</td>']
        for ch in CHAIN_ORDER:
            d = grid.get((st, ch), {"g": 0, "y": 0, "r": 0, "total": 0})
            clickable = (d["total"] > 0)
            cls = "num ah-cell" if clickable else "num"
            attrs = f' data-estado="{html_escape(st)}" data-cadena="{html_escape(ch)}"' if clickable else ""
            cells.append(f'<td class="{cls}"{attrs}>{cell_rich(d)}</td>')
            for k in row_tot:
                row_tot[k] += d[k]
                chain_tot[ch][k] += d[k]
                grand[k] += d[k]
        cls_tot = "num ah-cell" if row_tot["total"] > 0 else "num"
        attrs_tot = f' data-estado="{html_escape(st)}" data-cadena=""' if row_tot["total"] > 0 else ""
        cells.append(f'<td class="{cls_tot}"{attrs_tot}>{cell_rich(row_tot)}</td>')
        cells.append('</tr>')
        body_parts.append("".join(cells))
    tot_cells = ['<tr class="total"><td class="lbl-funnel"></td><td class="lbl-name">Total</td>']
    for ch in CHAIN_ORDER:
        d = chain_tot[ch]
        clickable = (d["total"] > 0)
        cls = "num ah-cell" if clickable else "num"
        attrs = f' data-estado="" data-cadena="{html_escape(ch)}"' if clickable else ""
        tot_cells.append(f'<td class="{cls}"{attrs}>{cell_rich(d)}</td>')
    g = grand
    cls_g = "num ah-cell" if g["total"] > 0 else "num"
    attrs_g = ' data-estado="" data-cadena=""' if g["total"] > 0 else ""
    tot_cells.append(f'<td class="{cls_g}"{attrs_g}>{cell_rich(g)}</td>')
    tot_cells.append('</tr>')
    body_parts.append("".join(tot_cells))

    return header, "\n".join(body_parts)


def build_detalle_section(cache):
    JIRA_URL = "https://leaseir.atlassian.net/browse/"
    rows = []
    for key, t in cache.get("tickets", {}).items():
        if not is_open(t.get("current_status")):
            continue
        chain = chain_of(t.get("cliente", ""), t.get("loc", ""))
        days = days_since(last_status_change_dt(t))
        created_dt = parse_iso(t.get("created"))
        created_s = created_dt.strftime("%d/%m/%Y") if created_dt else ""
        fventa_dt = parse_iso(t.get("fventa") or "")
        fventa_s = fventa_dt.strftime("%d/%m/%Y") if fventa_dt else (t.get("fventa") or "")
        st = t.get("current_status") or ""
        ft = FUNNEL_TAG.get(st, "?")
        ft_color = FUNNEL_COLOR.get(ft, "#888")
        days_color = COLOR_GREEN if (days is not None and days < 5) else (
            COLOR_YELLOW if (days is not None and days <= 15) else COLOR_RED)
        rows.append({
            "chain": chain, "key": key,
            "cliente": t.get("cliente", ""),
            "status": st, "ft": ft, "ft_color": ft_color,
            "days": days if days is not None else "",
            "days_color": days_color, "loc": t.get("loc", ""),
            "created": created_s, "tipo": t.get("tipo", ""),
            "bloq": t.get("bloq", ""), "susti": t.get("susti", ""),
            "fventa": fventa_s, "consola": t.get("consola", ""),
            "hp": t.get("hp", ""), "garantia": t.get("garantia", ""),
            "desc": (t.get("descripcion", "") or "")[:200],
            "tec_taller": t.get("tec_taller", ""),
            "importe": t.get("importe", ""),
            "tec_externo": t.get("tec_externo", ""),
            "ult_com": (t.get("ult_comentario", "") or "")[:300],
        })

    idx = {s: i for i, s in enumerate(FUNNEL_ORDER)}
    rows.sort(key=lambda r: (CHAIN_ORDER.index(r["chain"]) if r["chain"] in CHAIN_ORDER else 99,
                              idx.get(r["status"], 999),
                              -(r["days"] if isinstance(r["days"], int) else 0)))

    body = []
    for r in rows:
        link = f'<a href="{JIRA_URL}{r["key"]}" target="_blank" rel="noopener" style="color:#2a59c4;text-decoration:none">{r["key"]}</a>'
        body.append(
            '<tr>'
            f'<td>{html_escape(r["chain"])}</td>'
            f'<td>{link}</td>'
            f'<td title="{html_escape(r["cliente"])}" class="d-trunc">{html_escape(r["cliente"])}</td>'
            f'<td>{html_escape(r["status"])}</td>'
            f'<td style="color:{r["ft_color"]};font-weight:500;text-align:center">{r["ft"]}</td>'
            f'<td style="color:{r["days_color"]};font-weight:500;text-align:center">{r["days"]}</td>'
            f'<td title="{html_escape(r["loc"])}" class="d-trunc">{html_escape(r["loc"])}</td>'
            f'<td>{r["created"]}</td>'
            f'<td>{html_escape(r["tipo"])}</td>'
            f'<td style="text-align:center">{html_escape(r["bloq"])}</td>'
            f'<td style="text-align:center">{html_escape(r["susti"])}</td>'
            f'<td>{r["fventa"]}</td>'
            f'<td style="text-align:center">{html_escape(r["consola"])}</td>'
            f'<td style="text-align:center">{html_escape(r["hp"])}</td>'
            f'<td style="text-align:center">{html_escape(r["garantia"])}</td>'
            f'<td title="{html_escape(r["desc"])}" class="d-trunc">{html_escape(r["desc"])}</td>'
            f'<td>{html_escape(r["tec_taller"])}</td>'
            f'<td style="text-align:right">{html_escape(r["importe"])}</td>'
            f'<td>{html_escape(r["tec_externo"])}</td>'
            f'<td title="{html_escape(r["ult_com"])}" class="d-trunc">{html_escape(r["ult_com"])}</td>'
            '</tr>'
        )
    return "\n".join(body)


def build_html(cache, out_path, template_path):
    today = date.today()
    cutoffs = compute_cutoffs(today)
    opens = replay_opens(cache, cutoffs)

    present = set()
    for l, _ in cutoffs:
        for _, _, s in opens[l]:
            if s:
                present.add(s)
    states = [s for s in FUNNEL_ORDER if s in present] + sorted(s for s in present if s not in FUNNEL_ORDER)

    # Determinar cuántos cortes son fin-de-mes para clase CSS (cortes mensuales primero)
    monthly_count = 0
    for lbl, d in cutoffs:
        try:
            dnext = d + timedelta(days=1)
            if dnext.day == 1:
                monthly_count += 1
            else:
                break
        except Exception:
            break

    def hdr_cls(i):
        return "hdr-mes" if i < monthly_count else "hdr-dia"

    th_cells = "".join(
        f'<th class="{hdr_cls(i)}">{fmt_short(l)}</th>' for i, (l, _) in enumerate(cutoffs)
    )
    state_header = '<tr><th class="sticky-l1">Funnel</th><th class="sticky-l2">Estado</th>' + th_cells + '</tr>'
    chain_header = '<tr><th class="sticky-l1">Cadena</th>' + th_cells + '</tr>'

    def state_row(st):
        ft = FUNNEL_TAG.get(st, "?")
        color = FUNNEL_COLOR.get(ft, "#888")
        parts = [f'<tr><td class="lbl-funnel" style="color:{color}">{ft}</td>',
                 f'<td class="lbl-name">{st}</td>']
        for l, _ in cutoffs:
            n = sum(1 for _, _, s in opens[l] if s == st)
            cls = "num n0" if n == 0 else "num"
            parts.append(f'<td class="{cls}">{n if n else "·"}</td>')
        parts.append('</tr>')
        return "".join(parts)

    def chain_row(ch):
        cls_row = ' class="otros"' if ch == "Otros" else ''
        parts = [f'<tr{cls_row}><td class="lbl-name">{ch}</td>']
        for l, _ in cutoffs:
            n = sum(1 for _, x, _ in opens[l] if x == ch)
            cls = "num n0" if n == 0 else "num"
            parts.append(f'<td class="{cls}">{n if n else "·"}</td>')
        parts.append('</tr>')
        return "".join(parts)

    state_rows = "\n".join(state_row(s) for s in states)
    chain_rows = "\n".join(chain_row(ch) for ch in CHAIN_ORDER)
    state_total = ('<tr class="total"><td class="lbl-funnel"></td><td class="lbl-name">Total</td>'
                   + "".join(f'<td class="num">{len(opens[l])}</td>' for l, _ in cutoffs) + '</tr>')
    chain_total = ('<tr class="total"><td class="lbl-name">Total</td>'
                   + "".join(f'<td class="num">{len(opens[l])}</td>' for l, _ in cutoffs) + '</tr>')

    today_total = len(opens[cutoffs[-1][0]])
    week_ago_total = len(opens[cutoffs[-6][0]]) if len(cutoffs) >= 6 else today_total
    dweek = today_total - week_ago_total
    arrow = "▲" if dweek > 0 else ("▼" if dweek < 0 else "→")
    dcolor = "#c0392b" if dweek > 0 else ("#1f8a4c" if dweek < 0 else "#7d8590")
    TALLER_STATUSES = {
        "Pendiente asignar técnico", "En cola taller", "En preparación presupuesto",
        "Presupuesto preparado pendiente de enviar", "Pendiente confirmación presupuesto",
        "Esperando inicio reparación", "En reparación", "Inspección de salida",
    }
    en_taller = sum(
        1 for k, t in cache.get("tickets", {}).items()
        if t.get("current_status") in TALLER_STATUSES
    )

    estancadas_15 = sum(
        1 for k, t in cache.get("tickets", {}).items()
        if is_open(t.get("current_status"))
        and (days_since(last_status_change_dt(t)) or 0) > 15
    )

    peak = 0
    peak_lbl = ""
    for lbl, _ in cutoffs:
        n = len(opens[lbl])
        if n > peak:
            peak = n
            peak_lbl = lbl
    peak_pct = round(today_total / peak * 100) if peak else 0
    peak_color = "#1f8a4c" if peak_pct <= 70 else ("#b8860b" if peak_pct <= 90 else "#c0392b")

    is_demo = "DEMO" in (cache.get("_meta", {}).get("note") or "")
    demo = ('<div class="banner">⚠ Vista previa con cache DEMO. El GitHub Action poblará el histórico completo.</div>'
            if is_demo else '')

    ah_header, ah_rows = build_abiertas_hoy_section(cache)
    detalle_rows = build_detalle_section(cache)

    # Fase 2: pestaña "Por cadena"
    chains_html = ""
    try:
        from build_chains import build_chains_html
        airtable = None
        airtable_path = out_path.parent.parent / "cache" / "airtable_pedidos.json"
        if airtable_path.exists():
            airtable = json.loads(airtable_path.read_text(encoding="utf-8"))
        serial_year = {}
        sy_path = out_path.parent.parent / "data" / "serial_year.json"
        if sy_path.exists():
            serial_year = json.loads(sy_path.read_text(encoding="utf-8"))
        def _find_cache(name):
            from pathlib import Path as _P
            for cand in [_P("cache") / name, out_path.parent.parent / "cache" / name]:
                if cand.exists():
                    return json.loads(cand.read_text(encoding="utf-8"))
            return None
        sustis = _find_cache("sustis_activas.json")
        inmov = _find_cache("airtable_inmovilizado.json")
        chains_html = build_chains_html(cache, airtable, serial_year, sustis, inmov)
    except Exception as e:
        chains_html = f'<p style="color:#c0392b;padding:20px">Error generando Fase 2: {e}</p>'

    html = template_path.read_text(encoding="utf-8")
    repl = {
        "__TODAY__": today.isoformat(),
        "__DEMO_BANNER__": demo,
        "__TODAY_TOTAL__": str(today_total),
        "__DELTA_COLOR__": dcolor,
        "__DELTA_ARROW__": arrow,
        "__DELTA_ABS__": str(abs(dweek)),
        "__ESTANCADAS_15__": str(estancadas_15),
        "__PEAK__": str(peak),
        "__PEAK_LBL__": peak_lbl,
        "__PEAK_PCT__": str(peak_pct),
        "__PEAK_COLOR__": peak_color,
        "__EN_TALLER__": str(en_taller),
        "__AH_HEADER__": ah_header,
        "__AH_ROWS__": ah_rows,
        "__DETALLE_ROWS__": detalle_rows,
        "__STATE_HEADER__": state_header,
        "__STATE_ROWS__": state_rows,
        "__STATE_TOTAL__": state_total,
        "__CHAIN_HEADER__": chain_header,
        "__CHAIN_ROWS__": chain_rows,
        "__CHAIN_TOTAL__": chain_total,
        "__CHAINS_HTML__": chains_html,
    }
    for k, v in repl.items():
        html = html.replace(k, v)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")




def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", required=True)
    ap.add_argument("--output-xlsx", default=None)
    ap.add_argument("--output-html", default=None)
    ap.add_argument("--template", default="scripts/template.html")
    args = ap.parse_args()

    cache = json.loads(Path(args.cache).read_text(encoding="utf-8"))
    if args.output_xlsx:
        build_excel(cache, Path(args.output_xlsx))
        print(f"Wrote {args.output_xlsx}")
    if args.output_html:
        build_html(cache, Path(args.output_html), Path(args.template))
        print(f"Wrote {args.output_html}")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
