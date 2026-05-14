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
    if d <= 5:
        return "g"
    if d <= 15:
        return "y"
    return "r"


def is_created_today(ticket):
    """True si el ticket fue creado hoy (UTC)."""
    created = parse_iso(ticket.get("created"))
    if not created:
        return False
    today = datetime.now(timezone.utc).date()
    return created.date() == today


def compute_status_changes_today(cache):
    """Para cada (estado, cadena) cuenta tickets que ENTRARON hoy (>= 00:00 UTC hoy)
    y SALIERON hoy. Considera tickets abiertos hoy o que estaban abiertos ayer.
    Returns: dict {(estado, cadena): {'in': int, 'out': int}}
    """
    today = datetime.now(timezone.utc).date()
    cutoff_ayer = datetime(today.year, today.month, today.day, 0, 0, 0, tzinfo=timezone.utc) - timedelta(seconds=1)
    changes = {}
    for key, t in cache.get("tickets", {}).items():
        chain = chain_of(t.get("cliente", ""), t.get("loc", ""))
        status_ayer = status_at(t, cutoff_ayer)
        status_hoy = t.get("current_status")
        if status_ayer == status_hoy:
            continue
        if status_ayer and is_open(status_ayer):
            k_out = (status_ayer, chain)
            changes.setdefault(k_out, {"in": 0, "out": 0})["out"] += 1
        if status_hoy and is_open(status_hoy):
            k_in = (status_hoy, chain)
            changes.setdefault(k_in, {"in": 0, "out": 0})["in"] += 1
    return changes


# Estados donde la "persona" relevante es el assignee (Persona Asignada de Jira).
# Desde el siguiente estado en adelante, se usa "Técnico taller" (customfield_10247).
ASIGNADO_STATES = {
    "Abierto", "Recepcionado SAT", "Pendiente recogida", "Gestionado transporte",
    "Pendiente asignar técnico", "En cola taller", "En preparación presupuesto",
    "Presupuesto preparado pendiente de enviar", "Pendiente confirmación presupuesto",
    "Esperando inicio reparación",
}

# Estados donde "de verdad estamos esperando al técnico" (suma destacada en la tabla).
ESPERANDO_TECNICO_STATES = {
    "Pendiente asignar técnico", "En cola taller", "En preparación presupuesto",
    "Esperando inicio reparación", "En reparación",
}


def person_for_status(ticket, status):
    """Devuelve la persona relevante para un (ticket, estado).
    Hasta 'Esperando inicio reparación' (incluido) → assignee (Persona Asignada).
    Desde 'En reparación' en adelante → tec_taller."""
    if status in ASIGNADO_STATES:
        return (ticket.get("asignado") or "").strip() or "(Sin asignar)"
    return (ticket.get("tec_taller") or "").strip() or "(Sin asignar)"


def compute_changes_today_por_asignado(cache):
    """Entradas/salidas hoy agrupadas por (estado, persona) — usa person_for_status."""
    today = datetime.now(timezone.utc).date()
    cutoff_ayer = datetime(today.year, today.month, today.day, 0, 0, 0, tzinfo=timezone.utc) - timedelta(seconds=1)
    changes = {}
    for key, t in cache.get("tickets", {}).items():
        status_ayer = status_at(t, cutoff_ayer)
        status_hoy = t.get("current_status")
        if status_ayer == status_hoy:
            continue
        if status_ayer and is_open(status_ayer):
            p_ayer = person_for_status(t, status_ayer)
            k_out = (status_ayer, p_ayer)
            changes.setdefault(k_out, {"in": 0, "out": 0})["out"] += 1
        if status_hoy and is_open(status_hoy):
            p_hoy = person_for_status(t, status_hoy)
            k_in = (status_hoy, p_hoy)
            changes.setdefault(k_in, {"in": 0, "out": 0})["in"] += 1
    return changes


def build_tecnicos_section(cache, avances_extra=None):
    """Tabla estados x persona (assignee hasta Esperando inicio reparación, después tec_taller)."""
    grid = {}
    asignados_set = set()
    states_set = set()
    for key, t in cache.get("tickets", {}).items():
        if not is_open(t.get("current_status")):
            continue
        st = t.get("current_status") or ""
        persona = person_for_status(t, st)
        asignados_set.add(persona)
        states_set.add(st)
        grid[(st, persona)] = grid.get((st, persona), 0) + 1

    changes = compute_changes_today_por_asignado(cache)
    asignados_sorted = sorted(asignados_set, key=lambda x: (0 if x == "(Sin asignar)" else 1, x.lower()))
    states_order = [s for s in FUNNEL_ORDER if s in states_set] + sorted(s for s in states_set if s not in FUNNEL_ORDER)
    taller_set = set(TALLER_STATUSES_ORDER)

    ARROW_IN = "#c0392b"
    ARROW_OUT = "#1f8a4c"

    def flow_html(f_in, f_out):
        parts = []
        if f_in > 0:
            parts.append(f'<span style="color:{ARROW_IN};font-size:10px;font-weight:500" title="entran hoy">&#9650;{f_in}</span>')
        if f_out > 0:
            parts.append(f'<span style="color:{ARROW_OUT};font-size:10px;font-weight:500" title="salen hoy">&#9660;{f_out}</span>')
        return (' ' + ' '.join(parts)) if parts else ''

    # Totales por persona
    col_tot = {a: 0 for a in asignados_sorted}
    col_flow = {a: {"in": 0, "out": 0} for a in asignados_sorted}
    for (st, a), n in grid.items():
        if a in col_tot:
            col_tot[a] += n
    for (st, a), f in changes.items():
        if a in col_flow:
            col_flow[a]["in"] += f["in"]
            col_flow[a]["out"] += f["out"]
    grand_total = sum(col_tot.values())
    grand_flow = {"in": sum(f["in"] for f in col_flow.values()),
                  "out": sum(f["out"] for f in col_flow.values())}

    # Header (sin total grande — el total ya aparece en la fila Total inferior)
    parts = ['<tr><th class="sticky-l1 sticky-l2" style="text-align:left;left:0;min-width:240px">Estado</th>']
    for a in asignados_sorted:
        parts.append(
            f'<th style="padding:8px 10px"><div style="font-size:11px;line-height:1.2">{html_escape(a)}</div></th>'
        )
    parts.append('<th class="hdr-dia col-total" style="padding:8px 10px">Total</th>')
    parts.append('</tr>')
    header = "".join(parts)

    # Body
    body_parts = []
    for st in states_order:
        is_taller = st in taller_set
        idx = states_order.index(st)
        prev_taller = (idx > 0 and states_order[idx - 1] in taller_set)
        next_taller = (idx < len(states_order) - 1 and states_order[idx + 1] in taller_set)
        cls_list = []
        if is_taller:
            cls_list.append("row-taller")
            if not prev_taller:
                cls_list.append("row-taller-first")
            if not next_taller:
                cls_list.append("row-taller-last")
        cls_attr = f' class="{" ".join(cls_list)}"' if cls_list else ''
        row_tot = 0
        row_flow = {"in": 0, "out": 0}
        cells = [f'<tr{cls_attr}><td class="lbl-name" style="left:0;min-width:240px">{html_escape(st)}</td>']
        for a in asignados_sorted:
            n = grid.get((st, a), 0)
            f = changes.get((st, a), {"in": 0, "out": 0})
            row_tot += n
            row_flow["in"] += f["in"]
            row_flow["out"] += f["out"]
            if n > 0 or f["in"] > 0 or f["out"] > 0:
                content = f'<span style="font-weight:500">{n if n else "·"}</span>{flow_html(f["in"], f["out"])}'
                attrs = f' data-estado="{html_escape(st)}" data-asignado="{html_escape(a)}"'
                cells.append(f'<td class="num tec-cell"{attrs}>{content}</td>')
            else:
                cells.append('<td class="num n0"><span style="color:#c0d0e0">·</span></td>')
        # Total fila
        if row_tot > 0 or row_flow["in"] > 0 or row_flow["out"] > 0:
            content = f'<span style="font-weight:500">{row_tot}</span>{flow_html(row_flow["in"], row_flow["out"])}'
            attrs = f' data-estado="{html_escape(st)}" data-asignado=""'
            cells.append(f'<td class="num col-total tec-cell"{attrs}>{content}</td>')
        else:
            cells.append('<td class="num col-total n0"><span style="color:#c0d0e0">·</span></td>')
        cells.append('</tr>')
        body_parts.append("".join(cells))

    # Fila "Esperando técnico": suma de los 5 estados clave por persona
    esp_tot = {a: 0 for a in asignados_sorted}
    esp_flow = {a: {"in": 0, "out": 0} for a in asignados_sorted}
    for (st, a), n in grid.items():
        if st in ESPERANDO_TECNICO_STATES and a in esp_tot:
            esp_tot[a] += n
    for (st, a), f in changes.items():
        if st in ESPERANDO_TECNICO_STATES and a in esp_flow:
            esp_flow[a]["in"] += f["in"]
            esp_flow[a]["out"] += f["out"]
    esp_grand = sum(esp_tot.values())
    esp_grand_flow = {"in": sum(f["in"] for f in esp_flow.values()),
                      "out": sum(f["out"] for f in esp_flow.values())}
    esp_estado_list = "|".join(sorted(ESPERANDO_TECNICO_STATES))
    esp_cells = [f'<tr style="background:#fff4d6"><td class="lbl-name" style="left:0;min-width:240px;font-weight:500;background:#fff4d6" title="Suma de tickets en: {html_escape(esp_estado_list.replace("|", ", "))}">Esperando técnico</td>']
    for a in asignados_sorted:
        v = esp_tot.get(a, 0)
        f = esp_flow.get(a, {"in": 0, "out": 0})
        if v > 0 or f["in"] > 0 or f["out"] > 0:
            content = f'<span style="font-weight:600">{v}</span>{flow_html(f["in"], f["out"])}'
            attrs = f' data-estado-list="{html_escape(esp_estado_list)}" data-asignado="{html_escape(a)}"'
            esp_cells.append(f'<td class="num tec-cell" style="background:#fff4d6"{attrs}>{content}</td>')
        else:
            esp_cells.append('<td class="num n0" style="background:#fff4d6"><span style="color:#c0d0e0">·</span></td>')
    content = f'<span style="font-weight:600">{esp_grand}</span>{flow_html(esp_grand_flow["in"], esp_grand_flow["out"])}'
    esp_cells.append(f'<td class="num col-total tec-cell" style="background:#fff4d6;font-weight:600" data-estado-list="{html_escape(esp_estado_list)}" data-asignado="">{content}</td>')
    esp_cells.append('</tr>')
    body_parts.append("".join(esp_cells))

    # Fila Total
    tot_cells = ['<tr class="total"><td class="lbl-name" style="left:0;min-width:240px">Total</td>']
    for a in asignados_sorted:
        v = col_tot.get(a, 0)
        f = col_flow.get(a, {"in": 0, "out": 0})
        if v > 0 or f["in"] > 0 or f["out"] > 0:
            content = f'<span style="font-weight:500">{v}</span>{flow_html(f["in"], f["out"])}'
            attrs = f' data-estado="" data-asignado="{html_escape(a)}"'
            tot_cells.append(f'<td class="num tec-cell"{attrs}>{content}</td>')
        else:
            tot_cells.append('<td class="num n0"><span style="color:#c0d0e0">·</span></td>')
    content = f'<span style="font-weight:500">{grand_total}</span>{flow_html(grand_flow["in"], grand_flow["out"])}'
    tot_cells.append(f'<td class="num col-total tec-cell" data-estado="" data-asignado="">{content}</td>')
    tot_cells.append('</tr>')
    body_parts.append("".join(tot_cells))

    # === Filas adicionales de Avances (Presupuesto hecho, Equipo reparado) ===
    if avances_extra:
        # avances_extra = {"grid": {(label, persona): count}, "details": {(label, persona): [tickets]}, "labels": [(state, label)]}
        avg = avances_extra.get("grid", {})
        adetails = avances_extra.get("details", {})
        alabels = avances_extra.get("labels", [])
        for state_origin, lbl in alabels:
            row_cells = [f'<tr class="avances-row"><td class="lbl-name" style="left:0;min-width:240px;font-weight:500" title="Tickets unicos avanzados hoy desde {html_escape(state_origin)} hacia delante">{html_escape(lbl)} hoy</td>']
            row_tot = 0
            for a in asignados_sorted:
                n = avg.get((lbl, a), 0)
                row_tot += n
                if n > 0:
                    dkey = f"{lbl}__{a}".replace(" ", "_").replace("/", "_")
                    attrs = f' data-detail-key="{html_escape(dkey)}" data-estado="{html_escape(lbl)}" data-persona="{html_escape(a)}"'
                    row_cells.append(f'<td class="num avance-cell"{attrs}><span style="font-weight:500;color:#16a34a">{n}</span></td>')
                else:
                    row_cells.append('<td class="num n0"><span style="color:#c0d0e0">·</span></td>')
            # Total col
            if row_tot > 0:
                row_cells.append(f'<td class="num col-total"><span style="font-weight:600;color:#16a34a">{row_tot}</span></td>')
            else:
                row_cells.append('<td class="num col-total n0"><span style="color:#c0d0e0">·</span></td>')
            row_cells.append('</tr>')
            body_parts.append("".join(row_cells))

    return header, "\n".join(body_parts)


def build_avances_section(cache):
    """Métrica de productividad real:
    - "Presupuesto hecho" = transiciones hoy con to_status = "Presupuesto preparado pendiente de enviar"
    - "Equipo reparado" = transiciones hoy con to_status = "Inspección de salida"
    Dedupe por (ticket, persona): si un ticket toca el to_status 2 veces, cuenta 1 para esa persona.
    """
    # to_status objetivo -> label visible
    AVANCES_LABELS = [
        ("Presupuesto preparado pendiente de enviar", "Presupuesto hecho"),
        ("Inspección de salida", "Equipo reparado"),
    ]
    TO_STATES = {to for to, _ in AVANCES_LABELS}
    AVANCES_LABEL_BY_STATE = dict(AVANCES_LABELS)
    AVANCES_STATES_ORDER = [to for to, _ in AVANCES_LABELS]
    AVANCES_SET = set(AVANCES_STATES_ORDER)

    today = datetime.now(timezone.utc).date()
    cutoff = datetime(today.year, today.month, today.day, 0, 0, 0, tzinfo=timezone.utc)
    try:
        from zoneinfo import ZoneInfo
        madrid = ZoneInfo("Europe/Madrid")
    except Exception:
        madrid = None

    def fmt_hora(dt):
        try:
            return dt.astimezone(madrid).strftime("%H:%M") if madrid else dt.strftime("%H:%M")
        except Exception:
            return ""

    grid = {}
    detail_map = {}  # (to_state, person, ticket_key) -> {key, cliente, from, to, ts, dt}
    personas_set = set()

    for k, t in cache.get("tickets", {}).items():
        for tr in t.get("transitions", []):
            if len(tr) < 3:
                continue
            ts_str, from_s, to_s = tr[0], tr[1], tr[2]
            if to_s not in TO_STATES:
                continue
            dt = parse_iso(ts_str)
            if not dt or dt < cutoff:
                continue
            person = person_for_status(t, to_s)
            personas_set.add(person)
            cnt_key = (to_s, person, k)
            if cnt_key not in detail_map:
                grid[(to_s, person)] = grid.get((to_s, person), 0) + 1
            # Mantener la última transición a este to_state
            existing = detail_map.get(cnt_key)
            if not existing or dt > existing["dt"]:
                detail_map[cnt_key] = {
                    "key": k, "cliente": t.get("cliente", "") or "",
                    "from": from_s, "to": to_s,
                    "ts": fmt_hora(dt), "dt": dt,
                }

    details = {}
    for (to_s, person, k), info in detail_map.items():
        details.setdefault((to_s, person), []).append({
            "key": info["key"], "cliente": info["cliente"],
            "from": info["from"], "to": info["to"], "ts": info["ts"],
        })
    for lst in details.values():
        lst.sort(key=lambda x: x["ts"], reverse=True)

    if not personas_set:
        header = ('<tr><th class="sticky-l1 sticky-l2" style="text-align:left;left:0;min-width:240px">Tarea</th>'
                  '<th>(sin avances aún)</th><th class="col-total">Total</th></tr>')
        rows = ('<tr><td class="lbl-name" style="left:0;min-width:240px" colspan="3">'
                '<span style="color:#94a3b8;font-style:italic">Aún no hay tareas completadas hoy</span></td></tr>')
        return header, rows, "{}", {"grid": {}, "details": {}, "labels": []}

    personas_sorted = sorted(personas_set, key=lambda x: (0 if x == "(Sin asignar)" else 1, x.lower()))
    parts = ['<tr><th class="sticky-l1 sticky-l2" style="text-align:left;left:0;min-width:240px">Tarea</th>']
    for p in personas_sorted:
        parts.append(f'<th>{html_escape(p)}</th>')
    parts.append('<th class="hdr-dia col-total">Total</th></tr>')
    header = "".join(parts)
    body_parts = []
    col_tot = {p: 0 for p in personas_sorted}
    grand = 0
    detail_dict = {}
    for to_s in AVANCES_STATES_ORDER:
        lbl = AVANCES_LABEL_BY_STATE.get(to_s, to_s)
        row_tot = 0
        cells = [f'<tr><td class="lbl-name" style="left:0;min-width:240px">{html_escape(lbl)}</td>']
        for p in personas_sorted:
            n = grid.get((to_s, p), 0)
            row_tot += n
            col_tot[p] += n
            if n > 0:
                dkey = f"{lbl}__{p}".replace(" ", "_").replace("/", "_")
                detail_dict[dkey] = details.get((to_s, p), [])
                attrs = f' data-detail-key="{html_escape(dkey)}" data-estado="{html_escape(lbl)}" data-persona="{html_escape(p)}"'
                cells.append(f'<td class="num avance-cell"{attrs}><span style="font-weight:500">{n}</span></td>')
            else:
                cells.append('<td class="num n0"><span style="color:#c0d0e0">·</span></td>')
        grand += row_tot
        if row_tot > 0:
            cells.append(f'<td class="num col-total"><span style="font-weight:500">{row_tot}</span></td>')
        else:
            cells.append('<td class="num col-total n0"><span style="color:#c0d0e0">·</span></td>')
        cells.append('</tr>')
        body_parts.append("".join(cells))
    tot_cells = ['<tr class="total"><td class="lbl-name" style="left:0;min-width:240px">Total</td>']
    for p in personas_sorted:
        v = col_tot.get(p, 0)
        if v > 0:
            tot_cells.append(f'<td class="num"><span style="font-weight:600">{v}</span></td>')
        else:
            tot_cells.append('<td class="num n0"><span style="color:#c0d0e0">·</span></td>')
    tot_cells.append(f'<td class="num col-total"><span style="font-weight:600">{grand}</span></td>')
    tot_cells.append('</tr>')
    body_parts.append("".join(tot_cells))

    extra = {
        "grid": {(AVANCES_LABEL_BY_STATE.get(to, to), p): n for (to, p), n in grid.items()},
        "details": {(AVANCES_LABEL_BY_STATE.get(to, to), p): lst for (to, p), lst in details.items()},
        "labels": list(AVANCES_LABELS),
    }
    detail_dict_extra = {}
    for (to_s, p), lst in details.items():
        lbl = AVANCES_LABEL_BY_STATE.get(to_s, to_s)
        dkey = f"{lbl}__{p}".replace(" ", "_").replace("/", "_")
        detail_dict_extra[dkey] = lst
    detail_dict.update(detail_dict_extra)
    detail_json = json.dumps(detail_dict, ensure_ascii=False)
    return header, "\n".join(body_parts), detail_json, extra


def build_anual_averias_section(cache):
    """Tabla cruzada Año venta x Tipo averia para tickets 2026 que pasaron por
    Inspeccion de salida. Usa customfield_10815 (Cambios obs y mejoras),
    agrupado en 10 buckets clinicos. Genera datos embebidos en JS para que el
    filtro (cadena + importe) se aplique client-side sin recargar la pagina.
    """
    INSP_STATES = {"Inspeccion de salida", "Inspección de salida"}

    BUCKETS = [
        ("Umbi", "Umbilical",
         ["Cambio de umbilical", "Reparación umbilical", "Reparacion umbilical"]),
        ("Punt", "Puntera/Óptica",
         ["Cambio de puntera", "Cambio de zafiro", "Cambio de lente trasera", "Cambio de prisma"]),
        ("Buff", "Buffer", ["Cambio de buffer"]),
        ("Diod", "Diodo", ["Diodo nuevo", "Diodo antirretorno"]),
        ("Plac", "Placa/Electr.",
         ["Cambio de placa de carga", "MOSFET de disparo 1º", "MOSFET de disparo 2º",
          "MOSFET de carga", "Integrado LT1054", "Cambio de driver", "Pila BIOS",
          "Unidad SSD", "RAM"]),
        ("Gati", "Gatillo/Pistola",
         ["Cambio switch de gatillo", "Cambio de carcasa de pistola"]),
        ("Pant", "Pantalla/Control",
         ["Cambio de táctil + controlador", "Cambio de tactil + controlador",
          "Cambio de control", "Switch de pedal"]),
        ("Refr", "Refrigeración",
         ["Cambio de nevera", "Cambio de termostato", "Ventilador de 12V",
          "Mejora de hidráulica", "Mejora de hidraulica"]),
        ("Mant", "Mantenimiento",
         ["Limpieza interior y exterior", "Llenado de depósito", "Llenado de deposito",
          "Calibrado y medición de energías", "Calibrado y medicion de energias",
          "Mejora de masas", "Actualizar software", "Casquillo nuevo", "Lanyard", "Buzzer"]),
        ("Otro", "Otros", ["Otros"]),
    ]
    BUCKET_KEYS = [b[0] for b in BUCKETS]
    BUCKET_LABELS = {b[0]: b[1] for b in BUCKETS}
    val2bucket = {}
    for short, lbl, vals in BUCKETS:
        for v in vals:
            val2bucket[v.lower()] = short

    def importe_bucket(imp):
        try:
            v = float(imp) if imp not in ("", None) else 0
        except (TypeError, ValueError):
            v = 0
        if v == 0: return "free"
        if v < 1000: return "low"
        if v <= 6000: return "mid"
        return "high"

    YEAR_ROWS = ["<=2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025+", "Desconocido", "Abierta"]
    def year_row(y):
        if not y: return "Desconocido"
        try:
            yi = int(y)
        except (TypeError, ValueError):
            return "Desconocido"
        if yi <= 1999: return "Desconocido"
        if yi <= 2018: return "<=2018"
        if yi >= 2025: return "2025+"
        return str(yi)

    records = []
    for k, t in cache.get("tickets", {}).items():
        created = t.get("created") or ""
        if not created.startswith("2026"):
            continue
        passed = False
        for tr in t.get("transitions", []):
            if len(tr) >= 3 and tr[2] in INSP_STATES:
                passed = True
                break
        fv = (t.get("fventa") or "").strip()
        yr = None
        if fv and len(fv) >= 4 and fv[:4].isdigit():
            yr = int(fv[:4])
        buckets = set()
        if passed:
            for v in (t.get("cambios") or []):
                b = val2bucket.get(str(v).lower())
                if b: buckets.add(b)
            if not buckets:
                for v in (t.get("motivo") or []):
                    vl = str(v).lower()
                    if "umbilical" in vl: buckets.add("Umbi")
                    elif "buffer" in vl: buckets.add("Buff")
                    elif "puntera" in vl or "zafiro" in vl: buckets.add("Punt")
                    elif "diodo" in vl: buckets.add("Diod")
                    elif "placa" in vl: buckets.add("Plac")
                    elif "gatillo" in vl: buckets.add("Gati")
                    elif "pantalla" in vl: buckets.add("Pant")
                    elif "otros" in vl: buckets.add("Otro")
        ch = chain_of(t.get("cliente", ""), t.get("loc", ""))
        yrow = year_row(yr) if passed else "Abierta"
        records.append({"k": k, "y": yrow, "c": ch,
                        "i": importe_bucket(t.get("importe")),
                        "b": sorted(buckets)})

    if not records:
        return ('<div style="color:#94a3b8;padding:14px;font-style:italic">'
                'Aun no hay reparaciones 2026 con datos suficientes.</div>')

    chains_present = set(r["c"] for r in records)
    cadenas = [c for c in CHAIN_ORDER if c in chains_present]
    n_total = len(records)

    data_json = json.dumps(records, ensure_ascii=False)
    cad_options = ''.join(f'<option value="{html_escape(c)}">{html_escape(c)}</option>' for c in cadenas)
    bucket_headers = ''.join(
        f'<th style="background:var(--blue);color:white;font-weight:500;padding:6px 8px;text-align:right" title="{html_escape(BUCKET_LABELS[k])}">{k}</th>'
        for k in BUCKET_KEYS)
    year_rows_html = ''.join(
        f'<tr data-year="{y}"><td style="padding:6px 10px;font-weight:500">{y}</td>'
        f'<td class="num tickets-cell">0</td>'
        + ''.join(f'<td class="num bucket-cell" data-bucket="{k}">0</td>' for k in BUCKET_KEYS)
        + '</tr>'
        for y in YEAR_ROWS)

    html = (
        '<div style="display:flex;gap:14px;align-items:center;flex-wrap:wrap;margin:6px 0 10px;padding:9px 12px;background:#f1f5f9;border:1px solid var(--line);border-radius:7px;font-size:12px">'
        '<label style="display:inline-flex;align-items:center;gap:6px"><span style="color:#475569">Cadena:</span>'
        '<select id="av-filter-chain" style="padding:4px 8px;border:1px solid var(--line);border-radius:4px;font-size:12px;background:white">'
        '<option value="">Todas</option>'
        f'{cad_options}'
        '</select></label>'
        '<label style="display:inline-flex;align-items:center;gap:6px"><span style="color:#475569">Importe:</span>'
        '<select id="av-filter-importe" style="padding:4px 8px;border:1px solid var(--line);border-radius:4px;font-size:12px;background:white">'
        '<option value="">Todos</option>'
        '<option value="free">Sin coste</option>'
        '<option value="low">&lt; 1.000 €</option>'
        '<option value="mid">1.000 - 6.000 €</option>'
        '<option value="high">&gt; 6.000 €</option>'
        '</select></label>'
        '<span style="margin-left:auto;color:#475569">Tickets visibles: <b id="av-count" style="color:var(--blue)">'
        f'{n_total}</b> / {n_total}</span>'
        '</div>'
        '<div style="border:1px solid var(--line);border-radius:8px;overflow:hidden;background:white">'
        '<table id="av-table" style="width:100%;border-collapse:collapse;font-size:12.5px">'
        '<thead><tr>'
        '<th style="background:var(--blue);color:white;font-weight:500;padding:6px 10px;text-align:left">Año venta</th>'
        '<th style="background:var(--blue);color:white;font-weight:500;padding:6px 8px;text-align:right">Tickets</th>'
        f'{bucket_headers}'
        '</tr></thead>'
        f'<tbody>{year_rows_html}</tbody>'
        '<tfoot><tr class="total"><td style="padding:6px 10px;font-weight:600">TOTAL</td>'
        '<td class="num tickets-cell" style="font-weight:600">0</td>'
        + ''.join(f'<td class="num bucket-cell" data-bucket="{k}" style="font-weight:600">0</td>' for k in BUCKET_KEYS)
        + '</tr></tfoot>'
        '</table></div>'
        '<div class="legend" style="margin-top:6px">'
        + ' &middot; '.join(f'<b>{k}</b>={html_escape(BUCKET_LABELS[k])}' for k in BUCKET_KEYS) +
        '</div>'
        '<script id="av-data" type="application/json">' + data_json + '</script>'
        '<script>'
        '(function(){'
        'var DATA = JSON.parse(document.getElementById("av-data").textContent);'
        'var YEARS = ' + json.dumps(YEAR_ROWS) + ';'
        'var BUCKETS = ' + json.dumps(BUCKET_KEYS) + ';'
        'var tbl = document.getElementById("av-table");'
        'var selC = document.getElementById("av-filter-chain");'
        'var selI = document.getElementById("av-filter-importe");'
        'var cnt = document.getElementById("av-count");'
        'function recompute(){'
        '  var ch = selC.value, im = selI.value;'
        '  var grid = {}; YEARS.forEach(function(y){grid[y]={"_t":0}; BUCKETS.forEach(function(b){grid[y][b]=0;});});'
        '  var totals = {"_t":0}; BUCKETS.forEach(function(b){totals[b]=0;});'
        '  var visible = 0;'
        '  DATA.forEach(function(r){'
        '    if(ch && r.c!==ch) return;'
        '    if(im && r.i!==im) return;'
        '    visible++; grid[r.y]._t++; totals._t++;'
        '    r.b.forEach(function(b){ if(grid[r.y][b]!==undefined){grid[r.y][b]++; totals[b]++;} });'
        '  });'
        '  cnt.textContent = visible;'
        '  tbl.querySelectorAll("tbody tr").forEach(function(tr){'
        '    var y = tr.getAttribute("data-year"); var t = grid[y]._t;'
        '    var tc = tr.querySelector(".tickets-cell");'
        '    tc.textContent = t || "·"; tc.style.color = t ? "" : "#c0d0e0";'
        '    tr.querySelectorAll(".bucket-cell").forEach(function(td){'
        '      var v = grid[y][td.getAttribute("data-bucket")];'
        '      td.textContent = v || "·"; td.style.color = v ? "" : "#c0d0e0";'
        '    });'
        '  });'
        '  var tf = tbl.querySelector("tfoot tr");'
        '  tf.querySelector(".tickets-cell").textContent = totals._t;'
        '  tf.querySelectorAll(".bucket-cell").forEach(function(td){'
        '    td.textContent = totals[td.getAttribute("data-bucket")] || "·";'
        '  });'
        '}'
        'selC.addEventListener("change", recompute);'
        'selI.addEventListener("change", recompute);'
        'recompute();'
        '})();'
        '</script>'
    )
    return html


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


TALLER_STATUSES_ORDER = [
    "Pendiente asignar técnico", "En cola taller", "En preparación presupuesto",
    "Presupuesto preparado pendiente de enviar", "Pendiente confirmación presupuesto",
    "Esperando inicio reparación", "En reparación", "Inspección de salida",
]


def build_abiertas_hoy_section(cache):
    rows_data = []
    for key, t in cache.get("tickets", {}).items():
        if not is_open(t.get("current_status")):
            continue
        chain = chain_of(t.get("cliente", ""), t.get("loc", ""))
        days = days_since(last_status_change_dt(t))
        created_today = is_created_today(t)
        rows_data.append((chain, t.get("current_status"), days, created_today))

    grid = {}
    for ch, st, d, ct in rows_data:
        k = (st, ch)
        if k not in grid:
            grid[k] = {"t": 0, "g": 0, "y": 0, "r": 0, "total": 0}
        # Azul (creado hoy) tiene prioridad sobre verde/amarillo/rojo
        bucket = "t" if ct else days_bucket(d)
        grid[k][bucket] += 1
        grid[k]["total"] += 1

    changes_today = compute_status_changes_today(cache)

    present = sorted({k[0] for k in grid.keys()})
    states_order = [s for s in FUNNEL_ORDER if s in present]
    states_order += sorted(s for s in present if s not in FUNNEL_ORDER)
    taller_set = set(TALLER_STATUSES_ORDER)

    COLOR_TODAY = "#1e6fdb"
    ARROW_IN = "#c0392b"
    ARROW_OUT = "#1f8a4c"

    def cell_rich(d, ch_in=0, ch_out=0):
        if d["total"] == 0 and ch_in == 0 and ch_out == 0:
            return '<span style="color:#c0d0e0">·</span>'
        if d["total"] == 0:
            body = '<span style="color:#c0d0e0">·</span>'
        else:
            inner = ""
            if d["t"] > 0:
                inner += f'<span style="color:{COLOR_TODAY};font-weight:500">{d["t"]}</span><span style="font-size:11px;color:#7d8590">/</span>'
            inner += (f'<span style="color:{COLOR_GREEN};font-weight:500">{d["g"]}</span>'
                      f'<span style="font-size:11px;color:#7d8590">/</span>'
                      f'<span style="color:{COLOR_YELLOW};font-weight:500">{d["y"]}</span>'
                      f'<span style="font-size:11px;color:#7d8590">/</span>'
                      f'<span style="color:{COLOR_RED};font-weight:500">{d["r"]}</span>')
            body = (f'<span style="font-weight:500">{d["total"]}</span> '
                    f'<span style="font-size:11px;color:#7d8590">(</span>'
                    f'{inner}'
                    f'<span style="font-size:11px;color:#7d8590">)</span>')
        flow = []
        if ch_in > 0:
            flow.append(f'<span style="color:{ARROW_IN};font-weight:500;font-size:11px" title="entran hoy">&#9650;{ch_in}</span>')
        if ch_out > 0:
            flow.append(f'<span style="color:{ARROW_OUT};font-weight:500;font-size:11px" title="salen hoy">&#9660;{ch_out}</span>')
        if flow:
            body += ' <span style="font-size:11px;color:#7d8590">·</span> ' + ' '.join(flow)
        return body

    parts = ['<tr><th class="sticky-l1">Funnel</th><th class="sticky-l2">Estado</th>']
    for ch in CHAIN_ORDER:
        parts.append(f'<th>{ch}</th>')
    parts.append('<th class="col-total">Total</th></tr>')
    header = "".join(parts)

    body_parts = []
    chain_tot = {ch: {"t": 0, "g": 0, "y": 0, "r": 0, "total": 0} for ch in CHAIN_ORDER}
    chain_flow = {ch: {"in": 0, "out": 0} for ch in CHAIN_ORDER}
    grand = {"t": 0, "g": 0, "y": 0, "r": 0, "total": 0}
    grand_flow = {"in": 0, "out": 0}
    for st in states_order:
        ft = FUNNEL_TAG.get(st, "?")
        color = FUNNEL_COLOR.get(ft, "#888")
        row_tot = {"t": 0, "g": 0, "y": 0, "r": 0, "total": 0}
        row_flow = {"in": 0, "out": 0}
        is_taller = st in taller_set
        idx = states_order.index(st)
        prev_taller = (idx > 0 and states_order[idx-1] in taller_set)
        next_taller = (idx < len(states_order)-1 and states_order[idx+1] in taller_set)
        row_classes = []
        if is_taller:
            row_classes.append("row-taller")
            if not prev_taller:
                row_classes.append("row-taller-first")
            if not next_taller:
                row_classes.append("row-taller-last")
        row_cls_attr = f' class="{" ".join(row_classes)}"' if row_classes else ""
        cells = [f'<tr{row_cls_attr}><td class="lbl-funnel" style="color:{color}">{ft}</td>',
                 f'<td class="lbl-name">{html_escape(st)}</td>']
        for ch in CHAIN_ORDER:
            d = grid.get((st, ch), {"t": 0, "g": 0, "y": 0, "r": 0, "total": 0})
            ch_changes = changes_today.get((st, ch), {"in": 0, "out": 0})
            clickable = (d["total"] > 0)
            cls = "num ah-cell" if clickable else "num"
            attrs = f' data-estado="{html_escape(st)}" data-cadena="{html_escape(ch)}"' if clickable else ""
            cells.append(f'<td class="{cls}"{attrs}>{cell_rich(d, ch_changes["in"], ch_changes["out"])}</td>')
            for k in row_tot:
                row_tot[k] += d[k]
                chain_tot[ch][k] += d[k]
                grand[k] += d[k]
            row_flow["in"] += ch_changes["in"]
            row_flow["out"] += ch_changes["out"]
            chain_flow[ch]["in"] += ch_changes["in"]
            chain_flow[ch]["out"] += ch_changes["out"]
            grand_flow["in"] += ch_changes["in"]
            grand_flow["out"] += ch_changes["out"]
        cls_tot = "num col-total ah-cell" if row_tot["total"] > 0 else "num col-total"
        attrs_tot = f' data-estado="{html_escape(st)}" data-cadena=""' if row_tot["total"] > 0 else ""
        cells.append(f'<td class="{cls_tot}"{attrs_tot}>{cell_rich(row_tot, row_flow["in"], row_flow["out"])}</td>')
        cells.append('</tr>')
        body_parts.append("".join(cells))
    tot_cells = ['<tr class="total"><td class="lbl-funnel"></td><td class="lbl-name">Total</td>']
    for ch in CHAIN_ORDER:
        d = chain_tot[ch]
        cf = chain_flow[ch]
        clickable = (d["total"] > 0)
        cls = "num ah-cell" if clickable else "num"
        attrs = f' data-estado="" data-cadena="{html_escape(ch)}"' if clickable else ""
        tot_cells.append(f'<td class="{cls}"{attrs}>{cell_rich(d, cf["in"], cf["out"])}</td>')
    g = grand
    cls_g = "num col-total ah-cell" if g["total"] > 0 else "num col-total"
    attrs_g = ' data-estado="" data-cadena=""' if g["total"] > 0 else ""
    tot_cells.append(f'<td class="{cls_g}"{attrs_g}>{cell_rich(g, grand_flow["in"], grand_flow["out"])}</td>')
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
            "asignado": t.get("asignado", ""),
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
            f'<td style="color:{r["dt_color"]};font-weight:500;text-align:center" title="Entró en Pdte. asignar técnico el {r["dt_taller_label"] or "—"}">{r["dias_taller"] if r["dias_taller"] != "" else "—"}</td>'
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
            f'<td>{html_escape(r["asignado"])}</td>'
            '</tr>'
        )
    return "\n".join(body)


def build_html(cache, out_path, template_path):
    today = date.today()
    try:
        from zoneinfo import ZoneInfo
        _now_mad = datetime.now(ZoneInfo("Europe/Madrid"))
        today_label = _now_mad.strftime("%Y-%m-%d %H:%M") + " (Madrid)"
    except Exception:
        today_label = today.isoformat()
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

    taller_set_local = set(TALLER_STATUSES_ORDER)
    states_taller_idx = [i for i, s in enumerate(states) if s in taller_set_local]
    taller_first_idx = states_taller_idx[0] if states_taller_idx else -1
    taller_last_idx = states_taller_idx[-1] if states_taller_idx else -1

    def state_row(st):
        ft = FUNNEL_TAG.get(st, "?")
        color = FUNNEL_COLOR.get(ft, "#888")
        row_classes = []
        if st in taller_set_local:
            row_classes.append("row-taller")
            i = states.index(st)
            if i == taller_first_idx:
                row_classes.append("row-taller-first")
            if i == taller_last_idx:
                row_classes.append("row-taller-last")
        cls_attr = f' class="{" ".join(row_classes)}"' if row_classes else ""
        parts = [f'<tr{cls_attr}><td class="lbl-funnel" style="color:{color}">{ft}</td>',
                 f'<td class="lbl-name">{st}</td>']
        for l, _ in cutoffs:
            n = sum(1 for _, _, s in opens[l] if s == st)
            cls = "num n0" if n == 0 else "num"
            parts.append(f'<td class="{cls}">{n if n else "·"}</td>')
        parts.append('</tr>')
        return "".join(parts)

    # nuevas_por_chain_cutoff: {(ch, lbl): count_nuevas_esa_cadena_ese_corte}
    nuevas_por_chain_cutoff = {}
    for i, (lbl, d) in enumerate(cutoffs):
        if i < monthly_count:
            start_dt = datetime(d.year, d.month, 1, 0, 0, 0, tzinfo=timezone.utc)
        else:
            start_dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=timezone.utc)
        end_dt = datetime(d.year, d.month, d.day, 23, 59, 59, tzinfo=timezone.utc)
        for key, t in cache.get("tickets", {}).items():
            created = parse_iso(t.get("created"))
            if not created or not (start_dt <= created <= end_dt):
                continue
            ch = chain_of(t.get("cliente", ""), t.get("loc", ""))
            nuevas_por_chain_cutoff[(ch, lbl)] = nuevas_por_chain_cutoff.get((ch, lbl), 0) + 1

    def chain_row(ch):
        cls_row = ' class="otros"' if ch == "Otros" else ''
        parts = [f'<tr{cls_row}><td class="lbl-name">{ch}</td>']
        for l, _ in cutoffs:
            n = sum(1 for _, x, _ in opens[l] if x == ch)
            new = nuevas_por_chain_cutoff.get((ch, l), 0)
            cls = "num n0" if n == 0 else "num"
            if new > 0:
                nuevas_txt = f'<span style="color:#16a34a;font-size:9px;font-weight:500;margin-left:3px" title="Nuevas creadas">(+{new})</span>'
                parts.append(f'<td class="{cls}">{n if n else chr(183)}{nuevas_txt}</td>')
            else:
                parts.append(f'<td class="{cls}">{n if n else chr(183)}</td>')
        parts.append('</tr>')
        return "".join(parts)

    state_rows = "\n".join(state_row(s) for s in states)
    chain_rows = "\n".join(chain_row(ch) for ch in CHAIN_ORDER)
    nuevas_total = {}
    for i, (lbl, d) in enumerate(cutoffs):
        if i < monthly_count:
            start_dt = datetime(d.year, d.month, 1, 0, 0, 0, tzinfo=timezone.utc)
        else:
            start_dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=timezone.utc)
        end_dt = datetime(d.year, d.month, d.day, 23, 59, 59, tzinfo=timezone.utc)
        cnt = 0
        for key, t in cache.get("tickets", {}).items():
            created = parse_iso(t.get("created"))
            if not created:
                continue
            if start_dt <= created <= end_dt:
                cnt += 1
        nuevas_total[lbl] = cnt

    def _state_total_cell(l):
        n = len(opens[l])
        nv = nuevas_total.get(l, 0)
        if nv > 0:
            return f'<td class="num">{n} <span style="color:#16a34a;font-size:10px;font-weight:500" title="Nuevas creadas">(+{nv})</span></td>'
        return f'<td class="num">{n}</td>'
    state_total = ('<tr class="total"><td class="lbl-funnel"></td><td class="lbl-name">Total</td>'
                   + "".join(_state_total_cell(l) for l, _ in cutoffs) + '</tr>')
    chain_total = ('<tr class="total"><td class="lbl-name">Total</td>'
                   + "".join(f'<td class="num">{len(opens[l])}</td>' for l, _ in cutoffs) + '</tr>')

    today_total = len(opens[cutoffs[-1][0]])
    week_ago_total = len(opens[cutoffs[-6][0]]) if len(cutoffs) >= 6 else today_total
    dweek = today_total - week_ago_total
    arrow = "▲" if dweek > 0 else ("▼" if dweek < 0 else "→")
    dcolor = "#c0392b" if dweek > 0 else ("#1f8a4c" if dweek < 0 else "#7d8590")
    TALLER_STATUSES = set(TALLER_STATUSES_ORDER)
    en_taller = sum(
        1 for k, t in cache.get("tickets", {}).items()
        if t.get("current_status") in TALLER_STATUSES
    )

    estancadas_15 = sum(
        1 for k, t in cache.get("tickets", {}).items()
        if is_open(t.get("current_status"))
        and (days_since(last_status_change_dt(t)) or 0) > 15
    )

    # Pico = max histórico en taller (en lugar de max histórico de abiertas)
    peak = 0
    peak_lbl = ""
    for lbl, _ in cutoffs:
        n = sum(1 for _, _, s in opens[lbl] if s in TALLER_STATUSES)
        if n > peak:
            peak = n
            peak_lbl = lbl
    peak_pct = round(en_taller / peak * 100) if peak else 0
    peak_color = "#1f8a4c" if peak_pct <= 70 else ("#b8860b" if peak_pct <= 90 else "#c0392b")

    # Flechas in/out totales para el bloque "En taller"
    changes_today_kpi = compute_status_changes_today(cache)
    taller_in = sum(v["in"] for (st, _ch), v in changes_today_kpi.items() if st in TALLER_STATUSES)
    taller_out = sum(v["out"] for (st, _ch), v in changes_today_kpi.items() if st in TALLER_STATUSES)
    en_taller_flow_parts = []
    if taller_in > 0:
        en_taller_flow_parts.append(f'<span style="color:#c0392b;font-weight:500" title="entran hoy a taller">&#9650;{taller_in}</span>')
    if taller_out > 0:
        en_taller_flow_parts.append(f'<span style="color:#1f8a4c;font-weight:500" title="salen hoy de taller">&#9660;{taller_out}</span>')
    en_taller_flow = ('<div class="d" style="margin-top:2px;font-size:13px">'
                      + ' '.join(en_taller_flow_parts) + '</div>') if en_taller_flow_parts else ''

    # KPI Pendientes a taller: estados antes del taller (no asignados aún a técnico)
    PRE_TALLER_STATUSES = {"Abierto", "Recepcionado SAT", "Pendiente recogida", "Gestionado transporte"}
    BLOQ_TRUE = {"sí", "si", "yes", "true", "1"}
    pendientes_taller = 0
    pendientes_taller_bloq = 0
    for k, t in cache.get("tickets", {}).items():
        if t.get("current_status") not in PRE_TALLER_STATUSES:
            continue
        pendientes_taller += 1
        if (t.get("bloq") or "").strip().lower() in BLOQ_TRUE:
            pendientes_taller_bloq += 1

    # KPI Gestión externa
    EXTERNA_STATUSES = {"Enviado a técnico externo", "Esperando respuesta cliente a presupuesto", "Pendiente definir servicio externo"}
    gestion_externa = sum(
        1 for k, t in cache.get("tickets", {}).items()
        if t.get("current_status") in EXTERNA_STATUSES
    )

    # KPI Sustis solicitadas (pendientes de entregar) - del cache de sustis
    # KPI Salidas de taller hoy: transiciones hoy from=cualquier estado taller -> Devuelto/Resuelto/Finalizada
    SALIDA_DESTINOS = {"Devuelto a cliente", "Resuelto", "Finalizada"}
    salidas_taller = 0
    salidas_taller_tickets = set()
    _cutoff_hoy = datetime(today.year, today.month, today.day, 0, 0, 0, tzinfo=timezone.utc)
    for _kk, _t in cache.get("tickets", {}).items():
        for _tr in _t.get("transitions", []):
            if len(_tr) < 3: continue
            _ts, _fs, _to = _tr[0], _tr[1], _tr[2]
            if _fs not in TALLER_STATUSES: continue
            if _to not in SALIDA_DESTINOS: continue
            _dt = parse_iso(_ts)
            if not _dt or _dt < _cutoff_hoy: continue
            if _kk not in salidas_taller_tickets:
                salidas_taller_tickets.add(_kk)
                salidas_taller += 1

    sustis_solicitadas = 0
    try:
        from pathlib import Path as _P2
        for cand in [_P2("cache") / "sustis_activas.json", out_path.parent.parent / "cache" / "sustis_activas.json"]:
            if cand.exists():
                _s = json.loads(cand.read_text(encoding="utf-8"))
                sustis_solicitadas = sum(1 for it in (_s.get("items") or [])
                                          if (it.get("subtask_status") or "").lower() in ("solicitado",))
                break
    except Exception:
        pass

    # Nuevas creadas por corte (mensuales = todo el mes; diarios = ese dia)
    # KPI Nuevas hoy: tickets creados hoy Y abiertos (cuadra con bucket azul de la tabla Abiertas hoy)
    nuevas_hoy_count = sum(
        1 for k, t in cache.get("tickets", {}).items()
        if is_created_today(t) and is_open(t.get("current_status"))
    )

    # KPIs: Presupuestos hechos hoy y Equipos reparados hoy (= tickets únicos con to_status hoy)
    _cutoff_hoy_kpi = datetime(today.year, today.month, today.day, 0, 0, 0, tzinfo=timezone.utc)
    pres_hechos_hoy = set()
    eq_reparados_hoy = set()
    for _k, _t in cache.get("tickets", {}).items():
        for _tr in _t.get("transitions", []):
            if len(_tr) < 3: continue
            _ts, _fs, _to = _tr[0], _tr[1], _tr[2]
            _dt = parse_iso(_ts)
            if not _dt or _dt < _cutoff_hoy_kpi: continue
            if _to == "Presupuesto preparado pendiente de enviar":
                pres_hechos_hoy.add(_k)
            elif _to == "Inspección de salida":
                eq_reparados_hoy.add(_k)
    presupuestos_hechos_count = len(pres_hechos_hoy)
    equipos_reparados_count = len(eq_reparados_hoy)

    state_nuevas = ""  # Las nuevas se muestran en la fila Total entre paréntesis en verde

    chain_nuevas = ""  # ya no se usa (nuevas se muestran en cada celda chain_row)

    is_demo = "DEMO" in (cache.get("_meta", {}).get("note") or "")
    demo = ('<div class="banner">⚠ Vista previa con cache DEMO. El GitHub Action poblará el histórico completo.</div>'
            if is_demo else '')

    ah_header, ah_rows = build_abiertas_hoy_section(cache)
    avances_header, avances_rows, avances_detail_json, avances_extra = build_avances_section(cache)
    tec_header, tec_rows = build_tecnicos_section(cache, avances_extra)
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
        from sustis_etl import build_sustis_global_html
        sustis_html = build_sustis_global_html(sustis, inmov)
    except Exception as e:
        chains_html = f'<p style="color:#c0392b;padding:20px">Error generando Fase 2: {e}</p>'
        sustis_html = chains_html

    # Track record: guardar snapshot de KPIs hoy en cache/kpi_history.json (antes del render)
    history = {}
    try:
        hist_path = None
        for _cand in [Path("cache") / "kpi_history.json", out_path.parent.parent / "cache" / "kpi_history.json"]:
            if _cand.parent.exists():
                hist_path = _cand
                break
        if hist_path is not None:
            if hist_path.exists():
                try:
                    history = json.loads(hist_path.read_text(encoding="utf-8"))
                except Exception:
                    history = {}
            history[today.isoformat()] = {
                "ts": today_label,
                "abiertas_hoy": today_total,
                "nuevas_hoy": nuevas_hoy_count,
                "presupuestos_hechos": presupuestos_hechos_count,
                "equipos_reparados": equipos_reparados_count,
                "salidas_taller_hoy": salidas_taller,
                "pendientes_taller": pendientes_taller,
                "en_taller": en_taller,
                "peak_pct": peak_pct,
                "gestion_externa": gestion_externa,
                "sustis_solicitadas": sustis_solicitadas,
                "estancadas_15d": estancadas_15,
            }
            hist_path.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"[kpi_history] no se pudo guardar: {e}")

    # Generar tabla Evolución KPIs
    EVOL_KPI_COLS = [
        ("abiertas_hoy", "Abiertas", "#0b3d91"),
        ("nuevas_hoy", "Nuevas hoy", "#3b82f6"),
        ("presupuestos_hechos", "Presupuestos hechos", "#0891b2"),
        ("equipos_reparados", "Equipos reparados", "#059669"),
        ("salidas_taller_hoy", "Salidas taller", "#16a34a"),
        ("pendientes_taller", "Pdtes. taller", "#a23b72"),
        ("en_taller", "En taller", "#1f6feb"),
        ("peak_pct", "% pico taller", "#475569"),
        ("gestion_externa", "Gest. externa", "#cb6f0a"),
        ("sustis_solicitadas", "Sustis solic.", "#7c3aed"),
        ("estancadas_15d", "Estancadas >15d", "#c0392b"),
    ]
    dates_sorted = sorted(history.keys())
    if dates_sorted:
        ek_hdr = '<tr><th class="sticky-l1" style="text-align:left;left:0;min-width:120px">Fecha</th>'
        for _kk, _lbl, _col in EVOL_KPI_COLS:
            ek_hdr += f'<th style="color:#fff">{_lbl}</th>'
        ek_hdr += '</tr>'
        ek_rows_html = []
        for _date in dates_sorted:
            _row = history.get(_date, {})
            cells = [f'<td class="lbl-name" style="left:0;min-width:120px"><b>{_date}</b></td>']
            for _kk, _lbl, _col in EVOL_KPI_COLS:
                _v = _row.get(_kk, "")
                if _kk == "peak_pct" and _v != "":
                    cells.append(f'<td class="num" style="color:{_col};font-weight:500">{_v}%</td>')
                elif _v != "" and _v != 0:
                    cells.append(f'<td class="num" style="color:{_col};font-weight:500">{_v}</td>')
                else:
                    cells.append('<td class="num n0">·</td>')
            ek_rows_html.append('<tr>' + "".join(cells) + '</tr>')
        evolucion_kpis = ek_hdr + "\n".join(ek_rows_html)
    else:
        evolucion_kpis = '<tr><th colspan="9" style="color:#94a3b8;font-style:italic">Aún no hay historial. Se irá llenando a partir de hoy.</th></tr>'

    # Año venta x tipo averia (tabla con filtros cadena+importe)
    try:
        anual_averias_html = build_anual_averias_section(cache)
    except Exception as _e:
        anual_averias_html = f'<div style="color:#c0392b;padding:14px">Error tabla averias: {_e}</div>'

    html = template_path.read_text(encoding="utf-8")
    repl = {
        "__TODAY__": today_label,
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
        "__EN_TALLER_FLOW__": en_taller_flow,
        "__AH_HEADER__": ah_header,
        "__AH_ROWS__": ah_rows,
        "__TEC_HEADER__": tec_header,
        "__TEC_ROWS__": tec_rows,
        "__AVANCES_HEADER__": avances_header,
        "__AVANCES_ROWS__": avances_rows,
        "__AVANCES_DETAIL_JSON__": avances_detail_json,
        "__DETALLE_ROWS__": detalle_rows,
        "__STATE_HEADER__": state_header,
        "__STATE_ROWS__": state_rows,
        "__STATE_TOTAL__": state_total,
        "__STATE_NUEVAS__": state_nuevas,
        "__CHAIN_HEADER__": chain_header,
        "__CHAIN_ROWS__": chain_rows,
        "__CHAIN_TOTAL__": chain_total,
        "__CHAIN_NUEVAS__": chain_nuevas,
        "__PENDIENTES_TALLER__": str(pendientes_taller),
        "__PENDIENTES_TALLER_BLOQ__": str(pendientes_taller_bloq),
        "__GESTION_EXTERNA__": str(gestion_externa),
        "__SUSTIS_SOLICITADAS__": str(sustis_solicitadas),
        "__SALIDAS_TALLER__": str(salidas_taller),
        "__NUEVAS_HOY__": str(nuevas_hoy_count),
        "__PRES_HECHOS__": str(presupuestos_hechos_count),
        "__EQ_REPARADOS__": str(equipos_reparados_count),
        "__EVOLUCION_KPIS__": evolucion_kpis,
        "__ANUAL_AVERIAS__": anual_averias_html,
        "__CHAINS_HTML__": chains_html,
        "__SUSTIS_HTML__": sustis_html,
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
