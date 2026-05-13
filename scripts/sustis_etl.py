"""sustis_etl.py — Replica del Excel de Sustis (repo leaseir-dashboard-etl) para el dashboard web.

Genera 2 sub-vistas idénticas al Excel:
  - Resumen: bloques por cadena con tabla 14-col, backups permanentes pareados, stock SAT
  - Inmovilizado SAT: secciones A/B/C/D (Consolas, Manípulos, AHR, SRF)
"""
from __future__ import annotations
import re
from collections import defaultdict
from datetime import datetime, timezone

# ============================================================================
# Helpers de matching (portados del ETL)
# ============================================================================
CONSOLE_PREFIXES = ('MHR', 'AHR', 'MHP', 'HP', 'H', 'C')
ACTIVE_STATES = ('Equipo enviado', 'En préstamo')

CHAIN_ORDER_ETL = ['Elha', 'Sin Vello', 'Dermasana', 'Laser Factory',
                   'Epil Point', 'Smart Duck', 'Centros Unico Italia', 'Others']
ITALIA_CITIES = ['giugliano', 'ostia', 'orio', 'lecce', 'caserta', 'pescara',
                 'castellamare', 'andria', 'avellino', 'novara', 'palermo', 'roma',
                 'milano', 'napoli', 'colli aminei', 'chiaia', 'torino', 'bologna',
                 'firenze', 'genova', 'bari', 'catania']


def clean_serial(val):
    if val is None:
        return None
    s = re.sub(r'[\s\xa0​ ]+', '', str(val)).strip().upper()
    return None if s in ('NAN', 'NONE', '', '-') else s


def strip_prefix(s):
    if not s:
        return s
    for p in CONSOLE_PREFIXES:
        if s.startswith(p) and len(s) > len(p) and s[len(p)].isdigit():
            return s[len(p):]
    return s


def strip_bis(s):
    if not s:
        return s
    return s[:-3] if s.endswith('BIS') else s


def normalized_key(s):
    if not s:
        return None
    s = clean_serial(s)
    if not s:
        return None
    if s.startswith('SUST') or s.startswith('DEMO'):
        s2 = strip_bis(s)
        m = re.match(r'^([A-Z]+)(\d+)([A-Z]*)$', s2)
        if m:
            prefix, digits, suff = m.groups()
            s2 = prefix + (digits.lstrip('0') or '0') + suff
        return s2
    s2 = strip_bis(strip_prefix(s))
    m = re.match(r'^(\d+)([A-Z]*)$', s2)
    if m:
        digits, suff = m.groups()
        s2 = (digits.lstrip('0') or '0') + suff
    return s2


def all_keys(s):
    s = clean_serial(s)
    if not s:
        return []
    keys = {s, strip_bis(s)}
    nk = normalized_key(s)
    if nk:
        keys.add(nk)
    if not s.startswith('SUST') and not s.startswith('DEMO'):
        keys.add(strip_prefix(s))
        keys.add(strip_prefix(strip_bis(s)))
    return [k for k in keys if k]


def split_aliases(text):
    if not text:
        return []
    out = []
    for p in re.split(r'[,;\n/]+', str(text)):
        c = clean_serial(p)
        if c:
            out.append(c)
    return out


def core_num(s):
    if not s:
        return None
    ss = re.sub(r'\s+', '', s.upper())
    for p in CONSOLE_PREFIXES:
        if ss.startswith(p) and len(ss) > len(p) and ss[len(p)].isdigit():
            ss = ss[len(p):]
            break
    m = re.match(r'^(\d+)', ss)
    if m:
        return m.group(1).lstrip('0') or '0'
    return ss


def chain_for(cliente, loc=''):
    cl = (cliente or '').lower()
    ll = (loc or '').lower()
    if 'elha' in cl:
        return 'Elha'
    if 'sin vello' in cl or 'sinvello' in cl:
        return 'Sin Vello'
    if 'dermasana' in cl:
        return 'Dermasana'
    if 'laser factory' in cl or 'láser factory' in cl:
        return 'Laser Factory'
    if 'epil point' in cl:
        return 'Epil Point'
    if 'smart duck' in cl:
        return 'Smart Duck'
    if 'beauty cool' in cl:
        return 'Centros Unico Italia'
    if 'centri unico' in cl:
        return 'Centros Unico Italia'
    if ('unico' in cl or 'único' in cl):
        if any(city in cl or city in ll for city in ITALIA_CITIES):
            return 'Centros Unico Italia'
    return 'Others'


def parse_iso(s):
    if not s:
        return None
    s2 = str(s).replace('Z', '+00:00')
    s2 = re.sub(r'([+\-]\d{2})(\d{2})$', r'\1:\2', s2)
    try:
        return datetime.fromisoformat(s2)
    except ValueError:
        return None


def html_escape(s):
    if s is None:
        return ""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


# ============================================================================
# Indexes + matching
# ============================================================================
def build_idx(records):
    """records: list of dict con keys 'serial', 'aliases' (list)
    Devuelve (serial_idx, alias_idx) donde:
      serial_idx[key] -> list of records que tienen ese key como serial
      alias_idx[key]  -> list of (record, alias_str) que tienen ese key como alias
    """
    serial_idx = defaultdict(list)
    alias_idx = defaultdict(list)
    for r in records:
        if r.get('serial'):
            for k in all_keys(r['serial']):
                serial_idx[k].append(r)
        for a in r.get('aliases', []):
            for k in all_keys(a):
                alias_idx[k].append((r, a))
    return dict(serial_idx), dict(alias_idx)


def find_match(jira_serial, sat_serial_idx, sat_alias_idx):
    if not jira_serial:
        return None, None, None
    keys = all_keys(jira_serial)
    # Match directo
    for k in keys:
        if k in sat_serial_idx:
            for r in sat_serial_idx[k]:
                if r['serial'] == jira_serial:
                    return r, 'Direct', ''
    # Match con prefix-stripped (normalized)
    nk = normalized_key(jira_serial)
    if nk and nk in sat_serial_idx:
        r = sat_serial_idx[nk][0]
        return r, 'Prefix-stripped', f"AT={r['serial']}={jira_serial}"
    for k in keys:
        if k in sat_serial_idx:
            r = sat_serial_idx[k][0]
            return r, 'Prefix-stripped', f"AT={r['serial']}={jira_serial}"
    # Match por alias (Renombrada)
    for k in keys:
        if k in sat_alias_idx:
            r, a = sat_alias_idx[k][0]
            return r, 'Renombrada alias', f"alias={a}"
    return None, None, None


# ============================================================================
# Transformación de datos
# ============================================================================
def normalize_at_record(item):
    """Convierte un item del cache airtable_inmovilizado.json al formato del ETL."""
    return {
        'rec_id': item.get('id_rec', ''),
        'serial': clean_serial(item.get('serial')),
        'type': item.get('type_asset', ''),
        'console': item.get('console_model', ''),
        'spot': item.get('spot_size', ''),
        'activity': item.get('activity', ''),
        'customer': item.get('customer', ''),
        'aliases': split_aliases(item.get('renombrada', '')),
    }


def normalize_subtask(item):
    """Convierte un item del cache sustis_activas.json al formato del ETL."""
    return {
        'key': item.get('key', ''),
        'status': item.get('subtask_status', ''),
        'parent_key': item.get('parent_key', ''),
        'parent_status': item.get('parent_status', ''),
        'consola': clean_serial(item.get('consola_susti')),
        'handpiece': clean_serial(item.get('manipulo_susti')),
        'cliente': item.get('cliente', ''),
        'localizacion': item.get('loc', ''),
        'fecha_envio': parse_iso(item.get('fecha_envio')),
        'consola_av': item.get('consola_averiada', '') or '',
        'handpiece_av': item.get('manipulo_averiado', '') or '',
    }


def build_data(sustis_items, inmov_items):
    """Replica el cruce del ETL. Devuelve un dict con todas las estructuras
    necesarias para renderizar las 2 vistas."""
    subtasks = [normalize_subtask(s) for s in sustis_items]
    at_sat = [normalize_at_record(r) for r in inmov_items]

    # Build SAT indexes
    sat_serial_idx, sat_alias_idx = build_idx(at_sat)

    # Cruce sub-tasks <-> AT
    union_rows = []
    solo_jira = []
    sub_to_unions = defaultdict(list)
    sub_to_solojira = defaultdict(list)
    for s in subtasks:
        for kind, ser in [('consola', s['consola']), ('handpiece', s['handpiece'])]:
            if not ser:
                continue
            rec, method, detail = find_match(ser, sat_serial_idx, sat_alias_idx)
            base = {
                'jira_serial': ser, 'jira_kind': kind, 'sub_key': s['key'],
                'sub_status': s['status'], 'cliente': s['cliente'],
                'localizacion': s['localizacion'], 'parent_key': s['parent_key'],
                'parent_status': s['parent_status'],
                'consola_av_padre': s['consola_av'], 'handpiece_av_padre': s['handpiece_av'],
                'fecha_envio': s['fecha_envio'],
            }
            if rec:
                act = rec['activity'] or ''
                base.update({
                    'at_serial': rec['serial'], 'at_rec_id': rec['rec_id'],
                    'at_type': rec['type'], 'at_console': rec['console'],
                    'at_spot': rec['spot'], 'at_activity': act,
                    'at_customer': rec['customer'], 'method': method,
                })
                union_rows.append(base)
                sub_to_unions[s['key']].append(base)
            else:
                solo_jira.append(base)
                sub_to_solojira[s['key']].append(base)

    matched_recids = {u['at_rec_id'] for u in union_rows}
    solo_at = [r for r in at_sat if r['rec_id'] not in matched_recids]

    # Duplicados — solo el más antiguo
    serial_to_subs = defaultdict(list)
    for u in union_rows + solo_jira:
        serial_to_subs[u['jira_serial']].append(u)
    duplicate_groups = {ser: lst for ser, lst in serial_to_subs.items()
                        if len({u['sub_key'] for u in lst}) > 1}
    old_dup_pairs = set()  # set of (serial, sub_key) marked as old dup
    for ser, lst in duplicate_groups.items():
        sub_keys_sorted = sorted({u['sub_key'] for u in lst},
                                 key=lambda k: int(k.split('-')[1]) if '-' in k and k.split('-')[1].isdigit() else 0)
        for sk in sub_keys_sorted[:-1]:
            old_dup_pairs.add((ser, sk))

    matched_at_ids_norm = {re.sub(r'\s+', '', str(u['at_serial'] or '').upper())
                           for u in union_rows}

    serials_with_active = {u['jira_serial'] for u in union_rows + solo_jira
                           if u.get('sub_status') in ACTIVE_STATES}

    return {
        'subtasks': subtasks,
        'union_rows': union_rows,
        'solo_jira': solo_jira,
        'solo_at': solo_at,
        'sub_to_unions': sub_to_unions,
        'sub_to_solojira': sub_to_solojira,
        'old_dup_pairs': old_dup_pairs,
        'matched_at_ids_norm': matched_at_ids_norm,
        'serials_with_active': serials_with_active,
        'at_sat': at_sat,
    }


# ============================================================================
# Generación de filas para la vista "Resumen" (por cadena)
# ============================================================================
def build_subtask_row(s, sub_to_unions, sub_to_solojira, old_dup_pairs, now_utc):
    """Genera una fila por sub-task con todos los campos para la vista Resumen."""
    uns = sub_to_unions.get(s['key'], [])
    sjs = sub_to_solojira.get(s['key'], [])
    cmatch = next((u for u in uns if u['jira_kind'] == 'consola'), None)
    hmatch = next((u for u in uns if u['jira_kind'] == 'handpiece'), None)
    csolojira = next((u for u in sjs if u['jira_kind'] == 'consola'), None)
    hsolojira = next((u for u in sjs if u['jira_kind'] == 'handpiece'), None)
    consola_id = cmatch['at_serial'] if cmatch else (s['consola'] or '')
    handpiece_id = hmatch['at_serial'] if hmatch else (s['handpiece'] or '')
    modelo_parts = []
    if cmatch and cmatch.get('at_console'):
        modelo_parts.append(cmatch['at_console'])
    if hmatch and hmatch.get('at_spot'):
        modelo_parts.append(hmatch['at_spot'])
    modelo = ' / '.join(modelo_parts)
    acts = [u.get('at_activity', '') for u in [cmatch, hmatch] if u and u.get('at_activity')]
    activity = ' / '.join(acts) if len(set(acts)) > 1 else (acts[0] if acts else '')
    dias = ''
    fecha_str = ''
    if s['fecha_envio']:
        try:
            fecha_aware = s['fecha_envio'] if s['fecha_envio'].tzinfo else s['fecha_envio'].replace(tzinfo=timezone.utc)
            dias = (now_utc - fecha_aware).days
            fecha_str = s['fecha_envio'].strftime('%Y-%m-%d')
        except Exception:
            pass
    is_solojira = bool(csolojira or hsolojira)
    is_old_dup = False
    if s['consola'] and (s['consola'], s['key']) in old_dup_pairs:
        is_old_dup = True
    if s['handpiece'] and (s['handpiece'], s['key']) in old_dup_pairs:
        is_old_dup = True
    return {
        'cliente': s['cliente'] or '',
        'localizacion': s['localizacion'] or '',
        'consola_id': consola_id,
        'handpiece_id': handpiece_id,
        'modelo': modelo,
        'leas_principal': s['parent_key'] or '',
        'estado_principal': s['parent_status'] or '',
        'sub_key': s['key'],
        'sub_status': s['status'] or '',
        'consola_av': s['consola_av'],
        'handpiece_av': s['handpiece_av'],
        'fecha_envio': fecha_str,
        'dias': dias,
        'activity': activity,
        'is_solojira': is_solojira,
        'is_old_dup': is_old_dup,
    }


# ============================================================================
# HTML rendering
# ============================================================================
JIRA_URL = "https://leaseir.atlassian.net/browse/"

# Colores (ajustados a los del CSS del dashboard)
COLOR_RED_BG = "#fed4d4"
COLOR_YELLOW_BG = "#fff2a8"
COLOR_GREEN_BG = "#d6f0d7"
COLOR_BLOCK_HEADER = "#365e7d"
COLOR_SECTION_HEADER = "#1f3a5f"
COLOR_SUBSEC_HEADER = "#6c7d94"
COLOR_COL_HEADER = "#2e54a3"


HEADERS_14 = ['Cliente', 'Localización', 'Num Serie consola susti',
              'Num Serie manípulo susti', 'Modelo',
              'LEAS Incidencia principal', 'Estado principal',
              'Incidencia subtarea', 'Estado subtarea',
              'Consola averiada', 'Manípulo averiado',
              'Fecha y hora de envío', 'Días con sustitución',
              'Current Activity (Airtable)']

INV_HEADERS = ['Num. Serie', 'Modelo', 'Cliente', 'Incidencia (sub-task)',
               'Incidencia principal', 'Estado principal',
               'Consola principal (avería)', 'HP principal (avería)',
               'Estado Jira (sub-task)', 'Localización equipo averiado',
               'Fecha y hora de envío', 'Días con sustitución',
               'Current Activity (Airtable)']


def render_subtask_row_html(sr):
    """Renderiza una fila <tr> para la tabla de cadena (14 cols)."""
    bg = ""
    if sr['is_solojira']:
        bg = f' style="background:{COLOR_RED_BG}"'
    elif sr['is_old_dup']:
        bg = f' style="background:{COLOR_YELLOW_BG}"'
    sub_link = (f'<a href="{JIRA_URL}{sr["sub_key"]}" target="_blank" '
                f'style="color:#2a59c4;text-decoration:none">{sr["sub_key"]}</a>'
                if sr['sub_key'] else '')
    leas_link = (f'<a href="{JIRA_URL}{sr["leas_principal"]}" target="_blank" '
                 f'style="color:#2a59c4;text-decoration:none">{sr["leas_principal"]}</a>'
                 if sr['leas_principal'] else '')
    cells = [
        html_escape(sr['cliente']),
        html_escape(sr['localizacion']),
        html_escape(sr['consola_id']),
        html_escape(sr['handpiece_id']),
        html_escape(sr['modelo']),
        leas_link,
        html_escape(sr['estado_principal']),
        sub_link,
        html_escape(sr['sub_status']),
        html_escape(sr['consola_av']),
        html_escape(sr['handpiece_av']),
        html_escape(sr['fecha_envio']),
        str(sr['dias']) if sr['dias'] != '' else '',
        html_escape(sr['activity']),
    ]
    return f'<tr{bg}>' + ''.join(f'<td>{c}</td>' for c in cells) + '</tr>'


def render_paired_backups(by_cust):
    """Renderiza filas para backups permanentes pareando consola+manípulo por core_num."""
    rows = []
    for cust in sorted(by_cust.keys()):
        lst = by_cust[cust]
        consoles = [r for r in lst if r['type'] == 'Console']
        hps = [r for r in lst if r['type'] == 'Handpiece']
        paired_hps = set()
        for cons in consoles:
            cn = core_num(cons['serial'])
            pair = next((h for h in hps if core_num(h['serial']) == cn
                         and id(h) not in paired_hps), None)
            if pair:
                paired_hps.add(id(pair))
                modelo = ' / '.join(filter(None, [cons.get('console'), pair.get('spot')]))
                act = cons['activity'] if cons['activity'] == pair['activity'] \
                    else f"{cons['activity']} / {pair['activity']}"
                rows.append({
                    'cliente': cust, 'localizacion': '',
                    'consola_id': cons['serial'], 'handpiece_id': pair['serial'],
                    'modelo': modelo, 'activity': act,
                })
            else:
                rows.append({
                    'cliente': cust, 'localizacion': '',
                    'consola_id': cons['serial'], 'handpiece_id': '',
                    'modelo': cons.get('console', '') or '',
                    'activity': cons['activity'] or '',
                })
        for hp in hps:
            if id(hp) in paired_hps:
                continue
            rows.append({
                'cliente': cust, 'localizacion': '',
                'consola_id': '', 'handpiece_id': hp['serial'],
                'modelo': hp.get('spot', '') or '',
                'activity': hp['activity'] or '',
            })
    return rows


def render_backup_row_html(b):
    """Fila de backup permanente (subset de 14 cols, en blanco las que no aplican)."""
    cells = [
        html_escape(b['cliente']),
        html_escape(b['localizacion']),
        html_escape(b['consola_id']),
        html_escape(b['handpiece_id']),
        html_escape(b['modelo']),
        '', '', '', '', '', '', '', '',
        html_escape(b['activity']),
    ]
    return '<tr>' + ''.join(f'<td>{c}</td>' for c in cells) + '</tr>'


def render_disponible_row_html(r):
    """Fila para Stock SAT (6 cols, fondo verde)."""
    cells = [
        html_escape(r['serial']),
        html_escape(r['type']),
        html_escape(r.get('console') or ''),
        html_escape(r.get('spot') or ''),
        html_escape(r.get('customer') or ''),
        html_escape(r['activity'] or ''),
    ]
    return (f'<tr style="background:{COLOR_GREEN_BG}">'
            + ''.join(f'<td>{c}</td>' for c in cells) + '</tr>')


def render_resumen_html(data):
    """Renderiza la vista Resumen (bloques por cadena + backups + Stock SAT)."""
    now_utc = datetime.now(timezone.utc)
    subtasks = data['subtasks']
    sub_to_unions = data['sub_to_unions']
    sub_to_solojira = data['sub_to_solojira']
    old_dup_pairs = data['old_dup_pairs']
    at_sat = data['at_sat']
    solo_at = data['solo_at']
    matched_at_ids_norm = data['matched_at_ids_norm']

    # Filas por sub-task agrupadas por cadena
    all_rows = [build_subtask_row(s, sub_to_unions, sub_to_solojira, old_dup_pairs, now_utc)
                for s in subtasks]
    chain_to_rows = defaultdict(list)
    for sr in all_rows:
        chain_to_rows[chain_for(sr['cliente'])].append(sr)

    # Backups por cadena vs others
    backups_per_chain = defaultdict(list)
    backups_others = []
    for r in at_sat:
        if r['activity'] != 'Backups for customers':
            continue
        sid = re.sub(r'\s+', '', str(r['serial'] or '').upper())
        if sid in matched_at_ids_norm:
            continue
        ch = chain_for(r['customer'])
        if ch == 'Others' or not r['customer']:
            backups_others.append(r)
        else:
            backups_per_chain[ch].append(r)

    # Disponibles (Stock SAT)
    disponibles = [r for r in solo_at if r['activity'] == 'SAT']
    mhr_console = [r for r in disponibles
                   if r['type'] == 'Console' and (r.get('console') or '') not in ('AHR', 'SRF')]
    hps_main = [r for r in disponibles
                if r['type'] == 'Handpiece' and (r.get('spot') or '') in ('Single', 'Dual', 'Quad')]
    ahr = [r for r in disponibles if r['type'] == 'Console' and r.get('console') == 'AHR']
    srf = [r for r in disponibles if (r['type'] == 'Console' and r.get('console') == 'SRF')
           or (r['type'] == 'Handpiece' and r.get('spot') == 'SRF')]

    out = []
    table_style = ('font-size:11px;min-width:1500px;width:100%;'
                   'border-collapse:collapse;margin-bottom:20px')
    th_style = (f'background:{COLOR_COL_HEADER};color:white;padding:6px 10px;'
                'font-weight:500;text-align:left;border:1px solid #d0d7de')
    block_hdr_style = (f'background:{COLOR_BLOCK_HEADER};color:white;'
                       'padding:8px 12px;font-weight:600;font-size:13px;'
                       'margin:18px 0 0;border-radius:4px 4px 0 0')
    section_hdr_style = (f'background:{COLOR_SECTION_HEADER};color:white;'
                         'padding:8px 12px;font-weight:600;font-size:14px;'
                         'margin:24px 0 0;border-radius:4px 4px 0 0')
    subsec_hdr_style = (f'background:{COLOR_SUBSEC_HEADER};color:white;'
                        'padding:5px 10px;font-style:italic;font-size:11px;'
                        'margin:0;border-radius:0')
    total_style = ('background:#fff4d6;padding:5px 10px;font-weight:500;'
                   'font-size:11px;margin:0 0 8px;border-radius:0 0 4px 4px')

    # ===== Por cadena =====
    for chain in CHAIN_ORDER_ETL:
        chain_rows = chain_to_rows.get(chain, [])
        bks = backups_per_chain.get(chain, [])
        if chain == 'Others' and not chain_rows and not bks:
            continue
        out.append(f'<div style="{block_hdr_style}">CADENA — {chain.upper()}</div>')
        out.append('<div class="scroller" style="margin:0 0 4px;max-height:none">')
        out.append(f'<table style="{table_style}">')
        out.append('<tr>' + ''.join(f'<th style="{th_style}">{h}</th>' for h in HEADERS_14) + '</tr>')
        if chain_rows:
            for sr in chain_rows:
                out.append(render_subtask_row_html(sr))
        else:
            out.append(f'<tr><td colspan="14" style="color:#7d8590;font-style:italic">'
                       '(Sin tickets activos para esta cadena)</td></tr>')
        out.append('</table>')
        out.append('</div>')

        if bks:
            out.append(f'<div style="{subsec_hdr_style}">· Backups permanentes ({len(bks)})</div>')
            by_cust = defaultdict(list)
            for r in bks:
                by_cust[r.get('customer') or ''].append(r)
            paired = render_paired_backups(by_cust)
            if paired:
                out.append('<div class="scroller" style="margin:0;max-height:none">')
                out.append(f'<table style="{table_style}">')
                for b in paired:
                    out.append(render_backup_row_html(b))
                out.append('</table>')
                out.append('</div>')

        out.append(f'<div style="{total_style}">Total {chain}: '
                   f'{len(chain_rows)} ticket(s) + {len(bks)} backup(s)</div>')

    # ===== Backups permanentes — resto de clientes =====
    out.append(f'<div style="{section_hdr_style}">BACKUPS PERMANENTES — resto de clientes</div>')
    out.append('<div class="scroller" style="margin:0 0 4px;max-height:none">')
    out.append(f'<table style="{table_style}">')
    out.append('<tr>' + ''.join(f'<th style="{th_style}">{h}</th>' for h in HEADERS_14) + '</tr>')
    by_cust = defaultdict(list)
    for r in backups_others:
        by_cust[r.get('customer') or '(sin cliente)'].append(r)
    paired = render_paired_backups(by_cust)
    equipos_n = 0
    for b in paired:
        out.append(render_backup_row_html(b))
        equipos_n += (1 if not b['consola_id'] or not b['handpiece_id'] else 2)
    out.append('</table>')
    out.append('</div>')
    out.append(f'<div style="{total_style}">Total backups permanentes (others): '
               f'{len(paired)} fila(s) ({equipos_n} equipos)</div>')

    # ===== Equipos disponibles (Stock SAT) =====
    out.append(f'<div style="{section_hdr_style}">EQUIPOS DISPONIBLES (Stock SAT)</div>')
    out.append('<div class="scroller" style="margin:0 0 4px;max-height:none">')
    out.append(f'<table style="{table_style}">')
    disp_hdrs = ['Num Serie', 'Tipo', 'Console', 'Spot', 'Cliente', 'Activity']
    out.append('<tr>' + ''.join(f'<th style="{th_style}">{h}</th>' for h in disp_hdrs) + '</tr>')
    total_disp = 0
    ordered_main = (sorted(mhr_console, key=lambda x: x['serial'] or '')
                    + sorted([r for r in hps_main if r.get('spot') == 'Single'], key=lambda x: x['serial'] or '')
                    + sorted([r for r in hps_main if r.get('spot') == 'Dual'], key=lambda x: x['serial'] or '')
                    + sorted([r for r in hps_main if r.get('spot') == 'Quad'], key=lambda x: x['serial'] or ''))
    for r in ordered_main:
        out.append(render_disponible_row_html(r))
        total_disp += 1
    if ahr:
        out.append(f'<tr><td colspan="6" style="{subsec_hdr_style}">— AHR —</td></tr>')
        for r in sorted(ahr, key=lambda x: x['serial'] or ''):
            out.append(render_disponible_row_html(r))
            total_disp += 1
    if srf:
        out.append(f'<tr><td colspan="6" style="{subsec_hdr_style}">— SRF —</td></tr>')
        for r in sorted(srf, key=lambda x: x['serial'] or ''):
            out.append(render_disponible_row_html(r))
            total_disp += 1
    out.append('</table>')
    out.append('</div>')
    out.append(f'<div style="{total_style}">Total disponibles: {total_disp} equipos</div>')

    return '\n'.join(out)


# ============================================================================
# Vista Inmovilizado SAT (secciones A/B/C/D)
# ============================================================================
SECTION_TITLES = {
    'A': 'A — Consolas (MHR y otras, excl. AHR y SRF)',
    'B': 'B — Manípulos (Quad / Dual / Single, excl. SRF)',
    'C': 'C — Consolas AHR',
    'D': 'D — Tecnología SRF (consolas + manípulos)',
}


def section_for(r):
    if r.get('console') == 'SRF' or r.get('spot') == 'SRF':
        return 'D'
    if r['type'] == 'Console':
        return 'C' if r.get('console') == 'AHR' else 'A'
    return 'B'


def render_inv_row_html(r, color_bg=None):
    style = f' style="background:{color_bg}"' if color_bg else ''
    sub_link = (f'<a href="{JIRA_URL}{r["sub_key"]}" target="_blank" '
                f'style="color:#2a59c4;text-decoration:none">{r["sub_key"]}</a>'
                if r.get('sub_key') else '')
    leas_link = (f'<a href="{JIRA_URL}{r["leas_principal"]}" target="_blank" '
                 f'style="color:#2a59c4;text-decoration:none">{r["leas_principal"]}</a>'
                 if r.get('leas_principal') else '')
    cells = [
        html_escape(r['num_serie']),
        html_escape(r['modelo']),
        html_escape(r['cliente']),
        sub_link,
        leas_link,
        html_escape(r['estado_principal']),
        html_escape(r['consola_av']),
        html_escape(r['handpiece_av']),
        html_escape(r['sub_status']),
        html_escape(r['localizacion']),
        html_escape(r['fecha_envio']),
        str(r['dias']) if r['dias'] != '' else '',
        html_escape(r['activity']),
    ]
    return f'<tr{style}>' + ''.join(f'<td>{c}</td>' for c in cells) + '</tr>'


def render_inmovilizado_sat_html(data):
    """Renderiza la vista Inmovilizado SAT (4 secciones A/B/C/D)."""
    now_utc = datetime.now(timezone.utc)
    union_rows = data['union_rows']
    solo_jira = data['solo_jira']
    solo_at = data['solo_at']
    old_dup_pairs = data['old_dup_pairs']
    serials_with_active = data['serials_with_active']

    inventory = []
    for u in union_rows:
        fecha = u['fecha_envio']
        dias = ''
        fecha_str = ''
        if fecha:
            try:
                fecha_aware = fecha if fecha.tzinfo else fecha.replace(tzinfo=timezone.utc)
                dias = (now_utc - fecha_aware).days
                fecha_str = fecha.strftime('%Y-%m-%d')
            except Exception:
                pass
        modelo = u.get('at_console') if u.get('at_type') == 'Console' else u.get('at_spot')
        inventory.append({
            'kind': 'union',
            'num_serie': u['at_serial'],
            'modelo': modelo or '',
            'cliente': u['cliente'] or '',
            'sub_key': u['sub_key'],
            'leas_principal': u['parent_key'] or '',
            'estado_principal': u['parent_status'] or '',
            'consola_av': u['consola_av_padre'],
            'handpiece_av': u['handpiece_av_padre'],
            'sub_status': u['sub_status'] or '',
            'localizacion': u['localizacion'] or '',
            'fecha_envio': fecha_str,
            'dias': dias,
            'activity': u.get('at_activity', '') or '',
            'type': u.get('at_type'),
            'console': u.get('at_console'),
            'spot': u.get('at_spot'),
            'jira_serial': u['jira_serial'],
        })
    for j in solo_jira:
        fecha = j['fecha_envio']
        dias = ''
        fecha_str = ''
        if fecha:
            try:
                fecha_aware = fecha if fecha.tzinfo else fecha.replace(tzinfo=timezone.utc)
                dias = (now_utc - fecha_aware).days
                fecha_str = fecha.strftime('%Y-%m-%d')
            except Exception:
                pass
        inventory.append({
            'kind': 'solo_jira',
            'num_serie': j['jira_serial'],
            'modelo': j['jira_kind'],
            'cliente': j['cliente'] or '',
            'sub_key': j['sub_key'],
            'leas_principal': j['parent_key'] or '',
            'estado_principal': j['parent_status'] or '',
            'consola_av': j['consola_av_padre'],
            'handpiece_av': j['handpiece_av_padre'],
            'sub_status': j['sub_status'] or '',
            'localizacion': j['localizacion'] or '',
            'fecha_envio': fecha_str,
            'dias': dias,
            'activity': '(no en AT)',
            'type': 'Console' if j['jira_kind'] == 'consola' else 'Handpiece',
            'console': None,
            'spot': None,
            'jira_serial': j['jira_serial'],
        })
    for r in solo_at:
        inventory.append({
            'kind': 'solo_at',
            'num_serie': r['serial'],
            'modelo': (r.get('console') if r.get('type') == 'Console' else r.get('spot')) or '',
            'cliente': r.get('customer') or '',
            'sub_key': '',
            'leas_principal': '',
            'estado_principal': '',
            'consola_av': '',
            'handpiece_av': '',
            'sub_status': 'Sin ticket activo',
            'localizacion': '',
            'fecha_envio': '',
            'dias': '',
            'activity': r.get('activity') or '',
            'type': r.get('type'),
            'console': r.get('console'),
            'spot': r.get('spot'),
            'jira_serial': None,
        })

    for r in inventory:
        r['section'] = section_for(r)

    out = []
    section_hdr_style = (f'background:{COLOR_SECTION_HEADER};color:white;'
                        'padding:8px 12px;font-weight:600;font-size:14px;'
                        'margin:18px 0 0;border-radius:4px 4px 0 0')
    table_style = ('font-size:11px;min-width:1500px;width:100%;'
                   'border-collapse:collapse;margin-bottom:12px')
    th_style = (f'background:{COLOR_COL_HEADER};color:white;padding:6px 10px;'
                'font-weight:500;text-align:left;border:1px solid #d0d7de')

    out.append('<p class="chain-meta" style="margin:8px 0">'
               f'{len(inventory)} equipos · '
               f'<span style="background:{COLOR_RED_BG};padding:1px 6px">Solo Jira</span> · '
               f'<span style="background:{COLOR_YELLOW_BG};padding:1px 6px">Duplicado antiguo</span> · '
               f'<span style="background:{COLOR_GREEN_BG};padding:1px 6px">Disponible/Devuelto</span> · '
               'blanco: prestado activo</p>')

    for sec in ['A', 'B', 'C', 'D']:
        sec_rows = [r for r in inventory if r['section'] == sec]
        if not sec_rows:
            continue
        out.append(f'<div style="{section_hdr_style}">{SECTION_TITLES[sec]} ({len(sec_rows)})</div>')
        out.append('<div class="scroller" style="margin:0 0 4px;max-height:none">')
        out.append(f'<table style="{table_style}">')
        out.append('<tr>' + ''.join(f'<th style="{th_style}">{h}</th>' for h in INV_HEADERS) + '</tr>')

        def sort_key(r):
            is_avail = (r['kind'] == 'solo_at' and r.get('activity') == 'SAT')
            d = r['dias'] if isinstance(r['dias'], int) else -1
            return (1 if is_avail else 0, -d, r['num_serie'] or '')

        for r in sorted(sec_rows, key=sort_key):
            activity = r.get('activity') or ''
            serial = r.get('num_serie')
            jira_serial = r.get('jira_serial')
            has_active_elsewhere = (jira_serial in serials_with_active) and not (
                r['kind'] == 'union' and r['sub_status'] in ACTIVE_STATES)
            bg = None
            if r['kind'] == 'solo_jira':
                bg = COLOR_RED_BG
            elif (jira_serial, r['sub_key']) in old_dup_pairs:
                bg = COLOR_YELLOW_BG
            elif r['kind'] == 'solo_at' and activity == 'SAT' and not has_active_elsewhere:
                # Verde solo cuando el equipo está físicamente de vuelta en SAT
                # (sub-task ya cerrada, no aparece en cache) y no tiene otro ticket activo.
                bg = COLOR_GREEN_BG
            out.append(render_inv_row_html(r, bg))
        out.append('</table>')
        out.append('</div>')

    return '\n'.join(out)


# ============================================================================
# Entry point
# ============================================================================
def build_sustis_global_html(sustis, inmov):
    """HTML para la pestaña Sustis con 2 sub-vistas: Resumen + Inmovilizado SAT."""
    sustis_items = (sustis or {}).get('items', []) if sustis else []
    inmov_items = (inmov or {}).get('items', []) if inmov else []
    if not sustis_items and not inmov_items:
        return '<p style="color:#7d8590">No hay datos de sustituciones disponibles.</p>'

    data = build_data(sustis_items, inmov_items)
    resumen_html = render_resumen_html(data)
    inmov_html = render_inmovilizado_sat_html(data)

    out = []
    out.append('<div class="tabs2 sustis-tabs">')
    out.append('<button type="button" data-sustis="cadenas" class="active">Resumen por cadena</button>')
    out.append('<button type="button" data-sustis="inmovilizado">Inmovilizado SAT (A/B/C/D)</button>')
    out.append('</div>')
    out.append(f'<div class="sustis-pane active" data-sustis="cadenas">{resumen_html}</div>')
    out.append(f'<div class="sustis-pane" data-sustis="inmovilizado">{inmov_html}</div>')
    return ''.join(out)
