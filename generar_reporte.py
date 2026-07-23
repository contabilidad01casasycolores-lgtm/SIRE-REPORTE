#!/usr/bin/env python3
"""
Cruza SIRE (TXT con pipes |) vs SAP (Excel) y genera index.html.
Uso: python generar_reporte.py SIRE.txt SAP.xlsx
Tipos considerados: 01, 07, 08, 30, 42 — solo mes actual
"""
import sys, os, json
from datetime import datetime
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    print("❌ pip install openpyxl"); sys.exit(1)

if len(sys.argv) < 3:
    print("Uso: python generar_reporte.py SIRE.txt SAP.xlsx"); sys.exit(0)

SIRE_FILE = sys.argv[1]
SAP_FILE  = sys.argv[2]

for f in [SIRE_FILE, SAP_FILE]:
    if not os.path.exists(f):
        print(f"❌ No encontrado: {f}"); sys.exit(1)

print(f"📂 SIRE: {SIRE_FILE}")
print(f"📂 SAP:  {SAP_FILE}")

# ── Índices TXT SIRE (separado por |) ────────────────────────────────────────
T_PERIODO=2; T_FECHA=4; T_TIPO=6; T_SERIE=7; T_NRO=9
T_RUC=12; T_RAZON=13; T_BI=14; T_IGV=15; T_NG=20; T_TOTAL=24; T_MONEDA=25
T_ORIG_FECHA=27; T_ORIG_TIPO=28; T_ORIG_SERIE=29; T_ORIG_NRO=31

# ── Índices Excel SAP ─────────────────────────────────────────────────────────
A_TIPO=5; A_SERIE=6; A_NRO=8; A_RUC=11

# ── Reglas de cruce ──────────────────────────────────────────────────────────
# SIRE: 01/07/08/30/42 con serie que empieza con LETRA → Tipo+Serie+N°+RUC
# SIRE: 50/54 con serie que empieza con NÚMERO         → Tipo+Serie+N°
# SAP:  ignorar tipos 00, 05, 14, 46 siempre
# SAP:  ignorar cualquier doc con serie que empieza con 00
TIPOS_SIRE_CON_RUC    = {'01','07','08','30','42'}
TIPOS_SIRE_SIN_RUC    = {'50','53','54'}
TIPOS_VALIDOS         = TIPOS_SIRE_CON_RUC | TIPOS_SIRE_SIN_RUC
TIPOS_SAP_IGNORAR     = {'00','0','05','5','14','46'}

def fmt_s(v):
    try: return f"{float(v):,.2f}" if v else '0.00'
    except: return '0.00'

def flt(v):
    try: return float(str(v).replace(',','') or 0)
    except: return 0.0

def norm_nro(n):
    """Elimina ceros a la izquierda del número de comprobante."""
    return str(n or '').strip().lstrip('0') or '0'

def serie_valida_sire(tipo, serie):
    """SIRE: 01/07/08/30/42 deben tener serie con letra. 50/54 con número."""
    if not serie: return False
    tipo = tipo.zfill(2)
    if tipo in TIPOS_SIRE_CON_RUC:
        return serie[0].isalpha()     # E001, F016, etc.
    if tipo in TIPOS_SIRE_SIN_RUC:
        return serie[0].isdigit()     # 0001, etc.
    return False

def serie_valida_sap(tipo, serie):
    """SAP: ignorar tipos 00/05/14/46, y series que empiezan con 00."""
    tipo  = str(tipo or '').strip().zfill(2)
    serie = str(serie or '').strip()
    if tipo in TIPOS_SAP_IGNORAR:
        return False
    if serie.startswith('00'):
        return False
    return True

def tipo_label(t):
    m = {
        '01': 'Factura',
        '07': 'Nota de Crédito',
        '08': 'Nota de Débito',
        '30': 'Tarjeta de Crédito/Débito',
        '42': 'Tarjeta Propia',
        '50': 'Dec. Única de Aduanas',
        '53': 'Mensajería / Courier',
        '54': 'Liquidación de Cobranza',
    }
    return m.get(t.zfill(2), t)

TIPO_COLOR = {'01':'#1E40AF','07':'#991B1B','08':'#92400E','30':'#166534','42':'#5B21B6','50':'#0E7490','53':'#6D28D9','54':'#065F46'}
TIPO_BG    = {'01':'#EFF6FF','07':'#FEF2F2','08':'#FFFBEB','30':'#F0FDF4','42':'#F5F3FF','50':'#ECFEFF','53':'#F5F3FF','54':'#F0FDF4'}

# ── Detectar período directamente del archivo SIRE ───────────────────────────
# Lee el período del primer registro válido del TXT (columna 2)
meses_es = ['','ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
             'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']

periodo_id = None
with open(SIRE_FILE, 'r', encoding='utf-8', errors='ignore') as _f:
    for _i, _line in enumerate(_f):
        if _i == 0: continue
        _cols = _line.rstrip('\n').split('|')
        if len(_cols) < 3: continue
        _per = _cols[2].strip()
        if len(_per) == 6 and _per.isdigit():
            periodo_id = _per
            break

if not periodo_id:
    print("❌ No se pudo detectar el período del archivo SIRE"); sys.exit(1)

_anio = int(periodo_id[:4])
_mes  = int(periodo_id[4:])
periodo = f"{meses_es[_mes]} {_anio}"
hoy = datetime.now()
print(f"   Período detectado del SIRE: {periodo_id} → {periodo}")

# ── Leer SIRE TXT ─────────────────────────────────────────────────────────────
sire_rows = []
with open(SIRE_FILE, 'r', encoding='utf-8', errors='ignore') as f:
    for i, line in enumerate(f):
        if i == 0: continue          # cabecera
        cols = line.rstrip('\n').split('|')
        if len(cols) < 25: continue
        tipo = cols[T_TIPO].strip()
        if tipo not in TIPOS_VALIDOS: continue
        per  = cols[T_PERIODO].strip()
        if per != periodo_id: continue   # solo mes actual
        sire_rows.append(cols)

print(f"   SIRE período {periodo_id}: {len(sire_rows)} registros válidos")

# ── Leer SAP Excel ────────────────────────────────────────────────────────────
wb      = openpyxl.load_workbook(SAP_FILE, read_only=True)
ws      = wb.active if 'SAP' not in wb.sheetnames else wb['SAP']
sap_rows = [list(r) for i,r in enumerate(ws.iter_rows(values_only=True))
            if i > 0 and any(v is not None for v in r)]

print(f"   SAP: {len(sap_rows)} registros")

# ── Cruce TIPO + SERIE + NRO + RUC ───────────────────────────────────────────
def key_sap(r):
    tipo  = str(r[A_TIPO] or '').strip().zfill(2)
    serie = str(r[A_SERIE] or '').strip()
    nro   = norm_nro(r[A_NRO])
    ruc   = str(r[A_RUC] or '').strip()
    if not serie_valida_sap(tipo, serie):
        return None   # ignorar este registro del SAP
    if tipo in TIPOS_SIRE_SIN_RUC:
        return f"{tipo}|{serie}|{nro}"
    return f"{tipo}|{serie}|{nro}|{ruc}"

def key_sire(r):
    tipo  = r[T_TIPO].strip().zfill(2)
    serie = r[T_SERIE].strip()
    nro   = norm_nro(r[T_NRO])
    ruc   = r[T_RUC].strip()
    if tipo in TIPOS_SIRE_SIN_RUC:
        return f"{tipo}|{serie}|{nro}"
    return f"{tipo}|{serie}|{nro}|{ruc}"

sap_keys    = set(k for r in sap_rows if (k := key_sap(r)) is not None)
pendientes  = [r for r in sire_rows if key_sire(r) not in sap_keys]
registrados = [r for r in sire_rows if key_sire(r) in sap_keys]

total_sire = len(sire_rows)
ya_reg     = len(registrados)
pct_reg    = round(ya_reg / total_sire * 100) if total_sire else 0
pct_pend   = 100 - pct_reg

total_bi  = sum(flt(r[T_BI])    for r in pendientes)
total_igv = sum(flt(r[T_IGV])   for r in pendientes)
total_cp  = sum(flt(r[T_TOTAL]) for r in pendientes)

print(f"   ✅ Registrados: {ya_reg} ({pct_reg}%)")
print(f"   ❌ Pendientes:  {len(pendientes)} ({pct_pend}%)")

# Fechas del período
fechas_txt = [r[T_FECHA].strip() for r in sire_rows if r[T_FECHA].strip()]
fecha_min  = min(fechas_txt) if fechas_txt else '-'
fecha_max  = max(fechas_txt) if fechas_txt else '-'

# ── Agrupaciones ─────────────────────────────────────────────────────────────
by_tipo = defaultdict(list)
for r in pendientes:
    by_tipo[r[T_TIPO].strip().zfill(2)].append(r)

by_prov = defaultdict(lambda: {'count':0,'total':0,'igv':0,'tipos':set()})
for r in pendientes:
    razon = r[T_RAZON].strip()
    for s in [' SOCIEDAD ANONIMA CERRADA',' SOCIEDAD COMERCIAL DE RESPONSABILIDAD LIMITADA',
              ' EMPRESA INDIVIDUAL DE RESPONSABILIDAD LIMITADA',
              ' S.A.C.',' SAC',' S.A.',' SA',' E.I.R.L.',' S.R.L.',' SRL']:
        razon = razon.replace(s,'')
    if len(razon) > 42: razon = razon[:42]+'…'
    by_prov[razon]['count']  += 1
    by_prov[razon]['total']  += flt(r[T_TOTAL])
    by_prov[razon]['igv']    += flt(r[T_IGV])
    by_prov[razon]['tipos'].add(r[T_TIPO].strip().zfill(2))

top10    = sorted(by_prov.items(), key=lambda x: x[1]['total'], reverse=True)[:10]
num_prov = len(by_prov)

# ── JS data ───────────────────────────────────────────────────────────────────
# ── Ordenar pendientes alfabéticamente por razón social ───────────────────────
pendientes.sort(key=lambda r: str(r[T_RAZON] or '').strip().upper())

js_data = [{'fecha':r[T_FECHA].strip(),'tipo':r[T_TIPO].strip().zfill(2),
    'serie':r[T_SERIE].strip(),'nro':r[T_NRO].strip(),'ruc':r[T_RUC].strip(),
    'proveedor':r[T_RAZON].strip(),'bi':flt(r[T_BI]),'igv':flt(r[T_IGV]),
    'total':flt(r[T_TOTAL]),'moneda':r[T_MONEDA].strip(),
    'orig_fecha':r[T_ORIG_FECHA].strip() if len(r)>T_ORIG_FECHA else '',
    'orig_tipo': r[T_ORIG_TIPO].strip().zfill(2) if len(r)>T_ORIG_TIPO and r[T_ORIG_TIPO].strip() else '',
    'orig_serie':r[T_ORIG_SERIE].strip() if len(r)>T_ORIG_SERIE else '',
    'orig_nro':  r[T_ORIG_NRO].strip() if len(r)>T_ORIG_NRO else '',
    } for r in pendientes]

unique_dates   = sorted(set(d['fecha']  for d in js_data))
unique_monedas = sorted(set(d['moneda'] for d in js_data))
unique_tipos   = sorted(set(d['tipo']   for d in js_data))
razones        = sorted(set(d['proveedor'] for d in js_data if d['proveedor']))

date_opts  = '\n'.join(f'<option value="{d}">{d}</option>'             for d in unique_dates)
mon_opts   = '\n'.join(f'<option value="{m}">{m}</option>'             for m in unique_monedas)
tipo_opts  = '\n'.join(f'<option value="{t}">{tipo_label(t)}</option>' for t in unique_tipos)
razon_opts = '\n'.join(f'<option value="{r}"></option>'                for r in razones)

# Tarjetas por tipo
tipo_cards_html = ''
for t in ['01','07','08','30','42','50','53','54']:
    rows_t = by_tipo.get(t, [])
    if not rows_t: continue
    tot = sum(flt(r[T_TOTAL]) for r in rows_t)
    col = TIPO_COLOR.get(t,'#374151')
    bg  = TIPO_BG.get(t,'#F2F4F7')
    tipo_cards_html += f"""<div class="tipo-card" style="border-left-color:{col}">
      <div class="tipo-card-lbl">{tipo_label(t)}</div>
      <div class="tipo-card-num" style="color:{col}">{len(rows_t)}</div>
      <div class="tipo-card-total">{tot:,.2f}</div>
    </div>"""

# Top 10 HTML
def tipo_badge(t):
    col = TIPO_COLOR.get(t,'#64748b'); bg = TIPO_BG.get(t,'#f5f7fb')
    return f'<span style="background:{bg};color:{col};border:1px solid {col}33;border-radius:20px;padding:1px 8px;font-size:9.5px;font-weight:700">{tipo_label(t)}</span>'

prov_html = ''.join(f"""<tr>
  <td class="proveedor-cell">{n}</td>
  <td style="text-align:left;padding-left:13px">{''.join(tipo_badge(t) for t in sorted(d['tipos']))}</td>
  <td class="num">{d['count']}</td>
  <td class="num money">S/ {d['total']:,.2f}</td>
  <td class="num igv">S/ {d['igv']:,.2f}</td>
</tr>""" for n,d in top10)

now_str = datetime.now().strftime('%d/%m/%Y %H:%M')

HTML = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Comprobantes Pendientes SAP — {periodo}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

:root {{
  /* Paleta oficina: blanco roto, grafito, azul petróleo, rojo señal */
  --bg:        #F8FAFC;
  --surface:   #FFFFFF;
  --surface2:  #F1F5F9;
  --border:    #E2E8F0;
  --border2:   #CBD5E1;

  --ink:       #0F172A;
  --ink2:      #334155;
  --ink3:      #64748B;

  --navy:      #1E40AF;
  --navy2:     #1D4ED8;
  --navy3:     #2563EB;

  --red:       #DC2626;
  --red-bg:    #FEF2F2;
  --red-bdr:   #FECACA;

  --green:     #16A34A;
  --green-bg:  #F0FDF4;
  --green-bdr: #BBF7D0;

  --amber:     #D97706;
  --amber-bg:  #FFFBEB;

  --mono: 'JetBrains Mono', monospace;
  --sans: 'Inter', sans-serif;
  --radius: 6px;
  --shadow-sm: 0 1px 3px rgba(0,0,0,.06), 0 1px 2px rgba(0,0,0,.04);
  --shadow:    0 2px 6px rgba(0,0,0,.06), 0 1px 3px rgba(0,0,0,.04);
}}

* {{ margin:0; padding:0; box-sizing:border-box }}
body {{ font-family:var(--sans); background:var(--bg); color:var(--ink); font-size:13px; line-height:1.5 }}

/* ═══ HEADER ════════════════════════════════════════════════════════ */
.hdr {{
  background: var(--navy);
  border-bottom: 3px solid var(--red);
  padding: 0 24px;
  display: flex;
  align-items: stretch;
  min-height: 90px;
}}
.hdr-brand {{
  display: flex;
  flex-direction: column;
  justify-content: center;
  padding: 20px 32px 20px 0;
  border-right: 1px solid rgba(255,255,255,.15);
  margin-right: 32px;
  min-width: 220px;
}}
.hdr-eyebrow {{
  font-size: 9px;
  font-weight: 700;
  letter-spacing: 2px;
  text-transform: uppercase;
  color: rgba(255,255,255,.65);
  margin-bottom: 6px;
}}
.hdr-title {{
  font-size: 18px;
  font-weight: 700;
  color: #FFFFFF;
  letter-spacing: -.3px;
  line-height: 1.2;
}}


/* KPIs en header */
.hdr-kpis {{
  display: flex;
  align-items: stretch;
  flex: 1;
}}
.hdr-kpi {{
  display: flex;
  flex-direction: column;
  justify-content: center;
  padding: 0 28px;
  border-right: 1px solid rgba(255,255,255,.15);
}}
.hdr-kpi:last-child {{ border-right: none }}
.hdr-kpi-lbl {{
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 1px;
  text-transform: uppercase;
  color: #A8C4E0;
  margin-bottom: 5px;
}}
.hdr-kpi-val {{
  font-family: var(--mono);
  font-size: 26px;
  font-weight: 600;
  color: #FFFFFF;
  line-height: 1;
}}
.hdr-kpi-val.danger  {{ color: #FFA09A }}
.hdr-kpi-val.success {{ color: #7EEAAA }}
.hdr-kpi-sub {{
  font-size: 11px;
  color: rgba(255,255,255,.6);
  margin-top: 5px;
  font-weight: 400;
}}

.hdr-periodo {{
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: flex-end;
  padding: 20px 0 20px 28px;
  border-left: 1px solid rgba(255,255,255,.15);
  margin-left: auto;
}}
.hdr-mes {{
  font-family: var(--mono);
  font-size: 17px;
  font-weight: 600;
  color: #FFFFFF;
  letter-spacing: 1px;
}}
.hdr-fechas {{
  font-size: 11px;
  color: rgba(255,255,255,.65);
  margin-top: 5px;
  text-align: right;
}}

/* ═══ BARRA PROGRESO ════════════════════════════════════════════════ */
.prog-wrap {{ padding: 16px 36px 0 }}
.prog-card {{
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 12px 20px;
  box-shadow: var(--shadow-sm);
  display: flex;
  align-items: center;
  gap: 20px;
}}
.prog-lbl {{
  font-size: 10px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 1px;
  color: var(--ink3);
  white-space: nowrap;
  min-width: 110px;
}}
.prog-track {{
  flex: 1;
  height: 8px;
  background: var(--border);
  border-radius: 4px;
  overflow: hidden;
  display: flex;
}}
.prog-reg  {{ height:100%; background: #16A34A; width:{pct_reg}% }}
.prog-pend {{ height:100%; background: #DC2626; width:{pct_pend}% }}
.prog-legend {{
  display: flex;
  gap: 20px;
  font-size: 11px;
  white-space: nowrap;
}}
.prog-legend-item {{
  display: flex;
  align-items: center;
  gap: 6px;
  color: var(--ink2);
}}
.dot {{
  width: 8px; height: 8px;
  border-radius: 50%;
  display: inline-block;
  flex-shrink: 0;
}}

/* ═══ TARJETAS POR TIPO ═════════════════════════════════════════════ */
.tipo-cards {{
  display: flex;
  gap: 10px;
  padding: 10px 24px 0;
  flex-wrap: wrap;
}}
.tipo-card {{
  background: var(--surface);
  border: 1px solid var(--border);
  border-left: 4px solid #ccc;
  border-radius: var(--radius);
  padding: 11px 14px;
  box-shadow: var(--shadow-sm);
  min-width: 100px;
  flex: 1;
}}
.tipo-card-lbl {{
  font-size: 9px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 1px;
  color: var(--ink3);
  margin-bottom: 6px;
}}
.tipo-card-num {{
  font-family: var(--mono);
  font-size: 22px;
  font-weight: 500;
  line-height: 1;
  color: var(--ink);
}}
.tipo-card-total {{
  font-family: var(--mono);
  font-size: 11px;
  color: var(--ink3);
  margin-top: 4px;
}}



/* ═══ SECCIONES ══════════════════════════════════════════════════════ */
.sec {{ padding: 18px 36px 0 }}
.sec-hdr {{
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 10px;
  padding-bottom: 8px;
  border-bottom: 1px solid var(--border);
}}
.sec-title {{
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 1px;
  color: var(--navy);
}}
.sec-line {{ flex: 1 }}
.badge {{
  background: var(--ink2);
  color: #fff;
  border-radius: 3px;
  padding: 3px 10px;
  font-size: 11px;
  font-weight: 600;
  font-family: var(--mono);
  letter-spacing: .3px;
}}
.badge.red {{ background: var(--red) }}

/* ═══ FILTROS ════════════════════════════════════════════════════════ */
.fbar {{
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 14px 16px;
  margin-bottom: 8px;
  box-shadow: var(--shadow-sm);
}}
.fbar-row {{
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-items: flex-end;
}}
.fbar-row + .fbar-row {{
  margin-top: 10px;
  padding-top: 10px;
  border-top: 1px solid var(--border);
}}
.fg {{ display: flex; flex-direction: column; gap: 4px }}
.fg label {{
  font-size: 11px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 1px;
  color: var(--ink3);
}}
.fg input, .fg select {{
  border: 1.5px solid var(--border2);
  border-radius: var(--radius);
  padding: 8px 12px;
  font-size: 13px;
  font-family: var(--sans);
  color: var(--ink);
  background: var(--surface2);
  outline: none;
  transition: border-color .15s, box-shadow .15s;
  min-width: 140px;
}}
.fg input:focus, .fg select:focus {{
  border-color: var(--navy3);
  box-shadow: 0 0 0 3px rgba(37,99,235,.12);
  background: #fff;
}}
.fg-wide input {{ min-width: 210px }}
.fg-num input  {{ min-width: 100px }}
.razon-wrap {{ display: flex; gap: 6px; align-items: flex-end }}
.razon-wrap input {{ flex: 1; min-width: 200px }}

.btn-buscar {{
  padding: 7px 16px;
  border-radius: var(--radius);
  font-size: 12px;
  font-family: var(--sans);
  font-weight: 700;
  cursor: pointer;
  border: 1.5px solid var(--navy3);
  background: var(--navy3);
  color: #fff;
  display: flex;
  align-items: center;
  gap: 5px;
  transition: all .15s;
  letter-spacing: .3px;
}}
.btn-buscar:hover {{ background: var(--navy2); border-color: var(--navy2) }}

.fbar-actions {{ display: flex; gap: 8px; align-items: flex-end; margin-left: auto }}
.btn {{
  padding: 7px 16px;
  border-radius: var(--radius);
  font-size: 13px;
  font-family: var(--sans);
  font-weight: 600;
  cursor: pointer;
  border: none;
  transition: all .15s;
  display: flex;
  align-items: center;
  gap: 5px;
}}
.btn-clear {{
  background: var(--surface2);
  color: var(--ink3);
  border: 1.5px solid var(--border2);
}}
.btn-clear:hover {{ background: var(--border); color: var(--ink) }}
.btn-excel {{
  background: #1A5C35;
  color: #fff;
  font-weight: 700;
  box-shadow: 0 2px 6px rgba(0,0,0,.3);
  border: 1.5px solid #236B40;
}}
.btn-excel:hover {{ background: #145029 }}

/* CHIPS */
.chips {{ display:flex; flex-wrap:wrap; gap:5px; margin-top:8px; min-height:20px }}
.chip {{
  background: #EFF6FF;
  border: 1px solid #BFDBFE;
  border-radius: 3px;
  padding: 3px 9px;
  font-size: 10px;
  font-weight: 600;
  color: #1E40AF;
  display: flex;
  align-items: center;
  gap: 4px;
}}
.chip-x {{
  cursor: pointer;
  color: #60A5FA;
  font-size: 11px;
  line-height: 1;
  transition: color .1s;
}}
.chip-x:hover {{ color: var(--red) }}

/* RESULTADO */
.res-bar {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 6px;
  padding: 0 2px;
}}
.res-bar {{ background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 10px 16px; margin-bottom: 8px; box-shadow: var(--shadow-sm); }}
.res-count {{ font-size: 12px; color: var(--ink3) }}
.res-count strong {{ color: var(--ink); font-size: 15px; font-weight: 700 }}
.res-totals {{ display: flex; gap: 28px; font-family: var(--mono); font-size: 13px }}
.res-item {{ display: flex; flex-direction: column; align-items: flex-end }}
.res-lbl {{ font-size: 9px; text-transform: uppercase; letter-spacing: 1px; color: var(--ink3); margin-bottom: 2px; font-weight: 600 }}
.res-val {{ color: var(--ink2); font-weight: 600; font-size: 15px }}
.res-val.igv-v {{ color: var(--navy2) }}
.res-val.tot-v {{ color: var(--red); font-weight: 700; font-size: 17px }}

/* ═══ TABLA ══════════════════════════════════════════════════════════ */
.tbl-wrap {{
  overflow-x: auto;
  border-radius: var(--radius);
  box-shadow: var(--shadow-sm);
  margin-bottom: 24px;
  background: var(--surface);
  border: 1px solid var(--border);
}}
table {{ border-collapse: collapse; width: 100%; min-width: 860px }}
thead th {{
  background: var(--ink);
  color: rgba(255,255,255,.9);
  font-size: 10px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 1px;
  padding: 11px 13px;
  text-align: left;
  white-space: nowrap;
  border-right: 1px solid rgba(255,255,255,.1);
}}
thead th:last-child {{ border-right: none }}
thead th.r {{ text-align: right }}
thead th.igv-h  {{ background: #1D4ED8 }}
thead th.tot-h  {{ background: #991B1B }}

tbody tr {{ border-bottom: 1px solid var(--border); transition: background .1s }}
tbody tr:nth-child(even) {{ background: var(--surface2) }}
tbody tr:hover {{ background: #EEF2FA }}
tbody td {{ padding: 9px 13px; vertical-align: middle }}

.mono  {{ font-family: var(--mono); font-size: 12px }}
.muted {{ color: var(--ink2) }}
.num   {{ text-align: right; font-family: var(--mono); font-size: 13px }}
.igv   {{ background: #F0F5FF !important }}
.money {{ font-weight: 600; color: var(--red) }}
.total-cp {{ background: var(--red-bg) !important; font-weight: 600; color: var(--red) }}
.proveedor-cell {{
  max-width: 240px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  font-size: 13px;
}}
.col-fecha  {{ width: 88px  }}
.col-tipo   {{ width: 120px }}
.col-serie  {{ width: 52px  }}
.col-ncp    {{ width: 68px  }}
.col-ruc    {{ width: 100px }}
.col-importe{{ width: 14%   }}
.col-total  {{ width: 14%   }}
.col-mon    {{ width: 44px  }}
.empty-td   {{ text-align:center; padding:36px!important; color:var(--ink3); font-size:12px }}

/* Tipo badges */
.tipo-badge {{
  display: inline-block;
  padding: 1px 7px;
  border-radius: 3px;
  font-size: 9.5px;
  font-weight: 600;
  letter-spacing: .3px;
  white-space: nowrap;
  border: 1px solid transparent;
}}

/* ═══ FOOTER ═════════════════════════════════════════════════════════ */
.footer {{
  padding: 12px 24px 20px;
  font-size: 10px;
  color: var(--ink3);
  border-top: 1px solid var(--border);
  margin-top: 8px;
  display: flex;
  justify-content: space-between;
  align-items: center;
}}
.footer-brand {{ font-weight: 600; color: var(--ink2); letter-spacing: .5px }}

@media print {{
  .fbar, .fbar-actions, .chips {{ display: none !important }}
  body {{ background: #fff }}
  .hdr, thead th {{ -webkit-print-color-adjust: exact; print-color-adjust: exact }}
}}
</style>
</head>
<body>

<!-- HEADER -->
<div class="hdr">
  <div class="hdr-brand">
    <div class="hdr-eyebrow">Casas y Colores · Contabilidad</div>
    <div class="hdr-title">Comprobantes Pendientes</div>
  </div>
  <div class="hdr-kpis">
    <div class="hdr-kpi">
      <div class="hdr-kpi-lbl">Total SIRE</div>
      <div class="hdr-kpi-val">{total_sire}</div>
      <div class="hdr-kpi-sub">Tipos 01·07·08·30·42·50·53·54</div>
    </div>
    <div class="hdr-kpi">
      <div class="hdr-kpi-lbl">Ya en SAP</div>
      <div class="hdr-kpi-val success">{ya_reg}</div>
      <div class="hdr-kpi-sub">{pct_reg}% registrado</div>
    </div>
    <div class="hdr-kpi">
      <div class="hdr-kpi-lbl">Pendientes</div>
      <div class="hdr-kpi-val danger">{len(pendientes)}</div>
      <div class="hdr-kpi-sub">{pct_pend}% sin ingresar</div>
    </div>
    <div class="hdr-kpi">
      <div class="hdr-kpi-lbl">Total Pendiente</div>
      <div class="hdr-kpi-val danger">S/ {total_cp:,.0f}</div>
      <div class="hdr-kpi-sub">Por ingresar al SAP</div>
    </div>
    <div class="hdr-kpi">
      <div class="hdr-kpi-lbl">Proveedores</div>
      <div class="hdr-kpi-val">{num_prov}</div>
      <div class="hdr-kpi-sub">Distintos con pendientes</div>
    </div>
  </div>
  <div class="hdr-periodo">
    <div class="hdr-mes">{periodo}</div>
    <div class="hdr-fechas">{fecha_min} — {fecha_max}</div>
    <div class="hdr-fechas" style="margin-top:4px">Actualizado: {now_str}</div>
  </div>
</div>

<!-- BARRA PROGRESO -->
<div class="prog-wrap">
  <div class="prog-card">
    <span class="prog-lbl">Avance de registro</span>
    <div class="prog-track">
      <div class="prog-reg"></div>
      <div class="prog-pend"></div>
    </div>
    <div class="prog-legend">
      <div class="prog-legend-item">
        <span class="dot" style="background:var(--green)"></span>
        Registrados en SAP: <strong>{ya_reg}</strong> ({pct_reg}%)
      </div>
      <div class="prog-legend-item">
        <span class="dot" style="background:var(--red)"></span>
        Pendientes: <strong>{len(pendientes)}</strong> ({pct_pend}%)
      </div>
    </div>
  </div>
</div>

<!-- TARJETAS POR TIPO -->
<div class="tipo-cards">{tipo_cards_html}</div>



<!-- TOP PROVEEDORES -->
<div class="sec">
  <div class="sec-hdr">
    <span class="sec-title">Top 10 Proveedores con más pendientes</span>
    <span class="sec-line"></span>
    <span class="badge red">{num_prov} proveedores</span>
  </div>
  <div class="tbl-wrap">
    <table style="min-width:0;table-layout:fixed;width:100%">
      <thead><tr>
        <th style="width:36%">Proveedor / Razón Social</th>
        <th style="width:18%">Tipos</th>
        <th class="r" style="width:6%">N° CP</th>
        <th class="r tot-h" style="width:20%">Total Pendiente</th>
        <th class="r igv-h" style="width:20%">IGV</th>
      </tr></thead>
      <tbody>{prov_html}</tbody>
    </table>
  </div>
</div>

<!-- DETALLE -->
<div class="sec">
  <div class="sec-hdr">
    <span class="sec-title">Detalle de Comprobantes Pendientes</span>
    <span class="sec-line"></span>
    <span class="badge red" id="badge-count">{len(pendientes)} registros</span>
  </div>

  <div class="fbar">
    <div class="fbar-row">
      <div class="fg fg-wide">
        <label>Razón Social</label>
        <div class="razon-wrap">
          <input type="text" id="f-razon" list="razon-list"
                 placeholder="Escribe o selecciona proveedor…"
                 onkeydown="if(event.key==='Enter')applyFilters()">
          <datalist id="razon-list">{razon_opts}</datalist>
          <button class="btn-buscar" onclick="applyFilters()">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            Buscar
          </button>
        </div>
      </div>
      <div class="fg fg-wide">
        <label>RUC / Serie / Número</label>
        <input type="text" id="f-texto" placeholder="Ej: 20610024981 · E001 · 1080…" oninput="applyFilters()">
      </div>
      <div class="fbar-actions">
        <button class="btn btn-clear" onclick="clearFilters()">✕ Limpiar</button>
        <button class="btn btn-excel" onclick="exportExcel()">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg>
          Exportar Excel
        </button>
      </div>
    </div>
    <div class="fbar-row">
      <div class="fg">
        <label>Tipo CP</label>
        <select id="f-tipo" onchange="applyFilters()">
          <option value="">Todos los tipos</option>
          {tipo_opts}
        </select>
      </div>
      <div class="fg">
        <label>Fecha Emisión</label>
        <select id="f-fecha" onchange="applyFilters()">
          <option value="">Todas las fechas</option>
          {date_opts}
        </select>
      </div>
      <div class="fg">
        <label>Moneda</label>
        <select id="f-moneda" onchange="applyFilters()">
          <option value="">Todas</option>
          {mon_opts}
        </select>
      </div>
      <div class="fg fg-num">
        <label>Total mínimo</label>
        <input type="number" id="f-min" placeholder="0" step="100" oninput="applyFilters()">
      </div>
      <div class="fg fg-num">
        <label>Total máximo</label>
        <input type="number" id="f-max" placeholder="Sin límite" min="0" step="100" oninput="applyFilters()">
      </div>
    </div>
    <div class="chips" id="chips"></div>
  </div>

  <div class="res-bar">
    <div class="res-count">Mostrando <strong id="rc-n">{len(pendientes)}</strong> de {len(pendientes)} comprobantes pendientes</div>
    <div class="res-totals">
      <div class="res-item"><span class="res-lbl">Base Imponible</span><span class="res-val" id="s-bi">S/ {total_bi:,.2f}</span></div>
      <div class="res-item"><span class="res-lbl">IGV</span><span class="res-val igv-v" id="s-igv">S/ {total_igv:,.2f}</span></div>
      <div class="res-item"><span class="res-lbl">Total Pendiente</span><span class="res-val tot-v" id="s-tot">S/ {total_cp:,.2f}</span></div>
    </div>
  </div>

  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th class="col-fecha">Fecha</th>
        <th class="col-tipo">Tipo CP</th>
        <th class="col-serie">Serie</th>
        <th class="col-ncp">N° CP</th>
        <th class="col-ruc">RUC</th>
        <th>Proveedor / Razón Social</th>
        <th class="r">Base Imponible</th>
        <th class="r igv-h">IGV / IPM</th>
        <th class="r tot-h col-total">Total CP</th>
        <th class="col-mon">Mon.</th>
      </tr></thead>
      <tbody id="tbody-fact"></tbody>
    </table>
  </div>
</div>

<div class="footer">
  <span class="footer-brand">SIRE ↔ SAP · Casas y Colores · Contabilidad</span>
  <span>{periodo} · {len(pendientes)} pendientes de {total_sire} · Tipos 01·07·08·30·42·50·53·54 · Uso interno</span>
  <span>Actualizado: {now_str}</span>
</div>

<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
<script>
const DATA={json.dumps(js_data,ensure_ascii=False)};
const TOTAL=DATA.length;
const TIPO_LABEL={json.dumps({t:tipo_label(t) for t in ['01','07','08','30','42','50','53','54']},ensure_ascii=False)};
const TIPO_COLOR={json.dumps(TIPO_COLOR,ensure_ascii=False)};
const TIPO_BG={json.dumps(TIPO_BG,ensure_ascii=False)};
const fmt=v=>v.toLocaleString('es-PE',{{minimumFractionDigits:2,maximumFractionDigits:2}});

function tipoBadge(t){{
  const col=TIPO_COLOR[t]||'#374151';
  const bg=TIPO_BG[t]||'#F2F4F7';
  const lbl=TIPO_LABEL[t]||t;
  return`<span class="tipo-badge" style="background:${{bg}};color:${{col}};border-color:${{col}}44">${{lbl}}</span>`;
}}

function renderRows(list){{
  const tb=document.getElementById('tbody-fact');
  if(!list.length){{
    tb.innerHTML='<tr><td colspan="10" class="empty-td">Sin comprobantes pendientes con los filtros aplicados.</td></tr>';
    return;
  }}
  tb.innerHTML=list.map(d=>{{
    const prov=d.proveedor.length>55?d.proveedor.slice(0,55)+'…':d.proveedor;
    return`<tr>
      <td class="mono col-fecha">${{d.fecha}}</td>
      <td class="col-tipo">${{tipoBadge(d.tipo)}}</td>
      <td class="mono col-serie">${{d.serie}}</td>
      <td class="mono col-ncp">${{d.nro}}</td>
      <td class="mono col-ruc">${{d.ruc}}</td>
      <td class="proveedor-cell" title="${{d.proveedor}}">${{prov}}</td>
      <td class="num col-importe">${{fmt(d.bi)}}</td>
      <td class="num igv col-importe">${{fmt(d.igv)}}</td>
      <td class="num total-cp col-importe" style="${{d.total<0?'color:#7B2020;background:#FDF3F2':''}}">
        ${{d.total<0?'- ':''}}${{fmt(Math.abs(d.total))}}</td>
      <td class="mono col-mon" style="font-size:12px;font-weight:600">${{d.moneda}}</td>
    </tr>`;
  }}).join('');
}}

function updateSummary(list){{
  document.getElementById('rc-n').textContent=list.length;
  document.getElementById('badge-count').textContent=list.length+' registros';
  document.getElementById('s-bi').textContent=fmt(list.reduce((a,d)=>a+d.bi,0));
  document.getElementById('s-igv').textContent=fmt(list.reduce((a,d)=>a+d.igv,0));
  document.getElementById('s-tot').textContent=fmt(list.reduce((a,d)=>a+d.total,0));
}}

function getActive(){{
  return{{
    razon:  document.getElementById('f-razon').value.trim(),
    texto:  document.getElementById('f-texto').value.trim().toLowerCase(),
    tipo:   document.getElementById('f-tipo').value,
    fecha:  document.getElementById('f-fecha').value,
    moneda: document.getElementById('f-moneda').value,
    min:    document.getElementById('f-min').value !== '' ? parseFloat(document.getElementById('f-min').value) : -Infinity,
    max:    parseFloat(document.getElementById('f-max').value)||Infinity,
  }};
}}

function getFiltered(){{
  const f=getActive();
  return DATA.filter(d=>{{
    if(f.razon  && !d.proveedor.toLowerCase().includes(f.razon.toLowerCase())) return false;
    if(f.texto  && !d.ruc.includes(f.texto) && !d.proveedor.toLowerCase().includes(f.texto)
                && !d.serie.toLowerCase().includes(f.texto) && !d.nro.includes(f.texto)) return false;
    if(f.tipo   && d.tipo   !== f.tipo)   return false;
    if(f.fecha  && d.fecha  !== f.fecha)  return false;
    if(f.moneda && d.moneda !== f.moneda) return false;
    const absTotal = Math.abs(d.total);
    if(f.min !== -Infinity && absTotal < f.min) return false;
    if(f.max !== Infinity  && absTotal > f.max) return false;
    return true;
  }});
}}

function renderChips(){{
  const f=getActive(); const chips=[];
  const clr=id=>()=>{{document.getElementById(id).value='';applyFilters();}};
  if(f.razon)  chips.push(['Razón: '+f.razon.slice(0,25), clr('f-razon')]);
  if(f.texto)  chips.push(['Texto: '+f.texto.slice(0,22), clr('f-texto')]);
  if(f.tipo)   chips.push(['Tipo: '+(TIPO_LABEL[f.tipo]||f.tipo), clr('f-tipo')]);
  if(f.fecha)  chips.push(['Fecha: '+f.fecha, clr('f-fecha')]);
  if(f.moneda) chips.push(['Moneda: '+f.moneda, clr('f-moneda')]);
  if(f.min>0)  chips.push(['Mín: S/'+f.min, clr('f-min')]);
  if(f.max<Infinity) chips.push(['Máx: S/'+f.max, clr('f-max')]);
  const el=document.getElementById('chips');
  el.innerHTML=chips.map(([l],i)=>`<div class="chip">${{l}}<span class="chip-x" data-i="${{i}}">✕</span></div>`).join('');
  el.querySelectorAll('.chip-x').forEach(x=>{{x.onclick=chips[+x.dataset.i][1];}});
}}

function applyFilters(){{
  const l=getFiltered(); renderRows(l); updateSummary(l); renderChips();
}}

function clearFilters(){{
  ['f-razon','f-texto'].forEach(id=>document.getElementById(id).value='');
  ['f-tipo','f-fecha','f-moneda'].forEach(id=>document.getElementById(id).value='');
  ['f-min','f-max'].forEach(id=>document.getElementById(id).value='');
  applyFilters();
}}

function exportExcel(){{
  const list = getFiltered();
  const WB   = XLSX.utils.book_new();
  const hoy  = new Date().toLocaleDateString('es-PE');
  const ISO  = new Date().toISOString().slice(0,10);

  // ── Helpers de estilo ─────────────────────────────────────────────
  const NAVY  = '1E3A5F'; const WHITE = 'FFFFFF'; const LGRAY = 'F2F4F7';
  const RED   = '7B2020'; const GREEN = '1A5C35'; const IGVBL = '1E4878';
  const TOTBG = '6B2020'; const BORD  = 'C8CDD8';

  const border = {{
    top:   {{style:'thin', color:{{rgb:BORD}}}},
    bottom:{{style:'thin', color:{{rgb:BORD}}}},
    left:  {{style:'thin', color:{{rgb:BORD}}}},
    right: {{style:'thin', color:{{rgb:BORD}}}},
  }};

  function cell(v, s) {{ return {{v, s, t: typeof v==='number'?'n':'s'}}; }}

  const NF = '#,##0.00';

  // ── Construir filas ───────────────────────────────────────────────
  // Fila 1: título
  const R1 = [
    cell('COMPROBANTES PENDIENTES DE INGRESO AL SAP — {periodo}', {{
      font:{{bold:true, sz:13, color:{{rgb:WHITE}}, name:'Arial'}},
      fill:{{fgColor:{{rgb:NAVY}}, patternType:'solid'}},
      alignment:{{horizontal:'left', vertical:'center'}},
    }}),
    ...Array(13).fill(cell('',{{fill:{{fgColor:{{rgb:NAVY}},patternType:'solid'}}}}))
  ];

  // Fila 2: subtítulo
  const R2 = [
    cell('Casas y Colores  ·  Cruce SIRE vs SAP  ·  Tipos: 01·07·08·30·42·50·54  ·  ' + hoy, {{
      font:{{italic:true, sz:9, color:{{rgb:WHITE}}, name:'Arial'}},
      fill:{{fgColor:{{rgb:'254878'}}, patternType:'solid'}},
      alignment:{{horizontal:'left', vertical:'center'}},
    }}),
    ...Array(13).fill(cell('',{{fill:{{fgColor:{{rgb:'254878'}},patternType:'solid'}}}}))
  ];

  // Fila 3: vacía
  const R3 = Array(14).fill(cell('',{{}}));

  // Fila 4: cabeceras
  const hdrs = ['FECHA','TIPO CP','SERIE','N° CP','RUC','PROVEEDOR / RAZÓN SOCIAL','BASE IMPONIBLE','IGV / IPM','TOTAL CP','MONEDA','ORIG. TIPO','ORIG. FECHA','ORIG. SERIE','ORIG. N°'];
  const hdrBg = [NAVY,NAVY,NAVY,NAVY,NAVY,NAVY,NAVY,IGVBL,TOTBG,NAVY,'4A2060','4A2060','4A2060','4A2060'];
  const R4 = hdrs.map((h,i) => cell(h, {{
    font:    {{bold:true, sz:9, color:{{rgb:WHITE}}, name:'Arial'}},
    fill:    {{fgColor:{{rgb:hdrBg[i]}}, patternType:'solid'}},
    alignment:{{horizontal: i>=6&&i<=8?'right':'left', vertical:'center'}},
    border,
  }}));

  // Filas de datos
  const dataRows = list.map((d, idx) => {{
    const isEven = idx % 2 === 0;
    const rowBg  = isEven ? WHITE : LGRAY;
    const numStyle = (bg) => ({{
      font:{{sz:9, name:'Arial', color:{{rgb:RED}}, bold:true}},
      fill:{{fgColor:{{rgb:bg}}, patternType:'solid'}},
      alignment:{{horizontal:'right', vertical:'center'}},
      border, numFmt: NF,
    }});
    const txtStyle = (bg, clr='111827', bold=false) => ({{
      font:{{sz:9, name:'Arial', color:{{rgb:clr}}, bold}},
      fill:{{fgColor:{{rgb:bg}}, patternType:'solid'}},
      alignment:{{horizontal:'left', vertical:'center'}},
      border,
    }});
    return [
      cell(d.fecha,     txtStyle(rowBg)),
      cell(TIPO_LABEL[d.tipo]||d.tipo, txtStyle(rowBg)),
      cell(d.serie,     txtStyle(rowBg)),
      cell(d.nro,       txtStyle(rowBg)),
      cell(d.ruc,       txtStyle(rowBg, '6B7280')),
      cell(d.proveedor, txtStyle(rowBg)),
      {{v:d.bi,   t:'n', s:{{font:{{sz:9,name:'Arial',color:{{rgb:'111827'}}}},fill:{{fgColor:{{rgb:'F0F5FF'}},patternType:'solid'}},alignment:{{horizontal:'right',vertical:'center'}},border,numFmt:NF}}}},
      {{v:d.igv,  t:'n', s:{{font:{{sz:9,name:'Arial',color:{{rgb:'111827'}}}},fill:{{fgColor:{{rgb:'EEF2FA'}},patternType:'solid'}},alignment:{{horizontal:'right',vertical:'center'}},border,numFmt:NF}}}},
      {{v:d.total,t:'n', s:numStyle('FDF3F2')}},
      cell(d.moneda, txtStyle(rowBg)),
      cell(d.orig_tipo  ||'', {{font:{{sz:9,name:'Arial',color:{{rgb:d.orig_tipo?'4A2060':'9CA3AF'}}}},fill:{{fgColor:{{rgb:d.orig_tipo?'F3EEFF':rowBg}},patternType:'solid'}},alignment:{{horizontal:'left',vertical:'center'}},border}}),
      cell(d.orig_fecha ||'', {{font:{{sz:9,name:'Arial',color:{{rgb:d.orig_tipo?'4A2060':'9CA3AF'}}}},fill:{{fgColor:{{rgb:d.orig_tipo?'F3EEFF':rowBg}},patternType:'solid'}},alignment:{{horizontal:'left',vertical:'center'}},border}}),
      cell(d.orig_serie ||'', {{font:{{sz:9,name:'Arial',color:{{rgb:d.orig_tipo?'4A2060':'9CA3AF'}}}},fill:{{fgColor:{{rgb:d.orig_tipo?'F3EEFF':rowBg}},patternType:'solid'}},alignment:{{horizontal:'left',vertical:'center'}},border}}),
      cell(d.orig_nro   ||'', {{font:{{sz:9,name:'Arial',color:{{rgb:d.orig_tipo?'4A2060':'9CA3AF'}}}},fill:{{fgColor:{{rgb:d.orig_tipo?'F3EEFF':rowBg}},patternType:'solid'}},alignment:{{horizontal:'left',vertical:'center'}},border}}),
    ];
  }});

  // Fila vacía
  const RSEP = Array(14).fill(cell('',{{}}));

  // Fila totales
  const totBI  = list.reduce((a,d)=>a+d.bi,0);
  const totIGV = list.reduce((a,d)=>a+d.igv,0);
  const totCP  = list.reduce((a,d)=>a+d.total,0);
  const totStyle = (right=false, numFmt=null) => {{
    const s = {{
      font:{{bold:true, sz:10, color:{{rgb:WHITE}}, name:'Arial'}},
      fill:{{fgColor:{{rgb:GREEN}}, patternType:'solid'}},
      alignment:{{horizontal: right?'right':'left', vertical:'center'}},
      border,
    }};
    if (numFmt) s.numFmt = numFmt;
    return s;
  }};
  const RTOT = [
    cell('TOTALES',                 totStyle()),
    cell('',                        totStyle()),
    cell('',                        totStyle()),
    cell('',                        totStyle()),
    cell(list.length+' comprobantes', totStyle(false)),
    cell('',                        totStyle()),
    {{v:totBI,  t:'n', s:totStyle(true, NF)}},
    {{v:totIGV, t:'n', s:totStyle(true, NF)}},
    {{v:totCP,  t:'n', s:totStyle(true, NF)}},
    cell('',                        totStyle()),
  ];

  // ── Armar sheet ───────────────────────────────────────────────────
  const allRows = [R1, R2, R3, R4, ...dataRows, RSEP, RTOT];
  const WS = XLSX.utils.aoa_to_sheet(allRows);

  // Anchos
  WS['!cols'] = [
    {{wch:12}},{{wch:20}},{{wch:8}},{{wch:12}},{{wch:14}},
    {{wch:50}},{{wch:16}},{{wch:14}},{{wch:14}},{{wch:8}},
    {{wch:10}},{{wch:14}},{{wch:10}},{{wch:10}},
  ];

  // Merges título y subtítulo
  WS['!merges'] = [
    {{s:{{r:0,c:0}}, e:{{r:0,c:13}}}},
    {{s:{{r:1,c:0}}, e:{{r:1,c:13}}}},
  ];

  // Alturas
  WS['!rows'] = [
    {{hpt:22}}, // R1 título
    {{hpt:14}}, // R2 subtítulo
    {{hpt:6}},  // R3 vacía
    {{hpt:16}}, // R4 cabeceras
    ...list.map(()=>( {{hpt:14}} )),
    {{hpt:6}},  // sep
    {{hpt:16}}, // totales
  ];

  XLSX.utils.book_append_sheet(WB, WS, 'Pendientes SAP');
  XLSX.writeFile(WB, 'Pendientes_SAP_' + ISO + '.xlsx');
}}

applyFilters();
</script>
</body>
</html>"""


OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'index.html')
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(HTML)
print(f"\n{'='*50}\n  ✅ index.html listo ({len(HTML)//1024} KB)\n  📅 {periodo} · ❌ {len(pendientes)} pendientes · 💰 S/ {total_cp:,.2f}\n{'='*50}")
