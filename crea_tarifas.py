import io
import math
import zipfile
from datetime import date

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Generador de Tarifas Turaco",
    page_icon="📊",
    layout="centered",
)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS (lógica de cálculo)
# ─────────────────────────────────────────────────────────────────────────────

def parse_euro(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s.lower() in ('no vendible', 'no_vendible', ''):
        return None
    s = s.replace('€', '').replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def pvp_min(cost_log):
    return math.ceil(cost_log) - 0.1


def pvp_pub(pvp_min_val, pvpr):
    return pvp_min_val if pvp_min_val > pvpr else pvpr


def rentabilidad(neto, porte, pvp_sin_iva, comision_eur):
    denom = pvp_sin_iva - comision_eur
    if denom == 0:
        return None
    return 1 - (neto + porte) / denom


def num_cell(ws, row, col, value, fmt='#,##0.00'):
    cell = ws.cell(row=row, column=col, value=value)
    cell.number_format = fmt
    return cell


def pct_cell(ws, row, col, value):
    cell = ws.cell(row=row, column=col, value=value)
    cell.number_format = '0.00%'
    return cell


# ─────────────────────────────────────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_product_info(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb['Product_info']
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    headers = [str(h).strip() if h else '' for h in rows[0]]
    col = {h: i for i, h in enumerate(headers)}
    products = {}
    for row in rows[1:]:
        if not row[col.get('EAN', 1)]:
            continue
        try:
            ean = str(int(float(str(row[col['EAN']])))) if row[col['EAN']] else None
        except (ValueError, TypeError):
            continue
        if not ean:
            continue
        products[ean] = {
            'ref':        row[col['REFERENCIA']],
            'titulo':     row[col.get('TÍTULO DE PRODUCTO', col.get('TITULO DE PRODUCTO', 2))],
            'familia':    row[col.get('FAMILIA', 4)],
            'subfamilia': row[col.get('SUBFAMILIA', 5)],
            'com_amz':    row[col.get('COM_AZM', 7)],
            'com_mir':    row[col.get('COM_MIR', 8)],
            'com_mm':     row[col.get('COM_MM', 9)],
            'com_priv':   row[col.get('COM_PRIV', 10)],
            'com_c4':     row[col.get('COM_C4', 11)],
        }
    wb.close()
    return products


@st.cache_data(show_spinner=False)
def load_tarifa_nacional(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    tarifa = {}
    col = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i == 1:
            continue
        if i == 2:
            headers = [str(h).strip() if h else '' for h in row]
            col = {h: j for j, h in enumerate(headers)}
            continue
        ean_raw = row[col.get('EAN', 5)]
        if not ean_raw:
            continue
        try:
            ean = str(int(float(ean_raw)))
        except (ValueError, TypeError):
            continue
        pvpr  = row[col.get('PVPR ', col.get('PVPR', 10))]
        neto  = row[col.get('NETO', 11)]
        porte = row[col.get('PORTES', col.get('PORTE', 12))]
        if pvpr is None or neto is None:
            continue
        tarifa[ean] = {
            'pvpr':  float(pvpr),
            'neto':  float(neto),
            'porte': float(porte) if porte else 0.0,
        }
    wb.close()
    return tarifa


# ─────────────────────────────────────────────────────────────────────────────
# GENERACIÓN NACIONAL
# ─────────────────────────────────────────────────────────────────────────────

MARKETPLACES = [
    ('T_AMZ',  '(%) COM_AMZ', 'com_amz',  0.1815,   'FF6600'),
    ('T_MIR',  '(%) COM_MIR', 'com_mir',  0.15,     '0070C0'),
    ('T_C4',   '(%) COM_C4',  'com_c4',   0.10,     'C00000'),
    ('T_MM',   '(%) COM_MM',  'com_mm',   0.14956,  '00B050'),
    ('T_PRIV', '(%) COM_PRIV','com_priv', 0.1452,   '7030A0'),
]

NAC_HEADERS = [
    'REFERENCIA', 'EAN', 'NOMBRE COMPLETO', 'FAMILIA', 'SUBFAMILIA',
    'PVPR ', 'NETO', 'PORTES', 'MARGEN (%)', 'MARGEN (€)',
    'COST_log', 'COM (%)', 'COM (€)', 'IVA (21%)',
    'PVP MIN.', 'PVP PUB.', 'PVP SIN IVA', 'COMISION', 'RENTABILIDAD', 'DESPOSICIONADO'
]


def build_nacional(tarifa, products, margen=0.10):
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_id, com_sheet_id, com_key, com_default, bg in MARKETPLACES:
        ws = wb.create_sheet(sheet_id)
        fill = PatternFill('solid', start_color=bg)
        font_h = Font(bold=True, color='FFFFFF', name='Arial', size=10)

        for j, h in enumerate(NAC_HEADERS, 1):
            cell = ws.cell(row=1, column=j, value=h)
            cell.fill = fill
            cell.font = font_h
            cell.alignment = Alignment(horizontal='center')

        ws_com = wb.create_sheet(com_sheet_id)
        ws_com.cell(row=1, column=1, value='REFERENCIA').font = Font(bold=True)
        ws_com.cell(row=1, column=2, value='COMISION (%)').font = Font(bold=True)

        com_row = 2
        row = 2
        for ean, prod in products.items():
            if ean not in tarifa:
                continue
            tar = tarifa[ean]
            ref      = prod['ref']
            pvpr     = tar['pvpr']
            neto     = tar['neto']
            porte    = tar['porte']
            com_pct  = prod.get(com_key) or com_default

            margen_eur  = neto / (1 - margen)
            cost_log    = margen_eur + porte
            com_eur     = cost_log / (1 - com_pct)
            iva_eur     = com_eur * 1.21
            pmin        = pvp_min(iva_eur)
            ppub        = pvp_pub(pmin, pvpr)
            pvp_sin_iva = ppub / 1.21
            comision    = pvp_sin_iva * com_pct
            rent        = rentabilidad(neto, porte, pvp_sin_iva, comision)
            despos      = ppub > pvpr

            ws.cell(row=row, column=1,  value=ref)
            ws.cell(row=row, column=2,  value=int(ean) if ean.isdigit() else ean)
            ws.cell(row=row, column=3,  value=prod['titulo'])
            ws.cell(row=row, column=4,  value=prod['familia'])
            ws.cell(row=row, column=5,  value=prod['subfamilia'])
            num_cell(ws, row, 6,  pvpr)
            num_cell(ws, row, 7,  neto)
            num_cell(ws, row, 8,  porte)
            pct_cell(ws, row, 9,  margen)
            num_cell(ws, row, 10, margen_eur)
            num_cell(ws, row, 11, cost_log)
            pct_cell(ws, row, 12, com_pct)
            num_cell(ws, row, 13, com_eur)
            num_cell(ws, row, 14, iva_eur)
            num_cell(ws, row, 15, pmin)
            num_cell(ws, row, 16, ppub)
            num_cell(ws, row, 17, pvp_sin_iva)
            num_cell(ws, row, 18, comision)
            pct_cell(ws, row, 19, rent)
            ws.cell(row=row, column=20, value='SI' if despos else 'NO')

            ws_com.cell(row=com_row, column=1, value=ref)
            ws_com.cell(row=com_row, column=2, value=com_pct)
            com_row += 1
            row += 1

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['C'].width = 40
        for l in ['D', 'E']:
            ws.column_dimensions[l].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, row - 2


# ─────────────────────────────────────────────────────────────────────────────
# GENERACIÓN INTER
# ─────────────────────────────────────────────────────────────────────────────

MARGEN_INTER     = 0.05
MARGEN_INTER_BIG = 0.10
CONV_PL = 5
CONV_SE = 11

INTER_BG = {
    'ES-FR': 'C00000', 'FR-FR': 'FF0000',
    'ES-IT': '00B050', 'IT-IT': '008000',
    'ES-DE': '0070C0', 'DE-DE': '0000FF',
    'PT':    'FFC000', 'NL':    '7030A0',
    'BE':    '00B0F0', 'PL':    'FF6600', 'SE':    '4BACC6',
}

INTER_HDR_BASE = ['NOMBRE COMPLETO', 'Tipo Tarifa', 'REFERENCIA', 'EAN',
                  'PVP', 'NETO ES', 'PORTE ES', 'MARGEN', 'COST_log',
                  'COMISION', 'IVA', 'PVP MIN.', 'PVP PUB.']


def calc_inter_row(pvp, neto, porte, iva_factor, margen, com_pct=0.1815):
    if pvp is None or neto is None or porte is None:
        return None
    margen_eur = neto / (1 - margen)
    cost_log   = margen_eur + porte
    comision   = cost_log / (1 - com_pct)
    iva_val    = comision * iva_factor
    pmin       = pvp_min(iva_val)
    ppub       = pvp_pub(pmin, pvp)
    return {'margen': margen_eur, 'cost': cost_log,
            'comis': comision, 'iva': iva_val, 'pmin': pmin, 'ppub': ppub}


def load_inter_sheet(wb, sheet_name, col_pvp, col_porte_es, col_neto_es,
                     col_porte_loc=None, col_neto_loc=None, data_start=3):
    ws = wb[sheet_name]
    records = []
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        if not row or not row[0]:
            continue
        ean_raw = row[4]
        try:
            ean = str(int(float(str(ean_raw).replace(',', '.').replace(' ', '')))) if ean_raw else None
        except:
            ean = str(ean_raw) if ean_raw else None
        records.append({
            'nombre':    row[0],
            'tipo':      row[2],
            'ref_str':   str(row[3]) if row[3] else '',
            'ean':       ean,
            'pvp':       parse_euro(row[col_pvp]),
            'porte_es':  parse_euro(row[col_porte_es]),
            'neto_es':   parse_euro(row[col_neto_es]),
            'porte_loc': parse_euro(row[col_porte_loc]) if col_porte_loc is not None else None,
            'neto_loc':  parse_euro(row[col_neto_loc])  if col_neto_loc  is not None else None,
        })
    return records


def load_zona(wb, sheet_name, col_pvp=5, col_porte=6, col_neto=7):
    ws = wb[sheet_name]
    recs = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 8 or not row[0]:
            continue
        ean_raw = row[4]
        try:
            ean = str(int(float(str(ean_raw).replace(',', '.').replace(' ', '')))) if ean_raw else None
        except:
            ean = str(ean_raw) if ean_raw else None
        recs.append({
            'nombre': row[0], 'tipo': row[2],
            'ref_str': str(row[3]) if row[3] else '', 'ean': ean,
            'pvp': parse_euro(row[col_pvp]),
            'porte_es': parse_euro(row[col_porte]),
            'neto_es': parse_euro(row[col_neto]),
            'porte_loc': None, 'neto_loc': None,
        })
    return recs


def write_inter_ws(ws, records, ean_map, sheet_key, iva_factor, margen,
                   local_prefix=None, conv_factor=None):
    bg = INTER_BG.get(sheet_key, '1F4E79')
    hdrs = INTER_HDR_BASE[:]
    if local_prefix:
        hdrs[5] = 'NETO LOC'
        hdrs[6] = 'PORTE LOC'
        hdrs.append('SKU LOCAL')
    if conv_factor:
        hdrs.append('PVP (MONEDA LOCAL)')

    fill = PatternFill('solid', start_color=bg)
    font_h = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    for j, h in enumerate(hdrs, 1):
        cell = ws.cell(row=1, column=j, value=h)
        cell.fill = fill
        cell.font = font_h
        cell.alignment = Alignment(horizontal='center')

    row = 2
    for rec in records:
        ean   = rec['ean']
        neto  = rec['neto_loc'] if local_prefix else rec['neto_es']
        porte = rec['porte_loc'] if local_prefix else rec['porte_es']
        pvp   = rec['pvp']
        if pvp is None or neto is None or porte is None:
            continue
        m = MARGEN_INTER if neto < 30 else MARGEN_INTER_BIG
        calc = calc_inter_row(pvp, neto, porte, iva_factor, m)
        if not calc:
            continue

        prod    = ean_map.get(ean, {})
        ref_num = prod.get('ref', rec['ref_str'])

        ws.cell(row=row, column=1, value=rec['nombre'])
        ws.cell(row=row, column=2, value=rec['tipo'])
        ws.cell(row=row, column=3, value=ref_num)
        ws.cell(row=row, column=4, value=int(ean) if ean and ean.isdigit() else ean)
        num_cell(ws, row, 5,  pvp)
        num_cell(ws, row, 6,  neto)
        num_cell(ws, row, 7,  porte)
        num_cell(ws, row, 8,  calc['margen'])
        num_cell(ws, row, 9,  calc['cost'])
        num_cell(ws, row, 10, calc['comis'])
        num_cell(ws, row, 11, calc['iva'])
        num_cell(ws, row, 12, calc['pmin'])
        num_cell(ws, row, 13, calc['ppub'])

        col14 = 14
        if local_prefix and ref_num:
            ref_s = str(ref_num).replace(' ', '')
            ws.cell(row=row, column=col14, value=f"{local_prefix}{ref_s.zfill(5)}")
            col14 += 1
        if conv_factor:
            ws.cell(row=row, column=col14, value=round(calc['ppub'] * conv_factor, 1))

        row += 1

    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['C'].width = 20
    return row - 2


def build_inter(inter_bytes, products):
    wb_in  = load_workbook(io.BytesIO(inter_bytes), read_only=True, data_only=True)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    fr_recs = load_inter_sheet(wb_in, 'Francia',  5, 6, 7, 8, 9)
    it_recs = load_inter_sheet(wb_in, 'Italia',   5, 6, 7, 8, 9)
    de_recs = load_inter_sheet(wb_in, 'Alemania', 5, 6, 7, 8, 9)
    pt_recs = load_zona(wb_in, 'Portugal', 5, 6, 7)
    z3_recs = load_zona(wb_in, 'Zona 3',   5, 6, 7)
    z4_recs = load_zona(wb_in, 'Zona 4',   5, 6, 7)
    z5_recs = load_zona(wb_in, 'Zona 5',   5, 6, 7)
    wb_in.close()

    counts = {}
    sheets = [
        ('ES-FR', fr_recs, 1.20, None,  None),
        ('FR-FR', fr_recs, 1.20, 'FR',  None),
        ('ES-IT', it_recs, 1.22, None,  None),
        ('IT-IT', it_recs, 1.22, 'IT',  None),
        ('ES-DE', de_recs, 1.19, None,  None),
        ('DE-DE', de_recs, 1.19, 'DE',  None),
        ('PT',    pt_recs, 1.23, None,  None),
        ('NL',    z3_recs, 1.21, None,  None),
        ('BE',    z3_recs, 1.21, None,  None),
        ('PL',    z4_recs, 1.23, None,  CONV_PL),
        ('SE',    z5_recs, 1.25, None,  CONV_SE),
    ]
    for key, recs, iva, prefix, conv in sheets:
        ws = wb_out.create_sheet(key)
        n  = write_inter_ws(ws, recs, products, key, iva, MARGEN_INTER_BIG, prefix, conv)
        counts[key] = n

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf, counts


# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────

st.title("📊 Generador de Tarifas Turaco")
st.markdown("Sube los tres ficheros de entrada y descarga las tarifas completas generadas automáticamente.")

st.divider()

col1, col2 = st.columns(2)
with col1:
    f_calc    = st.file_uploader("🗂️ Calculadora (Product_info + comisiones)", type="xlsx",
                                  help="CALCULADORA_TARIFA_NACIONAL.xlsx")
    f_nacional = st.file_uploader("🇪🇸 Tarifa Nacional de Sistema", type="xlsx",
                                   help="TARIFA_TURACO_-_(MES)_2026.xlsx")
with col2:
    f_inter   = st.file_uploader("🌍 Tarifa Internacional de Sistema", type="xlsx",
                                  help="TARIFA_TURACO_INTER_-_(MES)2026.xlsx")
    margen    = st.number_input("Margen (%)", min_value=1, max_value=30, value=10,
                                 help="Margen de seguridad aplicado sobre el NETO. Por defecto 10%.") / 100

st.divider()

if st.button("⚡ Generar tarifas", type="primary",
             disabled=not (f_calc and f_nacional and f_inter)):

    with st.status("Procesando…", expanded=True) as status:

        st.write("📥 Cargando Product_info de la calculadora…")
        calc_bytes = f_calc.read()
        products   = load_product_info(calc_bytes)
        st.write(f"   ✅ {len(products):,} productos cargados")

        st.write("📥 Cargando tarifa nacional…")
        nac_bytes = f_nacional.read()
        tarifa    = load_tarifa_nacional(nac_bytes)
        st.write(f"   ✅ {len(tarifa):,} líneas cargadas")

        st.write("🔧 Generando TARIFA NACIONAL COMPLETA…")
        buf_nac, n_nac = build_nacional(tarifa, products, margen=margen)
        st.write(f"   ✅ {n_nac:,} productos × 5 marketplaces")

        st.write("🔧 Generando TARIFA INTER COMPLETA…")
        inter_bytes   = f_inter.read()
        buf_int, counts = build_inter(inter_bytes, products)
        for k, v in counts.items():
            st.write(f"   ✅ {k}: {v:,} filas")

        status.update(label="✅ ¡Tarifas generadas!", state="complete")

    st.divider()
    fecha = date.today().strftime('%d_%m_%Y')

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            label="⬇️ Tarifa Nacional Completa",
            data=buf_nac,
            file_name=f"TARIFA_NACIONAL_COMPLETA_{fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            label="⬇️ Tarifa Inter Completa",
            data=buf_int,
            file_name=f"TARIFA_TURACO_INTER_COMPLETA_{fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c3:
        # ZIP con los dos juntos
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            buf_nac.seek(0)
            buf_int.seek(0)
            zf.writestr(f"TARIFA_NACIONAL_COMPLETA_{fecha}.xlsx", buf_nac.read())
            zf.writestr(f"TARIFA_TURACO_INTER_COMPLETA_{fecha}.xlsx", buf_int.read())
        zip_buf.seek(0)
        st.download_button(
            label="⬇️ Descargar todo (.zip)",
            data=zip_buf,
            file_name=f"TARIFAS_TURACO_{fecha}.zip",
            mime="application/zip",
            use_container_width=True,
        )

elif not (f_calc and f_nacional and f_inter):
    st.info("👆 Sube los tres ficheros para habilitar la generación.", icon="ℹ️")
