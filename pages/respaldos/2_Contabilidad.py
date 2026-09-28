import glob
import io
import json
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Módulo Contable",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 Módulo Contable")

PRESETS_FILE = "vistas_personalizadas_er.json"

# ORDEN PREFERENTE DE EMPRESAS PARA EL MULTISELECT
ORDEN_EMPRESAS_PRIORIDAD = [
    "CIVLAT",
    "SERVYRE",
    "FERVIC",
    "CIV",
    "GPO SERVYRE",
    "FGS",
    "LATIN",
    "EFCO",
    "PESAZA",
    "LIMPIESPIN",
    "SERVYCARGO",
    "CIVMEX",
    "COMANA",
    "PROINA",
    "SEVILAT",
    "IPV",
    "SERSEÑAL",
    "INMOBILIARIA",
    "VIALTECNO",
    "LABORATORIO",
    "FPSB",
    "SIGNAL",
]


def ordenar_empresas_segun_prioridad(lista_empresas):
    def obtener_posicion(emp_nombre):
        emp_clean = str(emp_nombre).strip().upper()
        for idx, pref in enumerate(ORDEN_EMPRESAS_PRIORIDAD):
            if pref == emp_clean or pref in emp_clean:
                return idx
        return 999

    return sorted(lista_empresas, key=obtener_posicion)


# --- GESTIÓN DE VISTAS GUARDADAS (PRESETS) ---
def cargar_vistas_guardadas():
    vistas_predeterminados = {
        "📊 Resumen Ejecutivo": [
            "Ventas Netas Totales",
            "Total Costo",
            "Total Resultado Bruto",
            "Total Gastos",
            "Resultado Antes de Depreciacion",
            "Total Depreciaciones Y Amortización",
            "Total Costo Integral de Financiamiento",
            "Utilidad/Perdida antes de Impuestos",
        ]
    }

    if os.path.exists(PRESETS_FILE):
        try:
            with open(PRESETS_FILE, "r", encoding="utf-8") as f:
                guardadas = json.load(f)
                vistas_predeterminados.update(guardadas)
        except Exception:
            pass

    return vistas_predeterminados


def guardar_nueva_vista(nombre_vista, lista_conceptos):
    vistas = cargar_vistas_guardadas()
    vistas[f"⭐ {nombre_vista}"] = lista_conceptos

    custom_vistas = {k: v for k, v in vistas.items() if k.startswith("⭐ ")}
    with open(PRESETS_FILE, "w", encoding="utf-8") as f:
        json.dump(custom_vistas, f, ensure_ascii=False, indent=2)


# --- BUSCAR BALANZAS POR AÑO Y MES ---
def obtener_estructura_balanzas():
    archivos = glob.glob("Balanzas/**/Balanza.xlsx", recursive=True)
    estructura = []

    for path in archivos:
        path_norm = path.replace("\\", "/")
        partes = path_norm.split("/")

        if len(partes) >= 4:
            anio = partes[-3]
            mes = partes[-2]
            estructura.append({
                'anio': str(anio),
                'mes': str(mes),
                'ruta': path_norm,
                'mtime': os.path.getmtime(path),
            })

    if estructura:
        estructura.sort(key=lambda x: x['mtime'], reverse=True)

    return estructura


@st.cache_data(ttl=300)
def obtener_lista_empresas(ruta):
    xls = pd.ExcelFile(ruta)
    sheets = xls.sheet_names
    if len(sheets) > 1 and "Hoja1" in sheets:
        sheets.remove("Hoja1")

    return ordenar_empresas_segun_prioridad(sheets)


@st.cache_data(ttl=300)
def cargar_hoja_balanza(ruta, nombre_hoja):
    return pd.read_excel(ruta, sheet_name=nombre_hoja)


# --- PLANTILLA DE EXCEL CON AGRUPACIONES ---
@st.cache_data(ttl=3600)
def cargar_plantilla_formato():
    ruta_formato = "FORMATO EDO RESULTADOS.xlsx"
    if not os.path.exists(ruta_formato):
        return []

    wb = openpyxl.load_workbook(ruta_formato, data_only=False)
    sheet = wb['RESULTADOS ACUM'] if 'RESULTADOS ACUM' in wb.sheetnames else wb.active

    plantilla = []
    for i in range(6, sheet.max_row + 1):
        cta = sheet.cell(row=i, column=1).value
        concepto = sheet.cell(row=i, column=2).value
        c_formula = sheet.cell(row=i, column=3).value
        d_formula = sheet.cell(row=i, column=4).value

        str_cta = str(cta).strip() if cta else ""
        str_conc = str(concepto).strip() if concepto else ""

        if str_cta.upper() in ['CUENTA', '1', '2', '3'] or str_conc.upper() in [
            'CONCEPTO',
            '1',
            '2',
            'NONE',
            'NAN',
        ]:
            continue

        outline_lvl = sheet.row_dimensions[i].outlineLevel or 0

        if cta or concepto or c_formula:
            plantilla.append({
                'row_idx': i,
                'cuenta_patron': str_cta if str_cta else None,
                'concepto': str_conc,
                'c_formula': str(c_formula).strip() if c_formula else None,
                'd_formula': str(d_formula).strip() if d_formula else None,
                'outline_level': outline_lvl,
            })
    return plantilla


# --- PARSER NUMÉRICO ROBUSTO ---
def parse_monto_robusto(val):
    if pd.isnull(val):
        return 0.0
    val_str = (
        str(val).replace('$', '').replace(',', '').replace(' ', '').strip()
    )
    if not val_str or val_str.lower() in ('nan', 'none', '-'):
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0


# --- EVALUADOR DE FÓRMULAS EN CASCADA ---
def resolver_todas_las_formulas(mapa_valores, mapa_formulas):
    vals = dict(mapa_valores)

    for _ in range(5):
        for r_idx, f_str in mapa_formulas.items():
            f_clean = f_str.upper().replace(' ', '').replace('$', '')

            m_sub = re.match(r'^=SUBTOTAL\(9,C(\d+):C(\d+)\)$', f_clean)
            if m_sub:
                r_s, r_e = int(m_sub.group(1)), int(m_sub.group(2))
                vals[r_idx] = sum(vals.get(r, 0.0) for r in range(r_s, r_e + 1))
                continue

            expr_raw = re.sub(r'^=\+?', '', f_clean)

            def sustituir_celda(match):
                r_num = int(match.group(1))
                return str(vals.get(r_num, 0.0))

            expr = re.sub(r'C(\d+)', sustituir_celda, expr_raw)

            try:
                if re.match(r'^[0-9\.\+\-\*\/\(\)\s]+$', expr):
                    vals[r_idx] = float(eval(expr))
            except Exception:
                pass
    return vals


# --- COINCIDENCIA FLEXIBLE DE CUENTAS ---
def coincide_cuenta_robusta(cta_balanza, patron_template):
    if not patron_template or patron_template in ('None', 'CUENTA', ''):
        return False

    cb = str(cta_balanza).strip()
    pt = str(patron_template).strip()

    cb_clean = re.sub(r'[^0-9A-Za-z]', '', cb)
    pt_clean = re.sub(r'[^0-9A-Za-z?]', '', pt)

    if cb_clean == pt_clean:
        return True

    p_segs = pt.split('-')
    b_segs = cb.split('-')

    if len(p_segs) >= 3 and len(b_segs) >= 3:
        if p_segs[0] == b_segs[0]:
            p_seg2_is_wildcard = ('?' in p_segs[1]) or (
                p_segs[1] in ('00000', '0000', '000', '99999')
            )

            try:
                seg3_match = (p_segs[2] == b_segs[2]) or (
                    int(p_segs[2]) == int(b_segs[2])
                )
            except ValueError:
                seg3_match = p_segs[2] == b_segs[2]

            if p_seg2_is_wildcard and seg3_match:
                if len(p_segs) >= 4 and len(b_segs) >= 4:
                    try:
                        return (p_segs[3] == b_segs[3]) or (
                            int(p_segs[3]) == int(b_segs[3])
                        )
                    except ValueError:
                        return p_segs[3] == b_segs[3]
                return True

            try:
                if int(p_segs[1]) == int(b_segs[1]) and seg3_match:
                    if len(p_segs) >= 4 and len(b_segs) >= 4:
                        return int(p_segs[3]) == int(b_segs[3])
                    return True
            except ValueError:
                pass

    regex_str = "^" + pt.replace('?', '.').replace('-', r'\\-?') + "$"
    return bool(re.match(regex_str, cb))


# --- BUSCADOR DE MONTO EN BALANZA ---
def obtener_monto_cuenta_balanza(
    patron_template, balanza_records, tipo='mes', solo_intercos=False
):
    pt = str(patron_template).strip()

    # Si estamos en modo Intercos, verificar regla estricta de exclusión:
    if solo_intercos:
        segs_pt = pt.split('-')
        prefix_pt = segs_pt[0].strip() if segs_pt else ''

        # Si el primer nivel (3 dígitos) termina en '1' (ej. 411, 421, 441, 511, 521, 551, 561):
        if len(prefix_pt) == 3 and prefix_pt.endswith('1'):
            # Única excepción autorizada: la 531-?????-054-0000
            if prefix_pt == '531':
                if not (len(segs_pt) >= 3 and segs_pt[2] in ('054', '54', '0054')):
                    return 0.0
            else:
                # Todas las demás cuentas terminadas en 1 (411, 421, 441, 551, etc.) QUEDAN EN CERO / ELIMINADAS
                return 0.0

    p_prefix = pt.split('-')[0].strip() if '-' in pt else pt[:3]
    pt_clean = re.sub(r'[^0-9A-Za-z]', '', pt)

    exact_match = None
    for b in balanza_records:
        cb_clean = re.sub(r'[^0-9A-Za-z]', '', b['cta_raw'])
        if cb_clean == pt_clean:
            exact_match = b
            break

    records_a_sumar = []
    if exact_match:
        records_a_sumar = [exact_match]
    else:
        for b in balanza_records:
            if coincide_cuenta_robusta(b['cta_raw'], pt):
                records_a_sumar.append(b)

    monto = 0.0
    for b in records_a_sumar:
        if tipo == 'mes':
            if p_prefix.startswith(('420', '421', '422', '423', '450', '451')):
                monto += abs(b['cargos_m'] - b['abonos_m'])
            elif p_prefix.startswith(('4', '720', '730')):
                monto += b['abonos_m'] - b['cargos_m']
            else:
                monto += b['cargos_m'] - b['abonos_m']
        else:
            if p_prefix.startswith(('420', '421', '422', '423', '450', '451')):
                monto += abs(b['deudor_f'] - b['acreedor_f'])
            elif p_prefix.startswith(('4', '720', '730')):
                monto += b['acreedor_f'] - b['deudor_f']
            else:
                monto += b['deudor_f'] - b['acreedor_f']

    return monto


# --- CALCULAR VALORES POR EMPRESA ---
def calcular_mapa_valores_empresa(
    df_balanza, plantilla, tipo='mes', solo_intercos=False
):
    if df_balanza.empty or not plantilla:
        return {}

    num_cols = df_balanza.shape[1]

    col_cta = 0
    col_cargos_m = num_cols - 4
    col_abonos_m = num_cols - 3
    col_deudor_f = num_cols - 2
    col_acreedor_f = num_cols - 1

    balanza_records = []
    for idx, row in df_balanza.iterrows():
        cta_raw = str(row.iloc[col_cta]).strip()
        if not cta_raw or cta_raw.lower() in ('nan', 'cuenta', 'none'):
            continue

        if not re.search(r'\d{3}[-\s]?\d{3,5}', cta_raw):
            continue

        cargos_m = parse_monto_robusto(row.iloc[col_cargos_m])
        abonos_m = parse_monto_robusto(row.iloc[col_abonos_m])
        deudor_f = parse_monto_robusto(row.iloc[col_deudor_f])
        acreedor_f = parse_monto_robusto(row.iloc[col_acreedor_f])

        balanza_records.append({
            'cta_raw': cta_raw,
            'cargos_m': cargos_m,
            'abonos_m': abonos_m,
            'deudor_f': deudor_f,
            'acreedor_f': acreedor_f,
        })

    val_map = {}
    formulas_map = {}

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        c_form = row['c_formula']

        if patron:
            val_map[r_idx] = obtener_monto_cuenta_balanza(
                patron,
                balanza_records,
                tipo=tipo,
                solo_intercos=solo_intercos,
            )
        elif c_form and str(c_form).startswith('='):
            formulas_map[r_idx] = c_form
            val_map[r_idx] = 0.0
        else:
            val_map[r_idx] = 0.0

    return resolver_todas_las_formulas(val_map, formulas_map)


# --- GENERADOR MULTIEMPRESA ---
def generar_reporte_multiempresa(
    ruta_balanza,
    empresas_seleccionadas,
    plantilla,
    tipo='mes',
    solo_intercos=False,
):
    if not empresas_seleccionadas or not plantilla:
        return pd.DataFrame()

    mapas_empresas = {}
    for emp in empresas_seleccionadas:
        df_b = cargar_hoja_balanza(ruta_balanza, emp)
        mapas_empresas[emp] = calcular_mapa_valores_empresa(
            df_b, plantilla, tipo=tipo, solo_intercos=solo_intercos
        )

    reporte = []
    incluir_consolidado = len(empresas_seleccionadas) > 1

    current_group_id = 0

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        concepto = row['concepto']
        c_form = row['c_formula']
        outline_lvl = row.get('outline_level', 0)

        if not patron and not concepto and not c_form:
            continue

        es_formula = bool(c_form and str(c_form).startswith('=') and not patron)
        es_cuenta = bool(patron)
        es_dato = es_cuenta or es_formula

        if not es_cuenta:
            current_group_id += 1

        fila_dict = {
            'CUENTA': patron if patron else '',
            'CONCEPTO': concepto,
            'row_idx': r_idx,
            'group_id': f"grp_{current_group_id}",
            'is_detail': es_cuenta,
            'outline_level': outline_lvl,
        }

        monto_total_consolidado = 0.0

        for emp in empresas_seleccionadas:
            m_emp = mapas_empresas[emp].get(r_idx, 0.0)
            ventas_emp = mapas_empresas[emp].get(91, 0.0) or 1.0
            pct_emp = (m_emp / ventas_emp) * 100 if ventas_emp else 0.0

            fila_dict[f"{emp}"] = m_emp if es_dato else None
            fila_dict[f"% {emp}"] = pct_emp if es_dato else None

            if es_dato:
                monto_total_consolidado += m_emp

        if incluir_consolidado:
            tot_ventas_todas = sum(
                mapas_empresas[e].get(91, 0.0) for e in empresas_seleccionadas
            ) or 1.0
            pct_total = (
                (monto_total_consolidado / tot_ventas_todas) * 100
                if tot_ventas_todas
                else 0.0
            )

            fila_dict['TOTAL CONSOLIDADO'] = (
                monto_total_consolidado if es_dato else None
            )
            fila_dict['% TOTAL'] = pct_total if es_dato else None

        es_subtotal_o_total = es_formula or (not patron and bool(concepto))
        fila_dict['es_total'] = es_subtotal_o_total

        reporte.append(fila_dict)

    return pd.DataFrame(reporte)


# --- TABLA INTERACTIVA CON FILTRADO EXACTO DE SELECCIÓN ---
def renderizar_tabla_interactiva_agrupada(
    df_er, empresas, conceptos_a_mostrar=None
):
    if df_er.empty:
        return

    df_disp = df_er.copy()
    if conceptos_a_mostrar and "TODOS" not in conceptos_a_mostrar:
        df_disp = df_disp[df_disp['CONCEPTO'].isin(conceptos_a_mostrar)]

    if df_disp.empty:
        st.warning("No hay rubros seleccionados para mostrar.")
        return

    cols_empresas = list(empresas)
    incluir_consolidado = 'TOTAL CONSOLIDADO' in df_disp.columns

    cols_header = ['CUENTA', 'CONCEPTO']
    for e in cols_empresas:
        cols_header.extend([e, f"% {e}"])

    if incluir_consolidado:
        cols_header.extend(['TOTAL CONSOLIDADO', '% TOTAL'])

    html_code = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: transparent;
            margin: 0;
            padding: 0;
        }
        .tree-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .tree-table th {
            background-color: #1e293b;
            color: #ffffff;
            font-weight: 700;
            padding: 10px 12px;
            text-align: center;
            border: 1px solid #334155;
            font-size: 12px;
            text-transform: uppercase;
        }
        .tree-table td {
            padding: 8px 12px;
            border-bottom: 1px solid #e2e8f0;
            border-right: 1px solid #f1f5f9;
            color: #334155;
        }
        .group-header {
            background-color: #f8fafc;
            font-weight: 700;
            color: #0284c7;
        }
        .grand-total {
            background-color: #e2e8f0;
            font-weight: 800;
            color: #0f172a;
            border-top: 2px solid #475569;
            border-bottom: 2px double #0f172a;
        }
        .detail-row {
            background-color: #ffffff;
        }
        .detail-row:hover {
            background-color: #f8fafc;
        }
        .num-cell {
            text-align: right;
            font-variant-numeric: tabular-nums;
        }
        .text-cell {
            text-align: left;
        }
    </style>
    </head>
    <body>
        <table class="tree-table">
            <thead>
                <tr>
    """

    for h in cols_header:
        html_code += f"<th>{h}</th>"
    html_code += "</tr></thead><tbody>"

    for _, row in df_disp.iterrows():
        es_tot = row.get('es_total', False)
        outline_lvl = row.get('outline_level', 0)
        is_detail = row.get('is_detail', False)

        if es_tot and outline_lvl == 0:
            row_class = "grand-total"
        elif es_tot:
            row_class = "group-header"
        else:
            row_class = "detail-row"

        html_code += f"<tr class='{row_class}'>"
        html_code += f"<td class='text-cell'>{row.get('CUENTA', '')}</td>"

        concepto_str = row.get('CONCEPTO', '')
        if is_detail:
            html_code += f"<td class='text-cell' style='padding-left: 28px;'>{concepto_str}</td>"
        else:
            html_code += f"<td class='text-cell'><strong>{concepto_str}</strong></td>"

        for e in cols_empresas:
            val_m = row.get(e, None)
            val_pct = row.get(f"% {e}", None)
            m_str = (
                f"${val_m:,.2f}"
                if pd.notnull(val_m) and str(val_m) != 'None'
                else ""
            )
            p_str = (
                f"{val_pct:.1f}%"
                if pd.notnull(val_pct) and str(val_pct) != 'None'
                else ""
            )
            html_code += (
                f"<td class='num-cell'>{m_str}</td><td class='num-cell'>{p_str}</td>"
            )

        if incluir_consolidado:
            val_tot = row.get('TOTAL CONSOLIDADO', None)
            val_pct_tot = row.get('% TOTAL', None)
            tot_str = (
                f"${val_tot:,.2f}"
                if pd.notnull(val_tot) and str(val_tot) != 'None'
                else ""
            )
            pct_tot_str = (
                f"{val_pct_tot:.1f}%"
                if pd.notnull(val_pct_tot) and str(val_pct_tot) != 'None'
                else ""
            )
            html_code += f"<td class='num-cell'>{tot_str}</td><td class='num-cell'>{pct_tot_str}</td>"

        html_code += "</tr>"

    html_code += "</tbody></table></body></html>"

    calc_height = max(350, len(df_disp) * 36)
    components.html(html_code, height=calc_height, scrolling=True)


# --- EXPORTADOR A EXCEL FILTRADO ---
def exportar_excel_con_agrupaciones_openpyxl(
    df_er,
    empresas,
    anio,
    mes,
    tipo='mes',
    conceptos_a_mostrar=None,
    titulo_custom=None,
):
    df_exp = df_er.copy()
    if conceptos_a_mostrar and "TODOS" not in conceptos_a_mostrar:
        df_exp = df_exp[df_exp['CONCEPTO'].isin(conceptos_a_mostrar)]

    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EDO_RESULTADOS"

    ws.views.sheetView[0].showGridLines = True

    HEADER_FILL = PatternFill(
        start_color="1E293B", end_color="1E293B", fill_type="solid"
    )
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1E293B")
    SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="475569")

    TOTAL_PRINCIPAL_FILL = PatternFill(
        start_color="E2E8F0", end_color="E2E8F0", fill_type="solid"
    )
    TOTAL_PRINCIPAL_FONT = Font(
        name="Calibri", size=11, bold=True, color="0F172A"
    )

    TOTAL_FILL = PatternFill(
        start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"
    )
    TOTAL_FONT = Font(name="Calibri", size=11, bold=True, color="0284C7")
    REGULAR_FONT = Font(name="Calibri", size=11, color="334155")

    THIN_BORDER = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0'),
    )

    TOTAL_BORDER = Border(
        top=Side(style='thin', color='475569'),
        bottom=Side(style='double', color='0F172A'),
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
    )

    tipo_str = (
        titulo_custom
        if titulo_custom
        else ("DEL MES" if tipo == 'mes' else "ACUMULADO")
    )
    ws.cell(
        row=1,
        column=1,
        value=f"ESTADO DE RESULTADOS CONSOLIDADO ({tipo_str})",
    ).font = TITLE_FONT
    ws.cell(
        row=2,
        column=1,
        value=f"Periodo: {mes}/{anio} | Empresa(s): {', '.join(empresas)}",
    ).font = SUBTITLE_FONT

    start_row = 4
    df_clean = df_exp.drop(
        columns=[
            'es_total',
            'outline_level',
            'group_id',
            'is_detail',
            'row_idx',
        ],
        errors='ignore',
    )
    headers = list(df_clean.columns)

    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=h_text)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    ws.row_dimensions[start_row].height = 26

    for r_i, row_data in df_exp.reset_index(drop=True).iterrows():
        curr_row = start_row + 1 + r_i
        es_total = row_data.get('es_total', False)
        outline_lvl = row_data.get('outline_level', 0)

        for c_i, h_col in enumerate(headers, start=1):
            val = row_data[h_col]
            cell = ws.cell(row=curr_row, column=c_i)

            if pd.isnull(val) or str(val) == 'None':
                cell.value = ""
            else:
                cell.value = val

            if es_total and outline_lvl == 0:
                cell.font = TOTAL_PRINCIPAL_FONT
                cell.fill = TOTAL_PRINCIPAL_FILL
                cell.border = TOTAL_BORDER
            elif es_total:
                cell.font = TOTAL_FONT
                cell.fill = TOTAL_FILL
                cell.border = THIN_BORDER
            else:
                cell.font = REGULAR_FONT
                cell.border = THIN_BORDER

            if h_col in ('CUENTA', 'CONCEPTO'):
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif h_col.startswith('%'):
                cell.alignment = Alignment(
                    horizontal="right", vertical="center"
                )
                cell.number_format = '0.0%'
                if isinstance(val, (int, float)):
                    cell.value = val / 100.0
            else:
                cell.alignment = Alignment(
                    horizontal="right", vertical="center"
                )
                cell.number_format = '$#,##0.00'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 38

    wb.save(output)
    return output.getvalue()


# --- INTERFAZ PRINCIPAL DE STREAMLIT ---
estructura = obtener_estructura_balanzas()
plantilla = cargar_plantilla_formato()

if estructura:
    vistas_disponibles = cargar_vistas_guardadas()

    anios_disponibles = sorted(list(set(x['anio'] for x in estructura)), reverse=True)

    col_a, col_m = st.columns([1, 1])

    with col_a:
        anio_sel = st.selectbox("Año:", anios_disponibles)

    meses_disponibles = sorted(
        list(set(x['mes'] for x in estructura if x['anio'] == anio_sel)),
        reverse=True,
    )

    with col_m:
        mes_sel = st.selectbox("Mes:", meses_disponibles)

    ruta_balanza = next(
        (x['ruta'] for x in estructura if x['anio'] == anio_sel and x['mes'] == mes_sel),
        estructura[0]['ruta'],
    )

    lista_empresas = obtener_lista_empresas(ruta_balanza)

    col_emp1, col_emp2 = st.columns([1, 3])
    with col_emp1:
        st.write("")
        st.write("")
        seleccionar_todas_emp = st.checkbox(
            "☑️ Seleccionar Todas las Empresas", value=False
        )

    with col_emp2:
        default_empresas = (
            lista_empresas
            if seleccionar_todas_emp
            else ([lista_empresas[0]] if lista_empresas else [])
        )

        empresas_seleccionadas = st.multiselect(
            "Empresa(s):",
            lista_empresas,
            default=default_empresas,
        )

    if empresas_seleccionadas:
        df_preview = generar_reporte_multiempresa(
            ruta_balanza, empresas_seleccionadas, plantilla, tipo='mes'
        )
        todos_los_conceptos = []
        if not df_preview.empty:
            todos_los_conceptos = [
                c for c in df_preview['CONCEPTO'].unique() if c
            ]

        st.markdown("---")
        st.subheader("🎯 Configuración de Vista de Filtros")

        col_v1, col_v2 = st.columns([1.5, 3])

        opciones_plantillas = [
            "📊 Resumen Ejecutivo",
            "🔍 Detalle Completo",
            "✏️ Personalizada",
        ] + [k for k in vistas_disponibles.keys() if k.startswith("⭐ ")]

        with col_v1:
            plantilla_sel = st.selectbox(
                "Elegir Vista / Plantilla Guardada:", opciones_plantillas
            )

        conceptos_seleccionados = []

        if plantilla_sel == "🔍 Detalle Completo":
            conceptos_seleccionados = ["TODOS"]
        elif plantilla_sel in vistas_disponibles:
            conceptos_seleccionados = vistas_disponibles[plantilla_sel]
            with col_v2:
                st.info(
                    f"Mostrando **{len(conceptos_seleccionados)}** rubros predefinidos."
                )
        else:
            with col_v2:
                conceptos_seleccionados = st.multiselect(
                    "Selecciona exactamente los rubros que deseas ver:",
                    todos_los_conceptos,
                    default=[
                        "Ventas Netas Totales",
                        "Total Costo",
                        "Total Resultado Bruto",
                        "Total Gastos",
                        "Resultado Antes de Depreciacion",
                    ],
                )

                with st.expander("💾 Guardar esta selección como nueva vista"):
                    col_g1, col_g2 = st.columns([2, 1])
                    with col_g1:
                        nombre_nueva_vista = st.text_input(
                            "Nombre de la Vista:",
                            placeholder="Ej. Mi Filtro Especial",
                        )
                    with col_g2:
                        st.write("")
                        st.write("")
                        if st.button("Guardar Vista"):
                            if nombre_nueva_vista.strip():
                                guardar_nueva_vista(
                                    nombre_nueva_vista.strip(),
                                    conceptos_seleccionados,
                                )
                                st.success(
                                    f"¡Vista '{nombre_nueva_vista}' guardada exitosamente!"
                                )
                                st.rerun()

        st.markdown("---")

        # PESTAÑAS PRINCIPALES DEL MÓDULO CONTABLE
        tab_balanzas, tab_er_mes, tab_er_acum, tab_er_interco = st.tabs([
            "📑 Balanzas de Comprobación",
            "📈 Estado de Resultados (MES)",
            "📊 Estado de Resultados (ACUM)",
            "🔄 Estado de Resultados (Intercos)",
        ])

        with tab_balanzas:
            for emp in empresas_seleccionadas:
                st.subheader(
                    f"Balanza de Comprobación - {emp} ({mes_sel}/{anio_sel})"
                )
                df_b = cargar_hoja_balanza(ruta_balanza, emp)
                st.dataframe(df_b, use_container_width=True, hide_index=True)

        with tab_er_mes:
            empresas_str = ", ".join(empresas_seleccionadas)
            st.subheader(
                f"Estado de Resultados (DEL MES) - [{empresas_str}] ({mes_sel}/{anio_sel})"
            )

            if not plantilla:
                st.error("No se encontró el archivo 'FORMATO EDO RESULTADOS.xlsx' en la raíz.")
            else:
                with st.spinner("Procesando Estado de Resultados..."):
                    df_er_mes = generar_reporte_multiempresa(
                        ruta_balanza,
                        empresas_seleccionadas,
                        plantilla,
                        tipo='mes',
                    )

                if not df_er_mes.empty:
                    renderizar_tabla_interactiva_agrupada(
                        df_er_mes,
                        empresas_seleccionadas,
                        conceptos_a_mostrar=conceptos_seleccionados,
                    )

                    excel_mes = exportar_excel_con_agrupaciones_openpyxl(
                        df_er_mes,
                        empresas_seleccionadas,
                        anio_sel,
                        mes_sel,
                        tipo='mes',
                        conceptos_a_mostrar=conceptos_seleccionados,
                    )
                    st.download_button(
                        label=f"📥 Descargar ER Mes en Excel ({mes_sel}_{anio_sel})",
                        data=excel_mes,
                        file_name=f"Estado_Resultados_MES_Filtrado_{mes_sel}_{anio_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with tab_er_acum:
            empresas_str = ", ".join(empresas_seleccionadas)
            st.subheader(
                f"Estado de Resultados (ACUMULADO) - [{empresas_str}] ({mes_sel}/{anio_sel})"
            )

            if not plantilla:
                st.error("No se encontró el archivo 'FORMATO EDO RESULTADOS.xlsx' en la raíz.")
            else:
                with st.spinner("Procesando Estado de Resultados Acumulado..."):
                    df_er_acum = generar_reporte_multiempresa(
                        ruta_balanza,
                        empresas_seleccionadas,
                        plantilla,
                        tipo='acum',
                    )

                if not df_er_acum.empty:
                    renderizar_tabla_interactiva_agrupada(
                        df_er_acum,
                        empresas_seleccionadas,
                        conceptos_a_mostrar=conceptos_seleccionados,
                    )

                    excel_acum = exportar_excel_con_agrupaciones_openpyxl(
                        df_er_acum,
                        empresas_seleccionadas,
                        anio_sel,
                        mes_sel,
                        tipo='acum',
                        conceptos_a_mostrar=conceptos_seleccionados,
                    )
                    st.download_button(
                        label=f"📥 Descargar ER Acumulado en Excel ({mes_sel}_{anio_sel})",
                        data=excel_acum,
                        file_name=f"Estado_Resultados_ACUM_Filtrado_{mes_sel}_{anio_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with tab_er_interco:
            empresas_str = ", ".join(empresas_seleccionadas)
            st.subheader(
                f"Estado de Resultados (INTERCOMPAÑÍAS ACUMULADO) - [{empresas_str}] ({mes_sel}/{anio_sel})"
            )
            st.info(
                "💡 **Filtro Aplicado:** Se eliminan todas las cuentas de Intercompañías (`411`, `421`, `441`, `511`, `521`, `551`, `561`), conservando ÚNICAMENTE la cuenta `531-?????-054-0000` y las cuentas a terceros."
            )

            if not plantilla:
                st.error("No se encontró el archivo 'FORMATO EDO RESULTADOS.xlsx' en la raíz.")
            else:
                with st.spinner("Procesando Estado de Resultados Intercompañías (Acumulado)..."):
                    df_er_interco = generar_reporte_multiempresa(
                        ruta_balanza,
                        empresas_seleccionadas,
                        plantilla,
                        tipo='acum',
                        solo_intercos=True,
                    )

                if not df_er_interco.empty:
                    renderizar_tabla_interactiva_agrupada(
                        df_er_interco,
                        empresas_seleccionadas,
                        conceptos_a_mostrar=conceptos_seleccionados,
                    )

                    excel_interco = exportar_excel_con_agrupaciones_openpyxl(
                        df_er_interco,
                        empresas_seleccionadas,
                        anio_sel,
                        mes_sel,
                        tipo='acum',
                        conceptos_a_mostrar=conceptos_seleccionados,
                        titulo_custom="INTERCOMPAÑIAS ACUMULADO",
                    )
                    st.download_button(
                        label=f"📥 Descargar ER Intercompañías Acumulado en Excel ({mes_sel}_{anio_sel})",
                        data=excel_interco,
                        file_name=f"Estado_Resultados_INTERCOS_ACUM_{mes_sel}_{anio_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
else:
    st.info("Por favor selecciona al menos una empresa para mostrar el reporte.")