import glob
import io
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo Contable",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 Módulo Contable")


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
    return sheets


@st.cache_data(ttl=300)
def cargar_hoja_balanza(ruta, nombre_hoja):
    return pd.read_excel(ruta, sheet_name=nombre_hoja)


# --- PLANTILLA DE EXCEL ---
@st.cache_data(ttl=3600)
def cargar_plantilla_formato():
    ruta_formato = "FORMATO EDO RESULTADOS.xlsx"
    if not os.path.exists(ruta_formato):
        return []

    wb = openpyxl.load_workbook(ruta_formato, data_only=False)
    sheet = wb['RESULTADOS ACUM'] if 'RESULTADOS ACUM' in wb.sheetnames else wb.active

    plantilla = []
    for i in range(4, sheet.max_row + 1):
        cta = sheet.cell(row=i, column=1).value
        concepto = sheet.cell(row=i, column=2).value
        c_formula = sheet.cell(row=i, column=3).value
        d_formula = sheet.cell(row=i, column=4).value

        if cta or concepto or c_formula:
            plantilla.append({
                'row_idx': i,
                'cuenta_patron': str(cta).strip() if cta else None,
                'concepto': str(concepto).strip() if concepto else '',
                'c_formula': str(c_formula).strip() if c_formula else None,
                'd_formula': str(d_formula).strip() if d_formula else None,
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
def obtener_monto_cuenta_balanza(patron_template, balanza_records, tipo='mes'):
    pt = str(patron_template).strip()
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
def calcular_mapa_valores_empresa(df_balanza, plantilla, tipo='mes'):
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
                patron, balanza_records, tipo=tipo
            )
        elif c_form and str(c_form).startswith('='):
            formulas_map[r_idx] = c_form
            val_map[r_idx] = 0.0
        else:
            val_map[r_idx] = 0.0

    return resolver_todas_las_formulas(val_map, formulas_map)


# --- GENERADOR MULTIEMPRESA ---
def generar_reporte_multiempresa(
    ruta_balanza, empresas_seleccionadas, plantilla, tipo='mes'
):
    if not empresas_seleccionadas or not plantilla:
        return pd.DataFrame()

    mapas_empresas = {}
    for emp in empresas_seleccionadas:
        df_b = cargar_hoja_balanza(ruta_balanza, emp)
        mapas_empresas[emp] = calcular_mapa_valores_empresa(
            df_b, plantilla, tipo=tipo
        )

    reporte = []
    incluir_consolidado = len(empresas_seleccionadas) > 1

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        concepto = row['concepto']
        c_form = row['c_formula']

        if not patron and not concepto and not c_form:
            continue

        es_formula = bool(c_form and str(c_form).startswith('=') and not patron)
        es_cuenta = bool(patron)
        es_dato = es_cuenta or es_formula

        fila_dict = {
            'CUENTA': patron if patron else '',
            'CONCEPTO': concepto,
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

        es_subtotal_o_total = any(
            kw in concepto.lower()
            for kw in ['total', 'ventas a', 'utilidad', 'pérdida', 'netas', 'descuentos s/']
        )

        fila_dict['es_total'] = es_subtotal_o_total or es_formula
        fila_dict['es_encabezado'] = (
            not patron and not es_formula and bool(concepto)
        )

        reporte.append(fila_dict)

    return pd.DataFrame(reporte)


# --- FORMATO ESTILO FINANZAS (CLARO Y NEGRITAS EN SUMAS) ---
def renderizar_tabla_estilo_finanzas(df_er, empresas):
    if df_er.empty:
        return

    df_disp = df_er.copy()

    # Función para resaltar totales en negrita manteniendo el tema claro tipo Finanzas
    def aplicar_estilo_finanzas(row):
        styles = [''] * len(row)
        es_total = row.get('es_total', False)
        es_encabezado = row.get('es_encabezado', False)

        if es_total:
            # Resaltado en NEGRITA con fondo suave gris/azul para las sumas y totales
            styles = [
                'font-weight: 800; background-color: #f1f5f9; color: #0f172a; border-top: 1.5px solid #64748b; border-bottom: 2px double #0f172a;'
            ] * len(row)
        elif es_encabezado:
            # Títulos de sección en negrita azul
            styles = [
                'font-weight: 700; background-color: #f8fafc; color: #0369a1; text-transform: uppercase;'
            ] * len(row)
        else:
            # Filas normales
            styles = ['font-weight: 400; color: #334155;'] * len(row)

        return styles

    df_clean = df_disp.drop(columns=['es_total', 'es_encabezado'])

    format_dict = {}
    for emp in empresas:
        format_dict[emp] = (
            lambda x: f"${x:,.2f}" if pd.notnull(x) and str(x) != 'None' else ''
        )
        format_dict[f"% {emp}"] = (
            lambda x: f"{x:.1f}%" if pd.notnull(x) and str(x) != 'None' else ''
        )

    if 'TOTAL CONSOLIDADO' in df_clean.columns:
        format_dict['TOTAL CONSOLIDADO'] = (
            lambda x: f"${x:,.2f}" if pd.notnull(x) and str(x) != 'None' else ''
        )
        format_dict['% TOTAL'] = (
            lambda x: f"{x:.1f}%" if pd.notnull(x) and str(x) != 'None' else ''
        )

    styler = (
        df_clean.style.apply(aplicar_estilo_finanzas, axis=1)
        .format(format_dict)
        .set_properties(
            **{
                'padding': '6px 12px',
                'font-family': '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
                'font-size': '13px',
                'border': '1px solid #e2e8f0',
            }
        )
    )

    st.dataframe(styler, use_container_width=True, hide_index=True)


# --- EXPORTADOR A EXCEL EJECUTIVO (OPENPYXL) ---
def exportar_excel_ejecutivo_openpyxl(df_er, empresas, anio, mes, tipo='mes'):
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

    TOTAL_FILL = PatternFill(
        start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"
    )
    TOTAL_FONT = Font(name="Calibri", size=11, bold=True, color="0F172A")
    HEADER_ROW_FONT = Font(name="Calibri", size=11, bold=True, color="0284C7")

    REGULAR_FONT = Font(name="Calibri", size=11, color="1E293B")

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

    tipo_str = "DEL MES" if tipo == 'mes' else "ACUMULADO"
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
    df_clean = df_er.drop(
        columns=['es_total', 'es_encabezado'], errors='ignore'
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

    for r_i, row_data in df_er.iterrows():
        curr_row = start_row + 1 + r_i
        es_total = row_data.get('es_total', False)
        es_encabezado = row_data.get('es_encabezado', False)

        for c_i, h_col in enumerate(headers, start=1):
            val = row_data[h_col]
            cell = ws.cell(row=curr_row, column=c_i)

            if pd.isnull(val) or str(val) == 'None':
                cell.value = ""
            else:
                cell.value = val

            if es_total:
                cell.font = TOTAL_FONT
                cell.fill = TOTAL_FILL
                cell.border = TOTAL_BORDER
            elif es_encabezado:
                cell.font = HEADER_ROW_FONT
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


# --- INTERFAZ STREAMLIT ---
estructura = obtener_estructura_balanzas()
plantilla = cargar_plantilla_formato()

if estructura:
    anios_disponibles = sorted(list(set(x['anio'] for x in estructura)), reverse=True)

    col_a, col_m, col_e = st.columns([1, 1, 2])

    with col_a:
        anio_sel = st.selectbox("Selecciona Año:", anios_disponibles)

    meses_disponibles = sorted(
        list(set(x['mes'] for x in estructura if x['anio'] == anio_sel)),
        reverse=True,
    )

    with col_m:
        mes_sel = st.selectbox("Selecciona Mes:", meses_disponibles)

    ruta_balanza = next(
        (x['ruta'] for x in estructura if x['anio'] == anio_sel and x['mes'] == mes_sel),
        estructura[0]['ruta'],
    )

    lista_empresas = obtener_lista_empresas(ruta_balanza)

    with col_e:
        empresas_seleccionadas = st.multiselect(
            "Selecciona Empresa(s):",
            lista_empresas,
            default=[lista_empresas[0]] if lista_empresas else [],
        )

    if empresas_seleccionadas:
        tab_balanzas, tab_er_mes, tab_er_acum = st.tabs([
            "📑 Balanzas de Comprobación",
            "📈 Estado de Resultados (MES)",
            "📊 Estado de Resultados (ACUM)",
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
                with st.spinner("Procesando Estado de Resultados Comparativo (Del Mes)..."):
                    df_er_mes = generar_reporte_multiempresa(
                        ruta_balanza,
                        empresas_seleccionadas,
                        plantilla,
                        tipo='mes',
                    )

                if not df_er_mes.empty:
                    renderizar_tabla_estilo_finanzas(
                        df_er_mes, empresas_seleccionadas
                    )

                    excel_mes = exportar_excel_ejecutivo_openpyxl(
                        df_er_mes,
                        empresas_seleccionadas,
                        anio_sel,
                        mes_sel,
                        tipo='mes',
                    )
                    st.download_button(
                        label=f"📥 Descargar ER Mes en Excel ({mes_sel}_{anio_sel})",
                        data=excel_mes,
                        file_name=f"Estado_Resultados_MES_Ejecutivo_{mes_sel}_{anio_sel}.xlsx",
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
                with st.spinner("Procesando Estado de Resultados Comparativo (Acumulado)..."):
                    df_er_acum = generar_reporte_multiempresa(
                        ruta_balanza,
                        empresas_seleccionadas,
                        plantilla,
                        tipo='acum',
                    )

                if not df_er_acum.empty:
                    renderizar_tabla_estilo_finanzas(
                        df_er_acum, empresas_seleccionadas
                    )

                    excel_acum = exportar_excel_ejecutivo_openpyxl(
                        df_er_acum,
                        empresas_seleccionadas,
                        anio_sel,
                        mes_sel,
                        tipo='acum',
                    )
                    st.download_button(
                        label=f"📥 Descargar ER Acumulado en Excel ({mes_sel}_{anio_sel})",
                        data=excel_acum,
                        file_name=f"Estado_Resultados_ACUM_Ejecutivo_{mes_sel}_{anio_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
else:
    st.info("Por favor selecciona al menos una empresa para mostrar el reporte.")