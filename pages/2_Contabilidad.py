import glob
import io
import os
import re
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo Contable", layout="wide", initial_sidebar_state="expanded"
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


# --- BUSCADOR DE MONTO CON PRIORIDAD ---
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


# --- GENERADOR DE ESTADOS DE RESULTADOS ---
def generar_reporte_estado_resultados(df_balanza, plantilla, tipo='mes'):
    if not plantilla or df_balanza.empty:
        return pd.DataFrame()

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

    val_map = resolver_todas_las_formulas(val_map, formulas_map)

    ventas_totales = val_map.get(91, 0.0) or 1.0
    col_monto_hdr = 'DEL MES' if tipo == 'mes' else 'ACUMULADO'
    col_pct_hdr = '% MES' if tipo == 'mes' else '% ACUM'

    reporte = []
    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        concepto = row['concepto']
        c_form = row['c_formula']

        v = val_map.get(r_idx, 0.0)

        if not patron and not concepto and not c_form:
            continue

        pct = (v / ventas_totales) * 100 if ventas_totales else 0.0
        es_formula = bool(c_form and str(c_form).startswith('=') and not patron)
        es_cuenta = bool(patron)

        es_subtotal_o_total = any(
            kw in concepto.lower()
            for kw in ['total', 'ventas a', 'utilidad', 'pérdida', 'netas']
        )

        reporte.append({
            'CUENTA': patron if patron else '',
            'CONCEPTO': concepto,
            col_monto_hdr: v if (es_cuenta or es_formula) else None,
            col_pct_hdr: pct if (es_cuenta or es_formula) else None,
            'es_total': es_subtotal_o_total or es_formula,
            'es_encabezado': not patron and not es_formula and bool(concepto),
        })

    return pd.DataFrame(reporte)


# --- FORMATO DE TABLA CON FORMATO EXCEL Y ESTILOS ---
def renderizar_tabla_estilo_excel(df_er, col_monto, col_pct):
    if df_er.empty:
        return

    df_disp = df_er.copy()

    # Función de formato con colores para resaltar totales tipo Excel
    def aplicar_estilo_excel(row):
        styles = [''] * len(row)
        es_total = row.get('es_total', False)
        es_encabezado = row.get('es_encabezado', False)

        if es_total:
            # Fondo azul oscuro / gris con negrita para totales
            styles = [
                'font-weight: bold; background-color: #1e293b; color: #f8fafc; border-top: 1px solid #475569; border-bottom: 2px solid #94a3b8;'
            ] * len(row)
        elif es_encabezado:
            # Fondo suave para títulos de sección
            styles = [
                'font-weight: bold; background-color: #0f172a; color: #38bdf8; text-transform: uppercase;'
            ] * len(row)

        return styles

    # Preparar DataFrame para presentación
    df_clean = df_disp.drop(columns=['es_total', 'es_encabezado'])

    # Aplicar formato de moneda y porcentaje
    styler = (
        df_clean.style.apply(aplicar_estilo_excel, axis=1)
        .format(
            {
                col_monto: lambda x: (
                    f"${x:,.2f}" if pd.notnull(x) and x != '' else ''
                ),
                col_pct: lambda x: (
                    f"{x:.1f}%" if pd.notnull(x) and x != '' else ''
                ),
            }
        )
        .set_properties(
            **{
                'padding': '6px 12px',
                'font-family': 'Consolas, monospace',
                'font-size': '14px',
            }
        )
    )

    st.dataframe(styler, use_container_width=True, hide_index=True)


# --- GENERADOR DE EXCEL OFICIAL (.XLSX) ---
def exportar_excel_formateado(df_er, nombre_hoja):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_clean = df_er.drop(
            columns=['es_total', 'es_encabezado'], errors='ignore'
        )
        df_clean.to_excel(writer, sheet_name=nombre_hoja, index=False)

        workbook = writer.book
        worksheet = writer.sheets[nombre_hoja]

        # Formatos Excel
        hdr_format = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#1E293B',
            'border': 1,
        })
        currency_format = workbook.add_format({'num_format': '$#,##0.00'})
        pct_format = workbook.add_format({'num_format': '0.0%'})

        worksheet.set_column('A:A', 22)
        worksheet.set_column('B:B', 40)
        worksheet.set_column('C:C', 18, currency_format)
        worksheet.set_column('D:D', 12, pct_format)

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
        empresa_seleccionada = st.selectbox("Selecciona Empresa:", lista_empresas)

    if empresa_seleccionada:
        df_balanza = cargar_hoja_balanza(ruta_balanza, empresa_seleccionada)

        tab_balanzas, tab_er_mes, tab_er_acum = st.tabs([
            "📑 Balanzas de Comprobación",
            "📈 Estado de Resultados (MES)",
            "📊 Estado de Resultados (ACUM)",
        ])

        with tab_balanzas:
            st.subheader(
                f"Balanza de Comprobación - {empresa_seleccionada} ({mes_sel}/{anio_sel})"
            )
            st.dataframe(df_balanza, use_container_width=True, hide_index=True)

        with tab_er_mes:
            st.subheader(
                f"Estado de Resultados (DEL MES) - {empresa_seleccionada} ({mes_sel}/{anio_sel})"
            )

            if not plantilla:
                st.error("No se encontró el archivo 'FORMATO EDO RESULTADOS.xlsx' en la raíz.")
            else:
                with st.spinner("Procesando Estado de Resultados Del Mes..."):
                    df_er_mes = generar_reporte_estado_resultados(
                        df_balanza, plantilla, tipo='mes'
                    )

                if not df_er_mes.empty:
                    renderizar_tabla_estilo_excel(
                        df_er_mes, 'DEL MES', '% MES'
                    )

                    excel_mes = exportar_excel_formateado(df_er_mes, "ER_MES")
                    st.download_button(
                        label=f"📥 Descargar ER Mes en Excel ({empresa_seleccionada}_{mes_sel}_{anio_sel})",
                        data=excel_mes,
                        file_name=f"Estado_Resultados_MES_{empresa_seleccionada}_{mes_sel}_{anio_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with tab_er_acum:
            st.subheader(
                f"Estado de Resultados (ACUMULADO) - {empresa_seleccionada} ({mes_sel}/{anio_sel})"
            )

            if not plantilla:
                st.error("No se encontró el archivo 'FORMATO EDO RESULTADOS.xlsx' en la raíz.")
            else:
                with st.spinner("Procesando Estado de Resultados Acumulado..."):
                    df_er_acum = generar_reporte_estado_resultados(
                        df_balanza, plantilla, tipo='acum'
                    )

                if not df_er_acum.empty:
                    renderizar_tabla_estilo_excel(
                        df_er_acum, 'ACUMULADO', '% ACUM'
                    )

                    excel_acum = exportar_excel_formateado(
                        df_er_acum, "ER_ACUM"
                    )
                    st.download_button(
                        label=f"📥 Descargar ER Acumulado en Excel ({empresa_seleccionada}_{mes_sel}_{anio_sel})",
                        data=excel_acum,
                        file_name=f"Estado_Resultados_ACUM_{empresa_seleccionada}_{mes_sel}_{anio_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
else:
    st.error(
        "No se encontraron archivos de balanza en la carpeta 'Balanzas/ AÑO / MES /'."
    )