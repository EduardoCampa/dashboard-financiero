import glob
import os
import re
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo Contable", layout="wide", initial_sidebar_state="expanded"
)

st.title("📊 Módulo Contable")


# --- EXPLORADOR DE ESTRUCTURA: Balanzas / AÑO / MES / Balanza.xlsx ---
def obtener_estructura_balanzas():
    """Busca todos los archivos 'Balanza.xlsx' y extrae Año, Mes y Ruta."""
    archivos = glob.glob("Balanzas/**/Balanza.xlsx", recursive=True)
    estructura = []

    for path in archivos:
        # Normalizar barras de ruta (Windows/Linux)
        path_norm = path.replace("\\", "/")
        partes = path_norm.split("/")

        # Estructura esperada: Balanzas / AÑO / MES / Balanza.xlsx
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
        # Ordenar por fecha de modificación más reciente
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


# --- PLANTILLA EXCEL ---
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


# --- EVALUADOR DE FÓRMULAS EN CASCADA ---
def resolver_todas_las_formulas(mapa_valores, mapa_formulas):
    vals = dict(mapa_valores)

    for _ in range(5):
        for r_idx, f_str in mapa_formulas.items():
            f_clean = f_str.upper().replace(' ', '').replace('$', '')

            # 1. SUBTOTAL(9, Cstart:Cend)
            m_sub = re.match(r'^=SUBTOTAL\(9,C(\d+):C(\d+)\)$', f_clean)
            if m_sub:
                r_s, r_e = int(m_sub.group(1)), int(m_sub.group(2))
                vals[r_idx] = sum(vals.get(r, 0.0) for r in range(r_s, r_e + 1))
                continue

            # 2. Expresiones algebraicas (=+C29+C23+C17+C11)
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


# --- COINCIDENCIA DE CUENTAS ROBUSTA ---
def coincide_cuenta_robusta(cta_balanza, patron_template):
    if not patron_template or patron_template in ('None', 'CUENTA', ''):
        return False

    cb = str(cta_balanza).strip()
    pt = str(patron_template).strip()

    cb_clean = re.sub(r'[^0-9A-Za-z]', '', cb)
    pt_clean = re.sub(r'[^0-9A-Za-z?]', '', pt)

    if '?' not in pt_clean:
        return cb_clean == pt_clean

    p_segs = pt.split('-')
    b_segs = cb.split('-')

    if len(p_segs) >= 3 and len(b_segs) >= 3:
        if p_segs[0] == b_segs[0]:
            try:
                if int(p_segs[2]) == int(b_segs[2]):
                    if len(p_segs) >= 4 and len(b_segs) >= 4:
                        return int(p_segs[3]) == int(b_segs[3])
                    return True
            except ValueError:
                pass

    regex_str = "^" + pt.replace('?', '.').replace('-', r'\\-?') + "$"
    return bool(re.match(regex_str, cb))


# --- GENERADOR DEL ESTADO DE RESULTADOS ---
def generar_estado_resultados_completo(df_balanza, plantilla):
    if not plantilla or df_balanza.empty:
        return pd.DataFrame()

    balanza_records = []
    for idx, row in df_balanza.iterrows():
        cta_raw = str(row.iloc[0]).strip()
        if not cta_raw or cta_raw.lower() in ('nan', 'cuenta', 'none'):
            continue

        cargos_m = pd.to_numeric(row.iloc[4], errors='coerce') or 0.0
        abonos_m = pd.to_numeric(row.iloc[5], errors='coerce') or 0.0
        deudor_f = pd.to_numeric(row.iloc[6], errors='coerce') or 0.0
        acreedor_f = pd.to_numeric(row.iloc[7], errors='coerce') or 0.0

        balanza_records.append({
            'cta_raw': cta_raw,
            'cargos_m': cargos_m,
            'abonos_m': abonos_m,
            'deudor_f': deudor_f,
            'acreedor_f': acreedor_f,
        })

    val_mes_map = {}
    val_acum_map = {}
    formulas_map = {}

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        c_form = row['c_formula']

        if c_form and str(c_form).startswith('='):
            formulas_map[r_idx] = c_form
            val_mes_map[r_idx] = 0.0
            val_acum_map[r_idx] = 0.0
        elif patron:
            m_mes = 0.0
            m_acum = 0.0
            p_prefix = patron.split('-')[0].strip() if '-' in patron else patron[:3]

            for b in balanza_records:
                if coincide_cuenta_robusta(b['cta_raw'], patron):
                    if p_prefix.startswith(('4', '720', '730')):
                        m_mes += b['abonos_m'] - b['cargos_m']
                        m_acum += b['acreedor_f'] - b['deudor_f']
                    else:
                        m_mes += b['cargos_m'] - b['abonos_m']
                        m_acum += b['deudor_f'] - b['acreedor_f']

            val_mes_map[r_idx] = m_mes
            val_acum_map[r_idx] = m_acum
        else:
            val_mes_map[r_idx] = 0.0
            val_acum_map[r_idx] = 0.0

    val_mes_map = resolver_todas_las_formulas(val_mes_map, formulas_map)
    val_acum_map = resolver_todas_las_formulas(val_acum_map, formulas_map)

    reporte = []
    ventas_totales_mes = val_mes_map.get(91, 0.0) or 1.0
    ventas_totales_acum = val_acum_map.get(91, 0.0) or 1.0

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        concepto = row['concepto']
        c_form = row['c_formula']

        v_m = val_mes_map.get(r_idx, 0.0)
        v_a = val_acum_map.get(r_idx, 0.0)

        if not patron and not concepto and not c_form:
            continue

        pct_mes = (v_m / ventas_totales_mes) * 100 if ventas_totales_mes else 0.0
        pct_acum = (
            (v_a / ventas_totales_acum) * 100 if ventas_totales_acum else 0.0
        )

        es_formula = bool(c_form and str(c_form).startswith('='))
        es_cuenta = bool(patron)

        reporte.append({
            'CUENTA': patron if patron else '',
            'CONCEPTO': concepto,
            'DEL MES': v_m if (es_cuenta or es_formula) else None,
            '% MES': pct_mes if (es_cuenta or es_formula) else None,
            'ACUMULADO': v_a if (es_cuenta or es_formula) else None,
            '% ACUM': pct_acum if (es_cuenta or es_formula) else None,
        })

    return pd.DataFrame(reporte)


# --- INTERFAZ CON SELECTORES DE AÑO, MES Y EMPRESA ---
estructura = obtener_estructura_balanzas()
plantilla = cargar_plantilla_formato()

if estructura:
    # 1. Obtener lista de Años únicos
    anios_disponibles = sorted(list(set(x['anio'] for x in estructura)), reverse=True)

    col_a, col_m, col_e = st.columns([1, 1, 2])

    with col_a:
        anio_sel = st.selectbox("Selecciona Año:", anios_disponibles)

    # 2. Filtrar Meses disponibles para el Año seleccionado
    meses_disponibles = sorted(
        list(set(x['mes'] for x in estructura if x['anio'] == anio_sel)),
        reverse=True,
    )

    with col_m:
        mes_sel = st.selectbox("Selecciona Mes:", meses_disponibles)

    # 3. Obtener ruta del archivo coincidente
    ruta_balanza = next(
        (x['ruta'] for x in estructura if x['anio'] == anio_sel and x['mes'] mes_sel),
        estructura[0]['ruta'],
    )

    lista_empresas = obtener_lista_empresas(ruta_balanza)

    with col_e:
        empresa_seleccionada = st.selectbox("Selecciona Empresa:", lista_empresas)

    if empresa_seleccionada:
        df_balanza = cargar_hoja_balanza(ruta_balanza, empresa_seleccionada)

        tab_balanzas, tab_er = st.tabs(
            ["📑 Balanzas de Comprobación", "📈 Estado de Resultados"]
        )

        with tab_balanzas:
            st.subheader(
                f"Balanza de Comprobación - {empresa_seleccionada} ({mes_sel}/{anio_sel})"
            )
            st.dataframe(df_balanza, use_container_width=True, hide_index=True)

        with tab_er:
            st.subheader(
                f"Estado de Resultados - {empresa_seleccionada} ({mes_sel}/{anio_sel})"
            )

            if not plantilla:
                st.error(
                    "No se encontró el archivo 'FORMATO EDO RESULTADOS.xlsx' en la raíz del proyecto."
                )
            else:
                with st.spinner("Procesando Estado de Resultados..."):
                    df_er = generar_estado_resultados_completo(
                        df_balanza, plantilla
                    )

                if not df_er.empty:

                    def fmt_monto(val):
                        if pd.isnull(val) or val == '':
                            return ''
                        return f"${val:,.2f}"

                    def fmt_pct(val):
                        if pd.isnull(val) or val == '':
                            return ''
                        return f"{val:.1f}%"

                    df_disp = df_er.copy()
                    df_disp['DEL MES'] = df_disp['DEL MES'].apply(fmt_monto)
                    df_disp['% MES'] = df_disp['% MES'].apply(fmt_pct)
                    df_disp['ACUMULADO'] = df_disp['ACUMULADO'].apply(fmt_monto)
                    df_disp['% ACUM'] = df_disp['% ACUM'].apply(fmt_pct)

                    st.dataframe(
                        df_disp,
                        use_container_width=True,
                        hide_index=True,
                    )

                    csv_er = df_er.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label=f"📥 Descargar Estado de Resultados ({empresa_seleccionada}_{mes_sel}_{anio_sel})",
                        data=csv_er,
                        file_name=f"Estado_Resultados_{empresa_seleccionada}_{mes_sel}_{anio_sel}.csv",
                        mime="text/csv",
                    )
else:
    st.error(
        "No se encontraron archivos de balanza en la carpeta 'Balanzas/ AÑO / MES /'."
    )