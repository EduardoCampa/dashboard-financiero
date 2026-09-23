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


# --- OBTENER TODAS LAS BALANZAS DISPONIBLES EN LA CARPETA 'Balanzas/' ---
def obtener_archivos_balanzas():
    # Busca todos los archivos Balanza.xlsx
    archivos = glob.glob("Balanzas/**/Balanza.xlsx", recursive=True)
    if archivos:
        # Ordenar por fecha de modificación (más reciente primero)
        archivos.sort(key=os.path.getmtime, reverse=True)
    return archivos


@st.cache_data(ttl=300)
def obtener_lista_empresas(ruta):
    xls = pd.ExcelFile(ruta)
    # Filtra únicamente hojas que no sean vacías o 'Hoja1' por defecto si hay otras
    sheets = xls.sheet_names
    if len(sheets) > 1 and "Hoja1" in sheets:
        sheets.remove("Hoja1")
    return sheets


@st.cache_data(ttl=300)
def cargar_hoja_balanza(ruta, nombre_hoja):
    return pd.read_excel(ruta, sheet_name=nombre_hoja)


# --- CARGAR PLANTILLA EXCEL ---
@st.cache_data(ttl=3600)
def cargar_plantilla():
    ruta_formato = "FORMATO EDO RESULTADOS.xlsx"
    if not os.path.exists(ruta_formato):
        return []

    wb = openpyxl.load_workbook(ruta_formato, data_only=False)
    sheet = wb.active

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
                'c_formula': str(c_formula) if c_formula else None,
                'd_formula': str(d_formula) if d_formula else None,
            })
    return plantilla


def evaluar_formula_excel(formula, mapa_valores):
    if not formula or not str(formula).startswith('='):
        return 0.0

    f = (
        str(formula)
        .upper()
        .replace(' ', '')
        .replace('+$', '')
        .replace('+', '')
        .replace('$', '')
    )

    m_subtotal = re.match(r'^=SUBTOTAL\(9,C(\d+):C(\d+)\)$', f)
    if m_subtotal:
        r_start, r_end = int(m_subtotal.group(1)), int(m_subtotal.group(2))
        return sum(mapa_valores.get(r, 0.0) for r in range(r_start, r_end + 1))

    def reemplazar_celda(match):
        r_num = int(match.group(1))
        return str(mapa_valores.get(r_num, 0.0))

    expr = re.sub(r'C(\d+)', reemplazar_celda, f).lstrip('=')

    try:
        if re.match(r'^[0-9\.\+\-\*\/\(\)\s]+$', expr):
            return float(eval(expr))
    except Exception:
        pass
    return 0.0


def generar_estado_resultados_completo(df_balanza, plantilla):
    if not plantilla or df_balanza.empty:
        return pd.DataFrame()

    balanza_records = []
    for idx, row in df_balanza.iterrows():
        cta_raw = str(row.iloc[0]).strip()
        if not cta_raw or cta_raw.lower() in ('nan', 'cuenta', 'none'):
            continue

        cargos_m = pd.to_numeric(row.iloc[4], errors='coerce') or 0
        abonos_m = pd.to_numeric(row.iloc[5], errors='coerce') or 0
        cargos_a = pd.to_numeric(row.iloc[6], errors='coerce') or 0
        abonos_a = pd.to_numeric(row.iloc[7], errors='coerce') or 0

        digits_segments = re.findall(r'\d+', cta_raw)
        balanza_records.append({
            'cta_raw': cta_raw,
            'segments': digits_segments,
            'cargos_m': cargos_m,
            'abonos_m': abonos_m,
            'cargos_a': cargos_a,
            'abonos_a': abonos_a,
        })

    def coincide_cuenta(b_rec, patron):
        if not patron or '?' not in patron:
            return False
        p_segs = patron.split('-')
        if len(p_segs) < 3:
            return False

        p_prefix = p_segs[0].strip()
        p_sub = p_segs[2].strip()

        b_segs = b_rec['segments']
        if len(b_segs) >= 3:
            try:
                if b_segs[0] == p_prefix and int(b_segs[2]) == int(p_sub):
                    if len(p_segs) >= 4 and len(b_segs) >= 4:
                        return int(b_segs[3]) == int(p_segs[3])
                    return True
            except ValueError:
                pass

        regex_p = "^" + patron.replace('?', '.').replace('-', r'\\-?') + "$"
        return bool(re.match(regex_p, b_rec['cta_raw']))

    val_mes_map = {}
    val_acum_map = {}

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']

        if patron and '?' in patron:
            m_mes = 0.0
            m_acum = 0.0
            p_prefix = patron.split('-')[0].strip() if '-' in patron else ''

            for b in balanza_records:
                if coincide_cuenta(b, patron):
                    if p_prefix.startswith(('4', '720', '730')):
                        m_mes += b['abonos_m'] - b['cargos_m']
                        m_acum += b['abonos_a'] - b['cargos_a']
                    else:
                        m_mes += b['cargos_m'] - b['abonos_m']
                        m_acum += b['cargos_a'] - b['abonos_a']

            val_mes_map[r_idx] = m_mes
            val_acum_map[r_idx] = m_acum
        else:
            val_mes_map[r_idx] = 0.0
            val_acum_map[r_idx] = 0.0

    for row in plantilla:
        r_idx = row['row_idx']
        c_form = row['c_formula']

        if c_form and str(c_form).startswith('='):
            val_mes_map[r_idx] = evaluar_formula_excel(c_form, val_mes_map)
            val_acum_map[r_idx] = evaluar_formula_excel(c_form, val_acum_map)

    reporte = []
    ventas_totales_mes = val_mes_map.get(91, 1.0) or 1.0
    ventas_totales_acum = val_acum_map.get(91, 1.0) or 1.0

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        concepto = row['concepto']

        v_m = val_mes_map.get(r_idx, 0.0)
        v_a = val_acum_map.get(r_idx, 0.0)

        if not patron and not concepto:
            continue

        pct_mes = (v_m / ventas_totales_mes) * 100 if ventas_totales_mes else 0
        pct_acum = (
            (v_a / ventas_totales_acum) * 100 if ventas_totales_acum else 0
        )

        es_titulo = (
            True if (not patron or patron == 'CUENTA' or 'Total' in concepto) else False
        )

        reporte.append({
            'CUENTA': patron if patron else '',
            'CONCEPTO': concepto,
            'DEL MES': v_m if not (not patron and not c_form_es_valida(row['c_formula'])) else None,
            '% MES': pct_mes if patron or c_form_es_valida(row['c_formula']) else None,
            'ACUMULADO': v_a if not (not patron and not c_form_es_valida(row['c_formula'])) else None,
            '% ACUM': pct_acum if patron or c_form_es_valida(row['c_formula']) else None,
            'es_titulo': es_titulo,
        })

    return pd.DataFrame(reporte)


def c_form_es_valida(c_form):
    return bool(c_form and str(c_form).startswith('='))


# --- INTERFAZ CON SELECTOR DE ARCHIVO Y MES ---
archivos_balanza = obtener_archivos_balanzas()
plantilla = cargar_plantilla()

if archivos_balanza:
    col1, col2 = st.columns([2, 2])
    with col1:
        # Permite seleccionar el archivo o mes si hay varios (ej. Balanzas/2026/08/Balanza.xlsx)
        ruta_balanza = st.selectbox(
            "Selecciona la Balanza a consultar:",
            archivos_balanza,
            index=0,
        )

    lista_empresas = obtener_lista_empresas(ruta_balanza)

    with col2:
        empresa_seleccionada = st.selectbox("Selecciona la Empresa:", lista_empresas)

    if empresa_seleccionada:
        df_balanza = cargar_hoja_balanza(ruta_balanza, empresa_seleccionada)

        tab_balanzas, tab_er = st.tabs(
            ["📑 Balanzas de Comprobación", "📈 Estado de Resultados"]
        )

        with tab_balanzas:
            st.subheader(f"Balanza de Comprobación - {empresa_seleccionada}")
            st.dataframe(df_balanza, use_container_width=True, hide_index=True)

        with tab_er:
            st.subheader(f"Estado de Resultados - {empresa_seleccionada}")

            if not plantilla:
                st.error(
                    "No se encontró el archivo 'FORMATO EDO RESULTADOS.xlsx' en la raíz."
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
                        df_disp.drop(columns=['es_titulo']),
                        use_container_width=True,
                        hide_index=True,
                    )

                    csv_er = df_er.drop(columns=['es_titulo']).to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label=f"📥 Descargar Estado de Resultados ({empresa_seleccionada})",
                        data=csv_er,
                        file_name=f"Estado_Resultados_{empresa_seleccionada}.csv",
                        mime="text/csv",
                    )
else:
    st.error(
        "No se encontró ningún archivo de balanza dentro de la carpeta 'Balanzas/'."
    )