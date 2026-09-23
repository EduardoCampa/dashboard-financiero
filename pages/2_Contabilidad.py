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


# --- ARCHIVOS DE BALANZA ---
def obtener_archivos_balanzas():
    archivos = glob.glob("Balanzas/**/Balanza.xlsx", recursive=True)
    if archivos:
        archivos.sort(key=os.path.getmtime, reverse=True)
    return archivos


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


# --- EVALUADOR MATEMÁTICO DE FÓRMULAS DE EXCEL ---
def evaluar_formula_excel(formula_str, mapa_valores):
    if not formula_str or not str(formula_str).startswith('='):
        return 0.0

    f = (
        str(formula_str)
        .upper()
        .replace(' ', '')
        .replace('+$', '')
        .replace('+', '')
        .replace('$', '')
    )

    # 1. SUBTOTAL(9, Cstart:Cend)
    m_subtotal = re.match(r'^=SUBTOTAL\(9,C(\d+):C(\d+)\)$', f)
    if m_subtotal:
        r_start, r_end = int(m_subtotal.group(1)), int(m_subtotal.group(2))
        return sum(mapa_valores.get(r, 0.0) for r in range(r_start, r_end + 1))

    # 2. Reemplazo de referencias de celdas C<numero>
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


# --- COINCIDENCIA DE CUENTAS POR ESTRUCTURA Y NIVELES ---
def coincide_cuenta(cta_balanza, patron_template):
    if not patron_template or patron_template in ('None', 'CUENTA'):
        return False

    cb = str(cta_balanza).strip()
    pt = str(patron_template).strip()

    # Match 1: Regex con comodines
    regex_str = "^" + pt.replace('?', '.').replace('-', r'\\-?') + "$"
    if re.match(regex_str, cb):
        return True

    # Match 2: Segmentos contables
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
    return False


# --- GENERADOR DEL ESTADO DE RESULTADOS ---
def generar_estado_resultados_completo(df_balanza, plantilla):
    if not plantilla or df_balanza.empty:
        return pd.DataFrame()

    # Extraer balanza con índices de columnas requeridas:
    # Col A (0): Cuenta, Col E (4): Cargos Mes, Col F (5): Abonos Mes
    # Col G (6): Cargos Acum (Deudor F), Col H (7): Abonos Acum (Acreedor F)
    balanza_records = []
    for idx, row in df_balanza.iterrows():
        cta_raw = str(row.iloc[0]).strip()
        if not cta_raw or cta_raw.lower() in ('nan', 'cuenta', 'none'):
            continue

        cargos_m = pd.to_numeric(row.iloc[4], errors='coerce') or 0
        abonos_m = pd.to_numeric(row.iloc[5], errors='coerce') or 0
        cargos_a = pd.to_numeric(row.iloc[6], errors='coerce') or 0
        abonos_a = pd.to_numeric(row.iloc[7], errors='coerce') or 0

        balanza_records.append({
            'cta_raw': cta_raw,
            'cargos_m': cargos_m,
            'abonos_m': abonos_m,
            'cargos_a': cargos_a,
            'abonos_a': abonos_a,
        })

    val_mes_map = {}
    val_acum_map = {}

    # PASO 1: Llenar cuentas directas desde la Balanza
    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']

        if patron and '?' in patron:
            m_mes = 0.0
            m_acum = 0.0
            p_prefix = patron.split('-')[0].strip() if '-' in patron else ''

            for b in balanza_records:
                if coincide_cuenta(b['cta_raw'], patron):
                    # Fórmulas solicitadas:
                    # Ingresos (4xxx, 720, 730): Abonos - Cargos
                    # Costos/Gastos (5xxx, 6xxx, 710, 740): Cargos - Abonos
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

    # PASO 2: Evaluar Fórmulas y Subtotales secuencialmente
    for row in plantilla:
        r_idx = row['row_idx']
        c_form = row['c_formula']

        if c_form and str(c_form).startswith('='):
            val_mes_map[r_idx] = evaluar_formula_excel(c_form, val_mes_map)
            val_acum_map[r_idx] = evaluar_formula_excel(c_form, val_acum_map)

    # PASO 3: Construir reporte final con Porcentajes %
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

        # Si la fila es vacía, omitir
        if not patron and not concepto:
            continue

        pct_mes = (v_m / ventas_totales_mes) * 100 if ventas_totales_mes else 0
        pct_acum = (
            (v_a / ventas_totales_acum) * 100 if ventas_totales_acum else 0
        )

        es_formula = bool(c_form and str(c_form).startswith('='))
        es_cuenta = bool(patron and '?' in patron)

        reporte.append({
            'CUENTA': patron if patron else '',
            'CONCEPTO': concepto,
            'DEL MES': v_m if (es_cuenta or es_formula) else None,
            '% MES': pct_mes if (es_cuenta or es_formula) else None,
            'ACUMULADO': v_a if (es_cuenta or es_formula) else None,
            '% ACUM': pct_acum if (es_cuenta or es_formula) else None,
        })

    return pd.DataFrame(reporte)


# --- INTERFAZ GRAFICA STREAMLIT ---
archivos_balanza = obtener_archivos_balanzas()
plantilla = cargar_plantilla_formato()

if archivos_balanza:
    col1, col2 = st.columns([2, 2])
    with col1:
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
                with st.spinner("Generando Estado de Resultados..."):
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
                        label=f"📥 Descargar Estado de Resultados ({empresa_seleccionada})",
                        data=csv_er,
                        file_name=f"Estado_Resultados_{empresa_seleccionada}.csv",
                        mime="text/csv",
                    )
else:
    st.error(
        "No se encontró ningún archivo de balanza dentro de la carpeta 'Balanzas/'."
    )