import streamlit as st
import pandas as pd
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Contabilidad - Grupo SERVYRE", layout="wide")

# ==========================================
# 🎨 ESTILOS CSS PERSONALIZADOS (ESTILO ORACLE CLOUD / ENTERPRISE)
# ==========================================
st.markdown("""
    <style>
        .oracle-table-container {
            width: 100%;
            overflow-x: auto;
            margin-bottom: 20px;
            border: 1px solid #c0c0c0;
            border-radius: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .oracle-table {
            width: 100%;
            border-collapse: collapse;
            font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
            font-size: 13px;
            background-color: #ffffff;
            color: #333333;
        }
        .oracle-table th {
            background-color: #f0f2f5;
            color: #1f497d;
            font-weight: bold;
            text-align: center;
            padding: 8px 10px;
            border: 1px solid #d9d9d9;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .oracle-table td {
            padding: 6px 10px;
            border: 1px solid #e5e5e5;
            vertical-align: middle;
        }
        .oracle-table tr:nth-child(even) {
            background-color: #fcfcfc;
        }
        .oracle-table tr:hover {
            background-color: #f0f4f9;
        }
        .oracle-table tr.total-row {
            background-color: #e6ecf5 !important;
            font-weight: bold;
            color: #000000;
            border-top: 2px solid #1f497d;
            border-bottom: 2px solid #1f497d;
        }
        .oracle-table tr.total-row td {
            border: 1px solid #b0c4de;
        }
        .text-center { text-align: center; }
        .text-right { text-align: right; }
        .text-left { text-align: left; }
    </style>
""", unsafe_allow_html=True)

# --- FORMATO DE MONEDA REGIÓN MÉXICO ($1,234,567.89) ---
def formato_mx(val):
    if pd.isnull(val):
        return "$0.00"
    try:
        num = float(val)
        partes = f"{num:,.2f}".split(".")
        entero_formateado = f"{int(partes[0].replace(',', '')):,}"
        decimales = partes[1] if len(partes) > 1 else "00"
        prefix = "-$" if num < 0 else "$"
        return f"{prefix}{abs(int(partes[0].replace(',', ''))):,}.{decimales}"
    except (ValueError, TypeError):
        return str(val)

# --- FUNCIÓN AUXILIAR DE FILTRADO ESTRICTO DE REGISTROS ELIMINADOS ---
def filtrar_no_eliminados(df):
    if df is None or df.empty:
        return df
    col_del = next((c for c in ['Deleted', 'Delete', 'deleted', 'delete'] if c in df.columns), None)
    if col_del:
        df = df[pd.to_numeric(df[col_del], errors='coerce').fillna(0) == 0].copy()
    return df

# --- RENDERIZADOR DE TABLA ESTILO ORACLE ENTERPRISE ---
def mostrar_tabla_con_totales(df_entrada, cols_num, cols_orden=None):
    if df_entrada.empty:
        st.info("No hay registros para mostrar.")
        return

    df_calc = df_entrada.copy()
    if cols_orden:
        cols_existentes = [c for c in cols_orden if c in df_calc.columns]
    else:
        cols_existentes = list(df_calc.columns)
    
    html = ['<div class="oracle-table-container"><table class="oracle-table"><thead><tr>']
    for col in cols_existentes:
        html.append(f'<th>{col}</th>')
    html.append('</tr></thead><tbody>')

    for _, row in df_calc.iterrows():
        html.append('<tr>')
        for col in cols_existentes:
            val = row.get(col, '')
            if col in cols_num:
                html.append(f'<td class="text-right">{formato_mx(val)}</td>')
            elif any(k in str(col).lower() for k in ['doc', 'folio', 'fecha', 'date', 'tipo', 'moneda', 'curr', 'periodo', 'cuenta', 'id']):
                html.append(f'<td class="text-center">{val if not pd.isnull(val) else ""}</td>')
            else:
                html.append(f'<td class="text-left">{val if not pd.isnull(val) else ""}</td>')
        html.append('</tr>')

    # Fila de Totales
    html.append('<tr class="total-row">')
    for idx, col in enumerate(cols_existentes):
        if col in cols_num:
            tot_val = df_calc[col].sum() if col in df_calc.columns else 0.0
            html.append(f'<td class="text-right">{formato_mx(tot_val)}</td>')
        elif idx == 0:
            html.append('<td class="text-center">TOTALES</td>')
        else:
            html.append('<td></td>')
    html.append('</tr></tbody></table></div>')

    st.markdown("".join(html), unsafe_allow_html=True)

# --- BÚSQUEDA Y CARGA DE DATOS DESDE CONSOLIDADO MASTER ---
ruta_archivo = "Consolidado_Master.xlsx" if os.path.exists("Consolidado_Master.xlsx") else "../Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos_contabilidad(path):
    if not os.path.exists(path):
        return None, None, None, []
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names

        def buscar_hoja(palabras_clave):
            for s in sheets:
                s_clean = s.lower().replace(" ", "").replace("_", "").replace("ó", "o").replace("á", "a")
                if any(pk in s_clean for pk in palabras_clave):
                    return s
            return None

        s_balanza = buscar_hoja(['balanzacomprobacion', 'balanza'])
        s_er_men = buscar_hoja(['estadoresultadosmensual', 'ermensual', 'er_mensual', 'resultado_mensual'])
        s_er_acu = buscar_hoja(['estadoresultadosacumulado', 'eracumulado', 'er_acumulado', 'resultado_acumulado'])

        df_balanza = pd.read_excel(path, sheet_name=s_balanza) if s_balanza else pd.DataFrame()
        df_balanza = filtrar_no_eliminados(df_balanza)

        df_er_men = pd.read_excel(path, sheet_name=s_er_men) if s_er_men else pd.DataFrame()
        df_er_men = filtrar_no_eliminados(df_er_men)

        df_er_acu = pd.read_excel(path, sheet_name=s_er_acu) if s_er_acu else pd.DataFrame()
        df_er_acu = filtrar_no_eliminados(df_er_acu)

        return df_balanza, df_er_men, df_er_acu, sheets
    except Exception as e:
        st.error(f"Error al cargar datos contables: {e}")
        return None, None, None, []

df_balanza, df_er_men, df_er_acu, lista_hojas = cargar_datos_contabilidad(ruta_archivo)

# --- NAVEGACIÓN DEL MÓDULO CONTABLE ---
st.sidebar.title("📑 Módulo de Contabilidad")
st.sidebar.markdown("---")
submodulo = st.sidebar.radio(
    "Seleccione Reporte Contable:",
    ["📊 Balanza de Comprobación", "📅 Estado de Resultados Mensual", "📈 Estado de Resultados Acumulado"],
    key="sub_contabilidad_nav"
)

# ==========================================
# 1. BALANZA DE COMPROBACIÓN
# ==========================================
if submodulo == "📊 Balanza de Comprobación":
    st.title("📊 Balanza de Comprobación Contable")

    if df_balanza is not None and not df_balanza.empty:
        df_b = df_balanza.copy()

        cols_num_bal = [c for c in df_b.columns if any(k in str(c).lower() for k in ['saldo', 'debe', 'haber', 'cargo', 'abono', 'monto', 'total', 'import'])]
        for col_m in cols_num_bal:
            df_b[col_m] = pd.to_numeric(df_b[col_m], errors='coerce').fillna(0.0)

        st.markdown("#### ⚙️ Filtros de Selección")
        f1, f2, f3 = st.columns(3)
        with f1:
            col_emp = next((c for c in df_b.columns if 'empresa' in str(c).lower()), None)
            if col_emp:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_b[col_emp].dropna().unique()), key="bal_emp")
                if emp_sel: df_b = df_b[df_b[col_emp].isin(emp_sel)]
        with f2:
            col_per = next((c for c in df_b.columns if any(k in str(c).lower() for k in ['periodo', 'mes', 'año', 'anio'])), None)
            if col_per:
                per_sel = st.multiselect("Periodo Contable:", sorted(df_b[col_per].astype(str).dropna().unique()), key="bal_per")
                if per_sel: df_b = df_b[df_b[col_per].astype(str).isin(per_sel)]
        with f3:
            col_niv = next((c for c in df_b.columns if 'nivel' in str(c).lower()), None)
            if col_niv:
                niv_sel = st.multiselect("Nivel de Cuenta:", sorted(df_b[col_niv].dropna().unique()), key="bal_niv")
                if niv_sel: df_b = df_b[df_b[col_niv].isin(niv_sel)]

        st.markdown("---")

        if cols_num_bal:
            cols_kpi = st.columns(max(1, min(4, len(cols_num_bal))))
            for idx, col_m in enumerate(cols_num_bal[:4]):
                with cols_kpi[idx]:
                    st.metric(col_m, formato_mx(df_b[col_m].sum()))
            st.markdown("---")

        mostrar_tabla_con_totales(df_b, cols_num_bal)
    else:
        st.warning("⚠️ No se encontraron registros para la **Balanza de Comprobación**. Verifica que la pestaña se llame `BalanzaComprobacion` en el archivo Excel.")

# ==========================================
# 2. ESTADO DE RESULTADOS MENSUAL
# ==========================================
elif submodulo == "📅 Estado de Resultados Mensual":
    st.title("📅 Estado de Resultados Mensual")

    if df_er_men is not None and not df_er_men.empty:
        df_m = df_er_men.copy()

        cols_num_m = [c for c in df_m.columns if any(k in str(c).lower() for k in ['monto', 'total', 'import', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre', 'saldo', 'debe', 'haber'])]
        for col_m in cols_num_m:
            df_m[col_m] = pd.to_numeric(df_m[col_m], errors='coerce').fillna(0.0)

        st.markdown("#### ⚙️ Filtros de Selección")
        m1, m2 = st.columns(2)
        with m1:
            col_emp = next((c for c in df_m.columns if 'empresa' in str(c).lower()), None)
            if col_emp:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_m[col_emp].dropna().unique()), key="er_m_emp")
                if emp_sel: df_m = df_m[df_m[col_emp].isin(emp_sel)]
        with m2:
            col_con = next((c for c in df_m.columns if any(k in str(c).lower() for k in ['cuenta', 'concepto', 'nombre', 'descripcion'])), None)
            if col_con:
                con_fil = st.text_input("Buscar Concepto/Cuenta:", key="er_m_con")
                if con_fil: df_m = df_m[df_m[col_con].astype(str).str.contains(con_fil, case=False, na=False)]

        st.markdown("---")
        mostrar_tabla_con_totales(df_m, cols_num_m)
    else:
        st.warning("⚠️ No se encontraron registros para el **Estado de Resultados Mensual**. Verifica que la pestaña se llame `EstadoResultadosMensual` o similar en el archivo Excel.")

# ==========================================
# 3. ESTADO DE RESULTADOS ACUMULADO
# ==========================================
elif submodulo == "📈 Estado de Resultados Acumulado":
    st.title("📈 Estado de Resultados Acumulado")

    if df_er_acu is not None and not df_er_acu.empty:
        df_a = df_er_acu.copy()

        cols_num_a = [c for c in df_a.columns if any(k in str(c).lower() for k in ['monto', 'total', 'import', 'acumulado', 'saldo', 'debe', 'haber'])]
        for col_m in cols_num_a:
            df_a[col_m] = pd.to_numeric(df_a[col_m], errors='coerce').fillna(0.0)

        st.markdown("#### ⚙️ Filtros de Selección")
        a1, a2 = st.columns(2)
        with a1:
            col_emp = next((c for c in df_a.columns if 'empresa' in str(c).lower()), None)
            if col_emp:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_a[col_emp].dropna().unique()), key="er_a_emp")
                if emp_sel: df_a = df_a[df_a[col_emp].isin(emp_sel)]
        with a2:
            col_con = next((c for c in df_a.columns if any(k in str(c).lower() for k in ['cuenta', 'concepto', 'nombre', 'descripcion'])), None)
            if col_con:
                con_fil = st.text_input("Buscar Concepto/Cuenta:", key="er_a_con")
                if con_fil: df_a = df_a[df_a[col_con].astype(str).str.contains(con_fil, case=False, na=False)]

        st.markdown("---")
        mostrar_tabla_con_totales(df_a, cols_num_a)
    else:
        st.warning("⚠️ No se encontraron registros para el **Estado de Resultados Acumulado**. Verifica que la pestaña se llame `EstadoResultadosAcumulado` en el archivo Excel.")