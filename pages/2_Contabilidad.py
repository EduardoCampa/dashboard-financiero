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
            elif any(k in col.lower() for k in ['doc', 'folio', 'fecha', 'date', 'tipo', 'moneda', 'curr', 'periodo', 'cuenta', 'id']):
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

# --- BÚSQUEDA INTELIGENTE DE PESTAÑAS Y CARGA ---
ruta_archivo = "Consolidado_Master.xlsx" if os.path.exists("Consolidado_Master.xlsx") else "../Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos_contabilidad(path):
    if not os.path.exists(path):
        return None, None, None
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names

        # Función para buscar coincidencia flexible de hoja
        def buscar_hoja(palabras_clave):
            for s in sheets:
                s_clean = s.lower().replace(" ", "").replace("_", "").replace("ó", "o").replace("á", "a")
                if any(pk in s_clean for pk in palabras_clave):
                    return s
            return None

        sheet_balanza = buscar_hoja(['balanza', 'balanzacomprobacion', 'balanzacomp'])
        sheet_polizas = buscar_hoja(['poliza', 'polizas', 'librodiario', 'diario'])
        sheet_cuentas = buscar_hoja(['catalogocuentas', 'catcuenta', 'catcuentas', 'cuentas'])

        df_balanza = pd.read_excel(path, sheet_name=sheet_balanza) if sheet_balanza else pd.DataFrame()
        df_balanza = filtrar_no_eliminados(df_balanza)

        df_polizas = pd.read_excel(path, sheet_name=sheet_polizas) if sheet_polizas else pd.DataFrame()
        df_polizas = filtrar_no_eliminados(df_polizas)

        df_cuentas = pd.read_excel(path, sheet_name=sheet_cuentas) if sheet_cuentas else pd.DataFrame()
        df_cuentas = filtrar_no_eliminados(df_cuentas)

        return df_polizas, df_balanza, df_cuentas
    except Exception as e:
        st.error(f"Error al cargar datos contables: {e}")
        return None, None, None

df_polizas, df_balanza, df_cuentas = cargar_datos_contabilidad(ruta_archivo)

# --- NAVEGACIÓN DEL MÓDULO CONTABLE ---
st.sidebar.title("📑 Módulo de Contabilidad")
st.sidebar.markdown("---")
submodulo = st.sidebar.radio(
    "Seleccione Submódulo:",
    ["📊 Balanza de Comprobación", "📖 Libro Diario / Pólizas", "🗂️ Catálogo de Cuentas"],
    key="sub_contabilidad_nav"
)

# ==========================================
# 1. SUBMÓDULO: BALANZA DE COMPROBACIÓN
# ==========================================
if submodulo == "📊 Balanza de Comprobación":
    st.title("📊 Balanza de Comprobación Contable")
    
    if df_balanza is not None and not df_balanza.empty:
        df_b = df_balanza.copy()

        # Detección flexible de columnas numéricas
        cols_num_bal = [c for c in df_b.columns if any(k in c.lower() for k in ['saldo', 'debe', 'haber', 'cargo', 'abono', 'monto', 'total', 'import'])]
        for col_m in cols_num_bal:
            df_b[col_m] = pd.to_numeric(df_b[col_m], errors='coerce').fillna(0.0)

        st.markdown("#### ⚙️ Filtros de Selección")
        f1, f2, f3 = st.columns(3)
        with f1:
            col_emp = next((c for c in df_b.columns if 'empresa' in c.lower()), None)
            if col_emp:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_b[col_emp].dropna().unique()), key="bal_emp")
                if emp_sel: df_b = df_b[df_b[col_emp].isin(emp_sel)]
        with f2:
            col_per = next((c for c in df_b.columns if any(k in c.lower() for k in ['periodo', 'mes', 'año', 'anio'])), None)
            if col_per:
                per_sel = st.multiselect("Periodo Contable:", sorted(df_b[col_per].astype(str).dropna().unique()), key="bal_per")
                if per_sel: df_b = df_b[df_b[col_per].astype(str).isin(per_sel)]
        with f3:
            col_niv = next((c for c in df_b.columns if 'nivel' in c.lower()), None)
            if col_niv:
                niv_sel = st.multiselect("Nivel de Cuenta:", sorted(df_b[col_niv].dropna().unique()), key="bal_niv")
                if niv_sel: df_b = df_b[df_b[col_niv].isin(niv_sel)]

        st.markdown("---")

        # --- KPIS BALANZA ---
        cols_kpi = st.columns(max(1, min(4, len(cols_num_bal))))
        for idx, col_m in enumerate(cols_num_bal[:4]):
            with cols_kpi[idx]:
                st.metric(col_m, formato_mx(df_b[col_m].sum()))

        st.markdown("---")

        col_curr = next((c for c in df_b.columns if any(k in c.lower() for k in ['moneda', 'currency'])), None)
        if col_curr and df_b[col_curr].nunique() > 1:
            for curr in sorted(df_b[col_curr].dropna().unique()):
                st.markdown(f"### 💱 Balanza de Comprobación — Moneda: **{curr}**")
                df_curr = df_b[df_b[col_curr] == curr]
                mostrar_tabla_con_totales(df_curr, cols_num_bal)
                st.markdown("<br>", unsafe_allow_html=True)
        else:
            mostrar_tabla_con_totales(df_b, cols_num_bal)

    else:
        st.warning("⚠️ No se encontró la hoja de Balanza en `Consolidado_Master.xlsx`. Verifica que la pestaña se llame `BalanzaComprobacion` o `Balanza Comprobación`.")

# ==========================================
# 2. SUBMÓDULO: LIBRO DIARIO / PÓLIZAS
# ==========================================
elif submodulo == "📖 Libro Diario / Pólizas":
    st.title("📖 Libro Diario y Registro de Pólizas")

    if df_polizas is not None and not df_polizas.empty:
        df_p = df_polizas.copy()

        cols_num_pol = [c for c in df_p.columns if any(k in c.lower() for k in ['debe', 'haber', 'cargo', 'abono', 'monto', 'total', 'import'])]
        for col_m in cols_num_pol:
            df_p[col_m] = pd.to_numeric(df_p[col_m], errors='coerce').fillna(0.0)

        st.markdown("#### ⚙️ Filtros de Selección")
        p1, p2, p3, p4 = st.columns(4)
        with p1:
            col_emp = next((c for c in df_p.columns if 'empresa' in c.lower()), None)
            if col_emp:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_p[col_emp].dropna().unique()), key="pol_emp")
                if emp_sel: df_p = df_p[df_p[col_emp].isin(emp_sel)]
        with p2:
            col_tipo = next((c for c in df_p.columns if 'tipo' in c.lower()), None)
            if col_tipo:
                tipo_sel = st.multiselect("Tipo de Póliza:", sorted(df_p[col_tipo].dropna().unique()), key="pol_tipo")
                if tipo_sel: df_p = df_p[df_p[col_tipo].isin(tipo_sel)]
        with p3:
            col_fol = next((c for c in df_p.columns if any(k in c.lower() for k in ['folio', 'poliza', 'doc'])), None)
            if col_fol:
                fol_sel = st.multiselect("Folio Póliza:", sorted(df_p[col_fol].astype(str).dropna().unique()), key="pol_fol")
                if fol_sel: df_p = df_p[df_p[col_fol].astype(str).isin(fol_sel)]
        with p4:
            col_conc = next((c for c in df_p.columns if any(k in c.lower() for k in ['concepto', 'descrip', 'title'])), None)
            if col_conc:
                conc_filter = st.text_input("Buscar en Concepto:", key="pol_conc")
                if conc_filter: df_p = df_p[df_p[col_conc].astype(str).str.contains(conc_filter, case=False, na=False)]

        st.markdown("---")

        cols_kpi = st.columns(max(1, min(4, len(cols_num_pol) + 1)))
        with cols_kpi[0]:
            st.metric("Pólizas Registradas", f"{len(df_p):,}")
        for idx, col_m in enumerate(cols_num_pol[:3], 1):
            with cols_kpi[idx]:
                st.metric(f"Suma {col_m}", formato_mx(df_p[col_m].sum()))

        st.markdown("---")

        col_curr = next((c for c in df_p.columns if any(k in c.lower() for k in ['moneda', 'currency'])), None)
        if col_curr and df_p[col_curr].nunique() > 1:
            for curr in sorted(df_p[col_curr].dropna().unique()):
                st.markdown(f"### 💱 Pólizas — Moneda: **{curr}**")
                df_curr = df_p[df_p[col_curr] == curr]
                mostrar_tabla_con_totales(df_curr, cols_num_pol)
                st.markdown("<br>", unsafe_allow_html=True)
        else:
            mostrar_tabla_con_totales(df_p, cols_num_pol)

    else:
        st.warning("⚠️ No se encontró la hoja de Pólizas/Libro Diario en `Consolidado_Master.xlsx`. Verifica que la pestaña se llame `Polizas` o `Poliza`.")

# ==========================================
# 3. SUBMÓDULO: CATÁLOGO DE CUENTAS
# ==========================================
elif submodulo == "🗂️ Catálogo de Cuentas":
    st.title("🗂️ Catálogo de Cuentas Contables")

    if df_cuentas is not None and not df_cuentas.empty:
        df_c = df_cuentas.copy()

        st.markdown("#### ⚙️ Filtros de Selección")
        c1, c2, c3 = st.columns(3)
        with c1:
            col_emp = next((c for c in df_c.columns if 'empresa' in c.lower()), None)
            if col_emp:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_c[col_emp].dropna().unique()), key="cat_emp")
                if emp_sel: df_c = df_c[df_c[col_emp].isin(emp_sel)]
        with c2:
            col_tipo = next((c for c in df_c.columns if any(k in c.lower() for k in ['tipo', 'naturaleza', 'grupo'])), None)
            if col_tipo:
                tipo_sel = st.multiselect("Tipo / Naturaleza:", sorted(df_c[col_tipo].dropna().unique()), key="cat_tipo")
                if tipo_sel: df_c = df_c[df_c[col_tipo].isin(tipo_sel)]
        with c3:
            col_nom = next((c for c in df_c.columns if any(k in c.lower() for k in ['nombre', 'cuenta', 'descrip'])), None)
            if col_nom:
                nom_filter = st.text_input("Buscar Cuenta:", key="cat_nom")
                if nom_filter: df_c = df_c[df_c[col_nom].astype(str).str.contains(nom_filter, case=False, na=False)]

        st.markdown("---")

        cols_num_cat = [c for c in df_c.columns if any(k in c.lower() for k in ['saldo', 'monto', 'total'])]
        mostrar_tabla_con_totales(df_c, cols_num_cat)

    else:
        st.warning("⚠️ No se encontró la hoja de Catálogo de Cuentas en `Consolidado_Master.xlsx`. Verifica que la pestaña se llame `CatalogoCuentas` o `CatCuentas`.")