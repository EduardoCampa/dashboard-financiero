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
def mostrar_tabla_con_totales(df_entrada, cols_num, cols_orden):
    if df_entrada.empty:
        st.info("No hay registros para mostrar.")
        return

    df_calc = df_entrada.copy()
    cols_existentes = [c for c in cols_orden if c in df_calc.columns]
    
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
            elif col in ['DocFolio', 'DocumentID', 'DateDocument', 'DateOperation', 'TIPO DOC', 'Currency', 'Moneda', 'Fecha', 'Periodo', 'Cuenta']:
                html.append(f'<td class="text-center">{val if not pd.isnull(val) else ""}</td>')
            else:
                html.append(f'<td class="text-left">{val if not pd.isnull(val) else ""}</td>')
        html.append('</tr>')

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

# --- CARGA DE DATOS CONTABLES DESDE CONSOLIDADO MASTER ---
ruta_archivo = "Consolidado_Master.xlsx" if os.path.exists("Consolidado_Master.xlsx") else "../Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos_contabilidad(path):
    if not os.path.exists(path):
        return None, None, None
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names

        df_polizas = pd.read_excel(path, sheet_name='Polizas') if 'Polizas' in sheets else pd.DataFrame()
        df_polizas = filtrar_no_eliminados(df_polizas)

        df_balanza = pd.read_excel(path, sheet_name='BalanzaComprobacion') if 'BalanzaComprobacion' in sheets else pd.DataFrame()
        df_balanza = filtrar_no_eliminados(df_balanza)

        df_cuentas = pd.read_excel(path, sheet_name='CatalogoCuentas') if 'CatalogoCuentas' in sheets else pd.DataFrame()
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

        # Normalizar columnas
        for col_m in ['SaldoInicial', 'Debe', 'Haber', 'SaldoFinal']:
            if col_m in df_b.columns:
                df_b[col_m] = pd.to_numeric(df_b[col_m], errors='coerce').fillna(0.0)

        st.markdown("#### ⚙️ Filtros de Selección")
        f1, f2, f3 = st.columns(3)
        with f1:
            if 'EmpresaOrigen' in df_b.columns:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_b['EmpresaOrigen'].dropna().unique()), key="bal_emp")
                if emp_sel: df_b = df_b[df_b['EmpresaOrigen'].isin(emp_sel)]
        with f2:
            if 'Periodo' in df_b.columns:
                per_sel = st.multiselect("Periodo Contable:", sorted(df_b['Periodo'].astype(str).dropna().unique()), key="bal_per")
                if per_sel: df_b = df_b[df_b['Periodo'].astype(str).isin(per_sel)]
        with f3:
            if 'Nivel' in df_b.columns:
                niv_sel = st.multiselect("Nivel de Cuenta:", sorted(df_b['Nivel'].dropna().unique()), key="bal_niv")
                if niv_sel: df_b = df_b[df_b['Nivel'].isin(niv_sel)]

        st.markdown("---")

        # --- KPIS BALANZA ---
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.metric("Saldo Inicial Total", formato_mx(df_b['SaldoInicial'].sum() if 'SaldoInicial' in df_b.columns else 0))
        with k2:
            st.metric("Cargos (Debe)", formato_mx(df_b['Debe'].sum() if 'Debe' in df_b.columns else 0))
        with k3:
            st.metric("Abonos (Haber)", formato_mx(df_b['Haber'].sum() if 'Haber' in df_b.columns else 0))
        with k4:
            st.metric("Saldo Final Total", formato_mx(df_b['SaldoFinal'].sum() if 'SaldoFinal' in df_b.columns else 0))

        st.markdown("---")

        cols_num_bal = ['SaldoInicial', 'Debe', 'Haber', 'SaldoFinal']
        cols_orden_bal = ['EmpresaOrigen', 'Cuenta', 'NombreCuenta', 'Nivel', 'SaldoInicial', 'Debe', 'Haber', 'SaldoFinal']

        # Mostrar por Moneda si existe la columna
        if 'Currency' in df_b.columns and df_b['Currency'].nunique() > 1:
            for curr in sorted(df_b['Currency'].dropna().unique()):
                st.markdown(f"### 💱 Balanza de Comprobación — Moneda: **{curr}**")
                df_curr = df_b[df_b['Currency'] == curr]
                mostrar_tabla_con_totales(df_curr, cols_num_bal, cols_orden_bal)
                st.markdown("<br>", unsafe_allow_html=True)
        else:
            mostrar_tabla_con_totales(df_b, cols_num_bal, cols_orden_bal)

    else:
        st.info("No se encontraron registros en la hoja `BalanzaComprobacion`.")

# ==========================================
# 2. SUBMÓDULO: LIBRO DIARIO / PÓLIZAS
# ==========================================
elif submodulo == "📖 Libro Diario / Pólizas":
    st.title("📖 Libro Diario y Registro de Pólizas")

    if df_polizas is not None and not df_polizas.empty:
        df_p = df_polizas.copy()

        for col_m in ['Debe', 'Haber', 'Monto']:
            if col_m in df_p.columns:
                df_p[col_m] = pd.to_numeric(df_p[col_m], errors='coerce').fillna(0.0)

        st.markdown("#### ⚙️ Filtros de Selección")
        p1, p2, p3, p4 = st.columns(4)
        with p1:
            if 'EmpresaOrigen' in df_p.columns:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_p['EmpresaOrigen'].dropna().unique()), key="pol_emp")
                if emp_sel: df_p = df_p[df_p['EmpresaOrigen'].isin(emp_sel)]
        with p2:
            if 'TipoPoliza' in df_p.columns:
                tipo_sel = st.multiselect("Tipo de Póliza:", sorted(df_p['TipoPoliza'].dropna().unique()), key="pol_tipo")
                if tipo_sel: df_p = df_p[df_p['TipoPoliza'].isin(tipo_sel)]
        with p3:
            if 'Folio' in df_p.columns:
                fol_sel = st.multiselect("Folio Póliza:", sorted(df_p['Folio'].astype(str).dropna().unique()), key="pol_fol")
                if fol_sel: df_p = df_p[df_p['Folio'].astype(str).isin(fol_sel)]
        with p4:
            if 'Concepto' in df_p.columns:
                conc_filter = st.text_input("Buscar en Concepto:", key="pol_conc")
                if conc_filter: df_p = df_p[df_p['Concepto'].astype(str).str.contains(conc_filter, case=False, na=False)]

        st.markdown("---")

        # --- KPIS PÓLIZAS ---
        pk1, pk2, pk3 = st.columns(3)
        with pk1:
            st.metric("Total Pólizas Filtradas", f"{len(df_p):,}")
        with pk2:
            st.metric("Suma Cargos (Debe)", formato_mx(df_p['Debe'].sum() if 'Debe' in df_p.columns else 0))
        with pk3:
            st.metric("Suma Abonos (Haber)", formato_mx(df_p['Haber'].sum() if 'Haber' in df_p.columns else 0))

        st.markdown("---")

        cols_num_pol = ['Debe', 'Haber', 'Monto']
        cols_orden_pol = ['EmpresaOrigen', 'Fecha', 'TipoPoliza', 'Folio', 'Cuenta', 'NombreCuenta', 'Concepto', 'Debe', 'Haber', 'UUID']

        # Mostrar por Moneda si existe la columna
        if 'Currency' in df_p.columns and df_p['Currency'].nunique() > 1:
            for curr in sorted(df_p['Currency'].dropna().unique()):
                st.markdown(f"### 💱 Pólizas — Moneda: **{curr}**")
                df_curr = df_p[df_p['Currency'] == curr]
                mostrar_tabla_con_totales(df_curr, cols_num_pol, cols_orden_pol)
                st.markdown("<br>", unsafe_allow_html=True)
        else:
            mostrar_tabla_con_totales(df_p, cols_num_pol, cols_orden_pol)

    else:
        st.info("No se encontraron registros en la hoja `Polizas`.")

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
            if 'EmpresaOrigen' in df_c.columns:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_c['EmpresaOrigen'].dropna().unique()), key="cat_emp")
                if emp_sel: df_c = df_c[df_c['EmpresaOrigen'].isin(emp_sel)]
        with c2:
            if 'Tipo' in df_c.columns:
                tipo_sel = st.multiselect("Naturaleza/Tipo:", sorted(df_c['Tipo'].dropna().unique()), key="cat_tipo")
                if tipo_sel: df_c = df_c[df_c['Tipo'].isin(tipo_sel)]
        with c3:
            if 'Nombre' in df_c.columns or 'NombreCuenta' in df_c.columns:
                col_n = 'NombreCuenta' if 'NombreCuenta' in df_c.columns else 'Nombre'
                nom_filter = st.text_input("Buscar por Nombre de Cuenta:", key="cat_nom")
                if nom_filter: df_c = df_c[df_c[col_n].astype(str).str.contains(nom_filter, case=False, na=False)]

        st.markdown("---")

        cols_orden_cat = ['EmpresaOrigen', 'Cuenta', 'NombreCuenta', 'Naturaleza', 'Nivel', 'Estatus']
        cols_num_cat = []

        mostrar_tabla_con_totales(df_c, cols_num_cat, cols_orden_cat)

    else:
        st.info("No se encontraron registros en la hoja `CatalogoCuentas`.")