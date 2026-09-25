import streamlit as st
import pandas as pd
import io
import os
import glob

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

# --- FUNCIÓN AUXILIAR DE FILTRADO DE REGISTROS ELIMINADOS ---
def filtrar_no_eliminados(df):
    if df is None or df.empty:
        return df
    col_del = next((c for c in ['Deleted', 'Delete', 'deleted', 'delete'] if c in df.columns), None)
    if col_del:
        df = df[pd.to_numeric(df[col_del], errors='coerce').fillna(0) == 0].copy()
    return df

# --- RENDERIZADOR DE TABLA ESTILO ORACLE ENTERPRISE ---
def mostrar_tabla_con_totales_oracle(df_entrada, cols_num):
    if df_entrada.empty:
        st.info("No hay registros que coincidan con los filtros seleccionados.")
        return

    df_calc = df_entrada.copy()
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
            elif any(k in str(col).lower() for k in ['doc', 'folio', 'fecha', 'date', 'tipo', 'moneda', 'curr', 'periodo', 'cuenta', 'id', 'nivel', 'año', 'anio', 'mes']):
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

# --- CARGA INTELIGENTE DE ARCHIVOS Y PESTAÑAS ---
ruta_master = "Consolidado_Master.xlsx" if os.path.exists("Consolidado_Master.xlsx") else "../Consolidado_Master.xlsx"

@st.cache_data
def obtener_hojas_master(path_master):
    if os.path.exists(path_master):
        try:
            xls = pd.ExcelFile(path_master)
            return xls.sheet_names
        except Exception:
            return []
    return []

lista_hojas = obtener_hojas_master(ruta_master)

@st.cache_data
def cargar_tabla_contable(path_master, palabras_clave, patron_archivo):
    df_res = pd.DataFrame()

    # 1. Buscar dentro de Consolidado_Master.xlsx
    if os.path.exists(path_master):
        try:
            xls = pd.ExcelFile(path_master)
            for s in xls.sheet_names:
                s_clean = s.lower().replace(" ", "").replace("_", "").replace("ó", "o").replace("á", "a")
                if any(pk in s_clean for pk in palabras_clave):
                    df_res = pd.read_excel(path_master, sheet_name=s)
                    if not df_res.empty:
                        break
        except Exception:
            pass

    # 2. Si no se encontró en Master, buscar archivos independientes en las carpetas
    if df_res.empty:
        archivos = glob.glob(f"**/{patron_archivo}*.xlsx", recursive=True) + glob.glob(f"../**/{patron_archivo}*.xlsx", recursive=True)
        archivos = list(set([f for f in archivos if "~$" not in f]))
        
        dfs_list = []
        for arch in archivos:
            try:
                xls_ind = pd.ExcelFile(arch)
                for sheet in xls_ind.sheet_names:
                    df_temp = pd.read_excel(arch, sheet_name=sheet)
                    if not df_temp.empty:
                        dfs_list.append(df_temp)
            except Exception:
                continue
        if dfs_list:
            df_res = pd.concat(dfs_list, ignore_index=True)

    return filtrar_no_eliminados(df_res)

# Cargar DataFrames
df_balanza = cargar_tabla_contable(ruta_master, ['balanza'], 'Balanza')
df_er_men = cargar_tabla_contable(ruta_master, ['ermensual', 'er_mensual', 'mensual', 'p&l', 'resultadosmensual'], 'ER_Mensual')
df_er_acu = cargar_tabla_contable(ruta_master, ['eracumulado', 'er_acumulado', 'acumulado', 'resultadosacumulado'], 'ER_Acumulado')

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

    if df_balanza is None or df_balanza.empty:
        st.warning("⚠️ No se encontraron registros automáticos de Balanza.")
        if lista_hojas:
            hoja_manual = st.selectbox("Selecciona manualmente la pestaña de Balanza:", ["-- Seleccionar --"] + lista_hojas, key="sel_manual_bal")
            if hoja_manual != "-- Seleccionar --":
                df_balanza = pd.read_excel(ruta_master, sheet_name=hoja_manual)

    if df_balanza is not None and not df_balanza.empty:
        df_b = df_balanza.copy()

        # Normalizar columnas numéricas
        cols_num_bal = [c for c in df_b.columns if any(k in str(c).lower() for k in ['saldo', 'debe', 'haber', 'cargo', 'abono', 'monto', 'total', 'import', 'inicial', 'final'])]
        for col_m in cols_num_bal:
            df_b[col_m] = pd.to_numeric(df_b[col_m], errors='coerce').fillna(0.0)

        # Detectar columnas para filtros
        col_emp = next((c for c in df_b.columns if any(k in str(c).lower() for k in ['empresa', 'origen'])), None)
        col_anio = next((c for c in df_b.columns if any(k in str(c).lower() for k in ['año', 'anio', 'ejercicio'])), None)
        col_mes = next((c for c in df_b.columns if any(k in str(c).lower() for k in ['mes', 'periodo'])), None)
        col_con = next((c for c in df_b.columns if any(k in str(c).lower() for k in ['cuenta', 'nombre', 'concepto', 'descripcion'])), None)

        st.markdown("#### ⚙️ Filtros de Selección")
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            if col_emp:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_b[col_emp].dropna().astype(str).unique()), key="bal_emp")
                if emp_sel: df_b = df_b[df_b[col_emp].astype(str).isin(emp_sel)]
            else:
                st.info("Sin col. Empresa")

        with f2:
            if col_anio:
                anio_sel = st.multiselect("Año:", sorted(df_b[col_anio].dropna().astype(str).unique(), reverse=True), key="bal_anio")
                if anio_sel: df_b = df_b[df_b[col_anio].astype(str).isin(anio_sel)]
            else:
                st.info("Sin col. Año")

        with f3:
            if col_mes:
                mes_sel = st.multiselect("Mes / Periodo:", sorted(df_b[col_mes].dropna().astype(str).unique()), key="bal_mes")
                if mes_sel: df_b = df_b[df_b[col_mes].astype(str).isin(mes_sel)]
            else:
                st.info("Sin col. Mes")

        with f4:
            if col_con:
                con_fil = st.text_input("Buscar Cuenta/Concepto:", key="bal_con")
                if con_fil: df_b = df_b[df_b[col_con].astype(str).str.contains(con_fil, case=False, na=False)]

        st.markdown("---")

        if cols_num_bal:
            cols_kpi = st.columns(max(1, min(4, len(cols_num_bal))))
            for idx, col_m in enumerate(cols_num_bal[:4]):
                with cols_kpi[idx]:
                    st.metric(col_m, formato_mx(df_b[col_m].sum()))
            st.markdown("---")

        mostrar_tabla_con_totales_oracle(df_b, cols_num_bal)

# ==========================================
# 2. ESTADO DE RESULTADOS MENSUAL
# ==========================================
elif submodulo == "📅 Estado de Resultados Mensual":
    st.title("📅 Estado de Resultados Mensual")

    if df_er_men is None or df_er_men.empty:
        st.warning("⚠️ No se encontraron registros automáticos para Estado de Resultados Mensual.")
        if lista_hojas:
            hoja_manual_m = st.selectbox("Selecciona manualmente la pestaña del ER Mensual:", ["-- Seleccionar --"] + lista_hojas, key="sel_manual_erm")
            if hoja_manual_m != "-- Seleccionar --":
                df_er_men = pd.read_excel(ruta_master, sheet_name=hoja_manual_m)

    if df_er_men is not None and not df_er_men.empty:
        df_m = df_er_men.copy()

        cols_num_m = [c for c in df_m.columns if any(k in str(c).lower() for k in ['monto', 'total', 'import', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre', 'saldo', 'debe', 'haber'])]
        for col_m in cols_num_m:
            df_m[col_m] = pd.to_numeric(df_m[col_m], errors='coerce').fillna(0.0)

        col_emp_m = next((c for c in df_m.columns if any(k in str(c).lower() for k in ['empresa', 'origen'])), None)
        col_anio_m = next((c for c in df_m.columns if any(k in str(c).lower() for k in ['año', 'anio', 'ejercicio'])), None)
        col_mes_m = next((c for c in df_m.columns if any(k in str(c).lower() for k in ['mes', 'periodo'])), None)
        col_con_m = next((c for c in df_m.columns if any(k in str(c).lower() for k in ['cuenta', 'concepto', 'nombre', 'descripcion'])), None)

        st.markdown("#### ⚙️ Filtros de Selección")
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            if col_emp_m:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_m[col_emp_m].dropna().astype(str).unique()), key="er_m_emp")
                if emp_sel: df_m = df_m[df_m[col_emp_m].astype(str).isin(emp_sel)]
            else:
                st.info("Sin col. Empresa")

        with m2:
            if col_anio_m:
                anio_sel = st.multiselect("Año:", sorted(df_m[col_anio_m].dropna().astype(str).unique(), reverse=True), key="er_m_anio")
                if anio_sel: df_m = df_m[df_m[col_anio_m].astype(str).isin(anio_sel)]
            else:
                st.info("Sin col. Año")

        with m3:
            if col_mes_m:
                mes_sel = st.multiselect("Mes / Periodo:", sorted(df_m[col_mes_m].dropna().astype(str).unique()), key="er_m_mes")
                if mes_sel: df_m = df_m[df_m[col_mes_m].astype(str).isin(mes_sel)]
            else:
                st.info("Sin col. Mes")

        with m4:
            if col_con_m:
                con_fil = st.text_input("Buscar Concepto/Cuenta:", key="er_m_con")
                if con_fil: df_m = df_m[df_m[col_con_m].astype(str).str.contains(con_fil, case=False, na=False)]

        st.markdown("---")
        mostrar_tabla_con_totales_oracle(df_m, cols_num_m)

# ==========================================
# 3. ESTADO DE RESULTADOS ACUMULADO
# ==========================================
elif submodulo == "📈 Estado de Resultados Acumulado":
    st.title("📈 Estado de Resultados Acumulado")

    if df_er_acu is None or df_er_acu.empty:
        st.warning("⚠️ No se encontraron registros automáticos para Estado de Resultados Acumulado.")
        if lista_hojas:
            hoja_manual_a = st.selectbox("Selecciona manualmente la pestaña del ER Acumulado:", ["-- Seleccionar --"] + lista_hojas, key="sel_manual_era")
            if hoja_manual_a != "-- Seleccionar --":
                df_er_acu = pd.read_excel(ruta_master, sheet_name=hoja_manual_a)

    if df_er_acu is not None and not df_er_acu.empty:
        df_a = df_er_acu.copy()

        cols_num_a = [c for c in df_a.columns if any(k in str(c).lower() for k in ['monto', 'total', 'import', 'acumulado', 'saldo', 'debe', 'haber'])]
        for col_m in cols_num_a:
            df_a[col_m] = pd.to_numeric(df_a[col_m], errors='coerce').fillna(0.0)

        col_emp_a = next((c for c in df_a.columns if any(k in str(c).lower() for k in ['empresa', 'origen'])), None)
        col_anio_a = next((c for c in df_a.columns if any(k in str(c).lower() for k in ['año', 'anio', 'ejercicio'])), None)
        col_mes_a = next((c for c in df_a.columns if any(k in str(c).lower() for k in ['mes', 'periodo'])), None)
        col_con_a = next((c for c in df_a.columns if any(k in str(c).lower() for k in ['cuenta', 'concepto', 'nombre', 'descripcion'])), None)

        st.markdown("#### ⚙️ Filtros de Selección")
        a1, a2, a3, a4 = st.columns(4)
        with a1:
            if col_emp_a:
                emp_sel = st.multiselect("Empresa Origen:", sorted(df_a[col_emp_a].dropna().astype(str).unique()), key="er_a_emp")
                if emp_sel: df_a = df_a[df_a[col_emp_a].astype(str).isin(emp_sel)]
            else:
                st.info("Sin col. Empresa")

        with a2:
            if col_anio_a:
                anio_sel = st.multiselect("Año:", sorted(df_a[col_anio_a].dropna().astype(str).unique(), reverse=True), key="er_a_anio")
                if anio_sel: df_a = df_a[df_a[col_anio_a].astype(str).isin(anio_sel)]
            else:
                st.info("Sin col. Año")

        with a3:
            if col_mes_a:
                mes_sel = st.multiselect("Mes / Periodo:", sorted(df_a[col_mes_a].dropna().astype(str).unique()), key="er_a_mes")
                if mes_sel: df_a = df_a[df_a[col_mes_a].astype(str).isin(mes_sel)]
            else:
                st.info("Sin col. Mes")

        with a4:
            if col_con_a:
                con_fil = st.text_input("Buscar Concepto/Cuenta:", key="er_a_con")
                if con_fil: df_a = df_a[df_a[col_con_a].astype(str).str.contains(con_fil, case=False, na=False)]

        st.markdown("---")
        mostrar_tabla_con_totales_oracle(df_a, cols_num_a)