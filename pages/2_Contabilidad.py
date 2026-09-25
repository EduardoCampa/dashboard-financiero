import streamlit as st
import pandas as pd
import os
import glob

st.set_page_config(page_title="Contabilidad - Grupo SERVYRE", layout="wide")

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

# --- FUNCIÓN PARA MOSTRAR TABLA CON FILA DE TOTALES EN STREAMLIT NATIVO ---
def mostrar_tabla_con_totales_nativo(df_entrada, cols_num):
    if df_entrada.empty:
        st.info("No hay registros para mostrar.")
        return

    df_calc = df_entrada.copy()

    # Fila de Totales
    row_total = {}
    for col in df_calc.columns:
        if col in cols_num:
            row_total[col] = df_calc[col].sum()
        elif col in ['EmpresaOrigen', 'Cuenta', 'NombreCuenta', 'Concepto', 'Empresa', 'Descripcion']:
            row_total[col] = "TOTAL"
        else:
            row_total[col] = ""

    df_tot_row = pd.DataFrame([row_total])
    df_con_totales = pd.concat([df_calc, df_tot_row], ignore_index=True)

    # Formato visual
    df_view = df_con_totales.copy()
    for col_m in cols_num:
        if col_m in df_view.columns:
            df_view[col_m] = df_view[col_m].apply(formato_mx)

    st.dataframe(df_view, use_container_width=True)

# ==========================================
# 1. BALANZA DE COMPROBACIÓN
# ==========================================
if submodulo == "📊 Balanza de Comprobación":
    st.title("📊 Balanza de Comprobación Contable")

    if df_balanza is not None and not df_balanza.empty:
        df_b = df_balanza.copy()

        cols_num_bal = [c for c in df_b.columns if any(k in str(c).lower() for k in ['saldo', 'debe', 'haber', 'cargo', 'abono', 'monto', 'total', 'import', 'inicial', 'final'])]
        for col_m in cols_num_bal:
            df_b[col_m] = pd.to_numeric(df_b[col_m], errors='coerce').fillna(0.0)

        st.markdown("#### ⚙️ Filtros de Selección")
        f1, f2, f3 = st.columns(3)
        with f1:
            col_emp = next((c for c in df_b.columns if any(k in str(c).lower() for k in ['empresa', 'origen'])), None)
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

        mostrar_tabla_con_totales_nativo(df_b, cols_num_bal)
    else:
        st.warning("⚠️ No se encontraron registros de Balanza.")
        if lista_hojas:
            hoja_manual = st.selectbox("Selecciona manualmente la pestaña de Balanza:", ["-- Seleccionar --"] + lista_hojas, key="sel_manual_bal")
            if hoja_manual != "-- Seleccionar --":
                df_b_man = pd.read_excel(ruta_master, sheet_name=hoja_manual)
                cols_man = [c for c in df_b_man.columns if any(k in str(c).lower() for k in ['saldo', 'debe', 'haber', 'monto', 'total', 'import'])]
                for c in cols_man: df_b_man[c] = pd.to_numeric(df_b_man[c], errors='coerce').fillna(0.0)
                mostrar_tabla_con_totales_nativo(df_b_man, cols_man)

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
        mostrar_tabla_con_totales_nativo(df_m, cols_num_m)
    else:
        st.warning("⚠️ No se encontraron registros automáticos para Estado de Resultados Mensual.")
        if lista_hojas:
            hoja_manual_m = st.selectbox("Selecciona manualmente la pestaña del ER Mensual:", ["-- Seleccionar --"] + lista_hojas, key="sel_manual_erm")
            if hoja_manual_m != "-- Seleccionar --":
                df_m_man = pd.read_excel(ruta_master, sheet_name=hoja_manual_m)
                cols_man_m = [c for c in df_m_man.columns if any(k in str(c).lower() for k in ['monto', 'total', 'import', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre', 'saldo'])]
                for c in cols_man_m: df_m_man[c] = pd.to_numeric(df_m_man[c], errors='coerce').fillna(0.0)
                mostrar_tabla_con_totales_nativo(df_m_man, cols_man_m)

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
        mostrar_tabla_con_totales_nativo(df_a, cols_num_a)
    else:
        st.warning("⚠️ No se encontraron registros automáticos para Estado de Resultados Acumulado.")
        if lista_hojas:
            hoja_manual_a = st.selectbox("Selecciona manualmente la pestaña del ER Acumulado:", ["-- Seleccionar --"] + lista_hojas, key="sel_manual_era")
            if hoja_manual_a != "-- Seleccionar --":
                df_a_man = pd.read_excel(ruta_master, sheet_name=hoja_manual_a)
                cols_man_a = [c for c in df_a_man.columns if any(k in str(c).lower() for k in ['monto', 'total', 'import', 'acumulado', 'saldo'])]
                for c in cols_man_a: df_a_man[c] = pd.to_numeric(df_a_man[c], errors='coerce').fillna(0.0)
                mostrar_tabla_con_totales_nativo(df_a_man, cols_man_a)