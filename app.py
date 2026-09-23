import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Dashboard Financiero - Grupo SERVYRE", layout="wide")

def formato_mx(val):
    if pd.isnull(val):
        return "$0.00"
    try:
        num = float(val)
        partes = f"{num:,.2f}".split(".")
        entero_formateado = f"{int(partes[0].replace(',', '')):,}"
        decimales = partes[1] if len(partes) > 1 else "00"
        return f"${entero_formateado}.{decimales}"
    except (ValueError, TypeError):
        return str(val)

st.title("🏢 Grupo SERVYRE")
st.markdown("### 📊 Panel Ejecutivo y Consolidado General")

ruta_archivo = "Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos_panel(path):
    if not os.path.exists(path):
        return None, None, None
    try:
        df_factura = pd.read_excel(path, sheet_name='FacturaCliente')
        df_tesoreria = pd.read_excel(path, sheet_name='SolicitudPago')
        df_ordenes = pd.read_excel(path, sheet_name='OrdenCompra')
        return df_factura, df_tesoreria, df_ordenes
    except Exception as e:
        st.error(f"Error al cargar el panel: {e}")
        return None, None, None

df_factura, df_tesoreria, df_ordenes = cargar_datos_panel(ruta_archivo)

if df_factura is not None and df_tesoreria is not None:
    df_dash_fact = df_factura.copy()
    df_dash_tes = df_tesoreria.copy()
    df_dash_oc = df_ordenes.copy() if df_ordenes is not None else pd.DataFrame()

    for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
        if not df_t.empty and 'DateDocument' in df_t.columns:
            df_t['DateDocument'] = pd.to_datetime(df_t['DateDocument'], errors='coerce')
            df_t['Año'] = df_t['DateDocument'].dt.year.fillna(0).astype(int)
            df_t['Periodo'] = df_t['DateDocument'].dt.to_period('M').astype(str)
        else:
            if not df_t.empty:
                df_t['Año'] = 0
                df_t['Periodo'] = "Sin Periodo"

    empresas_set = set()
    for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
        if not df_t.empty and 'EmpresaOrigen' in df_t.columns:
            empresas_set.update(df_t['EmpresaOrigen'].dropna().unique())
    lista_empresas_dash = ["Todas"] + sorted(list(empresas_set))

    st.sidebar.title("Sistema Auditoría SAT")
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎛️ Filtros del Panel")

    empresa_dash_sel = st.sidebar.selectbox("Seleccione la Empresa Origen:", lista_empresas_dash, key="dash_empresa_sel")

    anios_set = set()
    for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
        if not df_t.empty and 'Año' in df_t.columns:
            anios_set.update(df_t['Año'].unique())
    anios_dash_validos = ["Todos"] + sorted([int(a) for a in anios_set if a > 0], reverse=True)
    anio_dash_sel = st.sidebar.selectbox("Seleccione el Año fiscal:", anios_dash_validos, key="dash_anio_sel")

    periodos_set = set()
    for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
        if not df_t.empty and 'Periodo' in df_t.columns:
            periodos_set.update(df_t['Periodo'].unique())
    periodos_dash_validos = ["Todos"] + [p for p in sorted(list(periodos_set)) if p != "Sin Periodo"]
    periodo_dash_sel = st.sidebar.selectbox("Seleccione el Periodo (Mes):", periodos_dash_validos, key="dash_periodo_sel")

    if empresa_dash_sel != "Todas":
        if not df_dash_fact.empty and 'EmpresaOrigen' in df_dash_fact.columns: df_dash_fact = df_dash_fact[df_dash_fact['EmpresaOrigen'] == empresa_dash_sel]
        if not df_dash_tes.empty and 'EmpresaOrigen' in df_dash_tes.columns: df_dash_tes = df_dash_tes[df_dash_tes['EmpresaOrigen'] == empresa_dash_sel]
        if not df_dash_oc.empty and 'EmpresaOrigen' in df_dash_oc.columns: df_dash_oc = df_dash_oc[df_dash_oc['EmpresaOrigen'] == empresa_dash_sel]

    if anio_dash_sel != "Todos":
        if not df_dash_fact.empty: df_dash_fact = df_dash_fact[df_dash_fact['Año'] == int(anio_dash_sel)]
        if not df_dash_tes.empty: df_dash_tes = df_dash_tes[df_dash_tes['Año'] == int(anio_dash_sel)]
        if not df_dash_oc.empty: df_dash_oc = df_dash_oc[df_dash_oc['Año'] == int(anio_dash_sel)]

    if periodo_dash_sel != "Todos":
        if not df_dash_fact.empty: df_dash_fact = df_dash_fact[df_dash_fact['Periodo'] == periodo_dash_sel]
        if not df_dash_tes.empty: df_dash_tes = df_dash_tes[df_dash_tes['Periodo'] == periodo_dash_sel]
        if not df_dash_oc.empty: df_dash_oc = df_dash_oc[df_dash_oc['Periodo'] == periodo_dash_sel]

    total_ingresos = pd.to_numeric(df_dash_fact['Total'], errors='coerce').fillna(0).sum() if not df_dash_fact.empty else 0.0
    total_oc = pd.to_numeric(df_dash_oc['Total'], errors='coerce').fillna(0).sum() if not df_dash_oc.empty else 0.0
    total_gastos_sp = pd.to_numeric(df_dash_tes['Total'], errors='coerce').fillna(0).sum() if not df_dash_tes.empty else 0.0
    utilidad_neta = total_ingresos - (total_oc + total_gastos_sp)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("INGRESOS", formato_mx(total_ingresos))
    with col2: st.metric("EGRESOS (OC)", formato_mx(total_oc))
    with col3: st.metric("GASTOS (SP)", formato_mx(total_gastos_sp))
    with col4: st.metric("UTILIDAD NETA", formato_mx(utilidad_neta))

    st.markdown("---")
    st.markdown("### 📈 Gráficas Financieras por Empresa Origen")
    empresas_grafica = sorted(list(set(
        list(df_dash_fact['EmpresaOrigen'].dropna().unique() if not df_dash_fact.empty and 'EmpresaOrigen' in df_dash_fact.columns else []) +
        list(df_dash_oc['EmpresaOrigen'].dropna().unique() if not df_dash_oc.empty and 'EmpresaOrigen' in df_dash_oc.columns else []) +
        list(df_dash_tes['EmpresaOrigen'].dropna().unique() if not df_dash_tes.empty and 'EmpresaOrigen' in df_dash_tes.columns else [])
    )))
    if empresas_grafica:
        data_grafica = []
        for emp in empresas_grafica:
            ing_emp = pd.to_numeric(df_dash_fact[df_dash_fact['EmpresaOrigen'] == emp]['Total'], errors='coerce').fillna(0).sum() if not df_dash_fact.empty and 'EmpresaOrigen' in df_dash_fact.columns else 0
            oc_emp = pd.to_numeric(df_dash_oc[df_dash_oc['EmpresaOrigen'] == emp]['Total'], errors='coerce').fillna(0).sum() if not df_dash_oc.empty and 'EmpresaOrigen' in df_dash_oc.columns else 0
            sp_emp = pd.to_numeric(df_dash_tes[df_dash_tes['EmpresaOrigen'] == emp]['Total'], errors='coerce').fillna(0).sum() if not df_dash_tes.empty and 'EmpresaOrigen' in df_dash_tes.columns else 0
            egr_emp = oc_emp + sp_emp
            ut_emp = ing_emp - egr_emp
            data_grafica.append({'EmpresaOrigen': emp, 'Ingresos': ing_emp, 'Egresos (OC + SP)': egr_emp, 'Utilidad': ut_emp})
        df_graf = pd.DataFrame(data_grafica).set_index('EmpresaOrigen')
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.subheader("Ingresos vs Egresos por Empresa")
            st.bar_chart(df_graf[['Ingresos', 'Egresos (OC + SP)']])
        with col_g2:
            st.subheader("Utilidad Neta por Empresa")
            st.bar_chart(df_graf['Utilidad'])
else:
    st.warning("No se pudieron cargar los datos del panel general.")