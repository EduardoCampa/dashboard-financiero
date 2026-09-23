import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Órdenes y Gastos - Grupo SERVYRE", layout="wide")

ruta_archivo = "../Consolidado_Master.xlsx" if not os.path.exists("Consolidado_Master.xlsx") else "Consolidado_Master.xlsx"

@st.cache_data
def cargar_oc_sp():
    if not os.path.exists(ruta_archivo): return None, None
    try:
        xls = pd.ExcelFile(ruta_archivo)
        sheets = xls.sheet_names
        df_tes = pd.read_excel(ruta_archivo, sheet_name='SolicitudPago') if 'SolicitudPago' in sheets else pd.DataFrame()
        df_ord = pd.read_excel(ruta_archivo, sheet_name='OrdenCompra') if 'OrdenCompra' in sheets else pd.DataFrame()
        return df_tes, df_ord
    except Exception:
        return None, None

df_tesoreria, df_ordenes = cargar_oc_sp()
columnas_oc_sp = ['EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument', 'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'Total', 'UUID', 'Saldo_Pendiente']

st.title("📦 Módulo de Órdenes de Compra y Solicitudes de Pago")

tab_oc, tab_sp = st.tabs(["OrdenCompra", "SolicitudPago"])

with tab_oc:
    if df_ordenes is not None and not df_ordenes.empty:
        df_o = df_ordenes.copy()
        if 'DateDocument' in df_o.columns:
            df_o['DateDocument'] = pd.to_datetime(df_o['DateDocument'], errors='coerce')
            df_o['Año'] = df_o['DateDocument'].dt.year.fillna(0).astype(int)
        else: df_o['Año'] = 0

        st.markdown("#### ⚙️ Filtros Orden de Compra")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            e_sel = st.multiselect("Empresa Origen (OC):", sorted(df_o['EmpresaOrigen'].dropna().unique()) if 'EmpresaOrigen' in df_o.columns else [], key="oc_e")
            if e_sel: df_o = df_o[df_o['EmpresaOrigen'].isin(e_sel)]
        with c2:
            p_sel = st.multiselect("Proveedor (OC):", sorted(df_o['BusinessEntityName'].dropna().unique()) if 'BusinessEntityName' in df_o.columns else [], key="oc_p")
            if p_sel: df_o = df_o[df_o['BusinessEntityName'].isin(p_sel)]
        with c3:
            a_sel = st.multiselect("Año (OC):", sorted([int(a) for a in df_o['Año'].unique() if a > 0], reverse=True), key="oc_a")
            if a_sel: df_o = df_o[df_o['Año'].isin(a_sel)]
        with c4:
            if st.checkbox("Saldo pendiente > $1.00 (OC)", key="oc_s") and 'Saldo_Pendiente' in df_o.columns:
                df_o = df_o[df_o['Saldo_Pendiente'] > 1.0]

        st.markdown("---")
        st.dataframe(df_o[[c for c in columnas_oc_sp if c in df_o.columns]], use_container_width=True)

with tab_sp:
    if df_tesoreria is not None and not df_tesoreria.empty:
        df_s = df_tesoreria.copy()
        if 'DateDocument' in df_s.columns:
            df_s['DateDocument'] = pd.to_datetime(df_s['DateDocument'], errors='coerce')
            df_s['Año'] = df_s['DateDocument'].dt.year.fillna(0).astype(int)
        else: df_s['Año'] = 0

        st.markdown("#### ⚙️ Filtros Solicitudes de Pago")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            e_sel = st.multiselect("Empresa Origen (SP):", sorted(df_s['EmpresaOrigen'].dropna().unique()) if 'EmpresaOrigen' in df_s.columns else [], key="sp_e")
            if e_sel: df_s = df_s[df_s['EmpresaOrigen'].isin(e_sel)]
        with c2:
            p_sel = st.multiselect("Proveedor (SP):", sorted(df_s['BusinessEntityName'].dropna().unique()) if 'BusinessEntityName' in df_s.columns else [], key="sp_p")
            if p_sel: df_s = df_s[df_s['BusinessEntityName'].isin(p_sel)]
        with c3:
            a_sel = st.multiselect("Año (SP):", sorted([int(a) for a in df_s['Año'].unique() if a > 0], reverse=True), key="sp_a")
            if a_sel: df_s = df_s[df_s['Año'].isin(a_sel)]
        with c4:
            if st.checkbox("Saldo pendiente > $1.00 (SP)", key="sp_s") and 'Saldo_Pendiente' in df_s.columns:
                df_s = df_s[df_s['Saldo_Pendiente'] > 1.0]

        st.markdown("---")
        st.dataframe(df_s[[c for c in columnas_oc_sp if c in df_s.columns]], use_container_width=True)