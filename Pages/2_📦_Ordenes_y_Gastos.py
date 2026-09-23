import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Órdenes y Gastos - Grupo SERVYRE", layout="wide")

ruta_archivo = "../Consolidado_Master.xlsx" if not os.path.exists("Consolidado_Master.xlsx") else "Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos_oc_sp():
    if not os.path.exists(ruta_archivo):
        return None, None, None, None, None
    try:
        xls = pd.ExcelFile(ruta_archivo)
        sheets = xls.sheet_names
        df_tesoreria = pd.read_excel(ruta_archivo, sheet_name='SolicitudPago') if 'SolicitudPago' in sheets else pd.DataFrame()
        df_ordenes = pd.read_excel(ruta_archivo, sheet_name='OrdenCompra') if 'OrdenCompra' in sheets else pd.DataFrame()
        df_edocuenta = pd.read_excel(ruta_archivo, sheet_name='EdoCuenta') if 'EdoCuenta' in sheets else pd.DataFrame()
        df_fact_compra = pd.read_excel(ruta_archivo, sheet_name='FacturaCompra') if 'FacturaCompra' in sheets else None
        df_gastos = pd.read_excel(ruta_archivo, sheet_name='Gastos') if 'Gastos' in sheets else None
        
        for df_chk in [df_fact_compra, df_gastos, df_ordenes, df_tesoreria]:
            if df_chk is not None and not df_chk.empty:
                col_del = 'Deleted' if 'Deleted' in df_chk.columns else ('Delete' if 'Delete' in df_chk.columns else None)
                if col_del:
                    df_chk.drop(df_chk[pd.to_numeric(df_chk[col_del], errors='coerce').fillna(0) == 1].index, inplace=True)
        return df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos
    except Exception as e:
        st.error(f"Error: {e}")
        return None, None, None, None, None

df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos = cargar_datos_oc_sp()

# Columnas originales exactas para OC y SP
columnas_oc = ['EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument', 'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'Total', 'UUID', 'Saldo_Pendiente']
columnas_sp = ['EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument', 'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'Total', 'UUID', 'Saldo_Pendiente']

st.title("📦 Módulo de Órdenes de Compra y Solicitudes de Pago")

tab_oc, tab_sp = st.tabs(["OrdenCompra", "SolicitudPago"])

with tab_oc:
    if df_ordenes is not None and not df_ordenes.empty:
        df_oc_f = df_ordenes.copy()
        if 'DateDocument' in df_oc_f.columns:
            df_oc_f['DateDocument'] = pd.to_datetime(df_oc_f['DateDocument'], errors='coerce')
            df_oc_f['Año'] = df_oc_f['DateDocument'].dt.year.fillna(0).astype(int)
        else: df_oc_f['Año'] = 0

        st.markdown("#### ⚙️ Filtros Orden de Compra")
        oc_c1, oc_c2, oc_c3, oc_c4 = st.columns(4)
        with oc_c1:
            l_emp_oc = sorted(df_oc_f['EmpresaOrigen'].dropna().unique()) if 'EmpresaOrigen' in df_oc_f.columns else []
            e_oc_sel = st.multiselect("Empresa Origen (OC):", l_emp_oc, default=[], key="oc_emp_f")
            if e_oc_sel: df_oc_f = df_oc_f[df_oc_f['EmpresaOrigen'].isin(e_oc_sel)]
        with oc_c2:
            l_prov_oc = sorted(df_oc_f['BusinessEntityName'].dropna().unique()) if 'BusinessEntityName' in df_oc_f.columns else []
            p_oc_sel = st.multiselect("Proveedor (OC):", l_prov_oc, default=[], key="oc_prov_f")
            if p_oc_sel: df_oc_f = df_oc_f[df_oc_f['BusinessEntityName'].isin(p_oc_sel)]
        with oc_c3:
            l_anio_oc = sorted([int(a) for a in df_oc_f['Año'].unique() if a > 0], reverse=True)
            a_oc_sel = st.multiselect("Año (OC):", l_anio_oc, default=[], key="oc_anio_f")
            if a_oc_sel: df_oc_f = df_oc_f[df_oc_f['Año'].isin(a_oc_sel)]
        with oc_c4:
            solo_saldo_oc = st.checkbox("Saldo pendiente > $1.00 (OC)", value=False, key="oc_saldo_chk")
            if solo_saldo_oc and 'Saldo_Pendiente' in df_oc_f.columns: df_oc_f = df_oc_f[df_oc_f['Saldo_Pendiente'] > 1.0]

        st.markdown("---")
        st.dataframe(df_oc_f[[c for c in columnas_oc if c in df_oc_f.columns]], use_container_width=True)

with tab_sp:
    if df_tesoreria is not None and not df_tesoreria.empty:
        df_sp_f = df_tesoreria.copy()
        if 'DateDocument' in df_sp_f.columns:
            df_sp_f['DateDocument'] = pd.to_datetime(df_sp_f['DateDocument'], errors='coerce')
            df_sp_f['Año'] = df_sp_f['DateDocument'].dt.year.fillna(0).astype(int)
        else: df_sp_f['Año'] = 0

        st.markdown("#### ⚙️ Filtros Solicitudes de Pago")
        sp_c1, sp_c2, sp_c3, sp_c4 = st.columns(4)
        with sp_c1:
            l_emp_sp = sorted(df_sp_f['EmpresaOrigen'].dropna().unique()) if 'EmpresaOrigen' in df_sp_f.columns else []
            e_sp_sel = st.multiselect("Empresa Origen (SP):", l_emp_sp, default=[], key="sp_emp_f")
            if e_sp_sel: df_sp_f = df_sp_f[df_sp_f['EmpresaOrigen'].isin(e_sp_sel)]
        with sp_c2:
            l_prov_sp = sorted(df_sp_f['BusinessEntityName'].dropna().unique()) if 'BusinessEntityName' in df_sp_f.columns else []
            p_sp_sel = st.multiselect("Proveedor (SP):", l_prov_sp, default=[], key="sp_prov_f")
            if p_sp_sel: df_sp_f = df_sp_f[df_sp_f['BusinessEntityName'].isin(p_sp_sel)]
        with sp_c3:
            l_anio_sp = sorted([int(a) for a in df_sp_f['Año'].unique() if a > 0], reverse=True)
            a_sp_sel = st.multiselect("Año (SP):", l_anio_sp, default=[], key="sp_anio_f")
            if a_sp_sel: df_sp_f = df_sp_f[df_sp_f['Año'].isin(a_sp_sel)]
        with sp_c4:
            solo_saldo_sp = st.checkbox("Saldo pendiente > $1.00 (SP)", value=False, key="sp_saldo_chk")
            if solo_saldo_sp and 'Saldo_Pendiente' in df_sp_f.columns: df_sp_f = df_sp_f[df_sp_f['Saldo_Pendiente'] > 1.0]

        st.markdown("---")
        st.dataframe(df_sp_f[[c for c in columnas_sp if c in df_sp_f.columns]], use_container_width=True)