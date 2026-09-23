import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Dashboard Financiero - Grupo SERVYRE", layout="wide")

st.title("🏢 Grupo SERVYRE — Sistema Auditoría SAT")
st.markdown("---")
st.markdown("### 📊 Panel Ejecutivo y Consolidado General")

ruta_archivo = "Consolidado_Master.xlsx"

@st.cache_data
def cargar_panel():
    if not os.path.exists(ruta_archivo):
        return None, None, None
    try:
        df_factura = pd.read_excel(ruta_archivo, sheet_name='FacturaCliente')
        df_tesoreria = pd.read_excel(ruta_archivo, sheet_name='SolicitudPago')
        df_ordenes = pd.read_excel(ruta_archivo, sheet_name='OrdenCompra')
        return df_factura, df_tesoreria, df_ordenes
    except Exception:
        return None, None, None

df_factura, df_tesoreria, df_ordenes = cargar_panel()

if df_factura is not None:
    total_ingresos = pd.to_numeric(df_factura['Total'], errors='coerce').fillna(0).sum()
    total_oc = pd.to_numeric(df_ordenes['Total'], errors='coerce').fillna(0).sum() if df_ordenes is not None else 0
    total_sp = pd.to_numeric(df_tesoreria['Total'], errors='coerce').fillna(0).sum() if df_tesoreria is not None else 0
    utilidad = total_ingresos - (total_oc + total_sp)

    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("INGRESOS", f"${total_ingresos:,.2f}")
    with col2: st.metric("EGRESOS (OC)", f"${total_oc:,.2f}")
    with col3: st.metric("GASTOS (SP)", f"${total_sp:,.2f}")
    with col4: st.metric("UTILIDAD NETA", f"${utilidad:,.2f}")

st.markdown("---")
st.success("👈 **Usa el menú de la izquierda (en la sección Pages)** para navegar entre Facturación, Órdenes de Compra y Reportes de Pagos de manera independiente.")