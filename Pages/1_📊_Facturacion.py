import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Facturación - Grupo SERVYRE", layout="wide")

def formato_mx(val):
    if pd.isnull(val): return "$0.00"
    try:
        num = float(val)
        partes = f"{num:,.2f}".split(".")
        return f"${int(partes[0].replace(',', '')):,}.{partes[1] if len(partes) > 1 else '00'}"
    except (ValueError, TypeError):
        return str(val)

ruta_archivo = "../Consolidado_Master.xlsx" if not os.path.exists("Consolidado_Master.xlsx") else "Consolidado_Master.xlsx"

@st.cache_data
def cargar_facturacion():
    if not os.path.exists(ruta_archivo): return None, None
    try:
        xls = pd.ExcelFile(ruta_archivo)
        sheets = xls.sheet_names
        df_fact = pd.read_excel(ruta_archivo, sheet_name='FacturaCliente') if 'FacturaCliente' in sheets else pd.DataFrame()
        df_edo = pd.read_excel(ruta_archivo, sheet_name='EdoCuenta') if 'EdoCuenta' in sheets else pd.DataFrame()
        if not df_fact.empty and 'Deleted' in df_fact.columns:
            df_fact = df_fact[pd.to_numeric(df_fact['Deleted'], errors='coerce').fillna(0) == 0]
        return df_fact, df_edo
    except Exception:
        return None, None

df_factura, df_edocuenta = cargar_facturacion()
columnas_facturacion = ['EmpresaOrigen', 'BusinessEntityName', 'DocFolio', 'DateDocument', 'Currency', 'SubTotal', 'TotalTax', 'Total', 'TotalPagado', 'SaldoPendiente', 'UUID', 'Status']

st.title("📊 Módulo de Facturación y Cobranza")

if df_factura is not None and not df_factura.empty:
    df_f = df_factura.copy()
    if df_edocuenta is not None and not df_edocuenta.empty and 'DocumentID' in df_f.columns and 'DocumentID' in df_edocuenta.columns:
        cols_e = ['EmpresaOrigen', 'DocumentID', 'Amount'] if 'EmpresaOrigen' in df_edocuenta.columns else ['DocumentID', 'Amount']
        df_e_sub = df_edocuenta[cols_e].copy()
        df_e_sub['Amount'] = pd.to_numeric(df_e_sub['Amount'], errors='coerce').fillna(0)
        if 'EmpresaOrigen' in df_e_sub.columns and 'EmpresaOrigen' in df_f.columns:
            df_pagos = df_e_sub.groupby(['EmpresaOrigen', 'DocumentID'])['Amount'].sum().reset_index()
            df_f = pd.merge(df_f, df_pagos, on=['EmpresaOrigen', 'DocumentID'], how='left')
        else:
            df_pagos = df_e_sub.groupby('DocumentID')['Amount'].sum().reset_index()
            df_f = pd.merge(df_f, df_pagos, on='DocumentID', how='left')
        df_f['TotalPagado'] = df_f['Amount'].fillna(0)
    else:
        df_f['TotalPagado'] = 0.0

    df_f['Total'] = pd.to_numeric(df_f['Total'], errors='coerce').fillna(0)
    df_f['SaldoPendiente'] = (df_f['Total'] - df_f['TotalPagado']).apply(lambda x: max(0.0, x))

    if 'DateDocument' in df_f.columns:
        df_f['DateDocument'] = pd.to_datetime(df_f['DateDocument'], errors='coerce')
        df_f['Año'] = df_f['DateDocument'].dt.year.fillna(0).astype(int)
    else:
        df_f['Año'] = 0

    st.markdown("#### ⚙️ Filtros de Selección")
    c1, c2, c3 = st.columns(3)
    with c1:
        if 'EmpresaOrigen' in df_f.columns:
            e_sel = st.multiselect("Filtrar por Empresa Origen:", sorted(df_f['EmpresaOrigen'].dropna().unique()), key="f_emp")
            if e_sel: df_f = df_f[df_f['EmpresaOrigen'].isin(e_sel)]
    with c2:
        if 'BusinessEntityName' in df_f.columns:
            c_sel = st.multiselect("Filtrar por Cliente(s):", sorted(df_f['BusinessEntityName'].dropna().unique()), key="f_cli")
            if c_sel: df_f = df_f[df_f['BusinessEntityName'].isin(c_sel)]
    with c3:
        a_sel = st.multiselect("Filtrar por Año(s):", sorted([int(a) for a in df_f['Año'].unique() if a > 0], reverse=True), key="f_anio")
        if a_sel: df_f = df_f[df_f['Año'].isin(a_sel)]

    st.markdown("---")
    st.metric("Total Facturado (Filtrado)", formato_mx(df_f['Total'].sum()))
    
    df_v = df_f.copy()
    for col in ['Total', 'TotalTax', 'SubTotal', 'TotalPagado', 'SaldoPendiente']:
        if col in df_v.columns: df_v[col] = df_v[col].apply(formato_mx)
    st.dataframe(df_v[[c for c in columnas_facturacion if c in df_v.columns]], use_container_width=True)
else:
    st.warning("No hay datos de facturación cargados.")