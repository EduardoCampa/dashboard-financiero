import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Facturación - Grupo SERVYRE", layout="wide")

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

ruta_archivo = "../Consolidado_Master.xlsx" if not os.path.exists("Consolidado_Master.xlsx") else "Consolidado_Master.xlsx"

@st.cache_data
def cargar_facturacion():
    if not os.path.exists(ruta_archivo):
        return None, None
    try:
        xls = pd.ExcelFile(ruta_archivo)
        sheets = xls.sheet_names
        df_factura = pd.read_excel(ruta_archivo, sheet_name='FacturaCliente') if 'FacturaCliente' in sheets else pd.DataFrame()
        df_edocuenta = pd.read_excel(ruta_archivo, sheet_name='EdoCuenta') if 'EdoCuenta' in sheets else pd.DataFrame()
        
        if not df_factura.empty:
            col_del_fac = 'Deleted' if 'Deleted' in df_factura.columns else ('Delete' if 'Delete' in df_factura.columns else None)
            if col_del_fac:
                df_factura = df_factura[pd.to_numeric(df_factura[col_del_fac], errors='coerce').fillna(0) == 0]
                
        return df_factura, df_edocuenta
    except Exception as e:
        st.error(f"Error al cargar datos: {e}")
        return None, None

df_factura, df_edocuenta = cargar_facturacion()

columnas_facturacion = [
    'EmpresaOrigen', 'BusinessEntityName', 'DocFolio', 'DateDocument', 
    'Currency', 'SubTotal', 'TotalTax', 'Total', 'TotalPagado', 'SaldoPendiente', 'UUID', 'Status'
]

st.title("📊 Módulo de Facturación y Dashboard de Cobranza")

if df_factura is not None and not df_factura.empty:
    df_fact_filtrado = df_factura.copy()

    if df_edocuenta is not None and not df_edocuenta.empty and 'DocumentID' in df_fact_filtrado.columns and 'DocumentID' in df_edocuenta.columns:
        cols_edo = ['EmpresaOrigen', 'DocumentID', 'Amount'] if 'EmpresaOrigen' in df_edocuenta.columns else ['DocumentID', 'Amount']
        df_edo_sub = df_edocuenta[cols_edo].copy()
        df_edo_sub['Amount'] = pd.to_numeric(df_edo_sub['Amount'], errors='coerce').fillna(0)
        
        if 'EmpresaOrigen' in df_edo_sub.columns and 'EmpresaOrigen' in df_fact_filtrado.columns:
            df_pagos_agr = df_edo_sub.groupby(['EmpresaOrigen', 'DocumentID'])['Amount'].sum().reset_index()
            df_fact_filtrado = pd.merge(df_fact_filtrado, df_pagos_agr, on=['EmpresaOrigen', 'DocumentID'], how='left')
        else:
            df_pagos_agr = df_edo_sub.groupby('DocumentID')['Amount'].sum().reset_index()
            df_fact_filtrado = pd.merge(df_fact_filtrado, df_pagos_agr, on='DocumentID', how='left')
        
        df_fact_filtrado['TotalPagado'] = df_fact_filtrado['Amount'].fillna(0)
        if 'Amount' in df_fact_filtrado.columns and 'Amount' != 'TotalPagado':
            df_fact_filtrado = df_fact_filtrado.drop(columns=['Amount'])
    else:
        df_fact_filtrado['TotalPagado'] = 0.0

    df_fact_filtrado['Total'] = pd.to_numeric(df_fact_filtrado['Total'], errors='coerce').fillna(0)
    df_fact_filtrado['SaldoPendiente'] = df_fact_filtrado['Total'] - df_fact_filtrado['TotalPagado']
    df_fact_filtrado['SaldoPendiente'] = df_fact_filtrado['SaldoPendiente'].apply(lambda x: max(0.0, x))

    if 'DateDocument' in df_fact_filtrado.columns:
        df_fact_filtrado['DateDocument'] = pd.to_datetime(df_fact_filtrado['DateDocument'], errors='coerce')
        df_fact_filtrado['Año'] = df_fact_filtrado['DateDocument'].dt.year.fillna(0).astype(int)
    else:
        df_fact_filtrado['Año'] = 0

    st.markdown("#### ⚙️ Filtros de Selección")
    col_f1, col_f2, col_f3 = st.columns(3)
    
    with col_f1:
        if 'EmpresaOrigen' in df_fact_filtrado.columns:
            lista_empresas_fac = sorted(df_fact_filtrado['EmpresaOrigen'].dropna().unique())
            empresas_fac_sel = st.multiselect("Filtrar por Empresa Origen:", lista_empresas_fac, default=[], key="fac_emp")
            if empresas_fac_sel:
                df_fact_filtrado = df_fact_filtrado[df_fact_filtrado['EmpresaOrigen'].isin(empresas_fac_sel)]
    
    with col_f2:
        if 'BusinessEntityName' in df_fact_filtrado.columns:
            lista_cli_fac = sorted(df_fact_filtrado['BusinessEntityName'].dropna().unique())
            cli_fac_sel = st.multiselect("Filtrar por Cliente(s):", lista_cli_fac, default=[], key="fac_cli")
            if cli_fac_sel:
                df_fact_filtrado = df_fact_filtrado[df_fact_filtrado['BusinessEntityName'].isin(cli_fac_sel)]

    with col_f3:
        anios_fac_disponibles = sorted([int(a) for a in df_fact_filtrado['Año'].unique() if a > 0], reverse=True)
        anios_fac_sel = st.multiselect("Filtrar por Año(s):", anios_fac_disponibles, default=[], key="fac_anio")
        if anios_fac_sel:
            df_fact_filtrado = df_fact_filtrado[df_fact_filtrado['Año'].isin(anios_fac_sel)]

    st.markdown("---")
    total_fact = df_fact_filtrado['Total'].sum()
    st.metric("Total Facturado (Filtrado)", formato_mx(total_fact))
    
    df_view_fact = df_fact_filtrado.copy()
    if 'Total' in df_view_fact.columns: df_view_fact['Total'] = df_view_fact['Total'].apply(formato_mx)
    if 'TotalTax' in df_view_fact.columns: df_view_fact['TotalTax'] = df_view_fact['TotalTax'].apply(formato_mx)
    if 'SubTotal' in df_view_fact.columns: df_view_fact['SubTotal'] = df_view_fact['SubTotal'].apply(formato_mx)
    if 'TotalPagado' in df_view_fact.columns: df_view_fact['TotalPagado'] = df_view_fact['TotalPagado'].apply(formato_mx)
    if 'SaldoPendiente' in df_view_fact.columns: df_view_fact['SaldoPendiente'] = df_view_fact['SaldoPendiente'].apply(formato_mx)

    cols_mostrar_fact = [c for c in columnas_facturacion if c in df_view_fact.columns]
    st.dataframe(df_view_fact[cols_mostrar_fact], use_container_width=True)
else:
    st.warning("No hay datos de facturación disponibles.")