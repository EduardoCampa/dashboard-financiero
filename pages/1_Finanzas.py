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
def cargar_datos_facturacion(path):
    if not os.path.exists(path):
        return None, None
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names

        # 1. Cargar FacturaCliente
        df_fac = pd.read_excel(path, sheet_name='FacturaCliente') if 'FacturaCliente' in sheets else pd.DataFrame()
        if not df_fac.empty:
            col_del_fac = 'Deleted' if 'Deleted' in df_fac.columns else ('Delete' if 'Delete' in df_fac.columns else None)
            if col_del_fac:
                df_fac = df_fac[pd.to_numeric(df_fac[col_del_fac], errors='coerce').fillna(0) == 0].copy()

            if 'ModuleName' in df_fac.columns:
                df_fac['TIPO DOC'] = df_fac['ModuleName'].astype(str).apply(
                    lambda x: 'NC' if 'nota' in x.lower() or 'credito' in x.lower() or 'crédito' in x.lower() else 'F'
                )
            else:
                df_fac['TIPO DOC'] = 'F'

        # 2. Cargar NotaCreditoCliente (Si existe como pestaña independiente)
        df_nc = pd.read_excel(path, sheet_name='NotaCreditoCliente') if 'NotaCreditoCliente' in sheets else pd.DataFrame()
        if not df_nc.empty:
            col_del_nc = 'Deleted' if 'Deleted' in df_nc.columns else ('Delete' if 'Delete' in df_nc.columns else None)
            if col_del_nc:
                df_nc = df_nc[pd.to_numeric(df_nc[col_del_nc], errors='coerce').fillna(0) == 0].copy()
            
            # Asignar explícitamente NC a los registros de esta pestaña
            df_nc['TIPO DOC'] = 'NC'

        # Consolidar ambas fuentes
        if not df_fac.empty and not df_nc.empty:
            df_facturacion_base = pd.concat([df_fac, df_nc], ignore_index=True)
        elif not df_fac.empty:
            df_facturacion_base = df_fac
        elif not df_nc.empty:
            df_facturacion_base = df_nc
        else:
            df_facturacion_base = pd.DataFrame()

        # 3. Cargar EdoCuenta
        df_edo = pd.read_excel(path, sheet_name='EdoCuenta') if 'EdoCuenta' in sheets else pd.DataFrame()
        if not df_edo.empty and 'Amount' in df_edo.columns:
            df_edo['Amount'] = pd.to_numeric(df_edo['Amount'], errors='coerce').fillna(0)
            df_edo = df_edo[df_edo['Amount'] != 0].copy()

        return df_facturacion_base, df_edo
    except Exception as e:
        st.error(f"Error al cargar datos de facturación: {e}")
        return None, None

df_facturacion_raw, df_edocuenta_raw = cargar_datos_facturacion(ruta_archivo)

columnas_requeridas = [
    'DateDocument', 'EmpresaOrigen', 'DocFolio', 'UUID', 'TIPO DOC',
    'CFDStatusCancelledName', 'BusinessEntityName', 'TotalRetention',
    'SubTotal', 'TotalDiscount', 'Subtotal2', 'TotalTax', 'Total',
    'StatusComplemento', 'CostCenterName', 'Amount', 'DateOperation', 'SaldoFactura'
]

st.title("📊 Módulo de Facturación y Cobranza")

if df_facturacion_raw is not None and not df_facturacion_raw.empty:
    df_f = df_facturacion_raw.copy()

    # Cálculo de Subtotal2 (SubTotal - TotalDiscount)
    subtotal_val = pd.to_numeric(df_f['SubTotal'], errors='coerce').fillna(0) if 'SubTotal' in df_f.columns else 0.0
    discount_val = pd.to_numeric(df_f['TotalDiscount'], errors='coerce').fillna(0) if 'TotalDiscount' in df_f.columns else 0.0
    df_f['Subtotal2'] = subtotal_val - discount_val

    # Formatear DateDocument sin hora
    if 'DateDocument' in df_f.columns:
        df_f['DateDocument_Fmt'] = pd.to_datetime(df_f['DateDocument'], errors='coerce').dt.strftime('%Y-%m-%d')
        df_f['Año'] = pd.to_datetime(df_f['DateDocument'], errors='coerce').dt.year.fillna(0).astype(int)
    else:
        df_f['DateDocument_Fmt'] = ""
        df_f['Año'] = 0

    # Cruce con EdoCuenta por DocumentID
    if df_edocuenta_raw is not None and not df_edocuenta_raw.empty and 'DocumentID' in df_f.columns and 'DocumentID' in df_edocuenta_raw.columns:
        cols_edo = ['DocumentID', 'Amount', 'DateOperation']
        if 'EmpresaOrigen' in df_edocuenta_raw.columns and 'EmpresaOrigen' in df_f.columns:
            cols_edo.append('EmpresaOrigen')

        df_edo_sub = df_edocuenta_raw[[c for c in cols_edo if c in df_edocuenta_raw.columns]].copy()
        
        if 'DateOperation' in df_edo_sub.columns:
            df_edo_sub['DateOperation'] = pd.to_datetime(df_edo_sub['DateOperation'], errors='coerce').dt.strftime('%Y-%m-%d')
        else:
            df_edo_sub['DateOperation'] = ""

        # Suma acumulada de cobros para calcular SaldoFactura
        group_keys = ['EmpresaOrigen', 'DocumentID'] if 'EmpresaOrigen' in df_edo_sub.columns and 'EmpresaOrigen' in df_f.columns else 'DocumentID'
        df_tot_pagado = df_edo_sub.groupby(group_keys)['Amount'].sum().reset_index().rename(columns={'Amount': 'Total_Pagos_Acumulados'})

        df_f = pd.merge(df_f, df_tot_pagado, on=group_keys, how='left')
        df_f['Total_Pagos_Acumulados'] = df_f['Total_Pagos_Acumulados'].fillna(0.0)

        # Merge para desglosar cada pago individual en un renglón
        df_f = pd.merge(df_f, df_edo_sub, on=group_keys, how='left')
        df_f['Amount'] = df_f['Amount'].fillna(0.0)
        df_f['DateOperation'] = df_f['DateOperation'].fillna("")
    else:
        df_f['Amount'] = 0.0
        df_f['Total_Pagos_Acumulados'] = 0.0
        df_f['DateOperation'] = ""

    # SaldoFactura = Total - Total_Pagos_Acumulados
    total_fac_val = pd.to_numeric(df_f['Total'], errors='coerce').fillna(0) if 'Total' in df_f.columns else 0.0
    df_f['SaldoFactura'] = (total_fac_val - df_f['Total_Pagos_Acumulados']).apply(lambda x: max(0.0, x))

    df_f['DateDocument'] = df_f['DateDocument_Fmt']

    st.markdown("#### ⚙️ Filtros de Selección")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        if 'EmpresaOrigen' in df_f.columns:
            e_sel = st.multiselect("Empresa Origen:", sorted(df_f['EmpresaOrigen'].dropna().unique()), key="fac_emp")
            if e_sel: df_f = df_f[df_f['EmpresaOrigen'].isin(e_sel)]
    with c2:
        if 'BusinessEntityName' in df_f.columns:
            cli_sel = st.multiselect("Cliente:", sorted(df_f['BusinessEntityName'].dropna().unique()), key="fac_cli")
            if cli_sel: df_f = df_f[df_f['BusinessEntityName'].isin(cli_sel)]
    with c3:
        if 'TIPO DOC' in df_f.columns:
            tipo_sel = st.multiselect("Tipo Doc (F/NC):", sorted(df_f['TIPO DOC'].dropna().unique()), key="fac_tipo")
            if tipo_sel: df_f = df_f[df_f['TIPO DOC'].isin(tipo_sel)]
    with c4:
        a_sel = st.multiselect("Año:", sorted([int(a) for a in df_f['Año'].unique() if a > 0], reverse=True), key="fac_anio")
        if a_sel: df_f = df_f[df_f['Año'].isin(a_sel)]
    with c5:
        st.markdown("<br>", unsafe_allow_html=True)
        solo_saldo_fac = st.checkbox("Saldo Factura > $1.00", value=False, key="fac_saldo_chk")
        if solo_saldo_fac:
            df_f = df_f[df_f['SaldoFactura'] > 1.0]

    st.markdown("---")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        df_unicos = df_f.drop_duplicates(subset=['EmpresaOrigen', 'DocumentID'] if 'EmpresaOrigen' in df_f.columns and 'DocumentID' in df_f.columns else ['DocFolio'])
        st.metric("Total Subtotal2", formato_mx(df_unicos['Subtotal2'].sum()))
    with col_m2:
        st.metric("Total Facturado", formato_mx(df_unicos['Total'].sum()))
    with col_m3:
        st.metric("Saldo Factura Pendiente", formato_mx(df_unicos['SaldoFactura'].sum()))

    # Formatear importes para despliegue en tabla
    df_view = df_f.copy()
    for col_m in ['TotalRetention', 'SubTotal', 'TotalDiscount', 'Subtotal2', 'TotalTax', 'Total', 'Amount', 'SaldoFactura']:
        if col_m in df_view.columns:
            df_view[col_m] = df_view[col_m].apply(formato_mx)

    cols_disponibles = [c for c in columnas_requeridas if c in df_view.columns]
    st.dataframe(df_view[cols_disponibles], use_container_width=True)
else:
    st.warning("No hay registros disponibles en FacturaCliente ni NotaCreditoCliente.")