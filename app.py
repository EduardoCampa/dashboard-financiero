import streamlit as st
import pandas as pd
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Finanzas - Grupo SERVYRE", layout="wide")

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

def generar_excel_ejecutivo(df_datos, empresa_nombre):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte de Pagos"
    ws.views.sheetView[0].showGridLines = True

    font_titulo = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    fill_titulo = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    font_empresa = Font(name="Calibri", size=12, bold=True, color="1F497D")
    fill_empresa = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    font_prov = Font(name="Calibri", size=11, bold=True, color="333333")
    fill_prov = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    font_moneda = Font(name="Calibri", size=10, bold=True, color="1F497D")
    fill_moneda = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    font_normal = Font(name="Calibri", size=10, color="000000")
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_total = PatternFill(start_color="8EA9DB", end_color="8EA9DB", fill_type="solid")

    borde_delgado = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    borde_total = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    row_idx = 1
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)
    cell = ws.cell(row=row_idx, column=1, value=f"REPORTE EJECUTIVO DE PAGOS — {empresa_nombre.upper()}")
    cell.font = font_titulo; cell.fill = fill_titulo; cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_idx].height = 35
    row_idx += 2

    headers = ["TIPO", "PROVEEDOR", "FECHA VENCIMIENTO", "FOLIO / DOCUMENTO", "UUID", "MONEDA", "DESCRIPCIÓN", "SALDO PENDIENTE"]
    col_emp = 'EmpresaOrigen' if 'EmpresaOrigen' in df_datos.columns else None
    col_prov = 'BusinessEntityName' if 'BusinessEntityName' in df_datos.columns else None
    col_curr = 'Currency' if 'Currency' in df_datos.columns else None

    empresas = sorted(df_datos[col_emp].unique()) if col_emp else ['General']

    for empresa in empresas:
        df_emp = df_datos[df_datos[col_emp] == empresa] if col_emp else df_datos
        total_empresa = df_emp['Saldo_Pendiente'].sum()

        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
        cell_emp = ws.cell(row=row_idx, column=1, value=f"EMPRESA: {str(empresa).upper()}")
        cell_emp.font = font_empresa; cell_emp.fill = fill_empresa; cell_emp.alignment = Alignment(horizontal="left", vertical="center")
        
        cell_emp_tot = ws.cell(row=row_idx, column=8, value=total_empresa)
        cell_emp_tot.font = font_empresa; cell_emp_tot.fill = fill_empresa; cell_emp_tot.number_format = '"$"#,##0.00'; cell_emp_tot.alignment = Alignment(horizontal="right", vertical="center")
        
        for c_idx in range(1, 9): ws.cell(row=row_idx, column=c_idx).border = borde_delgado
        ws.row_dimensions[row_idx].height = 26
        row_idx += 1

        proveedores = sorted(df_emp[col_prov].unique()) if col_prov else []
        for prov in proveedores:
            df_prov = df_emp[df_emp[col_prov] == prov]
            total_prov = df_prov['Saldo_Pendiente'].sum()

            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
            cell_prov = ws.cell(row=row_idx, column=1, value=f"   👤 {str(prov)}")
            cell_prov.font = font_prov; cell_prov.fill = fill_prov; cell_prov.alignment = Alignment(horizontal="left", vertical="center")
            
            cell_prov_tot = ws.cell(row=row_idx, column=8, value=total_prov)
            cell_prov_tot.font = font_prov; cell_prov_tot.fill = fill_prov; cell_prov_tot.number_format = '"$"#,##0.00'; cell_prov_tot.alignment = Alignment(horizontal="right", vertical="center")
            
            for c_idx in range(1, 9): ws.cell(row=row_idx, column=c_idx).border = borde_delgado
            ws.row_dimensions[row_idx].height = 22
            row_idx += 1

            monedas_prov = sorted(df_prov[col_curr].dropna().unique()) if col_curr else ['MXN']
            for moneda in monedas_prov:
                df_moneda = df_prov[df_prov[col_curr] == moneda] if col_curr else df_prov
                total_moneda = df_moneda['Saldo_Pendiente'].sum()

                ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
                cell_mon = ws.cell(row=row_idx, column=1, value=f"      💱 Moneda: {str(moneda)}")
                cell_mon.font = font_moneda; cell_mon.fill = fill_moneda; cell_mon.alignment = Alignment(horizontal="left", vertical="center")
                
                cell_mon_tot = ws.cell(row=row_idx, column=8, value=total_moneda)
                cell_mon_tot.font = font_moneda; cell_mon_tot.fill = fill_moneda; cell_mon_tot.number_format = '"$"#,##0.00'; cell_mon_tot.alignment = Alignment(horizontal="right", vertical="center")
                
                for c_idx in range(1, 9): ws.cell(row=row_idx, column=c_idx).border = borde_delgado
                ws.row_dimensions[row_idx].height = 20
                row_idx += 1

                for col_num, h in enumerate(headers, 1):
                    c = ws.cell(row=row_idx, column=col_num, value=h)
                    c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center"); c.border = borde_delgado
                ws.row_dimensions[row_idx].height = 20
                row_idx += 1

                for _, r in df_moneda.iterrows():
                    ws.cell(row=row_idx, column=1, value=str(r.get('Tipo_Movimiento', ''))).alignment = Alignment(horizontal="center")
                    ws.cell(row=row_idx, column=2, value=str(r.get('BusinessEntityName', ''))).alignment = Alignment(horizontal="left")
                    ws.cell(row=row_idx, column=3, value=str(r.get('Fecha_Fmt', ''))).alignment = Alignment(horizontal="center")
                    ws.cell(row=row_idx, column=4, value=str(r.get('DocFolio', ''))).alignment = Alignment(horizontal="center")
                    ws.cell(row=row_idx, column=5, value=str(r.get('UUID', ''))).alignment = Alignment(horizontal="center")
                    ws.cell(row=row_idx, column=6, value=str(r.get('Currency', 'MXN'))).alignment = Alignment(horizontal="center")
                    ws.cell(row=row_idx, column=7, value=str(r.get('Title', ''))).alignment = Alignment(horizontal="left")
                    
                    c_val = ws.cell(row=row_idx, column=8, value=float(r.get('Saldo_Pendiente', 0)))
                    c_val.number_format = '"$"#,##0.00'; c_val.alignment = Alignment(horizontal="right")

                    for col_num in range(1, 9): ws.cell(row=row_idx, column=col_num).font = font_normal; ws.cell(row=row_idx, column=col_num).border = borde_delgado
                    ws.row_dimensions[row_idx].height = 18
                    row_idx += 1
                row_idx += 1
            row_idx += 1
        row_idx += 1

    ws.cell(row=row_idx, column=1, value="TOTAL GENERAL").font = font_total
    for col_num in range(1, 8): ws.cell(row=row_idx, column=col_num).fill = fill_total; ws.cell(row=row_idx, column=col_num).border = borde_total
    
    cell_tot_gen = ws.cell(row=row_idx, column=8, value=df_datos['Saldo_Pendiente'].sum())
    cell_tot_gen.font = font_total; cell_tot_gen.fill = fill_total; cell_tot_gen.number_format = '"$"#,##0.00'
    cell_tot_gen.border = borde_total; cell_tot_gen.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[row_idx].height = 26

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

ruta_archivo = "Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos_financieros(path):
    if not os.path.exists(path):
        return None, None, None, None, None, None
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names
        df_factura = pd.read_excel(path, sheet_name='FacturaCliente') if 'FacturaCliente' in sheets else pd.DataFrame()
        df_tesoreria = pd.read_excel(path, sheet_name='SolicitudPago') if 'SolicitudPago' in sheets else pd.DataFrame()
        df_ordenes = pd.read_excel(path, sheet_name='OrdenCompra') if 'OrdenCompra' in sheets else pd.DataFrame()
        df_edocuenta = pd.read_excel(path, sheet_name='EdoCuenta') if 'EdoCuenta' in sheets else pd.DataFrame()
        df_fact_compra = pd.read_excel(path, sheet_name='FacturaCompra') if 'FacturaCompra' in sheets else None
        df_gastos = pd.read_excel(path, sheet_name='Gastos') if 'Gastos' in sheets else None
        
        for df_chk in [df_fact_compra, df_gastos, df_ordenes, df_tesoreria]:
            if df_chk is not None and not df_chk.empty:
                col_del = 'Deleted' if 'Deleted' in df_chk.columns else ('Delete' if 'Delete' in df_chk.columns else None)
                if col_del:
                    df_chk.drop(df_chk[pd.to_numeric(df_chk[col_del], errors='coerce').fillna(0) == 1].index, inplace=True)

        if not df_factura.empty:
            col_del_fac = 'Deleted' if 'Deleted' in df_factura.columns else ('Delete' if 'Delete' in df_factura.columns else None)
            if col_del_fac:
                df_factura = df_factura[pd.to_numeric(df_factura[col_del_fac], errors='coerce').fillna(0) == 0]

        return df_factura, df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos
    except Exception as e:
        st.error(f"Error: {e}")
        return None, None, None, None, None, None

df_factura, df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos = cargar_datos_financieros(ruta_archivo)

# --- BARRA LATERAL FINANCIERA CLÁSICA ---
st.sidebar.title("Sistema Auditoría SAT")
st.sidebar.markdown("---")
st.sidebar.markdown("### Módulos Financieros")

menu = st.sidebar.radio(
    "Seleccione módulo:", 
    ["Panel", "Facturación", "OC y SP", "Reporte Pagos"],
    key="menu_finanzas_sel"
)

columnas_oc = ['EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument', 'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'Total', 'UUID', 'Saldo_Pendiente']
columnas_sp = ['EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument', 'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'Total', 'UUID', 'Saldo_Pendiente']
columnas_facturacion = ['EmpresaOrigen', 'BusinessEntityName', 'DocFolio', 'DateDocument', 'Currency', 'SubTotal', 'TotalTax', 'Total', 'TotalPagado', 'SaldoPendiente', 'UUID', 'Status']

if menu == "Panel":
    st.title("Grupo SERVYRE")
    st.markdown("### 📊 Panel Ejecutivo y Consolidado General")
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
                if not df_t.empty: df_t['Año'] = 0; df_t['Periodo'] = "Sin Periodo"
        
        empresas_set = set()
        for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
            if not df_t.empty and 'EmpresaOrigen' in df_t.columns: empresas_set.update(df_t['EmpresaOrigen'].dropna().unique())
        lista_empresas_dash = ["Todas"] + sorted(list(empresas_set))

        col_d0, col_d1, col_d2 = st.columns(3)
        with col_d0: empresa_dash_sel = st.selectbox("Seleccione la Empresa Origen:", lista_empresas_dash, key="dash_empresa_sel")
        
        anios_set = set()
        for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
            if not df_t.empty and 'Año' in df_t.columns: anios_set.update(df_t['Año'].unique())
        anios_dash_validos = ["Todos"] + sorted([int(a) for a in anios_set if a > 0], reverse=True)
        with col_d1: anio_dash_sel = st.selectbox("Seleccione el Año fiscal:", anios_dash_validos, key="dash_anio_sel")
        
        periodos_set = set()
        for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
            if not df_t.empty and 'Periodo' in df_t.columns: periodos_set.update(df_t['Periodo'].unique())
        periodos_dash_validos = ["Todos"] + [p for p in sorted(list(periodos_set)) if p != "Sin Periodo"]
        with col_d2: periodo_dash_sel = st.selectbox("Seleccione el Periodo (Mes):", periodos_dash_validos, key="dash_periodo_sel")

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

elif menu == "Facturación":
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
            if 'Amount' in df_fact_filtrado.columns and 'Amount' != 'TotalPagado': df_fact_filtrado = df_fact_filtrado.drop(columns=['Amount'])
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
                if empresas_fac_sel: df_fact_filtrado = df_fact_filtrado[df_fact_filtrado['EmpresaOrigen'].isin(empresas_fac_sel)]
        with col_f2:
            if 'BusinessEntityName' in df_fact_filtrado.columns:
                lista_cli_fac = sorted(df_fact_filtrado['BusinessEntityName'].dropna().unique())
                cli_fac_sel = st.multiselect("Filtrar por Cliente(s):", lista_cli_fac, default=[], key="fac_cli")
                if cli_fac_sel: df_fact_filtrado = df_fact_filtrado[df_fact_filtrado['BusinessEntityName'].isin(cli_fac_sel)]
        with col_f3:
            anios_fac_disponibles = sorted([int(a) for a in df_fact_filtrado['Año'].unique() if a > 0], reverse=True)
            anios_fac_sel = st.multiselect("Filtrar por Año(s):", anios_fac_disponibles, default=[], key="fac_anio")
            if anios_fac_sel: df_fact_filtrado = df_fact_filtrado[df_fact_filtrado['Año'].isin(anios_fac_sel)]

        st.markdown("---")
        total_fact = df_fact_filtrado['Total'].sum()
        st.metric("Total Facturado (Filtrado)", formato_mx(total_fact))
        
        df_view_fact = df_fact_filtrado.copy()
        for col_m in ['Total', 'TotalTax', 'SubTotal', 'TotalPagado', 'SaldoPendiente']:
            if col_m in df_view_fact.columns: df_view_fact[col_m] = df_view_fact[col_m].apply(formato_mx)
        st.dataframe(df_view_fact[[c for c in columnas_facturacion if c in df_view_fact.columns]], use_container_width=True)

elif menu == "OC y SP":
    st.title("📦 Módulo de Órdenes de Compra y Solicitudes de Pago")
    tab_oc, tab_sp = st.tabs(["OrdenCompra", "SolicitudPago"])
    with tab_oc:
        if df_ordenes is not None and not df_ordenes.empty:
            st.dataframe(df_ordenes, use_container_width=True)
    with tab_sp:
        if df_tesoreria is not None and not df_tesoreria.empty:
            st.dataframe(df_tesoreria, use_container_width=True)

elif menu == "Reporte Pagos":
    st.title("📋 Reporte Ejecutivo de Pagos con Saldo Pendiente")
    st.info("Módulo de Reporte Ejecutivo de Pagos y Descarga en Excel con UUID.")
    # (Aquí puedes mantener la lógica robusta de descarga en Excel del Reporte Pagos)