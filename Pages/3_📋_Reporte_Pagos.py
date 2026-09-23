import streamlit as st
import pandas as pd
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Reporte Ejecutivo de Pagos - Grupo SERVYRE", layout="wide")

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

ruta_archivo = "../Consolidado_Master.xlsx" if not os.path.exists("Consolidado_Master.xlsx") else "Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos_reporte():
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
        return df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos
    except Exception as e:
        return None, None, None, None, None

df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos = cargar_datos_reporte()

st.title("📋 Reporte Ejecutivo de Pagos con Saldo Pendiente")

# Procesamiento OC
df_oc_base = pd.DataFrame()
if df_ordenes is not None and not df_ordenes.empty:
    df_oc_base = df_ordenes.copy()
    if df_fact_compra is not None:
        col_uuid_fc = 'UUID' if 'UUID' in df_fact_compra.columns else ('CFDIFolioFiscal' if 'CFDIFolioFiscal' in df_fact_compra.columns else None)
        col_total_fc = 'Total' if 'Total' in df_fact_compra.columns else None
        col_docid_fc = 'DocumentID' if 'DocumentID' in df_fact_compra.columns else None
        col_llave_fc = next((c for c in ['Solicitud de Pago', 'SolicitudPago', 'DocFolio'] if c in df_fact_compra.columns), None)
        if col_llave_fc and col_uuid_fc:
            cols_fc = [col_llave_fc, col_uuid_fc]
            if 'EmpresaOrigen' in df_fact_compra.columns: cols_fc.append('EmpresaOrigen')
            if col_total_fc: cols_fc.append(col_total_fc)
            if col_docid_fc: cols_fc.append(col_docid_fc)
            df_fc_sub = df_fact_compra[cols_fc].dropna(subset=[col_llave_fc]).copy()
            rename_dict = {col_llave_fc: 'DocFolio_Match', col_uuid_fc: 'UUID'}
            if col_total_fc: rename_dict[col_total_fc] = 'TotalFacturaCompra'
            if col_docid_fc: rename_dict[col_docid_fc] = 'DocumentID_FC'
            df_fc_sub = df_fc_sub.rename(columns=rename_dict)
            if 'DocumentID' in df_oc_base.columns: df_oc_base = df_oc_base.drop(columns=['DocumentID'])
            if 'EmpresaOrigen' in df_fc_sub.columns and 'EmpresaOrigen' in df_oc_base.columns:
                df_oc_base = pd.merge(df_oc_base, df_fc_sub, left_on=['EmpresaOrigen', 'DocFolio'], right_on=['EmpresaOrigen', 'DocFolio_Match'], how='left')
            else:
                df_oc_base = pd.merge(df_oc_base, df_fc_sub, left_on='DocFolio', right_on='DocFolio_Match', how='left')
            if 'DocumentID_FC' in df_oc_base.columns: df_oc_base['DocumentID'] = df_oc_base['DocumentID_FC']
    if 'UUID' not in df_oc_base.columns: df_oc_base['UUID'] = ""
    if 'TotalFacturaCompra' not in df_oc_base.columns: df_oc_base['TotalFacturaCompra'] = 0.0
    if 'DocumentID' not in df_oc_base.columns: df_oc_base['DocumentID'] = ""
    if df_edocuenta is not None and 'DocumentID' in df_oc_base.columns and 'DocumentID' in df_edocuenta.columns:
        cols_edo = ['EmpresaOrigen', 'DocumentID', 'Amount', 'DateOperation'] if 'EmpresaOrigen' in df_edocuenta.columns else ['DocumentID', 'Amount', 'DateOperation']
        df_edo_sub = df_edocuenta[[c for c in cols_edo if c in df_edocuenta.columns]].copy()
        if 'EmpresaOrigen' in df_edo_sub.columns and 'EmpresaOrigen' in df_oc_base.columns:
            df_oc_base = pd.merge(df_oc_base, df_edo_sub, on=['EmpresaOrigen', 'DocumentID'], how='left')
        else:
            df_oc_base = pd.merge(df_oc_base, df_edo_sub, on='DocumentID', how='left')
    else:
        df_oc_base['Amount'] = 0.0
    df_oc_base['Amount'] = pd.to_numeric(df_oc_base['Amount'], errors='coerce').fillna(0)
    df_oc_base['Total'] = pd.to_numeric(df_oc_base['Total'], errors='coerce').fillna(0)
    df_oc_base['TotalFacturaCompra'] = pd.to_numeric(df_oc_base['TotalFacturaCompra'], errors='coerce').fillna(0)
    group_keys_oc = ['EmpresaOrigen', 'DocFolio'] if 'EmpresaOrigen' in df_oc_base.columns else ['DocFolio']
    total_pagado_oc = df_oc_base.groupby(group_keys_oc + ['DocumentID'])['Amount'].transform('sum')
    df_oc_base['SaldoPagoOC'] = df_oc_base['TotalFacturaCompra'] - total_pagado_oc
    df_oc_base['SaldoPagoOC'] = df_oc_base['SaldoPagoOC'].apply(lambda x: max(0.0, x))
    df_oc_base['Tipo_Movimiento'] = 'Orden de Compra (OC)'
    df_oc_base['Saldo_Pendiente'] = df_oc_base['SaldoPagoOC']

# Procesamiento SP
df_sp_base = pd.DataFrame()
if df_tesoreria is not None and not df_tesoreria.empty:
    df_sp_base = df_tesoreria.copy()
    if df_gastos is not None:
        col_folio_gs = next((c for c in ['CFDIFolioFiscal', 'UUID', 'FolioFiscal'] if c in df_gastos.columns), None)
        col_total_gs = 'Total' if 'Total' in df_gastos.columns else None
        col_docid_gs = 'DocumentID' if 'DocumentID' in df_gastos.columns else None
        col_llave_gs = next((c for c in ['SolicitudPago', 'Solicitud de Pago', 'DocFolio'] if c in df_gastos.columns), None)
        if col_llave_gs and col_folio_gs:
            cols_gs = [col_llave_gs, col_folio_gs]
            if 'EmpresaOrigen' in df_gastos.columns: cols_gs.append('EmpresaOrigen')
            if col_total_gs: cols_gs.append(col_total_gs)
            if col_docid_gs: cols_gs.append(col_docid_gs)
            cols_gs_validas = [c for c in cols_gs if c in df_gastos.columns]
            df_gs_sub = df_gastos[cols_gs_validas].dropna(subset=[col_llave_gs]).copy()
            rename_gs = {col_llave_gs: 'DocFolio_Match', col_folio_gs: 'CFDIFolioFiscal'}
            if col_total_gs: rename_gs[col_total_gs] = 'TotalGastos'
            if col_docid_gs: rename_gs[col_docid_gs] = 'DocumentID_GS'
            df_gs_sub = df_gs_sub.rename(columns=rename_gs)
            if 'DocumentID' in df_sp_base.columns: df_sp_base = df_sp_base.drop(columns=['DocumentID'])
            if 'EmpresaOrigen' in df_gs_sub.columns and 'EmpresaOrigen' in df_sp_base.columns:
                df_sp_base = pd.merge(df_sp_base, df_gs_sub, left_on=['EmpresaOrigen', 'DocFolio'], right_on=['EmpresaOrigen', 'DocFolio_Match'], how='left')
            else:
                df_sp_base = pd.merge(df_sp_base, df_gs_sub, left_on='DocFolio', right_on='DocFolio_Match', how='left')
            if 'CFDIFolioFiscal' in df_sp_base.columns: df_sp_base['UUID'] = df_sp_base['CFDIFolioFiscal']
            if 'DocumentID_GS' in df_sp_base.columns: df_sp_base['DocumentID'] = df_sp_base['DocumentID_GS']
    if 'UUID' not in df_sp_base.columns: df_sp_base['UUID'] = ""
    if 'TotalGastos' not in df_sp_base.columns: df_sp_base['TotalGastos'] = 0.0
    if 'DocumentID' not in df_sp_base.columns: df_sp_base['DocumentID'] = ""
    if df_edocuenta is not None and 'DocumentID' in df_sp_base.columns and 'DocumentID' in df_edocuenta.columns:
        cols_edo_sp = ['EmpresaOrigen', 'DocumentID', 'Amount', 'DateOperation'] if 'EmpresaOrigen' in df_edocuenta.columns else ['DocumentID', 'Amount', 'DateOperation']
        df_edo_sub_sp = df_edocuenta[[c for c in cols_edo_sp if c in df_edocuenta.columns]].copy()
        if 'EmpresaOrigen' in df_edo_sub_sp.columns and 'EmpresaOrigen' in df_sp_base.columns:
            df_sp_base = pd.merge(df_sp_base, df_edo_sub_sp, on=['EmpresaOrigen', 'DocumentID'], how='left')
        else:
            df_sp_base = pd.merge(df_sp_base, df_edo_sub_sp, on='DocumentID', how='left')
    else:
        df_sp_base['Amount'] = 0.0
    df_sp_base['Amount'] = pd.to_numeric(df_sp_base['Amount'], errors='coerce').fillna(0)
    df_sp_base['Total'] = pd.to_numeric(df_sp_base['Total'], errors='coerce').fillna(0)
    df_sp_base['TotalGastos'] = pd.to_numeric(df_sp_base['TotalGastos'], errors='coerce').fillna(0)
    group_keys_sp = ['EmpresaOrigen', 'DocFolio'] if 'EmpresaOrigen' in df_sp_base.columns else ['DocFolio']
    total_pagado_por_doc = df_sp_base.groupby(group_keys_sp + ['DocumentID'])['Amount'].transform('sum')
    df_sp_base['SaldoPagoSP'] = df_sp_base['TotalGastos'] - total_pagado_por_doc
    df_sp_base['SaldoPagoSP'] = df_sp_base['SaldoPagoSP'].apply(lambda x: max(0.0, x))
    df_sp_base['Tipo_Movimiento'] = 'Solicitud de Pago (SP)'
    df_sp_base['Saldo_Pendiente'] = df_sp_base['SaldoPagoSP']

df_rep_list = []
if not df_oc_base.empty: df_rep_list.append(df_oc_base)
if not df_sp_base.empty: df_rep_list.append(df_sp_base)

if df_rep_list:
    df_rep_total = pd.concat(df_rep_list, ignore_index=True)
    if 'Saldo_Pendiente' in df_rep_total.columns:
        df_rep_total = df_rep_total[df_rep_total['Saldo_Pendiente'] > 1.0]

    col_emp_name = 'EmpresaOrigen' if 'EmpresaOrigen' in df_rep_total.columns else None
    col_prov_name = 'BusinessEntityName' if 'BusinessEntityName' in df_rep_total.columns else None
    col_folio_name = 'DocFolio' if 'DocFolio' in df_rep_total.columns else None
    col_currency = 'Currency' if 'Currency' in df_rep_total.columns else None

    if col_prov_name and col_folio_name:
        st.markdown("#### ⚙️ Filtros Independientes")
        lista_empresas = sorted(df_rep_total[col_emp_name].dropna().unique()) if col_emp_name else []
        tipos_disponibles = sorted(df_rep_total['Tipo_Movimiento'].dropna().unique())
        lista_proveedores = sorted(df_rep_total[col_prov_name].dropna().unique())
        lista_folios = sorted(df_rep_total[col_folio_name].dropna().unique())

        if col_emp_name:
            empresas_seleccionadas = st.multiselect("Filtrar por Empresa Origen:", lista_empresas, default=[], key="rep_emp")
            if empresas_seleccionadas: df_rep_total = df_rep_total[df_rep_total[col_emp_name].isin(empresas_seleccionadas)]

        tipos_seleccionados = st.multiselect("Filtrar por Tipo (OC / SP):", tipos_disponibles, default=[], key="rep_tipo")
        if tipos_seleccionados: df_rep_total = df_rep_total[df_rep_total['Tipo_Movimiento'].isin(tipos_seleccionados)]

        prov_seleccionados = st.multiselect("Filtrar por Proveedor(es):", lista_proveedores, default=[], key="rep_prov")
        if prov_seleccionados: df_rep_total = df_rep_total[df_rep_total[col_prov_name].isin(prov_seleccionados)]

        folios_seleccionados = st.multiselect("Filtrar por Folio(s) Específico(s):", lista_folios, default=[], key="rep_folio")
        if folios_seleccionados: df_rep_total = df_rep_total[df_rep_total[col_folio_name].isin(folios_seleccionados)]

        st.markdown("---")
        col_fecha = 'DateDocument' if 'DateDocument' in df_rep_total.columns else None
        col_desc = 'Title' if 'Title' in df_rep_total.columns else None

        if not df_rep_total.empty:
            if col_fecha: df_rep_total['Fecha_Fmt'] = pd.to_datetime(df_rep_total[col_fecha], errors='coerce').dt.strftime('%d/%m/%Y')
            else: df_rep_total['Fecha_Fmt'] = ""

            total_general_rep = df_rep_total['Saldo_Pendiente'].sum()
            st.markdown(f"### 💰 **Total Saldo Pendiente General:** {formato_mx(total_general_rep)}")
            
            empresas_agrupadas = df_rep_total[col_emp_name].dropna().unique() if col_emp_name else ['General']
            empresa_para_excel = empresas_agrupadas[0] if len(empresas_agrupadas) == 1 else "Consolidado"
            
            excel_file_bytes = generar_excel_ejecutivo(df_rep_total, empresa_para_excel)

            st.download_button(
                label="📥 Descargar Reporte en Excel con UUID",
                data=excel_file_bytes,
                file_name="Reporte_Ejecutivo_Pagos_UUID.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.markdown("---")

            for empresa in sorted(empresas_agrupadas):
                df_emp_subset = df_rep_total[df_rep_total[col_emp_name] == empresa] if col_emp_name else df_rep_total
                total_empresa = df_emp_subset['Saldo_Pendiente'].sum()
                
                st.markdown(f"## 🏢 **{empresa}** (Saldo Total Pendiente: {formato_mx(total_empresa)})")
                resumen_proveedor = df_emp_subset.groupby(col_prov_name)['Saldo_Pendiente'].sum().reset_index()
                resumen_proveedor = resumen_proveedor.sort_values(by='Saldo_Pendiente', ascending=False)

                for _, prov_row in resumen_proveedor.iterrows():
                    proveedor = prov_row[col_prov_name]
                    subtotal_prov = prov_row['Saldo_Pendiente']
                    
                    with st.expander(f"👤 {proveedor} — Saldo Pendiente Total: {formato_mx(subtotal_prov)}", expanded=True):
                        df_det_prov = df_emp_subset[df_emp_subset[col_prov_name] == proveedor]
                        monedas_del_prov = sorted(df_det_prov[col_currency].dropna().unique()) if col_currency else ['MXN']
                        
                        for moneda in monedas_del_prov:
                            df_det_moneda = df_det_prov[df_det_prov[col_currency] == moneda] if col_currency else df_det_prov
                            subtotal_moneda = df_det_moneda['Saldo_Pendiente'].sum()
                            
                            st.markdown(f"##### 💱 Moneda: **{moneda}** — Subtotal: {formato_mx(subtotal_moneda)}")
                            data_det_list = []
                            for _, row in df_det_moneda.iterrows():
                                data_det_list.append({
                                    "Tipo": row['Tipo_Movimiento'],
                                    "Fecha Vencimiento": row['Fecha_Fmt'],
                                    "Folio / Documento": row[col_folio_name],
                                    "UUID": row.get('UUID', ''),
                                    "Moneda": row[col_currency] if col_currency and not pd.isnull(row[col_currency]) else "MXN",
                                    "Descripción": row[col_desc] if col_desc else "",
                                    "Saldo Pendiente": formato_mx(row['Saldo_Pendiente'])
                                })
                            df_tabla_det = pd.DataFrame(data_det_list)
                            st.dataframe(df_tabla_det, use_container_width=True)
                            st.markdown("")
                st.markdown("---")