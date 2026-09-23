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
        prefix = "-$" if num < 0 else "$"
        return f"{prefix}{abs(int(partes[0].replace(',', ''))):,}.{decimales}"
    except (ValueError, TypeError):
        return str(val)

def generar_excel_facturacion_ejecutivo(df_datos):
    wb = Workbook()
    ws = wb.active
    ws.title = "Facturación y NC"
    ws.views.sheetView[0].showGridLines = True

    font_titulo = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    fill_titulo = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    font_normal = Font(name="Calibri", size=10, color="000000")
    font_nc = Font(name="Calibri", size=10, color="9C0006")
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_total = PatternFill(start_color="8EA9DB", end_color="8EA9DB", fill_type="solid")

    borde_delgado = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    borde_total = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    row_idx = 1
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=18)
    cell_t = ws.cell(row=row_idx, column=1, value="REPORTE EJECUTIVO DE FACTURACIÓN Y NOTAS DE CRÉDITO — GRUPO SERVYRE")
    cell_t.font = font_titulo; cell_t.fill = fill_titulo; cell_t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_idx].height = 35
    row_idx += 2

    headers = [
        'FECHA DOC', 'EMPRESA ORIGEN', 'FOLIO', 'UUID', 'TIPO DOC',
        'ESTATUS CANCELACIÓN', 'CLIENTE', 'RETENCIONES',
        'SUBTOTAL', 'DESCUENTO', 'SUBTOTAL 2', 'IMPUESTOS', 'TOTAL',
        'ESTATUS COMPLEMENTO', 'CENTRO COSTOS', 'MONTO COBRADO', 'FECHA PAGO', 'SALDO PENDIENTE'
    ]

    for col_num, h in enumerate(headers, 1):
        c = ws.cell(row=row_idx, column=col_num, value=h)
        c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde_delgado
    ws.row_dimensions[row_idx].height = 25
    row_idx += 1

    for _, r in df_datos.iterrows():
        es_nc = str(r.get('TIPO DOC', '')) == 'NC'
        f_usar = font_nc if es_nc else font_normal

        ws.cell(row=row_idx, column=1, value=str(r.get('DateDocument', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=2, value=str(r.get('EmpresaOrigen', ''))).alignment = Alignment(horizontal="left")
        ws.cell(row=row_idx, column=3, value=str(r.get('DocFolio', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=4, value=str(r.get('UUID', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=5, value=str(r.get('TIPO DOC', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=6, value=str(r.get('CFDStatusCancelledName', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=7, value=str(r.get('BusinessEntityName', ''))).alignment = Alignment(horizontal="left")
        
        cols_indices = [
            (8, 'TotalRetention'), (9, 'SubTotal'), (10, 'TotalDiscount'), (11, 'Subtotal2'),
            (12, 'TotalTax'), (13, 'Total'), (16, 'Amount'), (18, 'SaldoFactura')
        ]
        
        for c_idx, col_name in cols_indices:
            val = float(r.get(col_name, 0))
            cell_m = ws.cell(row=row_idx, column=c_idx, value=val)
            cell_m.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            cell_m.alignment = Alignment(horizontal="right")

        ws.cell(row=row_idx, column=14, value=str(r.get('StatusComplemento', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=15, value=str(r.get('CostCenterName', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=17, value=str(r.get('DateOperation', ''))).alignment = Alignment(horizontal="center")

        for c_num in range(1, 19):
            cell_curr = ws.cell(row=row_idx, column=c_num)
            cell_curr.font = f_usar
            cell_curr.border = borde_delgado

        ws.row_dimensions[row_idx].height = 18
        row_idx += 1

    ws.cell(row=row_idx, column=1, value="TOTALES").font = font_total
    for c_num in range(1, 19):
        ws.cell(row=row_idx, column=c_num).fill = fill_total
        ws.cell(row=row_idx, column=c_num).border = borde_total

    df_unicos = df_datos.drop_duplicates(subset=['EmpresaOrigen', 'DocumentID'] if 'EmpresaOrigen' in df_datos.columns and 'DocumentID' in df_datos.columns else ['DocFolio'])

    totales_map = {
        8: df_unicos['TotalRetention'].sum(),
        9: df_unicos['SubTotal'].sum(),
        10: df_unicos['TotalDiscount'].sum(),
        11: df_unicos['Subtotal2'].sum(),
        12: df_unicos['TotalTax'].sum(),
        13: df_unicos['Total'].sum(),
        16: df_datos['Amount'].sum(),
        18: df_unicos['SaldoFactura'].sum()
    }

    for c_idx, val_tot in totales_map.items():
        cell_t_val = ws.cell(row=row_idx, column=c_idx, value=val_tot)
        cell_t_val.font = font_total
        cell_t_val.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
        cell_t_val.alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[row_idx].height = 25

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

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

ruta_archivo = "Consolidado_Master.xlsx" if os.path.exists("Consolidado_Master.xlsx") else "../Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos_finanzas(path):
    if not os.path.exists(path):
        return None, None, None, None, None, None
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names

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

        df_nc = pd.read_excel(path, sheet_name='NotaCreditoCliente') if 'NotaCreditoCliente' in sheets else pd.DataFrame()
        if not df_nc.empty:
            col_del_nc = 'Deleted' if 'Deleted' in df_nc.columns else ('Delete' if 'Delete' in df_nc.columns else None)
            if col_del_nc:
                df_nc = df_nc[pd.to_numeric(df_nc[col_del_nc], errors='coerce').fillna(0) == 0].copy()
            df_nc['TIPO DOC'] = 'NC'

        df_facturacion = pd.concat([df_fac, df_nc], ignore_index=True) if not df_fac.empty or not df_nc.empty else pd.DataFrame()

        df_edo = pd.read_excel(path, sheet_name='EdoCuenta') if 'EdoCuenta' in sheets else pd.DataFrame()
        if not df_edo.empty and 'Amount' in df_edo.columns:
            df_edo['Amount'] = pd.to_numeric(df_edo['Amount'], errors='coerce').fillna(0)
            df_edo = df_edo[df_edo['Amount'] != 0].copy()

        df_tes = pd.read_excel(path, sheet_name='SolicitudPago') if 'SolicitudPago' in sheets else pd.DataFrame()
        df_ord = pd.read_excel(path, sheet_name='OrdenCompra') if 'OrdenCompra' in sheets else pd.DataFrame()
        df_fc = pd.read_excel(path, sheet_name='FacturaCompra') if 'FacturaCompra' in sheets else None
        df_gas = pd.read_excel(path, sheet_name='Gastos') if 'Gastos' in sheets else None

        # Filtrado de deleted en OC, SP, FC y Gastos
        for df_chk in [df_fc, df_gas, df_ord, df_tes]:
            if df_chk is not None and not df_chk.empty:
                col_del = 'Deleted' if 'Deleted' in df_chk.columns else ('Delete' if 'Delete' in df_chk.columns else None)
                if col_del:
                    df_chk.drop(df_chk[pd.to_numeric(df_chk[col_del], errors='coerce').fillna(0) == 1].index, inplace=True)

        return df_facturacion, df_edo, df_tes, df_ord, df_fc, df_gas
    except Exception as e:
        st.error(f"Error al cargar datos de finanzas: {e}")
        return None, None, None, None, None, None

df_factura, df_edocuenta, df_tesoreria, df_ordenes, df_fact_compra, df_gastos = cargar_datos_finanzas(ruta_archivo)

# LÓGICA DE VINCULACIÓN EN 3 NIVELES (OC / SP -> FacturaCompra/Gastos -> EdoCuenta)
@st.cache_data
def procesar_oc_sp_completo(_df_ord, _df_tes, _df_fc, _df_gas, _df_edo):
    # 1. Procesamiento de Órdenes de Compra (OC)
    df_oc_calc = pd.DataFrame()
    if _df_ord is not None and not _df_ord.empty:
        df_oc_calc = _df_ord.copy()
        if _df_fc is not None and not _df_fc.empty:
            col_uuid_fc = 'UUID' if 'UUID' in _df_fc.columns else ('CFDIFolioFiscal' if 'CFDIFolioFiscal' in _df_fc.columns else None)
            col_total_fc = 'Total' if 'Total' in _df_fc.columns else None
            col_docid_fc = 'DocumentID' if 'DocumentID' in _df_fc.columns else None
            col_llave_fc = next((c for c in ['Solicitud de Pago', 'SolicitudPago', 'DocFolio'] if c in _df_fc.columns), None)
            
            if col_llave_fc and col_uuid_fc:
                cols_fc = [col_llave_fc, col_uuid_fc]
                if 'EmpresaOrigen' in _df_fc.columns: cols_fc.append('EmpresaOrigen')
                if col_total_fc: cols_fc.append(col_total_fc)
                if col_docid_fc: cols_fc.append(col_docid_fc)
                
                df_fc_sub = _df_fc[cols_fc].dropna(subset=[col_llave_fc]).copy()
                rename_dict = {col_llave_fc: 'DocFolio_Match', col_uuid_fc: 'UUID'}
                if col_total_fc: rename_dict[col_total_fc] = 'TotalFacturaCompra'
                if col_docid_fc: rename_dict[col_docid_fc] = 'DocumentID_FC'
                df_fc_sub = df_fc_sub.rename(columns=rename_dict)
                
                if 'DocumentID' in df_oc_calc.columns: df_oc_calc = df_oc_calc.drop(columns=['DocumentID'])
                if 'EmpresaOrigen' in df_fc_sub.columns and 'EmpresaOrigen' in df_oc_calc.columns:
                    df_oc_calc = pd.merge(df_oc_calc, df_fc_sub, left_on=['EmpresaOrigen', 'DocFolio'], right_on=['EmpresaOrigen', 'DocFolio_Match'], how='left')
                else:
                    df_oc_calc = pd.merge(df_oc_calc, df_fc_sub, left_on='DocFolio', right_on='DocFolio_Match', how='left')
                
                if 'DocumentID_FC' in df_oc_calc.columns: df_oc_calc['DocumentID'] = df_oc_calc['DocumentID_FC']

        if 'UUID' not in df_oc_calc.columns: df_oc_calc['UUID'] = ""
        if 'TotalFacturaCompra' not in df_oc_calc.columns: df_oc_calc['TotalFacturaCompra'] = 0.0
        if 'DocumentID' not in df_oc_calc.columns: df_oc_calc['DocumentID'] = ""

        if _df_edo is not None and not _df_edo.empty and 'DocumentID' in df_oc_calc.columns and 'DocumentID' in _df_edo.columns:
            cols_edo = ['EmpresaOrigen', 'DocumentID', 'Amount'] if 'EmpresaOrigen' in _df_edo.columns else ['DocumentID', 'Amount']
            df_edo_sub = _df_edo[[c for c in cols_edo if c in _df_edo.columns]].copy()
            if 'EmpresaOrigen' in df_edo_sub.columns and 'EmpresaOrigen' in df_oc_calc.columns:
                df_oc_calc = pd.merge(df_oc_calc, df_edo_sub, on=['EmpresaOrigen', 'DocumentID'], how='left')
            else:
                df_oc_calc = pd.merge(df_oc_calc, df_edo_sub, on='DocumentID', how='left')
        else:
            df_oc_calc['Amount'] = 0.0

        df_oc_calc['Amount'] = pd.to_numeric(df_oc_calc['Amount'], errors='coerce').fillna(0)
        df_oc_calc['Total'] = pd.to_numeric(df_oc_calc['Total'], errors='coerce').fillna(0)
        df_oc_calc['TotalFacturaCompra'] = pd.to_numeric(df_oc_calc['TotalFacturaCompra'], errors='coerce').fillna(0)

        group_keys_oc = ['EmpresaOrigen', 'DocFolio'] if 'EmpresaOrigen' in df_oc_calc.columns else ['DocFolio']
        total_pagado_oc = df_oc_calc.groupby(group_keys_oc + ['DocumentID'])['Amount'].transform('sum')
        
        # Saldo = Total Factura Compra (o Total OC) menos Pagos en EdoCuenta
        base_monto_oc = df_oc_calc['TotalFacturaCompra'].apply(lambda x: x if x > 0 else 0)
        df_oc_calc['Saldo_Pendiente'] = (base_monto_oc - total_pagado_oc).apply(lambda x: max(0.0, x))
        df_oc_calc['Tipo_Movimiento'] = 'Orden de Compra (OC)'

    # 2. Procesamiento de Solicitudes de Pago (SP)
    df_sp_calc = pd.DataFrame()
    if _df_tes is not None and not _df_tes.empty:
        df_sp_calc = _df_tes.copy()
        if _df_gas is not None and not _df_gas.empty:
            col_folio_gs = next((c for c in ['CFDIFolioFiscal', 'UUID', 'FolioFiscal'] if c in _df_gas.columns), None)
            col_total_gs = 'Total' if 'Total' in _df_gas.columns else None
            col_docid_gs = 'DocumentID' if 'DocumentID' in _df_gas.columns else None
            col_llave_gs = next((c for c in ['SolicitudPago', 'Solicitud de Pago', 'DocFolio'] if c in _df_gas.columns), None)

            if col_llave_gs and col_folio_gs:
                cols_gs = [col_llave_gs, col_folio_gs]
                if 'EmpresaOrigen' in _df_gas.columns: cols_gs.append('EmpresaOrigen')
                if col_total_gs: cols_gs.append(col_total_gs)
                if col_docid_gs: cols_gs.append(col_docid_gs)

                df_gs_sub = _df_gas[cols_gs].dropna(subset=[col_llave_gs]).copy()
                rename_gs = {col_llave_gs: 'DocFolio_Match', col_folio_gs: 'CFDIFolioFiscal'}
                if col_total_gs: rename_gs[col_total_gs] = 'TotalGastos'
                if col_docid_gs: rename_gs[col_docid_gs] = 'DocumentID_GS'
                df_gs_sub = df_gs_sub.rename(columns=rename_gs)

                if 'DocumentID' in df_sp_calc.columns: df_sp_calc = df_sp_calc.drop(columns=['DocumentID'])
                if 'EmpresaOrigen' in df_gs_sub.columns and 'EmpresaOrigen' in df_sp_calc.columns:
                    df_sp_calc = pd.merge(df_sp_calc, df_gs_sub, left_on=['EmpresaOrigen', 'DocFolio'], right_on=['EmpresaOrigen', 'DocFolio_Match'], how='left')
                else:
                    df_sp_calc = pd.merge(df_sp_calc, df_gs_sub, left_on='DocFolio', right_on='DocFolio_Match', how='left')

                if 'CFDIFolioFiscal' in df_sp_calc.columns: df_sp_calc['UUID'] = df_sp_calc['CFDIFolioFiscal']
                if 'DocumentID_GS' in df_sp_calc.columns: df_sp_calc['DocumentID'] = df_sp_calc['DocumentID_GS']

        if 'UUID' not in df_sp_calc.columns: df_sp_calc['UUID'] = ""
        if 'TotalGastos' not in df_sp_calc.columns: df_sp_calc['TotalGastos'] = 0.0
        if 'DocumentID' not in df_sp_calc.columns: df_sp_calc['DocumentID'] = ""

        if _df_edo is not None and not _df_edo.empty and 'DocumentID' in df_sp_calc.columns and 'DocumentID' in _df_edo.columns:
            cols_edo_sp = ['EmpresaOrigen', 'DocumentID', 'Amount'] if 'EmpresaOrigen' in _df_edo.columns else ['DocumentID', 'Amount']
            df_edo_sub_sp = _df_edo[[c for c in cols_edo_sp if c in _df_edo.columns]].copy()
            if 'EmpresaOrigen' in df_edo_sub_sp.columns and 'EmpresaOrigen' in df_sp_calc.columns:
                df_sp_calc = pd.merge(df_sp_calc, df_edo_sub_sp, on=['EmpresaOrigen', 'DocumentID'], how='left')
            else:
                df_sp_calc = pd.merge(df_sp_calc, df_edo_sub_sp, on='DocumentID', how='left')
        else:
            df_sp_calc['Amount'] = 0.0

        df_sp_calc['Amount'] = pd.to_numeric(df_sp_calc['Amount'], errors='coerce').fillna(0)
        df_sp_calc['Total'] = pd.to_numeric(df_sp_calc['Total'], errors='coerce').fillna(0)
        df_sp_calc['TotalGastos'] = pd.to_numeric(df_sp_calc['TotalGastos'], errors='coerce').fillna(0)

        group_keys_sp = ['EmpresaOrigen', 'DocFolio'] if 'EmpresaOrigen' in df_sp_calc.columns else ['DocFolio']
        total_pagado_sp = df_sp_calc.groupby(group_keys_sp + ['DocumentID'])['Amount'].transform('sum')

        base_monto_sp = df_sp_calc['TotalGastos'].apply(lambda x: x if x > 0 else 0)
        df_sp_calc['Saldo_Pendiente'] = (base_monto_sp - total_pagado_sp).apply(lambda x: max(0.0, x))
        df_sp_calc['Tipo_Movimiento'] = 'Solicitud de Pago (SP)'

    return df_oc_calc, df_sp_calc

df_ordenes_proc, df_tesoreria_proc = procesar_oc_sp_completo(df_ordenes, df_tesoreria, df_fact_compra, df_gastos, df_edocuenta)

st.sidebar.title("💰 Módulo de Finanzas")
st.sidebar.markdown("---")
submodulo = st.sidebar.radio(
    "Seleccione Submódulo:",
    ["📊 Facturación", "📦 Órdenes de Compra y SP", "📋 Reporte Ejecutivo de Pagos"],
    key="sub_finanzas_nav"
)

# ==========================================
# 1. SUBMÓDULO: FACTURACIÓN Y NC
# ==========================================
if submodulo == "📊 Facturación":
    columnas_requeridas = [
        'DateDocument', 'EmpresaOrigen', 'DocFolio', 'UUID', 'TIPO DOC',
        'CFDStatusCancelledName', 'BusinessEntityName', 'TotalRetention',
        'SubTotal', 'TotalDiscount', 'Subtotal2', 'TotalTax', 'Total',
        'StatusComplemento', 'CostCenterName', 'Amount', 'DateOperation', 'SaldoFactura'
    ]

    st.title("📊 Módulo de Facturación y Cobranza")

    if df_factura is not None and not df_factura.empty:
        df_f = df_factura.copy()

        subtotal_val = pd.to_numeric(df_f['SubTotal'], errors='coerce').fillna(0) if 'SubTotal' in df_f.columns else 0.0
        discount_val = pd.to_numeric(df_f['TotalDiscount'], errors='coerce').fillna(0) if 'TotalDiscount' in df_f.columns else 0.0
        df_f['Subtotal2'] = subtotal_val - discount_val

        if 'DateDocument' in df_f.columns:
            df_f['DateDocument_Fmt'] = pd.to_datetime(df_f['DateDocument'], errors='coerce').dt.strftime('%Y-%m-%d')
            df_f['Año'] = pd.to_datetime(df_f['DateDocument'], errors='coerce').dt.year.fillna(0).astype(int)
        else:
            df_f['DateDocument_Fmt'] = ""
            df_f['Año'] = 0

        if df_edocuenta is not None and not df_edocuenta.empty and 'DocumentID' in df_f.columns and 'DocumentID' in df_edocuenta.columns:
            cols_edo = ['DocumentID', 'Amount', 'DateOperation']
            if 'EmpresaOrigen' in df_edocuenta.columns and 'EmpresaOrigen' in df_f.columns:
                cols_edo.append('EmpresaOrigen')

            df_edo_sub = df_edocuenta[[c for c in cols_edo if c in df_edocuenta.columns]].copy()
            if 'DateOperation' in df_edo_sub.columns:
                df_edo_sub['DateOperation'] = pd.to_datetime(df_edo_sub['DateOperation'], errors='coerce').dt.strftime('%Y-%m-%d')
            else:
                df_edo_sub['DateOperation'] = ""

            group_keys = ['EmpresaOrigen', 'DocumentID'] if 'EmpresaOrigen' in df_edo_sub.columns and 'EmpresaOrigen' in df_f.columns else 'DocumentID'
            df_tot_pagado = df_edo_sub.groupby(group_keys)['Amount'].sum().reset_index().rename(columns={'Amount': 'Total_Pagos_Acumulados'})

            df_f = pd.merge(df_f, df_tot_pagado, on=group_keys, how='left')
            df_f['Total_Pagos_Acumulados'] = df_f['Total_Pagos_Acumulados'].fillna(0.0)

            df_f = pd.merge(df_f, df_edo_sub, on=group_keys, how='left')
            df_f['Amount'] = df_f['Amount'].fillna(0.0)
            df_f['DateOperation'] = df_f['DateOperation'].fillna("")
        else:
            df_f['Amount'] = 0.0
            df_f['Total_Pagos_Acumulados'] = 0.0
            df_f['DateOperation'] = ""

        total_fac_val = pd.to_numeric(df_f['Total'], errors='coerce').fillna(0) if 'Total' in df_f.columns else 0.0
        df_f['SaldoFactura'] = (total_fac_val - df_f['Total_Pagos_Acumulados']).apply(lambda x: max(0.0, x))
        df_f['DateDocument'] = df_f['DateDocument_Fmt']

        # Multiplica Notas de Crédito por -1 para restar
        es_nc_mask = df_f['TIPO DOC'] == 'NC'
        cols_a_restar = ['SubTotal', 'TotalDiscount', 'Subtotal2', 'TotalTax', 'Total', 'SaldoFactura', 'TotalRetention']
        
        for col_r in cols_a_restar:
            if col_r in df_f.columns:
                df_f.loc[es_nc_mask, col_r] = -1 * df_f.loc[es_nc_mask, col_r].abs()

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

        # MÉTRICAS
        df_unicos = df_f.drop_duplicates(subset=['EmpresaOrigen', 'DocumentID'] if 'EmpresaOrigen' in df_f.columns and 'DocumentID' in df_f.columns else ['DocFolio'])
        
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Subtotal2 Total", formato_mx(df_unicos['Subtotal2'].sum()))
        with col_m2:
            st.metric("Total Facturado (Neto)", formato_mx(df_unicos['Total'].sum()))
        with col_m3:
            st.metric("Cobrado en EdoCuenta", formato_mx(df_f['Amount'].sum()))
        with col_m4:
            st.metric("Saldo Pendiente Neto", formato_mx(df_unicos['SaldoFactura'].sum()))

        st.markdown("---")

        # DASHBOARD GRÁFICO
        g1, g2 = st.columns(2)
        with g1:
            st.subheader("📈 Top 10 Clientes por Saldo Pendiente")
            if 'BusinessEntityName' in df_unicos.columns:
                top_clientes = df_unicos.groupby('BusinessEntityName')['SaldoFactura'].sum().nlargest(10).reset_index()
                st.bar_chart(top_clientes.set_index('BusinessEntityName'))
        with g2:
            st.subheader("🏢 Distribución por Empresa Origen")
            if 'EmpresaOrigen' in df_unicos.columns:
                emp_totales = df_unicos.groupby('EmpresaOrigen')['Total'].sum().reset_index()
                st.bar_chart(emp_totales.set_index('EmpresaOrigen'))

        st.markdown("---")
        st.subheader("📋 Tabla General de Facturación y Notas de Crédito")

        st.download_button(
            label="📊 Descargar Reporte Ejecutivo de Facturación y NC en Excel",
            data=generar_excel_facturacion_ejecutivo(df_f),
            file_name="Reporte_Ejecutivo_Facturacion_NC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_exp_fact_ejec_v3",
            use_container_width=True
        )

        st.markdown("<br>", unsafe_allow_html=True)

        df_view = df_f.copy()
        for col_m in ['TotalRetention', 'SubTotal', 'TotalDiscount', 'Subtotal2', 'TotalTax', 'Total', 'Amount', 'SaldoFactura']:
            if col_m in df_view.columns:
                df_view[col_m] = df_view[col_m].apply(formato_mx)

        cols_disponibles = [c for c in columnas_requeridas if c in df_view.columns]
        st.dataframe(df_view[cols_disponibles], use_container_width=True)
    else:
        st.warning("No hay registros disponibles de Facturación.")

# ==========================================
# 2. SUBMÓDULO: ÓRDENES DE COMPRA Y SP
# ==========================================
elif submodulo == "📦 Órdenes de Compra y SP":
    st.title("📦 Módulo de Órdenes de Compra y Solicitudes de Pago")
    columnas_oc_sp = ['EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument', 'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'Total', 'UUID', 'Saldo_Pendiente']
    
    tab_oc, tab_sp = st.tabs(["OrdenCompra", "SolicitudPago"])
    
    with tab_oc:
        if df_ordenes_proc is not None and not df_ordenes_proc.empty:
            df_o = df_ordenes_proc.copy()
            if 'DateDocument' in df_o.columns:
                df_o['DateDocument'] = pd.to_datetime(df_o['DateDocument'], errors='coerce').dt.strftime('%Y-%m-%d')
                df_o['Año'] = pd.to_datetime(df_o['DateDocument'], errors='coerce').dt.year.fillna(0).astype(int)
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
            df_v_oc = df_o.copy()
            for col_m in ['SubTotal', 'Total', 'Saldo_Pendiente']:
                if col_m in df_v_oc.columns: df_v_oc[col_m] = df_v_oc[col_m].apply(formato_mx)
            st.dataframe(df_v_oc[[c for c in columnas_oc_sp if c in df_v_oc.columns]], use_container_width=True)

    with tab_sp:
        if df_tesoreria_proc is not None and not df_tesoreria_proc.empty:
            df_s = df_tesoreria_proc.copy()
            if 'DateDocument' in df_s.columns:
                df_s['DateDocument'] = pd.to_datetime(df_s['DateDocument'], errors='coerce').dt.strftime('%Y-%m-%d')
                df_s['Año'] = pd.to_datetime(df_s['DateDocument'], errors='coerce').dt.year.fillna(0).astype(int)
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
            df_v_sp = df_s.copy()
            for col_m in ['SubTotal', 'Total', 'Saldo_Pendiente']:
                if col_m in df_v_sp.columns: df_v_sp[col_m] = df_v_sp[col_m].apply(formato_mx)
            st.dataframe(df_v_sp[[c for c in columnas_oc_sp if c in df_v_sp.columns]], use_container_width=True)

# ==========================================
# 3. SUBMÓDULO: REPORTE EJECUTIVO DE PAGOS
# ==========================================
elif submodulo == "📋 Reporte Ejecutivo de Pagos":
    st.title("📋 Reporte Ejecutivo de Pagos con Saldo Pendiente")
    
    df_oc_b = df_ordenes_proc.copy() if df_ordenes_proc is not None and not df_ordenes_proc.empty else pd.DataFrame()
    df_sp_b = df_tesoreria_proc.copy() if df_tesoreria_proc is not None and not df_tesoreria_proc.empty else pd.DataFrame()
    
    df_list = [df for df in [df_oc_b, df_sp_b] if not df.empty]
    if df_list:
        df_tot = pd.concat(df_list, ignore_index=True)
        if 'Saldo_Pendiente' in df_tot.columns:
            df_tot['Saldo_Pendiente'] = pd.to_numeric(df_tot['Saldo_Pendiente'], errors='coerce').fillna(0)
            df_tot = df_tot[df_tot['Saldo_Pendiente'] > 1.0]
        
        st.markdown(f"### 💰 **Total Saldo Pendiente General:** {formato_mx(df_tot['Saldo_Pendiente'].sum())}")
        
        empresa_excel = df_tot['EmpresaOrigen'].iloc[0] if 'EmpresaOrigen' in df_tot.columns and len(df_tot['EmpresaOrigen'].dropna().unique()) == 1 else "Consolidado"
        st.download_button(
            label="📥 Descargar Reporte Ejecutivo de Pagos en Excel (Con UUID)",
            data=generar_excel_ejecutivo(df_tot, empresa_excel),
            file_name="Reporte_Ejecutivo_Pagos_UUID.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("---")
        
        df_view_rep = df_tot.copy()
        if 'DateDocument' in df_view_rep.columns:
            df_view_rep['Fecha_Fmt'] = pd.to_datetime(df_view_rep['DateDocument'], errors='coerce').dt.strftime('%d/%m/%Y')
        else:
            df_view_rep['Fecha_Fmt'] = ""

        # Presentación agrupada por Empresa, Proveedor y Moneda
        col_emp_n = 'EmpresaOrigen' if 'EmpresaOrigen' in df_view_rep.columns else None
        col_prov_n = 'BusinessEntityName' if 'BusinessEntityName' in df_view_rep.columns else None
        col_curr_n = 'Currency' if 'Currency' in df_view_rep.columns else None
        col_fol_n = 'DocFolio' if 'DocFolio' in df_view_rep.columns else None
        col_desc_n = 'Title' if 'Title' in df_view_rep.columns else None

        empresas_agrupadas = sorted(df_view_rep[col_emp_n].dropna().unique()) if col_emp_n else ['General']

        for empresa in empresas_agrupadas:
            df_emp_s = df_view_rep[df_view_rep[col_emp_n] == empresa] if col_emp_n else df_view_rep
            st.markdown(f"## 🏢 **{empresa}** (Saldo Total Pendiente: {formato_mx(df_emp_s['Saldo_Pendiente'].sum())})")
            
            resumen_prov = df_emp_s.groupby(col_prov_n)['Saldo_Pendiente'].sum().reset_index().sort_values(by='Saldo_Pendiente', ascending=False)
            
            for _, prov_row in resumen_prov.iterrows():
                proveedor = prov_row[col_prov_n]
                subtotal_prov = prov_row['Saldo_Pendiente']
                
                with st.expander(f"👤 {proveedor} — Saldo Pendiente Total: {formato_mx(subtotal_prov)}", expanded=True):
                    df_det_p = df_emp_s[df_emp_s[col_prov_n] == proveedor]
                    monedas_p = sorted(df_det_p[col_curr_n].dropna().unique()) if col_curr_n else ['MXN']
                    
                    for moneda in monedas_p:
                        df_det_m = df_det_p[df_det_p[col_curr_n] == moneda] if col_curr_n else df_det_p
                        st.markdown(f"##### 💱 Moneda: **{moneda}** — Subtotal: {formato_mx(df_det_m['Saldo_Pendiente'].sum())}")
                        
                        det_list = []
                        for _, row in df_det_m.iterrows():
                            det_list.append({
                                "Tipo": row.get('Tipo_Movimiento', ''),
                                "Fecha": row.get('Fecha_Fmt', ''),
                                "Folio": row.get(col_fol_n, ''),
                                "UUID": row.get('UUID', ''),
                                "Moneda": row.get(col_curr_n, 'MXN'),
                                "Descripción": row.get(col_desc_n, ''),
                                "Saldo Pendiente": formato_mx(row.get('Saldo_Pendiente', 0))
                            })
                        st.dataframe(pd.DataFrame(det_list), use_container_width=True)
            st.markdown("---")
    else:
        st.warning("No hay registros pendientes para el reporte de pagos.")