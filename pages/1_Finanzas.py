import streamlit as st

import pandas as pd

import io

import os

from openpyxl import Workbook

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from openpyxl.utils import get_column_letter



st.set_page_config(page_title="Finanzas - Grupo SERVYRE", layout="wide")



# ==========================================

# 🎨 ESTILOS CSS PERSONALIZADOS (ESTILO ORACLE CLOUD / ENTERPRISE)

# ==========================================

st.markdown("""

    <style>

        /* Estilo para tablas HTML personalizadas */

        .oracle-table-container {

            width: 100%;

            overflow-x: auto;

            margin-bottom: 20px;

            border: 1px solid #c0c0c0;

            border-radius: 4px;

            box-shadow: 0 1px 3px rgba(0,0,0,0.05);

        }

        .oracle-table {

            width: 100%;

            border-collapse: collapse;

            font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;

            font-size: 13px;

            background-color: #ffffff;

            color: #333333;

        }

        .oracle-table th {

            background-color: #f0f2f5;

            color: #1f497d;

            font-weight: bold;

            text-align: center;

            padding: 8px 10px;

            border: 1px solid #d9d9d9;

            font-size: 12px;

            text-transform: uppercase;

            letter-spacing: 0.5px;

        }

        .oracle-table td {

            padding: 6px 10px;

            border: 1px solid #e5e5e5;

            vertical-align: middle;

        }

        .oracle-table tr:nth-child(even) {

            background-color: #fcfcfc;

        }

        .oracle-table tr:hover {

            background-color: #f0f4f9;

        }

        .oracle-table tr.total-row {

            background-color: #e6ecf5 !important;

            font-weight: bold;

            color: #000000;

            border-top: 2px solid #1f497d;

            border-bottom: 2px solid #1f497d;

        }

        .oracle-table tr.total-row td {

            border: 1px solid #b0c4de;

        }

        .text-center { text-align: center; }

        .text-right { text-align: right; }

        .text-left { text-align: left; }

    </style>

""", unsafe_allow_html=True)



# --- FORMATO DE MONEDA REGIÓN MÉXICO ($1,234,567.89) ---

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



# --- FUNCIÓN AUXILIAR DE FILTRADO ESTRICTO DE REGISTROS ELIMINADOS (DELETED == 1) ---

def filtrar_no_eliminados(df):

    if df is None or df.empty:

        return df

    col_del = next((c for c in ['Deleted', 'Delete', 'deleted', 'delete'] if c in df.columns), None)

    if col_del:

        df = df[pd.to_numeric(df[col_del], errors='coerce').fillna(0) == 0].copy()

    return df



# --- FUNCIÓN GENERAR EXCEL FACTURACIÓN ---

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



# --- FUNCIÓN GENERAR EXCEL EJECUTIVO DE PAGOS ---

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

    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=9)

    cell = ws.cell(row=row_idx, column=1, value=f"REPORTE EJECUTIVO DE PAGOS — {empresa_nombre.upper()}")

    cell.font = font_titulo; cell.fill = fill_titulo; cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[row_idx].height = 35

    row_idx += 2



    headers = ["TIPO", "PROVEEDOR", "FECHA VENCIMIENTO", "FOLIO / DOCUMENTO", "DOCUMENT ID", "UUID", "MONEDA", "DESCRIPCIÓN", "SALDO PENDIENTE"]

    col_emp = 'EmpresaOrigen' if 'EmpresaOrigen' in df_datos.columns else None

    col_prov = 'BusinessEntityName' if 'BusinessEntityName' in df_datos.columns else None

    col_curr = 'Currency' if 'Currency' in df_datos.columns else None



    empresas = sorted(df_datos[col_emp].unique()) if col_emp else ['General']



    for empresa in empresas:

        df_emp = df_datos[df_datos[col_emp] == empresa] if col_emp else df_datos

        total_empresa = df_emp['Saldo_Pendiente'].sum()



        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)

        cell_emp = ws.cell(row=row_idx, column=1, value=f"EMPRESA: {str(empresa).upper()}")

        cell_emp.font = font_empresa; cell_emp.fill = fill_empresa; cell_emp.alignment = Alignment(horizontal="left", vertical="center")

        

        cell_emp_tot = ws.cell(row=row_idx, column=9, value=total_empresa)

        cell_emp_tot.font = font_empresa; cell_emp_tot.fill = fill_empresa; cell_emp_tot.number_format = '"$"#,##0.00'; cell_emp_tot.alignment = Alignment(horizontal="right", vertical="center")

        

        for c_idx in range(1, 10): ws.cell(row=row_idx, column=c_idx).border = borde_delgado

        ws.row_dimensions[row_idx].height = 26

        row_idx += 1



        proveedores = sorted(df_emp[col_prov].unique()) if col_prov else []

        for prov in proveedores:

            df_prov = df_emp[df_emp[col_prov] == prov]

            total_prov = df_prov['Saldo_Pendiente'].sum()



            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)

            cell_prov = ws.cell(row=row_idx, column=1, value=f"   👤 {str(prov)}")

            cell_prov.font = font_prov; cell_prov.fill = fill_prov; cell_prov.alignment = Alignment(horizontal="left", vertical="center")

            

            cell_prov_tot = ws.cell(row=row_idx, column=9, value=total_prov)

            cell_prov_tot.font = font_prov; cell_prov_tot.fill = fill_prov; cell_prov_tot.number_format = '"$"#,##0.00'; cell_prov_tot.alignment = Alignment(horizontal="right", vertical="center")

            

            for c_idx in range(1, 10): ws.cell(row=row_idx, column=c_idx).border = borde_delgado

            ws.row_dimensions[row_idx].height = 22

            row_idx += 1



            monedas_prov = sorted(df_prov[col_curr].dropna().unique()) if col_curr else ['MXN']

            for moneda in monedas_prov:

                df_moneda = df_prov[df_prov[col_curr] == moneda] if col_curr else df_prov

                total_moneda = df_moneda['Saldo_Pendiente'].sum()



                ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)

                cell_mon = ws.cell(row=row_idx, column=1, value=f"      💱 Moneda: {str(moneda)}")

                cell_mon.font = font_moneda; cell_mon.fill = fill_moneda; cell_mon.alignment = Alignment(horizontal="left", vertical="center")

                

                cell_mon_tot = ws.cell(row=row_idx, column=9, value=total_moneda)

                cell_mon_tot.font = font_moneda; cell_mon_tot.fill = fill_moneda; cell_mon_tot.number_format = '"$"#,##0.00'; cell_mon_tot.alignment = Alignment(horizontal="right", vertical="center")

                

                for c_idx in range(1, 10): ws.cell(row=row_idx, column=c_idx).border = borde_delgado

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

                    ws.cell(row=row_idx, column=5, value=str(r.get('DocumentID', ''))).alignment = Alignment(horizontal="center")

                    ws.cell(row=row_idx, column=6, value=str(r.get('UUID', ''))).alignment = Alignment(horizontal="center")

                    ws.cell(row=row_idx, column=7, value=str(r.get('Currency', 'MXN'))).alignment = Alignment(horizontal="center")

                    ws.cell(row=row_idx, column=8, value=str(r.get('Title', ''))).alignment = Alignment(horizontal="left")

                    

                    c_val = ws.cell(row=row_idx, column=9, value=float(r.get('Saldo_Pendiente', 0)))

                    c_val.number_format = '"$"#,##0.00'; c_val.alignment = Alignment(horizontal="right")



                    for col_num in range(1, 10): ws.cell(row=row_idx, column=col_num).font = font_normal; ws.cell(row=row_idx, column=col_num).border = borde_delgado

                    ws.row_dimensions[row_idx].height = 18

                    row_idx += 1

                row_idx += 1

            row_idx += 1

        row_idx += 1



    ws.cell(row=row_idx, column=1, value="TOTAL GENERAL").font = font_total

    for col_num in range(1, 9): ws.cell(row=row_idx, column=col_num).fill = fill_total; ws.cell(row=row_idx, column=col_num).border = borde_total

    

    cell_tot_gen = ws.cell(row=row_idx, column=9, value=df_datos['Saldo_Pendiente'].sum())

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

        df_fac = filtrar_no_eliminados(df_fac)

        if df_fac is not None and not df_fac.empty:

            if 'ModuleName' in df_fac.columns:

                df_fac['TIPO DOC'] = df_fac['ModuleName'].astype(str).apply(

                    lambda x: 'NC' if 'nota' in x.lower() or 'credito' in x.lower() or 'crédito' in x.lower() else 'F'

                )

            else:

                df_fac['TIPO DOC'] = 'F'



        df_nc = pd.read_excel(path, sheet_name='NotaCreditoCliente') if 'NotaCreditoCliente' in sheets else pd.DataFrame()

        df_nc = filtrar_no_eliminados(df_nc)

        if df_nc is not None and not df_nc.empty:

            df_nc['TIPO DOC'] = 'NC'



        df_facturacion = pd.concat([df_fac, df_nc], ignore_index=True) if not df_fac.empty or not df_nc.empty else pd.DataFrame()



        df_edo = pd.read_excel(path, sheet_name='EdoCuenta') if 'EdoCuenta' in sheets else pd.DataFrame()

        df_edo = filtrar_no_eliminados(df_edo)

        if df_edo is not None and not df_edo.empty and 'Amount' in df_edo.columns:

            df_edo['Amount'] = pd.to_numeric(df_edo['Amount'], errors='coerce').fillna(0)

            df_edo = df_edo[df_edo['Amount'] != 0].copy()



        df_tes = pd.read_excel(path, sheet_name='SolicitudPago') if 'SolicitudPago' in sheets else pd.DataFrame()

        df_tes = filtrar_no_eliminados(df_tes)



        df_ord = pd.read_excel(path, sheet_name='OrdenCompra') if 'OrdenCompra' in sheets else pd.DataFrame()

        df_ord = filtrar_no_eliminados(df_ord)



        df_fc = pd.read_excel(path, sheet_name='FacturaCompra') if 'FacturaCompra' in sheets else None

        df_fc = filtrar_no_eliminados(df_fc)



        df_gas = pd.read_excel(path, sheet_name='Gastos') if 'Gastos' in sheets else None

        df_gas = filtrar_no_eliminados(df_gas)



        return df_facturacion, df_edo, df_tes, df_ord, df_fc, df_gas

    except Exception as e:

        st.error(f"Error al cargar datos de finanzas: {e}")

        return None, None, None, None, None, None



df_factura, df_edocuenta, df_tesoreria, df_ordenes, df_fact_compra, df_gastos = cargar_datos_finanzas(ruta_archivo)



# --- PROCESAMIENTO CON DESGLOSE POR RENGLÓN SI HAY MÚLTIPLES FACTURAS/PAGOS ---

@st.cache_data

def procesar_oc_sp_definitivo(_df_ord, _df_tes, _df_fc, _df_gas, _df_edo):

    pagos_edo_map = {}

    if _df_edo is not None and not _df_edo.empty:

        col_doc_edo = 'DocumentID' if 'DocumentID' in _df_edo.columns else ('Document' if 'Document' in _df_edo.columns else None)

        if col_doc_edo:

            df_edo_c = _df_edo.copy()

            df_edo_c['DocID_Clean'] = df_edo_c[col_doc_edo].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()

            df_edo_c['Amount'] = pd.to_numeric(df_edo_c['Amount'], errors='coerce').fillna(0)

            

            if 'EmpresaOrigen' in df_edo_c.columns:

                df_edo_c['Emp_Clean'] = df_edo_c['EmpresaOrigen'].astype(str).str.strip()

                pagos_edo_map = df_edo_c.groupby(['Emp_Clean', 'DocID_Clean'])['Amount'].sum().to_dict()



    # PREPARAR HOJA DE GASTOS

    df_gs_m = pd.DataFrame()

    if _df_gas is not None and not _df_gas.empty:

        col_sp_gs = next((c for c in ['SolicitudPago', 'Solicitud de Pago', 'OrdenCompra', 'DocFolio'] if c in _df_gas.columns), None)

        col_doc_gs = 'DocumentID' if 'DocumentID' in _df_gas.columns else ('Document' if 'Document' in _df_gas.columns else None)

        col_uuid_gs = next((c for c in ['CFDIFolioFiscal', 'UUID', 'FolioFiscal'] if c in _df_gas.columns), None)

        col_emp_gs = next((c for c in ['EmpresaOrigen', 'Empresa Origen', 'Empresa'] if c in _df_gas.columns), None)

        col_tot_gs = next((c for c in ['TotalM', 'Total', 'SubTotal'] if c in _df_gas.columns), None)



        if col_sp_gs and col_doc_gs:

            df_gs_m = _df_gas.copy()

            df_gs_m['Folio_Match'] = df_gs_m[col_sp_gs].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()

            df_gs_m['DocumentID_GS'] = df_gs_m[col_doc_gs].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()

            if col_emp_gs:

                df_gs_m['Emp_GS'] = df_gs_m[col_emp_gs].astype(str).str.strip()

            if col_uuid_gs:

                df_gs_m['UUID_GS'] = df_gs_m[col_uuid_gs].astype(str).str.strip()

            if col_tot_gs:

                df_gs_m['Monto_GS'] = pd.to_numeric(df_gs_m[col_tot_gs], errors='coerce').fillna(0)



    # PREPARAR HOJA DE FACTURACOMPRA

    df_fc_m = pd.DataFrame()

    if _df_fc is not None and not _df_fc.empty:

        col_sp_fc = next((c for c in ['Solicitud de Pago', 'SolicitudPago', 'OrdenCompra', 'DocFolio'] if c in _df_fc.columns), None)

        col_doc_fc = 'DocumentID' if 'DocumentID' in _df_fc.columns else ('Document' if 'Document' in _df_fc.columns else None)

        col_uuid_fc = 'UUID' if 'UUID' in _df_fc.columns else ('CFDIFolioFiscal' if 'CFDIFolioFiscal' in _df_fc.columns else None)

        col_emp_fc = next((c for c in ['EmpresaOrigen', 'Empresa Origen', 'Empresa'] if c in _df_fc.columns), None)

        col_tot_fc = next((c for c in ['TotalM', 'Total', 'SubTotal'] if c in _df_fc.columns), None)



        if col_sp_fc and col_doc_fc:

            df_fc_m = _df_fc.copy()

            df_fc_m['Folio_Match'] = df_fc_m[col_sp_fc].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()

            df_fc_m['DocumentID_FC'] = df_fc_m[col_doc_fc].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()

            if col_emp_fc:

                df_fc_m['Emp_FC'] = df_fc_m[col_emp_fc].astype(str).str.strip()

            if col_uuid_fc:

                df_fc_m['UUID_FC'] = df_fc_m[col_uuid_fc].astype(str).str.strip()

            if col_tot_fc:

                df_fc_m['Monto_FC'] = pd.to_numeric(df_fc_m[col_tot_fc], errors='coerce').fillna(0)



    # 1. SOLICITUDES DE PAGO (SP)

    rows_sp = []

    if _df_tes is not None and not _df_tes.empty:

        for _, row in _df_tes.iterrows():

            folio_sp = str(row.get('DocFolio', '')).replace('.0', '').strip().upper()

            emp = str(row.get('EmpresaOrigen', '')).strip()

            total_doc = pd.to_numeric(row.get('Total', 0), errors='coerce') or 0.0



            sub_items = []



            if not df_gs_m.empty:

                sub_gs = df_gs_m[(df_gs_m['Folio_Match'] == folio_sp) & (df_gs_m['Emp_GS'] == emp)] if 'Emp_GS' in df_gs_m.columns else df_gs_m[df_gs_m['Folio_Match'] == folio_sp]

                for _, r_gs in sub_gs.iterrows():

                    d_id = str(r_gs.get('DocumentID_GS', '')).strip()

                    u_id = str(r_gs.get('UUID_GS', '')).strip()

                    m_item = float(r_gs.get('Monto_GS', 0.0))

                    sub_items.append({'doc_id': d_id, 'uuid': u_id, 'monto': m_item})



            if not df_fc_m.empty:

                sub_fc = df_fc_m[(df_fc_m['Folio_Match'] == folio_sp) & (df_fc_m['Emp_FC'] == emp)] if 'Emp_FC' in df_fc_m.columns else df_fc_m[df_fc_m['Folio_Match'] == folio_sp]

                for _, r_fc in sub_fc.iterrows():

                    d_id = str(r_fc.get('DocumentID_FC', '')).strip()

                    u_id = str(r_fc.get('UUID_FC', '')).strip()

                    m_item = float(r_fc.get('Monto_FC', 0.0))

                    sub_items.append({'doc_id': d_id, 'uuid': u_id, 'monto': m_item})



            if not sub_items:

                doc_nat = str(row.get('DocumentID', '')).replace('.0', '').strip()

                sub_items.append({'doc_id': doc_nat, 'uuid': '', 'monto': total_doc})



            num_sub = len(sub_items)

            for item in sub_items:

                r_copy = row.to_dict()

                d_id = item['doc_id']

                m_pagado = pagos_edo_map.get((emp, d_id), 0.0) if d_id else 0.0



                r_copy['DocumentID'] = d_id

                r_copy['UUID'] = item['uuid'] if item['uuid'] != 'nan' else ''

                

                m_fact = item['monto'] if item['monto'] > 0 else (total_doc / num_sub)

                r_copy['Total'] = m_fact

                r_copy['Amount'] = m_pagado

                r_copy['SaldoPagoSP'] = max(0.0, m_fact - m_pagado)

                r_copy['Saldo_Pendiente'] = r_copy['SaldoPagoSP']

                r_copy['Tipo_Movimiento'] = 'Solicitud de Pago (SP)'

                rows_sp.append(r_copy)



    df_sp_calc = pd.DataFrame(rows_sp) if rows_sp else pd.DataFrame()



    # 2. ÓRDENES DE COMPRA (OC)

    rows_oc = []

    if _df_ord is not None and not _df_ord.empty:

        for _, row in _df_ord.iterrows():

            folio_oc = str(row.get('DocFolio', '')).replace('.0', '').strip().upper()

            emp = str(row.get('EmpresaOrigen', '')).strip()

            total_doc = pd.to_numeric(row.get('Total', 0), errors='coerce') or 0.0



            sub_items = []



            if not df_fc_m.empty:

                sub_fc = df_fc_m[(df_fc_m['Folio_Match'] == folio_oc) & (df_fc_m['Emp_FC'] == emp)] if 'Emp_FC' in df_fc_m.columns else df_fc_m[df_fc_m['Folio_Match'] == folio_oc]

                for _, r_fc in sub_fc.iterrows():

                    d_id = str(r_fc.get('DocumentID_FC', '')).strip()

                    u_id = str(r_fc.get('UUID_FC', '')).strip()

                    m_item = float(r_fc.get('Monto_FC', 0.0))

                    sub_items.append({'doc_id': d_id, 'uuid': u_id, 'monto': m_item})



            if not df_gs_m.empty:

                sub_gs = df_gs_m[(df_gs_m['Folio_Match'] == folio_oc) & (df_gs_m['Emp_GS'] == emp)] if 'Emp_GS' in df_gs_m.columns else df_gs_m[df_gs_m['Folio_Match'] == folio_oc]

                for _, r_gs in sub_gs.iterrows():

                    d_id = str(r_gs.get('DocumentID_GS', '')).strip()

                    u_id = str(r_gs.get('UUID_GS', '')).strip()

                    m_item = float(r_gs.get('Monto_GS', 0.0))

                    sub_items.append({'doc_id': d_id, 'uuid': u_id, 'monto': m_item})



            if not sub_items:

                doc_nat = str(row.get('DocumentID', '')).replace('.0', '').strip()

                sub_items.append({'doc_id': doc_nat, 'uuid': '', 'monto': total_doc})



            num_sub = len(sub_items)

            for item in sub_items:

                r_copy = row.to_dict()

                d_id = item['doc_id']

                m_pagado = pagos_edo_map.get((emp, d_id), 0.0) if d_id else 0.0



                r_copy['DocumentID'] = d_id

                r_copy['UUID'] = item['uuid'] if item['uuid'] != 'nan' else ''



                m_fact = item['monto'] if item['monto'] > 0 else (total_doc / num_sub)

                r_copy['Total'] = m_fact

                r_copy['Amount'] = m_pagado

                r_copy['SaldoPagoOC'] = max(0.0, m_fact - m_pagado)

                r_copy['Saldo_Pendiente'] = r_copy['SaldoPagoOC']

                r_copy['Tipo_Movimiento'] = 'Orden de Compra (OC)'

                rows_oc.append(r_copy)



    df_oc_calc = pd.DataFrame(rows_oc) if rows_oc else pd.DataFrame()



    return df_oc_calc, df_sp_calc



df_ordenes_proc, df_tesoreria_proc = procesar_oc_sp_definitivo(df_ordenes, df_tesoreria, df_fact_compra, df_gastos, df_edocuenta)



columnas_oc_visuales = [

    'EmpresaOrigen', 'DocFolio', 'DocumentID', 'BusinessEntityName', 'DateDocument', 'Title',

    'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'TotalDiscount', 'TotalTax',

    'TotalRetention', 'Total', 'UUID', 'Amount', 'SaldoPagoOC'

]



columnas_sp_visuales = [

    'EmpresaOrigen', 'DocFolio', 'DocumentID', 'BusinessEntityName', 'DateDocument', 'Title',

    'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'TotalDiscount', 'TotalTax',

    'TotalRetention', 'Total', 'UUID', 'Amount', 'SaldoPagoSP'

]



# --- RENDERIZADOR DE TABLA ESTILO ORACLE ENTERPRISE ---

def mostrar_tabla_con_totales(df_entrada, cols_num, cols_orden):

    if df_entrada.empty:

        st.info("No hay registros para mostrar.")

        return



    df_calc = df_entrada.copy()



    # Construir encabezados HTML

    cols_existentes = [c for c in cols_orden if c in df_calc.columns]

    

    html = ['<div class="oracle-table-container"><table class="oracle-table"><thead><tr>']

    for col in cols_existentes:

        html.append(f'<th>{col}</th>')

    html.append('</tr></thead><tbody>')



    # Filas de Datos

    for _, row in df_calc.iterrows():

        html.append('<tr>')

        for col in cols_existentes:

            val = row.get(col, '')

            if col in cols_num:

                val_fmt = formato_mx(val)

                html.append(f'<td class="text-right">{val_fmt}</td>')

            elif col in ['DocFolio', 'DocumentID', 'DateDocument', 'TIPO DOC', 'Currency', 'Moneda', 'Fecha Vencimiento', 'Tipo']:

                html.append(f'<td class="text-center">{val if not pd.isnull(val) else ""}</td>')

            else:

                html.append(f'<td class="text-left">{val if not pd.isnull(val) else ""}</td>')

        html.append('</tr>')



    # Fila de Totales

    html.append('<tr class="total-row">')

    for idx, col in enumerate(cols_existentes):

        if col in cols_num:

            tot_val = df_calc[col].sum() if col in df_calc.columns else 0.0

            html.append(f'<td class="text-right">{formato_mx(tot_val)}</td>')

        elif idx == 0:

            html.append('<td class="text-center">TOTALES</td>')

        else:

            html.append('<td></td>')

    html.append('</tr></tbody></table></div>')



    st.markdown("".join(html), unsafe_allow_html=True)



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



        subtotal_val = pd.t 

