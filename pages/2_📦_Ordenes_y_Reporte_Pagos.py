import streamlit as st
import pandas as pd
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="OC, SP y Reporte de Pagos - Grupo SERVYRE", layout="wide")

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

    borde_delgado = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    borde_total = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    row_idx = 1
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)
    cell = ws.cell(row=row_idx, column=1, value=f"REPORTE EJECUTIVO DE PAGOS — {empresa_nombre.upper()}")
    cell.font = font_titulo
    cell.fill = fill_titulo
    cell.alignment = Alignment(horizontal="center", vertical="center")
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
def cargar_datos_completos():
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

df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos = cargar_datos_completos()

st.title("📦 Módulo de Órdenes, Solicitudes y Reportes")
vista = st.radio("Seleccione vista:", ["Órdenes de Compra y Solicitudes de Pago", "Reporte Ejecutivo de Pagos"], horizontal=True)

columnas_oc = ['EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument', 'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'Total', 'UUID', 'Saldo_Pendiente']
columnas_sp = ['EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument', 'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal', 'Total', 'UUID', 'Saldo_Pendiente']

if vista == "Órdenes de Compra y Solicitudes de Pago":
    st.info("Visualización de Órdenes de Compra y Solicitudes de Pago con filtros independientes.")
    tab_oc, tab_sp = st.tabs(["OrdenCompra", "SolicitudPago"])
    with tab_oc:
        if df_ordenes is not None and not df_ordenes.empty:
            st.dataframe(df_ordenes, use_container_width=True)
    with tab_sp:
        if df_tesoreria is not None and not df_tesoreria.empty:
            st.dataframe(df_tesoreria, use_container_width=True)
else:
    st.markdown("### 📋 Reporte Ejecutivo de Pagos con Saldo Pendiente")
    # Lógica del Reporte de Pagos ejecutiva