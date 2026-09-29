import streamlit as st
import pandas as pd
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Finanzas - Grupo SERVYRE", layout="wide")

# ==========================================
# 🎨 ESTILOS CSS PERSONALIZADOS (ESTILO NARANJA CORPORATIVO)
# ==========================================
st.markdown("""
    <style>
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
            background-color: #fde9d9;
            color: #833C0C;
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
            background-color: #fef2ed;
        }
        .oracle-table tr.total-row {
            background-color: #fce4d6 !important;
            font-weight: bold;
            color: #000000;
            border-top: 2px solid #c65911;
            border-bottom: 2px solid #c65911;
        }
        .oracle-table tr.total-row td {
            border: 1px solid #f8cbad;
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
    fill_titulo = PatternFill(start_color="C65911", end_color="C65911", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="D66011", end_color="D66011", fill_type="solid")
    font_normal = Font(name="Calibri", size=10, color="000000")
    font_nc = Font(name="Calibri", size=10, color="9C0006")
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_total = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")

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
        ws.cell(row=row_idx, column=3, value=str(r.get('DocFolio', r.get('Folio', '')))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=4, value=str(r.get('UUID', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=5, value=str(r.get('TIPO DOC', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=6, value=str(r.get('CFDStatusCancelledName', ''))).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=7, value=str(r.get('BusinessEntityName', ''))).alignment = Alignment(horizontal="left")
        
        cols_indices = [
            (8, 'TotalRetention'), (9, 'SubTotal'), (10, 'TotalDiscount'), (11, 'Subtotal2'),
            (12, 'TotalTax'), (13, 'Total'), (16, 'Amount'), (18, 'SaldoFactura')
        ]
        
        for c_idx, col_name in cols_indices:
            val = float(r.get(col_name, 0)) if col_name in r else 0.0
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
        8: df_unicos['TotalRetention'].sum() if 'TotalRetention' in df_unicos.columns else 0.0,
        9: df_unicos['SubTotal'].sum() if 'SubTotal' in df_unicos.columns else 0.0,
        10: df_unicos['TotalDiscount'].sum() if 'TotalDiscount' in df_unicos.columns else 0.0,
        11: df_unicos['Subtotal2'].sum() if 'Subtotal2' in df_unicos.columns else 0.0,
        12: df_unicos['TotalTax'].sum() if 'TotalTax' in df_unicos.columns else 0.0,
        13: df_unicos['Total'].sum() if 'Total' in df_unicos.columns else 0.0,
        16: df_datos['Amount'].sum() if 'Amount' in df_datos.columns else 0.0,
        18: df_unicos['SaldoFactura'].sum() if 'SaldoFactura' in df_unicos.columns else 0.0
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
            # Garantizar que exista DocFolio
            if 'DocFolio' not in df_fac.columns and 'Folio' in df_fac.columns:
                df_fac['DocFolio'] = df_fac['Folio']

        df_nc = pd.read_excel(path, sheet_name='NotaCreditoCliente') if 'NotaCreditoCliente' in sheets else pd.DataFrame()
        df_nc = filtrar_no_eliminados(df_nc)
        if df_nc is not None and not df_nc.empty:
            df_nc['TIPO DOC'] = 'NC'
            if 'DocFolio' not in df_nc.columns and 'Folio' in df_nc.columns:
                df_nc['DocFolio'] = df_nc['Folio']

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

# --- PROCESAMIENTO FACTURACIÓN ---
@st.cache_data
def procesar_facturacion_definitivo(_df_fac, _df_edo):
    if _df_fac is None or _df_fac.empty:
        return pd.DataFrame()
    
    df_f = _df_fac.copy()
    if 'DocFolio' not in df_f.columns and 'Folio' in df_f.columns:
        df_f['DocFolio'] = df_f['Folio']

    sub_v = pd.to_numeric(df_f['SubTotal'], errors='coerce').fillna(0) if 'SubTotal' in df_f.columns else 0.0
    desc_v = pd.to_numeric(df_f['TotalDiscount'], errors='coerce').fillna(0) if 'TotalDiscount' in df_f.columns else 0.0
    df_f['Subtotal2'] = sub_v - desc_v

    if _df_edo is not None and not _df_edo.empty and 'DocumentID' in df_f.columns and 'DocumentID' in _df_edo.columns:
        cols_edo = ['DocumentID', 'Amount', 'DateOperation']
        if 'EmpresaOrigen' in _df_edo.columns and 'EmpresaOrigen' in df_f.columns:
            cols_edo.append('EmpresaOrigen')

        df_edo_sub = _df_edo[[c for c in cols_edo if c in _df_edo.columns]].copy()
        
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
    return df_f

df_factura_proc = procesar_facturacion_definitivo(df_factura, df_edocuenta)

# --- RENDERIZADOR DE TABLA CON FILTROS INTERACTIVOS Y KPIs ---
def mostrar_tabla_filtrada_con_totales(df_entrada, cols_num, cols_orden, clave_prefijo, campo_saldo):
    if df_entrada.empty:
        st.info("No hay registros para mostrar.")
        return

    df_calc = df_entrada.copy()
    if 'DocFolio' not in df_calc.columns and 'Folio' in df_calc.columns:
        df_calc['DocFolio'] = df_calc['Folio']

    # --- FILTROS DE BÚSQUEDA INTERACTIVOS ---
    st.markdown("#### ⚙️ Filtros de Búsqueda")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        empresas = sorted(df_calc['EmpresaOrigen'].dropna().unique()) if 'EmpresaOrigen' in df_calc.columns else []
        sel_emp = st.multiselect("Filtrar por Empresa:", empresas, key=f"{clave_prefijo}_emp")
        if sel_emp: df_calc = df_calc[df_calc['EmpresaOrigen'].isin(sel_emp)]

    with c2:
        provs = sorted(df_calc['BusinessEntityName'].dropna().unique()) if 'BusinessEntityName' in df_calc.columns else []
        sel_prov = st.multiselect("Filtrar por Proveedor / Cliente:", provs, key=f"{clave_prefijo}_prov")
        if sel_prov: df_calc = df_calc[df_calc['BusinessEntityName'].isin(sel_prov)]

    with c3:
        folios = sorted(df_calc['DocFolio'].dropna().unique()) if 'DocFolio' in df_calc.columns else []
        sel_folio = st.multiselect("Filtrar por Folio:", folios, key=f"{clave_prefijo}_folio")
        if sel_folio: df_calc = df_calc[df_calc['DocFolio'].isin(sel_folio)]

    st.markdown("---")

    # --- KPIs SUPERIORES DE LA TABLA ---
    total_docs = len(df_calc)
    suma_saldo = df_calc[campo_saldo].sum() if campo_saldo in df_calc.columns else 0.0

    kpi_1, kpi_2 = st.columns(2)
    kpi_1.metric("Documentos Filtrados", f"{total_docs:,}")
    kpi_2.metric("Saldo Pendiente Filtrado", formato_mx(suma_saldo))
    st.markdown("---")

    cols_existentes = [c for c in cols_orden if c in df_calc.columns]
    
    html = ['<div class="oracle-table-container"><table class="oracle-table"><thead><tr>']
    for col in cols_existentes:
        html.append(f'<th>{col}</th>')
    html.append('</tr></thead><tbody>')

    for _, row in df_calc.iterrows():
        html.append('<tr>')
        for col in cols_existentes:
            val = row.get(col, '')
            if col in cols_num:
                val_fmt = formato_mx(val)
                html.append(f'<td class="text-right">{val_fmt}</td>')
            elif col in ['DocFolio', 'DocumentID', 'DateDocument', 'TIPO DOC', 'Currency', 'Moneda']:
                html.append(f'<td class="text-center">{val if not pd.isnull(val) else ""}</td>')
            else:
                html.append(f'<td class="text-left">{val if not pd.isnull(val) else ""}</td>')
        html.append('</tr>')

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
    st.markdown("Muestra exclusivamente las facturas que tienen un saldo pendiente real mayor a $1.00")

    if df_factura_proc is not None and not df_factura_proc.empty:
        df_f = df_factura_proc[df_factura_proc['SaldoFactura'] > 1.0].copy()

        if not df_f.empty:
            excel_data = generar_excel_facturacion_ejecutivo(df_f)
            st.download_button(
                label="📥 Descargar Reporte de Facturación en Excel",
                data=excel_data,
                file_name="Reporte_Facturacion_Servyre.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            cols_num_fac = ['TotalRetention', 'SubTotal', 'TotalDiscount', 'Subtotal2', 'TotalTax', 'Total', 'Amount', 'SaldoFactura']
            mostrar_tabla_filtrada_con_totales(df_f, cols_num_fac, columnas_requeridas, "fac", "SaldoFactura")
        else:
            st.info("No hay facturas con saldo pendiente mayor a $1.00.")
    else:
        st.warning("No hay datos disponibles en Facturación.")