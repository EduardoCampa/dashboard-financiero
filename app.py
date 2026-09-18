import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuración inicial del dashboard en ancho completo
st.set_page_config(page_title="Dashboard Financiero - Contpaq", layout="wide")

# --- REGLA FIJA ESTRICTA PARA FORMATO DE MONEDA REGIÓN MÉXICO ($1,234,567.89) ---
def formato_mx(val):
    """
    Formato contable estándar para México:
    Fuerza el formato a $1,234,567.89 (Comas para miles, punto para decimales, símbolo de pesos).
    """
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

# --- FUNCIÓN PARA GENERAR EXCEL CON FORMATO MONETARIO DE MÉXICO ---
def generar_excel_ejecutivo(df_datos, empresa_nombre):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte de Pagos"
    
    ws.views.sheetView[0].showGridLines = True

    font_titulo_empresa = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    fill_titulo_empresa = PatternFill(start_color="B25A2B", end_color="B25A2B", fill_type="solid")
    
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="B25A2B", end_color="B25A2B", fill_type="solid")
    
    font_prov = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_prov = PatternFill(start_color="F2D9D0", end_color="F2D9D0", fill_type="solid")
    
    font_normal = Font(name="Calibri", size=10, color="000000")
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_total = PatternFill(start_color="F2D9D0", end_color="F2D9D0", fill_type="solid")

    borde_delgado = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    borde_total = Border(
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )

    row_idx = 1
    
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
    cell = ws.cell(row=row_idx, column=1, value=f"REPORTE DE PAGOS — {empresa_nombre.upper()}")
    cell.font = font_titulo_empresa
    cell.fill = fill_titulo_empresa
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_idx].height = 30
    row_idx += 2

    headers = ["TIPO", "PROVEEDOR", "FECHA VENCIMIENTO", "FOLIO / DOCUMENTO", "MONEDA", "DESCRIPCIÓN", "SALDO PENDIENTE"]
    proveedores = df_datos['BusinessEntityName'].unique() if 'BusinessEntityName' in df_datos.columns else []

    for prov in sorted(proveedores):
        df_prov = df_datos[df_datos['BusinessEntityName'] == prov]
        total_prov = df_prov['Saldo_Pendiente'].sum()

        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
        cell_prov = ws.cell(row=row_idx, column=1, value=str(prov))
        cell_prov.font = font_prov
        cell_prov.fill = fill_prov
        cell_prov.alignment = Alignment(horizontal="left", vertical="center")
        
        cell_prov_tot = ws.cell(row=row_idx, column=7, value=total_prov)
        cell_prov_tot.font = font_prov
        cell_prov_tot.fill = fill_prov
        cell_prov_tot.number_format = '"$"#,##0.00'
        cell_prov_tot.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[row_idx].height = 22
        row_idx += 1

        for col_num, h in enumerate(headers, 1):
            c = ws.cell(row=row_idx, column=col_num, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = borde_delgado
        ws.row_dimensions[row_idx].height = 20
        row_idx += 1

        for _, r in df_prov.iterrows():
            ws.cell(row=row_idx, column=1, value=str(r.get('Tipo_Movimiento', ''))).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=2, value=str(r.get('BusinessEntityName', ''))).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=3, value=str(r.get('Fecha_Fmt', ''))).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=4, value=str(r.get('DocFolio', ''))).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=5, value=str(r.get('Currency', 'MXN'))).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=6, value=str(r.get('Title', ''))).alignment = Alignment(horizontal="left")
            
            c_val = ws.cell(row=row_idx, column=7, value=float(r.get('Saldo_Pendiente', 0)))
            c_val.number_format = '"$"#,##0.00'
            c_val.alignment = Alignment(horizontal="right")

            for col_num in range(1, 8):
                cell_det = ws.cell(row=row_idx, column=col_num)
                cell_det.font = font_normal
                cell_det.border = borde_delgado
            
            ws.row_dimensions[row_idx].height = 18
            row_idx += 1

        row_idx += 1

    total_general = df_datos['Saldo_Pendiente'].sum()
    ws.cell(row=row_idx, column=1, value="TOTAL GENERAL").font = font_total
    for col_num in range(1, 7):
        ws.cell(row=row_idx, column=col_num).fill = fill_total
        ws.cell(row=row_idx, column=col_num).border = borde_total
        
    cell_tot_gen = ws.cell(row=row_idx, column=7, value=total_general)
    cell_tot_gen.font = font_total
    cell_tot_gen.fill = fill_total
    cell_tot_gen.number_format = '"$"#,##0.00'
    cell_tot_gen.border = borde_total
    cell_tot_gen.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[row_idx].height = 24

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- BARRA LATERAL (MENÚ LATERAL ORGANIZADO) ---
st.sidebar.title("Sistema Auditoría SAT")
st.sidebar.markdown("---")

area_principal = st.sidebar.radio("Área Principal", ["Finanzas", "Contabilidad"], key="area_ppal")

menu = ""
menu_contabilidad = ""

if area_principal == "Finanzas":
    st.sidebar.markdown("### Módulos Financieros")
    menu = st.sidebar.radio(
        "Seleccione módulo:", 
        ["Panel", "Facturación", "OC y SP", "Reporte Pagos"],
        key="menu_finanzas_sel"
    )
else:
    st.sidebar.markdown("### Módulos Contables")
    menu_contabilidad = st.sidebar.radio(
        "Seleccione módulo:",
        ["Polizas", "Balanza de Comprobación", "Conciliación Bancaria"],
        key="menu_contabilidad_sel"
    )

# Ruta de tu archivo maestro consolidado en tu equipo
ruta_archivo = r"C:\Users\User\Downloads\PYTHON\Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos(path):
    try:
        df_factura = pd.read_excel(path, sheet_name='FacturaCliente')
        df_tesoreria = pd.read_excel(path, sheet_name='SolicitudPago')
        df_ordenes = pd.read_excel(path, sheet_name='OrdenCompra')
        df_edocuenta = pd.read_excel(path, sheet_name='EdoCuenta')
        
        df_fact_compra = pd.read_excel(path, sheet_name='FacturaCompra') if 'FacturaCompra' in pd.ExcelFile(path).sheet_names else None
        df_gastos = pd.read_excel(path, sheet_name='Gastos') if 'Gastos' in pd.ExcelFile(path).sheet_names else None
        
        return df_factura, df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos
    except Exception as e:
        st.error(f"Error al cargar las pestañas del archivo: {e}")
        return None, None, None, None, None, None

df_factura, df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos = cargar_datos(ruta_archivo)

# Columnas para OrdenCompra
columnas_oc = [
    'EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument',
    'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal',
    'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total', 'TotalFacturaCompra', 'SaldoOC', 'DocumentID', 'UUID', 'Amount', 'DateOperation', 'SaldoPagoOC'
]

# Columnas para SolicitudPago
columnas_sp = [
    'EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument',
    'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal',
    'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total', 'TotalGastos', 'SaldoSP', 'DocumentID', 'CFDIFolioFiscal', 'Amount', 'DateOperation', 'SaldoPagoSP'
]

# --- RENDERIZADO SEGÚN LA SELECCIÓN ---

if area_principal == "Contabilidad":
    st.title("📚 Módulo de Contabilidad")
    st.info(f"El módulo de **{menu_contabilidad}** se encuentra actualmente en desarrollo.")

elif menu == "Panel":
    st.title("Testing Solutions S.A.")
    st.markdown("### 📊 Panel Ejecutivo y Consolidado General (Ingresos, Órdenes de Compra y Gastos)")

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
                if not df_t.empty:
                    df_t['Año'] = 0
                    df_t['Periodo'] = "Sin Periodo"

        st.markdown("### 📅 Filtros Generales (Empresa Origen, Año y Periodo)")
        
        empresas_set = set()
        for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
            if not df_t.empty and 'EmpresaOrigen' in df_t.columns:
                empresas_set.update(df_t['EmpresaOrigen'].dropna().unique())
        lista_empresas_dash = ["Todas"] + sorted(list(empresas_set))

        col_d0, col_d1, col_d2 = st.columns(3)
        with col_d0:
            empresa_dash_sel = st.selectbox("Seleccione la Empresa Origen:", lista_empresas_dash, key="dash_empresa_sel")
        
        anios_set = set()
        for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
            if not df_t.empty and 'Año' in df_t.columns:
                anios_set.update(df_t['Año'].unique())
        anios_dash_validos = ["Todos"] + sorted([int(a) for a in anios_set if a > 0], reverse=True)
        
        with col_d1:
            anio_dash_sel = st.selectbox("Seleccione el Año fiscal:", anios_dash_validos, key="dash_anio_sel")
        
        periodos_set = set()
        for df_t in [df_dash_fact, df_dash_tes, df_dash_oc]:
            if not df_t.empty and 'Periodo' in df_t.columns:
                periodos_set.update(df_t['Periodo'].unique())
        periodos_dash_validos = ["Todos"] + [p for p in sorted(list(periodos_set)) if p != "Sin Periodo"]
        
        with col_d2:
            periodo_dash_sel = st.selectbox("Seleccione el Periodo (Mes):", periodos_dash_validos, key="dash_periodo_sel")

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
        
        egresos_totales = total_oc + total_gastos_sp
        utilidad_neta = total_ingresos - egresos_totales

        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric(label="INGRESOS (Facturación)", value=formato_mx(total_ingresos), delta="Clientes")
        with col2: st.metric(label="EGRESOS (Órdenes Compra)", value=formato_mx(total_oc), delta="Proveedores / OC")
        with col3: st.metric(label="GASTOS / PAGOS (Tesorería)", value=formato_mx(total_gastos_sp), delta="Solicitudes")
        with col4: 
            st.metric(label="UTILIDAD NETA", value=formato_mx(utilidad_neta), delta="Margen calculado")
            st.progress(0.83)

        st.markdown("---")
        st.markdown("### 📈 Gráficas Financieras por Empresa Origen")

        empresas_grafica = sorted(list(set(
            list(df_dash_fact['EmpresaOrigen'].dropna().unique() if not df_dash_fact.empty and 'EmpresaOrigen' in df_dash_fact.columns else []) +
            list(df_dash_oc['EmpresaOrigen'].dropna().unique() if not df_dash_oc.empty and 'EmpresaOrigen' in df_dash_oc.columns else []) +
            list(df_dash_tes['EmpresaOrigen'].dropna().unique() if not df_dash_tes.empty and 'EmpresaOrigen' in df_dash_tes.columns else [])
        )))

        if empresas_grafica:
            data_grafica = []
            for emp in empresas_grafica:
                ing_emp = pd.to_numeric(df_dash_fact[df_dash_fact['EmpresaOrigen'] == emp]['Total'], errors='coerce').fillna(0).sum() if not df_dash_fact.empty and 'EmpresaOrigen' in df_dash_fact.columns else 0
                oc_emp = pd.to_numeric(df_dash_oc[df_dash_oc['EmpresaOrigen'] == emp]['Total'], errors='coerce').fillna(0).sum() if not df_dash_oc.empty and 'EmpresaOrigen' in df_dash_oc.columns else 0
                sp_emp = pd.to_numeric(df_dash_tes[df_dash_tes['EmpresaOrigen'] == emp]['Total'], errors='coerce').fillna(0).sum() if not df_dash_tes.empty and 'EmpresaOrigen' in df_dash_tes.columns else 0
                egr_emp = oc_emp + sp_emp
                ut_emp = ing_emp - egr_emp
                data_grafica.append({'EmpresaOrigen': emp, 'Ingresos': ing_emp, 'Egresos (OC + SP)': egr_emp, 'Utilidad': ut_emp})

            df_graf = pd.DataFrame(data_grafica).set_index('EmpresaOrigen')

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.subheader("Ingresos vs Egresos por Empresa")
                st.bar_chart(df_graf[['Ingresos', 'Egresos (OC + SP)']])
            with col_g2:
                st.subheader("Utilidad Neta por Empresa")
                st.bar_chart(df_graf['Utilidad'])
        else:
            st.info("No hay datos suficientes para mostrar las gráficas con los filtros seleccionados.")

elif menu == "Facturación":
    st.title("📊 Módulo de Facturación y Dashboard de Cobranza")
    if df_factura is not None:
        if 'DateDocument' in df_factura.columns:
            df_factura['DateDocument'] = pd.to_datetime(df_factura['DateDocument'], errors='coerce')
            df_factura['Año'] = df_factura['DateDocument'].dt.year.fillna(0).astype(int)
        else:
            df_factura['Año'] = 0

        st.markdown("### 📅 Filtrar por Periodo y Empresa")
        empresas_disponibles = sorted(df_factura['EmpresaOrigen'].dropna().unique()) if 'EmpresaOrigen' in df_factura.columns else []
        
        col_f_emp, col_f_anio = st.columns(2)
        with col_f_emp:
            empresa_sel_fact = st.selectbox("Seleccione la Empresa Origen:", ["Todas"] + empresas_disponibles, key="select_empresa_fact")
        with col_f_anio:
            anios_disponibles = sorted([int(a) for a in df_factura['Año'].unique() if a > 0], reverse=True)
            anio_seleccionado = st.selectbox("Seleccione el Año fiscal:", ["Todos"] + anios_disponibles)

        df_fact_anio = df_factura.copy()
        if empresa_sel_fact != "Todas":
            df_fact_anio = df_fact_anio[df_fact_anio['EmpresaOrigen'] == empresa_sel_fact]
            
        if anio_seleccionado != "Todos":
            df_fact_anio = df_fact_anio[df_fact_anio['Año'] == int(anio_seleccionado)]

        if df_edocuenta is not None and 'DocumentID' in df_fact_anio.columns and 'DocumentID' in df_edocuenta.columns:
            cols_edo = ['EmpresaOrigen', 'DocumentID', 'DateOperation', 'Amount', 'FinancialEntityName'] if 'EmpresaOrigen' in df_edocuenta.columns else ['DocumentID', 'DateOperation', 'Amount', 'FinancialEntityName']
            df_edo_subset = df_edocuenta[[c for c in cols_edo if c in df_edocuenta.columns]].copy()
            if 'Amount' in df_edo_subset.columns:
                df_edo_subset['Amount'] = pd.to_numeric(df_edo_subset['Amount'], errors='coerce').fillna(0)
                df_edo_subset = df_edo_subset.rename(columns={'Amount': 'Credit'})

            if 'EmpresaOrigen' in df_edo_subset.columns and 'EmpresaOrigen' in df_fact_anio.columns:
                df_unido = pd.merge(df_fact_anio, df_edo_subset, on=['EmpresaOrigen', 'DocumentID'], how='left')
            else:
                df_unido = pd.merge(df_fact_anio, df_edo_subset, on='DocumentID', how='left')
        else:
            df_unido = df_fact_anio.copy()
            for col in ['DateOperation', 'Credit', 'FinancialEntityName']:
                if col not in df_unido.columns: df_unido[col] = None

        group_keys_fact = ['EmpresaOrigen', 'DocumentID'] if 'EmpresaOrigen' in df_unido.columns else ['DocumentID']
        if 'Total' in df_unido.columns and 'Credit' in df_unido.columns:
            df_unido['Credit'] = pd.to_numeric(df_unido['Credit'], errors='coerce').fillna(0)
            pagos_totales = df_unido.groupby(group_keys_fact)['Credit'].transform('sum')
            df_unido['TotalPagado'] = pagos_totales
            df_unido['SaldoFactura'] = pd.to_numeric(df_unido['Total'], errors='coerce').fillna(0) - df_unido['TotalPagado']
        else:
            df_unido['SaldoFactura'] = pd.to_numeric(df_unido.get('Total', 0), errors='coerce').fillna(0)

        st.markdown("---")
        col_f1, col_f2 = st.columns(2)
        with col_f1: texto_busqueda = st.text_input("Buscar por Folio, Cliente o RFC:", "")
        with col_f2:
            if 'DateDocument' in df_unido.columns and not df_unido['DateDocument'].dropna().empty:
                min_f = df_unido['DateDocument'].min().date()
                max_f = df_unido['DateDocument'].max().date()
                rango_fechas = st.date_input("Rango exacto de fechas:", value=(min_f, max_f))
            else:
                rango_fechas = None

        columnas_deseadas = [
            'EmpresaOrigen', 'DocumentID', 'RFC', 'BusinessEntityName', 'DateDocument',
            'Title', 'Currency', 'Rate', 'SubTotal', 'TotalTax', 'Total',
            'CostCenterName', 'UUID', 'CFDStatusCancelledName',
            'DateOperation', 'Credit', 'FinancialEntityName', 'SaldoFactura'
        ]
        df_filtrado = df_unido[[col for col in columnas_deseadas if col in df_unido.columns]].copy()

        if texto_busqueda:
            filtro_mask = df_filtrado.astype(str).apply(lambda col: col.str.contains(texto_busqueda, case=False, na=False)).any(axis=1)
            df_filtrado = df_filtrado[filtro_mask]

        if rango_fechas and len(rango_fechas) == 2 and 'DateDocument' in df_filtrado.columns:
            inicio, fin = rango_fechas
            df_filtrado = df_filtrado[(df_filtrado['DateDocument'].dt.date >= inicio) & (df_filtrado['DateDocument'].dt.date <= fin)]

        df_unicas_filtrado = df_filtrado.drop_duplicates(subset=group_keys_fact)
        total_facturado_per = pd.to_numeric(df_unicas_filtrado['Total'], errors='coerce').fillna(0).sum() if 'Total' in df_unicas_filtrado.columns else 0.0
        total_cobrado_per = pd.to_numeric(df_filtrado['Credit'], errors='coerce').fillna(0).sum() if 'Credit' in df_filtrado.columns else 0.0
        total_saldo_per = pd.to_numeric(df_unicas_filtrado['SaldoFactura'], errors='coerce').fillna(0).sum() if 'SaldoFactura' in df_unicas_filtrado.columns else 0.0

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        with kpi1: st.metric("Total Facturado", formato_mx(total_facturado_per))
        with kpi2: st.metric("Total Cobrado / Pagos", formato_mx(total_cobrado_per))
        with kpi3: st.metric("Saldo Pendiente", formato_mx(max(0, total_saldo_per)))
        with kpi4: st.metric("Total Facturas", f"{len(df_unicas_filtrado):,}")

        st.markdown("---")

        df_view = df_filtrado.copy()
        for col in ['SubTotal', 'TotalTax', 'Total', 'Credit', 'SaldoFactura', 'Rate']:
            if col in df_view.columns:
                df_view[col] = pd.to_numeric(df_view[col], errors='coerce').apply(formato_mx)

        st.dataframe(df_view, use_container_width=True)

elif menu == "OC y SP":
    st.title("📦 Módulo de Órdenes de Compra y Solicitudes de Pago")
    
    st.markdown("### 🔎 Filtros de Búsqueda, Periodo y Empresa")
    
    empresas_oc_sp = set()
    for df_temp in [df_ordenes, df_tesoreria]:
        if df_temp is not None and 'EmpresaOrigen' in df_temp.columns:
            empresas_oc_sp.update(df_temp['EmpresaOrigen'].dropna().unique())
    lista_empresas_oc_sp = sorted(list(empresas_oc_sp))

    col_oc0, col_oc1, col_oc2, col_oc3 = st.columns(4)
    with col_oc0:
        empresa_sel_oc_sp = st.selectbox("Empresa Origen:", ["Todas"] + lista_empresas_oc_sp, key="select_empresa_oc_sp")
    with col_oc1:
        texto_busqueda_oc = st.text_input("Buscar por Proveedor, Folio o Título:", "", key="search_oc_sp")
    with col_oc2:
        anios_oc_sp = set()
        for df_temp in [df_ordenes, df_tesoreria]:
            if df_temp is not None and 'DateDocument' in df_temp.columns:
                anios_oc_sp.update(pd.to_datetime(df_temp['DateDocument'], errors='coerce').dt.year.dropna().unique())
        anio_oc_sel = st.selectbox("Seleccione el Año fiscal:", ["Todos"] + sorted([int(a) for a in anios_oc_sp if a > 0], reverse=True), key="select_anio_oc_sp")
    with col_oc3:
        filtro_estado_saldo = st.selectbox(
            "Filtrar por Saldo de Pago:", 
            ["Todos", "Con saldo pendiente (> $0)", "Liquidados / En cero ($0)"],
            key="filtro_estado_saldo_key"
        )

    st.markdown("---")
    tab_oc, tab_sp = st.tabs(["OrdenCompra", "SolicitudPago"])
    
    # --- PESTAÑA ORDENCOMPRA ---
    with tab_oc:
        st.subheader("Registros de OrdenCompra")
        if df_ordenes is not None:
            df_oc_base = df_ordenes.copy()
            
            if empresa_sel_oc_sp != "Todas" and 'EmpresaOrigen' in df_oc_base.columns:
                df_oc_base = df_oc_base[df_oc_base['EmpresaOrigen'] == empresa_sel_oc_sp]
            
            if df_fact_compra is not None:
                col_uuid_fc = 'UUID' if 'UUID' in df_fact_compra.columns else ('CFDIFolioFiscal' if 'CFDIFolioFiscal' in df_fact_compra.columns else None)
                col_total_fc = 'Total' if 'Total' in df_fact_compra.columns else None
                col_docid_fc = 'DocumentID' if 'DocumentID' in df_fact_compra.columns else None
                
                col_llave_fc = None
                for cand in ['Solicitud de Pago', 'SolicitudPago', 'DocFolio']:
                    if cand in df_fact_compra.columns:
                        col_llave_fc = cand
                        break
                
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
                    
                    if 'DocumentID' in df_oc_base.columns:
                        df_oc_base = df_oc_base.drop(columns=['DocumentID'])
                        
                    if 'EmpresaOrigen' in df_fc_sub.columns and 'EmpresaOrigen' in df_oc_base.columns:
                        df_oc_base = pd.merge(df_oc_base, df_fc_sub, left_on=['EmpresaOrigen', 'DocFolio'], right_on=['EmpresaOrigen', 'DocFolio_Match'], how='left')
                    else:
                        df_oc_base = pd.merge(df_oc_base, df_fc_sub, left_on='DocFolio', right_on='DocFolio_Match', how='left')
                    
                    if 'DocumentID_FC' in df_oc_base.columns:
                        df_oc_base['DocumentID'] = df_oc_base['DocumentID_FC']
                else:
                    df_oc_base['UUID'] = ""
                    df_oc_base['TotalFacturaCompra'] = 0.0
                    df_oc_base['DocumentID'] = ""
            else:
                df_oc_base['UUID'] = ""
                df_oc_base['TotalFacturaCompra'] = 0.0
                df_oc_base['DocumentID'] = ""

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
                df_oc_base['DateOperation'] = ""

            if 'Amount' not in df_oc_base.columns: df_oc_base['Amount'] = 0.0
            if 'DateOperation' not in df_oc_base.columns: df_oc_base['DateOperation'] = ""

            df_oc_base['Amount'] = pd.to_numeric(df_oc_base['Amount'], errors='coerce').fillna(0)
            df_oc_base['Total'] = pd.to_numeric(df_oc_base['Total'], errors='coerce').fillna(0)
            df_oc_base['TotalFacturaCompra'] = pd.to_numeric(df_oc_base['TotalFacturaCompra'], errors='coerce').fillna(0)
            
            group_keys_oc = ['EmpresaOrigen', 'DocFolio'] if 'EmpresaOrigen' in df_oc_base.columns else ['DocFolio']
            acumulado_fc = df_oc_base.groupby(group_keys_oc)['TotalFacturaCompra'].cumsum()
            df_oc_base['SaldoOC'] = df_oc_base['Total'] - acumulado_fc
            df_oc_base['SaldoOC'] = df_oc_base['SaldoOC'].apply(lambda x: max(0.0, x))
            
            total_pagado_oc = df_oc_base.groupby(group_keys_oc + ['DocumentID'])['Amount'].transform('sum')
            df_oc_base['SaldoPagoOC'] = df_oc_base['TotalFacturaCompra'] - total_pagado_oc
            df_oc_base['SaldoPagoOC'] = df_oc_base['SaldoPagoOC'].apply(lambda x: max(0.0, x))

            if anio_oc_sel != "Todos" and 'DateDocument' in df_oc_base.columns:
                df_oc_base['DateDocument_dt'] = pd.to_datetime(df_oc_base['DateDocument'], errors='coerce')
                df_oc_base = df_oc_base[df_oc_base['DateDocument_dt'].dt.year == int(anio_oc_sel)]

            if texto_busqueda_oc:
                cols_buscar = [c for c in ['DocFolio', 'BusinessEntityName', 'Title'] if c in df_oc_base.columns]
                if cols_buscar:
                    mask_oc = df_oc_base[cols_buscar].astype(str).apply(lambda col: col.str.contains(texto_busqueda_oc, case=False, na=False)).any(axis=1)
                    df_oc_base = df_oc_base[mask_oc]

            total_oc_val = df_oc_base.drop_duplicates(subset=group_keys_oc)['Total'].sum()
            total_fact_compra_val = df_oc_base.drop_duplicates(subset=group_keys_oc + ['UUID', 'TotalFacturaCompra'])['TotalFacturaCompra'].sum()
            total_amount_oc_val = df_oc_base['Amount'].sum()
            
            cols_grupo_saldo = group_keys_oc + ['DocumentID'] if 'DocumentID' in df_oc_base.columns else group_keys_oc
            total_saldo_pago_oc_val = df_oc_base.drop_duplicates(subset=cols_grupo_saldo)['SaldoPagoOC'].sum()

            kpi_oc1, kpi_oc2, kpi_oc3, kpi_oc4 = st.columns(4)
            with kpi_oc1: st.metric("Total Órdenes", formato_mx(total_oc_val))
            with kpi_oc2: st.metric("Total Facturas Compra", formato_mx(total_fact_compra_val))
            with kpi_oc3: st.metric("Total Pagado (Amount)", formato_mx(total_amount_oc_val))
            with kpi_oc4: st.metric("Saldo Pendiente Pago", formato_mx(total_saldo_pago_oc_val))
            st.markdown("---")

            df_oc_filtered = df_oc_base.copy()
            if filtro_estado_saldo == "Con saldo pendiente (> $0)":
                df_oc_filtered = df_oc_filtered[df_oc_filtered['SaldoPagoOC'] > 0]
            elif filtro_estado_saldo == "Liquidados / En cero ($0)":
                df_oc_filtered = df_oc_filtered[df_oc_filtered['SaldoPagoOC'] <= 0]

            df_oc_view = df_oc_filtered[[c for c in columnas_oc if c in df_oc_filtered.columns]].copy()
            for col_m in ['SubTotal', 'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total', 'TotalFacturaCompra', 'SaldoOC', 'Rate', 'Amount', 'SaldoPagoOC']:
                if col_m in df_oc_view.columns:
                    df_oc_view[col_m] = pd.to_numeric(df_oc_view[col_m], errors='coerce').apply(formato_mx)

            st.dataframe(df_oc_view, use_container_width=True)
        else:
            st.warning("No se encontró información de OrdenCompra.")
            
    # --- PESTAÑA SOLICITUDPAGO ---
    with tab_sp:
        st.subheader("Registros de SolicitudPago")
        if df_tesoreria is not None:
            df_sp_base = df_tesoreria.copy()
            
            if empresa_sel_oc_sp != "Todas" and 'EmpresaOrigen' in df_sp_base.columns:
                df_sp_base = df_sp_base[df_sp_base['EmpresaOrigen'] == empresa_sel_oc_sp]
            
            if df_gastos is not None:
                col_folio_gs = None
                for candidate in ['CFDIFolioFiscal', 'UUID', 'FolioFiscal']:
                    if candidate in df_gastos.columns:
                        col_folio_gs = candidate
                        break
                
                col_total_gs = 'Total' if 'Total' in df_gastos.columns else None
                col_docid_gs = 'DocumentID' if 'DocumentID' in df_gastos.columns else None
                
                col_llave_gs = None
                for cand in ['SolicitudPago', 'Solicitud de Pago', 'DocFolio']:
                    if cand in df_gastos.columns:
                        col_llave_gs = cand
                        break
                
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
                    
                    if 'DocumentID' in df_sp_base.columns:
                        df_sp_base = df_sp_base.drop(columns=['DocumentID'])
                        
                    if 'EmpresaOrigen' in df_gs_sub.columns and 'EmpresaOrigen' in df_sp_base.columns:
                        df_sp_base = pd.merge(df_sp_base, df_gs_sub, left_on=['EmpresaOrigen', 'DocFolio'], right_on=['EmpresaOrigen', 'DocFolio_Match'], how='left')
                    else:
                        df_sp_base = pd.merge(df_sp_base, df_gs_sub, left_on='DocFolio', right_on='DocFolio_Match', how='left')
                    
                    if 'DocumentID_GS' in df_sp_base.columns:
                        df_sp_base['DocumentID'] = df_sp_base['DocumentID_GS']
                else:
                    df_sp_base['CFDIFolioFiscal'] = ""
                    df_sp_base['TotalGastos'] = 0.0
                    df_sp_base['DocumentID'] = ""
            else:
                df_sp_base['CFDIFolioFiscal'] = ""
                df_sp_base['TotalGastos'] = 0.0
                df_sp_base['DocumentID'] = ""

            if 'CFDIFolioFiscal' not in df_sp_base.columns: df_sp_base['CFDIFolioFiscal'] = ""
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
                df_sp_base['DateOperation'] = ""

            if 'Amount' not in df_sp_base.columns: df_sp_base['Amount'] = 0.0
            if 'DateOperation' not in df_sp_base.columns: df_sp_base['DateOperation'] = ""

            df_sp_base['Amount'] = pd.to_numeric(df_sp_base['Amount'], errors='coerce').fillna(0)
            df_sp_base['Total'] = pd.to_numeric(df_sp_base['Total'], errors='coerce').fillna(0)
            df_sp_base['TotalGastos'] = pd.to_numeric(df_sp_base['TotalGastos'], errors='coerce').fillna(0)
            
            group_keys_sp = ['EmpresaOrigen', 'DocFolio'] if 'EmpresaOrigen' in df_sp_base.columns else ['DocFolio']
            acumulado_gs = df_sp_base.groupby(group_keys_sp)['TotalGastos'].cumsum()
            df_sp_base['SaldoSP'] = df_sp_base['Total'] - acumulado_gs
            df_sp_base['SaldoSP'] = df_sp_base['SaldoSP'].apply(lambda x: max(0.0, x))
            
            total_pagado_por_doc = df_sp_base.groupby(group_keys_sp + ['DocumentID'])['Amount'].transform('sum')
            df_sp_base['SaldoPagoSP'] = df_sp_base['TotalGastos'] - total_pagado_por_doc
            df_sp_base['SaldoPagoSP'] = df_sp_base['SaldoPagoSP'].apply(lambda x: max(0.0, x))

            if anio_oc_sel != "Todos" and 'DateDocument' in df_sp_base.columns:
                df_sp_base['DateDocument_dt'] = pd.to_datetime(df_sp_base['DateDocument'], errors='coerce')
                df_sp_base = df_sp_base[df_sp_base['DateDocument_dt'].dt.year == int(anio_oc_sel)]

            if texto_busqueda_oc:
                cols_buscar_sp = [c for c in ['DocFolio', 'BusinessEntityName', 'Title'] if c in df_sp_base.columns]
                if cols_buscar_sp:
                    mask_sp = df_sp_base[cols_buscar_sp].astype(str).apply(lambda col: col.str.contains(texto_busqueda_oc, case=False, na=False)).any(axis=1)
                    df_sp_base = df_sp_base[mask_sp]

            total_sp_val = df_sp_base.drop_duplicates(subset=group_keys_sp)['Total'].sum()
            total_gastos_val = df_sp_base.drop_duplicates(subset=group_keys_sp + ['CFDIFolioFiscal', 'TotalGastos'])['TotalGastos'].sum()
            total_amount_sp_val = df_sp_base['Amount'].sum()
            
            cols_grupo_saldo_sp = group_keys_sp + ['DocumentID'] if 'DocumentID' in df_sp_base.columns else group_keys_sp
            total_saldo_pago_sp_val = df_sp_base.drop_duplicates(subset=cols_grupo_saldo_sp)['SaldoPagoSP'].sum()

            kpi_sp1, kpi_sp2, kpi_sp3, kpi_sp4 = st.columns(4)
            with kpi_sp1: st.metric("Total Solicitudes", formato_mx(total_sp_val))
            with kpi_sp2: st.metric("Total Gastos", formato_mx(total_gastos_val))
            with kpi_sp3: st.metric("Total Pagado (Amount)", formato_mx(total_amount_sp_val))
            with kpi_sp4: st.metric("Saldo Pendiente Pago", formato_mx(total_saldo_pago_sp_val))
            st.markdown("---")

            df_sp_filtered = df_sp_base.copy()
            if filtro_estado_saldo == "Con saldo pendiente (> $0)":
                df_sp_filtered = df_sp_filtered[df_sp_filtered['SaldoPagoSP'] > 0]
            elif filtro_estado_saldo == "Liquidados / En cero ($0)":
                df_sp_filtered = df_sp_filtered[df_sp_filtered['SaldoPagoSP'] <= 0]

            df_sp_view = df_sp_filtered[[c for c in columnas_sp if c in df_sp_filtered.columns]].copy()
            for col_m in ['SubTotal', 'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total', 'TotalGastos', 'SaldoSP', 'Rate', 'Amount', 'SaldoPagoSP']:
                if col_m in df_sp_view.columns:
                    df_sp_view[col_m] = pd.to_numeric(df_sp_view[col_m], errors='coerce').apply(formato_mx)

            st.dataframe(df_sp_view, use_container_width=True)
        else:
            st.warning("No se encontró información de SolicitudPago.")

elif menu == "Reporte Pagos":
    st.title("📋 Reporte Ejecutivo de Pagos con Saldo Pendiente")
    st.markdown("### Muestra exclusivamente las OC y SP que tienen saldo pendiente (> $0.00)")

    df_lista_combinada = []
    
    if df_ordenes is not None:
        df_oc_rep = df_ordenes.copy()
        df_oc_rep['Tipo_Movimiento'] = 'Orden de Compra (OC)'
        df_oc_rep['Total_Num'] = pd.to_numeric(df_oc_rep['Total'], errors='coerce').fillna(0)
        df_oc_rep['Saldo_Pendiente'] = df_oc_rep['Total_Num']
        df_lista_combinada.append(df_oc_rep)
        
    if df_tesoreria is not None:
        df_sp_rep = df_tesoreria.copy()
        df_sp_rep['Tipo_Movimiento'] = 'Solicitud de Pago (SP)'
        df_sp_rep['Total_Num'] = pd.to_numeric(df_sp_rep['Total'], errors='coerce').fillna(0)
        df_sp_rep['Saldo_Pendiente'] = df_sp_rep['Total_Num']
        df_lista_combinada.append(df_sp_rep)

    if df_lista_combinada:
        df_rep_total = pd.concat(df_lista_combinada, ignore_index=True)
        
        col_emp_name = 'EmpresaOrigen' if 'EmpresaOrigen' in df_rep_total.columns else None
        col_prov_name = 'BusinessEntityName' if 'BusinessEntityName' in df_rep_total.columns else ('Provider' if 'Provider' in df_rep_total.columns else None)
        col_folio_name = 'DocFolio' if 'DocFolio' in df_rep_total.columns else None
        col_currency = 'Currency' if 'Currency' in df_rep_total.columns else None
        
        if 'Saldo_Pendiente' in df_rep_total.columns:
            df_rep_total = df_rep_total[df_rep_total['Saldo_Pendiente'] > 0]

        if col_prov_name and col_folio_name:
            st.markdown("#### ⚙️ Filtros de Selección")
            
            if col_emp_name:
                lista_empresas = sorted(df_rep_total[col_emp_name].dropna().unique())
                empresas_seleccionadas = st.multiselect("Filtrar por Empresa Origen:", lista_empresas, default=lista_empresas)
                if empresas_seleccionadas:
                    df_rep_total = df_rep_total[df_rep_total[col_emp_name].isin(empresas_seleccionadas)]

            tipos_disponibles = sorted(df_rep_total['Tipo_Movimiento'].dropna().unique())
            tipos_seleccionados = st.multiselect("Filtrar por Tipo (OC / SP):", tipos_disponibles, default=tipos_disponibles)
            if tipos_seleccionados:
                df_rep_total = df_rep_total[df_rep_total['Tipo_Movimiento'].isin(tipos_seleccionados)]

            lista_proveedores = sorted(df_rep_total[col_prov_name].dropna().unique())
            prov_seleccionados = st.multiselect("Filtrar por Proveedor(es):", lista_proveedores, default=lista_proveedores)
            if prov_seleccionados:
                df_rep_total = df_rep_total[df_rep_total[col_prov_name].isin(prov_seleccionados)]

            lista_folios = sorted(df_rep_total[col_folio_name].dropna().unique())
            folios_seleccionados = st.multiselect("Filtrar por Folio(s) Específico(s) (OC / SP):", lista_folios, default=lista_folios)
            if folios_seleccionados:
                df_rep_total = df_rep_total[df_rep_total[col_folio_name].isin(folios_seleccionados)]

            st.markdown("---")

            col_fecha = 'DateDocument' if 'DateDocument' in df_rep_total.columns else None
            col_desc = 'Title' if 'Title' in df_rep_total.columns else ('Description' if 'Description' in df_rep_total.columns else None)

            if not df_rep_total.empty:
                if col_fecha:
                    df_rep_total['Fecha_Fmt'] = pd.to_datetime(df_rep_total[col_fecha], errors='coerce').dt.strftime('%d/%m/%Y')
                else:
                    df_rep_total['Fecha_Fmt'] = ""

                total_general_rep = df_rep_total['Saldo_Pendiente'].sum()
                st.markdown(f"### 💰 **Total Saldo Pendiente General:** {formato_mx(total_general_rep)}")
                
                empresas_agrupadas = df_rep_total[col_emp_name].dropna().unique() if col_emp_name else ['General']
                
                empresa_para_excel = empresas_agrupadas[0] if len(empresas_agrupadas) == 1 else "Consolidado"
                excel_file_bytes = generar_excel_ejecutivo(df_rep_total, empresa_para_excel)

                st.download_button(
                    label="📥 Descargar Reporte en Excel (Formato Ejecutivo Estilizado)",
                    data=excel_file_bytes,
                    file_name="Reporte_Ejecutivo_Pagos.xlsx",
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
                        
                        with st.expander(f"👤 {proveedor} — Saldo Pendiente: {formato_mx(subtotal_prov)}", expanded=True):
                            df_det = df_emp_subset[df_emp_subset[col_prov_name] == proveedor]
                            
                            data_det_list = []
                            for _, row in df_det.iterrows():
                                data_det_list.append({
                                    "Tipo": row['Tipo_Movimiento'],
                                    "Fecha Vencimiento": row['Fecha_Fmt'],
                                    "Folio / Documento": row[col_folio_name],
                                    "Moneda": row[col_currency] if col_currency and not pd.isnull(row[col_currency]) else "MXN",
                                    "Descripción": row[col_desc] if col_desc else "",
                                    "Saldo Pendiente": formato_mx(row['Saldo_Pendiente'])
                                })
                            
                            df_tabla_det = pd.DataFrame(data_det_list)
                            st.dataframe(df_tabla_det, use_container_width=True)
                    st.markdown("---")
            else:
                st.warning("No hay registros con saldo pendiente mayor a cero con los filtros seleccionados.")
        else:
            st.warning("Faltan columnas de proveedor o folios en los archivos.")
    else:
        st.warning("No hay datos cargados para generar el reporte.")