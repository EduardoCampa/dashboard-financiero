import streamlit as st
import pandas as pd
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuración inicial del dashboard en ancho completo
st.set_page_config(page_title="Dashboard Financiero - Contpaq", layout="wide")

# --- REGLA FIJA ESTRICTA PARA FORMATO DE MONEDA REGIÓN MÉXICO ($1,234,567.89) ---
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

# --- FUNCIÓN PARA GENERAR EXCEL CON DISEÑO JERÁRQUICO (EMPRESA -> PROVEEDOR -> MONEDA) ---
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
        cell_emp.font = font_empresa
        cell_emp.fill = fill_empresa
        cell_emp.alignment = Alignment(horizontal="left", vertical="center")
        
        cell_emp_tot = ws.cell(row=row_idx, column=8, value=total_empresa)
        cell_emp_tot.font = font_empresa
        cell_emp_tot.fill = fill_empresa
        cell_emp_tot.number_format = '"$"#,##0.00'
        cell_emp_tot.alignment = Alignment(horizontal="right", vertical="center")
        
        for c_idx in range(1, 9):
            ws.cell(row=row_idx, column=c_idx).border = borde_delgado
        ws.row_dimensions[row_idx].height = 26
        row_idx += 1

        proveedores = sorted(df_emp[col_prov].unique()) if col_prov else []

        for prov in proveedores:
            df_prov = df_emp[df_emp[col_prov] == prov]
            total_prov = df_prov['Saldo_Pendiente'].sum()

            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
            cell_prov = ws.cell(row=row_idx, column=1, value=f"   👤 {str(prov)}")
            cell_prov.font = font_prov
            cell_prov.fill = fill_prov
            cell_prov.alignment = Alignment(horizontal="left", vertical="center")
            
            cell_prov_tot = ws.cell(row=row_idx, column=8, value=total_prov)
            cell_prov_tot.font = font_prov
            cell_prov_tot.fill = fill_prov
            cell_prov_tot.number_format = '"$"#,##0.00'
            cell_prov_tot.alignment = Alignment(horizontal="right", vertical="center")
            
            for c_idx in range(1, 9):
                ws.cell(row=row_idx, column=c_idx).border = borde_delgado
            ws.row_dimensions[row_idx].height = 22
            row_idx += 1

            monedas_prov = sorted(df_prov[col_curr].dropna().unique()) if col_curr else ['MXN']
            for moneda in monedas_prov:
                df_moneda = df_prov[df_prov[col_curr] == moneda] if col_curr else df_prov
                total_moneda = df_moneda['Saldo_Pendiente'].sum()

                ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
                cell_mon = ws.cell(row=row_idx, column=1, value=f"      💱 Moneda: {str(moneda)}")
                cell_mon.font = font_moneda
                cell_mon.fill = fill_moneda
                cell_mon.alignment = Alignment(horizontal="left", vertical="center")
                
                cell_mon_tot = ws.cell(row=row_idx, column=8, value=total_moneda)
                cell_mon_tot.font = font_moneda
                cell_mon_tot.fill = fill_moneda
                cell_mon_tot.number_format = '"$"#,##0.00'
                cell_mon_tot.alignment = Alignment(horizontal="right", vertical="center")
                
                for c_idx in range(1, 9):
                    ws.cell(row=row_idx, column=c_idx).border = borde_delgado
                ws.row_dimensions[row_idx].height = 20
                row_idx += 1

                for col_num, h in enumerate(headers, 1):
                    c = ws.cell(row=row_idx, column=col_num, value=h)
                    c.font = font_header
                    c.fill = fill_header
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border = borde_delgado
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
                    c_val.number_format = '"$"#,##0.00'
                    c_val.alignment = Alignment(horizontal="right")

                    for col_num in range(1, 9):
                        cell_det = ws.cell(row=row_idx, column=col_num)
                        cell_det.font = font_normal
                        cell_det.border = borde_delgado
                    
                    ws.row_dimensions[row_idx].height = 18
                    row_idx += 1

                row_idx += 1
            row_idx += 1
        row_idx += 1

    ws.cell(row=row_idx, column=1, value="TOTAL GENERAL").font = font_total
    for col_num in range(1, 8):
        ws.cell(row=row_idx, column=col_num).fill = fill_total
        ws.cell(row=row_idx, column=col_num).border = borde_total
        
    cell_tot_gen = ws.cell(row=row_idx, column=8, value=df_datos['Saldo_Pendiente'].sum())
    cell_tot_gen.font = font_total
    cell_tot_gen.fill = fill_total
    cell_tot_gen.number_format = '"$"#,##0.00'
    cell_tot_gen.border = borde_total
    cell_tot_gen.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[row_idx].height = 26

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- BARRA LATERAL ---
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

ruta_archivo = "Consolidado_Master.xlsx"

@st.cache_data
def cargar_datos(path):
    try:
        df_factura = pd.read_excel(path, sheet_name='FacturaCliente')
        df_tesoreria = pd.read_excel(path, sheet_name='SolicitudPago')
        df_ordenes = pd.read_excel(path, sheet_name='OrdenCompra')
        df_edocuenta = pd.read_excel(path, sheet_name='EdoCuenta')
        
        df_fact_compra = pd.read_excel(path, sheet_name='FacturaCompra') if 'FacturaCompra' in pd.ExcelFile(path).sheet_names else None
        df_gastos = pd.read_excel(path, sheet_name='Gastos') if 'Gastos' in pd.ExcelFile(path).sheet_names else None
        
        for df_chk in [df_fact_compra, df_gastos, df_ordenes, df_tesoreria, df_factura]:
            if df_chk is not None:
                col_del = 'Deleted' if 'Deleted' in df_chk.columns else ('Delete' if 'Delete' in df_chk.columns else None)
                if col_del:
                    df_chk.drop(df_chk[pd.to_numeric(df_chk[col_del], errors='coerce').fillna(0) == 1].index, inplace=True)

        return df_factura, df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos
    except Exception as e:
        st.error(f"Error al cargar las pestañas del archivo: {e}")
        return None, None, None, None, None, None

df_factura, df_tesoreria, df_ordenes, df_edocuenta, df_fact_compra, df_gastos = cargar_datos(ruta_archivo)

columnas_oc = [
    'EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument',
    'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal',
    'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total', 'TotalFacturaCompra', 'SaldoOC', 'DocumentID', 'UUID', 'Amount', 'DateOperation', 'SaldoPagoOC'
]

columnas_sp = [
    'EmpresaOrigen', 'DocFolio', 'BusinessEntityName', 'DateDocument',
    'Title', 'CostCenterName', 'Currency', 'Rate', 'SubTotal',
    'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total', 'TotalGastos', 'SaldoSP', 'DocumentID', 'CFDIFolioFiscal', 'Amount', 'DateOperation', 'SaldoPagoSP'
]

# --- RENDERIZADO SEGÚN LA SELECCIÓN ---

if area_principal == "Contabilidad":
    st.title("📚 Módulo de Contabilidad")
    if menu_contabilidad == "Balanza de Comprobación":
        st.markdown("### 📊 Consulta Automática de Balanzas de Comprobación")
        
        carpeta_principal_balanzas = "Balanzas"
        if os.path.exists(carpeta_principal_balanzas):
            anios_disponibles = sorted([d for d in os.listdir(carpeta_principal_balanzas) if os.path.isdir(os.path.join(carpeta_principal_balanzas, d))])
        else:
            anios_disponibles = []

        if anios_disponibles:
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                anio_con = st.selectbox("Seleccione el Año:", anios_disponibles, key="balanza_anio_con")
            
            ruta_anio = os.path.join(carpeta_principal_balanzas, str(anio_con))
            meses_disponibles = sorted([d for d in os.listdir(ruta_anio) if os.path.isdir(os.path.join(ruta_anio, d))])
            
            meses_dict_nombres = {
                "01": "01 - Enero", "02": "02 - Febrero", "03": "03 - Marzo", "04": "04 - Abril",
                "05": "05 - Mayo", "06": "06 - Junio", "07": "07 - Julio", "08": "08 - Agosto",
                "09": "09 - Septiembre", "10": "10 - Octubre", "11": "11 - Noviembre", "12": "12 - Diciembre"
            }
            
            meses_opciones = [m for m in meses_disponibles if m in meses_dict_nombres]
            
            with col_c2:
                mes_con = st.selectbox("Seleccione el Mes:", meses_opciones, format_func=lambda x: meses_dict_nombres.get(x, x), key="balanza_mes_con")
            
            if mes_con:
                ruta_a_consultar = os.path.join(carpeta_principal_balanzas, str(anio_con), mes_con, "Balanza.xlsx")
                
                if os.path.exists(ruta_a_consultar):
                    try:
                        xls_temp = pd.ExcelFile(ruta_a_consultar)
                        hojas_guardadas = xls_temp.sheet_names
                        
                        st.info(f"📂 Archivo detectado y listo: `Balanzas/{anio_con}/{mes_con}/Balanza.xlsx`")
                        
                        empresa_sel_con = st.selectbox("Seleccione la empresa a visualizar:", hojas_guardadas, key="visor_empresa_guardada")
                        if empresa_sel_con:
                            # header=7 toma exactamente la fila 8 de Excel como cabecera limpia
                            df_vista = pd.read_excel(ruta_a_consultar, sheet_name=empresa_sel_con, header=7)
                            
                            df_vista = df_vista.dropna(how='all')
                            if not df_vista.empty:
                                primer_col = df_vista.columns[0]
                                df_vista = df_vista[df_vista[primer_col].notnull()]
                            
                            st.markdown(f"#### Empresa: **{empresa_sel_con}** (Periodo: {meses_dict_nombres.get(mes_con, mes_con)} {anio_con})")
                            st.dataframe(df_vista, use_container_width=True)
                    except Exception as e:
                        st.error(f"Error al leer el archivo de balanza: {e}")
                else:
                    st.warning(f"⚠️ No se encontró el archivo `Balanza.xlsx` en la ruta `Balanzas/{anio_con}/{mes_con}/`.")
        else:
            st.warning("⚠️ No se encontró la carpeta `Balanzas/` o no contiene subcarpetas de años registradas.")
    else:
        st.info(f"El módulo de **{menu_contabilidad}** se encuentra en desarrollo.")

elif menu == "Panel":
    st.title("Testing Solutions S.A.")
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
                if not df_t.empty:
                    df_t['Año'] = 0
                    df_t['Periodo'] = "Sin Periodo"
        
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
        utilidad_neta = total_ingresos - (total_oc + total_gastos_sp)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("INGRESOS", formato_mx(total_ingresos))
        with col2: st.metric("EGRESOS (OC)", formato_mx(total_oc))
        with col3: st.metric("GASTOS (SP)", formato_mx(total_gastos_sp))
        with col4: st.metric("UTILIDAD NETA", formato_mx(utilidad_neta))

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

elif menu == "Facturación":
    st.title("📊 Módulo de Facturación y Dashboard de Cobranza")
    if df_factura is not None:
        df_fact_anio = df_factura.copy()
        if 'Total' in df_fact_anio.columns:
            total_fact = pd.to_numeric(df_fact_anio['Total'], errors='coerce').fillna(0).sum()
            st.metric("Total Facturado", formato_mx(total_fact))
        st.dataframe(df_fact_anio.head(50), use_container_width=True)

elif menu == "OC y SP" or menu == "Reporte Pagos":
    df_oc_base = pd.DataFrame()
    if df_ordenes is not None:
        df_oc_base = df_ordenes.copy()
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

    df_sp_base = pd.DataFrame()
    if df_tesoreria is not None:
        df_sp_base = df_tesoreria.copy()
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

    if menu == "OC y SP":
        st.title("📦 Módulo de Órdenes de Compra y Solicitudes de Pago")
        tab_oc, tab_sp = st.tabs(["OrdenCompra", "SolicitudPago"])
        with tab_oc:
            if not df_oc_base.empty:
                st.dataframe(df_oc_base[[c for c in columnas_oc if c in df_oc_base.columns]], use_container_width=True)
        with tab_sp:
            if not df_sp_base.empty:
                st.dataframe(df_sp_base[[c for c in columnas_sp if c in df_sp_base.columns]], use_container_width=True)

    elif menu == "Reporte Pagos":
        st.title("📋 Reporte Ejecutivo de Pagos con Saldo Pendiente")
        st.markdown("### Muestra exclusivamente las OC y SP que tienen saldo pendiente real (> $1.00)")

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
                st.markdown("#### ⚙️ Filtros de Selección Independientes")
                
                lista_empresas = sorted(df_rep_total[col_emp_name].dropna().unique()) if col_emp_name else []
                tipos_disponibles = sorted(df_rep_total['Tipo_Movimiento'].dropna().unique())
                lista_proveedores = sorted(df_rep_total[col_prov_name].dropna().unique())
                lista_folios = sorted(df_rep_total[col_folio_name].dropna().unique())

                if col_emp_name:
                    empresas_seleccionadas = st.multiselect("Filtrar por Empresa Origen:", lista_empresas, default=lista_empresas)
                    if empresas_seleccionadas:
                        df_rep_total = df_rep_total[df_rep_total[col_emp_name].isin(empresas_seleccionadas)]

                tipos_seleccionados = st.multiselect("Filtrar por Tipo (OC / SP):", tipos_disponibles, default=tipos_disponibles)
                if tipos_seleccionados:
                    df_rep_total = df_rep_total[df_rep_total['Tipo_Movimiento'].isin(tipos_seleccionados)]

                prov_seleccionados = st.multiselect("Filtrar por Proveedor(es):", lista_proveedores, default=lista_proveedores)
                if prov_seleccionados:
                    df_rep_total = df_rep_total[df_rep_total[col_prov_name].isin(prov_seleccionados)]

                folios_seleccionados = st.multiselect("Filtrar por Folio(s) Específico(s) (OC / SP):", lista_folios, default=lista_folios)
                if folios_seleccionados:
                    df_rep_total = df_rep_total[df_rep_total[col_folio_name].isin(folios_seleccionados)]

                st.markdown("---")

                col_fecha = 'DateDocument' if 'DateDocument' in df_rep_total.columns else None
                col_desc = 'Title' if 'Title' in df_rep_total.columns else None

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
                else:
                    st.warning("No hay registros que coincidan con los filtros seleccionados.")
            else:
                st.warning("Faltan columnas requeridas en los archivos.")
        else:
            st.warning("No hay datos cargados para generar el reporte.")