import glob
import os
import io
import re
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Módulo SAT y Conciliación Master vs XML",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- MENÚ LATERAL DE NAVEGACIÓN ---
st.sidebar.markdown("### 📑 Módulo SAT")
submodulo_sat = st.sidebar.radio(
    "Seleccione Submódulo:",
    ["📊 Amarre Ingresos", "👥 Amarre Nóminas"]
)

nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# --- FUNCIÓN PARA GENERAR EXCEL EJECUTIVO MULTI-PESTAÑA CON LAS 4 EMPRESAS Y CONSOLIDADO ---
def generar_excel_ejecutivo_completo_nomina(dict_dfs_empresas, periodo_str):
    wb = Workbook()
    wb.remove(wb.active) # Eliminar la hoja por defecto

    # Estilos Ejecutivos
    font_titulo = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    fill_titulo = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    
    font_normal = Font(name="Calibri", size=10, color="000000")
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

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

    headers = ["Concepto SAT", "Monto SAT", "Cuenta Contable", "Monto Contabilidad", "Diferencia (SAT - Contab)"]

    # 1. PESTAÑA CONSOLIDADO GLOBAL
    ws_cons = wb.create_sheet(title="CONSOLIDADO")
    ws_cons.views.sheetView[0].showGridLines = True

    ws_cons.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    cell_t_c = ws_cons.cell(row=1, column=1, value=f"REPORTE CONSOLIDADO DE AMARRE DE NÓMINA — GRUPO ({periodo_str.upper()})")
    cell_t_c.font = font_titulo
    cell_t_c.fill = fill_titulo
    cell_t_c.alignment = Alignment(horizontal="center", vertical="center")
    ws_cons.row_dimensions[1].height = 32

    for col_num, h in enumerate(headers, 1):
        c = ws_cons.cell(row=3, column=col_num, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borde_delgado
    ws_cons.row_dimensions[3].height = 25

    # Sumar dataframe consolidado
    df_first = list(dict_dfs_empresas.values())[0] if dict_dfs_empresas else pd.DataFrame()
    if not df_first.empty:
        df_cons_data = df_first[['Concepto SAT', 'Cuenta Contable']].copy()
        df_cons_data['Monto SAT'] = 0.0
        df_cons_data['Monto Contabilidad'] = 0.0
        df_cons_data['Diferencia (SAT - Contab)'] = 0.0

        for df_emp in dict_dfs_empresas.values():
            if not df_emp.empty:
                df_cons_data['Monto SAT'] += df_emp['Monto SAT']
                df_cons_data['Monto Contabilidad'] += df_emp['Monto Contabilidad']
                df_cons_data['Diferencia (SAT - Contab)'] += df_emp['Diferencia (SAT - Contab)']

        row_idx_c = 4
        for idx_f, (_, r) in enumerate(df_cons_data.iterrows()):
            ws_cons.cell(row=row_idx_c, column=1, value=str(r.get('Concepto SAT', ''))).alignment = Alignment(horizontal="left", vertical="center")
            
            c1 = ws_cons.cell(row=row_idx_c, column=2, value=float(r.get('Monto SAT', 0.0)))
            c1.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            c1.alignment = Alignment(horizontal="right", vertical="center")

            ws_cons.cell(row=row_idx_c, column=3, value=str(r.get('Cuenta Contable', ''))).alignment = Alignment(horizontal="center", vertical="center")

            c2 = ws_cons.cell(row=row_idx_c, column=4, value=float(r.get('Monto Contabilidad', 0.0)))
            c2.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            c2.alignment = Alignment(horizontal="right", vertical="center")

            c3 = ws_cons.cell(row=row_idx_c, column=5, value=float(r.get('Diferencia (SAT - Contab)', 0.0)))
            c3.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            c3.alignment = Alignment(horizontal="right", vertical="center")

            fill_row = fill_zebra if idx_f % 2 == 1 else None
            for c_num in range(1, 6):
                cell_curr = ws_cons.cell(row=row_idx_c, column=c_num)
                cell_curr.font = font_normal
                cell_curr.border = borde_delgado
                if fill_row:
                    cell_curr.fill = fill_row
            ws_cons.row_dimensions[row_idx_c].height = 20
            row_idx_c += 1

        # Totales Consolidado
        ws_cons.cell(row=row_idx_c, column=1, value="TOTALES").font = font_total
        ws_cons.cell(row=row_idx_c, column=1).alignment = Alignment(horizontal="center", vertical="center")

        for col_i, col_key in [(2, 'Monto SAT'), (4, 'Monto Contabilidad'), (5, 'Diferencia (SAT - Contab)')]:
            c_t = ws_cons.cell(row=row_idx_c, column=col_i, value=df_cons_data[col_key].sum())
            c_t.font = font_total
            c_t.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            c_t.alignment = Alignment(horizontal="right", vertical="center")

        for col_num in range(1, 6):
            c_tot_item = ws_cons.cell(row=row_idx_c, column=col_num)
            c_tot_item.fill = fill_total
            c_tot_item.border = borde_total
        ws_cons.row_dimensions[row_idx_c].height = 24

        for col in ws_cons.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_cons.column_dimensions[col_letter].width = max(max_len + 4, 18)

    # 2. PESTAÑAS INDIVIDUALES POR EMPRESA
    for emp_nombre, df_datos in dict_dfs_empresas.items():
        ws = wb.create_sheet(title=f"Amarre {emp_nombre}"[:31])
        ws.views.sheetView[0].showGridLines = True

        row_idx = 1
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=5)
        cell_t = ws.cell(row=row_idx, column=1, value=f"REPORTE EJECUTIVO DE AMARRE DE NÓMINA — {emp_nombre.upper()} ({periodo_str.upper()})")
        cell_t.font = font_titulo
        cell_t.fill = fill_titulo
        cell_t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row_idx].height = 32
        row_idx += 2

        for col_num, h in enumerate(headers, 1):
            c = ws.cell(row=row_idx, column=col_num, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = borde_delgado
        ws.row_dimensions[row_idx].height = 25
        row_idx += 1

        for idx_f, (_, r) in enumerate(df_datos.iterrows()):
            ws.cell(row=row_idx, column=1, value=str(r.get('Concepto SAT', ''))).alignment = Alignment(horizontal="left", vertical="center")
            
            c_m_sat = ws.cell(row=row_idx, column=2, value=float(r.get('Monto SAT', 0.0)))
            c_m_sat.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            c_m_sat.alignment = Alignment(horizontal="right", vertical="center")

            ws.cell(row=row_idx, column=3, value=str(r.get('Cuenta Contable', ''))).alignment = Alignment(horizontal="center", vertical="center")

            c_m_cont = ws.cell(row=row_idx, column=4, value=float(r.get('Monto Contabilidad', 0.0)))
            c_m_cont.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            c_m_cont.alignment = Alignment(horizontal="right", vertical="center")

            c_dif = ws.cell(row=row_idx, column=5, value=float(r.get('Diferencia (SAT - Contab)', 0.0)))
            c_dif.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            c_dif.alignment = Alignment(horizontal="right", vertical="center")

            fill_row = fill_zebra if idx_f % 2 == 1 else None
            for c_num in range(1, 6):
                cell_curr = ws.cell(row=row_idx, column=c_num)
                cell_curr.font = font_normal
                cell_curr.border = borde_delgado
                if fill_row:
                    cell_curr.fill = fill_row

            ws.row_dimensions[row_idx].height = 20
            row_idx += 1

        # Totales Empresa
        ws.cell(row=row_idx, column=1, value="TOTALES").font = font_total
        ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")
        
        tot_sat = df_datos['Monto SAT'].sum() if 'Monto SAT' in df_datos.columns else 0.0
        tot_cont = df_datos['Monto Contabilidad'].sum() if 'Monto Contabilidad' in df_datos.columns else 0.0
        tot_dif = df_datos['Diferencia (SAT - Contab)'].sum() if 'Diferencia (SAT - Contab)' in df_datos.columns else 0.0

        c_tot_sat = ws.cell(row=row_idx, column=2, value=tot_sat)
        c_tot_sat.font = font_total
        c_tot_sat.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
        c_tot_sat.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=row_idx, column=3, value="").font = font_total

        c_tot_cont = ws.cell(row=row_idx, column=4, value=tot_cont)
        c_tot_cont.font = font_total
        c_tot_cont.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
        c_tot_cont.alignment = Alignment(horizontal="right", vertical="center")

        c_tot_dif = ws.cell(row=row_idx, column=5, value=tot_dif)
        c_tot_dif.font = font_total
        c_tot_dif.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
        c_tot_dif.alignment = Alignment(horizontal="right", vertical="center")

        for col_num in range(1, 6):
            c_tot_item = ws.cell(row=row_idx, column=col_num)
            c_tot_item.fill = fill_total
            c_tot_item.border = borde_total

        ws.row_dimensions[row_idx].height = 24

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 18)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ==========================================
# 1. SUBMÓDULO: AMARRE INGRESOS
# ==========================================
if submodulo_sat == "📊 Amarre Ingresos":

    st.title("📑 Módulo SAT - Comparativos, Conciliación y Reporte Ejecutivo")

    # --- FUNCIONES DE PROCESAMIENTO SAT (XMLs) ---
    def cargar_y_procesar_sat(carpeta_xmls="XML", anio_filtro=None, mes_ini=None, mes_fin=None):
        archivos_excel = (
            glob.glob(os.path.join(carpeta_xmls, "**", "*.xlsx"), recursive=True) + 
            glob.glob(os.path.join(carpeta_xmls, "*.xlsx")) + 
            glob.glob("*.xlsx")
        )
        archivos_excel = list(set(archivos_excel))
        
        registros_ingresos = []
        registros_egresos = []
        
        for archivo in archivos_excel:
            if any(x in archivo for x in ["Consolidado", "Balanza", "FORMATO"]):
                continue
            try:
                df_sat = pd.read_excel(archivo)
                df_sat.columns = [str(c).strip() for c in df_sat.columns]
                
                for _, row in df_sat.iterrows():
                    uuid = ""
                    for col in df_sat.columns:
                        if 'uuid' in col.lower() or 'folio fiscal' in col.lower():
                            val_uuid = str(row.get(col, '')).strip().upper()
                            if val_uuid and val_uuid != 'NAN':
                                uuid = val_uuid
                                break
                    if not uuid:
                        continue

                    estado = "VIGENTE"
                    for col in df_sat.columns:
                        if col.lower() == 'estado' or 'estatus' in col.lower():
                            val_est = str(row.get(col, '')).strip().upper()
                            if val_est and val_est != 'NAN':
                                estado = val_est
                                break
                    
                    if 'VIGENTE' not in estado and 'ERROR' not in estado:
                        continue

                    fecha_emision_str = ""
                    dt_fecha = None
                    for col in df_sat.columns:
                        if 'fecha' in col.lower() and 'emision' in col.lower():
                            val_fecha = row.get(col, '')
                            if pd.notnull(val_fecha):
                                dt_fecha = pd.to_datetime(val_fecha, errors='coerce')
                                if pd.notnull(dt_fecha):
                                    fecha_emision_str = dt_fecha.strftime('%Y-%m-%d')
                            break
                    
                    if dt_fecha is None:
                        continue

                    if anio_filtro and dt_fecha.year != int(anio_filtro):
                        continue
                    if mes_ini and mes_fin and (dt_fecha.month < int(mes_ini) or dt_fecha.month > int(mes_fin)):
                        continue

                    tipo_doc = ""
                    for col in df_sat.columns:
                        if col.lower() == 'tipo':
                            tipo_doc = str(row.get(col, '')).strip()
                            break

                    razon_emisor = ""
                    for col in df_sat.columns:
                        if 'razon' in col.lower() and 'emisor' in col.lower():
                            val_razon = str(row.get(col, '')).strip()
                            if val_razon and val_razon.upper() != 'NAN':
                                razon_emisor = val_razon.upper()
                                break
                    
                    if not razon_emisor:
                        razon_emisor = "SIN RAZÓN EMISOR"

                    subtotal = 0.0
                    descuento = 0.0
                    for col in df_sat.columns:
                        c_low = col.lower()
                        if c_low == 'subtotal' or c_low == 'sub total':
                            val = str(row.get(col, 0)).replace('$', '').replace(',', '').strip()
                            subtotal = float(val) if val and val.lower() != 'nan' else 0.0
                        elif 'descuento' in c_low:
                            val = str(row.get(col, 0)).replace('$', '').replace(',', '').strip()
                            descuento = float(val) if val and val.lower() != 'nan' else 0.0

                    subtotal_neto = subtotal - descuento

                    registro = {
                        'Empresa': razon_emisor,
                        'UUID': uuid,
                        'Fecha emision': fecha_emision_str,
                        'Estado_SAT': estado,
                        'SubTotal': subtotal_neto
                    }

                    if 'I - Ingreso' in tipo_doc or tipo_doc.startswith('I'):
                        registros_ingresos.append(registro)
                    elif 'E - Egreso' in tipo_doc or tipo_doc.startswith('E'):
                        registros_egresos.append(registro)

            except Exception:
                continue

        return pd.DataFrame(registros_ingresos), pd.DataFrame(registros_egresos)


    # --- FUNCIONES DE PROCESAMIENTO CONSOLIDADO MASTER ---
    def cargar_y_procesar_master(ruta_master="Consolidado_Master.xlsx", anio_filtro=None, mes_ini=None, mes_fin=None):
        if not os.path.exists(ruta_master):
            return pd.DataFrame(), pd.DataFrame()

        def procesar_hoja(nombre_hoja, es_egreso=False):
            try:
                df = pd.read_excel(ruta_master, sheet_name=nombre_hoja)
                if df.empty:
                    return pd.DataFrame()

                if 'UUID' in df.columns:
                    df = df[df['UUID'].notnull() & (df['UUID'].astype(str).str.strip() != '') & (df['UUID'].astype(str).str.upper() != 'NAN')].copy()
                else:
                    return pd.DataFrame()

                if 'CFDStatusCancelledName' in df.columns:
                    df = df[df['CFDStatusCancelledName'].astype(str).str.strip().str.upper().str.contains('VIGENTE|ERROR', na=False)].copy()

                fecha_col_name = 'DateDocument' if 'DateDocument' in df.columns else 'CFDIFechaCertificacion'
                if fecha_col_name in df.columns:
                    df['Fecha_Doc'] = pd.to_datetime(df[fecha_col_name], errors='coerce')
                    df = df[df['Fecha_Doc'].notnull()].copy()
                    if anio_filtro:
                        df = df[df['Fecha_Doc'].dt.year == int(anio_filtro)]
                    if mes_ini and mes_fin:
                        df = df[(df['Fecha_Doc'].dt.month >= int(mes_ini)) & (df['Fecha_Doc'].dt.month <= int(mes_fin))]
                    df['Fecha_Doc'] = df['Fecha_Doc'].dt.strftime('%Y-%m-%d')
                else:
                    df['Fecha_Doc'] = ''

                subtotal = pd.to_numeric(df.get('SubTotal', 0), errors='coerce').fillna(0.0)
                descuento = pd.to_numeric(df.get('TotalDiscount', 0), errors='coerce').fillna(0.0)
                subtotal_neto = subtotal - descuento

                if es_egreso:
                    subtotal_neto = subtotal_neto.abs() * -1

                empresa_raw = df['EmpresaOrigen'] if 'EmpresaOrigen' in df.columns else 'SIN EMPRESA'
                df['Empresa'] = empresa_raw.astype(str).str.strip().str.upper()

                estado_master = df['CFDStatusCancelledName'] if 'CFDStatusCancelledName' in df.columns else 'VIGENTE'

                df_final = pd.DataFrame({
                    'Empresa': df['Empresa'],
                    'UUID': df['UUID'],
                    'CFDIFechaCertificacion': df['Fecha_Doc'],
                    'Estado_Master': estado_master,
                    'SubTotal': subtotal_neto
                })

                return df_final
            except Exception:
                return pd.DataFrame()

        df_ingresos_master = procesar_hoja("FacturaCliente", es_egreso=False)
        df_egresos_master = procesar_hoja("NotaCreditoCliente", es_egreso=True)

        return df_ingresos_master, df_egresos_master


    # --- CARGAR BALANZAS Y DETALLE POR MES ---
    def cargar_balanzas_por_mes(anio=2026, mes=8, ruta_base="Balanzas"):
        mes_str = f"{int(mes):02d}"
        ruta_balanza = os.path.join(ruta_base, str(anio), mes_str, "Balanza.xlsx")
        
        if not os.path.exists(ruta_balanza):
            posibles = glob.glob(os.path.join(ruta_base, "**", mes_str, "*.xlsx"), recursive=True)
            if posibles:
                ruta_balanza = posibles[0]
            else:
                return pd.DataFrame(columns=['Empresa', 'Contabilidad Ingresos', 'Contabilidad Egresos']), pd.DataFrame()

        detalles = []
        try:
            xls = pd.ExcelFile(ruta_balanza)
            for hoja in xls.sheet_names:
                if hoja.lower() in ['hoja1', 'resumen']:
                    continue
                
                nombre_pestana = hoja.strip().upper()

                df_hoja = pd.read_excel(ruta_balanza, sheet_name=hoja)
                if df_hoja.empty:
                    continue
                
                num_cols = df_hoja.shape[1]
                if num_cols < 4:
                    continue
                    
                col_cta = 0
                col_nom = 1
                col_deudor_f = num_cols - 2
                col_acreedor_f = num_cols - 1
                
                for _, row in df_hoja.iterrows():
                    cta = str(row.iloc[col_cta]).strip()
                    if not cta or cta.lower() in ('nan', 'cuenta', 'none'):
                        continue
                    
                    cta_limpia = cta.replace(' ', '')
                    nom = str(row.iloc[col_nom]) if num_cols > col_nom else ""
                    
                    monto_ac = float(pd.to_numeric(row.iloc[col_acreedor_f], errors='coerce') or 0.0)
                    monto_de = float(pd.to_numeric(row.iloc[col_deudor_f], errors='coerce') or 0.0)
                    
                    match_cuenta = None
                    tipo_reg = None
                    
                    if cta_limpia.startswith('410-00000-000-0000'):
                        match_cuenta = '410-00000-000-0000'
                        tipo_reg = 'INGRESOS (+)'
                    elif cta_limpia.startswith('411-00000-000-0000'):
                        match_cuenta = '411-00000-000-0000'
                        tipo_reg = 'INGRESOS (+)'
                    elif cta_limpia.startswith('423-00000-000-0000'):
                        match_cuenta = '423-00000-000-0000'
                        tipo_reg = 'INGRESOS (-)'
                    elif cta_limpia.startswith('420-00000-000-0000'):
                        match_cuenta = '420-00000-000-0000'
                        tipo_reg = 'EGRESOS (+)'
                    elif cta_limpia.startswith('421-00000-000-0000'):
                        match_cuenta = '421-00000-000-0000'
                        tipo_reg = 'EGRESOS (+)'

                    if match_cuenta:
                        detalles.append({
                            'Empresa': nombre_pestana,
                            'Hoja Balanza': nombre_pestana,
                            'Tipo': tipo_reg,
                            'Cuenta': match_cuenta,
                            'Nombre Cuenta': nom,
                            'Saldo Deudor Final': monto_de,
                            'Saldo Acreedor Final': monto_ac
                        })
        except Exception:
            pass

        df_det = pd.DataFrame(detalles)
        if df_det.empty:
            return pd.DataFrame(columns=['Empresa', 'Contabilidad Ingresos', 'Contabilidad Egresos']), df_det

        res_list = []
        empresas_unicas = df_det['Empresa'].unique()
        for empresa in empresas_unicas:
            df_emp = df_det[df_det['Empresa'] == empresa]
            
            v_410 = df_emp[df_emp['Cuenta'].str.startswith('410-00000-000-0000')]['Saldo Acreedor Final'].sum()
            v_411 = df_emp[df_emp['Cuenta'].str.startswith('411-00000-000-0000')]['Saldo Acreedor Final'].sum()
            
            df_423 = df_emp[df_emp['Cuenta'].str.startswith('423-00000-000-0000')]
            v_423_ac = df_423['Saldo Acreedor Final'].sum()
            v_423_de = df_423['Saldo Deudor Final'].sum()
            v_423 = v_423_ac if v_423_ac > 0 else v_423_de
            
            v_420 = df_emp[df_emp['Cuenta'].str.startswith('420-00000-000-0000')]['Saldo Deudor Final'].sum()
            v_421 = df_emp[df_emp['Cuenta'].str.startswith('421-00000-000-0000')]['Saldo Deudor Final'].sum()
            
            ingresos_cont = float(v_410) + float(v_411) - float(v_423)
            egresos_cont = -1 * (float(v_420) + float(v_421))
            
            res_list.append({
                'Empresa': empresa,
                'Contabilidad Ingresos': ingresos_cont,
                'Contabilidad Egresos': egresos_cont
            })

        return pd.DataFrame(res_list), df_det


    # --- CONTROLES DE FILTRO POR PERIODOS ---
    col_a, col_mini, col_mfin = st.columns([1, 1, 1])

    with col_a:
        anio_sel = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0)
    with col_mini:
        mes_inicial = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: nombres_meses[x-1])
    with col_mfin:
        mes_final = st.selectbox("Mes Final:", list(range(1, 13)), index=7, format_func=lambda x: nombres_meses[x-1])

    st.markdown("---")
    carpeta_input = st.text_input("Carpeta o ubicación de los archivos del SAT:", value="XML")
    ruta_master_input = st.text_input("Archivo Consolidado Master:", value="Consolidado_Master.xlsx")
    ruta_balanzas_input = st.text_input("Carpeta Raíz de Balanzas:", value="Balanzas")

    if st.button("🚀 Ejecutar Amarre Ingresos"):
        with st.spinner("Procesando información, balanzas por mes y conciliando..."):
            df_ingresos_sat, df_egresos_sat = cargar_y_procesar_sat(carpeta_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
            df_ingresos_master, df_egresos_master = cargar_y_procesar_master(ruta_master_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
            
            df_balanzas, df_detalle_cont = cargar_balanzas_por_mes(anio=anio_sel, mes=mes_final, ruta_base=ruta_balanzas_input)

            st.success("✅ Procesamiento completado con éxito.")

            tab_comp, tab_xml, tab_master, tab_conc, tab_cont = st.tabs([
                "⚖ 1.- Comparativos", 
                "📑 2.- Ingresos y Egresos de los XML", 
                "📁 3.- Ingresos y Egresos Master",
                "🔍 4.- Conciliacion por UUID",
                "📊 5.- Contabilidad"
            ])

            with tab_comp:
                st.markdown(f"### ⚖ Tablas de Resumen General - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
                
                mapeo_tabla = [
                    {"XML": "CIVMEX", "MASTER": "CIVMEX", "CONTABILIDAD": "CIVMEX"},
                    {"XML": "COMERCIALIZADORA DE INFRAESTRUCTURA VIAL", "MASTER": "CIV", "CONTABILIDAD": "CIV"},
                    {"XML": "COMERCIALIZADORA DE INFRAESTRUCTURA VIAL LATINOAMERICANA", "MASTER": "CIVLA", "CONTABILIDAD": "CIVLAT"},
                    {"XML": "ELEMENTOS FABRICADOS Y CONSTRUCCIONES", "MASTER": "EFCO", "CONTABILIDAD": "EFCO"},
                    {"XML": "FGS SISTEMAS INTEGRALES DE MANTENIMIENTO", "MASTER": "FGS", "CONTABILIDAD": "FGS"},
                    {"XML": "GRUPO FERVIC", "MASTER": "FERVIC", "CONTABILIDAD": "FERVIC"},
                    {"XML": "GRUPO FPSB", "MASTER": "GRUPO FPSB", "CONTABILIDAD": "FPSB"},
                    {"XML": "GRUPO PESAZA", "MASTER": "PESAZA", "CONTABILIDAD": "PESAZA"},
                    {"XML": "GRUPO SERVYRE", "MASTER": "GRUPOSERVYRE", "CONTABILIDAD": "GPO SERVYRE"},
                    {"XML": "INMOBILIARIA PSZ", "MASTER": "INMOBILIARIA", "CONTABILIDAD": "INMOBILIARIA"},
                    {"XML": "LABORATORIO MAJONINMAR", "MASTER": "LABORATORIO", "CONTABILIDAD": "LABORATORIO"},
                    {"XML": "LATIN AMERICAN ITS", "MASTER": "LAITS", "CONTABILIDAD": "LATIN"},
                    {"XML": "LIMPIESPIN", "MASTER": "LIMPIESPIN", "CONTABILIDAD": "LIMPIESPIN"},
                    {"XML": "SERVYCARGO", "MASTER": "SERVYCARGO", "CONTABILIDAD": "SERVYCARGO"},
                    {"XML": "SERVYRE", "MASTER": "SERVYRE", "CONTABILIDAD": "SERVYRE"}
                ]
                df_map = pd.DataFrame(mapeo_tabla)

                # Agregados XML
                agg_xml_ing = df_ingresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                agg_xml_ing.columns = ['XML', 'XML_Ingresos']
                agg_xml_eg = df_egresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                agg_xml_eg.columns = ['XML', 'XML_Egresos']

                # Agregados Master
                agg_mas_ing = df_ingresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                agg_mas_ing.columns = ['MASTER', 'Master_Ingresos']
                agg_mas_eg = df_egresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                agg_mas_eg.columns = ['MASTER', 'Master_Egresos']

                # Agregados Contabilidad
                if not df_balanzas.empty:
                    df_bal_ren = df_balanzas.rename(columns={'Empresa': 'CONTABILIDAD', 'Contabilidad Ingresos': 'Cont_Ingresos', 'Contabilidad Egresos': 'Cont_Egresos'})
                else:
                    df_bal_ren = pd.DataFrame(columns=['CONTABILIDAD', 'Cont_Ingresos', 'Cont_Egresos'])

                # Unir al mapa base
                df_res = pd.merge(df_map, agg_xml_ing, on='XML', how='left')
                df_res = pd.merge(df_res, agg_xml_eg, on='XML', how='left')
                df_res = pd.merge(df_res, agg_mas_ing, on='MASTER', how='left')
                df_res = pd.merge(df_res, agg_mas_eg, on='MASTER', how='left')
                df_res = pd.merge(df_res, df_bal_ren, on='CONTABILIDAD', how='left')
                df_res = df_res.fillna(0.0)

                # --- TABLA 1: INGRESOS ---
                st.markdown("#### 📈 1. Tabla de Resumen - Ingresos")
                df_ingresos_final = pd.DataFrame({
                    'Empresa (XML / Master / Contab)': df_res['XML'] + " / " + df_res['MASTER'] + " / " + df_res['CONTABILIDAD'],
                    'XML Ingresos': df_res['XML_Ingresos'],
                    'Master Ingresos': df_res['Master_Ingresos'],
                    'Contabilidad Ingresos': df_res['Cont_Ingresos'],
                    'Dif. (XML vs Master)': df_res['XML_Ingresos'] - df_res['Master_Ingresos'],
                    'Dif. (XML vs Contab)': df_res['XML_Ingresos'] - df_res['Cont_Ingresos']
                })
                
                st.dataframe(
                    df_ingresos_final.style.format({
                        'XML Ingresos': '${:,.2f}',
                        'Master Ingresos': '${:,.2f}', 
                        'Contabilidad Ingresos': '${:,.2f}',
                        'Dif. (XML vs Master)': '${:,.2f}', 
                        'Dif. (XML vs Contab)': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("---")

                # --- TABLA 2: EGRESOS ---
                st.markdown("#### 📉 2. Tabla de Resumen - Egresos")
                df_egresos_final = pd.DataFrame({
                    'Empresa (XML / Master / Contab)': df_res['XML'] + " / " + df_res['MASTER'] + " / " + df_res['CONTABILIDAD'],
                    'XML Egresos': df_res['XML_Egresos'],
                    'Master Egresos': df_res['Master_Egresos'],
                    'Contabilidad Egresos': df_res['Cont_Egresos'],
                    'Dif. (XML vs Master)': df_res['XML_Egresos'] - df_res['Master_Egresos'],
                    'Dif. (XML vs Contab)': df_res['XML_Egresos'] - df_res['Cont_Egresos']
                })
                
                st.dataframe(
                    df_egresos_final.style.format({
                        'XML Egresos': '${:,.2f}',
                        'Master Egresos': '${:,.2f}', 
                        'Contabilidad Egresos': '${:,.2f}',
                        'Dif. (XML vs Master)': '${:,.2f}', 
                        'Dif. (XML vs Contab)': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("---")

                # --- BOTÓN DE EXPORTACIÓN A EXCEL CON FORMATO EJECUTIVO ---
                def generar_excel_amarre_ingresos(df_ing, df_eg, periodo_texto):
                    wb = Workbook()
                    ws_default = wb.active
                    wb.remove(ws_default)

                    # Paleta y estilos
                    fill_titulo = PatternFill("solid", fgColor="1F4E78")
                    fill_header = PatternFill("solid", fgColor="2F75B5")
                    fill_subtitulo = PatternFill("solid", fgColor="D9EAF7")
                    fill_total = PatternFill("solid", fgColor="D9E1F2")
                    fill_ok = PatternFill("solid", fgColor="E2F0D9")
                    fill_dif = PatternFill("solid", fgColor="FCE4D6")
                    fill_zebra = PatternFill("solid", fgColor="F7F9FC")

                    font_titulo = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
                    font_subtitulo = Font(name="Calibri", size=10, italic=True, color="404040")
                    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
                    font_normal = Font(name="Calibri", size=10, color="000000")
                    font_total = Font(name="Calibri", size=10, bold=True, color="000000")
                    font_ok = Font(name="Calibri", size=10, bold=True, color="006100")
                    font_error = Font(name="Calibri", size=10, bold=True, color="9C0006")

                    thin_gray = Side(style="thin", color="D9E1F2")
                    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
                    border_total = Border(top=Side(style="thin", color="7F7F7F"), bottom=Side(style="double", color="1F1F1F"))
                    money_fmt = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'

                    def preparar_hoja(ws, titulo, df, columnas_moneda, incluir_estado=True):
                        ws.sheet_view.showGridLines = False
                        ultima_col = len(df.columns)

                        # Título
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ultima_col)
                        c = ws.cell(1, 1, titulo)
                        c.fill = fill_titulo
                        c.font = font_titulo
                        c.alignment = Alignment(horizontal="center", vertical="center")
                        ws.row_dimensions[1].height = 28

                        # Periodo
                        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ultima_col)
                        c = ws.cell(2, 1, f"Periodo: {periodo_texto}    |    Generado: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}")
                        c.fill = fill_subtitulo
                        c.font = font_subtitulo
                        c.alignment = Alignment(horizontal="center", vertical="center")
                        ws.row_dimensions[2].height = 21

                        fila_header = 4
                        for col_idx, nombre in enumerate(df.columns, 1):
                            c = ws.cell(fila_header, col_idx, nombre)
                            c.fill = fill_header
                            c.font = font_header
                            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                            c.border = border
                        ws.row_dimensions[fila_header].height = 34

                        fila_inicio = fila_header + 1
                        for r_idx, (_, row) in enumerate(df.iterrows(), fila_inicio):
                            for col_idx, nombre in enumerate(df.columns, 1):
                                valor = row.get(nombre, "")
                                c = ws.cell(r_idx, col_idx, valor)
                                c.font = font_normal
                                c.border = border
                                c.alignment = Alignment(horizontal="right" if nombre in columnas_moneda else "left", vertical="center")

                                if nombre in columnas_moneda:
                                    try:
                                        c.value = float(valor)
                                    except Exception:
                                        c.value = 0.0
                                    c.number_format = money_fmt

                                if (r_idx - fila_inicio) % 2 == 1:
                                    c.fill = fill_zebra

                                if incluir_estado and nombre.startswith("Dif."):
                                    try:
                                        diferencia = float(valor)
                                    except Exception:
                                        diferencia = 0.0
                                    if abs(diferencia) <= 0.01:
                                        c.fill = fill_ok
                                        c.font = font_ok
                                    else:
                                        c.fill = fill_dif
                                        c.font = font_error

                        fila_total = fila_inicio + len(df)
                        ws.cell(fila_total, 1, "TOTALES")
                        ws.cell(fila_total, 1).font = font_total
                        ws.cell(fila_total, 1).alignment = Alignment(horizontal="center", vertical="center")

                        for col_idx, nombre in enumerate(df.columns, 1):
                            c = ws.cell(fila_total, col_idx)
                            c.fill = fill_total
                            c.border = border_total
                            if nombre in columnas_moneda:
                                try:
                                    total = pd.to_numeric(df[nombre], errors="coerce").fillna(0).sum()
                                except Exception:
                                    total = 0.0
                                c.value = float(total)
                                c.number_format = money_fmt
                                c.font = font_total
                                c.alignment = Alignment(horizontal="right", vertical="center")

                        ws.row_dimensions[fila_total].height = 24
                        ws.freeze_panes = "A5"
                        ws.auto_filter.ref = f"A4:{get_column_letter(ultima_col)}{fila_total - 1 if len(df) else 4}"
                        ws.print_title_rows = "1:4"
                        ws.sheet_properties.pageSetUpPr.fitToPage = True
                        ws.page_setup.fitToWidth = 1
                        ws.page_setup.fitToHeight = 0
                        ws.page_setup.orientation = "landscape"
                        ws.page_margins.left = 0.25
                        ws.page_margins.right = 0.25
                        ws.page_margins.top = 0.5
                        ws.page_margins.bottom = 0.5

                        # Anchos
                        for col_idx, nombre in enumerate(df.columns, 1):
                            letra = get_column_letter(col_idx)
                            if nombre == "Empresa (XML / Master / Contab)":
                                ancho = 48
                            elif nombre.startswith("Dif."):
                                ancho = 22
                            else:
                                ancho = 20
                            ws.column_dimensions[letra].width = ancho

                        return fila_total

                    # 1) Resumen ejecutivo
                    ws_res = wb.create_sheet("Resumen Ejecutivo")
                    ws_res.sheet_view.showGridLines = False
                    ws_res.merge_cells("A1:F1")
                    c = ws_res["A1"]
                    c.value = "REPORTE EJECUTIVO — AMARRE DE INGRESOS SAT"
                    c.fill = fill_titulo
                    c.font = font_titulo
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    ws_res.row_dimensions[1].height = 30

                    ws_res.merge_cells("A2:F2")
                    c = ws_res["A2"]
                    c.value = f"Periodo: {periodo_texto}    |    Generado: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}"
                    c.fill = fill_subtitulo
                    c.font = font_subtitulo
                    c.alignment = Alignment(horizontal="center")

                    tot_xml = pd.to_numeric(df_ing.get("XML Ingresos", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
                    tot_master = pd.to_numeric(df_ing.get("Master Ingresos", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
                    tot_cont = pd.to_numeric(df_ing.get("Contabilidad Ingresos", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
                    dif_master = tot_xml - tot_master
                    dif_cont = tot_xml - tot_cont

                    kpis = [
                        ("Total XML Ingresos", tot_xml),
                        ("Total Master Ingresos", tot_master),
                        ("Total Contabilidad Ingresos", tot_cont),
                        ("Diferencia XML vs Master", dif_master),
                        ("Diferencia XML vs Contabilidad", dif_cont),
                    ]
                    for i, (label, value) in enumerate(kpis, 4):
                        ws_res.cell(i, 1, label).font = font_header
                        ws_res.cell(i, 1).fill = fill_header
                        ws_res.cell(i, 2, float(value)).number_format = money_fmt
                        ws_res.cell(i, 2).font = font_total
                        ws_res.cell(i, 2).alignment = Alignment(horizontal="right")
                        ws_res.cell(i, 1).border = border
                        ws_res.cell(i, 2).border = border
                        if "Diferencia" in label:
                            if abs(value) <= 0.01:
                                ws_res.cell(i, 2).fill = fill_ok
                                ws_res.cell(i, 2).font = font_ok
                            else:
                                ws_res.cell(i, 2).fill = fill_dif
                                ws_res.cell(i, 2).font = font_error

                    estado = "CONCILIADO" if abs(dif_master) <= 0.01 and abs(dif_cont) <= 0.01 else "CON DIFERENCIAS"
                    ws_res.cell(10, 1, "ESTATUS DEL AMARRE").font = font_header
                    ws_res.cell(10, 1).fill = fill_header
                    ws_res.cell(10, 2, estado).font = font_ok if estado == "CONCILIADO" else font_error
                    ws_res.cell(10, 2).fill = fill_ok if estado == "CONCILIADO" else fill_dif
                    ws_res.cell(10, 1).border = border
                    ws_res.cell(10, 2).border = border

                    ws_res.column_dimensions["A"].width = 38
                    ws_res.column_dimensions["B"].width = 25
                    ws_res.column_dimensions["C"].width = 4
                    ws_res.column_dimensions["D"].width = 4
                    ws_res.column_dimensions["E"].width = 4
                    ws_res.column_dimensions["F"].width = 4
                    ws_res.freeze_panes = "A4"
                    ws_res.sheet_properties.pageSetUpPr.fitToPage = True
                    ws_res.page_setup.fitToWidth = 1
                    ws_res.page_setup.fitToHeight = 1

                    # 2) Hojas detalladas originales, pero con formato ejecutivo
                    columnas_ing = ["XML Ingresos", "Master Ingresos", "Contabilidad Ingresos", "Dif. (XML vs Master)", "Dif. (XML vs Contab)"]
                    columnas_eg = ["XML Egresos", "Master Egresos", "Contabilidad Egresos", "Dif. (XML vs Master)", "Dif. (XML vs Contab)"]

                    ws_ing = wb.create_sheet("Resumen Ingresos")
                    preparar_hoja(ws_ing, "REPORTE EJECUTIVO — RESUMEN DE INGRESOS SAT", df_ing, columnas_ing)

                    ws_eg = wb.create_sheet("Resumen Egresos")
                    preparar_hoja(ws_eg, "REPORTE EJECUTIVO — RESUMEN DE EGRESOS SAT", df_eg, columnas_eg)

                    # Vista inicial
                    wb.active = 0

                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    return output.getvalue()

                excel_data = generar_excel_amarre_ingresos(
                    df_ingresos_final,
                    df_egresos_final,
                    f"{nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} de {anio_sel}"
                )

                st.download_button(
                    label="📥 Descargar Reporte Ejecutivo en Excel",
                    data=excel_data,
                    file_name=f"Reporte_Ejecutivo_Amarre_Ingresos_{nombres_meses[mes_final-1]}_{anio_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with tab_xml:
                st.markdown(f"### 📑 Ingresos y Egresos de los XML (Vigentes y Error) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
                
                st.markdown("#### 📊 Resumen de Ingresos y Egresos por Emisor (XML)")
                if not df_ingresos_sat.empty or not df_egresos_sat.empty:
                    res_ing_x = df_ingresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                    res_ing_x.columns = ['Empresa', 'Ingresos XML']
                    
                    res_eg_x = df_egresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                    res_eg_x.columns = ['Empresa', 'Egresos XML']

                    df_res_xml = pd.merge(res_ing_x, res_eg_x, on='Empresa', how='outer').fillna(0.0)
                    st.dataframe(
                        df_res_xml.style.format({
                            'Ingresos XML': '${:,.2f}', 
                            'Egresos XML': '${:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("Sin datos para resumir.")

                st.markdown("---")
                st.markdown("#### 📋 Detalle Completo de XMLs")
                subtab_x_ing, subtab_x_eg = st.tabs(["📈 Detalle Ingresos XML", "📉 Detalle Egresos XML"])
                
                with subtab_x_ing:
                    if not df_ingresos_sat.empty:
                        st.dataframe(df_ingresos_sat.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                        st.metric("Total Ingresos XML", f"${df_ingresos_sat['SubTotal'].sum():,.2f}")
                    else:
                        st.warning("Sin registros de ingresos en XML.")

                with subtab_x_eg:
                    if not df_egresos_sat.empty:
                        st.dataframe(df_egresos_sat.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                        st.metric("Total Egresos XML", f"${df_egresos_sat['SubTotal'].sum():,.2f}")
                    else:
                        st.warning("Sin registros de egresos en XML.")

            with tab_master:
                st.markdown(f"### 📁 Ingresos y Egresos Master - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
                
                st.markdown("#### 📊 Resumen de Ingresos y Egresos por Empresa Origen (Master)")
                if not df_ingresos_master.empty or not df_egresos_master.empty:
                    res_ing_m = df_ingresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                    res_ing_m.columns = ['Empresa', 'Ingresos Master']
                    
                    res_eg_m = df_egresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                    res_eg_m.columns = ['Empresa', 'Egresos Master']

                    df_res_mas = pd.merge(res_ing_m, res_eg_m, on='Empresa', how='outer').fillna(0.0)
                    st.dataframe(
                        df_res_mas.style.format({
                            'Ingresos Master': '${:,.2f}', 
                            'Egresos Master': '${:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("Sin datos para resumir.")

                st.markdown("---")
                st.markdown("#### 📋 Detalle Completo Master")
                subtab_m_ing, subtab_m_eg = st.tabs(["📈 Detalle Ingresos Master", "📉 Detalle Egresos Master"])
                
                with subtab_m_ing:
                    if not df_ingresos_master.empty:
                        st.dataframe(df_ingresos_master.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                        st.metric("Total Ingresos Master", f"${df_ingresos_master['SubTotal'].sum():,.2f}")
                    else:
                        st.warning("Sin registros de ingresos en Master.")

                with subtab_m_eg:
                    if not df_egresos_master.empty:
                        st.dataframe(df_egresos_master.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                        st.metric("Total Egresos Master", f"${df_egresos_master['SubTotal'].sum():,.2f}")
                    else:
                        st.warning("Sin registros de egresos en Master.")

            with tab_conc:
                st.markdown(f"### 🔍 Conciliación por UUID (Ingresos Vigentes y Error) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
                
                if not df_ingresos_master.empty and not df_ingresos_sat.empty:
                    df_m_agg = df_ingresos_master[['UUID', 'Empresa', 'SubTotal']].rename(
                        columns={'SubTotal': 'SubTotal_Master'}
                    )
                    df_s_agg = df_ingresos_sat[['UUID', 'Empresa', 'Estado_SAT', 'SubTotal']].rename(
                        columns={'SubTotal': 'SubTotal_SAT', 'Empresa': 'Empresa_SAT'}
                    )
                    
                    df_concil = pd.merge(df_m_agg, df_s_agg, on='UUID', how='outer')
                    df_concil['SubTotal_Master'] = df_concil['SubTotal_Master'].fillna(0.0)
                    df_concil['SubTotal_SAT'] = df_concil['SubTotal_SAT'].fillna(0.0)
                    df_concil['Empresa'] = df_concil['Empresa'].fillna(df_concil['Empresa_SAT']).fillna('OTRAS')
                    
                    df_concil['Diferencia'] = df_concil['SubTotal_Master'] - df_concil['SubTotal_SAT']
                    
                    df_con_dif = df_concil[df_concil['Diferencia'].round(2) != 0.0].copy()

                    st.markdown("#### 📊 Resumen de UUIDs con Diferencias / Faltantes")
                    col_d1, col_d2 = st.columns(2)
                    col_d1.metric("Total UUIDs Analizados", len(df_concil))
                    col_d2.metric("UUIDs con Diferencia / Faltantes", len(df_con_dif))

                    st.markdown("---")
                    st.markdown("#### 📋 Detalle Exclusivo de UUIDs con Diferencias o Faltantes en una Base")
                    if not df_con_dif.empty:
                        st.dataframe(
                            df_con_dif[[
                                'UUID', 'Empresa', 'Estado_SAT', 'SubTotal_Master', 'SubTotal_SAT', 'Diferencia'
                            ]].style.format({
                                'SubTotal_Master': '${:,.2f}',
                                'SubTotal_SAT': '${:,.2f}', 
                                'Diferencia': '${:,.2f}'
                            }),
                            use_container_width=True,
                            hide_index=True
                        )
                    else:
                        st.success("🎉 ¡Excelente! No se encontraron diferencias en los UUIDs cruzados.")
                else:
                    st.info("Se requiere información tanto del Master como de los XMLs para ejecutar la conciliación.")

            with tab_cont:
                st.markdown(f"### 📊 Contabilidad y Balanzas (Mes: {nombres_meses[mes_final-1]} {anio_sel})")
                st.markdown("Fórmula aplicada:\n* **Ingresos:** `410-00000-000-0000` (+) `411-00000-000-0000` (-) `423-00000-000-0000`\n* **Egresos:** `420-00000-000-0000` (+) `421-00000-000-0000`")
                
                st.markdown("#### 📊 Resumen Contable por Pestaña de Balanza")
                if not df_balanzas.empty:
                    st.dataframe(
                        df_balanzas.style.format({
                            'Contabilidad Ingresos': '${:,.2f}',                              'Contabilidad Egresos': '${:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("Sin datos contables para resumir.")

                st.markdown("---")
                st.markdown("#### 📋 Detalle Completo de Cuentas Contables")
                if not df_detalle_cont.empty:
                    st.dataframe(
                        df_detalle_cont.style.format({
                            'Saldo Deudor Final': '${:,.2f}',                              'Saldo Acreedor Final': '${:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning("No se encontraron registros contables para las cuentas especificadas en el periodo seleccionado.")


# ==========================================
# 2. SUBMÓDULO: AMARRE NÓMINAS
# ==========================================
elif submodulo_sat == "👥 Amarre Nóminas":

    st.title("👥 Módulo SAT - Amarre de Nómina vs Contabilidad")

    # MAPEO CONSOLIDADO DE EQUIVALENCIAS (Agrupado por Cuenta Contable)
    MAPEO_EQUIVALENCIAS = [
        {"Conceptos SAT": ["001/001/Sueldo"], "Cuenta Contable": "-001-0001 Sueldo"},
        {"Conceptos SAT": ["001/019/Vacaciones a tiempo"], "Cuenta Contable": "-001-0009 Vacaciones"},
        {"Conceptos SAT": ["002/024/Aguinaldo"], "Cuenta Contable": "-001-0011 Aguinaldo"},
        {"Conceptos SAT": ["010/033/Premio Puntualidad"], "Cuenta Contable": "-001-0006 Premio Puntualidad"},
        {"Conceptos SAT": ["021/020/Prima de vacaciones a tiempo", "021/022/Prima de vacaciones reportada"], "Cuenta Contable": "-001-0010 Prima Vacacional"},
        {"Conceptos SAT": ["029/032/Vales Despensa"], "Cuenta Contable": "-001-0002 Despensa"},
        {"Conceptos SAT": ["038/012/Gratificación", "038/013/Compensación"], "Cuenta Contable": "-001-0004 Compensación"},
        {"Conceptos SAT": ["046/002/Asimilados a Salarios"], "Cuenta Contable": "-029-0000 Asimilados"},
        {"Conceptos SAT": ["049/034/Premio Asistencia"], "Cuenta Contable": "-001-0005 Premio Asistencia"}
    ]

    # MAPEO DE RFCS A EMPRESAS DE CONTABILIDAD
    MAPEO_RFC_EMPRESA = {
        'CIV141222JD5': 'CIVLAT',
        'EFC840210UI4': 'EFCO',
        'GFE811209FZ2': 'FERVIC',
        'SER970728JN8': 'SERVYRE'
    }

    # --- 1. PROCESAMIENTO SAT NÓMINAS ---
    def cargar_y_procesar_nomina_sat(carpeta_nomina="Nomina", anio_filtro=None, mes_ini=None, mes_fin=None):
        if not os.path.exists(carpeta_nomina):
            return pd.DataFrame()

        archivos = (
            glob.glob(os.path.join(carpeta_nomina, "**", "*.xlsx"), recursive=True) + 
            glob.glob(os.path.join(carpeta_nomina, "*.xlsx"))
        )
        archivos = list(set(archivos))
        dfs_procesados = []

        for arch in archivos:
            try:
                nombre_base = os.path.basename(arch).upper()
                rfc_archivo = nombre_base.split('-')[0].strip() if '-' in nombre_base else ""

                xls = pd.ExcelFile(arch)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    if df.empty:
                        continue

                    df.columns = [str(c).strip() for c in df.columns]

                    col_rfc = next((c for c in df.columns if 'rfc' in c.lower() and 'emisor' in c.lower()), None)
                    if not col_rfc:
                        col_rfc = next((c for c in df.columns if 'rfc' in c.lower()), None)

                    if col_rfc:
                        rfc_val = df[col_rfc].astype(str).str.strip().str.upper()
                        df['Empresa_RFC_Clean'] = rfc_val.map(lambda r: MAPEO_RFC_EMPRESA.get(r, r))
                    elif rfc_archivo:
                        df['Empresa_RFC_Clean'] = MAPEO_RFC_EMPRESA.get(rfc_archivo, rfc_archivo)
                    else:
                        df['Empresa_RFC_Clean'] = "EMPRESA NÓMINA"

                    col_fecha = next((c for c in df.columns if 'fecha' in c.lower() and ('pago' in c.lower() or 'emision' in c.lower() or 'certific' in c.lower())), None)
                    if not col_fecha:
                        col_fecha = next((c for c in df.columns if 'fecha' in c.lower()), None)

                    if not col_fecha:
                        continue

                    df['Fecha_Dt'] = pd.to_datetime(df[col_fecha], errors='coerce')
                    df = df[df['Fecha_Dt'].notnull()].copy()

                    if anio_filtro:
                        df = df[df['Fecha_Dt'].dt.year == int(anio_filtro)]
                    if mes_ini and mes_fin:
                        df = df[(df['Fecha_Dt'].dt.month >= int(mes_ini)) & (df['Fecha_Dt'].dt.month <= int(mes_fin))]

                    if df.empty:
                        continue

                    df['Mes_Pago'] = df['Fecha_Dt'].dt.strftime('%Y-%m')

                    cols_conceptos = []
                    inicio_conceptos = False

                    for c in df.columns:
                        c_trim = c.strip()
                        if '001/001' in c_trim or c_trim.startswith('001/001'):
                            inicio_conceptos = True

                        if inicio_conceptos:
                            val_num = pd.to_numeric(df[c].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False), errors='coerce')
                            if val_num.notnull().any() and val_num.abs().sum() > 0:
                                df[c] = val_num.fillna(0.0)
                                cols_conceptos.append(c)

                    if not cols_conceptos:
                        for c in df.columns:
                            c_trim = c.strip()
                            if any(c_trim.startswith(prefix) for prefix in ['001/', '002/', '003/', '004/', '005/', '009/', 'Total']):
                                val_num = pd.to_numeric(df[c].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False), errors='coerce')
                                if val_num.notnull().any() and val_num.abs().sum() > 0:
                                    df[c] = val_num.fillna(0.0)
                                    cols_conceptos.append(c)

                    if not cols_conceptos:
                        continue

                    df_agg = df.groupby(['Empresa_RFC_Clean', 'Mes_Pago'])[cols_conceptos].sum().reset_index()
                    dfs_procesados.append(df_agg)

            except Exception:
                continue

        if not dfs_procesados:
            return pd.DataFrame()

        df_consolidado = pd.concat(dfs_procesados, ignore_index=True)
        cols_totales = [c for c in df_consolidado.columns if c not in ['Empresa_RFC_Clean', 'Mes_Pago']]
        df_final = df_consolidado.groupby(['Empresa_RFC_Clean', 'Mes_Pago'])[cols_totales].sum().reset_index()

        df_final = df_final.rename(columns={'Empresa_RFC_Clean': 'Empresa / RFC'})
        
        for c in cols_totales:
            if df_final[c].abs().sum() == 0:
                df_final.drop(columns=[c], inplace=True)

        return df_final


    # --- 2. PROCESAMIENTO CONTABILIDAD NÓMINAS (USANDO BALANZA RH) ---
    def cargar_y_procesar_nomina_contabilidad(ruta_base="Balanzas", archivo_balanza_rh="Balanza RH.xlsx", anio_filtro=2026, mes_ini=1, mes_fin=8):
        registros_contables = []

        mapeo_conceptos_fijos = {
            '0001': '-001-0001 Sueldo',
            '0002': '-001-0002 Despensa',
            '0004': '-001-0004 Compensación',
            '0005': '-001-0005 Premio Asistencia',
            '0006': '-001-0006 Premio Puntualidad',
            '0008': '-001-0008 Gratificación Extraordinaria',
            '0009': '-001-0009 Vacaciones',
            '0010': '-001-0010 Prima Vacacional',
            '0011': '-001-0011 Aguinaldo'
        }

        for m in range(int(mes_ini), int(mes_fin) + 1):
            mes_str = f"{m:02d}"
            
            posibles = glob.glob(os.path.join(ruta_base, str(anio_filtro), mes_str, archivo_balanza_rh)) + \
                       glob.glob(os.path.join(ruta_base, "**", mes_str, archivo_balanza_rh), recursive=True)
            
            # IMPORTANTE: para Nóminas SOLO se debe utilizar Balanza RH.xlsx.
            # No hacemos fallback a cualquier *.xlsx porque en la misma carpeta
            # también existe Balanza.xlsx y ese archivo corresponde a otra fuente.
            posibles = list(dict.fromkeys(posibles))
            if not posibles:
                continue

            for ruta_balanza in posibles:
                try:
                    nombre_archivo = os.path.basename(ruta_balanza).upper()
                    xls = pd.ExcelFile(ruta_balanza)

                    for hoja in xls.sheet_names:
                        hoja_upper = hoja.strip().upper()

                        if hoja_upper not in ['HOJA1', 'HOJA 1', 'RESUMEN', 'BALANZA', 'SHEET1']:
                            nombre_empresa = hoja_upper
                        elif 'EFCO' in nombre_archivo:
                            nombre_empresa = 'EFCO'
                        elif 'CIVLAT' in nombre_archivo:
                            nombre_empresa = 'CIVLAT'
                        elif 'FERVIC' in nombre_archivo:
                            nombre_empresa = 'FERVIC'
                        elif 'SERVYRE' in nombre_archivo:
                            nombre_empresa = 'SERVYRE'
                        else:
                            nombre_empresa = hoja_upper

                        df_hoja = pd.read_excel(ruta_balanza, sheet_name=hoja)
                        if df_hoja.empty or df_hoja.shape[1] < 3:
                            continue

                        idx_cargos_e = 4
                        for idx_c, col_name in enumerate(df_hoja.columns):
                            c_str = str(col_name).strip().upper()
                            if 'CARGO' in c_str or 'DEBITO' in c_str:
                                idx_cargos_e = idx_c
                                break

                        for _, row in df_hoja.iterrows():
                            cta = str(row.iloc[0]).strip()
                            if not cta or cta.lower() in ('nan', 'cuenta', 'none'):
                                continue

                            cta_limpia = cta.replace(' ', '')
                            nom_cuenta = str(row.iloc[1]).strip().title() if df_hoja.shape[1] > 1 else ""

                            match = re.match(r'^(5\d\d|6\d\d)-(\d+)-(001|029|056)-(\d{4})$', cta_limpia)

                            if match:
                                grupo, obra, dept, subcta = match.groups()

                                es_valida = False
                                label_resumen = ""

                                if dept in ['001', '056'] and subcta != '0000':
                                    es_valida = True
                                    if subcta in mapeo_conceptos_fijos:
                                        label_resumen = mapeo_conceptos_fijos[subcta]
                                    else:
                                        label_resumen = f"-{dept}-{subcta} {nom_cuenta}".strip()

                                elif dept == '029' and subcta == '0000':
                                    es_valida = True
                                    label_resumen = "-029-0000 Asimilados"

                                if es_valida:
                                    val_cargo = row.iloc[idx_cargos_e] if df_hoja.shape[1] > idx_cargos_e else 0.0
                                    monto_cargo = float(pd.to_numeric(str(val_cargo).replace('$', '').replace(',', ''), errors='coerce') or 0.0)

                                    if monto_cargo > 0.0:
                                        registros_contables.append({
                                            'Cuenta Completa': f"{cta_limpia} {nom_cuenta}".strip(),
                                            'Subcuenta Resumen': label_resumen,
                                            'Empresa': nombre_empresa,
                                            'Monto': monto_cargo
                                        })
                except Exception:
                    continue

        df_cont = pd.DataFrame(registros_contables)
        if df_cont.empty:
            return pd.DataFrame(), pd.DataFrame()

        pivot_resumen = pd.pivot_table(
            df_cont,
            index='Subcuenta Resumen',
            columns='Empresa',
            values='Monto',
            aggfunc='sum',
            fill_value=0.0
        ).reset_index().rename(columns={'Subcuenta Resumen': 'Concepto / Subcuenta'})

        pivot_detalle = pd.pivot_table(
            df_cont,
            index='Cuenta Completa',
            columns='Empresa',
            values='Monto',
            aggfunc='sum',
            fill_value=0.0
        ).reset_index().rename(columns={'Cuenta Completa': 'Cuenta Contable'})

        return pivot_resumen, pivot_detalle


    # --- CONTROLES DE FILTRO ---
    col_a_nom, col_mini_nom, col_mfin_nom = st.columns([1, 1, 1])

    with col_a_nom:
        anio_sel_nom = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0, key="nom_anio")
    with col_mini_nom:
        mes_inicial_nom = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: nombres_meses[x-1], key="nom_mini")
    with col_mfin_nom:
        mes_final_nom = st.selectbox("Mes Final:", list(range(1, 13)), index=7, format_func=lambda x: nombres_meses[x-1], key="nom_mfin")

    st.markdown("---")
    carpeta_nomina_input = st.text_input("Carpeta Raíz de Nómina (XMLs / Excel):", value="Nomina", key="nom_dir")
    ruta_balanzas_nom_input = st.text_input("Carpeta Raíz de Balanzas:", value="Balanzas", key="nom_bal_dir")
    archivo_balanza_rh_input = st.text_input("Nombre del archivo de Balanza RH:", value="Balanza RH.xlsx", key="nom_bal_rh_file")

    if st.button("🚀 Ejecutar Amarre Nóminas"):
        with st.spinner("Procesando archivos del SAT y Balanzas RH para Nómina..."):
            
            # Cargar SAT Nómina
            df_nomina_sat = cargar_y_procesar_nomina_sat(
                carpeta_nomina=carpeta_nomina_input,
                anio_filtro=anio_sel_nom,
                mes_ini=mes_inicial_nom,
                mes_fin=mes_final_nom
            )

            # Cargar Contabilidad Nómina usando preferentemente Balanza RH.xlsx
            df_nomina_resumen, df_nomina_detalle = cargar_y_procesar_nomina_contabilidad(
                ruta_base=ruta_balanzas_nom_input,
                archivo_balanza_rh=archivo_balanza_rh_input,
                anio_filtro=anio_sel_nom,
                mes_ini=mes_inicial_nom,
                mes_fin=mes_final_nom
            )

            st.success("✅ Procesamiento de Amarre de Nóminas completado.")

            subtab_nom_comp, subtab_nom_sat, subtab_nom_cont = st.tabs([
                "⚖️ Resumen Comparativo", 
                "📑 SAT", 
                "📊 Contabilidad"
            ])

            # PESTAÑA 1: RESUMEN COMPARATIVO CON PESTAÑAS POR EMPRESA
            with subtab_nom_comp:
                periodo_texto = f"{nombres_meses[mes_inicial_nom-1]} a {nombres_meses[mes_final_nom-1]} {anio_sel_nom}"
                st.markdown(f"### ⚖️ Amarre Comparativo SAT vs Contabilidad ({periodo_texto})")
                
                # Obtener la lista unificada de empresas
                empresas_sat = df_nomina_sat['Empresa / RFC'].unique() if not df_nomina_sat.empty else []
                empresas_cont = [c for c in df_nomina_resumen.columns if c != 'Concepto / Subcuenta'] if not df_nomina_resumen.empty else []
                
                todas_empresas = sorted(list(set(empresas_sat).union(set(empresas_cont))))

                if todas_empresas:
                    dict_dfs_para_excel = {}
                    tabs_empresas = st.tabs([f"🏢 {emp}" for emp in todas_empresas])

                    for idx, emp_nombre in enumerate(todas_empresas):
                        with tabs_empresas[idx]:
                            st.markdown(f"#### 🏢 Resumen de Amarre para: `{emp_nombre}`")

                            if not df_nomina_sat.empty and emp_nombre in df_nomina_sat['Empresa / RFC'].values:
                                df_sat_emp = df_nomina_sat[df_nomina_sat['Empresa / RFC'] == emp_nombre]
                                cols_conceptos_sat = [c for c in df_sat_emp.columns if c not in ['Empresa / RFC', 'Mes_Pago']]
                                sumas_sat = df_sat_emp[cols_conceptos_sat].sum()
                            else:
                                sumas_sat = pd.Series(dtype=float)

                            if not df_nomina_resumen.empty and emp_nombre in df_nomina_resumen.columns:
                                df_cont_emp = df_nomina_resumen[['Concepto / Subcuenta', emp_nombre]].rename(
                                    columns={emp_nombre: 'Contabilidad'}
                                )
                            else:
                                df_cont_emp = pd.DataFrame(columns=['Concepto / Subcuenta', 'Contabilidad'])

                            filas_comparativas = []

                            for eq in MAPEO_EQUIVALENCIAS:
                                lista_sat_nombres = eq['Conceptos SAT']
                                c_cont_nombre = eq['Cuenta Contable']

                                monto_sat = 0.0
                                if not sumas_sat.empty:
                                    for c_sat_nombre in lista_sat_nombres:
                                        col_match = next((col for col in sumas_sat.index if c_sat_nombre.lower() in col.lower() or col.lower().startswith(c_sat_nombre[:7].lower())), None)
                                        if col_match:
                                            monto_sat += float(sumas_sat[col_match])

                                monto_cont = 0.0
                                if not df_cont_emp.empty:
                                    row_c = df_cont_emp[df_cont_emp['Concepto / Subcuenta'].str.startswith(c_cont_nombre[:9])]
                                    if not row_c.empty:
                                        monto_cont = float(row_c['Contabilidad'].sum())

                                label_sat_mostrar = " / ".join(lista_sat_nombres)

                                filas_comparativas.append({
                                    'Concepto SAT': label_sat_mostrar,
                                    'Monto SAT': monto_sat,
                                    'Cuenta Contable': c_cont_nombre,
                                    'Monto Contabilidad': monto_cont,
                                    'Diferencia (SAT - Contab)': monto_sat - monto_cont
                                })

                            df_comp_final = pd.DataFrame(filas_comparativas)
                            dict_dfs_para_excel[emp_nombre] = df_comp_final

                            st.dataframe(
                                df_comp_final.style.format({
                                    'Monto SAT': '${:,.2f}',
                                    'Monto Contabilidad': '${:,.2f}',                                     'Diferencia (SAT - Contab)': '${:,.2f}'
                                }),
                                use_container_width=True,
                                hide_index=True
                            )

                    st.markdown("---")

                    # BOTÓN ÚNICO DE DESCARGA CON TODAS LAS PESTAÑAS (EMPRESAS + CONSOLIDADO)
                    excel_completo_bytes = generar_excel_ejecutivo_completo_nomina(
                        dict_dfs_empresas=dict_dfs_para_excel,
                        periodo_str=periodo_texto
                    )

                    st.download_button(
                        label="📥 Descargar Reporte Completo de Nómina en Excel (Todas las Empresas + Consolidado)",
                        data=excel_completo_bytes,
                        file_name=f"Amarre_Nomina_Ejecutivo_{nombres_meses[mes_final_nom-1]}_{anio_sel_nom}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_completo_nomina",
                        use_container_width=True
                    )

                else:
                    st.info("No se encontraron empresas con datos para comparar.")


            # PESTAÑA 2: SAT NÓMINA
            with subtab_nom_sat:
                st.markdown(f"#### 📋 Consolidado de Nómina SAT por RFC / Empresa, Mes y Conceptos ({nombres_meses[mes_inicial_nom-1]} a {nombres_meses[mes_final_nom-1]} {anio_sel_nom})")
                if not df_nomina_sat.empty:
                    num_cols = [col for col in df_nomina_sat.columns if col not in ['Empresa / RFC', 'Mes_Pago']]
                    format_dict = {col: '${:,.2f}' for col in num_cols}
                    
                    st.dataframe(
                        df_nomina_sat.style.format(format_dict),
                        use_container_width=True,
                        hide_index=True
                    )

                    output_nom = io.BytesIO()
                    with pd.ExcelWriter(output_nom, engine='openpyxl') as writer:
                        df_nomina_sat.to_excel(writer, sheet_name='Consolidado Nómina SAT', index=False)
                    excel_data_nom = output_nom.getvalue()

                    st.download_button(
                        label="📥 Descargar Consolidado Nómina SAT en Excel",
                        data=excel_data_nom,
                        file_name=f"Consolidado_Nomina_SAT_{nombres_meses[mes_final_nom-1]}_{anio_sel_nom}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No se encontraron registros de nómina a partir de los conceptos 001/001 en la carpeta especificada para el periodo seleccionado.")

            # PESTAÑA 3: CONTABILIDAD NÓMINA
            with subtab_nom_cont:
                if not df_nomina_resumen.empty:
                    st.markdown(f"#### 📊 1. Resumen Contable de Nómina por Empresa y Concepto ({nombres_meses[mes_inicial_nom-1]} a {nombres_meses[mes_final_nom-1]} {anio_sel_nom})")
                    empresas_cols_res = [col for col in df_nomina_resumen.columns if col != 'Concepto / Subcuenta']
                    format_dict_res = {col: '${:,.2f}' for col in empresas_cols_res}

                    st.dataframe(
                        df_nomina_resumen.style.format(format_dict_res),
                        use_container_width=True,
                        hide_index=True
                    )

                    st.markdown("---")

                    st.markdown(f"#### 📋 2. Detalle Completo de Cuentas por Obra ({nombres_meses[mes_inicial_nom-1]} a {nombres_meses[mes_final_nom-1]} {anio_sel_nom})")
                    empresas_cols_det = [col for col in df_nomina_detalle.columns if col != 'Cuenta Contable']
                    format_dict_det = {col: '${:,.2f}' for col in empresas_cols_det}

                    st.dataframe(
                        df_nomina_detalle.style.format(format_dict_det),
                        use_container_width=True,
                        hide_index=True
                    )

                    output_cont_nom = io.BytesIO()
                    with pd.ExcelWriter(output_cont_nom, engine='openpyxl') as writer:
                        df_nomina_resumen.to_excel(writer, sheet_name='Resumen Nómina', index=False)
                        df_nomina_detalle.to_excel(writer, sheet_name='Detalle Cuentas', index=False)
                    excel_data_cont_nom = output_cont_nom.getvalue()

                    st.download_button(
                        label="📥 Descargar Consolidado Contable de Nómina en Excel",
                        data=excel_data_cont_nom,
                        file_name=f"Consolidado_Contable_Nomina_{nombres_meses[mes_final_nom-1]}_{anio_sel_nom}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No se encontraron cuentas que coincidan con las reglas definidas en las balanzas del periodo seleccionado.")