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
    ["📊 Amarre Ingresos", "👥 Amarre Nóminas", "🏛️ Impuestos"]
)

nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# --- FUNCIÓN AUXILIAR PARA OBTENER SALDOS DE BALANZA POR CUENTA ---
def obtener_saldos_balanza(ruta_balanza):
    if not ruta_balanza or not os.path.exists(ruta_balanza):
        return pd.DataFrame()

    detalles = []
    try:
        xls = pd.ExcelFile(ruta_balanza)
        for hoja in xls.sheet_names:
            if hoja.lower() in ['hoja1', 'hoja 1', 'resumen', 'sheet1']:
                continue
            
            nombre_pestana = hoja.strip().upper()
            df_hoja = pd.read_excel(ruta_balanza, sheet_name=hoja)
            if df_hoja.empty or df_hoja.shape[1] < 4:
                continue

            cols_upper = [str(c).strip().upper() for c in df_hoja.columns]
            
            idx_ac_f = next((i for i, c in enumerate(cols_upper) if 'ACREEDOR F' in c or 'ACREEDOR FINAL' in c), None)
            if idx_ac_f is None:
                idx_ac_f = next((i for i, c in enumerate(cols_upper) if 'ACREEDOR' in c), df_hoja.shape[1] - 1)
                
            idx_de_f = next((i for i, c in enumerate(cols_upper) if 'DEUDOR F' in c or 'DEUDOR FINAL' in c), None)
            if idx_de_f is None:
                idx_de_f = next((i for i, c in enumerate(cols_upper) if 'DEUDOR' in c), df_hoja.shape[1] - 2)

            for _, row in df_hoja.iterrows():
                cta = str(row.iloc[0]).strip()
                if not cta or cta.lower() in ('nan', 'cuenta', 'none'):
                    continue
                
                cta_limpia = cta.replace(' ', '')
                nom = str(row.iloc[1]).strip() if df_hoja.shape[1] > 1 else ""
                
                monto_ac = float(pd.to_numeric(str(row.iloc[idx_ac_f]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)
                monto_de = float(pd.to_numeric(str(row.iloc[idx_de_f]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)

                detalles.append({
                    'Empresa': nombre_pestana,
                    'Cuenta': cta_limpia,
                    'Nombre Cuenta': nom,
                    'Saldo Deudor Final': monto_de,
                    'Saldo Acreedor Final': monto_ac
                })
    except Exception:
        pass

    return pd.DataFrame(detalles)


# --- FUNCIÓN PARA GENERAR EXCEL EJECUTIVO MULTI-PESTAÑA CON LAS EMPRESAS Y CONSOLIDADO ---
def generar_excel_ejecutivo_completo_nomina(dict_dfs_empresas, periodo_str):
    wb = Workbook()
    wb.remove(wb.active)  # Eliminar la hoja por defecto

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


    def cargar_balanzas_por_mes(anio=2026, mes=8, ruta_base="Balanzas"):
        mes_str = f"{int(mes):02d}"
        
        posibles = (
            glob.glob(os.path.join(ruta_base, str(anio), mes_str, "*.xlsx"), recursive=True) +
            glob.glob(os.path.join(ruta_base, "**", mes_str, "*.xlsx"), recursive=True)
        )
        posibles_gen = [f for f in posibles if "rh" not in os.path.basename(f).lower()]
        ruta_balanza = posibles_gen[0] if posibles_gen else (posibles[0] if posibles else None)
        
        if not ruta_balanza or not os.path.exists(ruta_balanza):
            return pd.DataFrame(columns=['Empresa', 'Contabilidad Ingresos', 'Contabilidad Egresos']), pd.DataFrame()

        df_det = obtener_saldos_balanza(ruta_balanza)
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


    col_a, col_mini, col_mfin = st.columns([1, 1, 1])

    with col_a:
        anio_sel = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0)
    with col_mini:
        mes_inicial = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: nombres_meses[x-1])
    with col_mfin:
        mes_final = st.selectbox("Mes Final:", list(range(1, 13)), index=8, format_func=lambda x: nombres_meses[x-1])

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

                agg_xml_ing = df_ingresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                agg_xml_ing.columns = ['XML', 'XML_Ingresos']
                agg_xml_eg = df_egresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                agg_xml_eg.columns = ['XML', 'XML_Egresos']

                agg_mas_ing = df_ingresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                agg_mas_ing.columns = ['MASTER', 'Master_Ingresos']
                agg_mas_eg = df_egresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                agg_mas_eg.columns = ['MASTER', 'Master_Egresos']

                if not df_balanzas.empty:
                    df_bal_ren = df_balanzas.rename(columns={'Empresa': 'CONTABILIDAD', 'Contabilidad Ingresos': 'Cont_Ingresos', 'Contabilidad Egresos': 'Cont_Egresos'})
                else:
                    df_bal_ren = pd.DataFrame(columns=['CONTABILIDAD', 'Cont_Ingresos', 'Cont_Egresos'])

                df_res = pd.merge(df_map, agg_xml_ing, on='XML', how='left')
                df_res = pd.merge(df_res, agg_xml_eg, on='XML', how='left')
                df_res = pd.merge(df_res, agg_mas_ing, on='MASTER', how='left')
                df_res = pd.merge(df_res, agg_mas_eg, on='MASTER', how='left')
                df_res = pd.merge(df_res, df_bal_ren, on='CONTABILIDAD', how='left')
                df_res = df_res.fillna(0.0)

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

            with tab_xml:
                st.markdown(f"### 📑 Ingresos y Egresos de los XML (Vigentes y Error)")
                if not df_ingresos_sat.empty:
                    st.dataframe(df_ingresos_sat.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)

            with tab_master:
                st.markdown(f"### 📁 Ingresos y Egresos Master")
                if not df_ingresos_master.empty:
                    st.dataframe(df_ingresos_master.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)

            with tab_conc:
                st.markdown(f"### 🔍 Conciliación por UUID")
                st.info("Visualización de conciliación por UUID.")

            with tab_cont:
                st.markdown(f"### 📊 Contabilidad y Balanzas")
                if not df_balanzas.empty:
                    st.dataframe(df_balanzas.style.format({'Contabilidad Ingresos': '${:,.2f}', 'Contabilidad Egresos': '${:,.2f}'}), use_container_width=True, hide_index=True)


# ==========================================
# 2. SUBMÓDULO: AMARRE NÓMINAS
# ==========================================
elif submodulo_sat == "👥 Amarre Nóminas":

    st.title("👥 Módulo SAT - Amarre de Nómina vs Contabilidad")

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

    MAPEO_RFC_EMPRESA = {
        'CIV141222JD5': 'CIVLAT',
        'EFC840210UI4': 'EFCO',
        'GFE811209FZ2': 'FERVIC',
        'SER970728JN8': 'SERVYRE'
    }

    col_a_nom, col_mini_nom, col_mfin_nom = st.columns([1, 1, 1])

    with col_a_nom:
        anio_sel_nom = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0, key="nom_anio")
    with col_mini_nom:
        mes_inicial_nom = st.selectbox("Mes Inicial:", list(range(1, 13)), index=8, format_func=lambda x: nombres_meses[x-1], key="nom_mini")
    with col_mfin_nom:
        mes_final_nom = st.selectbox("Mes Final:", list(range(1, 13)), index=8, format_func=lambda x: nombres_meses[x-1], key="nom_mfin")

    st.markdown("---")
    carpeta_nomina_input = st.text_input("Carpeta Raíz de Nómina (XMLs / Excel):", value="Nomina", key="nom_dir")
    ruta_balanzas_nom_input = st.text_input("Carpeta Raíz de Balanzas:", value="Balanzas", key="nom_bal_dir")
    archivo_balanza_rh_input = st.text_input("Nombre del archivo de Balanza RH:", value="Balanza RH.xlsx", key="nom_bal_rh_file")

    st.info("Seleccione las opciones e inicie el proceso para consultar la nómina.")


# ==========================================
# 3. SUBMÓDULO: IMPUESTOS (NUEVO)
# ==========================================
elif submodulo_sat == "🏛️ Impuestos":

    st.title("🏛️ Módulo de Cálculos de Impuestos")

    col_a_imp, col_m_imp = st.columns([1, 1])

    with col_a_imp:
        anio_imp = st.selectbox("Año de Consulta:", [2026, 2025, 2024], index=0, key="imp_anio")
    with col_m_imp:
        mes_imp = st.selectbox("Mes de Consulta:", list(range(1, 13)), index=8, format_func=lambda x: nombres_meses[x-1], key="imp_mes")

    st.markdown("---")
    ruta_balanzas_imp = st.text_input("Carpeta Raíz de Balanzas:", value="Balanzas", key="imp_bal_dir")

    # PESTAÑAS PRINCIPALES DEL SUBMÓDULO DE IMPUESTOS
    tab_isr, tab_iva, tab_retenciones = st.tabs([
        "📈 ISR (Pagos Provicionales)", 
        "💵 IVA (Cobrado vs Pagado)", 
        "📋 RETENCIONES"
    ])

    # 1. PESTAÑA ISR
    with tab_isr:
        st.markdown(f"### 📈 Cálculo de Pago Provisional de ISR — {nombres_meses[mes_imp-1]} {anio_imp}")

        mes_str = f"{int(mes_imp):02d}"
        posibles = (
            glob.glob(os.path.join(ruta_balanzas_imp, str(anio_imp), mes_str, "*.xlsx"), recursive=True) +
            glob.glob(os.path.join(ruta_balanzas_imp, "**", mes_str, "*.xlsx"), recursive=True)
        )
        posibles_gen = [f for f in posibles if "rh" not in os.path.basename(f).lower()]
        ruta_balanza_sel = posibles_gen[0] if posibles_gen else (posibles[0] if posibles else None)

        if ruta_balanza_sel and os.path.exists(ruta_balanza_sel):
            df_saldos = obtener_saldos_balanza(ruta_balanza_sel)

            if not df_saldos.empty:
                empresas_unicas = df_saldos['Empresa'].unique()
                empresa_sel = st.selectbox("Seleccione Empresa:", empresas_unicas, key="emp_isr")

                df_emp = df_saldos[df_saldos['Empresa'] == empresa_sel]

                # Cuentas ISR: 410-00000-000-0000 + 411-00000-000-0000 - 423-00000-000-0000
                v_410 = df_emp[df_emp['Cuenta'].str.startswith('410-00000-000-0000')]['Saldo Acreedor Final'].sum()
                v_411 = df_emp[df_emp['Cuenta'].str.startswith('411-00000-000-0000')]['Saldo Acreedor Final'].sum()

                df_423 = df_emp[df_emp['Cuenta'].str.startswith('423-00000-000-0000')]
                v_423_ac = df_423['Saldo Acreedor Final'].sum()
                v_423_de = df_423['Saldo Deudor Final'].sum()
                v_423 = v_423_ac if v_423_ac > 0 else v_423_de

                ingresos_isr = float(v_410) + float(v_411) - float(v_423)

                st.markdown("#### 📊 Detalle de Cuentas Contables para Ingresos")
                col_i1, col_i2, col_i3, col_i4 = st.columns(4)
                col_i1.metric("Cuenta 410 (Ingresos Obra)", f"${v_410:,.2f}")
                col_i2.metric("Cuenta 411 (Otros Ingresos)", f"${v_411:,.2f}")
                col_i3.metric("Cuenta 423 (Devoluciones/Descuentos)", f"${v_423:,.2f}")
                col_i4.metric("Ingresos Nominales ISR", f"${ingresos_isr:,.2f}")

                st.markdown("---")
                st.markdown("#### ⚙️ Parámetros de Cálculo ISR")
                
                # Input manual del Coeficiente de Utilidad
                coeficiente_utilidad = st.number_input(
                    "Coeficiente de Utilidad (Manual):", 
                    min_value=0.0000, 
                    max_value=1.0000, 
                    value=0.0500, 
                    step=0.0001, 
                    format="%.4f"
                )

                utilidad_fiscal = ingresos_isr * coeficiente_utilidad
                tasa_isr = 0.30
                pago_provisional = utilidad_fiscal * tasa_isr

                st.markdown("#### 🧮 Resultado del Cálculo")
                col_r1, col_r2, col_r3 = st.columns(3)
                col_r1.metric("Utilidad Fiscal Estimada", f"${utilidad_fiscal:,.2f}")
                col_r2.metric("Tasa ISR", "30.00%")
                col_r3.metric("Pago Provisional Impuesto (ISR)", f"${pago_provisional:,.2f}")

            else:
                st.warning("No se pudieron extraer datos de la balanza seleccionada.")
        else:
            st.error(f"No se encontró ninguna balanza para el periodo {nombres_meses[mes_imp-1]} {anio_imp} en la ruta especificada.")

    # 2. PESTAÑA IVA
    with tab_iva:
        st.markdown(f"### 💵 Determinación de IVA (Cobrado vs Pagado) — {nombres_meses[mes_imp-1]} {anio_imp}")

        if 'ruta_balanza_sel' in locals() and ruta_balanza_sel and os.path.exists(ruta_balanza_sel):
            df_saldos_iva = obtener_saldos_balanza(ruta_balanza_sel)

            if not df_saldos_iva.empty:
                empresas_unicas_iva = df_saldos_iva['Empresa'].unique()
                empresa_sel_iva = st.selectbox("Seleccione Empresa:", empresas_unicas_iva, key="emp_iva")

                df_emp_iva = df_saldos_iva[df_saldos_iva['Empresa'] == empresa_sel_iva]

                # Cuenta IVA Cobrado: 201-00010-001-0000
                v_iva_cobrado = df_emp_iva[df_emp_iva['Cuenta'].str.startswith('201-00010-001-0000')]['Saldo Acreedor Final'].sum()

                # Cuentas IVA Pagado: 101-00011-001-0000 + 101-00011-002-0000
                v_iva_pag1 = df_emp_iva[df_emp_iva['Cuenta'].str.startswith('101-00011-001-0000')]['Saldo Deudor Final'].sum()
                v_iva_pag2 = df_emp_iva[df_emp_iva['Cuenta'].str.startswith('101-00011-002-0000')]['Saldo Deudor Final'].sum()
                v_iva_pagado = float(v_iva_pag1) + float(v_iva_pag2)

                diferencia_iva = v_iva_cobrado - v_iva_pagado

                st.markdown("#### 📊 Cuentas Contables de IVA")
                col_v1, col_v2 = st.columns(2)
                
                with col_v1:
                    st.subheader("IVA Cobrado / Trasladado")
                    st.metric("Cuenta 201-00010-001-0000", f"${v_iva_cobrado:,.2f}")

                with col_v2:
                    st.subheader("IVA Pagado / Acreditable")
                    st.write(f"• **101-00011-001-0000:** ${v_iva_pag1:,.2f}")
                    st.write(f"• **101-00011-002-0000:** ${v_iva_pag2:,.2f}")
                    st.metric("Total IVA Pagado", f"${v_iva_pagado:,.2f}")

                st.markdown("---")
                st.markdown("#### ⚖️ Resumen de Determinación de IVA")
                col_res1, col_res2 = st.columns(2)

                if diferencia_iva > 0:
                    col_res1.metric("IVA a Cargo (A Pagar)", f"${diferencia_iva:,.2f}")
                    col_res2.info("El IVA Cobrado fue mayor al IVA Pagado en el periodo.")
                else:
                    col_res1.metric("IVA a Favor", f"${abs(diferencia_iva):,.2f}")
                    col_res2.success("El IVA Pagado fue mayor al IVA Cobrado en el periodo.")

            else:
                st.warning("No se encontraron registros de IVA en la balanza.")
        else:
            st.error("No se encontró la balanza contable para el periodo seleccionado.")

    # 3. PESTAÑA RETENCIONES
    with tab_retenciones:
        st.markdown(f"### 📋 Detalle de Retenciones de Impuestos — {nombres_meses[mes_imp-1]} {anio_imp}")
        st.info("Módulo configurado para consultar retenciones de ISR e IVA (Servicios Profesionales, Arrendamiento, Fletes, etc.).")