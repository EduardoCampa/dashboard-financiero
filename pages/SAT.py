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

# --- VALIDACIÓN DE SESIÓN ---
if "autenticado" not in st.session_state or not st.session_state.autenticado:
    st.warning("⚠️ Debes iniciar sesión en la página principal para acceder a este módulo.")
    st.stop()

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
            nombre_pestana = str(hoja).strip().upper()
            
            try:
                df_raw = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=20)
                if df_raw.empty:
                    continue

                header_row_idx = 0
                for r_idx, row in df_raw.iterrows():
                    row_str = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
                    if any(k in row_str for k in ['CUENTA', 'CTA', 'NOMBRE', 'DESCRIPCION', 'SALDO', 'CARGO', 'ABONO']):
                        header_row_idx = r_idx
                        break

                df_hoja = pd.read_excel(xls, sheet_name=hoja, skiprows=header_row_idx)
                if df_hoja.empty:
                    continue

                cols_upper = [str(c).strip().upper() for c in df_hoja.columns]

                idx_de_f = next((i for i, c in enumerate(cols_upper) if any(x in c for x in ['DEUDOR F', 'FINAL D', 'SALDO D'])), None)
                idx_ac_f = next((i for i, c in enumerate(cols_upper) if any(x in c for x in ['ACREEDOR F', 'FINAL A', 'SALDO A'])), None)
                idx_col_e = next((i for i, c in enumerate(cols_upper) if any(x in c for x in ['CARGO', 'DEBITO'])), None)
                idx_col_f = next((i for i, c in enumerate(cols_upper) if any(x in c for x in ['ABONO', 'CREDITO'])), None)

                if idx_de_f is None: idx_de_f = min(6, df_hoja.shape[1] - 1)
                if idx_ac_f is None: idx_ac_f = min(7, df_hoja.shape[1] - 1)
                if idx_col_e is None: idx_col_e = min(4, df_hoja.shape[1] - 1)
                if idx_col_f is None: idx_col_f = min(5, df_hoja.shape[1] - 1)

                for _, row in df_hoja.iterrows():
                    cta = str(row.iloc[0]).strip()
                    if not cta or cta.lower() in ('nan', 'cuenta', 'none', 'total', 'totales'):
                        continue
                    
                    cta_original = cta
                    cta_limpia = cta.replace(' ', '').replace('-', '')
                    nom = str(row.iloc[1]).strip() if df_hoja.shape[1] > 1 else ""
                    
                    def parse_monto(val):
                        try:
                            v = str(val).replace('$', '').replace(',', '')
                            return float(pd.to_numeric(v, errors='coerce') or 0.0)
                        except:
                            return 0.0

                    monto_ac = parse_monto(row.iloc[idx_ac_f])
                    monto_de = parse_monto(row.iloc[idx_de_f])
                    monto_col_e = parse_monto(row.iloc[idx_col_e])
                    monto_col_f = parse_monto(row.iloc[idx_col_f])

                    detalles.append({
                        'Empresa': nombre_pestana,
                        'Cuenta_Original': cta_original,
                        'Cuenta': cta_limpia,
                        'Nombre Cuenta': nom,
                        'Saldo Deudor Final': monto_de,
                        'Saldo Acreedor Final': monto_ac,
                        'Columna E (Cargos)': monto_col_e,
                        'Columna F (Abonos)': monto_col_f
                    })
            except Exception:
                continue
    except Exception:
        pass

    return pd.DataFrame(detalles)


# --- FUNCIÓN PARA GENERAR EXCEL EJECUTIVO DE AMARRE DE INGRESOS ---
def generar_excel_ejecutivo_ingresos(df_ingresos_final, df_egresos_final, periodo_str):
    wb = Workbook()
    wb.remove(wb.active)

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

    def estilizar_hoja(ws, titulo_texto, df_datos):
        ws.views.sheetView[0].showGridLines = True
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_datos.columns))
        cell_t = ws.cell(row=1, column=1, value=titulo_texto)
        cell_t.font = font_titulo
        cell_t.fill = fill_titulo
        cell_t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32

        headers = list(df_datos.columns)
        for col_num, h in enumerate(headers, 1):
            c = ws.cell(row=3, column=col_num, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = borde_delgado
        ws.row_dimensions[3].height = 25

        row_idx = 4
        for idx_f, (_, r) in enumerate(df_datos.iterrows()):
            fill_row = fill_zebra if idx_f % 2 == 1 else None
            for c_idx, val in enumerate(r, 1):
                cell = ws.cell(row=row_idx, column=c_idx)
                if c_idx == 1:
                    cell.value = str(val)
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.value = float(val) if pd.notnull(val) else 0.0
                    cell.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
                    cell.alignment = Alignment(horizontal="right", vertical="center")

                cell.font = font_normal
                cell.border = borde_delgado
                if fill_row:
                    cell.fill = fill_row

            ws.row_dimensions[row_idx].height = 20
            row_idx += 1

        ws.cell(row=row_idx, column=1, value="TOTALES").font = font_total
        ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=1).border = borde_total
        ws.cell(row=row_idx, column=1).fill = fill_total

        for c_idx in range(2, len(headers) + 1):
            col_letter = get_column_letter(c_idx)
            cell_tot = ws.cell(row=row_idx, column=c_idx, value=f"=SUM({col_letter}4:{col_letter}{row_idx-1})")
            cell_tot.font = font_total
            cell_tot.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
            cell_tot.alignment = Alignment(horizontal="right", vertical="center")
            cell_tot.border = borde_total
            cell_tot.fill = fill_total

        ws.row_dimensions[row_idx].height = 24

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 18)

    ws_ing = wb.create_sheet(title="Resumen Ingresos")
    estilizar_hoja(ws_ing, f"REPORTE EJECUTIVO DE AMARRE DE INGRESOS — ({periodo_str.upper()})", df_ingresos_final)

    ws_eg = wb.create_sheet(title="Resumen Egresos")
    estilizar_hoja(ws_eg, f"REPORTE EJECUTIVO DE AMARRE DE EGRESOS — ({periodo_str.upper()})", df_egresos_final)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# --- FUNCIÓN PARA GENERAR EXCEL EJECUTIVO MULTI-PESTAÑA (NÓMINAS) ---
def generar_excel_ejecutivo_completo_nomina(dict_dfs_empresas, periodo_str):
    wb = Workbook()
    wb.remove(wb.active)

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
            nombre_archivo = os.path.basename(archivo)
            if any(x in nombre_archivo for x in ["Consolidado", "Balanza", "FORMATO"]) or nombre_archivo.startswith("~$"):
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
        posibles = list(set(posibles))
        posibles_gen = [f for f in posibles if "rh" not in os.path.basename(f).lower() and not os.path.basename(f).startswith("~$")]
        
        ruta_balanza = posibles_gen[0] if posibles_gen else None
        
        if not ruta_balanza or not os.path.exists(ruta_balanza):
            return pd.DataFrame(), pd.DataFrame()

        df_det = obtener_saldos_balanza(ruta_balanza)
        if df_det.empty:
            return pd.DataFrame(), df_det

        df_det['Cuenta_Original'] = df_det['Cuenta_Original'].astype(str)
        df_det['Cuenta'] = df_det['Cuenta'].astype(str)

        res_list = []
        empresas_unicas = df_det['Empresa'].unique()
        for empresa in empresas_unicas:
            df_emp = df_det[df_det['Empresa'] == empresa]
            
            def obtener_saldo_exacto(df_sub, prefijo_cta):
                match = df_sub[
                    (df_sub['Cuenta_Original'] == f"{prefijo_cta}-00000-000-0000") | 
                    (df_sub['Cuenta'] == f"{prefijo_cta}000000000000")
                ]
                return match['Saldo Acreedor Final'].sum() if not match.empty else 0.0

            def obtener_saldo_deudor_exacto(df_sub, prefijo_cta):
                match = df_sub[
                    (df_sub['Cuenta_Original'] == f"{prefijo_cta}-00000-000-0000") | 
                    (df_sub['Cuenta'] == f"{prefijo_cta}000000000000")
                ]
                return match['Saldo Deudor Final'].sum() if not match.empty else 0.0

            v_410 = obtener_saldo_exacto(df_emp, "410")
            v_411 = obtener_saldo_exacto(df_emp, "411")
            
            v_423_ac = obtener_saldo_exacto(df_emp, "423")
            v_423_de = obtener_saldo_deudor_exacto(df_emp, "423")
            v_423 = v_423_ac if v_423_ac > 0 else v_423_de
            
            v_420 = obtener_saldo_deudor_exacto(df_emp, "420")
            v_421 = obtener_saldo_deudor_exacto(df_emp, "421")
            
            ingresos_cont = float(v_410) + float(v_411) - float(v_423)
            egresos_cont = -1 * (float(v_420) + float(v_421))
            
            res_list.append({
                'Empresa': empresa,
                'Cuenta 410': float(v_410),
                'Cuenta 411': float(v_411),
                'Cuenta 423': float(v_423),
                'Contabilidad Ingresos': ingresos_cont,
                'Cuenta 420': float(v_420),
                'Cuenta 421': float(v_421),
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

                df_res = pd.merge(df_map, agg_xml_ing, on='XML', how='left')
                df_res = pd.merge(df_res, agg_xml_eg, on='XML', how='left')
                df_res = pd.merge(df_res, agg_mas_ing, on='MASTER', how='left')
                df_res = pd.merge(df_res, agg_mas_eg, on='MASTER', how='left')

                if not df_balanzas.empty:
                    res_cont_ing, res_cont_eg = [], []
                    for _, r_map in df_res.iterrows():
                        target = str(r_map['CONTABILIDAD']).strip().upper()
                        match = df_balanzas[df_balanzas['Empresa'] == target]
                        if not match.empty:
                            res_cont_ing.append(match['Contabilidad Ingresos'].sum())
                            res_cont_eg.append(match['Contabilidad Egresos'].sum())
                        else:
                            res_cont_ing.append(0.0)
                            res_cont_eg.append(0.0)
                    df_res['Cont_Ingresos'] = res_cont_ing
                    df_res['Cont_Egresos'] = res_cont_eg
                else:
                    df_res['Cont_Ingresos'] = 0.0
                    df_res['Cont_Egresos'] = 0.0

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

                st.markdown("---")

                periodo_str_ing = f"{nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} {anio_sel}"
                excel_ingresos_bytes = generar_excel_ejecutivo_ingresos(
                    df_ingresos_final=df_ingresos_final,
                    df_egresos_final=df_egresos_final,
                    periodo_str=periodo_str_ing
                )

                st.download_button(
                    label="📥 Descargar Reporte Ejecutivo de Ingresos y Egresos en Excel",
                    data=excel_ingresos_bytes,
                    file_name=f"Amarre_Ingresos_Egresos_Ejecutivo_{nombres_meses[mes_final-1]}_{anio_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_ejecutivo_ingresos",
                    use_container_width=True
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
                    st.markdown("**Detalle de Integración de Cuentas:**")
                    st.markdown("`Ingresos = 410 + 411 - 423`  |  `Egresos = 420 + 421`")
                    st.dataframe(
                        df_balanzas.style.format({
                            'Cuenta 410': '${:,.2f}',
                            'Cuenta 411': '${:,.2f}',
                            'Cuenta 423': '${:,.2f}',
                            'Contabilidad Ingresos': '${:,.2f}',
                            'Cuenta 420': '${:,.2f}',
                            'Cuenta 421': '${:,.2f}',
                            'Contabilidad Egresos': '${:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning("⚠️ No se encontraron saldos en la Balanza o las cuentas no coinciden. Asegúrate de que el archivo Balanza.xlsx tenga el formato esperado.")


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


    def cargar_y_procesar_nomina_contabilidad(ruta_base="Balanzas", archivo_balanza_rh="Balanza RH.xlsx", anio_filtro=2026, mes_ini=1, mes_fin=9):
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
            
            ruta_directa = os.path.join(ruta_base, str(anio_filtro), mes_str, archivo_balanza_rh)
            
            posibles = []
            if os.path.exists(ruta_directa):
                posibles = [ruta_directa]
            else:
                encontrados = glob.glob(os.path.join(ruta_base, "**", mes_str, archivo_balanza_rh), recursive=True)
                posibles = [f for f in encontrados if os.path.basename(f).lower() == archivo_balanza_rh.lower()]

            posibles = list(dict.fromkeys(posibles))
            if not posibles:
                continue

            for ruta_balanza in posibles:
                try:
                    xls = pd.ExcelFile(ruta_balanza)

                    for hoja in xls.sheet_names:
                        hoja_upper = hoja.strip().upper()

                        if hoja_upper in ['HOJA1', 'HOJA 1', 'RESUMEN', 'BALANZA', 'SHEET1']:
                            continue

                        nombre_empresa = hoja_upper

                        df_hoja = pd.read_excel(xls, sheet_name=hoja)
                        if df_hoja.empty or df_hoja.shape[1] < 5:
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
                                    val_cargo = row.iloc[idx_cargos_e]
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

    if st.button("🚀 Ejecutar Amarre Nóminas"):
        with st.spinner("Procesando archivos del SAT y Balanzas RH para Nómina..."):
            
            df_nomina_sat = cargar_y_procesar_nomina_sat(
                carpeta_nomina=carpeta_nomina_input,
                anio_filtro=anio_sel_nom,
                mes_ini=mes_inicial_nom,
                mes_fin=mes_final_nom
            )

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

            with subtab_nom_comp:
                periodo_texto = f"{nombres_meses[mes_inicial_nom-1]} a {nombres_meses[mes_final_nom-1]} {anio_sel_nom}"
                st.markdown(f"### ⚖️ Amarre Comparativo SAT vs Contabilidad ({periodo_texto})")
                
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
                                    clave_cta = c_cont_nombre.split(' ')[0].strip()
                                    row_c = df_cont_emp[df_cont_emp['Concepto / Subcuenta'].astype(str).str.contains(clave_cta, regex=False, na=False)]
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


# ==========================================
# 3. SUBMÓDULO: IMPUESTOS
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

    tab_isr, tab_iva, tab_retenciones = st.tabs([
        "📈 ISR (Pagos Provisionales)", 
        "💵 IVA (Cobrado vs Pagado)", 
        "📋 RETENCIONES"
    ])

    with tab_isr:
        st.markdown(f"### 📈 Cálculo de Pago Provisional de ISR — {nombres_meses[mes_imp-1]} {anio_imp}")

        mes_str = f"{int(mes_imp):02d}"
        posibles = (
            glob.glob(os.path.join(ruta_balanzas_imp, str(anio_imp), mes_str, "*.xlsx"), recursive=True) +
            glob.glob(os.path.join(ruta_balanzas_imp, "**", mes_str, "*.xlsx"), recursive=True)
        )
        posibles = list(set(posibles))
        posibles_gen = [f for f in posibles if "rh" not in os.path.basename(f).lower() and not os.path.basename(f).startswith("~$")]
        ruta_balanza_sel = posibles_gen[0] if posibles_gen else None

        if ruta_balanza_sel and os.path.exists(ruta_balanza_sel):
            df_saldos = obtener_saldos_balanza(ruta_balanza_sel)

            if not df_saldos.empty:
                df_saldos['Cuenta_Original'] = df_saldos['Cuenta_Original'].astype(str)
                df_saldos['Cuenta'] = df_saldos['Cuenta'].astype(str)
                empresas_unicas = df_saldos['Empresa'].unique()
                empresa_sel = st.selectbox("Seleccione Empresa:", empresas_unicas, key="emp_isr")

                df_emp = df_saldos[df_saldos['Empresa'] == empresa_sel]

                def obtener_saldo_exacto(df_sub, prefijo_cta):
                    match = df_sub[
                        (df_sub['Cuenta_Original'] == f"{prefijo_cta}-00000-000-0000") | 
                        (df_sub['Cuenta'] == f"{prefijo_cta}000000000000")
                    ]
                    return match['Saldo Acreedor Final'].sum() if not match.empty else 0.0

                def obtener_saldo_deudor_exacto(df_sub, prefijo_cta):
                    match = df_sub[
                        (df_sub['Cuenta_Original'] == f"{prefijo_cta}-00000-000-0000") | 
                        (df_sub['Cuenta'] == f"{prefijo_cta}000000000000")
                    ]
                    return match['Saldo Deudor Final'].sum() if not match.empty else 0.0

                v_410 = obtener_saldo_exacto(df_emp, "410")
                v_411 = obtener_saldo_exacto(df_emp, "411")
                
                v_423_ac = obtener_saldo_exacto(df_emp, "423")
                v_423_de = obtener_saldo_deudor_exacto(df_emp, "423")
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

    with tab_iva:
        st.markdown(f"### 💵 Determinación de IVA por Movimientos del Mes (Cargos y Abonos) — {nombres_meses[mes_imp-1]} {anio_imp}")

        if ruta_balanza_sel and os.path.exists(ruta_balanza_sel):
            df_saldos_iva = obtener_saldos_balanza(ruta_balanza_sel)

            if not df_saldos_iva.empty:
                df_saldos_iva['Cuenta_Original'] = df_saldos_iva['Cuenta_Original'].astype(str)
                df_saldos_iva['Cuenta'] = df_saldos_iva['Cuenta'].astype(str)

                def extraer_movimientos_iva(df_sub):
                    mask_cobrado = df_sub['Cuenta_Original'].str.startswith('201') | df_sub['Cuenta'].str.startswith('201')
                    mask_pagado = df_sub['Cuenta_Original'].str.startswith('101') | df_sub['Cuenta'].str.startswith('101')
                    
                    mask_nombre_cob = df_sub['Nombre Cuenta'].str.upper().str.contains('IVA.*COBRADO|IVA.*TRASLADADO', na=False)
                    mask_nombre_pag = df_sub['Nombre Cuenta'].str.upper().str.contains('IVA.*PAGADO|IVA.*ACREDITABLE', na=False)

                    df_cob = df_sub[mask_cobrado | mask_nombre_cob]
                    df_pag = df_sub[mask_pagado | mask_nombre_pag]

                    tot_cobrado = df_cob['Columna F (Abonos)'].sum() - df_cob['Columna E (Cargos)'].sum()
                    tot_pagado = df_pag['Columna E (Cargos)'].sum() - df_pag['Columna F (Abonos)'].sum()

                    return max(tot_cobrado, 0.0), max(tot_pagado, 0.0)

                empresas_lista = df_saldos_iva['Empresa'].unique()
                lista_resultados_iva = []

                for emp in empresas_lista:
                    df_emp_sub = df_saldos_iva[df_saldos_iva['Empresa'] == emp]
                    iva_cob, iva_pag = extraer_movimientos_iva(df_emp_sub)
                    iva_neto = iva_cob - iva_pag

                    lista_resultados_iva.append({
                        'Empresa': emp,
                        'IVA COBRADO': iva_cob,
                        'IVA PAGADO': iva_pag,
                        'IVA NETO (Cargo/Favor)': iva_neto
                    })

                df_matriz_iva = pd.DataFrame(lista_resultados_iva)

                if not df_matriz_iva.empty:
                    df_pivot_iva = df_matriz_iva.set_index('Empresa')[['IVA COBRADO', 'IVA PAGADO', 'IVA NETO (Cargo/Favor)']].T
                    
                    st.markdown("#### 📋 Concentrado de IVA por Empresa (Movimientos del Mes)")
                    st.dataframe(
                        df_pivot_iva.style.format('${:,.2f}'),
                        use_container_width=True
                    )

                    st.markdown("---")
                    st.markdown("#### 🔍 Detalle Analítico por Empresa Seleccionada")
                    empresa_sel_iva = st.selectbox("Seleccione Empresa para detalle:", empresas_lista, key="emp_iva_det")
                    df_emp_detalle_iva = df_saldos_iva[df_saldos_iva['Empresa'] == empresa_sel_iva]

                    mask_iva_all = (
                        df_emp_detalle_iva['Cuenta_Original'].str.startswith(('201', '101')) |
                        df_emp_detalle_iva['Nombre Cuenta'].str.upper().str.contains('IVA', na=False)
                    )
                    df_cuentas_iva = df_emp_detalle_iva[mask_iva_all][['Cuenta_Original', 'Nombre Cuenta', 'Columna E (Cargos)', 'Columna F (Abonos)', 'Saldo Deudor Final', 'Saldo Acreedor Final']]
                    
                    st.dataframe(
                        df_cuentas_iva.style.format({
                            'Columna E (Cargos)': '${:,.2f}',                             'Columna F (Abonos)': '${:,.2f}',
                            'Saldo Deudor Final': '${:,.2f}',                             'Saldo Acreedor Final': '${:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning("No se pudieron procesar los importes de IVA para las empresas.")

            else:
                st.warning("No se encontraron registros en la balanza del mes.")
        else:
            st.error("No se encontró la balanza contable para el periodo seleccionado.")

    with tab_retenciones:
        st.markdown(f"### 📋 Detalle de Retenciones de Impuestos — {nombres_meses[mes_imp-1]} {anio_imp}")
        st.info("Módulo configurado para consultar retenciones de ISR e IVA basadas en los movimientos del periodo.")