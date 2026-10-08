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

st.sidebar.markdown("### 📑 Módulo SAT")
submodulo_sat = st.sidebar.radio(
    "Seleccione Submódulo:",
    ["📊 Amarre Ingresos", "👥 Amarre Nóminas", "🏛️ Impuestos"]
)

nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# --- FUNCIÓN DE LECTURA ROBUSTA DE BALANZAS ---
def obtener_saldos_balanza(ruta_balanza):
    if not ruta_balanza or not os.path.exists(ruta_balanza):
        return pd.DataFrame()

    detalles = []
    try:
        xls = pd.ExcelFile(ruta_balanza)
        for hoja in xls.sheet_names:
            hoja_clean = hoja.strip().upper()
            if hoja_clean in ['HOJA1', 'HOJA 1', 'RESUMEN', 'SHEET1', 'BALANZA']:
                continue
            
            df_hoja = pd.read_excel(xls, sheet_name=hoja)
            if df_hoja.empty or df_hoja.shape[1] < 4:
                continue

            # Identificar dinámicamente o por posición predeterminada
            idx_de_f = 6 if df_hoja.shape[1] > 6 else df_hoja.shape[1] - 2
            idx_ac_f = 7 if df_hoja.shape[1] > 7 else df_hoja.shape[1] - 1
            idx_col_e = 4 if df_hoja.shape[1] > 4 else 0 # Cargos
            idx_col_f = 5 if df_hoja.shape[1] > 5 else 0 # Abonos

            for idx_col, col in enumerate(df_hoja.columns):
                c_str = str(col).strip().upper()
                if 'DEUDOR F' in c_str or 'DEUDOR FINAL' in c_str:
                    idx_de_f = idx_col
                elif 'ACREEDOR F' in c_str or 'ACREEDOR FINAL' in c_str:
                    idx_ac_f = idx_col

            for _, row in df_hoja.iterrows():
                cta = str(row.iloc[0]).strip()
                if not cta or cta.lower() in ('nan', 'cuenta', 'none', 'total', 'totales'):
                    continue
                
                # Normalización de la cuenta removiendo espacios y guiones para comparaciones flexibles
                cta_sin_formato = re.sub(r'[^0-9]', '', cta)
                nom = str(row.iloc[1]).strip() if df_hoja.shape[1] > 1 else ""
                
                monto_de_f = float(pd.to_numeric(str(row.iloc[idx_de_f]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)
                monto_ac_f = float(pd.to_numeric(str(row.iloc[idx_ac_f]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)
                monto_col_e = float(pd.to_numeric(str(row.iloc[idx_col_e]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)
                monto_col_f = float(pd.to_numeric(str(row.iloc[idx_col_f]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)

                detalles.append({
                    'Empresa': hoja_clean,
                    'Cuenta_Raw': cta,
                    'Cuenta_Clean': cta_sin_formato,
                    'Nombre Cuenta': nom,
                    'Saldo Deudor Final': monto_de_f,
                    'Saldo Acreedor Final': monto_ac_f,
                    'Columna E (Cargos)': monto_col_e,
                    'Columna F (Abonos)': monto_col_f
                })
    except Exception:
        pass

    return pd.DataFrame(detalles)


# --- EXPORTACIÓN A EXCEL DE IVA ---
def generar_excel_iva(df_transpuesto, periodo_str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen IVA"
    ws.views.sheetView[0].showGridLines = True

    font_titulo = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    fill_titulo = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    font_normal = Font(name="Calibri", size=10, color="000000")
    font_total_col = Font(name="Calibri", size=10, bold=True, color="000000")
    fill_total_col = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    borde_delgado = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    total_cols = len(df_transpuesto.columns)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    cell_t = ws.cell(row=1, column=1, value=f"REPORTE EJECUTIVO DE DETERMINACIÓN DE IVA — ({periodo_str.upper()})")
    cell_t.font = font_titulo
    cell_t.fill = fill_titulo
    cell_t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    for col_num, col_name in enumerate(df_transpuesto.columns, 1):
        c = ws.cell(row=3, column=col_num, value=col_name)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borde_delgado
    ws.row_dimensions[3].height = 25

    for idx_f, (_, r) in enumerate(df_transpuesto.iterrows(), start=4):
        for col_num, col_name in enumerate(df_transpuesto.columns, 1):
            val = r[col_name]
            c = ws.cell(row=idx_f, column=col_num)

            if col_num == 1:
                c.value = str(val)
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.font = Font(name="Calibri", size=10, bold=True)
            else:
                c.value = float(val) if pd.notnull(val) else 0.0
                c.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.font = font_total_col if col_name == 'TOTAL CONSOLIDADO' else font_normal

            c.border = borde_delgado
            if col_name == 'TOTAL CONSOLIDADO':
                c.fill = fill_total_col
            elif idx_f % 2 == 1:
                c.fill = fill_zebra

        ws.row_dimensions[idx_f].height = 20

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 16)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ==========================================
# 1. AMARRE INGRESOS
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
        registros_ingresos, registros_egresos = [], []
        
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

                    subtotal, descuento = 0.0, 0.0
                    for col in df_sat.columns:
                        c_low = col.lower()
                        if c_low == 'subtotal' or c_low == 'sub total':
                            val = str(row.get(col, 0)).replace('$', '').replace(',', '').strip()
                            subtotal = float(val) if val and val.lower() != 'nan' else 0.0
                        elif 'descuento' in c_low:
                            val = str(row.get(col, 0)).replace('$', '').replace(',', '').strip()
                            descuento = float(val) if val and val.lower() != 'nan' else 0.0

                    registro = {
                        'Empresa': razon_emisor, 'UUID': uuid,
                        'Fecha emision': fecha_emision_str, 'Estado_SAT': estado,
                        'SubTotal': subtotal - descuento
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
                if df.empty or 'UUID' not in df.columns:
                    return pd.DataFrame()

                df = df[df['UUID'].notnull() & (df['UUID'].astype(str).str.strip() != '') & (df['UUID'].astype(str).str.upper() != 'NAN')].copy()

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

                return pd.DataFrame({
                    'Empresa': df['Empresa'], 'UUID': df['UUID'],
                    'CFDIFechaCertificacion': df['Fecha_Doc'],
                    'Estado_Master': df['CFDStatusCancelledName'] if 'CFDStatusCancelledName' in df.columns else 'VIGENTE',
                    'SubTotal': subtotal_neto
                })
            except Exception:
                return pd.DataFrame()

        return procesar_hoja("FacturaCliente", es_egreso=False), procesar_hoja("NotaCreditoCliente", es_egreso=True)

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
        for empresa in df_det['Empresa'].unique():
            df_emp = df_det[df_det['Empresa'] == empresa]
            
            # Filtro flexible para cuentas de ingresos 410, 411 y devoluciones 423
            v_410 = df_emp[df_emp['Cuenta_Clean'].str.startswith('410')]['Saldo Acreedor Final'].sum()
            v_411 = df_emp[df_emp['Cuenta_Clean'].str.startswith('411')]['Saldo Acreedor Final'].sum()
            
            df_423 = df_emp[df_emp['Cuenta_Clean'].str.startswith('423')]
            v_423_ac = df_423['Saldo Acreedor Final'].sum()
            v_423_de = df_423['Saldo Deudor Final'].sum()
            v_423 = v_423_ac if v_423_ac > 0 else v_423_de
            
            v_420 = df_emp[df_emp['Cuenta_Clean'].str.startswith('420')]['Saldo Deudor Final'].sum()
            v_421 = df_emp[df_emp['Cuenta_Clean'].str.startswith('421')]['Saldo Deudor Final'].sum()
            
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
                "⚖ 1.- Comparativos", "📑 2.- Ingresos y Egresos de los XML", 
                "📁 3.- Ingresos y Egresos Master", "🔍 4.- Conciliacion por UUID", "📊 5.- Contabilidad"
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

                # BÚSQUEDA Y CRUCE FLEXIBLE DE CONTABILIDAD
                if not df_balanzas.empty:
                    df_bal_ren = df_balanzas.rename(columns={'Empresa': 'CONTABILIDAD', 'Contabilidad Ingresos': 'Cont_Ingresos', 'Contabilidad Egresos': 'Cont_Egresos'})
                else:
                    df_bal_ren = pd.DataFrame(columns=['CONTABILIDAD', 'Cont_Ingresos', 'Cont_Egresos'])

                df_res = pd.merge(df_map, agg_xml_ing, on='XML', how='left')
                df_res = pd.merge(df_res, agg_xml_eg, on='XML', how='left')
                df_res = pd.merge(df_res, agg_mas_ing, on='MASTER', how='left')
                df_res = pd.merge(df_res, agg_mas_eg, on='MASTER', how='left')
                
                # Coincidencia flexible de nombre de pestaña vs nombre del mapa
                if not df_bal_ren.empty:
                    res_cont_ing, res_cont_eg = [], []
                    for _, r_map in df_res.iterrows():
                        nombre_target = str(r_map['CONTABILIDAD']).strip().upper()
                        match = df_bal_ren[df_bal_ren['CONTABILIDAD'].str.contains(nombre_target, na=False)]
                        if not match.empty:
                            res_cont_ing.append(match['Cont_Ingresos'].sum())
                            res_cont_eg.append(match['Cont_Egresos'].sum())
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
                st.dataframe(df_ingresos_final.style.format({'XML Ingresos': '${:,.2f}', 'Master Ingresos': '${:,.2f}', 'Contabilidad Ingresos': '${:,.2f}', 'Dif. (XML vs Master)': '${:,.2f}', 'Dif. (XML vs Contab)': '${:,.2f}'}), use_container_width=True, hide_index=True)

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
                st.dataframe(df_egresos_final.style.format({'XML Egresos': '${:,.2f}', 'Master Egresos': '${:,.2f}', 'Contabilidad Egresos': '${:,.2f}', 'Dif. (XML vs Master)': '${:,.2f}', 'Dif. (XML vs Contab)': '${:,.2f}'}), use_container_width=True, hide_index=True)

            with tab_xml:
                st.markdown("### 📑 Ingresos y Egresos XML")
                if not df_ingresos_sat.empty:
                    st.dataframe(df_ingresos_sat.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)

            with tab_master:
                st.markdown("### 📁 Ingresos y Egresos Master")
                if not df_ingresos_master.empty:
                    st.dataframe(df_ingresos_master.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)

            with tab_conc:
                st.markdown("### 🔍 Conciliación por UUID")
                st.info("Visualización por UUID disponible.")

            with tab_cont:
                st.markdown("### 📊 Contabilidad y Balanzas Detalle")
                if not df_balanzas.empty:
                    st.dataframe(df_balanzas.style.format({'Contabilidad Ingresos': '${:,.2f}', 'Contabilidad Egresos': '${:,.2f}'}), use_container_width=True, hide_index=True)


# ==========================================
# 2. AMARRE NÓMINAS
# ==========================================
elif submodulo_sat == "👥 Amarre Nóminas":
    st.title("👥 Módulo SAT - Amarre de Nómina vs Contabilidad")
    st.info("Módulo listo para ejecutar.")


# ==========================================
# 3. IMPUESTOS (RESUMEN MULTI-EMPRESA)
# ==========================================
elif submodulo_sat == "🏛️ Impuestos":

    st.title("🏛️ Módulo de Cálculos de Impuestos — Resumen Ejecutivo")

    col_a_imp, col_m_imp = st.columns([1, 1])
    with col_a_imp:
        anio_imp = st.selectbox("Año de Consulta:", [2026, 2025, 2024], index=0, key="imp_anio")
    with col_m_imp:
        mes_imp = st.selectbox("Mes de Consulta:", list(range(1, 13)), index=8, format_func=lambda x: nombres_meses[x-1], key="imp_mes")

    st.markdown("---")
    ruta_balanzas_imp = st.text_input("Carpeta Raíz de Balanzas:", value="Balanzas", key="imp_bal_dir")

    mes_str = f"{int(mes_imp):02d}"
    posibles = (
        glob.glob(os.path.join(ruta_balanzas_imp, str(anio_imp), mes_str, "*.xlsx"), recursive=True) +
        glob.glob(os.path.join(ruta_balanzas_imp, "**", mes_str, "*.xlsx"), recursive=True)
    )
    posibles_gen = [f for f in posibles if "rh" not in os.path.basename(f).lower()]
    ruta_balanza_sel = posibles_gen[0] if posibles_gen else (posibles[0] if posibles else None)

    tab_isr, tab_iva, tab_retenciones = st.tabs(["📈 ISR (Pagos Provisionales)", "💵 IVA (Cobrado vs Pagado)", "📋 RETENCIONES"])

    with tab_isr:
        st.markdown(f"### 📈 Resumen General de ISR — {nombres_meses[mes_imp-1]} {anio_imp}")

        if ruta_balanza_sel and os.path.exists(ruta_balanza_sel):
            df_saldos = obtener_saldos_balanza(ruta_balanza_sel)

            if not df_saldos.empty:
                datos_isr = []
                for emp in sorted(df_saldos['Empresa'].unique()):
                    df_emp = df_saldos[df_saldos['Empresa'] == emp]

                    v_410_f = float(df_emp[df_emp['Cuenta_Clean'].str.startswith('410')]['Saldo Acreedor Final'].sum())
                    v_410 = v_410_f if v_410_f != 0 else float(df_emp[df_emp['Cuenta_Clean'].str.startswith('410')]['Columna F (Abonos)'].sum())

                    v_411_f = float(df_emp[df_emp['Cuenta_Clean'].str.startswith('411')]['Saldo Acreedor Final'].sum())
                    v_411 = v_411_f if v_411_f != 0 else float(df_emp[df_emp['Cuenta_Clean'].str.startswith('411')]['Columna F (Abonos)'].sum())

                    df_423 = df_emp[df_emp['Cuenta_Clean'].str.startswith('423')]
                    v_423_ac = float(df_423['Saldo Acreedor Final'].sum())
                    v_423_de = float(df_423['Saldo Deudor Final'].sum())
                    v_423 = v_423_ac if v_423_ac > 0 else v_423_de

                    ingresos_isr = v_410 + v_411 - v_423

                    datos_isr.append({
                        'Empresa': emp,
                        'Cuenta 410 (Ventas/Ingresos Obra)': v_410,
                        'Cuenta 411 (Otros Ingresos)': v_411,
                        'Cuenta 423 (-) Dev. y Desc.': v_423,
                        'Ingresos Nominales ISR': ingresos_isr,
                        'Coeficiente de Utilidad': 0.0500
                    })

                df_base_isr = pd.DataFrame(datos_isr)
                st.markdown("#### ✏️ Capture o Modifique el Coeficiente de Utilidad por Empresa:")
                df_edited = st.data_editor(
                    df_base_isr[['Empresa', 'Ingresos Nominales ISR', 'Coeficiente de Utilidad']],
                    column_config={
                        "Empresa": st.column_config.TextColumn(disabled=True),
                        "Ingresos Nominales ISR": st.column_config.NumberColumn(format="$%,.2f", disabled=True),
                        "Coeficiente de Utilidad": st.column_config.NumberColumn(format="%.4f", step=0.0001, min_value=0.0, max_value=1.0)
                    },
                    hide_index=True, use_container_width=True
                )

                df_merged = pd.merge(df_base_isr.drop(columns=['Coeficiente de Utilidad']), df_edited[['Empresa', 'Coeficiente de Utilidad']], on='Empresa')
                df_merged['Utilidad Fiscal Estimada'] = df_merged['Ingresos Nominales ISR'] * df_merged['Coeficiente de Utilidad']
                df_merged['Pago Provisional ISR (30%)'] = df_merged['Utilidad Fiscal Estimada'] * 0.30

                conceptos_isr = [
                    'Cuenta 410 (Ventas/Ingresos Obra)', 'Cuenta 411 (Otros Ingresos)',
                    'Cuenta 423 (-) Dev. y Desc.', 'Ingresos Nominales ISR',
                    'Coeficiente de Utilidad', 'Utilidad Fiscal Estimada', 'Pago Provisional ISR (30%)'
                ]

                tabla_resumen_isr = pd.DataFrame({'Concepto': conceptos_isr})
                for _, row in df_merged.iterrows():
                    tabla_resumen_isr[row['Empresa']] = [row[c] for c in conceptos_isr]

                tabla_resumen_isr['TOTAL CONSOLIDADO'] = [None if c == 'Coeficiente de Utilidad' else df_merged[c].sum() for c in conceptos_isr]

                st.markdown("---")
                st.markdown("#### 📊 Tabla de Resumen Comparativo ISR — Grupo")
                st.dataframe(tabla_resumen_isr.style.format({col: "${:,.2f}" for col in tabla_resumen_isr.columns if col != 'Concepto'}, na_rep="-"), use_container_width=True, hide_index=True)
            else:
                st.warning("No se pudieron extraer saldos de la balanza seleccionada.")
        else:
            st.error(f"No se encontró archivo de balanza en el periodo {nombres_meses[mes_imp-1]} {anio_imp}.")

    with tab_iva:
        st.markdown(f"### 💵 Resumen General de IVA — {nombres_meses[mes_imp-1]} {anio_imp}")

        if ruta_balanza_sel and os.path.exists(ruta_balanza_sel):
            df_saldos_iva = obtener_saldos_balanza(ruta_balanza_sel)

            if not df_saldos_iva.empty:
                datos_iva = []
                for emp in sorted(df_saldos_iva['Empresa'].unique()):
                    df_emp_iva = df_saldos_iva[df_saldos_iva['Empresa'] == emp]

                    # Búsqueda flexible de cuentas IVA
                    v_cobrado = float(df_emp_iva[df_emp_iva['Cuenta_Clean'].str.startswith('20100010001')]['Columna F (Abonos)'].sum())
                    if v_cobrado == 0:
                        v_cobrado = float(df_emp_iva[df_emp_iva['Cuenta_Clean'].str.startswith('201') & df_emp_iva['Nombre Cuenta'].str.contains('COBRADO|TRASLADADO', case=False, na=False)]['Columna F (Abonos)'].sum())

                    v_pag1 = float(df_emp_iva[df_emp_iva['Cuenta_Clean'].str.startswith('10100011001')]['Columna E (Cargos)'].sum())
                    v_pag2 = float(df_emp_iva[df_emp_iva['Cuenta_Clean'].str.startswith('10100011002')]['Columna E (Cargos)'].sum())
                    if (v_pag1 + v_pag2) == 0:
                        v_pag1 = float(df_emp_iva[df_emp_iva['Cuenta_Clean'].str.startswith('101') & df_emp_iva['Nombre Cuenta'].str.contains('PAGADO|ACREDITABLE', case=False, na=False)]['Columna E (Cargos)'].sum())

                    v_pagado_total = v_pag1 + v_pag2
                    diferencia = v_cobrado - v_pagado_total

                    datos_iva.append({
                        'Empresa': emp,
                        'IVA Cobrado (Cuenta 201-00010-001-0000) [Col. F]': v_cobrado,
                        'IVA Pagado Gastos (Cuenta 101-00011-001-0000) [Col. E]': v_pag1,
                        'IVA Pagado Costos (Cuenta 101-00011-002-0000) [Col. E]': v_pag2,
                        'Total IVA Pagado / Acreditable': v_pagado_total,
                        'IVA a Pagar (A Cargo)': diferencia if diferencia > 0 else 0.0,
                        'IVA a Favor': abs(diferencia) if diferencia < 0 else 0.0
                    })

                df_base_iva = pd.DataFrame(datos_iva)

                conceptos_iva = [
                    'IVA Cobrado (Cuenta 201-00010-001-0000) [Col. F]',
                    'IVA Pagado Gastos (Cuenta 101-00011-001-0000) [Col. E]',
                    'IVA Pagado Costos (Cuenta 101-00011-002-0000) [Col. E]',
                    'Total IVA Pagado / Acreditable',
                    'IVA a Pagar (A Cargo)',
                    'IVA a Favor'
                ]

                tabla_resumen_iva = pd.DataFrame({'Concepto': conceptos_iva})
                for _, row in df_base_iva.iterrows():
                    tabla_resumen_iva[row['Empresa']] = [row[c] for c in conceptos_iva]

                tabla_resumen_iva['TOTAL CONSOLIDADO'] = [df_base_iva[c].sum() for c in conceptos_iva]

                st.dataframe(tabla_resumen_iva.style.format({col: "${:,.2f}" for col in tabla_resumen_iva.columns if col != 'Concepto'}), use_container_width=True, hide_index=True)
                st.markdown("---")

                excel_iva_bytes = generar_excel_iva(tabla_resumen_iva, f"{nombres_meses[mes_imp-1]} {anio_imp}")
                st.download_button(
                    label="📥 Descargar Reporte Ejecutivo de IVA en Excel",
                    data=excel_iva_bytes,
                    file_name=f"Reporte_Ejecutivo_IVA_{nombres_meses[mes_imp-1]}_{anio_imp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_iva", use_container_width=True
                )
            else:
                st.warning("No se encontraron saldos de IVA en la balanza.")
        else:
            st.error("No se encontró el archivo de balanza especificado.")

    with tab_retenciones:
        st.markdown(f"### 📋 Resumen General de Retenciones — {nombres_meses[mes_imp-1]} {anio_imp}")
        st.info("Espacio para consolidar cuentas de Retenciones.")