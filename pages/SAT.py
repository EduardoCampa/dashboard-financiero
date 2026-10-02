import glob
import os
import io
import re
import pandas as pd
import streamlit as st

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

                # --- BOTÓN DE EXPORTACIÓN A EXCEL ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_ingresos_final.to_excel(writer, sheet_name='Resumen Ingresos', index=False)
                    df_egresos_final.to_excel(writer, sheet_name='Resumen Egresos', index=False)
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Descargar Reporte Ejecutivo en Excel",
                    data=excel_data,
                    file_name=f"Reporte_Ejecutivo_SAT_{nombres_meses[mes_final-1]}_{anio_sel}.xlsx",
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
                            'Contabilidad Ingresos': '${:,.2f}', 
                            'Contabilidad Egresos': '${:,.2f}'
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
                            'Saldo Deudor Final': '${:,.2f}', 
                            'Saldo Acreedor Final': '${:,.2f}'
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
                        df['Empresa_RFC_Clean'] = df[col_rfc].astype(str).str.strip().str.upper()
                    elif rfc_archivo:
                        df['Empresa_RFC_Clean'] = rfc_archivo
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


    # --- 2. PROCESAMIENTO CONTABILIDAD NÓMINAS (CUENTA COMPLETA, DEPTOS 001, 029, 056, SIN 0000) ---
    def cargar_y_procesar_nomina_contabilidad(ruta_base="Balanzas", anio_filtro=2026, mes_ini=1, mes_fin=8):
        registros_contables = []

        # Recorrer mes a mes dentro del rango para acumular la Columna E (Cargos del Mes)
        for m in range(int(mes_ini), int(mes_fin) + 1):
            mes_str = f"{m:02d}"
            posibles = glob.glob(os.path.join(ruta_base, str(anio_filtro), mes_str, "*.xlsx")) + \
                       glob.glob(os.path.join(ruta_base, "**", mes_str, "*.xlsx"), recursive=True) + \
                       glob.glob(os.path.join(ruta_base, "*.xlsx"))
            
            posibles = list(set(posibles))
            if not posibles:
                continue

            for ruta_balanza in posibles:
                try:
                    nombre_archivo = os.path.basename(ruta_balanza).upper()
                    xls = pd.ExcelFile(ruta_balanza)

                    for hoja in xls.sheet_names:
                        hoja_upper = hoja.strip().upper()

                        # Identificación de la empresa
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

                        # Localizar Columna E: Cargos del Mes (Índice 4 por defecto)
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

                            # PATRÓN DE VALIDACIÓN DE CUENTA:
                            # 1. Grupo principal: 500 a 699 (\d{3})
                            # 2. Obra / Subnivel 2
                            # 3. Departamento Nivel 3: 001, 029 o 056
                            # 4. Nivel 4 (Subcuenta): Excluye explícitamente el 0000 (?!0000)
                            match = re.match(r'^(5\d\d|6\d\d)-(\d+)-(001|029|056)-(?!0000)(\d{4})$', cta_limpia)

                            if match:
                                # Conservar el código completo de la cuenta tal como en la balanza
                                label_cuenta = f"{cta_limpia} {nom_cuenta}".strip()

                                val_cargo = row.iloc[idx_cargos_e] if df_hoja.shape[1] > idx_cargos_e else 0.0
                                monto_cargo = float(pd.to_numeric(str(val_cargo).replace('$', '').replace(',', ''), errors='coerce') or 0.0)

                                if monto_cargo > 0.0:
                                    registros_contables.append({
                                        'Cuenta Contable': label_cuenta,
                                        'Empresa': nombre_empresa,
                                        'Monto': monto_cargo
                                    })
                except Exception:
                    continue

        df_cont = pd.DataFrame(registros_contables)
        if df_cont.empty:
            return pd.DataFrame()

        # Crear la tabla pivote con Cuentas completas en filas y Empresas en columnas
        pivot_cont = pd.pivot_table(
            df_cont,
            index='Cuenta Contable',
            columns='Empresa',
            values='Monto',
            aggfunc='sum',
            fill_value=0.0
        ).reset_index()

        return pivot_cont


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

    if st.button("🚀 Ejecutar Amarre Nóminas"):
        with st.spinner("Procesando archivos del SAT y Balanzas Contables para Nómina..."):
            
            # Cargar SAT Nómina
            df_nomina_sat = cargar_y_procesar_nomina_sat(
                carpeta_nomina=carpeta_nomina_input,
                anio_filtro=anio_sel_nom,
                mes_ini=mes_inicial_nom,
                mes_fin=mes_final_nom
            )

            # Cargar Contabilidad Nómina mostrando cuentas completas
            df_nomina_cont = cargar_y_procesar_nomina_contabilidad(
                ruta_base=ruta_balanzas_nom_input,
                anio_filtro=anio_sel_nom,
                mes_ini=mes_inicial_nom,
                mes_fin=mes_final_nom
            )

            st.success("✅ Procesamiento de Amarre de Nóminas completado.")

            subtab_nom_sat, subtab_nom_cont = st.tabs(["📑 SAT", "📊 Contabilidad"])

            # SUBPESTAÑA 1: SAT NÓMINA
            with subtab_nom_sat:
                st.markdown(f"#### 📋 Consolidado de Nómina SAT por RFC, Mes y Conceptos ({nombres_meses[mes_inicial_nom-1]} a {nombres_meses[mes_final_nom-1]} {anio_sel_nom})")
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

            # SUBPESTAÑA 2: CONTABILIDAD NÓMINA
            with subtab_nom_cont:
                st.markdown(f"#### 📊 Consolidado Contable de Cuentas (Grupo 500 a 699, Deptos -001-, -029- y -056-, sin -0000) ({nombres_meses[mes_inicial_nom-1]} a {nombres_meses[mes_final_nom-1]} {anio_sel_nom})")
                if not df_nomina_cont.empty:
                    empresas_cols = [col for col in df_nomina_cont.columns if col != 'Cuenta Contable']
                    format_dict_c = {col: '${:,.2f}' for col in empresas_cols}

                    st.dataframe(
                        df_nomina_cont.style.format(format_dict_c),
                        use_container_width=True,
                        hide_index=True
                    )

                    output_cont_nom = io.BytesIO()
                    with pd.ExcelWriter(output_cont_nom, engine='openpyxl') as writer:
                        df_nomina_cont.to_excel(writer, sheet_name='Contabilidad Nómina', index=False)
                    excel_data_cont_nom = output_cont_nom.getvalue()

                    st.download_button(
                        label="📥 Descargar Consolidado Contable de Nómina en Excel",
                        data=excel_data_cont_nom,
                        file_name=f"Consolidado_Contable_Nomina_{nombres_meses[mes_final_nom-1]}_{anio_sel_nom}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No se encontraron cuentas del grupo 500 al 699 con nivel 3 (-001- / -029- / -056-) en las balanzas del periodo seleccionado.")