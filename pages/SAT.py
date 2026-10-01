import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT y Conciliación Master vs XML",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📑 Módulo SAT - Comparativos y Conciliación (Lectura Directa)")


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
                
                if 'VIGENTE' not in estado:
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

                # Lectura directa del nombre del emisor
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

        except Exception as e:
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
                df = df[df['CFDStatusCancelledName'].astype(str).str.strip().str.upper() == 'VIGENTE'].copy()

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

            # Lectura directa de EmpresaOrigen
            empresa_raw = df['EmpresaOrigen'] if 'EmpresaOrigen' in df.columns else 'SIN EMPRESA'
            df['Empresa'] = empresa_raw.astype(str).str.strip().str.upper()

            df_final = pd.DataFrame({
                'Empresa': df['Empresa'],
                'UUID': df['UUID'],
                'CFDIFechaCertificacion': df['Fecha_Doc'],
                'Estado_Master': 'VIGENTE',
                'SubTotal': subtotal_neto
            })

            return df_final
        except Exception as e:
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
            
            # Lectura directa del nombre de la pestaña del Excel de balanza
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
    except Exception as e:
        pass

    df_det = pd.DataFrame(detalles)
    if df_det.empty:
        return pd.DataFrame(columns=['Empresa', 'Contabilidad Ingresos', 'Contabilidad Egresos']), df_det

    # Cálculo directo agrupado por el nombre exacto de la pestaña
    res_list = []
    empresas_unicas = df_det['Empresa'].unique()
    for empresa in empresas_unicas:
        df_emp = df_det[df_det['Empresa'] == empresa]
        
        v_410 = df_emp[df_emp['Cuenta'].str.startswith('410-00000-000-0000')]['Saldo Acreedor Final'].sum()
        v_411 = df_emp[df_emp['Cuenta'].str.startswith('411-00000-000-0000')]['Saldo Acreedor Final'].sum()
        v_423 = df_emp[df_emp['Cuenta'].str.startswith('423-00000-000-0000')]['Saldo Acreedor Final'].sum()
        
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


# --- CONTROLES DE FILTRO POR PERIODO ---
nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

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

if st.button("🚀 Ejecutar Procesamiento Completo"):
    with st.spinner("Procesando información, balanzas por mes y conciliando..."):
        df_ingresos_sat, df_egresos_sat = cargar_y_procesar_sat(carpeta_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
        df_ingresos_master, df_egresos_master = cargar_y_procesar_master(ruta_master_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
        
        df_balanzas, df_detalle_cont = cargar_balanzas_por_mes(anio=anio_sel, mes=mes_final, ruta_base=ruta_balanzas_input)

        st.success("✅ Procesamiento completado con éxito.")

        tab_comp, tab_xml, tab_master, tab_conc, tab_cont = st.tabs([
            "⚖️ 1.- Comparativos", 
            "📑 2.- Ingresos y Egresos de los XML", 
            "📁 3.- Ingresos y Egresos Master",
            "🔍 4.- Conciliacion por UUID",
            "📊 5.- Contabilidad"
        ])

        with tab_comp:
            st.markdown(f"### ⚖ Comparativos Generales - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
            st.info("Pestaña de comparativos limpia.")

        with tab_xml:
            st.markdown(f"### 📑 Ingresos y Egresos de los XML (Vigentes) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
            
            st.markdown("#### 📊 Resumen de Ingresos y Egresos por Emisor (XML)")
            if not df_ingresos_sat.empty or not df_egresos_sat.empty:
                res_ing_x = df_ingresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                res_ing_x.columns = ['Empresa', 'Ingresos XML']
                
                res_eg_x = df_egresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                res_eg_x.columns = ['Empresa', 'Egresos XML']

                df_res_xml = pd.merge(res_ing_x, res_eg_x, on='Empresa', how='outer').fillna(0.0)
                st.dataframe(
                    df_res_xml.style.format({
                        'Ingresos XML': '${:,.2f}',                         'Egresos XML': '${:,.2f}'
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
            st.markdown(f"### 📁 Ingresos y Egresos Master (Vigentes) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
            
            st.markdown("#### 📊 Resumen de Ingresos y Egresos por Empresa Origen (Master)")
            if not df_ingresos_master.empty or not df_egresos_master.empty:
                res_ing_m = df_ingresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                res_ing_m.columns = ['Empresa', 'Ingresos Master']
                
                res_eg_m = df_egresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
                res_eg_m.columns = ['Empresa', 'Egresos Master']

                df_res_mas = pd.merge(res_ing_m, res_eg_m, on='Empresa', how='outer').fillna(0.0)
                st.dataframe(
                    df_res_mas.style.format({
                        'Ingresos Master': '${:,.2f}',                         'Egresos Master': '${:,.2f}'
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
            st.markdown(f"### 🔍 Conciliación por UUID (Ingresos Vigentes) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
            
            if not df_ingresos_master.empty and not df_ingresos_sat.empty:
                df_m_agg = df_ingresos_master[['UUID', 'Empresa', 'SubTotal']].rename(
                    columns={'SubTotal': 'SubTotal_Master'}
                )
                df_s_agg = df_ingresos_sat[['UUID', 'Empresa', 'SubTotal']].rename(
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
                            'UUID', 'Empresa', 'SubTotal_Master', 'SubTotal_SAT', 'Diferencia'
                        ]].style.format({
                            'SubTotal_Master': '${:,.2f}',
                            'SubTotal_SAT': '${:,.2f}',                             'Diferencia': '${:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.success("🎉 ¡Excelente! No se encontraron diferencias en los UUIDs vigentes cruzados.")
            else:
                st.info("Se requiere información tanto del Master como de los XMLs para ejecutar la conciliación.")

        with tab_cont:
            st.markdown(f"### 📊 Contabilidad y Balanzas (Mes: {nombres_meses[mes_final-1]} {anio_sel})")
            st.markdown("Fórmula aplicada:\n* **Ingresos:** `410-00000-000-0000` (+) `411-00000-000-0000` (-) `423-00000-000-0000`\n* **Egresos:** `420-00000-000-0000` (+) `421-00000-000-0000`")
            
            st.markdown("#### 📊 Resumen Contable por Pestaña de Balanza")
            if not df_balanzas.empty:
                st.dataframe(
                    df_balanzas.style.format({
                        'Contabilidad Ingresos': '${:,.2f}',                         'Contabilidad Egresos': '${:,.2f}'
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
                        'Saldo Deudor Final': '${:,.2f}',                         'Saldo Acreedor Final': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("No se encontraron registros contables para las cuentas especificadas en el periodo seleccionado.")