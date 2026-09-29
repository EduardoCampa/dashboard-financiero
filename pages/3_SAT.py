import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT y Conciliación Master vs XML",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📑 Módulo SAT - Comparativo Master vs XML, Detalle y Conciliación")

# --- DICCIONARIO OFICIAL DE MAPEO DE RAZONES SOCIALES Y EMPRESAS ---
MAPEO_RAZON_A_EMPRESA = {
    "COMERCIALIZADORA DE INFRAESTRUCTURA VIAL": "CIV",
    "COMERCIALIZADORA DE INFRAESTRUCTURA VIAL LATINOAMERICANA": "CIVLAT",
    "ELEMENTOS FABRICADOS Y CONSTRUCCIONES": "EFCO",
    "FGS SISTEMAS INTEGRALES DE MANTENIMIENTO": "FGS",
    "GRUPO FERVIC": "FERVIC",
    "GRUPO SERVYRE": "GRUPOSERVYRE",
    "INMOBILIARIA PSZ": "INMOBILIARIA",
    "LABORATORIO MAJONINMAR": "LABORATORIO",
    "LATIN AMERICAN ITS": "LAITS",
    "GRUPO PESAZA": "PESAZA"
}

MAPEO_EMPRESA_ORIGEN = {
    "CIV": "CIV",
    "CIVLA": "CIVLAT",
    "CIVLAT": "CIVLAT",
    "CIVMEX": "CIVMEX",
    "EFCO": "EFCO",
    "FERVIC": "FERVIC",
    "FGS": "FGS",
    "GRUPO FPSB": "GRUPO FPSB",
    "GRUPOSERVYRE": "GRUPOSERVYRE",
    "INMOBILIARIA": "INMOBILIARIA",
    "LABORATORIO": "LABORATORIO",
    "LAITS": "LAITS",
    "LIMPIESPIN": "LIMPIESPIN",
    "PESAZA": "PESAZA",
    "SERSENAL": "SERSENAL",
    "SERVYCARGO": "SERVYCARGO",
    "SERVYRE": "SERVYRE"
}

def normalizar_empresa(razon_o_empresa):
    val = str(razon_o_empresa).strip().upper()
    if not val or val in ('NAN', '0', '0.0', '0.00000', 'NONE'):
        return "OTRAS"
    if val in MAPEO_EMPRESA_ORIGEN:
        return MAPEO_EMPRESA_ORIGEN[val]
    for razon, empresa_corta in MAPEO_RAZON_A_EMPRESA.items():
        if razon in val or empresa_corta in val:
            return empresa_corta
    return val


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
                            razon_emisor = val_razon
                            break

                empresa_normalizada = normalizar_empresa(razon_emisor)
                if empresa_normalizada == "OTRAS":
                    base_nom = os.path.basename(archivo).upper()
                    for k, v in MAPEO_EMPRESA_ORIGEN.items():
                        if k in base_nom:
                            empresa_normalizada = v
                            break

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
                    'Empresa': empresa_normalizada,
                    'Razon emisor': razon_emisor if razon_emisor else "SIN RAZÓN EMISOR",
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

            if 'CFDIFechaCertificacion' in df.columns:
                df['Fecha_Cert'] = pd.to_datetime(df['CFDIFechaCertificacion'], errors='coerce')
                df = df[df['Fecha_Cert'].notnull()].copy()
                if anio_filtro:
                    df = df[df['Fecha_Cert'].dt.year == int(anio_filtro)]
                if mes_ini and mes_fin:
                    df = df[(df['Fecha_Cert'].dt.month >= int(mes_ini)) & (df['Fecha_Cert'].dt.month <= int(mes_fin))]
                df['Fecha_Cert'] = df['Fecha_Cert'].dt.strftime('%Y-%m-%d')
            else:
                df['Fecha_Cert'] = ''

            subtotal = pd.to_numeric(df.get('SubTotal', 0), errors='coerce').fillna(0.0)
            descuento = pd.to_numeric(df.get('TotalDiscount', 0), errors='coerce').fillna(0.0)
            subtotal_neto = subtotal - descuento

            status_canc = df.get('CFDStatusCancelledName', 'VIGENTE')
            df['Estado_Master'] = status_canc.apply(lambda x: str(x).strip().upper() if pd.notnull(x) else 'VIGENTE')

            if es_egreso:
                subtotal_neto = subtotal_neto.abs() * -1

            empresa_raw = df['EmpresaOrigen'] if 'EmpresaOrigen' in df.columns else 'SIN EMPRESA'
            empresa_norm = empresa_raw.apply(normalizar_empresa)

            df_final = pd.DataFrame({
                'Empresa': empresa_norm,
                'UUID': df['UUID'],
                'CFDIFechaCertificacion': df['Fecha_Cert'],
                'Estado_Master': df['Estado_Master'],
                'SubTotal': subtotal_neto
            })

            return df_final
        except Exception as e:
            return pd.DataFrame()

    df_ingresos_master = procesar_hoja("FacturaCliente", es_egreso=False)
    df_egresos_master = procesar_hoja("NotaCreditoCliente", es_egreso=True)

    return df_ingresos_master, df_egresos_master


# --- CONTROLES DE FILTRO POR PERIODO ---
col_a, col_mini, col_mfin = st.columns([1, 1, 1])

with col_a:
    anio_sel = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0)
with col_mini:
    mes_inicial = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'][x-1])
with col_mfin:
    mes_final = st.selectbox("Mes Final:", list(range(1, 13)), index=7, format_func=lambda x: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'][x-1])

st.markdown("---")
carpeta_input = st.text_input("Carpeta o ubicación de los archivos del SAT:", value="XML")
ruta_master_input = st.text_input("Archivo Consolidado Master:", value="Consolidado_Master.xlsx")

if st.button("🚀 Ejecutar Procesamiento Completo"):
    with st.spinner("Procesando información y unificando empresas..."):
        df_ingresos_sat, df_egresos_sat = cargar_y_procesar_sat(carpeta_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
        df_ingresos_master, df_egresos_master = cargar_y_procesar_master(ruta_master_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)

        st.success("✅ Procesamiento completado con éxito.")

        tab_comp, tab_xml, tab_master, tab_conc = st.tabs([
            "⚖️ 1.- Comparativo Master vs XML", 
            "📑 2.- Ingresos y Egresos de los XML", 
            "📁 3.- Ingresos y Egresos Master",
            "🔍 4.- Conciliacion por UUID"
        ])

        with tab_comp:
            st.markdown(f"### ⚖️️ Comparativo Master vs XML - Periodo: {mes_inicial} a {mes_final} del {anio_sel}")
            
            sat_vig = df_ingresos_sat[df_ingresos_sat['Estado_SAT'].str.contains('VIGENTE', na=False)]
            mast_vig = df_ingresos_master[df_ingresos_master['Estado_Master'].str.contains('VIGENTE', na=False)]

            res_ing_sat = sat_vig.groupby('Empresa', as_index=False)['SubTotal'].sum() if not sat_vig.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
            res_ing_sat.columns = ['Empresa', 'XML Ingresos']

            res_ing_mast = mast_vig.groupby('Empresa', as_index=False)['SubTotal'].sum() if not mast_vig.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
            res_ing_mast.columns = ['Empresa', 'Master Ingresos']

            st.markdown("#### 📈 Comparativo de Ingresos (Master vs XML)")
            if not res_ing_sat.empty or not res_ing_mast.empty:
                df_comp_ing = pd.merge(res_ing_mast, res_ing_sat, on='Empresa', how='outer').fillna(0.0)
                df_comp_ing['Diferencia (Master - XML)'] = df_comp_ing['Master Ingresos'] - df_comp_ing['XML Ingresos']
                st.dataframe(
                    df_comp_ing.style.format({
                        'Master Ingresos': '${:,.2f}',
                        'XML Ingresos': '${:,.2f}',
                        'Diferencia (Master - XML)': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No hay datos de ingresos para comparar.")

            st.markdown("---")
            sat_eg_vig = df_egresos_sat[df_egresos_sat['Estado_SAT'].str.contains('VIGENTE', na=False)]
            mast_eg_vig = df_egresos_master[df_egresos_master['Estado_Master'].str.contains('VIGENTE', na=False)]

            res_eg_sat = sat_eg_vig.groupby('Empresa', as_index=False)['SubTotal'].sum() if not sat_eg_vig.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
            res_eg_sat.columns = ['Empresa', 'XML Egresos']

            res_eg_mast = mast_eg_vig.groupby('Empresa', as_index=False)['SubTotal'].sum() if not mast_eg_vig.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
            res_eg_mast.columns = ['Empresa', 'Master Egresos']

            st.markdown("#### 📉 Comparativo de Egresos / Notas de Crédito (Master vs XML)")
            if not res_eg_sat.empty or not res_eg_mast.empty:
                df_comp_eg = pd.merge(res_eg_mast, res_eg_sat, on='Empresa', how='outer').fillna(0.0)
                df_comp_eg['Diferencia (Master - XML)'] = df_comp_eg['Master Egresos'] - df_comp_eg['XML Egresos']
                st.dataframe(
                    df_comp_eg.style.format({
                        'Master Egresos': '${:,.2f}',
                        'XML Egresos': '${:,.2f}',
                        'Diferencia (Master - XML)': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No hay datos de egresos para comparar.")

        with tab_xml:
            st.markdown(f"### 📑 Ingresos y Egresos de los XML - Periodo: {mes_inicial} a {mes_final} del {anio_sel}")
            subtab_x_ing, subtab_x_eg = st.tabs(["📈 Ingresos XML", "📉 Egresos XML"])
            
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
            st.markdown(f"### 📁 Ingresos y Egresos Master - Periodo: {mes_inicial} a {mes_final} del {anio_sel}")
            subtab_m_ing, subtab_m_eg = st.tabs(["📈 Ingresos Master", "📉 Egresos Master"])
            
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
            st.markdown(f"### 🔍 Conciliación por UUID (Ingresos) - Periodo: {mes_inicial} a {mes_final} del {anio_sel}")
            
            if not df_ingresos_master.empty and not df_ingresos_sat.empty:
                df_m_agg = df_ingresos_master[['UUID', 'Empresa', 'SubTotal', 'Estado_Master']].rename(
                    columns={'SubTotal': 'SubTotal_Master'}
                )
                df_s_agg = df_ingresos_sat[['UUID', 'Empresa', 'SubTotal', 'Estado_SAT']].rename(
                    columns={'SubTotal': 'SubTotal_SAT', 'Empresa': 'Empresa_SAT'}
                )
                
                df_concil = pd.merge(df_m_agg, df_s_agg, on='UUID', how='outer')
                df_concil['SubTotal_Master'] = df_concil['SubTotal_Master'].fillna(0.0)
                df_concil['SubTotal_SAT'] = df_concil['SubTotal_SAT'].fillna(0.0)
                df_concil['Empresa'] = df_concil['Empresa'].fillna(df_concil['Empresa_SAT']).fillna('OTRAS')
                df_concil['Estado_Master'] = df_concil['Estado_Master'].fillna('NO ENCONTRADO EN MASTER')
                df_concil['Estado_SAT'] = df_concil['Estado_SAT'].fillna('NO ENCONTRADO EN SAT')
                
                df_concil['Diferencia'] = df_concil['SubTotal_Master'] - df_concil['SubTotal_SAT']
                
                # Filtrar ESTRICTAMENTE solo aquellos UUIDs que tienen diferencia real
                df_con_dif = df_concil[
                    (df_concil['Diferencia'].round(2) != 0.0) | 
                    (df_concil['Estado_Master'] != df_concil['Estado_SAT'])
                ].copy()

                st.markdown("#### 📊 Resumen de UUIDs con Diferencias o Discrepancias")
                col_d1, col_d2 = st.columns(2)
                col_d1.metric("Total UUIDs Analizados", len(df_concil))
                col_d2.metric("UUIDs con Diferencia Estricta", len(df_con_dif))

                st.markdown("---")
                st.markdown("#### 📋 Detalle Exclusivo de UUIDs con Diferencias")
                if not df_con_dif.empty:
                    st.dataframe(
                        df_con_dif[[
                            'UUID', 'Empresa', 'Estado_Master', 'Estado_SAT', 
                            'SubTotal_Master', 'SubTotal_SAT', 'Diferencia'
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