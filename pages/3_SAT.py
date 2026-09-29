import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT y Consolidado Master",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📑 Módulo SAT - Resumen SAT y Detalle Consolidado Master")

# --- FUNCIONES DE PROCESAMIENTO SAT ---
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

                estado = ""
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

                razon_emisor = ""
                for col in df_sat.columns:
                    if 'razon' in col.lower() and 'emisor' in col.lower():
                        val_razon = str(row.get(col, '')).strip()
                        if val_razon and val_razon.upper() != 'NAN':
                            razon_emisor = val_razon
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
                    'Razon emisor': razon_emisor if razon_emisor else "SIN RAZÓN EMISOR",
                    'UUID': uuid,
                    'Fecha emision': fecha_emision_str,
                    'Estado': estado,
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

    def procesar_hoja(nombre_hoja):
        try:
            df = pd.read_excel(ruta_master, sheet_name=nombre_hoja)
            if df.empty:
                return pd.DataFrame()

            # 1. Omitir UUIDs vacíos o nulos
            if 'UUID' in df.columns:
                df = df[df['UUID'].notnull() & (df['UUID'].astype(str).str.strip() != '') & (df['UUID'].astype(str).str.upper() != 'NAN')].copy()
            else:
                return pd.DataFrame()

            # 2. Filtrar por estatus Vigente
            if 'CFDStatusCancelledName' in df.columns:
                df = df[df['CFDStatusCancelledName'].astype(str).str.strip().str.upper() == 'VIGENTE'].copy()

            # 3. Filtrar por fecha y rango (CFDIFechaCertificacion)
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

            # 4. Calcular SubTotal - TotalDiscount
            subtotal = pd.to_numeric(df.get('SubTotal', 0), errors='coerce').fillna(0.0)
            descuento = pd.to_numeric(df.get('TotalDiscount', 0), errors='coerce').fillna(0.0)
            df['SubTotal_Neto'] = subtotal - descuento

            # 5. Seleccionar columnas requeridas
            empresa = df['EmpresaOrigen'] if 'EmpresaOrigen' in df.columns else 'SIN EMPRESA'
            uuid_col = df['UUID']
            fecha_col = df['Fecha_Cert']
            status_col = df['CFDStatusCancelledName']
            neto_col = df['SubTotal_Neto']

            df_final = pd.DataFrame({
                'EmpresaOrigen': empresa,
                'UUID': uuid_col,
                'CFDIFechaCertificacion': fecha_col,
                'CFDStatusCancelledName': status_col,
                'SubTotal': neto_col
            })

            return df_final
        except Exception as e:
            return pd.DataFrame()

    df_ingresos_master = procesar_hoja("FacturaCliente")
    df_egresos_master = procesar_hoja("NotaCreditoCliente")

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

if st.button("🚀 Ejecutar Análisis SAT y Master"):
    with st.spinner("Procesando información del SAT y del Consolidado Master..."):
        # 1. Datos SAT
        df_ingresos_sat, df_egresos_sat = cargar_y_procesar_sat(carpeta_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
        
        # 2. Datos Consolidado Master
        df_ingresos_master, df_egresos_master = cargar_y_procesar_master(ruta_master_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)

        st.success("✅ Procesamiento completado correctamente.")

        tab_sat, tab_master = st.tabs([
            "📊 Módulo SAT (Ingresos / Egresos)", 
            "📁 Detalle Consolidado Master"
        ])

        with tab_sat:
            st.markdown(f"### 📊 Resumen SAT - Periodo: {mes_inicial} a {mes_final} del {anio_sel}")
            col_res1, col_res2 = st.columns(2)

            with col_res1:
                st.markdown("#### 📈 Ingresos por Razón Emisor")
                if not df_ingresos_sat.empty:
                    res_ing = df_ingresos_sat.groupby('Razon emisor', as_index=False)['SubTotal'].sum()
                    res_ing.columns = ['Razon emisor', 'Total Ingresos']
                    st.dataframe(res_ing.style.format({'Total Ingresos': '${:,.2f}'}), use_container_width=True, hide_index=True)
                    st.metric("Total Ingresos Vigentes", f"${res_ing['Total Ingresos'].sum():,.2f}")
                else:
                    st.info("No hay ingresos vigentes.")

            with col_res2:
                st.markdown("#### 📉 Egresos por Razón Emisor")
                if not df_egresos_sat.empty:
                    res_eg = df_egresos_sat.groupby('Razon emisor', as_index=False)['SubTotal'].sum()
                    res_eg.columns = ['Razon emisor', 'Total Egresos']
                    st.dataframe(res_eg.style.format({'Total Egresos': '${:,.2f}'}), use_container_width=True, hide_index=True)
                    st.metric("Total Egresos Vigentes", f"${res_eg['Total Egresos'].sum():,.2f}")
                else:
                    st.info("No hay egresos vigentes.")

        with tab_master:
            st.markdown(f"### 📁 Detalle Consolidado Master (Vigentes con UUID) - Periodo: {mes_inicial} a {mes_final} del {anio_sel}")
            
            st.markdown("#### 📈 Detalle de Ingresos (FacturaCliente)")
            if not df_ingresos_master.empty:
                st.dataframe(df_ingresos_master.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                st.metric("Suma SubTotal Neto Ingresos Master", f"${df_ingresos_master['SubTotal'].sum():,.2f}")
            else:
                st.warning("No se encontraron registros de ingresos vigentes en el Master.")

            st.markdown("---")
            st.markdown("#### 📉 Detalle de Egresos (NotaCreditoCliente)")
            if not df_egresos_master.empty:
                st.dataframe(df_egresos_master.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                st.metric("Suma SubTotal Neto Egresos Master", f"${df_egresos_master['SubTotal'].sum():,.2f}")
            else:
                st.warning("No se encontraron registros de egresos vigentes en el Master.")