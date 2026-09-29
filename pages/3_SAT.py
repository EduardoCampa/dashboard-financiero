import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT - Resumen Filtrado por Periodo",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📑 Módulo SAT - Resumen por Razón Emisor (Ingresos y Egresos con Filtros)")

def cargar_y_procesar_sat(carpeta_xmls="XML", anio_filtro=None, mes_ini=None, mes_fin=None):
    """Consolida, filtra por periodo y separa los CFDis de Ingresos y Egresos."""
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
                # Validar UUID
                uuid = ""
                for col in df_sat.columns:
                    if 'uuid' in col.lower() or 'folio fiscal' in col.lower():
                        val_uuid = str(row.get(col, '')).strip().upper()
                        if val_uuid and val_uuid != 'NAN':
                            uuid = val_uuid
                            break
                if not uuid:
                    continue

                # Estado (Solo Vigentes)
                estado = ""
                for col in df_sat.columns:
                    if col.lower() == 'estado' or 'estatus' in col.lower():
                        val_est = str(row.get(col, '')).strip().upper()
                        if val_est and val_est != 'NAN':
                            estado = val_est
                            break
                
                if 'VIGENTE' not in estado:
                    continue

                # Fecha Emisión y filtrado por Periodo
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

                # Tipo (Ingreso o Egreso)
                tipo_doc = ""
                for col in df_sat.columns:
                    if col.lower() == 'tipo':
                        tipo_doc = str(row.get(col, '')).strip()
                        break

                # Razón Emisor
                razon_emisor = ""
                for col in df_sat.columns:
                    if 'razon' in col.lower() and 'emisor' in col.lower():
                        val_razon = str(row.get(col, '')).strip()
                        if val_razon and val_razon.upper() != 'NAN':
                            razon_emisor = val_razon
                            break

                # SubTotal y Descuento
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

if st.button("🚀 Ejecutar Amarre y Resumen"):
    with st.spinner("Procesando y filtrando información por periodo..."):
        df_ingresos, df_egresos = cargar_y_procesar_sat(carpeta_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)

        st.success("✅ Filtros y procesamiento completados con éxito.")

        tab_res, tab_ing, tab_eg = st.tabs([
            "📊 Resumen por Razón Emisor", 
            "📈 Detalle de Ingresos", 
            "📉 Detalle de Egresos"
        ])

        with tab_res:
            st.markdown(f"### 📊 Resumen Ejecutivo del Periodo: {mes_inicial} a {mes_final} del {anio_sel}")
            
            col_res1, col_res2 = st.columns(2)

            with col_res1:
                st.markdown("#### 📈 Resumen de Ingresos por Razón Emisor")
                if not df_ingresos.empty:
                    res_ing = df_ingresos.groupby('Razon emisor', as_index=False)['SubTotal'].sum()
                    res_ing.columns = ['Razon emisor', 'Total Ingresos']
                    st.dataframe(
                        res_ing.style.format({'Total Ingresos': '${:,.2f}'}),
                        use_container_width=True,
                        hide_index=True
                    )
                    st.metric("Total Ingresos Vigentes", f"${res_ing['Total Ingresos'].sum():,.2f}")
                else:
                    st.info("No hay ingresos vigentes en este periodo.")

            with col_res2:
                st.markdown("#### 📉 Resumen de Egresos por Razón Emisor")
                if not df_egresos.empty:
                    res_eg = df_egresos.groupby('Razon emisor', as_index=False)['SubTotal'].sum()
                    res_eg.columns = ['Razon emisor', 'Total Egresos']
                    st.dataframe(
                        res_eg.style.format({'Total Egresos': '${:,.2f}'}),
                        use_container_width=True,
                        hide_index=True
                    )
                    st.metric("Total Egresos Vigentes", f"${res_eg['Total Egresos'].sum():,.2f}")
                else:
                    st.info("No hay egresos vigentes en este periodo.")

        with tab_ing:
            st.markdown("### 📋 Detalle de Ingresos Vigentes")
            if not df_ingresos.empty:
                st.dataframe(
                    df_ingresos.style.format({'SubTotal': '${:,.2f}'}),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("No se encontraron ingresos vigentes para el periodo.")

        with tab_eg:
            st.markdown("### 📋 Detalle de Egresos Vigentes")
            if not df_egresos.empty:
                st.dataframe(
                    df_egresos.style.format({'SubTotal': '${:,.2f}'}),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("No se encontraron egresos vigentes para el periodo.")