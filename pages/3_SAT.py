import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT - Resumen por Razón Emisor",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📑 Módulo SAT - Resumen Consolidado por Razón Emisor (Ingresos y Egresos)")

def cargar_y_procesar_sat(carpeta_xmls="XML"):
    """Consolida y separa los CFDis de Ingresos y Egresos desde los archivos Excel del SAT."""
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

                # Fecha Emisión (Solo Fecha)
                fecha_emision = ""
                for col in df_sat.columns:
                    if 'fecha' in col.lower() and 'emision' in col.lower():
                        val_fecha = row.get(col, '')
                        if pd.notnull(val_fecha):
                            dt = pd.to_datetime(val_fecha, errors='coerce')
                            if pd.notnull(dt):
                                fecha_emision = dt.strftime('%Y-%m-%d')
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
                    'Fecha emision': fecha_emision,
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

# Interfaz en Streamlit
carpeta_input = st.text_input("Carpeta o ubicación de los archivos del SAT:", value="XML")

if st.button("🚀 Cargar, Clasificar y Resumir"):
    with st.spinner("Procesando y generando tabla resumen por Razón Emisor..."):
        df_ingresos, df_egresos = cargar_y_procesar_sat(carpeta_input)

        st.success("✅ Datos procesados correctamente.")

        tab_res, tab_ing, tab_eg = st.tabs([
            "📊 Resumen por Razón Emisor", 
            "📈 Detalle de Ingresos", 
            "📉 Detalle de Egresos"
        ])

        with tab_res:
            st.markdown("### 📊 Resumen de Ingresos y Egresos Vigentes por Razón Emisor")
            
            if not df_ingresos.empty or not df_egresos.empty:
                # Agrupar ingresos
                df_ing_sum = pd.DataFrame(columns=['Razon emisor', 'Total Ingresos'])
                if not df_ingresos.empty:
                    df_ing_sum = df_ingresos.groupby('Razon emisor', as_index=False)['SubTotal'].sum()
                    df_ing_sum.columns = ['Razon emisor', 'Total Ingresos']

                # Agrupar egresos
                df_eg_sum = pd.DataFrame(columns=['Razon emisor', 'Total Egresos'])
                if not df_egresos.empty:
                    df_eg_sum = df_egresos.groupby('Razon emisor', as_index=False)['SubTotal'].sum()
                    df_eg_sum.columns = ['Razon emisor', 'Total Egresos']

                # Unir ambas tablas por Razón Emisor (Outer join)
                df_resumen = pd.merge(df_ing_sum, df_eg_sum, on='Razon emisor', how='outer').fillna(0.0)
                df_resumen['Neto (Ingresos - Egresos)'] = df_resumen['Total Ingresos'] - df_resumen['Total Egresos']

                st.dataframe(
                    df_resumen.style.format({
                        'Total Ingresos': '${:,.2f}',
                        'Total Egresos': '${:,.2f}',
                        'Neto (Ingresos - Egresos)': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                tot_ing = df_resumen['Total Ingresos'].sum()
                tot_eg = df_resumen['Total Egresos'].sum()
                neto_gen = df_resumen['Neto (Ingresos - Egresos)'].sum()

                col_m1, col_m2, col_m3 = st.columns(3)
                col_m1.metric("Suma Total Ingresos Vigentes", f"${tot_ing:,.2f}")
                col_m2.metric("Suma Total Egresos Vigentes", f"${tot_eg:,.2f}")
                col_m3.metric("Balance Neto General", f"${neto_gen:,.2f}")
            else:
                st.info("No se encontraron registros vigentes para resumir.")

        with tab_ing:
            st.markdown("### 📋 Detalle de Comprobantes de Ingresos Vigentes")
            if not df_ingresos.empty:
                st.dataframe(
                    df_ingresos.style.format({'SubTotal': '${:,.2f}'}),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("No se encontraron ingresos vigentes.")

        with tab_eg:
            st.markdown("### 📋 Detalle de Comprobantes de Egresos Vigentes")
            if not df_egresos.empty:
                st.dataframe(
                    df_egresos.style.format({'SubTotal': '${:,.2f}'}),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("No se encontraron egresos vigentes.")