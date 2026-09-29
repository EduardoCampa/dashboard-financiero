import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT - Ingresos y Egresos",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📑 Módulo SAT - Tablas de Ingresos y Egresos Vigentes")

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
                    continue # Excluir todo lo que no sea vigente

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
                    'Razon emisor': razon_emisor,
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

if st.button("🚀 Cargar y Clasificar CFDis"):
    with st.spinner("Procesando y filtrando comprobantes vigentes..."):
        df_ingresos, df_egresos = cargar_y_procesar_sat(carpeta_input)

        st.success("✅ Datos procesados correctamente.")

        tab_ing, tab_eg = st.tabs(["📈 Tabla de Ingresos", "📉 Tabla de Egresos"])

        with tab_ing:
            st.markdown("### Comprobantes de Ingresos Vigentes")
            if not df_ingresos.empty:
                st.dataframe(
                    df_ingresos.style.format({'SubTotal': '${:,.2f}'}),
                    use_container_width=True,
                    hide_index=True
                )
                st.info(f"Total de registros de ingresos vigentes: {len(df_ingresos)} | Suma SubTotal Neto: ${df_ingresos['SubTotal'].sum():,.2f}")
            else:
                st.warning("No se encontraron ingresos vigentes.")

        with tab_eg:
            st.markdown("### Comprobantes de Egresos Vigentes")
            if not df_egresos.empty:
                st.dataframe(
                    df_egresos.style.format({'SubTotal': '${:,.2f}'}),
                    use_container_width=True,
                    hide_index=True
                )
                st.info(f"Total de registros de egresos vigentes: {len(df_egresos)} | Suma SubTotal Neto: ${df_egresos['SubTotal'].sum():,.2f}")
            else:
                st.warning("No se encontraron egresos vigentes.")