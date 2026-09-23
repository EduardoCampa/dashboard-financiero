import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Módulo Contable", layout="wide")

st.title("📊 Estados Financieros")

archivo_balanzas = "Consolidado_Master.xlsx"


# 1. Función con caché para obtener solo la lista de pestañas (empresas)
@st.cache_data(ttl=600)
def obtener_lista_empresas(ruta):
    xls = pd.ExcelFile(ruta)
    return xls.sheet_names


# 2. Función con caché para cargar la pestaña seleccionada sin saturar la memoria
@st.cache_data(ttl=600)
def cargar_hoja_empresa(ruta, nombre_hoja):
    return pd.read_excel(ruta, sheet_name=nombre_hoja)


if os.path.exists(archivo_balanzas):
    try:
        lista_empresas = obtener_lista_empresas(archivo_balanzas)

        # Selector de empresa
        empresa_seleccionada = st.selectbox(
            "Selecciona la Empresa:", lista_empresas
        )

        if empresa_seleccionada:
            # Carga rápida utilizando caché
            with st.spinner(
                f"Cargando datos de {empresa_seleccionada}..."
            ):
                df_empresa = cargar_hoja_empresa(
                    archivo_balanzas, empresa_seleccionada
                )

            st.subheader(f"Estado Financiero - {empresa_seleccionada}")

            # Mostrar la tabla de forma eficiente
            st.dataframe(
                df_empresa,
                use_container_width=True,
                hide_index=True,
            )

            # Botón de descarga preparado sobre demanda
            csv = df_empresa.to_csv(index=False).encode("utf-8")
            st.download_button(
                label=f"📥 Descargar Estado Financiero ({empresa_seleccionada})",
                data=csv,
                file_name=f"Estado_Financiero_{empresa_seleccionada}.csv",
                mime="text/csv",
            )
    except Exception as e:
        st.error(f"Error al procesar el archivo Excel: {e}")
else:
    st.warning(
        f"No se encontró el archivo '{archivo_balanzas}'. Asegúrate de ejecutar primero la consolidación."
    )