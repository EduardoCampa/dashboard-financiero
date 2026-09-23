import os
import glob
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Módulo Contable", layout="wide")

st.title("📊 Estados Financieros")


# Función para buscar de forma automática el archivo de Balanza más reciente en la carpeta Balanzas
def obtener_ruta_balanza():
    # Busca cualquier archivo Balanza.xlsx dentro de la estructura Balanzas/
    archivos = glob.glob("Balanzas/**/Balanza.xlsx", recursive=True)
    if archivos:
        # Ordena por fecha de modificación para tomar el más reciente
        archivos.sort(key=os.path.getmtime, reverse=True)
        return archivos[0]
    return None


ruta_balanza = obtener_ruta_balanza()


@st.cache_data(ttl=300)
def obtener_lista_empresas(ruta):
    xls = pd.ExcelFile(ruta)
    return xls.sheet_names


@st.cache_data(ttl=300)
def cargar_hoja_empresa(ruta, nombre_hoja):
    return pd.read_excel(ruta, sheet_name=nombre_hoja)


if ruta_balanza and os.path.exists(ruta_balanza):
    st.caption(f"📁 Cargando balanzas desde: `{ruta_balanza}`")
    try:
        lista_empresas = obtener_lista_empresas(ruta_balanza)

        # Selector de empresa (cada pestaña del Excel de Balanzas es una empresa)
        empresa_seleccionada = st.selectbox(
            "Selecciona la Empresa:", lista_empresas
        )

        if empresa_seleccionada:
            with st.spinner(
                f"Cargando balanza de {empresa_seleccionada}..."
            ):
                df_empresa = cargar_hoja_empresa(
                    ruta_balanza, empresa_seleccionada
                )

            st.subheader(f"Estado Financiero - {empresa_seleccionada}")

            # Mostrar la tabla formateada con las cuentas y saldos
            st.dataframe(
                df_empresa,
                use_container_width=True,
                hide_index=True,
            )

            # Botón para descargar
            csv = df_empresa.to_csv(index=False).encode("utf-8")
            st.download_button(
                label=f"📥 Descargar Balanza ({empresa_seleccionada})",
                data=csv,
                file_name=f"Balanza_{empresa_seleccionada}.csv",
                mime="text/csv",
            )
    except Exception as e:
        st.error(f"Error al leer la balanza: {e}")
else:
    st.error(
        "No se encontró ningún archivo 'Balanza.xlsx' dentro de la carpeta 'Balanzas'."
    )