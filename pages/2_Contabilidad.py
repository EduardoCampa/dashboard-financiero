import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Módulo Contable", layout="wide")

st.title("📊 Estados Financieros")


# Función para cargar la lista de empresas (hojas de Excel)
@st.cache_data
def cargar_datos_balanzas(ruta_archivo):
    # Lee todas las pestañas/hojas del archivo Excel
    xls = pd.ExcelFile(ruta_archivo)
    return xls


archivo_balanzas = "Consolidado_Master.xlsx"

if os.path.exists(archivo_balanzas):
    excel_balanzas = cargar_datos_balanzas(archivo_balanzas)
    lista_empresas = excel_balanzas.sheet_names

    # Selector de empresa (cada empresa corresponde a una pestaña)
    empresa_seleccionada = st.selectbox(
        "Selecciona la Empresa:", lista_empresas
    )

    if empresa_seleccionada:
        df_empresa = pd.read_excel(
            archivo_balanzas, sheet_name=empresa_seleccionada
        )

        st.subheader(f"Estado Financiero - {empresa_seleccionada}")

        # Mostrar tabla formateada
        st.dataframe(
            df_empresa,
            use_container_width=True,
            hide_index=True,
        )

        # Botón para descargar el reporte de la empresa actual
        csv = df_empresa.to_csv(index=False).encode("utf-8")
        st.download_button(
            label=f"📥 Descargar Estado Financiero ({empresa_seleccionada})",
            data=csv,
            file_name=f"Estado_Financiero_{empresa_seleccionada}.csv",
            mime="text/csv",
        )
else:
    st.warning(
        f"No se encontró el archivo '{archivo_balanzas}'. Asegúrate de ejecutar primero la consolidación."
    )