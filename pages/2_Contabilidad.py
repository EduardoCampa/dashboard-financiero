import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Módulo Contable", layout="wide")

st.title("📊 Estados Financieros")


# Retornamos directamente la lista de nombres de pestañas (que sí se puede guardar en caché)
@st.cache_data
def obtener_pestanas_empresas(ruta_archivo):
    xls = pd.ExcelFile(ruta_archivo)
    return xls.sheet_names


archivo_balanzas = "Consolidado_Master.xlsx"

if os.path.exists(archivo_balanzas):
    # Obtenemos la lista de nombres de las empresas
    lista_empresas = obtener_pestanas_empresas(archivo_balanzas)

    # Selector de empresa
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