import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Contabilidad - Grupo SERVYRE", layout="wide")

st.sidebar.title("📚 Módulo de Contabilidad")
st.sidebar.markdown("---")
sub_cont = st.sidebar.radio("Seleccione Submódulo:", ["Balanza de Comprobación", "Pólizas", "Conciliación Bancaria"], key="sub_contabilidad")

if sub_cont == "Balanza de Comprobación":
    st.title("📚 Módulo de Contabilidad")
    st.markdown("### 📊 Consulta Automática de Balanzas de Comprobación")
    
    col_btn1, col_btn2 = st.columns([6, 1])
    with col_btn2:
        if st.button("🔄 Refrescar"):
            st.cache_data.clear()
            st.rerun()

    carpeta_principal_balanzas = "Balanzas"

    if os.path.exists(carpeta_principal_balanzas):
        anios_disponibles = sorted([d for d in os.listdir(carpeta_principal_balanzas) if os.path.isdir(os.path.join(carpeta_principal_balanzas, d))])
    else:
        anios_disponibles = []

    if anios_disponibles:
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            anio_con = st.selectbox("Seleccione el Año:", anios_disponibles, key="balanza_anio_con")
        
        ruta_anio = os.path.join(carpeta_principal_balanzas, str(anio_con))
        meses_disponibles = sorted([d for d in os.listdir(ruta_anio) if os.path.isdir(os.path.join(ruta_anio, d))])
        
        meses_dict_nombres = {
            "01": "01 - Enero", "02": "02 - Febrero", "03": "03 - Marzo", "04": "04 - Abril",
            "05": "05 - Mayo", "06": "06 - Junio", "07": "07 - Julio", "08": "08 - Agosto",
            "09": "09 - Septiembre", "10": "10 - Octubre", "11": "11 - Noviembre", "12": "12 - Diciembre"
        }
        meses_opciones = [m for m in meses_disponibles if m in meses_dict_nombres]
        
        with col_c2:
            mes_con = st.selectbox("Seleccione el Mes:", meses_opciones, format_func=lambda x: meses_dict_nombres.get(x, x), key="balanza_mes_con")
        
        if mes_con:
            ruta_a_consultar = os.path.join(carpeta_principal_balanzas, str(anio_con), mes_con, "Balanza.xlsx")
            st.info(f"📂 Ruta activa de lectura automática: `Balanzas/{anio_con}/{mes_con}/Balanza.xlsx`")

            if os.path.exists(ruta_a_consultar):
                try:
                    xls_temp = pd.ExcelFile(ruta_a_consultar)
                    hojas_guardadas = xls_temp.sheet_names
                    
                    empresa_sel_con = st.selectbox("Seleccione la empresa a visualizar:", hojas_guardadas, key="visor_empresa_guardada")
                    if empresa_sel_con:
                        df_raw = pd.read_excel(ruta_a_consultar, sheet_name=empresa_sel_con, header=None)
                        fila_header = 0
                        for idx, row in df_raw.iterrows():
                            if "Cuenta" in str(row.values):
                                fila_header = idx
                                break
                        
                        df_vista = pd.read_excel(ruta_a_consultar, sheet_name=empresa_sel_con, header=fila_header)
                        df_vista = df_vista.dropna(how='all')
                        
                        st.markdown(f"#### Empresa: **{empresa_sel_con}** (Periodo: {meses_dict_nombres.get(mes_con, mes_con)} {anio_con})")
                        st.dataframe(df_vista, use_container_width=True)
                except Exception as e:
                    st.error(f"Error al leer el archivo de balanza: {e}")
            else:
                st.warning(f"⚠️ No se encontró el archivo `Balanza.xlsx` en la ruta especificada.")
    else:
        st.warning("⚠️ No se encontró la carpeta `Balanzas/` en el proyecto.")
else:
    st.title("📚 Módulo de Contabilidad")
    st.info(f"El submódulo de **{sub_cont}** se encuentra en desarrollo.")