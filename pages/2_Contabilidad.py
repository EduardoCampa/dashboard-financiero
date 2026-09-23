import os
import glob
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(page_title="Módulo Contable - Estado de Resultados", layout="wide")

st.title("📊 Estado de Resultados")


# --- FUNCIONES DE CARGA Y PROCESAMIENTO ---

# Función para buscar de forma automática el archivo de Balanza más reciente
def obtener_ruta_balanza():
    # Busca cualquier archivo Balanza.xlsx dentro de la estructura Balanzas/
    archivos = glob.glob("Balanzas/**/Balanza.xlsx", recursive=True)
    if archivos:
        # Ordena por fecha de modificación para tomar el más reciente
        archivos.sort(key=os.path.getmtime, reverse=True)
        return archivos[0]
    return None


@st.cache_data(ttl=300)
def obtener_lista_empresas(ruta):
    """Obtiene los nombres de las pestañas (empresas) del Excel."""
    if not ruta: return []
    xls = pd.ExcelFile(ruta)
    return xls.sheet_names


def procesar_estado_resultados(df_balanza):
    """
    Toma una balanza de comprobación plana y la transforma en un 
    Estado de Resultados formateado como la imagen de ejemplo.
    
    Se asume que df_balanza tiene columnas como: 'CUENTA', 'NOMBRE', 'SALDO FINAL'
    y que se pueden identificar Ingresos, Costos y Gastos por el código de cuenta 
    o una columna de 'TIPO'.
    """
    
    # 1. Limpieza inicial: Asegurar nombres de columnas estándar y quitar filas vacías
    # Ajusta 'CUENTA', 'NOMBRE', 'SALDO FINAL' a los nombres reales en tu Excel
    try:
        df = df_balanza[['CUENTA', 'NOMBRE', 'SALDO FINAL']].copy()
        df.columns = ['Cuenta', 'Descripción', 'Monto']
        df = df.dropna(subset=['Cuenta', 'Monto']) # Quitar filas sin cuenta o monto
        df['Monto'] = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)
    except KeyError:
        st.error("El archivo de Balanza no tiene las columnas esperadas ('CUENTA', 'NOMBRE', 'SALDO FINAL').")
        return pd.DataFrame()

    # 2. Definición de Rangos de Cuentas (AJUSTA ESTO SEGÚN TU CATÁLOGO)
    # Ejemplo estándar: 4xxx Ingresos, 5xxx Costos, 6xxx Gastos
    def clasificar_cuenta(cuenta):
        cta_str = str(cuenta)
        if cta_str.startswith('4'): return 'Ingresos'
        if cta_str.startswith('5'): return 'Costos'
        if cta_str.startswith('6'): return 'Gastos'
        return 'Otro'

    df['Tipo'] = df['Cuenta'].apply(clasificar_cuenta)
    
    # Filtrar solo cuentas de resultados
    df_res = df[df['Tipo'].isin(['Ingresos', 'Costos', 'Gastos'])].copy()

    # 3. Agrupar y Calcular Subtotales
    ingresos_df = df_res[df_res['Tipo'] == 'Ingresos']
    costos_df = df_res[df_res['Tipo'] == 'Costos']
    gastos_df = df_res[df_res['Tipo'] == 'Gastos']
    
    total_ingresos = ingresos_df['Monto'].sum()
    total_costos = costos_df['Monto'].sum()
    utilidad_bruta = total_ingresos - total_costos
    
    total_gastos = gastos_df['Monto'].sum()
    utilidad_operacion = utilidad_bruta - total_gastos

    # 4. Construir la Tabla Formateada Final (como la imagen)
    final_data = []
    
    # Sección INGRESOS
    final_data.append({'Concepto': 'INGRESOS', 'Monto': None, 'Nivel': 0})
    for _, row in ingresos_df.iterrows():
        final_data.append({'Concepto': f"  {row['Cuenta']} - {row['Descripción']}", 'Monto': row['Monto'], 'Nivel': 1})
    final_data.append({'Concepto': 'TOTAL INGRESOS', 'Monto': total_ingresos, 'Nivel': 2})
    
    # Espacio
    final_data.append({'Concepto': '', 'Monto': None, 'Nivel': 0})
    
    # Sección COSTOS
    final_data.append({'Concepto': 'COSTOS', 'Monto': None, 'Nivel': 0})
    for _, row in costos_df.iterrows():
        final_data.append({'Concepto': f"  {row['Cuenta']} - {row['Descripción']}", 'Monto': row['Monto'], 'Nivel': 1})
    final_data.append({'Concepto': 'TOTAL COSTOS', 'Monto': total_costos, 'Nivel': 2})
    
    # Espacio y UTILIDAD BRUTA
    final_data.append({'Concepto': '', 'Monto': None, 'Nivel': 0})
    final_data.append({'Concepto': 'UTILIDAD BRUTA', 'Monto': utilidad_bruta, 'Nivel': 3})
    final_data.append({'Concepto': '', 'Monto': None, 'Nivel': 0})
    
    # Sección GASTOS
    final_data.append({'Concepto': 'GASTOS DE OPERACIÓN', 'Monto': None, 'Nivel': 0})
    for _, row in gastos_df.iterrows():
        final_data.append({'Concepto': f"  {row['Cuenta']} - {row['Descripción']}", 'Monto': row['Monto'], 'Nivel': 1})
    final_data.append({'Concepto': 'TOTAL GASTOS DE OPERACIÓN', 'Monto': total_gastos, 'Nivel': 2})

    # Espacio y UTILIDAD DE OPERACIÓN
    final_data.append({'Concepto': '', 'Monto': None, 'Nivel': 0})
    final_data.append({'Concepto': 'UTILIDAD DE OPERACIÓN', 'Monto': utilidad_operacion, 'Nivel': 3})

    df_final = pd.DataFrame(final_data)
    return df_final


# --- LÓGICA DE LA INTERFAZ ---

ruta_balanza = obtener_ruta_balanza()

if ruta_balanza and os.path.exists(ruta_balanza):
    st.caption(f"📁 Origen de datos: `{ruta_balanza}`")
    
    try:
        lista_empresas = obtener_lista_empresas(ruta_balanza)
        empresa_seleccionada = st.selectbox("Selecciona la Empresa:", lista_empresas)

        if empresa_seleccionada:
            # Cargar la balanza plana de la empresa (sin caché aquí para procesar)
            df_balanza_plana = pd.read_excel(ruta_balanza, sheet_name=empresa_seleccionada)
            
            # Transformar la balanza plana en el Estado de Resultados formateado
            with st.spinner(f"Generando Estado de Resultados para {empresa_seleccionada}..."):
                df_estado_resultados = procesar_estado_resultados(df_balanza_plana)

            if not df_estado_resultados.empty:
                st.subheader(f"Estado de Resultados - {empresa_seleccionada}")
                
                # --- FORMATEO VISUAL ESTILO EXCEL (Estilo de la imagen) ---
                # Definimos estilos para resaltar subtotales y totales
                def styler_estado_resultados(df):
                    # Crear copia para no modificar datos
                    sf = df.copy()
                    
                    # Formato de moneda
                    numeric_cols = ['Monto']
                    sf[numeric_cols] = sf[numeric_cols].applymap(lambda x: f"${x:,.2f}" if pd.notnull(x) else "")
                    
                    # Quitar columna Nivel de la visualización
                    sf = sf.drop(columns=['Nivel'])
                    return sf

                # Aplicar estilos y mostrar
                # Nota: st.dataframe no soporta todo el estilizado complejo de negritas por fila fácilmente,
                # pero st.table o formatear el texto antes ayuda. Usaremos formateo de texto para simplicidad.
                
                # Formatear montos como moneda antes de mostrar
                df_mostrar = df_estado_resultados.copy()
                df_mostrar['Monto'] = df_mostrar['Monto'].apply(lambda x: f"${x:,.2f}" if pd.notnull(x) else "")
                
                # Ocultar columna de nivel y mostrar tabla
                st.dataframe(
                    df_mostrar.drop(columns=['Nivel']),
                    use_container_width=True,
                    hide_index=True
                )

                # Botón de descarga del Estado de Resultados generado
                csv = df_estado_resultados.drop(columns=['Nivel']).to_csv(index=False).encode("utf-8")
                st.download_button(
                    label=f"📥 Descargar Estado de Resultados ({empresa_seleccionada})",
                    data=csv,
                    file_name=f"Estado_Resultados_{empresa_seleccionada}.csv",
                    mime="text/csv",
                )
            else:
                st.warning("No se pudieron generar datos para el Estado de Resultados. Revisa la estructura de la balanza.")

    except Exception as e:
        st.error(f"Error al procesar el Estado de Resultados: {e}")
else:
    st.error("No se encontró el archivo de balanzas necesario en la ruta esperada.")