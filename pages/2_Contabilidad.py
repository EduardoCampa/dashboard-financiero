import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo Contable", layout="wide", initial_sidebar_state="expanded"
)

st.title("📊 Módulo Contable")


# --- OBTENER ARCHIVO DE BALANZA MÁS RECIENTE ---
def obtener_ruta_balanza():
    archivos = glob.glob("Balanzas/**/Balanza.xlsx", recursive=True)
    if archivos:
        archivos.sort(key=os.path.getmtime, reverse=True)
        return archivos[0]
    return None


@st.cache_data(ttl=300)
def obtener_lista_empresas(ruta):
    xls = pd.ExcelFile(ruta)
    return xls.sheet_names


@st.cache_data(ttl=300)
def cargar_hoja_balanza(ruta, nombre_hoja):
    return pd.read_excel(ruta, sheet_name=nombre_hoja)


# --- LÓGICA DEL ESTADO DE RESULTADOS ---
def generar_estado_resultados(df):
    """Procesa la balanza por índices de columna (A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7)

    Aplica la fórmula contable indicada para Del Mes y Acumulado.
    """
    try:
        # Copia de trabajo
        data = df.copy()

        # Aseguramos que trabajamos con filas numéricas a partir de donde empiezan las cuentas
        # Columna A (0): Cuenta, Columna B (1): Nombre
        # Columna E (4): Cargos Mes, Columna F (5): Abonos Mes
        # Columna G (6): Cargos Acum, Columna H (7): Abonos Acum

        filas_procesadas = []

        for idx, row in data.iterrows():
            cta_str = str(row.iloc[0]).strip()
            nombre_str = str(row.iloc[1]).strip()

            # Filtrar solo cuentas que inicien con 4, 5 o 6
            if cta_str.startswith(('4', '5', '6')):
                # Convertir montos a numéricos (col E, F, G, H -> índices 4, 5, 6, 7)
                cargos_mes = pd.to_numeric(row.iloc[4], errors='coerce') or 0
                abonos_mes = pd.to_numeric(row.iloc[5], errors='coerce') or 0
                cargos_acum = pd.to_numeric(row.iloc[6], errors='coerce') or 0
                abonos_acum = pd.to_numeric(row.iloc[7], errors='coerce') or 0

                # Clasificación y fórmulas exactas requeridas:
                if cta_str.startswith('4'):
                    tipo = 'Ingresos'
                    monto_mes = abonos_mes - cargos_mes  # Abonos (F) - Cargos (E)
                    monto_acum = (
                        abonos_acum - cargos_acum
                    )  # Abonos (H) - Cargos (G)
                elif cta_str.startswith('5'):
                    tipo = 'Costos'
                    monto_mes = cargos_mes - abonos_mes  # Cargos (E) - Abonos (F)
                    monto_acum = (
                        cargos_acum - abonos_acum
                    )  # Cargos (G) - Abonos (H)
                else:  # Inicia con 6
                    tipo = 'Gastos'
                    monto_mes = cargos_mes - abonos_mes  # Cargos (E) - Abonos (F)
                    monto_acum = (
                        cargos_acum - abonos_acum
                    )  # Cargos (G) - Abonos (H)

                filas_procesadas.append({
                    'Cuenta': cta_str,
                    'Descripción': nombre_str,
                    'Tipo': tipo,
                    'Del Mes': monto_mes,
                    'Acumulado': monto_acum,
                })

        df_res = pd.DataFrame(filas_procesadas)

        if df_res.empty:
            return pd.DataFrame()

        # Agrupar por rubros
        ingresos_df = df_res[df_res['Tipo'] == 'Ingresos']
        costos_df = df_res[df_res['Tipo'] == 'Costos']
        gastos_df = df_res[df_res['Tipo'] == 'Gastos']

        # Totales
        tot_ing_mes, tot_ing_acum = (
            ingresos_df['Del Mes'].sum(),
            ingresos_df['Acumulado'].sum(),
        )
        tot_cos_mes, tot_cos_acum = (
            costos_df['Del Mes'].sum(),
            costos_df['Acumulado'].sum(),
        )
        tot_gas_mes, tot_gas_acum = (
            gastos_df['Del Mes'].sum(),
            gastos_df['Acumulado'].sum(),
        )

        util_bruta_mes = tot_ing_mes - tot_cos_mes
        util_bruta_acum = tot_ing_acum - tot_cos_acum

        util_op_mes = util_bruta_mes - tot_gas_mes
        util_op_acum = util_bruta_acum - tot_gas_acum

        # Construcción final del reporte
        tabla_final = []

        # --- INGRESOS ---
        tabla_final.append({
            'Concepto': 'INGRESOS',
            'Del Mes': None,
            'Acumulado': None,
        })
        for _, r in ingresos_df.iterrows():
            tabla_final.append({
                'Concepto': f"  {r['Cuenta']} - {r['Descripción']}",
                'Del Mes': r['Del Mes'],
                'Acumulado': r['Acumulado'],
            })
        tabla_final.append({
            'Concepto': 'TOTAL INGRESOS',
            'Del Mes': tot_ing_mes,
            'Acumulado': tot_ing_acum,
        })
        tabla_final.append(
            {'Concepto': '', 'Del Mes': None, 'Acumulado': None}
        )

        # --- COSTOS ---
        tabla_final.append({
            'Concepto': 'COSTOS',
            'Del Mes': None,
            'Acumulado': None,
        })
        for _, r in costos_df.iterrows():
            tabla_final.append({
                'Concepto': f"  {r['Cuenta']} - {r['Descripción']}",
                'Del Mes': r['Del Mes'],
                'Acumulado': r['Acumulado'],
            })
        tabla_final.append({
            'Concepto': 'TOTAL COSTOS',
            'Del Mes': tot_cos_mes,
            'Acumulado': tot_cos_acum,
        })
        tabla_final.append(
            {'Concepto': '', 'Del Mes': None, 'Acumulado': None}
        )

        # --- UTILIDAD BRUTA ---
        tabla_final.append({
            'Concepto': 'UTILIDAD BRUTA',
            'Del Mes': util_bruta_mes,
            'Acumulado': util_bruta_acum,
        })
        tabla_final.append(
            {'Concepto': '', 'Del Mes': None, 'Acumulado': None}
        )

        # --- GASTOS ---
        tabla_final.append({
            'Concepto': 'GASTOS DE OPERACIÓN',
            'Del Mes': None,
            'Acumulado': None,
        })
        for _, r in gastos_df.iterrows():
            tabla_final.append({
                'Concepto': f"  {r['Cuenta']} - {r['Descripción']}",
                'Del Mes': r['Del Mes'],
                'Acumulado': r['Acumulado'],
            })
        tabla_final.append({
            'Concepto': 'TOTAL GASTOS DE OPERACIÓN',
            'Del Mes': tot_gas_mes,
            'Acumulado': tot_gas_acum,
        })
        tabla_final.append(
            {'Concepto': '', 'Del Mes': None, 'Acumulado': None}
        )

        # --- UTILIDAD DE OPERACIÓN ---
        tabla_final.append({
            'Concepto': 'UTILIDAD DE OPERACIÓN',
            'Del Mes': util_op_mes,
            'Acumulado': util_op_acum,
        })

        return pd.DataFrame(tabla_final)

    except Exception as e:
        st.error(f"Error procesando la balanza: {e}")
        return pd.DataFrame()


# --- INTERFAZ PRINCIPAL ---

ruta_balanza = obtener_ruta_balanza()

if ruta_balanza and os.path.exists(ruta_balanza):
    st.caption(f"📁 Archivo de origen: `{ruta_balanza}`")

    lista_empresas = obtener_lista_empresas(ruta_balanza)
    empresa_seleccionada = st.selectbox("Selecciona la Empresa:", lista_empresas)

    if empresa_seleccionada:
        df_balanza = cargar_hoja_balanza(ruta_balanza, empresa_seleccionada)

        # MÓDULO CON PESTAÑAS (TABS)
        tab_balanzas, tab_er = st.tabs(
            ["📑 Balanzas de Comprobación", "📈 Estado de Resultados"]
        )

        # PESTAÑA 1: BALANZAS
        with tab_balanzas:
            st.subheader(f"Balanza de Comprobación - {empresa_seleccionada}")
            st.dataframe(
                df_balanza,
                use_container_width=True,
                hide_index=True,
            )

        # PESTAÑA 2: ESTADO DE RESULTADOS
        with tab_er:
            st.subheader(f"Estado de Resultados - {empresa_seleccionada}")

            df_er = generar_estado_resultados(df_balanza)

            if not df_er.empty:
                # Formatear números a moneda ($#,###.##) para presentación
                df_er_display = df_er.copy()
                df_er_display['Del Mes'] = df_er_display['Del Mes'].apply(
                    lambda x: f"${x:,.2f}" if pd.notnull(x) else ""
                )
                df_er_display['Acumulado'] = df_er_display['Acumulado'].apply(
                    lambda x: f"${x:,.2f}" if pd.notnull(x) else ""
                )

                st.dataframe(
                    df_er_display,
                    use_container_width=True,
                    hide_index=True,
                )

                # Descarga CSV
                csv_er = df_er.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Descargar Estado de Resultados ({empresa_seleccionada})",
                    data=csv_er,
                    file_name=f"Estado_Resultados_{empresa_seleccionada}.csv",
                    mime="text/csv",
                )
            else:
                st.warning(
                    "No se encontraron cuentas de resultados (4xxx, 5xxx, 6xxx) en esta pestaña."
                )

else:
    st.error(
        "No se encontró el archivo de balanzas en la carpeta 'Balanzas/'."
    )