import glob
import os
import re
import openpyxl
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


# --- CARGAR PLANTILLA EXACTA DE EXCEL ---
@st.cache_data(ttl=3600)
def cargar_plantilla_formato():
    ruta_formato = "FORMATO EDO RESULTADOS.xlsx"
    if not os.path.exists(ruta_formato):
        return []

    wb = openpyxl.load_workbook(ruta_formato, data_only=True)
    sheet = wb.active

    filas_plantilla = []
    for i in range(4, sheet.max_row + 1):
        cta = sheet.cell(row=i, column=1).value
        concepto = sheet.cell(row=i, column=2).value
        if cta or concepto:
            filas_plantilla.append({
                'row_idx': i,
                'cuenta_patron': str(cta).strip() if cta else None,
                'concepto': str(concepto).strip() if concepto else '',
            })
    return filas_plantilla


def coindice_patron(cuenta_balanza, patron):
    if not patron or patron == 'CUENTA':
        return False
    # Transforma 410-?????-001-0000 a Regex
    regex_patron = f"^{patron.replace('?', '.').replace('-', r'\\-?')}$"
    return bool(re.match(regex_patron, str(cuenta_balanza).strip()))


# --- GENERAR ESTADO DE RESULTADOS BASADO EN EL FORMATO ---
def generar_estado_resultados_formato(df_balanza, plantilla):
    if not plantilla or df_balanza.empty:
        return pd.DataFrame()

    # Mapeo de columnas por índice en la balanza
    # Col A (0): Cuenta, Col B (1): Nombre
    # Col E (4): Cargos Mes, Col F (5): Abonos Mes
    # Col G (6): Cargos Acum, Col H (7): Abonos Acum

    cuentas_balanza = []
    for idx, row in df_balanza.iterrows():
        cta = str(row.iloc[0]).strip()
        cargos_mes = pd.to_numeric(row.iloc[4], errors='coerce') or 0
        abonos_mes = pd.to_numeric(row.iloc[5], errors='coerce') or 0
        cargos_acum = pd.to_numeric(row.iloc[6], errors='coerce') or 0
        abonos_acum = pd.to_numeric(row.iloc[7], errors='coerce') or 0

        # Determinación de fórmula por inicio de cuenta
        if cta.startswith(('4', '720', '730')):
            val_mes = abonos_mes - cargos_mes  # Ingresos Del Mes
            val_acum = abonos_acum - cargos_acum  # Ingresos Acumulado
        else:
            val_mes = cargos_mes - abonos_mes  # Costos/Gastos Del Mes
            val_acum = cargos_acum - abonos_acum  # Costos/Gastos Acumulado

        cuentas_balanza.append({
            'cuenta': cta,
            'val_mes': val_mes,
            'val_acum': val_acum,
        })

    df_cuentas = pd.DataFrame(cuentas_balanza)

    # Construir el reporte respetando cada renglón de la plantilla
    reporte = []
    acumuladores_mes = {}
    acumuladores_acum = {}

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        concepto = row['concepto']

        # Si es encabezado principal o título
        if patron == 'CUENTA':
            reporte.append({
                'CUENTA': 'CUENTA',
                'CONCEPTO': 'CONCEPTO',
                'DEL MES': 'DEL MES',
                'ACUMULADO': 'ACUMULADO',
                'es_total': True,
            })
            continue

        # Renglón de cuenta individual
        if patron and '?' in patron:
            coincidencias = df_cuentas[
                df_cuentas['cuenta'].apply(
                    lambda c: coindice_patron(c, patron)
                )
            ]
            m_mes = coincidencias['val_mes'].sum()
            m_acum = coincidencias['val_acum'].sum()

            acumuladores_mes[r_idx] = m_mes
            acumuladores_acum[r_idx] = m_acum

            reporte.append({
                'CUENTA': patron,
                'CONCEPTO': concepto,
                'DEL MES': m_mes,
                'ACUMULADO': m_acum,
                'es_total': False,
            })
        else:
            # Es una fila de Subtotal / Total / Encabezado de Sección
            reporte.append({
                'CUENTA': patron if patron else '',
                'CONCEPTO': concepto,
                'DEL MES': None,
                'ACUMULADO': None,
                'es_total': True,
            })

    return pd.DataFrame(reporte)


# --- INTERFAZ PRINCIPAL ---
ruta_balanza = obtener_ruta_balanza()
plantilla_formato = cargar_plantilla_formato()

if ruta_balanza and os.path.exists(ruta_balanza):
    st.caption(f"📁 Archivo de Origen: `{ruta_balanza}`")

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

            if not plantilla_formato:
                st.error(
                    "No se encontró el archivo 'FORMATO EDO RESULTADOS.xlsx' en la raíz del proyecto."
                )
            else:
                df_er = generar_estado_resultados_formato(
                    df_balanza, plantilla_formato
                )

                if not df_er.empty:
                    # Formatear montos a moneda ($#,###.##)
                    df_display = df_er.copy()

                    def fmt(val):
                        if pd.isnull(val) or val == '' or isinstance(val, str):
                            return val
                        return f"${val:,.2f}"

                    df_display['DEL MES'] = df_display['DEL MES'].apply(fmt)
                    df_display['ACUMULADO'] = df_display['ACUMULADO'].apply(
                        fmt
                    )

                    st.dataframe(
                        df_display.drop(columns=['es_total']),
                        use_container_width=True,
                        hide_index=True,
                    )

                    # Botón de Descarga
                    csv_er = df_er.drop(columns=['es_total']).to_csv(
                        index=False
                    ).encode('utf-8')
                    st.download_button(
                        label=f"📥 Descargar Estado de Resultados ({empresa_seleccionada})",
                        data=csv_er,
                        file_name=f"Estado_Resultados_{empresa_seleccionada}.csv",
                        mime="text/csv",
                    )
else:
    st.error(
        "No se encontró el archivo de balanzas en la carpeta 'Balanzas/'."
    )