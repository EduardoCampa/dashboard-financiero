import glob
import os
import io
import re
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Módulo SAT y Conciliación Master vs XML",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- MENÚ LATERAL DE NAVEGACIÓN ---
st.sidebar.markdown("### 📑 Módulo SAT")
submodulo_sat = st.sidebar.radio(
    "Seleccione Submódulo:",
    ["📊 Amarre Ingresos", "👥 Amarre Nóminas", "🏛️ Impuestos"]
)

nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# --- FUNCIÓN AUXILIAR GENERAL MEJORADA PARA OBTENER SALDOS DE BALANZA ---
def obtener_saldos_balanza(ruta_balanza):
    if not ruta_balanza or not os.path.exists(ruta_balanza):
        return pd.DataFrame()

    detalles = []
    try:
        xls = pd.ExcelFile(ruta_balanza)
        for hoja in xls.sheet_names:
            if hoja.lower() in ['hoja1', 'hoja 1', 'resumen', 'sheet1']:
                continue
            
            nombre_pestana = hoja.strip().upper()
            df_hoja = pd.read_excel(ruta_balanza, sheet_name=hoja)
            if df_hoja.empty or df_hoja.shape[1] < 4:
                continue

            cols_upper = [str(c).strip().upper() for c in df_hoja.columns]
            
            # Buscar dinámicamente índices
            idx_ac_f = next((i for i, c in enumerate(cols_upper) if 'ACREEDOR F' in c or 'ACREEDOR FINAL' in c or 'SALDO F' in c), None)
            if idx_ac_f is None:
                idx_ac_f = 7 if df_hoja.shape[1] > 7 else df_hoja.shape[1] - 1
                
            idx_de_f = next((i for i, c in enumerate(cols_upper) if 'DEUDOR F' in c or 'DEUDOR FINAL' in c), None)
            if idx_de_f is None:
                idx_de_f = 6 if df_hoja.shape[1] > 6 else df_hoja.shape[1] - 2

            idx_col_e = 4 if df_hoja.shape[1] > 4 else 0 # Column E (Cargos)
            idx_col_f = 5 if df_hoja.shape[1] > 5 else 0 # Column F (Abonos)

            for _, row in df_hoja.iterrows():
                cta = str(row.iloc[0]).strip()
                if not cta or cta.lower() in ('nan', 'cuenta', 'none', 'total', 'totales'):
                    continue
                
                cta_limpia = cta.replace(' ', '')
                nom = str(row.iloc[1]).strip() if df_hoja.shape[1] > 1 else ""
                
                monto_de_f = float(pd.to_numeric(str(row.iloc[idx_de_f]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)
                monto_ac_f = float(pd.to_numeric(str(row.iloc[idx_ac_f]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)
                
                monto_col_e = float(pd.to_numeric(str(row.iloc[idx_col_e]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)
                monto_col_f = float(pd.to_numeric(str(row.iloc[idx_col_f]).replace('$', '').replace(',', ''), errors='coerce') or 0.0)

                detalles.append({
                    'Empresa': nombre_pestana,
                    'Cuenta': cta_limpia,
                    'Nombre Cuenta': nom,
                    'Saldo Deudor Final': monto_de_f,
                    'Saldo Acreedor Final': monto_ac_f,
                    'Columna E (Cargos)': monto_col_e,
                    'Columna F (Abonos)': monto_col_f
                })
    except Exception:
        pass

    return pd.DataFrame(detalles)


# --- FUNCIÓN PARA GENERAR EXCEL EJECUTIVO DE IVA ---
def generar_excel_iva(df_transpuesto, periodo_str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen IVA"
    ws.views.sheetView[0].showGridLines = True

    font_titulo = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    fill_titulo = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    
    font_normal = Font(name="Calibri", size=10, color="000000")
    font_total_col = Font(name="Calibri", size=10, bold=True, color="000000")
    fill_total_col = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    borde_delgado = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    total_cols = len(df_transpuesto.columns)
    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    cell_t = ws.cell(row=1, column=1, value=f"REPORTE EJECUTIVO DE DETERMINACIÓN DE IVA — ({periodo_str.upper()})")
    cell_t.font = font_titulo
    cell_t.fill = fill_titulo
    cell_t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    for col_num, col_name in enumerate(df_transpuesto.columns, 1):
        c = ws.cell(row=3, column=col_num, value=col_name)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borde_delgado
    ws.row_dimensions[3].height = 25

    for idx_f, (_, r) in enumerate(df_transpuesto.iterrows(), start=4):
        for col_num, col_name in enumerate(df_transpuesto.columns, 1):
            val = r[col_name]
            c = ws.cell(row=idx_f, column=col_num)

            if col_num == 1:
                c.value = str(val)
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.font = Font(name="Calibri", size=10, bold=True)
            else:
                c.value = float(val) if pd.notnull(val) else 0.0
                c.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.font = font_total_col if col_name == 'TOTAL CONSOLIDADO' else font_normal

            c.border = borde_delgado
            if col_name == 'TOTAL CONSOLIDADO':
                c.fill = fill_total_col
            elif idx_f % 2 == 1:
                c.fill = fill_zebra

        ws.row_dimensions[idx_f].height = 20

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 16)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ==========================================
# 1. SUBMÓDULO: AMARRE INGRESOS
# ==========================================
if submodulo_sat == "📊 Amarre Ingresos":
    st.title("📑 Módulo SAT - Comparativos, Conciliación y Reporte Ejecutivo")
    st.info("Seleccione las opciones e inicie el proceso para consultar los Ingresos.")


# ==========================================
# 2. SUBMÓDULO: AMARRE NÓMINAS
# ==========================================
elif submodulo_sat == "👥 Amarre Nóminas":
    st.title("👥 Módulo SAT - Amarre de Nómina vs Contabilidad")
    st.info("Seleccione las opciones e inicie el proceso para consultar la nómina.")


# ==========================================
# 3. SUBMÓDULO: IMPUESTOS (RESUMEN MULTI-EMPRESA)
# ==========================================
elif submodulo_sat == "🏛️ Impuestos":

    st.title("🏛️ Módulo de Cálculos de Impuestos — Resumen Ejecutivo")

    col_a_imp, col_m_imp = st.columns([1, 1])

    with col_a_imp:
        anio_imp = st.selectbox("Año de Consulta:", [2026, 2025, 2024], index=0, key="imp_anio")
    with col_m_imp:
        mes_imp = st.selectbox("Mes de Consulta:", list(range(1, 13)), index=8, format_func=lambda x: nombres_meses[x-1], key="imp_mes")

    st.markdown("---")
    ruta_balanzas_imp = st.text_input("Carpeta Raíz de Balanzas:", value="Balanzas", key="imp_bal_dir")

    mes_str = f"{int(mes_imp):02d}"
    posibles = (
        glob.glob(os.path.join(ruta_balanzas_imp, str(anio_imp), mes_str, "*.xlsx"), recursive=True) +
        glob.glob(os.path.join(ruta_balanzas_imp, "**", mes_str, "*.xlsx"), recursive=True)
    )
    posibles_gen = [f for f in posibles if "rh" not in os.path.basename(f).lower()]
    ruta_balanza_sel = posibles_gen[0] if posibles_gen else (posibles[0] if posibles else None)

    tab_isr, tab_iva, tab_retenciones = st.tabs([
        "📈 ISR (Pagos Provisionales)", 
        "💵 IVA (Cobrado vs Pagado)", 
        "📋 RETENCIONES"
    ])

    # ---------------------------------------------------------
    # 1. PESTAÑA ISR (TABLA COMPARATIVA MULTI-EMPRESA)
    # ---------------------------------------------------------
    with tab_isr:
        st.markdown(f"### 📈 Resumen General de ISR — {nombres_meses[mes_imp-1]} {anio_imp}")

        if ruta_balanza_sel and os.path.exists(ruta_balanza_sel):
            df_saldos = obtener_saldos_balanza(ruta_balanza_sel)

            if not df_saldos.empty:
                empresas_unicas = sorted(df_saldos['Empresa'].unique())
                
                datos_isr = []
                for emp in empresas_unicas:
                    df_emp = df_saldos[df_saldos['Empresa'] == emp]

                    # Extracción flexible: intenta Saldo Acreedor Final, si es 0 prueba con Columna F (Abonos)
                    v_410_f = float(df_emp[df_emp['Cuenta'].str.startswith('410-00000-000-0000')]['Saldo Acreedor Final'].sum())
                    v_410 = v_410_f if v_410_f != 0 else float(df_emp[df_emp['Cuenta'].str.startswith('410-00000-000-0000')]['Columna F (Abonos)'].sum())

                    v_411_f = float(df_emp[df_emp['Cuenta'].str.startswith('411-00000-000-0000')]['Saldo Acreedor Final'].sum())
                    v_411 = v_411_f if v_411_f != 0 else float(df_emp[df_emp['Cuenta'].str.startswith('411-00000-000-0000')]['Columna F (Abonos)'].sum())

                    df_423 = df_emp[df_emp['Cuenta'].str.startswith('423-00000-000-0000')]
                    v_423_ac = float(df_423['Saldo Acreedor Final'].sum())
                    v_423_de = float(df_423['Saldo Deudor Final'].sum())
                    v_423 = v_423_ac if v_423_ac > 0 else v_423_de

                    ingresos_isr = v_410 + v_411 - v_423

                    datos_isr.append({
                        'Empresa': emp,
                        'Cuenta 410 (Ventas/Ingresos Obra)': v_410,
                        'Cuenta 411 (Otros Ingresos)': v_411,
                        'Cuenta 423 (-) Dev. y Desc.': v_423,
                        'Ingresos Nominales ISR': ingresos_isr,
                        'Coeficiente de Utilidad': 0.0500
                    })

                df_base_isr = pd.DataFrame(datos_isr)

                st.markdown("#### ✏️ Capture o Modifique el Coeficiente de Utilidad por Empresa:")
                
                df_edited = st.data_editor(
                    df_base_isr[['Empresa', 'Ingresos Nominales ISR', 'Coeficiente de Utilidad']],
                    column_config={
                        "Empresa": st.column_config.TextColumn(disabled=True),
                        "Ingresos Nominales ISR": st.column_config.NumberColumn(format="$%,.2f", disabled=True),
                        "Coeficiente de Utilidad": st.column_config.NumberColumn(format="%.4f", step=0.0001, min_value=0.0, max_value=1.0)
                    },
                    hide_index=True,
                    use_container_width=True
                )

                df_merged = pd.merge(df_base_isr.drop(columns=['Coeficiente de Utilidad']), df_edited[['Empresa', 'Coeficiente de Utilidad']], on='Empresa')
                
                df_merged['Utilidad Fiscal Estimada'] = df_merged['Ingresos Nominales ISR'] * df_merged['Coeficiente de Utilidad']
                df_merged['Tasa ISR'] = 0.30
                df_merged['Pago Provisional ISR (30%)'] = df_merged['Utilidad Fiscal Estimada'] * 0.30

                conceptos_isr = [
                    'Cuenta 410 (Ventas/Ingresos Obra)',
                    'Cuenta 411 (Otros Ingresos)',
                    'Cuenta 423 (-) Dev. y Desc.',
                    'Ingresos Nominales ISR',
                    'Coeficiente de Utilidad',
                    'Utilidad Fiscal Estimada',
                    'Pago Provisional ISR (30%)'
                ]

                tabla_resumen_isr = pd.DataFrame({'Concepto': conceptos_isr})

                for _, row in df_merged.iterrows():
                    emp_col = row['Empresa']
                    tabla_resumen_isr[emp_col] = [
                        row['Cuenta 410 (Ventas/Ingresos Obra)'],
                        row['Cuenta 411 (Otros Ingresos)'],
                        row['Cuenta 423 (-) Dev. y Desc.'],
                        row['Ingresos Nominales ISR'],
                        row['Coeficiente de Utilidad'],
                        row['Utilidad Fiscal Estimada'],
                        row['Pago Provisional ISR (30%)']
                    ]

                totales_row = []
                for concepto in conceptos_isr:
                    if concepto == 'Coeficiente de Utilidad':
                        totales_row.append(None)
                    else:
                        totales_row.append(df_merged[concepto].sum() if concepto in df_merged.columns else 0.0)

                tabla_resumen_isr['TOTAL CONSOLIDADO'] = totales_row

                st.markdown("---")
                st.markdown(f"#### 📊 Tabla de Resumen Comparativo ISR — Grupo")

                format_dict_isr = {col: "${:,.2f}" for col in tabla_resumen_isr.columns if col != 'Concepto'}
                
                st.dataframe(
                    tabla_resumen_isr.style.format(format_dict_isr, na_rep="-"),
                    use_container_width=True,
                    hide_index=True
                )

            else:
                st.warning("No se pudieron extraer saldos de la balanza seleccionada.")
        else:
            st.error(f"No se encontró archivo de balanza en el periodo {nombres_meses[mes_imp-1]} {anio_imp}.")


    # ---------------------------------------------------------
    # 2. PESTAÑA IVA
    # ---------------------------------------------------------
    with tab_iva:
        st.markdown(f"### 💵 Resumen General de IVA — {nombres_meses[mes_imp-1]} {anio_imp}")

        if ruta_balanza_sel and os.path.exists(ruta_balanza_sel):
            df_saldos_iva = obtener_saldos_balanza(ruta_balanza_sel)

            if not df_saldos_iva.empty:
                empresas_unicas_iva = sorted(df_saldos_iva['Empresa'].unique())

                datos_iva = []
                for emp in empresas_unicas_iva:
                    df_emp_iva = df_saldos_iva[df_saldos_iva['Empresa'] == emp]

                    # 201-00010-001-0000 IVA Cobrado -> Columna F (Abonos)
                    v_cobrado = float(df_emp_iva[df_emp_iva['Cuenta'].str.startswith('201-00010-001-0000')]['Columna F (Abonos)'].sum())

                    # 101-00011-001-0000 + 101-00011-002-0000 IVA Pagado -> Columna E (Cargos)
                    v_pag1 = float(df_emp_iva[df_emp_iva['Cuenta'].str.startswith('101-00011-001-0000')]['Columna E (Cargos)'].sum())
                    v_pag2 = float(df_emp_iva[df_emp_iva['Cuenta'].str.startswith('101-00011-002-0000')]['Columna E (Cargos)'].sum())
                    v_pagado_total = v_pag1 + v_pag2

                    diferencia = v_cobrado - v_pagado_total
                    iva_a_favor = abs(diferencia) if diferencia < 0 else 0.0
                    iva_a_pagar = diferencia if diferencia > 0 else 0.0

                    datos_iva.append({
                        'Empresa': emp,
                        'IVA Cobrado (Cuenta 201-00010-001-0000) [Col. F]': v_cobrado,
                        'IVA Pagado Gastos (Cuenta 101-00011-001-0000) [Col. E]': v_pag1,
                        'IVA Pagado Costos (Cuenta 101-00011-002-0000) [Col. E]': v_pag2,
                        'Total IVA Pagado / Acreditable': v_pagado_total,
                        'IVA a Pagar (A Cargo)': iva_a_pagar,
                        'IVA a Favor': iva_a_favor
                    })

                df_base_iva = pd.DataFrame(datos_iva)

                conceptos_iva = [
                    'IVA Cobrado (Cuenta 201-00010-001-0000) [Col. F]',
                    'IVA Pagado Gastos (Cuenta 101-00011-001-0000) [Col. E]',
                    'IVA Pagado Costos (Cuenta 101-00011-002-0000) [Col. E]',
                    'Total IVA Pagado / Acreditable',
                    'IVA a Pagar (A Cargo)',
                    'IVA a Favor'
                ]

                tabla_resumen_iva = pd.DataFrame({'Concepto': conceptos_iva})

                for _, row in df_base_iva.iterrows():
                    emp_col = row['Empresa']
                    tabla_resumen_iva[emp_col] = [
                        row['IVA Cobrado (Cuenta 201-00010-001-0000) [Col. F]'],
                        row['IVA Pagado Gastos (Cuenta 101-00011-001-0000) [Col. E]'],
                        row['IVA Pagado Costos (Cuenta 101-00011-002-0000) [Col. E]'],
                        row['Total IVA Pagado / Acreditable'],
                        row['IVA a Pagar (A Cargo)'],
                        row['IVA a Favor']
                    ]

                tabla_resumen_iva['TOTAL CONSOLIDADO'] = [
                    df_base_iva[c].sum() for c in conceptos_iva
                ]

                format_dict_iva = {col: "${:,.2f}" for col in tabla_resumen_iva.columns if col != 'Concepto'}

                st.dataframe(
                    tabla_resumen_iva.style.format(format_dict_iva),
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("---")

                periodo_texto_iva = f"{nombres_meses[mes_imp-1]} {anio_imp}"
                excel_iva_bytes = generar_excel_iva(tabla_resumen_iva, periodo_texto_iva)

                st.download_button(
                    label="📥 Descargar Reporte Ejecutivo de IVA en Excel",
                    data=excel_iva_bytes,
                    file_name=f"Reporte_Ejecutivo_IVA_{nombres_meses[mes_imp-1]}_{anio_imp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_iva",
                    use_container_width=True
                )

            else:
                st.warning("No se encontraron saldos de IVA en la balanza.")
        else:
            st.error("No se encontró el archivo de balanza especificado.")


    # ---------------------------------------------------------
    # 3. PESTAÑA RETENCIONES
    # ---------------------------------------------------------
    with tab_retenciones:
        st.markdown(f"### 📋 Resumen General de Retenciones — {nombres_meses[mes_imp-1]} {anio_imp}")
        st.info("Espacio preparado para consolidar cuentas de Retenciones de ISR/IVA (Fletes, Servicios Profesionales, Arrendamiento) en formato multi-empresa.")