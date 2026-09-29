import glob
import io
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📑 Módulo SAT & Amarres")

# Diccionario de equivalencia: Empresa Origen (BusinessEntityName) -> RFC Emisor SAT
MAPEO_EMPRESA_RFC = {
    "CIV": "CIV1009089B0",
    "CIVLA": "CIV141222JD5",
    "CIVMEX": "CIV141204DN5",
    "EFCO": "EFC840210UI4",
    "FERVIC": "GFE811209FZ2",
    "FGS": "FSI100908CM7",
    "GRUPO FPSB": "GFP1409118E0",
    "GRUPOSERVYRE": "GSE0201315C9",
    "INMOBILIARIA": "IPS1802069U3",
    "LABORATORIO": "LMA160202RS4",
    "LAITS": "LAI130114KT7",
    "LIMPIESPIN": "LIM100513860",
    "PESAZA": "GPE100907IMA",
    "SERSENAL": "", 
    "SERVYCARGO": "SER100803D12",
    "SERVYRE": "SER970728JN8"
}

def parse_monto_robusto(val):
    if pd.isnull(val):
        return 0.0
    val_str = str(val).replace('$', '').replace(',', '').replace(' ', '').strip()
    if not val_str or val_str.lower() in ('nan', 'none', '-'):
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def consolidar_excels_sat(carpeta_xmls="XML"):
    """Lee y consolida los archivos de Excel del SAT, extrayendo UUID, RFC, Nombre del Emisor y Estado."""
    registros_xml = []
    archivos_excel = glob.glob(os.path.join(carpeta_xmls, "**", "*.xlsx"), recursive=True) or glob.glob(os.path.join(carpeta_xmls, "*.xlsx"))
    
    for archivo in archivos_excel:
        try:
            df_sat = pd.read_excel(archivo)
            df_sat.columns = [str(c).strip() for c in df_sat.columns]
            for _, row in df_sat.iterrows():
                uuid = ""
                for col in df_sat.columns:
                    if 'uuid' in col.lower() or 'folio fiscal' in col.lower():
                        val_uuid = str(row.get(col, '')).strip().upper()
                        if val_uuid and val_uuid != 'NAN':
                            uuid = val_uuid
                            break
                if not uuid:
                    continue

                rfc_emisor = ""
                for col in df_sat.columns:
                    if 'rfc' in col.lower() and ('emisor' in col.lower() or 'rfc' == col.lower()):
                        val_rfc = str(row.get(col, '')).strip().upper()
                        if val_rfc and val_rfc != 'NAN':
                            rfc_emisor = val_rfc
                            break

                nombre_emisor_sat = ""
                for col in df_sat.columns:
                    if 'emisor' in col.lower() and ('nombre' in col.lower() or 'razon' in col.lower()):
                        val_nom = str(row.get(col, '')).strip()
                        if val_nom and val_nom.lower() != 'nan':
                            nombre_emisor_sat = val_nom
                            break

                estado_sat = "VIGENTE"
                for col in df_sat.columns:
                    if col.lower() == 'estado' or 'estatus' in col.lower():
                        val_est = str(row.get(col, '')).strip().upper()
                        if val_est and val_est != 'NAN':
                            estado_sat = val_est
                            break

                fecha_emision = ""
                for col in df_sat.columns:
                    if 'fecha' in col.lower():
                        f_raw = str(row.get(col, ''))
                        if f_raw and f_raw != 'NAN':
                            fecha_emision = f_raw.split('T')[0].split(' ')[0]
                            break

                razon_receptor = ""
                for col in df_sat.columns:
                    if 'receptor' in col.lower() and ('nombre' in col.lower() or 'razon' in col.lower()):
                        razon_receptor = str(row.get(col, ''))
                        break

                def obtener_val(keywords):
                    for col in df_sat.columns:
                        if any(k in col.lower() for k in keywords):
                            return parse_monto_robusto(row.get(col, 0.0))
                    return 0.0

                es_cancelada = 'CANCELAD' in estado_sat
                
                registros_xml.append({
                    'UUID': uuid,
                    'RFC_Emisor': rfc_emisor,
                    'Nombre_Emisor_SAT': nombre_emisor_sat,
                    'Estado_SAT': estado_sat,
                    'Fecha emision': fecha_emision,
                    'Razon receptor': razon_receptor if razon_receptor != 'nan' else '',
                    'SubTotal': 0.0 if es_cancelada else obtener_val(['subtotal', 'sub total']),
                    'Descuento': 0.0 if es_cancelada else obtener_val(['descuento', 'desc']),
                    'IVA Trasladado': 0.0 if es_cancelada else obtener_val(['iva trasladado', 'trasladado 002', 'iva 16']),
                    'IVA Retenido': 0.0 if es_cancelada else obtener_val(['iva retenido', 'retenido 002']),
                    'ISR Retenido': 0.0 if es_cancelada else obtener_val(['isr retenido', 'retenido 001']),
                    'Local retenido': 0.0 if es_cancelada else obtener_val(['local retenido', 'impuesto local']),
                    'Total': 0.0 if es_cancelada else obtener_val(['total'])
                })
        except Exception:
            continue
    return pd.DataFrame(registros_xml)


def cargar_base_master_general(ruta_master="Consolidado_Master.xlsx", anio_filtro=None, mes_ini=None, mes_fin=None):
    """Carga y consolida FacturaCliente y NotaCreditoCliente (negativo), filtrando por periodo."""
    if not os.path.exists(ruta_master):
        return pd.DataFrame()
    
    dfs_totales = []
    
    try:
        df_fact = pd.read_excel(ruta_master, sheet_name="FacturaCliente")
        if not df_fact.empty:
            dfs_totales.append(df_fact)
    except Exception:
        pass

    try:
        df_nc = pd.read_excel(ruta_master, sheet_name="NotaCreditoCliente")
        if not df_nc.empty:
            cols_numericas = [c for c in ['SubTotal', 'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total'] if c in df_nc.columns]
            for col in cols_numericas:
                df_nc[col] = pd.to_numeric(df_nc[col], errors='coerce').fillna(0.0) * -1
            dfs_totales.append(df_nc)
    except Exception:
        pass

    if not dfs_totales:
        return pd.DataFrame()

    df_master = pd.concat(dfs_totales, ignore_index=True)

    if 'DateDocument' in df_master.columns:
        df_master['DateDocument'] = pd.to_datetime(df_master['DateDocument'], errors='coerce')
        if anio_filtro:
            df_master = df_master[df_master['DateDocument'].dt.year == int(anio_filtro)]
        if mes_ini and mes_fin:
            df_master = df_master[(df_master['DateDocument'].dt.month >= int(mes_ini)) & (df_master['DateDocument'].dt.month <= int(mes_fin))]
            
    if 'UUID' in df_master.columns:
        df_master = df_master[df_master['UUID'].notnull() & (df_master['UUID'].astype(str).str.strip() != '')]
    
    cols_a_agrupar = [c for c in ['UUID', 'BusinessEntityName', 'CFDStatusCancelledName'] if c in df_master.columns]
    cols_numericas = [c for c in ['SubTotal', 'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total'] if c in df_master.columns]
    
    if not cols_a_agrupar:
        return pd.DataFrame()

    df_agrupado = df_master.groupby(cols_a_agrupar, as_index=False)[cols_numericas].sum()
    
    if 'BusinessEntityName' in df_agrupado.columns:
        df_agrupado['RFC_Esperado'] = df_agrupado['BusinessEntityName'].map(MAPEO_EMPRESA_RFC).fillna('')
        
    return df_agrupado


def exportar_excel_multi_pestana(df_detalle, df_resumen, titulo_reporte, subtitulo_reporte):
    """Genera un archivo Excel profesional con múltiples pestañas y estilos aplicados."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    fill_encabezado = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    font_encabezado = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    borde_delgado = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # --- Pestaña 1: Resumen por Empresa ---
    ws_res = wb.active
    ws_res.title = "Resumen por Empresa"
    ws_res.views.sheetView[0].showGridLines = True
    
    ws_res.cell(row=1, column=1, value=titulo_reporte).font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws_res.cell(row=2, column=1, value=subtitulo_reporte + " | Resumen Ejecutivo").font = Font(name="Calibri", size=10, italic=True, color="475569")
    
    start_row = 4
    for c_idx, col_name in enumerate(list(df_resumen.columns), start=1):
        cell = ws_res.cell(row=start_row, column=c_idx, value=str(col_name))
        cell.fill = fill_encabezado
        cell.font = font_encabezado
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_i, (_, row_series) in enumerate(df_resumen.iterrows()):
        curr_row = start_row + 1 + r_i
        for c_i, col_name in enumerate(list(df_resumen.columns), start=1):
            val = row_series[col_name]
            cell = ws_res.cell(row=curr_row, column=c_i, value=val if pd.notnull(val) else 0.0)
            cell.border = borde_delgado
            if isinstance(val, (int, float)):
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal="right")

    # --- Pestaña 2: Detalle Amarre ---
    ws_det = wb.create_sheet(title="Detalle Amarre")
    ws_det.views.sheetView[0].showGridLines = True
    
    ws_det.cell(row=1, column=1, value=titulo_reporte).font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws_det.cell(row=2, column=1, value=subtitulo_reporte + " | Detalle por UUID").font = Font(name="Calibri", size=10, italic=True, color="475569")
    
    for c_idx, col_name in enumerate(list(df_detalle.columns), start=1):
        cell = ws_det.cell(row=start_row, column=c_idx, value=str(col_name))
        cell.fill = fill_encabezado
        cell.font = font_encabezado
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_i, (_, row_series) in enumerate(df_detalle.iterrows()):
        curr_row = start_row + 1 + r_i
        for c_i, col_name in enumerate(list(df_detalle.columns), start=1):
            val = row_series[col_name]
            cell = ws_det.cell(row=curr_row, column=c_i, value=val if pd.notnull(val) else 0.0)
            cell.border = borde_delgado
            if isinstance(val, (int, float)):
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal="right")

    wb.save(output)
    return output.getvalue()


# --- INTERFAZ DEL MÓDULO SAT ---
col_anio, col_mini, col_mfin = st.columns([1, 1, 1])

with col_anio:
    anio_sel = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0)
with col_mini:
    mes_inicial = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'][x-1])
with col_mfin:
    mes_final = st.selectbox("Mes Final:", list(range(1, 13)), index=7, format_func=lambda x: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'][x-1])

st.markdown("---")

subtab_rh, subtab_ingresos = st.tabs(["👥 Amarre RH", "💰 Amarre Ingresos"])

with subtab_rh:
    st.markdown("### 👥 Amarre de Recursos Humanos (RH)")
    st.info("Módulo para validación de nóminas y retenciones de sueldos y salarios contra CONTPAQi y SAT.")

with subtab_ingresos:
    st.markdown("### 💰 Amarre de Ingresos (Facturas + Notas de Crédito vs Excel del SAT)")
    st.info(f"Filtra la base maestra para el año **{anio_sel}** (meses {mes_inicial} a {mes_final}), aplicando notas de crédito en negativo y comparando contra el SAT.")

    carpeta_excel_input = st.text_input("Carpeta que contiene los Excel del SAT (ej. XML):", value="XML")

    if st.button("🚀 Ejecutar Amarre y Resumen por Empresa"):
        with st.spinner("Procesando facturas, notas de crédito y archivos Excel del SAT..."):
            df_fact_base = cargar_base_master_general("Consolidado_Master.xlsx", anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
            df_xml_sat = consolidar_excels_sat(carpeta_excel_input)

            if df_fact_base.empty:
                st.warning(f"No se encontraron registros en el maestro para el periodo seleccionado ({mes_inicial} a {mes_fin} de {anio_sel}).")
            elif df_xml_sat.empty:
                st.warning(f"No se encontraron archivos Excel en la carpeta '{carpeta_excel_input}'.")
            else:
                df_amarre = pd.merge(df_fact_base, df_xml_sat, on="UUID", how="outer", suffixes=('_Fact', '_SAT'))
                
                tot_fact = 'Total_Fact' if 'Total_Fact' in df_amarre.columns else 'Total_x'
                tot_sat = 'Total_SAT' if 'Total_SAT' in df_amarre.columns else 'Total_y'
                
                df_amarre['Diferencia_Total'] = df_amarre[tot_fact].fillna(0.0) - df_amarre[tot_sat].fillna(0.0)
                
                if 'RFC_Esperado' in df_amarre.columns and 'RFC_Emisor' in df_amarre.columns:
                    df_amarre['Validacion_RFC'] = df_amarre.apply(
                        lambda row: 'Coincide' if str(row['RFC_Emisor']).strip() == str(row['RFC_Esperado']).strip() else 'Diferente/Revisar',
                        axis=1
                    )

                col_empresa = 'BusinessEntityName' if 'BusinessEntityName' in df_amarre.columns else None
                if col_empresa:
                    df_resumen_empresa = df_amarre.groupby(col_empresa, as_index=False).agg({
                        tot_fact: 'sum',
                        tot_sat: 'sum',
                        'Diferencia_Total': 'sum'
                    }).rename(columns={
                        tot_fact: 'Total_Maestro',
                        tot_sat: 'Total_SAT',
                        'Diferencia_Total': 'Diferencia'
                    })
                else:
                    df_resumen_empresa = pd.DataFrame()

                st.success("✅ Amarre y resumen por empresa realizados correctamente.")

                tab_res_ui, tab_det_ui = st.tabs(["📊 Resumen por Empresa", "📋 Detalle por UUID"])
                
                with tab_res_ui:
                    st.markdown("### Resumen Total por Empresa (Maestro vs SAT)")
                    st.dataframe(df_resumen_empresa, use_container_width=True)

                with tab_det_ui:
                    st.markdown("### Detalle Completo de Amarre")
                    st.dataframe(df_amarre, use_container_width=True)

                output_excel = exportar_excel_multi_pestana(
                    df_amarre,
                    df_resumen_empresa,
                    "REPORTE DE AMARRE DE INGRESOS",
                    f"Periodo: {mes_inicial} a {mes_final} del {anio_sel}"
                )
                
                st.download_button(
                    label="📥 Descargar Reporte Completo (Excel con Resumen y Detalle)",
                    data=output_excel,
                    file_name=f"Amarre_Ingresos_Resumen_Y_Detalle_{mes_inicial}_a_{mes_final}_{anio_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )