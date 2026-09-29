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
    """Lee y consolida los archivos de Excel del SAT, extrayendo UUID, RFC Emisor y Estado."""
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

                # Extracción del RFC Emisor
                rfc_emisor = ""
                for col in df_sat.columns:
                    if 'rfc' in col.lower() and ('emisor' in col.lower() or 'rfc' == col.lower()):
                        val_rfc = str(row.get(col, '')).strip().upper()
                        if val_rfc and val_rfc != 'NAN':
                            rfc_emisor = val_rfc
                            break

                # Extracción del Estado de la factura en el SAT
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

                # Si el estado es cancelado, los montos deben ser 0.0 por regla de negocio solicitada
                es_cancelada = 'CANCELAD' in estado_sat
                
                registros_xml.append({
                    'UUID': uuid,
                    'RFC_Emisor': rfc_emisor,
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


def cargar_base_facturacion_master(ruta_master="Consolidado_Master.xlsx", anio_filtro=None):
    """Carga y filtra la pestaña FacturaCliente por Año usando DateDocument, incluyendo BusinessEntityName y CFDStatusCancelledName."""
    if not os.path.exists(ruta_master):
        return pd.DataFrame()
    try:
        df = pd.read_excel(ruta_master, sheet_name="FacturaCliente")
        if 'DateDocument' in df.columns:
            df['DateDocument'] = pd.to_datetime(df['DateDocument'], errors='coerce')
            if anio_filtro:
                df = df[df['DateDocument'].dt.year == int(anio_filtro)]
        if 'UUID' in df.columns:
            df = df[df['UUID'].notnull() & (df['UUID'].astype(str).str.strip() != '')]
        
        # Identificar columnas disponibles para agrupar y sumar
        cols_a_agrupar = [c for c in ['UUID', 'BusinessEntityName', 'CFDStatusCancelledName'] if c in df.columns]
        cols_numericas = [c for c in ['SubTotal', 'TotalDiscount', 'TotalTax', 'TotalRetention', 'Total'] if c in df.columns]
        
        if not cols_a_agrupar:
            return pd.DataFrame()

        # Agrupación conservando Empresa Origen y estatus de cancelación del maestro
        df_agrupado = df.groupby(cols_a_agrupar, as_index=False)[cols_numericas].sum()
        
        if 'BusinessEntityName' in df_agrupado.columns:
            df_agrupado['RFC_Esperado'] = df_agrupado['BusinessEntityName'].map(MAPEO_EMPRESA_RFC).fillna('')
            
        return df_agrupado
    except Exception:
        return pd.DataFrame()


def exportar_excel_matriz_individual(df_matriz, titulo_reporte, subtitulo_reporte):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte"
    ws.views.sheetView[0].showGridLines = True
    ws.cell(row=1, column=1, value=titulo_reporte).font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws.cell(row=2, column=1, value=subtitulo_reporte).font = Font(name="Calibri", size=10, italic=True, color="475569")
    
    start_row = 4
    ws.cell(row=start_row, column=1, value="EMPRESA").fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws.cell(row=start_row, column=1).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for c_idx, col_name in enumerate(list(df_matriz.columns), start=2):
        cell = ws.cell(row=start_row, column=c_idx, value=str(col_name))
        cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    for r_i, (idx_name, row_series) in enumerate(df_matriz.iterrows()):
        curr_row = start_row + 1 + r_i
        ws.cell(row=curr_row, column=1, value=str(idx_name))
        for c_i, col_name in enumerate(list(df_matriz.columns), start=2):
            val = row_series[col_name]
            cell = ws.cell(row=curr_row, column=c_i, value=val if pd.notnull(val) else 0.0)
            if isinstance(val, (int, float)):
                cell.number_format = '$#,##0.00'
    wb.save(output)
    return output.getvalue()


# --- INTERFAZ DEL MÓDULO SAT ---
anio_sel = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0)

st.markdown("---")

subtab_rh, subtab_ingresos = st.tabs(["👥 Amarre RH", "💰 Amarre Ingresos"])

with subtab_rh:
    st.markdown("### 👥 Amarre de Recursos Humanos (RH)")
    st.info("Módulo para validación de nóminas y retenciones de sueldos y salarios contra CONTPAQi y SAT.")

with subtab_ingresos:
    st.markdown("### 💰 Amarre de Ingresos (Base de Facturación por Empresa vs Excel del SAT - Anual)")
    st.info(f"Filtra la pestaña `FacturaCliente` de `Consolidado_Master.xlsx` para el año **{anio_sel}**, considerando empresa de origen, estatus de cancelación y validación de RFC.")

    carpeta_excel_input = st.text_input("Carpeta que contiene los Excel del SAT (ej. XML):", value="XML")

    if st.button("🚀 Ejecutar Amarre por Empresa / RFC"):
        with st.spinner("Procesando base de facturación y archivos Excel del SAT..."):
            df_fact_base = cargar_base_facturacion_master("Consolidado_Master.xlsx", anio_filtro=anio_sel)
            df_xml_sat = consolidar_excels_sat(carpeta_excel_input)

            if df_fact_base.empty:
                st.warning(f"No se encontraron registros en 'FacturaCliente' para el año {anio_sel}.")
            elif df_xml_sat.empty:
                st.warning(f"No se encontraron archivos Excel en la carpeta '{carpeta_excel_input}'.")
            else:
                # Cruzamos por UUID manteniendo la trazabilidad completa
                df_amarre = pd.merge(df_fact_base, df_xml_sat, on="UUID", how="outer", suffixes=('_Fact', '_SAT'))
                
                tot_fact = 'Total_Fact' if 'Total_Fact' in df_amarre.columns else 'Total_x'
                tot_sat = 'Total_SAT' if 'Total_SAT' in df_amarre.columns else 'Total_y'
                
                df_amarre['Diferencia_Total'] = df_amarre[tot_fact].fillna(0.0) - df_amarre[tot_sat].fillna(0.0)
                
                # Validación cruzada de RFC
                if 'RFC_Esperado' in df_amarre.columns and 'RFC_Emisor' in df_amarre.columns:
                    df_amarre['Validacion_RFC'] = df_amarre.apply(
                        lambda row: 'Coincide' if str(row['RFC_Emisor']).strip() == str(row['RFC_Esperado']).strip() else 'Diferente/Revisar',
                        axis=1
                    )

                st.success("✅ Amarre por empresa, estatus y RFC realizado correctamente.")
                st.dataframe(df_amarre, use_container_width=True)

                output_ingresos = exportar_excel_matriz_individual(
                    df_amarre.set_index('UUID') if 'UUID' in df_amarre.columns else df_amarre,
                    "REPORTE DE AMARRE DE INGRESOS - ANUAL",
                    f"Ejercicio Fiscal: {anio_sel} | Base Facturación vs Excel SAT"
                )
                st.download_button(
                    label="📥 Descargar Excel - Amarre por Empresa",
                    data=output_ingresos,
                    file_name=f"Amarre_Ingresos_Empresas_{anio_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )