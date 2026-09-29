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

# Diccionario de equivalencia oficial: Empresa Origen (Nombre Corto) -> RFC Emisor SAT
MAPEO_EMPRESA_RFC = {
    "CIV": "CIV1009089B0",
    "CIVLAT": "CIV141222JD5",
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

RFC_TO_EMPRESA = {rfc: emp for emp, rfc in MAPEO_EMPRESA_RFC.items() if rfc}

# Mapeo inverso o equivalente para balanzas
MAPEO_NOMBRE_A_CODIGO = {
    "CIVLAT": "0468",
    "CIV": "0467",
    "CIVMEX": "0469",
    "SERVYRE": "2872",
    "SERSEÑAL": "1636",
    "EFCO": "0813",
    "FGS": "0942",
    "FERVIC": "1127",
    "FPSB": "1216",
    "PESAZA": "1144",
    "GPO SERVYRE": "1149",
    "INMOBILIARIA": "1404",
    "LABORATORIO": "1603",
    "LATIN": "1616",
    "LIMPIESPIN": "1626",
    "SERVYCARGO": "2871",
    "VIALTECNO": "3329",
    "SEVILAT": "2937",
    "PROINA": "2396",
    "IPV": "1428",
    "COMANA": "0665"
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


def obtener_estructura_balanzas():
    archivos = glob.glob("Balanzas/**/Balanza.xlsx", recursive=True)
    estructura = []
    for path in archivos:
        path_norm = path.replace("\\", "/")
        partes = path_norm.split("/")
        if len(partes) >= 4:
            anio = partes[-3]
            mes = partes[-2]
            estructura.append({
                'anio': str(anio),
                'mes': str(mes),
                'ruta': path_norm,
                'mtime': os.path.getmtime(path),
            })
    if estructura:
        estructura.sort(key=lambda x: x['mtime'], reverse=True)
    return estructura


@st.cache_data(ttl=600)
def cargar_hoja_balanza(ruta, nombre_hoja):
    return pd.read_excel(ruta, sheet_name=nombre_hoja)


def extraer_datos_contabilidad_balanzas(anio_sel, mes_fin_sel):
    """Extrae la contabilidad de ingresos (410, 411, 423) y notas de crédito (420, 421) de las balanzas."""
    estructura = obtener_estructura_balanzas()
    # Buscar balanza del año y mes final seleccionado (o el más cercano disponible)
    mes_str = f"{int(mes_fin_sel):02d}"
    ruta_balanza = next(
        (x['ruta'] for x in estructura if x['anio'] == str(anio_sel) and (x['mes'] == mes_str or str(int(x['mes'])) == str(mes_fin_sel))),
        None
    )
    if not ruta_balanza and estructura:
        ruta_balanza = estructura[0]['ruta']

    if not ruta_balanza or not os.path.exists(ruta_balanza):
        return {}

    xls = pd.ExcelFile(ruta_balanza)
    sheets = xls.sheet_names
    if len(sheets) > 1 and "Hoja1" in sheets:
        sheets.remove("Hoja1")

    datos_contables = {}
    for emp in sheets:
        df_b_raw = cargar_hoja_balanza(ruta_balanza, emp)
        num_cols = df_b_raw.shape[1]
        col_cta = 0
        col_deudor_f = 6 if num_cols > 6 else num_cols - 2
        col_acreedor_f = 7 if num_cols > 7 else num_cols - 1

        ingresos_cont = 0.0
        nc_cont = 0.0

        for _, row in df_b_raw.iterrows():
            cta_raw = str(row.iloc[col_cta]).strip()
            if not cta_raw or cta_raw.lower() in ('nan', 'cuenta', 'none'):
                continue
            if not re.search(r'\d{3}[-\s]?\d{3,5}', cta_raw):
                continue
            
            d_f = parse_monto_robusto(row.iloc[col_deudor_f])
            a_f = parse_monto_robusto(row.iloc[col_acreedor_f])
            saldo_cta = a_f - d_f  # Naturaleza acreedora para ingresos

            segs = cta_raw.split('-')
            if len(segs) >= 4:
                prefix = segs[0].strip()
                seg2 = segs[1].strip()
                seg3 = segs[2].strip()
                seg4 = segs[3].strip()

                es_base = (seg2 in ('00000', '0000', '0') and seg3 in ('000', '0') and seg4 in ('0000', '0'))

                if es_base:
                    if prefix in ('410', '411', '423'):
                        ingresos_cont += saldo_cta
                    elif prefix in ('420', '421'):
                        nc_cont += abs(saldo_cta)

        datos_contables[emp.upper()] = {
            'Ingresos_Cont': ingresos_cont,
            'NC_Cont': nc_cont
        }

    return datos_contables


def consolidar_excels_sat(carpeta_xmls="XML"):
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

                estado_sat = "VIGENTE"
                for col in df_sat.columns:
                    if col.lower() == 'estado' or 'estatus' in col.lower():
                        val_est = str(row.get(col, '')).strip().upper()
                        if val_est and val_est != 'NAN':
                            estado_sat = val_est
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
                    'Estado_SAT': estado_sat,
                    'SubTotal_SAT': 0.0 if es_cancelada else obtener_val(['subtotal', 'sub total']),
                    'Total_SAT': 0.0 if es_cancelada else obtener_val(['total'])
                })
        except Exception:
            continue
    return pd.DataFrame(registros_xml)


def cargar_base_master_general(ruta_master="Consolidado_Master.xlsx", anio_filtro=None, mes_ini=None, mes_fin=None):
    if not os.path.exists(ruta_master):
        return pd.DataFrame()
    
    dfs_totales = []
    
    try:
        df_fact = pd.read_excel(ruta_master, sheet_name="FacturaCliente")
        if not df_fact.empty:
            df_fact['Tipo_Doc'] = 'Factura'
            dfs_totales.append(df_fact)
    except Exception:
        pass

    try:
        df_nc = pd.read_excel(ruta_master, sheet_name="NotaCreditoCliente")
        if not df_nc.empty:
            df_nc['Tipo_Doc'] = 'NotaCredito'
            cols_numericas = [c for c in ['SubTotal', 'Total'] if c in df_nc.columns]
            for col in cols_numericas:
                df_nc[col] = pd.to_numeric(df_nc[col], errors='coerce').fillna(0.0).abs() * -1
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
    
    return df_master


def exportar_excel_completo(df_facturas_det, df_notas_det, df_facturas_res, df_notas_res, df_dif, titulo_reporte, subtitulo_reporte):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    fill_encabezado = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    font_encabezado = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    borde_delgado = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    def escribir_tabla(ws, start_r, titulo_sec, df):
        ws.cell(row=start_r, column=1, value=titulo_sec).font = Font(name="Calibri", size=12, bold=True, color="1E293B")
        r = start_r + 1
        for c_idx, col_name in enumerate(list(df.columns), start=1):
            cell = ws.cell(row=r, column=c_idx, value=str(col_name))
            cell.fill = fill_encabezado
            cell.font = font_encabezado
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        for r_i, (_, row_series) in enumerate(df.iterrows()):
            curr_row = r + 1 + r_i
            for c_i, col_name in enumerate(list(df.columns), start=1):
                val = row_series[col_name]
                cell = ws.cell(row=curr_row, column=c_i, value=val if pd.notnull(val) else 0.0)
                cell.border = borde_delgado
                if isinstance(val, (int, float)):
                    cell.number_format = '$#,##0.00'
                    cell.alignment = Alignment(horizontal="right")
        return r + len(df) + 3

    ws_res = wb.active
    ws_res.title = "Resumen"
    ws_res.views.sheetView[0].showGridLines = True
    ws_res.cell(row=1, column=1, value=titulo_reporte).font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws_res.cell(row=2, column=1, value=subtitulo_reporte + " | Resumen Ejecutivo").font = Font(name="Calibri", size=10, italic=True, color="475569")
    
    next_r = escribir_tabla(ws_res, 4, "RESUMEN INGRESOS (VIGENTES)", df_facturas_res)
    escribir_tabla(ws_res, next_r, "RESUMEN EGRESOS / NOTAS DE CRÉDITO (VIGENTES)", df_notas_res)

    ws_det = wb.create_sheet(title="Acumulado")
    ws_det.views.sheetView[0].showGridLines = True
    ws_det.cell(row=1, column=1, value=titulo_reporte).font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws_det.cell(row=2, column=1, value=subtitulo_reporte + " | Detalle General").font = Font(name="Calibri", size=10, italic=True, color="475569")
    
    next_r_det = escribir_tabla(ws_det, 4, "ACUMULADO - FACTURAS", df_facturas_det)
    escribir_tabla(ws_det, next_r_det, "ACUMULADO - NOTAS DE CRÉDITO", df_notas_det)

    ws_dif = wb.create_sheet(title="Diferencias UUID")
    ws_dif.views.sheetView[0].showGridLines = True
    ws_dif.cell(row=1, column=1, value=titulo_reporte).font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws_dif.cell(row=2, column=1, value=subtitulo_reporte + " | Auditoría de UUIDs con Diferencia").font = Font(name="Calibri", size=10, italic=True, color="475569")
    
    escribir_tabla(ws_dif, 4, "UUIDs QUE CONFORMAN LA DIFERENCIA", df_dif)

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
    st.markdown("### 💰 Amarre de Ingresos (Facturas y Notas de Crédito)")
    st.info(f"Filtra la base maestra para el año **{anio_sel}** (meses {mes_inicial} a {mes_final}).")

    carpeta_excel_input = st.text_input("Carpeta que contiene los Excel del SAT (ej. XML):", value="XML")

    if st.button("🚀 Ejecutar Amarre y Auditoría de Diferencias"):
        with st.spinner("Procesando información, balanzas contables y auditando diferencias..."):
            df_fact_base = cargar_base_master_general("Consolidado_Master.xlsx", anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
            df_xml_sat = consolidar_excels_sat(carpeta_excel_input)
            datos_contables_dict = extraer_datos_contabilidad_balanzas(anio_sel, mes_final)

            if df_fact_base.empty:
                st.warning(f"No se encontraron registros en el maestro para el periodo seleccionado.")
            elif df_xml_sat.empty:
                st.warning(f"No se encontraron archivos Excel en la carpeta '{carpeta_excel_input}'.")
            else:
                df_amarre = pd.merge(df_fact_base, df_xml_sat, on="UUID", how="outer", suffixes=('', '_SAT'))
                
                def determinar_empresa(row):
                    ben = str(row.get('BusinessEntityName', '')).strip().upper()
                    if ben in MAPEO_EMPRESA_RFC:
                        return ben
                    rfc_sat = str(row.get('RFC_Emisor', '')).strip().upper()
                    if rfc_sat in RFC_TO_EMPRESA:
                        return RFC_TO_EMPRESA[rfc_sat]
                    for emp in MAPEO_EMPRESA_RFC.keys():
                        if emp in ben:
                            return emp
                    return "OTRAS"

                df_amarre['Empresa'] = df_amarre.apply(determinar_empresa, axis=1)
                df_amarre['Origen'] = df_amarre['Empresa'].apply(lambda x: f"En Acumulado {x}" if x != "OTRAS" else "En Acumulado General")

                df_amarre['SubTotal'] = pd.to_numeric(df_amarre.get('SubTotal', 0), errors='coerce').fillna(0.0)
                df_amarre['SubTotal_SAT'] = pd.to_numeric(df_amarre.get('SubTotal_SAT', 0), errors='coerce').fillna(0.0)

                mask_nc = df_amarre['Tipo_Doc'] == 'NotaCredito'
                df_amarre.loc[mask_nc, 'SubTotal'] = df_amarre.loc[mask_nc, 'SubTotal'].abs() * -1
                df_amarre.loc[mask_nc, 'SubTotal_SAT'] = df_amarre.loc[mask_nc, 'SubTotal_SAT'].abs() * -1

                df_amarre['Diferencia'] = df_amarre['SubTotal'] - df_amarre['SubTotal_SAT']

                if 'CFDStatusCancelledName' in df_amarre.columns:
                    df_amarre['Estatus'] = df_amarre['CFDStatusCancelledName'].apply(lambda x: 'Cancelado' if pd.notnull(x) and 'CANCELAD' in str(x).upper() else 'VIGENTE')
                else:
                    df_amarre['Estatus'] = 'VIGENTE'

                df_facturas = df_amarre[df_amarre['Tipo_Doc'] == 'Factura'].copy()
                df_notas = df_amarre[df_amarre['Tipo_Doc'] == 'NotaCredito'].copy()

                def preparar_detalle(df):
                    if df.empty:
                        return pd.DataFrame(columns=['Origen', 'Empresa', 'UUID', 'Tipo Doc', 'Estatus', 'SubTotal Origen', 'SubTotal Destino', 'Diferencia'])
                    cols = ['Origen', 'Empresa', 'UUID', 'Tipo_Doc', 'Estatus', 'SubTotal', 'SubTotal_SAT', 'Diferencia']
                    d = df[[c for c in cols if c in df.columns]].copy()
                    d.columns = ['Origen', 'Empresa', 'UUID', 'Tipo Doc', 'Estatus', 'SubTotal Origen', 'SubTotal Destino', 'Diferencia']
                    return d

                def preparar_resumen_con_contabilidad(df, tipo_doc='Factura'):
                    if df.empty:
                        return pd.DataFrame(columns=['Empresa', 'CONTABILIDAD', 'SISTEMA (VIG)', 'SAT (VIG)', 'DIF. SIST vs SAT'])
                    vig = df[df['Estatus'] == 'VIGENTE']
                    res = vig.groupby('Empresa', as_index=False).agg({
                        'SubTotal': 'sum',
                        'SubTotal_SAT': 'sum',
                        'Diferencia': 'sum'
                    }).rename(columns={
                        'SubTotal': 'SISTEMA (VIG)',
                        'SubTotal_SAT': 'SAT (VIG)',
                        'Diferencia': 'DIF. SIST vs SAT'
                    })

                    # Agregar columna CONTABILIDAD desde las balanzas
                    cont_list = []
                    for emp in res['Empresa']:
                        emp_key = emp.upper()
                        val_cont = 0.0
                        if emp_key in datos_contables_dict:
                            if tipo_doc == 'Factura':
                                val_cont = datos_contables_dict[emp_key]['Ingresos_Cont']
                            else:
                                val_cont = datos_contables_dict[emp_key]['NC_Cont'] * -1 # Mostrar negativo para egresos
                        cont_list.append(val_cont)
                    
                    res.insert(1, 'CONTABILIDAD', cont_list)
                    return res

                df_fact_det = preparar_detalle(df_facturas)
                df_nota_det = preparar_detalle(df_notas)
                df_fact_res = preparar_resumen_con_contabilidad(df_facturas, 'Factura')
                df_nota_res = preparar_resumen_con_contabilidad(df_notas, 'NotaCredito')

                df_diferencias = df_amarre[df_amarre['Diferencia'].round(2) != 0.0].copy()
                df_dif_final = preparar_detalle(df_diferencias)

                st.success("✅ Amarre, notas de crédito, contabilidad y auditoría de diferencias realizadas correctamente.")

                tab_res_ui, tab_det_ui, tab_dif_ui = st.tabs(["📊 Resumen Ejecutivo", "📋 Detalle Acumulado", "🔍 UUIDs con Diferencias"])
                
                with tab_res_ui:
                    st.markdown("### 📈 RESUMEN INGRESOS (VIGENTES)")
                    st.dataframe(df_fact_res.style.format({
                        'CONTABILIDAD': '${:,.2f}',
                        'SISTEMA (VIG)': '${:,.2f}',
                        'SAT (VIG)': '${:,.2f}',
                        'DIF. SIST vs SAT': '${:,.2f}'
                    }), use_container_width=True, hide_index=True)

                    st.markdown("### 📉 RESUMEN EGRESOS / NOTAS DE CRÉDITO (VIGENTES)")
                    st.dataframe(df_nota_res.style.format({
                        'CONTABILIDAD': '${:,.2f}',
                        'SISTEMA (VIG)': '${:,.2f}',
                        'SAT (VIG)': '${:,.2f}',
                        'DIF. SIST vs SAT': '${:,.2f}'
                    }), use_container_width=True, hide_index=True)

                with tab_det_ui:
                    st.markdown("### 📋 ACUMULADO - FACTURAS")
                    st.dataframe(df_fact_det, use_container_width=True)
                    st.markdown("### 📋 ACUMULADO - NOTAS DE CRÉDITO")
                    st.dataframe(df_nota_det, use_container_width=True)

                with tab_dif_ui:
                    st.markdown("### 🔍 UUIDs QUE CONFORMAN LA DIFERENCIA")
                    st.info(f"Se encontraron {len(df_dif_final)} registros con discrepancias entre el Sistema y el SAT.")
                    st.dataframe(df_dif_final, use_container_width=True)

                output_excel = exportar_excel_completo(
                    df_fact_det, df_nota_det, df_fact_res, df_nota_res, df_dif_final,
                    "REPORTE DE AMARRE DE INGRESOS",
                    f"Periodo: {mes_inicial} a {mes_final} del {anio_sel}"
                )
                
                st.download_button(
                    label="📥 Descargar Reporte Completo (Excel con Columna Contabilidad)",
                    data=output_excel,
                    file_name=f"Amarre_Ingresos_Auditoria_{mes_inicial}_a_{mes_final}_{anio_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )