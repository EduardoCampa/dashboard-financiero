import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT y Conciliación Master vs XML",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📑 Módulo SAT - Comparativo Master vs XML, Contabilidad y Conciliación")

# --- DICCIONARIO OFICIAL UNIFICADO SEGÚN TU TABLA ---
MAPEO_UNIFICADO = {
    "COMERCIALIZADORA DE INFRAESTRUCTURA VIAL LATINOAMERICANA": "CIVLAT",
    "CIVLAT": "CIVLAT",
    "CIVLA": "CIVLAT",
    "COMERCIALIZADORA DE INFRAESTRUCTURA VIAL": "CIV",
    "CIV": "CIV",
    "CIVMEX": "CIVMEX",
    "ELEMENTOS FABRICADOS Y CONSTRUCCIONES": "EFCO",
    "EFCO": "EFCO",
    "FGS SISTEMAS INTEGRALES DE MANTENIMIENTO": "FGS",
    "FGS": "FGS",
    "GRUPO FERVIC": "FERVIC",
    "FERVIC": "FERVIC",
    "GRUPO FPSB": "FPSB",
    "FPSB": "FPSB",
    "GRUPO PESAZA": "PESAZA",
    "PESAZA": "PESAZA",
    "GRUPO SERVYRE": "GPO SERVYRE",
    "GPO SERVYRE": "GPO SERVYRE",
    "GRUPOSERVYRE": "GPO SERVYRE",
    "INMOBILIARIA PSZ": "INMOBILIARIA",
    "INMOBILIARIA": "INMOBILIARIA",
    "LABORATORIO MAJONINMAR": "LABORATORIO",
    "LABORATORIO": "LABORATORIO",
    "LATIN AMERICAN ITS": "LATIN",
    "LATIN": "LATIN",
    "LAITS": "LATIN",
    "LIMPIESPIN": "LIMPIESPIN",
    "SERVYCARGO": "SERVYCARGO",
    "SERVYRE": "SERVYRE"
}

def normalizar_empresa(razon_o_empresa):
    val = str(razon_o_empresa).strip().upper()
    if not val or val in ('NAN', '0', '0.0', '0.00000', 'NONE'):
        return "OTRAS"
    
    if val in MAPEO_UNIFICADO:
        return MAPEO_UNIFICADO[val]
    
    for clave in sorted(MAPEO_UNIFICADO.keys(), key=len, reverse=True):
        if clave in val:
            return MAPEO_UNIFICADO[clave]
            
    return val


# --- FUNCIONES DE PROCESAMIENTO SAT (XMLs) ---
def cargar_y_procesar_sat(carpeta_xmls="XML", anio_filtro=None, mes_ini=None, mes_fin=None):
    archivos_excel = (
        glob.glob(os.path.join(carpeta_xmls, "**", "*.xlsx"), recursive=True) + 
        glob.glob(os.path.join(carpeta_xmls, "*.xlsx")) + 
        glob.glob("*.xlsx")
    )
    archivos_excel = list(set(archivos_excel))
    
    registros_ingresos = []
    registros_egresos = []
    
    for archivo in archivos_excel:
        if any(x in archivo for x in ["Consolidado", "Balanza", "FORMATO"]):
            continue
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

                estado = "VIGENTE"
                for col in df_sat.columns:
                    if col.lower() == 'estado' or 'estatus' in col.lower():
                        val_est = str(row.get(col, '')).strip().upper()
                        if val_est and val_est != 'NAN':
                            estado = val_est
                            break
                
                if 'VIGENTE' not in estado:
                    continue

                fecha_emision_str = ""
                dt_fecha = None
                for col in df_sat.columns:
                    if 'fecha' in col.lower() and 'emision' in col.lower():
                        val_fecha = row.get(col, '')
                        if pd.notnull(val_fecha):
                            dt_fecha = pd.to_datetime(val_fecha, errors='coerce')
                            if pd.notnull(dt_fecha):
                                fecha_emision_str = dt_fecha.strftime('%Y-%m-%d')
                        break
                
                if dt_fecha is None:
                    continue

                if anio_filtro and dt_fecha.year != int(anio_filtro):
                    continue
                if mes_ini and mes_fin and (dt_fecha.month < int(mes_ini) or dt_fecha.month > int(mes_fin)):
                    continue

                tipo_doc = ""
                for col in df_sat.columns:
                    if col.lower() == 'tipo':
                        tipo_doc = str(row.get(col, '')).strip()
                        break

                razon_emisor = ""
                for col in df_sat.columns:
                    if 'razon' in col.lower() and 'emisor' in col.lower():
                        val_razon = str(row.get(col, '')).strip()
                        if val_razon and val_razon.upper() != 'NAN':
                            razon_emisor = val_razon
                            break

                empresa_normalizada = normalizar_empresa(razon_emisor)
                if empresa_normalizada == "OTRAS":
                    base_nom = os.path.basename(archivo).upper()
                    for k, v in MAPEO_UNIFICADO.items():
                        if k in base_nom:
                            empresa_normalizada = v
                            break

                subtotal = 0.0
                descuento = 0.0
                for col in df_sat.columns:
                    c_low = col.lower()
                    if c_low == 'subtotal' or c_low == 'sub total':
                        val = str(row.get(col, 0)).replace('$', '').replace(',', '').strip()
                        subtotal = float(val) if val and val.lower() != 'nan' else 0.0
                    elif 'descuento' in c_low:
                        val = str(row.get(col, 0)).replace('$', '').replace(',', '').strip()
                        descuento = float(val) if val and val.lower() != 'nan' else 0.0

                subtotal_neto = subtotal - descuento

                registro = {
                    'Empresa': empresa_normalizada,
                    'Razon emisor': razon_emisor if razon_emisor else "SIN RAZÓN EMISOR",
                    'UUID': uuid,
                    'Fecha emision': fecha_emision_str,
                    'Estado_SAT': estado,
                    'SubTotal': subtotal_neto
                }

                if 'I - Ingreso' in tipo_doc or tipo_doc.startswith('I'):
                    registros_ingresos.append(registro)
                elif 'E - Egreso' in tipo_doc or tipo_doc.startswith('E'):
                    registros_egresos.append(registro)

        except Exception as e:
            continue

    return pd.DataFrame(registros_ingresos), pd.DataFrame(registros_egresos)


# --- FUNCIONES DE PROCESAMIENTO CONSOLIDADO MASTER ---
def cargar_y_procesar_master(ruta_master="Consolidado_Master.xlsx", anio_filtro=None, mes_ini=None, mes_fin=None):
    if not os.path.exists(ruta_master):
        return pd.DataFrame(), pd.DataFrame()

    def procesar_hoja(nombre_hoja, es_egreso=False):
        try:
            df = pd.read_excel(ruta_master, sheet_name=nombre_hoja)
            if df.empty:
                return pd.DataFrame()

            if 'UUID' in df.columns:
                df = df[df['UUID'].notnull() & (df['UUID'].astype(str).str.strip() != '') & (df['UUID'].astype(str).str.upper() != 'NAN')].copy()
            else:
                return pd.DataFrame()

            if 'CFDStatusCancelledName' in df.columns:
                df = df[df['CFDStatusCancelledName'].astype(str).str.strip().str.upper() == 'VIGENTE'].copy()

            fecha_col_name = 'DateDocument' if 'DateDocument' in df.columns else 'CFDIFechaCertificacion'
            if fecha_col_name in df.columns:
                df['Fecha_Doc'] = pd.to_datetime(df[fecha_col_name], errors='coerce')
                df = df[df['Fecha_Doc'].notnull()].copy()
                if anio_filtro:
                    df = df[df['Fecha_Doc'].dt.year == int(anio_filtro)]
                if mes_ini and mes_fin:
                    df = df[(df['Fecha_Doc'].dt.month >= int(mes_ini)) & (df['Fecha_Doc'].dt.month <= int(mes_fin))]
                df['Fecha_Doc'] = df['Fecha_Doc'].dt.strftime('%Y-%m-%d')
            else:
                df['Fecha_Doc'] = ''

            subtotal = pd.to_numeric(df.get('SubTotal', 0), errors='coerce').fillna(0.0)
            descuento = pd.to_numeric(df.get('TotalDiscount', 0), errors='coerce').fillna(0.0)
            subtotal_neto = subtotal - descuento

            if es_egreso:
                subtotal_neto = subtotal_neto.abs() * -1

            empresa_raw = df['EmpresaOrigen'] if 'EmpresaOrigen' in df.columns else 'SIN EMPRESA'
            empresa_norm = empresa_raw.apply(normalizar_empresa)

            df_final = pd.DataFrame({
                'Empresa': empresa_norm,
                'UUID': df['UUID'],
                'CFDIFechaCertificacion': df['Fecha_Doc'],
                'Estado_Master': 'VIGENTE',
                'SubTotal': subtotal_neto
            })

            return df_final
        except Exception as e:
            return pd.DataFrame()

    df_ingresos_master = procesar_hoja("FacturaCliente", es_egreso=False)
    df_egresos_master = procesar_hoja("NotaCreditoCliente", es_egreso=True)

    return df_ingresos_master, df_egresos_master


# --- CARGAR BALANZAS Y DETALLE POR MES ---
def cargar_balanzas_por_mes(anio=2026, mes=8, ruta_base="Balanzas"):
    mes_str = f"{int(mes):02d}"
    ruta_balanza = os.path.join(ruta_base, str(anio), mes_str, "Balanza.xlsx")
    
    if not os.path.exists(ruta_balanza):
        posibles = glob.glob(os.path.join(ruta_base, "**", mes_str, "*.xlsx"), recursive=True)
        if posibles:
            ruta_balanza = posibles[0]
        else:
            return pd.DataFrame(columns=['Empresa', 'Contabilidad Ingresos', 'Contabilidad Egresos']), pd.DataFrame()

    resultados = []
    detalles = []
    try:
        xls = pd.ExcelFile(ruta_balanza)
        for hoja in xls.sheet_names:
            if hoja.lower() in ['hoja1', 'resumen']:
                continue
            empresa_normalizada = normalizar_empresa(hoja)
            df_hoja = pd.read_excel(ruta_balanza, sheet_name=hoja)
            if df_hoja.empty:
                continue
            
            num_cols = df_hoja.shape[1]
            if num_cols < 4:
                continue
                
            col_cta = 0
            col_nom = 1
            col_deudor_f = num_cols - 2
            col_acreedor_f = num_cols - 1
            
            val_410 = 0.0
            val_411 = 0.0
            val_423 = 0.0
            val_420 = 0.0
            val_421 = 0.0
            
            for _, row in df_hoja.iterrows():
                cta = str(row.iloc[col_cta]).strip()
                if not cta or cta.lower() in ('nan', 'cuenta', 'none'):
                    continue
                
                cta_limpia = cta.replace(' ', '')
                nom = str(row.iloc[col_nom]) if num_cols > col_nom else ""
                
                monto_ac = float(pd.to_numeric(row.iloc[col_acreedor_f], errors='coerce') or 0.0)
                monto_de = float(pd.to_numeric(row.iloc[col_deudor_f], errors='coerce') or 0.0)
                
                match_cuenta = None
                tipo_reg = None
                
                if cta_limpia.startswith('410-00000-000-0000'):
                    val_410 += monto_ac
                    match_cuenta = '410-00000-000-0000'
                    tipo_reg = 'INGRESOS (+)'
                elif cta_limpia.startswith('411-00000-000-0000'):
                    val_411 += monto_ac
                    match_cuenta = '411-00000-000-0000'
                    tipo_reg = 'INGRESOS (+)'
                elif cta_limpia.startswith('423-00000-000-0000'):
                    val_423 += monto_ac
                    match_cuenta = '423-00000-000-0000'
                    tipo_reg = 'INGRESOS (-)'
                elif cta_limpia.startswith('420-00000-000-0000'):
                    val_420 += monto_de
                    match_cuenta = '420-00000-000-0000'
                    tipo_reg = 'EGRESOS (+)'
                elif cta_limpia.startswith('421-00000-000-0000'):
                    val_421 += monto_de
                    match_cuenta = '421-00000-000-0000'
                    tipo_reg = 'EGRESOS (+)'

                if match_cuenta:
                    detalles.append({
                        'Empresa': empresa_normalizada,
                        'Hoja Balanza': hoja,
                        'Tipo': tipo_reg,
                        'Cuenta': match_cuenta,
                        'Nombre Cuenta': nom,
                        'Saldo Deudor Final': monto_de,
                        'Saldo Acreedor Final': monto_ac
                    })
            
            ingresos_cont = float(val_410) + float(val_411) - float(val_423)
            # Egresos en negativo para el resumen general
            egresos_cont = -1 * (float(val_420) + float(val_421))
            
            resultados.append({
                'Empresa': empresa_normalizada,
                'Contabilidad Ingresos': ingresos_cont,
                'Contabilidad Egresos': egresos_cont
            })
    except Exception as e:
        pass

    df_res_bal = pd.DataFrame(resultados)
    if not df_res_bal.empty:
        df_res_bal = df_res_bal.groupby('Empresa', as_index=False)[['Contabilidad Ingresos', 'Contabilidad Egresos']].sum()

    return df_res_bal, pd.DataFrame(detalles)


# --- CONTROLES DE FILTRO POR PERIODO ---
nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

col_a, col_mini, col_mfin = st.columns([1, 1, 1])

with col_a:
    anio_sel = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0)
with col_mini:
    mes_inicial = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: nombres_meses[x-1])
with col_mfin:
    mes_final = st.selectbox("Mes Final:", list(range(1, 13)), index=7, format_func=lambda x: nombres_meses[x-1])

st.markdown("---")
carpeta_input = st.text_input("Carpeta o ubicación de los archivos del SAT:", value="XML")
ruta_master_input = st.text_input("Archivo Consolidado Master:", value="Consolidado_Master.xlsx")
ruta_balanzas_input = st.text_input("Carpeta Raíz de Balanzas:", value="Balanzas")

if st.button("🚀 Ejecutar Procesamiento Completo"):
    with st.spinner("Procesando información, balanzas por mes y conciliando..."):
        df_ingresos_sat, df_egresos_sat = cargar_y_procesar_sat(carpeta_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
        df_ingresos_master, df_egresos_master = cargar_y_procesar_master(ruta_master_input, anio_filtro=anio_sel, mes_ini=mes_inicial, mes_fin=mes_final)
        
        df_balanzas, df_detalle_cont = cargar_balanzas_por_mes(anio=anio_sel, mes=mes_final, ruta_base=ruta_balanzas_input)

        st.success("✅ Procesamiento completado con éxito.")

        tab_comp, tab_xml, tab_master, tab_conc, tab_cont = st.tabs([
            "⚖️ 1.- Comparativo General", 
            "📑 2.- Ingresos y Egresos de los XML", 
            "📁 3.- Ingresos y Egresos Master",
            "🔍 4.- Conciliacion por UUID",
            "📊 5.- Contabilidad"
        ])

        with tab_comp:
            st.markdown(f"### ⚖ Comparativo General (XML vs Master vs Contabilidad) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
            
            res_ing_sat = df_ingresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
            res_ing_sat.columns = ['Empresa', 'XML Ingresos']

            res_ing_mast = df_ingresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_ingresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
            res_ing_mast.columns = ['Empresa', 'Master Ingresos']

            st.markdown("#### 📈 Comparativo de Ingresos Vigentes")
            if not res_ing_sat.empty or not res_ing_mast.empty or not df_balanzas.empty:
                # Generar lista completa de empresas unificadas
                empresas_todas = set(
                    list(res_ing_sat['Empresa'].unique() if not res_ing_sat.empty else []) + 
                    list(res_ing_mast['Empresa'].unique() if not res_ing_mast.empty else []) + 
                    list(df_balanzas['Empresa'].unique() if not df_balanzas.empty else [])
                )
                df_base_empresas = pd.DataFrame({'Empresa': list(empresas_todas)})

                df_comp_ing = df_base_empresas.copy()
                if not res_ing_mast.empty:
                    df_comp_ing = pd.merge(df_comp_ing, res_ing_mast, on='Empresa', how='left')
                else:
                    df_comp_ing['Master Ingresos'] = 0.0

                if not res_ing_sat.empty:
                    df_comp_ing = pd.merge(df_comp_ing, res_ing_sat, on='Empresa', how='left')
                else:
                    df_comp_ing['XML Ingresos'] = 0.0

                if not df_balanzas.empty and 'Contabilidad Ingresos' in df_balanzas.columns:
                    df_comp_ing = pd.merge(df_comp_ing, df_balanzas[['Empresa', 'Contabilidad Ingresos']], on='Empresa', how='left')
                else:
                    df_comp_ing['Contabilidad Ingresos'] = 0.0

                df_comp_ing = df_comp_ing.fillna(0.0)
                    
                df_comp_ing['Diferencia (Master vs XML)'] = df_comp_ing['Master Ingresos'] - df_comp_ing['XML Ingresos']
                df_comp_ing['Diferencia (Contabilidad vs XML)'] = df_comp_ing['Contabilidad Ingresos'] - df_comp_ing['XML Ingresos']
                
                df_comp_ing = df_comp_ing[[
                    'Empresa', 'Master Ingresos', 'XML Ingresos', 'Contabilidad Ingresos', 
                    'Diferencia (Master vs XML)', 'Diferencia (Contabilidad vs XML)'
                ]]
                
                st.dataframe(
                    df_comp_ing.style.format({
                        'Master Ingresos': '${:,.2f}',
                        'XML Ingresos': '${:,.2f}',                         'Contabilidad Ingresos': '${:,.2f}',
                        'Diferencia (Master vs XML)': '${:,.2f}',                         'Diferencia (Contabilidad vs XML)': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No hay datos de ingresos para comparar.")

            st.markdown("---")
            res_eg_sat = df_egresos_sat.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_sat.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
            res_eg_sat.columns = ['Empresa', 'XML Egresos']

            res_eg_mast = df_egresos_master.groupby('Empresa', as_index=False)['SubTotal'].sum() if not df_egresos_master.empty else pd.DataFrame(columns=['Empresa', 'SubTotal'])
            res_eg_mast.columns = ['Empresa', 'Master Egresos']

            st.markdown("#### 📉 Comparativo de Egresos Vigentes / Notas de Crédito")
            if not res_eg_sat.empty or not res_eg_mast.empty or not df_balanzas.empty:
                empresas_eg_todas = set(
                    list(res_eg_sat['Empresa'].unique() if not res_eg_sat.empty else []) + 
                    list(res_eg_mast['Empresa'].unique() if not res_eg_mast.empty else []) + 
                    list(df_balanzas['Empresa'].unique() if not df_balanzas.empty else [])
                )
                df_base_eg_empresas = pd.DataFrame({'Empresa': list(empresas_eg_todas)})

                df_comp_eg = df_base_eg_empresas.copy()
                if not res_eg_mast.empty:
                    df_comp_eg = pd.merge(df_comp_eg, res_eg_mast, on='Empresa', how='left')
                else:
                    df_comp_eg['Master Egresos'] = 0.0

                if not res_eg_sat.empty:
                    df_comp_eg = pd.merge(df_comp_eg, res_eg_sat, on='Empresa', how='left')
                else:
                    df_comp_eg['XML Egresos'] = 0.0

                if not df_balanzas.empty and 'Contabilidad Egresos' in df_balanzas.columns:
                    df_comp_eg = pd.merge(df_comp_eg, df_balanzas[['Empresa', 'Contabilidad Egresos']], on='Empresa', how='left')
                else:
                    df_comp_eg['Contabilidad Egresos'] = 0.0

                df_comp_eg = df_comp_eg.fillna(0.0)
                    
                df_comp_eg['Diferencia (Master vs XML)'] = df_comp_eg['Master Egresos'] - df_comp_eg['XML Egresos']
                df_comp_eg['Diferencia (Contabilidad vs XML)'] = df_comp_eg['Contabilidad Egresos'] - df_comp_eg['XML Egresos']
                
                df_comp_eg = df_comp_eg[[
                    'Empresa', 'Master Egresos', 'XML Egresos', 'Contabilidad Egresos', 
                    'Diferencia (Master vs XML)', 'Diferencia (Contabilidad vs XML)'
                ]]
                
                st.dataframe(
                    df_comp_eg.style.format({
                        'Master Egresos': '${:,.2f}',
                        'XML Egresos': '${:,.2f}',                         'Contabilidad Egresos': '${:,.2f}',
                        'Diferencia (Master vs XML)': '${:,.2f}',                         'Diferencia (Contabilidad vs XML)': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No hay datos de egresos para comparar.")

        with tab_xml:
            st.markdown(f"### 📑 Ingresos y Egresos de los XML (Vigentes) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
            subtab_x_ing, subtab_x_eg = st.tabs(["📈 Ingresos XML", "📉 Egresos XML"])
            
            with subtab_x_ing:
                if not df_ingresos_sat.empty:
                    st.dataframe(df_ingresos_sat.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                    st.metric("Total Ingresos XML", f"${df_ingresos_sat['SubTotal'].sum():,.2f}")
                else:
                    st.warning("Sin registros de ingresos en XML.")

            with subtab_x_eg:
                if not df_egresos_sat.empty:
                    st.dataframe(df_egresos_sat.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                    st.metric("Total Egresos XML", f"${df_egresos_sat['SubTotal'].sum():,.2f}")
                else:
                    st.warning("Sin registros de egresos en XML.")

        with tab_master:
            st.markdown(f"### 📁 Ingresos y Egresos Master (Vigentes) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
            subtab_m_ing, subtab_m_eg = st.tabs(["📈 Ingresos Master", "📉 Egresos Master"])
            
            with subtab_m_ing:
                if not df_ingresos_master.empty:
                    st.dataframe(df_ingresos_master.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                    st.metric("Total Ingresos Master", f"${df_ingresos_master['SubTotal'].sum():,.2f}")
                else:
                    st.warning("Sin registros de ingresos en Master.")

            with subtab_m_eg:
                if not df_egresos_master.empty:
                    st.dataframe(df_egresos_master.style.format({'SubTotal': '${:,.2f}'}), use_container_width=True, hide_index=True)
                    st.metric("Total Egresos Master", f"${df_egresos_master['SubTotal'].sum():,.2f}")
                else:
                    st.warning("Sin registros de egresos en Master.")

        with tab_conc:
            st.markdown(f"### 🔍 Conciliación por UUID (Ingresos Vigentes) - Periodo: {nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} del {anio_sel}")
            
            if not df_ingresos_master.empty and not df_ingresos_sat.empty:
                df_m_agg = df_ingresos_master[['UUID', 'Empresa', 'SubTotal']].rename(
                    columns={'SubTotal': 'SubTotal_Master'}
                )
                df_s_agg = df_ingresos_sat[['UUID', 'Empresa', 'SubTotal']].rename(
                    columns={'SubTotal': 'SubTotal_SAT', 'Empresa': 'Empresa_SAT'}
                )
                
                df_concil = pd.merge(df_m_agg, df_s_agg, on='UUID', how='outer')
                df_concil['SubTotal_Master'] = df_concil['SubTotal_Master'].fillna(0.0)
                df_concil['SubTotal_SAT'] = df_concil['SubTotal_SAT'].fillna(0.0)
                df_concil['Empresa'] = df_concil['Empresa'].fillna(df_concil['Empresa_SAT']).fillna('OTRAS')
                
                df_concil['Diferencia'] = df_concil['SubTotal_Master'] - df_concil['SubTotal_SAT']
                
                df_con_dif = df_concil[df_concil['Diferencia'].round(2) != 0.0].copy()

                st.markdown("#### 📊 Resumen de UUIDs con Diferencias / Faltantes")
                col_d1, col_d2 = st.columns(2)
                col_d1.metric("Total UUIDs Analizados", len(df_concil))
                col_d2.metric("UUIDs con Diferencia / Faltantes", len(df_con_dif))

                st.markdown("---")
                st.markdown("#### 📋 Detalle Exclusivo de UUIDs con Diferencias o Faltantes en una Base")
                if not df_con_dif.empty:
                    st.dataframe(
                        df_con_dif[[
                            'UUID', 'Empresa', 'SubTotal_Master', 'SubTotal_SAT', 'Diferencia'
                        ]].style.format({
                            'SubTotal_Master': '${:,.2f}',
                            'SubTotal_SAT': '${:,.2f}',                             'Diferencia': '${:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.success("🎉 ¡Excelente! No se encontraron diferencias en los UUIDs vigentes cruzados.")
            else:
                st.info("Se requiere información tanto del Master como de los XMLs para ejecutar la conciliación.")

        with tab_cont:
            st.markdown(f"### 📊 Detalle de Cuentas Contables (Mes: {nombres_meses[mes_final-1]} {anio_sel})")
            st.markdown("Fórmula aplicada:\n* **Ingresos:** `410-00000-000-0000` (+) `411-00000-000-0000` (-) `423-00000-000-0000`\n* **Egresos:** `420-00000-000-0000` (+) `421-00000-000-0000`")
            
            if not df_detalle_cont.empty:
                st.dataframe(
                    df_detalle_cont.style.format({
                        'Saldo Deudor Final': '${:,.2f}',                         'Saldo Acreedor Final': '${:,.2f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("No se encontraron registros contables para las cuentas especificadas en el periodo seleccionado.")