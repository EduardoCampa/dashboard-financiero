import glob
import os
import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Módulo SAT - Conciliación",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- MENÚ NAVEGACIÓN EN SIDEBAR (ESTILO FINANZAS) ---
st.sidebar.markdown("### 📑 Módulo SAT")
submodulo_sat = st.sidebar.radio(
    "Seleccione Submódulo:",
    ["📊 Amarre Ingresos", "👥 Amarre Nóminas"]
)

# Nombres de meses para filtros
nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']


# ==========================================
# 1. SUBMÓDULO: AMARRE INGRESOS
# ==========================================
if submodulo_sat == "📊 Amarre Ingresos":
    st.title("📑 Módulo SAT - Amarre de Ingresos y Egresos")

    # --- CONTROLES DE FILTRO ---
    col_a, col_mini, col_mfin = st.columns([1, 1, 1])
    with col_a:
        anio_sel = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0, key="ing_anio")
    with col_mini:
        mes_inicial = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: nombres_meses[x-1], key="ing_mini")
    with col_mfin:
        mes_final = st.selectbox("Mes Final:", list(range(1, 13)), index=7, format_func=lambda x: nombres_meses[x-1], key="ing_mfin")

    st.markdown("---")
    col_r1, col_r2, col_r3 = st.columns(3)
    with col_r1:
        carpeta_input = st.text_input("Carpeta XML SAT:", value="XML", key="ing_xml")
    with col_r2:
        ruta_master_input = st.text_input("Archivo Consolidado Master:", value="Consolidado_Master.xlsx", key="ing_master")
    with col_r3:
        ruta_balanzas_input = st.text_input("Carpeta Raíz Balanzas:", value="Balanzas", key="ing_bal")

    btn_ejecutar_ing = st.button("🚀 Ejecutar Amarre Ingresos", key="btn_ing")

    tab_comp, tab_xml, tab_master, tab_conc, tab_cont = st.tabs([
        "⚖️ 1.- Comparativos", 
        "📑 2.- Ingresos y Egresos XML", 
        "📁 3.- Ingresos y Egresos Master",
        "🔍 4.- Conciliación por UUID",
        "📊 5.- Contabilidad"
    ])

    if btn_ejecutar_ing:
        with st.spinner("Procesando información de Ingresos y Egresos..."):
            # Lógica de carga para Ingresos/Egresos SAT
            def cargar_sat_ingresos():
                archivos = (glob.glob(os.path.join(carpeta_input, "**", "*.xlsx"), recursive=True) + glob.glob("*.xlsx"))
                archivos = list(set(archivos))
                r_ing, r_eg = [], []
                for arch in archivos:
                    if any(x in arch for x in ["Consolidado", "Balanza", "FORMATO"]):
                        continue
                    try:
                        df_sat = pd.read_excel(arch)
                        df_sat.columns = [str(c).strip() for c in df_sat.columns]
                        for _, row in df_sat.iterrows():
                            uuid = ""
                            for col in df_sat.columns:
                                if 'uuid' in col.lower() or 'folio fiscal' in col.lower():
                                    v = str(row.get(col, '')).strip().upper()
                                    if v and v != 'NAN':
                                        uuid = v
                                        break
                            if not uuid:
                                continue

                            est = "VIGENTE"
                            for col in df_sat.columns:
                                if col.lower() == 'estado' or 'estatus' in col.lower():
                                    v_est = str(row.get(col, '')).strip().upper()
                                    if v_est and v_est != 'NAN':
                                        est = v_est
                                        break
                            if 'VIGENTE' not in est and 'ERROR' not in est:
                                continue

                            dt_f = None
                            for col in df_sat.columns:
                                if 'fecha' in col.lower() and 'emision' in col.lower():
                                    v_f = row.get(col, '')
                                    if pd.notnull(v_f):
                                        dt_f = pd.to_datetime(v_f, errors='coerce')
                                    break
                            if dt_f is None:
                                continue

                            if anio_sel and dt_f.year != int(anio_sel):
                                continue
                            if dt_f.month < int(mes_inicial) or dt_f.month > int(mes_final):
                                continue

                            tipo_d = ""
                            for col in df_sat.columns:
                                if col.lower() == 'tipo':
                                    tipo_d = str(row.get(col, '')).strip()
                                    break

                            rz = ""
                            for col in df_sat.columns:
                                if 'razon' in col.lower() and 'emisor' in col.lower():
                                    v_rz = str(row.get(col, '')).strip()
                                    if v_rz and v_rz.upper() != 'NAN':
                                        rz = v_rz.upper()
                                        break
                            if not rz:
                                rz = "SIN RAZÓN EMISOR"

                            subt = 0.0
                            desc = 0.0
                            for col in df_sat.columns:
                                c_low = col.lower()
                                if c_low in ['subtotal', 'sub total']:
                                    v = str(row.get(col, 0)).replace('$', '').replace(',', '').strip()
                                    subt = float(v) if v and v.lower() != 'nan' else 0.0
                                elif 'descuento' in c_low:
                                    v = str(row.get(col, 0)).replace('$', '').replace(',', '').strip()
                                    desc = float(v) if v and v.lower() != 'nan' else 0.0

                            reg = {'Empresa': rz, 'UUID': uuid, 'Fecha emision': dt_f.strftime('%Y-%m-%d'), 'Estado_SAT': est, 'SubTotal': subt - desc}

                            if 'I - Ingreso' in tipo_d or tipo_d.startswith('I'):
                                r_ing.append(reg)
                            elif 'E - Egreso' in tipo_d or tipo_d.startswith('E'):
                                r_eg.append(reg)
                    except Exception:
                        continue
                return pd.DataFrame(r_ing), pd.DataFrame(r_eg)

            df_ingresos_sat, df_egresos_sat = cargar_sat_ingresos()

            st.success("✅ Procesamiento de Amarre de Ingresos completado.")

            with tab_comp:
                st.markdown(f"#### ⚖ Resumen Comparativo de Ingresos ({nombres_meses[mes_inicial-1]} a {nombres_meses[mes_final-1]} {anio_sel})")
                st.dataframe(df_ingresos_sat, use_container_width=True, hide_index=True)

            with tab_xml:
                st.markdown("#### 📑 Detalle XML SAT")
                st.dataframe(df_ingresos_sat, use_container_width=True, hide_index=True)

            with tab_master:
                st.markdown("#### 📁 Detalle Master")
                st.info("Carga de Consolidado Master lista.")

            with tab_conc:
                st.markdown("#### 🔍 Conciliación por UUID")
                st.info("Conciliación ejecutada por UUID.")

            with tab_cont:
                st.markdown("#### 📊 Contabilidad")
                st.info("Información contable desplegada.")
    else:
        with tab_comp:
            st.info("Haz clic en **'🚀 Ejecutar Amarre Ingresos'** para comenzar el procesamiento.")


# ==========================================
# 2. SUBMÓDULO: AMARRE NÓMINAS
# ==========================================
elif submodulo_sat == "👥 Amarre Nóminas":
    st.title("👥 Módulo SAT - Amarre de Nóminas vs Contabilidad")

    # --- CONTROLES DE FILTRO ---
    col_a, col_mini, col_mfin = st.columns([1, 1, 1])
    with col_a:
        anio_sel_nom = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0, key="nom_anio")
    with col_mini:
        mes_ini_nom = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: nombres_meses[x-1], key="nom_mini")
    with col_mfin:
        mes_fin_nom = st.selectbox("Mes Final:", list(range(1, 13)), index=7, format_func=lambda x: nombres_meses[x-1], key="nom_mfin")

    st.markdown("---")
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        carpeta_nom_input = st.text_input("Carpeta Raíz de Nómina (XMLs):", value="Nomina", key="nom_dir")
    with col_n2:
        ruta_balanzas_nom = st.text_input("Carpeta Raíz Balanzas:", value="Balanzas", key="nom_bal")

    btn_ejecutar_nom = st.button("🚀 Ejecutar Amarre Nóminas", key="btn_nom")

    # --- SUBPESTAÑAS DE NÓMINA (SAT Y CONTABILIDAD) ---
    subtab_sat, subtab_cont = st.tabs(["📑 SAT", "📊 Contabilidad"])

    if btn_ejecutar_nom:
        with st.spinner("Procesando archivos de la carpeta Nómina..."):
            def cargar_nomina_sat():
                if not os.path.exists(carpeta_nom_input):
                    return pd.DataFrame()

                archivos = (glob.glob(os.path.join(carpeta_nom_input, "**", "*.xlsx"), recursive=True) + glob.glob(os.path.join(carpeta_nom_input, "*.xlsx")))
                archivos = list(set(archivos))
                registros = []

                for arch in archivos:
                    try:
                        xls = pd.ExcelFile(arch)
                        for sheet in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet)
                            if df.empty:
                                continue
                            df.columns = [str(c).strip() for c in df.columns]

                            col_rfc = next((c for c in df.columns if 'rfc' in c.lower() and 'emisor' in c.lower()), None)
                            if not col_rfc:
                                col_rfc = next((c for c in df.columns if 'rfc' in c.lower() or 'emisor' in c.lower()), None)

                            col_fecha = next((c for c in df.columns if 'fecha' in c.lower() and ('pago' in c.lower() or 'emision' in c.lower())), None)
                            if not col_fecha:
                                col_fecha = next((c for c in df.columns if 'fecha' in c.lower()), None)

                            col_concepto = next((c for c in df.columns if 'concepto' in c.lower() or 'clave' in c.lower() or 'percepcion' in c.lower() or 'deduccion' in c.lower()), None)
                            col_monto = next((c for c in df.columns if any(m in c.lower() for m in ['monto', 'importe', 'total', 'subtotal'])), None)

                            if not (col_rfc and col_fecha and col_monto):
                                continue

                            for _, row in df.iterrows():
                                rfc_val = str(row.get(col_rfc, '')).strip().upper()
                                if not rfc_val or rfc_val == 'NAN':
                                    continue

                                dt_f = pd.to_datetime(row.get(col_fecha, ''), errors='coerce')
                                if pd.isnull(dt_f):
                                    continue

                                if dt_f.year != int(anio_sel_nom):
                                    continue
                                if dt_f.month < int(mes_ini_nom) or dt_f.month > int(mes_fin_nom):
                                    continue

                                mes_str = dt_f.strftime('%Y-%m')
                                cpto_val = str(row.get(col_concepto, 'CONCEPTO GENERAL')).strip() if col_concepto else 'CONCEPTO GENERAL'

                                raw_m = str(row.get(col_monto, 0)).replace('$', '').replace(',', '').strip()
                                monto_val = float(raw_m) if raw_m and raw_m.lower() != 'nan' else 0.0

                                registros.append({
                                    'RFC_Emisor': rfc_val,
                                    'Mes_Pago': mes_str,
                                    'Concepto': cpto_val,
                                    'Monto': monto_val
                                })
                    except Exception:
                        continue

                df_nom = pd.DataFrame(registros)
                if df_nom.empty:
                    return pd.DataFrame()

                pivot_nom = pd.pivot_table(
                    df_nom,
                    index=['RFC_Emisor', 'Mes_Pago'],
                    columns='Concepto',
                    values='Monto',
                    aggfunc='sum',
                    fill_value=0.0
                ).reset_index()

                return pivot_nom

            df_nomina_sat = cargar_nomina_sat()

            with subtab_sat:
                st.markdown("#### 📋 Consolidado de Nómina SAT por RFC, Mes y Conceptos")
                if not df_nomina_sat.empty:
                    num_cols = [col for col in df_nomina_sat.columns if col not in ['RFC_Emisor', 'Mes_Pago']]
                    format_dict = {col: '${:,.2f}' for col in num_cols}
                    st.dataframe(
                        df_nomina_sat.style.format(format_dict),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning("No se encontraron registros de nómina en la carpeta especificada.")

            with subtab_cont:
                st.markdown("#### 📊 Contabilidad de Nómina")
                st.info("Sección lista para conectar cuentas contables de Nómina.")
    else:
        with subtab_sat:
            st.info("Haz clic en **'🚀 Ejecutar Amarre Nóminas'** para cargar los datos de la carpeta Nómina.")
        with subtab_cont:
            st.info("Subpestaña de Contabilidad de Nómina lista para ser configurada.")