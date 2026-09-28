import glob
import io
import json
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Módulo Contable",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 Módulo Contable")

PRESETS_FILE = "vistas_personalizadas_er.json"

# ORDEN PREFERENTE DE EMPRESAS PARA LAS MATRICES
ORDEN_EMPRESAS_PRIORIDAD = [
    "CIVLAT",
    "SERVYRE",
    "FERVIC",
    "CIV",
    "GPO SERVYRE",
    "FGS",
    "LATIN",
    "EFCO",
    "PESAZA",
    "LIMPIESPIN",
    "SERVYCARGO",
    "CIVMEX",
    "COMANA",
    "PROINA",
    "SEVILAT",
    "IPV",
    "SERSEÑAL",
    "INMOBILIARIA",
    "VIALTECNO",
    "LABORATORIO",
    "FPSB",
    "SIGNAL",
]

# MAPEO DE CÓDIGOS DE CONTPAQI (Último segmento) A NOMBRE DE EMPRESA
MAPEO_CODIGO_EMPRESA = {
    "0468": "CIVLAT", "2": "CIVLAT", "002": "CIVLAT",
    "2872": "SERVYRE",
    "1626": "LIMPIESPIN",
    "1616": "LAITS",
    "1144": "PESAZA",
    "0813": "EFCO",
    "1428": "IPV",
    "1404": "INMOBILIARIA",
    "1216": "FPSB",
    "1603": "LABORATORIOS",
    "1127": "FERVIC",
    "0469": "CIVMEX",
    "0942": "FGS",
    "2937": "SEVILAT",
    "0467": "CIV",
    "2871": "SERVYCARGO",
    "3329": "VIALTECNO",
    "0665": "COMANA",
    "1149": "GPO SERVYRE",
    "2396": "PROINA",
}


def ordenar_empresas_segun_prioridad(lista_empresas):
    def obtener_posicion(emp_nombre):
        emp_clean = str(emp_nombre).strip().upper()
        for idx, pref in enumerate(ORDEN_EMPRESAS_PRIORIDAD):
            if pref == emp_clean or pref in emp_clean:
                return idx
        return 999

    return sorted(lista_empresas, key=obtener_posicion)


# --- GESTIÓN DE VISTAS GUARDADAS (PRESETS) ---
def cargar_vistas_guardadas():
    vistas_predeterminados = {
        "📊 Resumen Ejecutivo": [
            "Ventas Netas Totales",
            "Total Costo",
            "Total Resultado Bruto",
            "Total Gastos",
            "Resultado Antes de Depreciacion",
            "Total Depreciaciones Y Amortización",
            "Total Costo Integral de Financiamiento",
            "Utilidad/Perdida antes de Impuestos",
        ]
    }

    if os.path.exists(PRESETS_FILE):
        try:
            with open(PRESETS_FILE, "r", encoding="utf-8") as f:
                guardadas = json.load(f)
                vistas_predeterminados.update(guardadas)
        except Exception:
            pass

    return vistas_predeterminados


def guardar_nueva_vista(nombre_vista, lista_conceptos):
    vistas = cargar_vistas_guardadas()
    vistas[f"⭐ {nombre_vista}"] = lista_conceptos

    custom_vistas = {k: v for k, v in vistas.items() if k.startswith("⭐ ")}
    with open(PRESETS_FILE, "w", encoding="utf-8") as f:
        json.dump(custom_vistas, f, ensure_ascii=False, indent=2)


# --- BUSCAR BALANZAS POR AÑO Y MES ---
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
def obtener_lista_empresas(ruta):
    xls = pd.ExcelFile(ruta)
    sheets = xls.sheet_names
    if len(sheets) > 1 and "Hoja1" in sheets:
        sheets.remove("Hoja1")

    return ordenar_empresas_segun_prioridad(sheets)


@st.cache_data(ttl=600)
def extraer_registros_balanza_pestana(ruta, nombre_hoja):
    df_b = pd.read_excel(ruta, sheet_name=nombre_hoja)
    if df_b.empty:
        return []

    num_cols = df_b.shape[1]
    col_cta = 0
    col_deudor_f = num_cols - 2
    col_acreedor_f = num_cols - 1

    balanza_records = []
    for _, row in df_b.iterrows():
        cta_raw = str(row.iloc[col_cta]).strip()
        if not cta_raw or cta_raw.lower() in ('nan', 'cuenta', 'none'):
            continue

        if not re.search(r'\d{3}[-\s]?\d{3,5}', cta_raw):
            continue

        deudor_f = parse_monto_robusto(row.iloc[col_deudor_f])
        acreedor_f = parse_monto_robusto(row.iloc[col_acreedor_f])
        saldo_neto = deudor_f - acreedor_f

        balanza_records.append({
            'cta_raw': cta_raw,
            'saldo_neto': saldo_neto,
        })
    return balanza_records


# --- PARSER NUMÉRICO ROBUSTO ---
def parse_monto_robusto(val):
    if pd.isnull(val):
        return 0.0
    val_str = (
        str(val).replace('$', '').replace(',', '').replace(' ', '').strip()
    )
    if not val_str or val_str.lower() in ('nan', 'none', '-'):
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0


# --- INTERFAZ PRINCIPAL DE STREAMLIT ---
estructura = obtener_estructura_balanzas()

if estructura:
    anios_disponibles = sorted(list(set(x['anio'] for x in estructura)), reverse=True)

    col_a, col_m = st.columns([1, 1])

    with col_a:
        anio_sel = st.selectbox("Año:", anios_disponibles)

    meses_disponibles = sorted(
        list(set(x['mes'] for x in estructura if x['anio'] == anio_sel)),
        reverse=True,
    )

    with col_m:
        mes_sel = st.selectbox("Mes:", meses_disponibles)

    ruta_balanza = next(
        (x['ruta'] for x in estructura if x['anio'] == anio_sel and x['mes'] == mes_sel),
        estructura[0]['ruta'],
    )

    lista_empresas = obtener_lista_empresas(ruta_balanza)

    col_emp1, col_emp2 = st.columns([1, 3])
    with col_emp1:
        st.write("")
        st.write("")
        seleccionar_todas_emp = st.checkbox(
            "☑️ Seleccionar Todas las Empresas", value=True
        )

    with col_emp2:
        default_empresas = (
            lista_empresas
            if seleccionar_todas_emp
            else ([lista_empresas[0]] if lista_empresas else [])
        )

        empresas_seleccionadas = st.multiselect(
            "Empresa(s):",
            lista_empresas,
            default=default_empresas,
        )

    st.markdown("---")

    seccion_contable = st.radio(
        "📌 **Selecciona la Vista Contable:**",
        ["📊 Estados de Resultados", "🔗 Amarres Contables"],
        horizontal=True,
    )

    st.markdown("---")

    if seccion_contable == "📊 Estados de Resultados":
        st.info("Módulo de Estados de Resultados activo.")
    elif seccion_contable == "🔗 Amarres Contables":
        st.subheader(f"🔗 Módulo de Amarres Contables ({mes_sel}/{anio_sel})")

        subtab_rh, subtab_ingresos, subtab_intercos, subtab_ig_intercos = st.tabs([
            "👥 Amarre RH",
            "💰 Amarre Ingresos",
            "🔄 Amarre Intercos",
            "📑 Amarre I y G Intercos",
        ])

        with subtab_rh:
            st.markdown("### 👥 Amarre Recursos Humanos")
            st.info("Conciliación de Sueldos, Salarios y Provisiones de Nómina.")

        with subtab_ingresos:
            st.markdown("### 💰 Amarre de Ingresos y Facturación")
            st.info("Cruce de XMLs vs Ingresos Registrados en Cuentas 410 / 411.")

        with subtab_intercos:
            st.markdown("### 🔄 Amarre Intercompañías (Cuentas por Cobrar / Pagar y Préstamos)")
            st.info(f"Cruce dinámico leyendo directamente cada pestaña de la balanza mensual: `{ruta_balanza}`.")

            empresas_disp = obtener_lista_empresas(ruta_balanza)

            if empresas_disp:
                # --- TABLA 1: FACTURACIÓN ---
                st.markdown("---")
                st.markdown("#### 🟢 1. Amarre de Facturación (Receptora VS Origen | Clientes 101-00004 & Proveedores 201-00002, 201-00004)")

                matriz_fact = pd.DataFrame(
                    0.0, index=empresas_disp, columns=empresas_disp
                )

                for emp_receptora in empresas_disp:
                    recs = extraer_registros_balanza_pestana(ruta_balanza, emp_receptora)
                    for r in recs:
                        cta = r['cta_raw']
                        saldo = r['saldo_neto']

                        if cta.startswith("101-00004-") or cta.startswith("201-00002-") or cta.startswith("201-00004-"):
                            segs = cta.split('-')
                            if len(segs) >= 3:
                                # Identificar empresa origen (tercer segmento o sufijo)
                                sub_cta = segs[2] if len(segs) >= 3 else segs[1]
                                emp_origen = MAPEO_ID_EMPRESA.get(sub_cta, None) if 'MAPEO_ID_EMPRESA' in globals() else None
                                
                                if not emp_origen:
                                    for k, v in MAPEO_CODIGO_EMPRESA.items():
                                        if k in cta:
                                            emp_origen = v
                                            break

                                if not emp_origen:
                                    for ed in empresas_disp:
                                        if sub_cta in ed or sub_cta.lstrip('0') in ed.lstrip('0'):
                                            emp_origen = ed
                                            break

                                if emp_origen and emp_origen in empresas_disp:
                                    if cta.startswith("201-"):
                                        saldo = -saldo
                                    matriz_fact.loc[emp_receptora, emp_origen] += saldo

                matriz_fact['TOTAL'] = matriz_fact.sum(axis=1)
                st.dataframe(
                    matriz_fact.style.format("${:,.2f}"),
                    use_container_width=True,
                )

                # --- TABLA 2: PRÉSTAMOS ---
                st.markdown("---")
                st.markdown("#### 🟡 2. Amarre de Préstamos (Receptora VS Origen | Deudores 101-00015 & Pasivo LP 202-00001)")

                matriz_prest = pd.DataFrame(
                    0.0, index=empresas_disp, columns=empresas_disp
                )

                for emp_receptora in empresas_disp:
                    recs = extraer_registros_balanza_pestana(ruta_balanza, emp_receptora)
                    for r in recs:
                        cta = r['cta_raw']
                        saldo = r['saldo_neto']

                        if cta.startswith("101-00015-") or cta.startswith("202-00001-"):
                            segs = cta.split('-')
                            if len(segs) >= 3:
                                sub_cta = segs[2] if len(segs) >= 3 else segs[1]
                                emp_origen = None
                                for k, v in MAPEO_CODIGO_EMPRESA.items():
                                    if k in cta:
                                        emp_origen = v
                                        break

                                if not emp_origen:
                                    for ed in empresas_disp:
                                        if sub_cta in ed or sub_cta.lstrip('0') in ed.lstrip('0'):
                                            emp_origen = ed
                                            break

                                if emp_origen and emp_origen in empresas_disp:
                                    if cta.startswith("202-"):
                                        saldo = -saldo
                                    matriz_prest.loc[emp_receptora, emp_origen] += saldo

                matriz_prest['TOTAL'] = matriz_prest.sum(axis=1)
                st.dataframe(
                    matriz_prest.style.format("${:,.2f}"),
                    use_container_width=True,
                )

                st.success("✅ Matrices de Facturación y Préstamos generadas correctamente leyendo pestaña por pestaña de la balanza mensual.")

        with subtab_ig_intercos:
            st.markdown("### 📑 Amarre Ingresos y Gastos Intercompañías")
            st.info("Conciliación cruzada de Ingresos Intercos vs Gastos y Costos Intercos.")
else:
    st.info("No se encontraron carpetas de Balanzas en la ruta especificada.")