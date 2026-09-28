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
EXCEL_RESUMEN_PATH = "PARTES RELACIONADAS 08-2026 RESUMEN.xlsx"

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
    "0468": "CIVLAT",
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


@st.cache_data(ttl=600)
def cargar_balanzas_acumuladas():
    if not os.path.exists(EXCEL_RESUMEN_PATH):
        return pd.DataFrame()
    df = pd.read_excel(EXCEL_RESUMEN_PATH, sheet_name="BALANZAS ACUMULADAS")
    return df


@st.cache_data(ttl=600)
def obtener_lista_empresas_excel():
    df = cargar_balanzas_acumuladas()
    if df.empty:
        return []
    empresas = df['EMPRESA'].dropna().unique().tolist()
    return ordenar_empresas_segun_prioridad(empresas)


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
lista_empresas = obtener_lista_empresas_excel()

if lista_empresas:
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

    # NAVEGACIÓN PRINCIPAL DEL MÓDULO CONTABLE
    seccion_contable = st.radio(
        "📌 **Selecciona la Vista Contable:**",
        ["📊 Estados de Resultados", "🔗 Amarres Contables"],
        horizontal=True,
    )

    st.markdown("---")

    if seccion_contable == "📊 Estados de Resultados":
        st.info("Módulo de Estados de Resultados activo.")
    elif seccion_contable == "🔗 Amarres Contables":
        st.subheader("🔗 Módulo de Amarres Contables (Ejercicio Acumulado)")

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
            st.info("Cruce dinámico desde la pestaña `BALANZAS ACUMULADAS` del archivo Excel adjunto.")

            df_balanzas = cargar_balanzas_acumuladas()

            if not df_balanzas.empty:
                empresas_disp = obtener_lista_empresas_excel()

                # --- TABLA 1: FACTURACIÓN ---
                st.markdown("---")
                st.markdown("#### 🟢 1. Amarre de Facturación (Receptora VS Origen | Clientes 101-00004 & Proveedores 201-00002, 201-00004)")

                matriz_fact = pd.DataFrame(
                    0.0, index=empresas_disp, columns=empresas_disp
                )

                # Iterar sobre las filas de balanzas acumuladas
                for _, row in df_balanzas.iterrows():
                    emp_receptora = str(row.get('EMPRESA', '')).strip()
                    cuenta = str(row.iloc[1]).strip()
                    saldo_final = parse_monto_robusto(row.get('Saldo final', 0.0))

                    if emp_receptora in empresas_disp:
                        if cuenta.startswith("101-00004-") or cuenta.startswith("201-00002-") or cuenta.startswith("201-00004-"):
                            segs = cuenta.split('-')
                            if len(segs) >= 4:
                                cod_contraparte = segs[3].strip()
                                emp_origen = MAPEO_ID_EMPRESA.get(
                                    cod_contraparte, None
                                )

                                if not emp_origen:
                                    for ed in empresas_disp:
                                        if (
                                            cod_contraparte in ed
                                            or cod_contraparte.lstrip('0')
                                            in ed.lstrip('0')
                                        ):
                                            emp_origen = ed
                                            break

                                if emp_origen and emp_origen in empresas_disp:
                                    if cuenta.startswith("201-"):
                                        saldo_final = -saldo_final
                                    matriz_fact.loc[
                                        emp_receptora, emp_origen
                                    ] += saldo_final

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

                for _, row in df_balanzas.iterrows():
                    emp_receptora = str(row.get('EMPRESA', '')).strip()
                    cuenta = str(row.iloc[1]).strip()
                    saldo_final = parse_monto_robusto(row.get('Saldo final', 0.0))

                    if emp_receptora in empresas_disp:
                        if cuenta.startswith("101-00015-") or cuenta.startswith("202-00001-"):
                            segs = cuenta.split('-')
                            if len(segs) >= 4:
                                cod_contraparte = segs[3].strip()
                                emp_origen = MAPEO_ID_EMPRESA.get(
                                    cod_contraparte, None
                                )

                                if not emp_origen:
                                    for ed in empresas_disp:
                                        if (
                                            cod_contraparte in ed
                                            or cod_contraparte.lstrip('0')
                                            in ed.lstrip('0')
                                        ):
                                            emp_origen = ed
                                            break

                                if emp_origen and emp_origen in empresas_disp:
                                    if cuenta.startswith("202-"):
                                        saldo_final = -saldo_final
                                    matriz_prest.loc[
                                        emp_receptora, emp_origen
                                    ] += saldo_final

                matriz_prest['TOTAL'] = matriz_prest.sum(axis=1)
                st.dataframe(
                    matriz_prest.style.format("${:,.2f}"),
                    use_container_width=True,
                )

                st.success("✅ Matrices de Facturación y Préstamos cargadas exitosamente desde `BALANZAS ACUMULADAS`.")

        with subtab_ig_intercos:
            st.markdown("### 📑 Amarre Ingresos y Gastos Intercompañías")
            st.info("Conciliación cruzada de Ingresos Intercos vs Gastos y Costos Intercos.")
else:
    st.info("No se encontraron empresas en el archivo de balanzas acumuladas.")