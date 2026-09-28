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

# ORDEN PREFERENTE DE EMPRESAS PARA EL MULTISELECT
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

# MAPEO OFICIAL DE CÓDIGOS DE CONTPAQI (Último segmento de la cuenta) A NOMBRE DE EMPRESA
MAPEO_CODIGO_EMPRESA = {
    "0467": "CIV",
    "0468": "CIVLAT",
    "0469": "CIVMEX",
    "0665": "COMANA",
    "0813": "EFCO",
    "0942": "FGS",
    "1127": "FERVIC",
    "1144": "PESAZA",
    "1149": "GPO SERVYRE",
    "1216": "FPSB",
    "1404": "INMOBILIARIA",
    "1428": "IPV",
    "1603": "LABORATORIOS",
    "1616": "LAITS",
    "1626": "LIMPIESPIN",
    "1636": "LOGISTICA",
    "2396": "PROINA",
    "2871": "SERVYCARGO",
    "2872": "SERVYRE",
    "2937": "SEVILAT",
    "3329": "VIALTECNO",
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
def cargar_hoja_balanza(ruta, nombre_hoja):
    return pd.read_excel(ruta, sheet_name=nombre_hoja)


@st.cache_data(ttl=600)
def extraer_registros_balanza(ruta, nombre_hoja):
    df_b = cargar_hoja_balanza(ruta, nombre_hoja)
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
            'deudor_f': deudor_f,
            'acreedor_f': acreedor_f,
        })
    return balanza_records


@st.cache_data(ttl=600)
def obtener_lista_obras(ruta, empresas_tuple):
    obras = set()
    for emp in empresas_tuple:
        records = extraer_registros_balanza(ruta, emp)
        for r in records:
            segs = r['cta_raw'].split('-')
            if len(segs) >= 3:
                seg2 = segs[1].strip()
                if seg2 != '00000' and seg2.isdigit() and len(seg2) >= 3:
                    obras.add(seg2)
    return sorted(list(obras))


# --- PLANTILLA DE EXCEL CON AGRUPACIONES ---
@st.cache_data(ttl=3600)
def cargar_plantilla_formato():
    ruta_formato = "FORMATO EDO RESULTADOS.xlsx"
    if not os.path.exists(ruta_formato):
        return []

    wb = openpyxl.load_workbook(ruta_formato, data_only=False)
    sheet = wb['RESULTADOS ACUM'] if 'RESULTADOS ACUM' in wb.sheetnames else wb.active

    plantilla = []
    for i in range(6, sheet.max_row + 1):
        cta = sheet.cell(row=i, column=1).value
        concepto = sheet.cell(row=i, column=2).value
        c_formula = sheet.cell(row=i, column=3).value
        d_formula = sheet.cell(row=i, column=4).value

        str_cta = str(cta).strip() if cta else ""
        str_conc = str(concepto).strip() if concepto else ""

        if str_cta.upper() in ['CUENTA', '1', '2', '3'] or str_conc.upper() in [
            'CONCEPTO',
            '1',
            '2',
            'NONE',
            'NAN',
        ]:
            continue

        outline_lvl = sheet.row_dimensions[i].outlineLevel or 0

        if cta or concepto or c_formula:
            plantilla.append({
                'row_idx': i,
                'cuenta_patron': str_cta if str_cta else None,
                'concepto': str_conc,
                'c_formula': str(c_formula).strip() if c_formula else None,
                'd_formula': str(d_formula).strip() if d_formula else None,
                'outline_level': outline_lvl,
            })
    return plantilla


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


# --- EVALUADOR DE FÓRMULAS EN CASCADA ---
def resolver_todas_las_formulas(mapa_valores, mapa_formulas):
    vals = dict(mapa_valores)

    for _ in range(5):
        for r_idx, f_str in mapa_formulas.items():
            f_clean = f_str.upper().replace(' ', '').replace('$', '')

            m_sub = re.match(r'^=SUBTOTAL\(9,C(\d+):C(\d+)\)$', f_clean)
            if m_sub:
                r_s, r_e = int(m_sub.group(1)), int(m_sub.group(2))
                vals[r_idx] = sum(vals.get(r, 0.0) for r in range(r_s, r_e + 1))
                continue

            expr_raw = re.sub(r'^=\+?', '', f_clean)

            def sustituir_celda(match):
                r_num = int(match.group(1))
                return str(vals.get(r_num, 0.0))

            expr = re.sub(r'C(\d+)', sustituir_celda, expr_raw)

            try:
                if re.match(r'^[0-9\.\+\-\*\/\(\)\s]+$', expr):
                    vals[r_idx] = float(eval(expr))
            except Exception:
                pass
    return vals


def coincide_cuenta_robusta(cta_balanza, patron_template):
    if not patron_template or patron_template in ('None', 'CUENTA', ''):
        return False

    cb = str(cta_balanza).strip()
    pt = str(patron_template).strip()

    cb_clean = re.sub(r'[^0-9A-Za-z]', '', cb)
    pt_clean = re.sub(r'[^0-9A-Za-z?]', '', pt)

    if cb_clean == pt_clean:
        return True

    p_segs = pt.split('-')
    b_segs = cb.split('-')

    if len(p_segs) >= 3 and len(b_segs) >= 3:
        if p_segs[0] == b_segs[0]:
            p_seg2_is_wildcard = ('?' in p_segs[1]) or (
                p_segs[1] in ('00000', '0000', '000', '99999')
            )

            try:
                seg3_match = (p_segs[2] == b_segs[2]) or (
                    int(p_segs[2]) == int(b_segs[2])
                )
            except ValueError:
                seg3_match = p_segs[2] == b_segs[2]

            if p_seg2_is_wildcard and seg3_match:
                if len(p_segs) >= 4 and len(b_segs) >= 4:
                    try:
                        return (p_segs[3] == b_segs[3]) or (
                            int(p_segs[3]) == int(b_segs[3])
                        )
                    except ValueError:
                        return p_segs[3] == b_segs[3]
                return True

            try:
                if int(p_segs[1]) == int(b_segs[1]) and seg3_match:
                    if len(p_segs) >= 4 and len(b_segs) >= 4:
                        return int(p_segs[3]) == int(b_segs[3])
                    return True
            except ValueError:
                pass

    regex_str = "^" + pt.replace('?', '.').replace('-', r'\\-?') + "$"
    return bool(re.match(regex_str, cb))


def obtener_monto_cuenta_balanza(
    patron_template,
    balanza_records,
    tipo='mes',
    solo_intercos=False,
    obras_filtro=None,
):
    pt = str(patron_template).strip()
    p_prefix = pt.split('-')[0].strip() if '-' in pt else pt[:3]
    pt_clean = re.sub(r'[^0-9A-Za-z]', '', pt)

    exact_match = None
    for b in balanza_records:
        cb_clean = re.sub(r'[^0-9A-Za-z]', '', b['cta_raw'])
        if cb_clean == pt_clean:
            exact_match = b
            break

    records_a_sumar = []
    if exact_match:
        records_a_sumar = [exact_match]
    else:
        for b in balanza_records:
            if coincide_cuenta_robusta(b['cta_raw'], pt):
                records_a_sumar.append(b)

    if obras_filtro:
        obras_set = set(obras_filtro)
        filtrados_obra = []
        for b in records_a_sumar:
            segs_cb = str(b['cta_raw']).strip().split('-')
            if len(segs_cb) >= 2 and segs_cb[1] in obras_set:
                filtrados_obra.append(b)
        records_a_sumar = filtrados_obra

    monto = 0.0
    for b in records_a_sumar:
        if tipo == 'mes':
            monto += b.get('cargos_m', 0.0) - b.get('abonos_m', 0.0)
        else:
            monto += b['saldo_neto']

    return monto


def calcular_mapa_valores_empresa(
    balanza_records, plantilla, tipo='mes', solo_intercos=False, obras_filtro=None
):
    if not balanza_records or not plantilla:
        return {}

    val_map = {}
    formulas_map = {}

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        c_form = row['c_formula']

        if patron:
            val_map[r_idx] = obtener_monto_cuenta_balanza(
                patron,
                balanza_records,
                tipo=tipo,
                solo_intercos=solo_intercos,
                obras_filtro=obras_filtro,
            )
        elif c_form and str(c_form).startswith('='):
            formulas_map[r_idx] = c_form
            val_map[r_idx] = 0.0
        else:
            val_map[r_idx] = 0.0

    return resolver_todas_las_formulas(val_map, formulas_map)


@st.cache_data(ttl=600)
def generar_reporte_multiempresa(
    ruta_balanza,
    empresas_tuple,
    plantilla,
    tipo='mes',
    solo_intercos=False,
    obras_tuple=None,
):
    if not empresas_tuple or not plantilla:
        return pd.DataFrame()

    empresas_list = list(empresas_tuple)
    obras_a_excluir = list(obras_tuple) if obras_tuple else None

    mapas_empresas = {}
    mapas_obras_excluidas = {}

    for emp in empresas_list:
        records_b = extraer_registros_balanza(ruta_balanza, emp)
        mapas_empresas[emp] = calcular_mapa_valores_empresa(
            records_b, plantilla, tipo=tipo, solo_intercos=solo_intercos
        )

        if obras_a_excluir:
            mapas_obras_excluidas[emp] = calcular_mapa_valores_empresa(
                records_b,
                plantilla,
                tipo=tipo,
                solo_intercos=solo_intercos,
                obras_filtro=obras_a_excluir,
            )

    reporte = []
    current_group_id = 0

    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        concepto = row['concepto']
        c_form = row['c_formula']
        outline_lvl = row.get('outline_level', 0)

        if not patron and not concepto and not c_form:
            continue

        es_formula = bool(c_form and str(c_form).startswith('=') and not patron)
        es_cuenta = bool(patron)
        es_dato = es_cuenta or es_formula

        if not es_cuenta:
            current_group_id += 1

        fila_dict = {
            'CUENTA': patron if patron else '',
            'CONCEPTO': concepto,
            'row_idx': r_idx,
            'group_id': f"grp_{current_group_id}",
            'is_detail': es_cuenta,
            'outline_level': outline_lvl,
        }

        monto_total_consolidado = 0.0

        for emp in empresas_list:
            m_emp = mapas_empresas[emp].get(r_idx, 0.0)
            ventas_emp = mapas_empresas[emp].get(91, 0.0) or 1.0
            pct_emp = (m_emp / ventas_emp) * 100 if ventas_emp else 0.0

            fila_dict[f"{emp}"] = m_emp if es_dato else None
            fila_dict[f"% {emp}"] = pct_emp if es_dato else None

            if es_dato:
                monto_total_consolidado += m_emp

        tot_ventas_todas = sum(
            mapas_empresas[e].get(91, 0.0) for e in empresas_list
        ) or 1.0
        pct_total = (
            (monto_total_consolidado / tot_ventas_todas) * 100
            if tot_ventas_todas
            else 0.0
        )

        fila_dict['TOTAL CONSOLIDADO'] = (
            monto_total_consolidado if es_dato else None
        )
        fila_dict['% TOTAL'] = pct_total if es_dato else None
        fila_dict['es_total'] = es_formula or (not patron and bool(concepto))

        reporte.append(fila_dict)

    return pd.DataFrame(reporte)


def renderizar_tabla_interactiva_agrupada(
    df_er, empresas, conceptos_a_mostrar=None
):
    if df_er.empty:
        return

    df_disp = df_er.copy()
    if conceptos_a_mostrar and "TODOS" not in conceptos_a_mostrar:
        df_disp = df_disp[df_disp['CONCEPTO'].isin(conceptos_a_mostrar)]

    if df_disp.empty:
        st.warning("No hay rubros seleccionados para mostrar.")
        return

    cols_empresas = list(empresas)
    cols_header = ['CUENTA', 'CONCEPTO']
    for e in cols_empresas:
        cols_header.extend([e, f"% {e}"])

    html_code = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: transparent; }
        .tree-table { width: 100%; border-collapse: collapse; font-size: 13px; background: #fff; border: 1px solid #e2e8f0; border-radius: 8px; }
        .tree-table th { background: #1e293b; color: #fff; padding: 10px; text-align: center; font-size: 12px; }
        .tree-table td { padding: 8px 12px; border-bottom: 1px solid #e2e8f0; color: #334155; }
        .group-header { background: #f8fafc; font-weight: 700; color: #0284c7; }
        .grand-total { background: #e2e8f0; font-weight: 800; color: #0f172a; }
        .num-cell { text-align: right; font-variant-numeric: tabular-nums; }
    </style>
    </head>
    <body>
        <table class="tree-table">
            <thead><tr>
    """
    for h in cols_header:
        html_code += f"<th>{h}</th>"
    html_code += "</tr></thead><tbody>"

    for _, row in df_disp.iterrows():
        es_tot = row.get('es_total', False)
        row_class = "grand-total" if es_tot else "detail-row"
        html_code += f"<tr class='{row_class}'>"
        html_code += f"<td>{row.get('CUENTA', '')}</td><td>{row.get('CONCEPTO', '')}</td>"
        for e in cols_empresas:
            val_m = row.get(e, None)
            m_str = f"${val_m:,.2f}" if pd.notnull(val_m) and str(val_m) != 'None' else ""
            html_code += f"<td class='num-cell'>{m_str}</td><td></td>"
        html_code += "</tr>"

    html_code += "</tbody></table></body></html>"
    components.html(html_code, height=400, scrolling=True)


# --- INTERFAZ PRINCIPAL DE STREAMLIT ---
estructura = obtener_estructura_balanzas()
plantilla = cargar_plantilla_formato()

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
        seleccionar_todas_emp = st.checkbox("☑️ Seleccionar Todas las Empresas", value=True)

    with col_emp2:
        default_empresas = lista_empresas if seleccionar_todas_emp else ([lista_empresas[0]] if lista_empresas else [])
        empresas_seleccionadas = st.multiselect("Empresa(s):", lista_empresas, default=default_empresas)

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
            "🔄 Amarre Intercompañías",
            "📑 Amarre I y G Intercos",
        ])

        with subtab_intercos:
            st.markdown("### 🔄 Amarre Intercompañías (Cuentas de Balance: Clientes, Proveedores y Préstamos)")
            st.info(f"Leyendo saldos de balance directamente de la balanza mensual: `{ruta_balanza}`.")

            todas_empresas_balanza = obtener_lista_empresas(ruta_balanza)

            if todas_empresas_balanza:
                datos_empresas_recs = {}
                for emp in todas_empresas_balanza:
                    datos_empresas_recs[emp] = extraer_registros_balanza(ruta_balanza, emp)

                def extraer_codigo_contraparte(cta_str):
                    segs = cta_str.split('-')
                    if len(segs) >= 4:
                        return segs[3].strip()
                    elif len(segs) == 3:
                        return segs[2].strip()
                    return ""

                # --- TABLA 1: FACTURACIÓN (Clientes 101-00004 & Proveedores/Acreedores 201-00002, 201-00004) ---
                st.markdown("---")
                st.markdown("#### 🟢 1. Amarre de Facturación (Receptora VS Origen | Clientes 101-00004 & Proveedores 201-00002, 201-00004)")

                matriz_fact = pd.DataFrame(0.0, index=todas_empresas_balanza, columns=todas_empresas_balanza)

                for emp_receptora in todas_empresas_balanza:
                    recs = datos_empresas_recs[emp_receptora]
                    for r in recs:
                        cta = r['cta_raw']
                        # Excluir cuentas de resultados
                        if cta.startswith(('4', '5', '6', '7')):
                            continue

                        if cta.startswith(("101-00004-", "201-00002-", "201-00004-")):
                            cod_contraparte = extraer_codigo_contraparte(cta)
                            emp_origen = MAPEO_CODIGO_EMPRESA.get(cod_contraparte, None)

                            if emp_origen and emp_origen in todas_empresas_balanza:
                                saldo = r['saldo_neto']
                                if cta.startswith("201-"):
                                    saldo = -saldo
                                matriz_fact.loc[emp_receptora, emp_origen] += saldo

                matriz_fact['TOTAL'] = matriz_fact.sum(axis=1)
                st.dataframe(matriz_fact.style.format("${:,.2f}"), use_container_width=True)

                # --- TABLA 2: PRÉSTAMOS (Deudores 101-00015 & Pasivo LP 202-00001) ---
                st.markdown("---")
                st.markdown("#### 🟡 2. Amarre de Préstamos (Receptora VS Origen | Deudores 101-00015 & Pasivo LP 202-00001)")

                matriz_prest = pd.DataFrame(0.0, index=todas_empresas_balanza, columns=todas_empresas_balanza)

                for emp_receptora in todas_empresas_balanza:
                    recs = datos_empresas_recs[emp_receptora]
                    for r in recs:
                        cta = r['cta_raw']
                        if cta.startswith(('4', '5', '6', '7')):
                            continue

                        if cta.startswith(("101-00015-", "202-00001-")):
                            cod_contraparte = extraer_codigo_contraparte(cta)
                            emp_origen = MAPEO_CODIGO_EMPRESA.get(cod_contraparte, None)

                            if emp_origen and emp_origen in todas_empresas_balanza:
                                saldo = r['saldo_neto']
                                if cta.startswith("202-"):
                                    saldo = -saldo
                                matriz_prest.loc[emp_receptora, emp_origen] += saldo

                matriz_prest['TOTAL'] = matriz_prest.sum(axis=1)
                st.dataframe(matriz_prest.style.format("${:,.2f}"), use_container_width=True)

                st.success("✅ Matrices de Facturación y Préstamos calculadas estrictamente con cuentas de Balance y códigos de contraparte.")
else:
    st.info("Por favor selecciona al menos una empresa para mostrar el reporte.")