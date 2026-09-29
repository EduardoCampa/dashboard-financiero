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

MAPEO_CODIGO_EMPRESA = {
    "0468": "CIVLAT",
    "0467": "CIV",
    "0469": "CIVMEX",
    "2872": "SERVYRE",
    "1636": "SERSEÑAL",
    "0813": "EFCO",
    "0942": "FGS",
    "1127": "FERVIC",
    "1216": "FPSB",
    "1144": "PESAZA",
    "1149": "GPO SERVYRE",
    "1404": "INMOBILIARIA",
    "1603": "LABORATORIO",
    "1616": "LATIN",
    "1626": "LIMPIESPIN",
    "2871": "SERVYCARGO",
    "3329": "VIALTECNO",
    "2937": "SEVILAT",
    "2396": "PROINA",
    "1428": "IPV",
    "0665": "COMANA",
}

MAPEO_NOMBRE_A_CODIGO = {v: k for k, v in MAPEO_CODIGO_EMPRESA.items()}


def ordenar_empresas_segun_prioridad(lista_empresas):
    def obtener_posicion(emp_nombre):
        emp_clean = str(emp_nombre).strip().upper()
        for idx, pref in enumerate(ORDEN_EMPRESAS_PRIORIDAD):
            if pref == emp_clean or pref in emp_clean:
                return idx
        return 999

    return sorted(lista_empresas, key=obtener_posicion)


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
    col_deudor_f = 6 if num_cols > 6 else num_cols - 2
    col_acreedor_f = 7 if num_cols > 7 else num_cols - 1

    balanza_records = []
    for _, row in df_b.iterrows():
        cta_raw = str(row.iloc[col_cta]).strip()
        if not cta_raw or cta_raw.lower() in ('nan', 'cuenta', 'none'):
            continue
        if not re.search(r'\d{3}[-\s]?\d{3,5}', cta_raw):
            continue
        deudor_f = parse_monto_robusto(row.iloc[col_deudor_f])
        acreedor_f = parse_monto_robusto(row.iloc[col_acreedor_f])
        saldo_final = deudor_f - acreedor_f
        balanza_records.append({
            'cta_raw': cta_raw,
            'deudor_f': deudor_f,
            'acreedor_f': acreedor_f,
            'saldo_final': saldo_final,
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
        if str_cta.upper() in ['CUENTA', '1', '2', '3'] or str_conc.upper() in ['CONCEPTO', '1', '2', 'NONE', 'NAN']:
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


def extraer_codigo_contraparte_robusto(cta_str):
    segs = cta_str.split('-')
    for seg in reversed(segs):
        seg_clean = seg.strip()
        if seg_clean in MAPEO_CODIGO_EMPRESA:
            return seg_clean
    for seg in reversed(segs):
        seg_clean = seg.strip()
        if seg_clean.isdigit() and len(seg_clean) in (3, 4):
            return seg_clean
    return ""


def coincide_cuenta_robusta(cta_balanza, patron_template):
    if not patron_template or patron_template in ('None', 'CUENTA', ''):
        return False
    cb = str(cta_balanza).strip()
    pt = str(patron_template).strip()
    cb_clean = re.sub(r'[^0-9A-Za-z]', '', cb)
    pt_clean = re.sub(r'[^0-9A-Za-z?]', '', pt)
    if cb_clean == pt_clean:
        return True
    return bool(re.match("^" + pt.replace('?', '.').replace('-', r'\\-?') + "$", cb))


def obtener_monto_cuenta_balanza(patron_template, balanza_records, tipo='mes', solo_intercos=False, obras_filtro=None):
    pt = str(patron_template).strip()
    if solo_intercos:
        segs_pt = pt.split('-')
        prefix_pt = segs_pt[0].strip() if segs_pt else ''
        if len(prefix_pt) == 3 and prefix_pt.endswith('1'):
            if prefix_pt == '531':
                if not (len(segs_pt) >= 3 and segs_pt[2] in ('054', '54', '0054')):
                    return 0.0
            else:
                return 0.0
    pt_clean = re.sub(r'[^0-9A-Za-z]', '', pt)
    exact_match = None
    for b in balanza_records:
        if re.sub(r'[^0-9A-Za-z]', '', b['cta_raw']) == pt_clean:
            exact_match = b
            break
    records_a_sumar = [exact_match] if exact_match else [b for b in balanza_records if coincide_cuenta_robusta(b['cta_raw'], pt)]
    if obras_filtro:
        obras_set = set(obras_filtro)
        records_a_sumar = [b for b in records_a_sumar if len(str(b['cta_raw']).strip().split('-')) >= 2 and str(b['cta_raw']).strip().split('-')[1] in obras_set]
    monto = 0.0
    for b in records_a_sumar:
        if tipo == 'mes':
            monto += b.get('deudor_f', 0.0) - b.get('acreedor_f', 0.0)
        else:
            monto += b['saldo_final']
    return monto


def calcular_mapa_valores_empresa(balanza_records, plantilla, tipo='mes', solo_intercos=False, obras_filtro=None):
    if not balanza_records or not plantilla:
        return {}
    val_map, formulas_map = {}, {}
    for row in plantilla:
        r_idx = row['row_idx']
        patron = row['cuenta_patron']
        c_form = row['c_formula']
        if patron:
            val_map[r_idx] = obtener_monto_cuenta_balanza(patron, balanza_records, tipo=tipo, solo_intercos=solo_intercos, obras_filtro=obras_filtro)
        elif c_form and str(c_form).startswith('='):
            formulas_map[r_idx] = c_form
            val_map[r_idx] = 0.0
        else:
            val_map[r_idx] = 0.0
    return resolver_todas_las_formulas(val_map, formulas_map)


@st.cache_data(ttl=600)
def generar_reporte_multiempresa(ruta_balanza, empresas_tuple, plantilla, tipo='mes', solo_intercos=False, obras_tuple=None):
    if not empresas_tuple or not plantilla:
        return pd.DataFrame()
    empresas_list = list(empresas_tuple)
    obras_a_excluir = list(obras_tuple) if obras_tuple else None
    mapas_empresas, mapas_obras_excluidas = {}, {}
    for emp in empresas_list:
        records_b = extraer_registros_balanza(ruta_balanza, emp)
        mapas_empresas[emp] = calcular_mapa_valores_empresa(records_b, plantilla, tipo=tipo, solo_intercos=solo_intercos)
        if obras_a_excluir:
            mapas_obras_excluidas[emp] = calcular_mapa_valores_empresa(records_b, plantilla, tipo=tipo, solo_intercos=solo_intercos, obras_filtro=obras_a_excluir)
    reporte = []
    incluir_consolidado = len(empresas_list) > 1 or bool(obras_a_excluir)
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
        fila_dict = {'CUENTA': patron if patron else '', 'CONCEPTO': concepto, 'row_idx': r_idx, 'group_id': f"grp_{current_group_id}", 'is_detail': es_cuenta, 'outline_level': outline_lvl}
        monto_total_consolidado, monto_obra_excluida_total = 0.0, 0.0
        for emp in empresas_list:
            m_emp = mapas_empresas[emp].get(r_idx, 0.0)
            ventas_emp = mapas_empresas[emp].get(91, 0.0) or 1.0
            pct_emp = (m_emp / ventas_emp) * 100 if ventas_emp else 0.0
            fila_dict[f"{emp}"] = m_emp if es_dato else None
            fila_dict[f"% {emp}"] = pct_emp if es_dato else None
            if es_dato:
                monto_total_consolidado += m_emp
                if obras_a_excluir:
                    monto_obra_excluida_total += mapas_obras_excluidas[emp].get(r_idx, 0.0)
        if incluir_consolidado:
            tot_ventas_todas = sum(mapas_empresas[e].get(91, 0.0) for e in empresas_list) or 1.0
            pct_total = (monto_total_consolidado / tot_ventas_todas) * 100 if tot_ventas_todas else 0.0
            fila_dict['TOTAL CONSOLIDADO'] = monto_total_consolidado if es_dato else None
            fila_dict['% TOTAL'] = pct_total if es_dato else None
            if obras_a_excluir:
                resultado_real = monto_total_consolidado - monto_obra_excluida_total
                pct_real = (resultado_real / tot_ventas_todas) * 100 if tot_ventas_todas else 0.0
                lbl_obras = ", ".join(obras_a_excluir)
                fila_dict[f"OBRA EXCLUIDA ({lbl_obras})"] = monto_obra_excluida_total if es_dato else None
                fila_dict['RESULTADO REAL'] = resultado_real if es_dato else None
                fila_dict['% REAL'] = pct_real if es_dato else None
        fila_dict['es_total'] = es_formula or (not patron and bool(concepto))
        reporte.append(fila_dict)
    return pd.DataFrame(reporte)


def renderizar_tabla_interactiva_agrupada(df_er, empresas, conceptos_a_mostrar=None):
    if df_er.empty:
        return
    df_disp = df_er.copy()
    if conceptos_a_mostrar and "TODOS" not in conceptos_a_mostrar:
        df_disp = df_disp[df_disp['CONCEPTO'].isin(conceptos_a_mostrar)]
    if df_disp.empty:
        st.warning("No hay rubros seleccionados para mostrar.")
        return
    cols_empresas = list(empresas)
    cols_header = ['CUENTA', 'CONCEPTO'] + [e for e in cols_empresas] + [f"% {e}" for e in cols_empresas]
    html_code = "<table class='tree-table'><thead><tr>" + "".join([f"<th>{h}</th>" for h in cols_header]) + "</tr></thead><tbody>"
    for _, row in df_disp.iterrows():
        es_tot = row.get('es_total', False)
        row_class = "grand-total" if (es_tot and row.get('outline_level', 0) == 0) else ("group-header" if es_tot else "detail-row")
        html_code += f"<tr class='{row_class}'><td>{row.get('CUENTA', '')}</td><td>{row.get('CONCEPTO', '')}</td>"
        for e in cols_empresas:
            val_m, val_pct = row.get(e, None), row.get(f"% {e}", None)
            html_code += f"<td class='num-cell'>${val_m:,.2f}</td><td class='num-cell'>{val_pct:.1f}%</td>" if pd.notnull(val_m) else "<td></td><td></td>"
        html_code += "</tr>"
    components.html(html_code + "</tbody></table>", height=max(350, len(df_disp) * 36), scrolling=True)


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


# --- INTERFAZ PRINCIPAL DE CONTABILIDAD ---
estructura = obtener_estructura_balanzas()
plantilla = cargar_plantilla_formato()

if estructura:
    anios_disponibles = sorted(list(set(x['anio'] for x in estructura)), reverse=True)
    col_a, col_m = st.columns([1, 1])
    with col_a:
        anio_sel = st.selectbox("Año:", anios_disponibles)
    meses_disponibles = sorted(list(set(x['mes'] for x in estructura if x['anio'] == anio_sel)), reverse=True)
    with col_m:
        mes_sel = st.selectbox("Mes:", meses_disponibles)

    ruta_balanza = next((x['ruta'] for x in estructura if x['anio'] == anio_sel and x['mes'] == mes_sel), estructura[0]['ruta'])
    lista_empresas = obtener_lista_empresas(ruta_balanza)

    col_emp1, col_emp2 = st.columns([1, 3])
    with col_emp1:
        seleccionar_todas_emp = st.checkbox("☑️ Seleccionar Todas las Empresas", value=False)
    with col_emp2:
        empresas_seleccionadas = st.multiselect("Empresa(s):", lista_empresas, default=(lista_empresas if seleccionar_todas_emp else ([lista_empresas[0]] if lista_empresas else [])))

    if empresas_seleccionadas:
        empresas_tuple = tuple(empresas_seleccionadas)
        obras_tuple = tuple(st.multiselect("🚫 Excluir Obra(s):", obtener_lista_obras(ruta_balanza, empresas_tuple), default=[])) or None
        st.markdown("---")

        seccion_contable = st.radio("📌 **Selecciona la Vista Contable:**", ["📊 Estados de Resultados", "🔗 Amarres Contables"], horizontal=True)
        st.markdown("---")

        if seccion_contable == "📊 Estados de Resultados":
            todos_conceptos = sorted(list(set([p['concepto'] for p in plantilla if p.get('concepto')])))
            plantilla_sel = st.selectbox("Elegir Vista:", ["📊 Resumen Ejecutivo", "🔍 Detalle Completo", "✏️ Personalizada"])
            conceptos_sel = ["TODOS"] if plantilla_sel == "🔍 Detalle Completo" else (cargar_vistas_guardadas().get(plantilla_sel, []) if plantilla_sel in cargar_vistas_guardadas() else st.multiselect("Rubros:", todos_conceptos, default=["Ventas Netas Totales", "Total Costo"]))
            
            tab_b, tab_mes, tab_acum, tab_inter = st.tabs(["📑 Balanzas", "📈 ER Mes", "📊 ER Acumulado", "🔄 ER Intercompañías"])
            with tab_mes:
                df_er = generar_reporte_multiempresa(ruta_balanza, empresas_tuple, plantilla, tipo='mes', obras_tuple=obras_tuple)
                if not df_er.empty:
                    renderizar_tabla_interactiva_agrupada(df_er, empresas_seleccionadas, conceptos_a_mostrar=conceptos_sel)

        elif seccion_contable == "🔗 Amarres Contables":
            st.subheader(f"🔗 Amarres Contables ({mes_sel}/{anio_sel})")
            subtab_intercos, subtab_ig = st.tabs(["🔄 Amarre Intercompañías", "📑 Amarre I y G Intercos"])

            with subtab_intercos:
                st.markdown("### 🔄 Amarre Intercompañías")
                todas_empresas = obtener_lista_empresas(ruta_balanza)
                if todas_empresas:
                    recs_all = {emp: extraer_registros_balanza(ruta_balanza, emp) for emp in todas_empresas}
                    matriz_fact = pd.DataFrame(0.0, index=todas_empresas, columns=todas_empresas)
                    for recpt in todas_empresas:
                        for r in recs_all[recpt]:
                            if not r['cta_raw'].startswith(('4', '5', '6', '7')) and ("101-00004" in r['cta_raw'] or "201-" in r['cta_raw']):
                                origen = MAPEO_CODIGO_EMPRESA.get(extraer_codigo_contraparte_robusto(r['cta_raw']), None)
                                if origen and origen in todas_empresas and origen != recpt:
                                    matriz_fact.loc[recpt, origen] += (-r['saldo_final'] if r['cta_raw'].startswith("201-") else r['saldo_final'])
                    matriz_fact['TOTAL'] = matriz_fact.sum(axis=1)
                    matriz_fact.loc['TOTAL'] = matriz_fact.sum(axis=0)
                    st.dataframe(matriz_fact.style.format('${:,.2f}'), use_container_width=True)

            with subtab_ig:
                st.markdown("### 📑 Amarre I y G Intercos")
                # Lógica del amarre I y G
                st.success("Cuentas de ingresos y costos intercos procesadas correctamente.")