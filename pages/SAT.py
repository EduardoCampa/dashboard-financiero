# --- 2. PROCESAMIENTO CONTABILIDAD NÓMINAS (CORREGIDO PARA LEER 'CARGOS') ---
def cargar_y_procesar_nomina_contabilidad(ruta_base="Balanzas", archivo_balanza_rh="Balanza RH.xlsx", anio_filtro=2026, mes_ini=1, mes_fin=8):
    registros_contables = []

    mapeo_conceptos_fijos = {
        '0001': '-001-0001 Sueldo',
        '0002': '-001-0002 Despensa',
        '0004': '-001-0004 Compensación',
        '0005': '-001-0005 Premio Asistencia',
        '0006': '-001-0006 Premio Puntualidad',
        '0008': '-001-0008 Gratificación Extraordinaria',
        '0009': '-001-0009 Vacaciones',
        '0010': '-001-0010 Prima Vacacional',
        '0011': '-001-0011 Aguinaldo'
    }

    for m in range(int(mes_ini), int(mes_fin) + 1):
        mes_str = f"{m:02d}"
        
        posibles = glob.glob(os.path.join(ruta_base, str(anio_filtro), mes_str, archivo_balanza_rh)) + \
                   glob.glob(os.path.join(ruta_base, "**", mes_str, archivo_balanza_rh), recursive=True)
        
        posibles = list(dict.fromkeys(posibles))
        if not posibles:
            continue

        for ruta_balanza in posibles:
            try:
                nombre_archivo = os.path.basename(ruta_balanza).upper()
                xls = pd.ExcelFile(ruta_balanza)

                for hoja in xls.sheet_names:
                    hoja_upper = hoja.strip().upper()

                    if hoja_upper not in ['HOJA1', 'HOJA 1', 'RESUMEN', 'BALANZA', 'SHEET1']:
                        nombre_empresa = hoja_upper
                    elif 'EFCO' in nombre_archivo:
                        nombre_empresa = 'EFCO'
                    elif 'CIVLAT' in nombre_archivo:
                        nombre_empresa = 'CIVLAT'
                    elif 'FERVIC' in nombre_archivo:
                        nombre_empresa = 'FERVIC'
                    elif 'SERVYRE' in nombre_archivo:
                        nombre_empresa = 'SERVYRE'
                    else:
                        nombre_empresa = hoja_upper

                    df_hoja = pd.read_excel(ruta_balanza, sheet_name=hoja)
                    if df_hoja.empty or df_hoja.shape[1] < 3:
                        continue

                    # BUSCAR DINÁMICAMENTE LA COLUMNA DE CARGOS DEL MES (NO SALDO DEUDOR FINAL)
                    idx_cargos_e = None
                    for idx_c, col_name in enumerate(df_hoja.columns):
                        c_str = str(col_name).strip().upper()
                        if 'CARGO' in c_str or 'DEBITO' in c_str:
                            idx_cargos_e = idx_c
                            break
                    
                    # Si no la encuentra por nombre, tomar la posición de la columna Cargos (columna 4 / E)
                    if idx_cargos_e is None:
                        idx_cargos_e = 4

                    for _, row in df_hoja.iterrows():
                        cta = str(row.iloc[0]).strip()
                        if not cta or cta.lower() in ('nan', 'cuenta', 'none'):
                            continue

                        cta_limpia = cta.replace(' ', '')
                        nom_cuenta = str(row.iloc[1]).strip().title() if df_hoja.shape[1] > 1 else ""

                        match = re.match(r'^(5\d\d|6\d\d)-(\d+)-(001|029|056)-(\d{4})$', cta_limpia)

                        if match:
                            grupo, obra, dept, subcta = match.groups()

                            es_valida = False
                            label_resumen = ""

                            if dept in ['001', '056'] and subcta != '0000':
                                es_valida = True
                                if subcta in mapeo_conceptos_fijos:
                                    label_resumen = mapeo_conceptos_fijos[subcta]
                                else:
                                    label_resumen = f"-{dept}-{subcta} {nom_cuenta}".strip()

                            elif dept == '029' and subcta == '0000':
                                es_valida = True
                                label_resumen = "-029-0000 Asimilados"

                            if es_valida:
                                # Extraer ÚNICAMENTE el movimiento de Cargos del periodo
                                val_cargo = row.iloc[idx_cargos_e] if df_hoja.shape[1] > idx_cargos_e else 0.0
                                monto_cargo = float(pd.to_numeric(str(val_cargo).replace('$', '').replace(',', ''), errors='coerce') or 0.0)

                                if monto_cargo > 0.0:
                                    registros_contables.append({
                                        'Cuenta Completa': f"{cta_limpia} {nom_cuenta}".strip(),
                                        'Subcuenta Resumen': label_resumen,
                                        'Empresa': nombre_empresa,
                                        'Monto': monto_cargo
                                    })
            except Exception:
                continue

    df_cont = pd.DataFrame(registros_contables)
    if df_cont.empty:
        return pd.DataFrame(), pd.DataFrame()

    pivot_resumen = pd.pivot_table(
        df_cont,
        index='Subcuenta Resumen',
        columns='Empresa',
        values='Monto',
        aggfunc='sum',
        fill_value=0.0
    ).reset_index().rename(columns={'Subcuenta Resumen': 'Concepto / Subcuenta'})

    pivot_detalle = pd.pivot_table(
        df_cont,
        index='Cuenta Completa',
        columns='Empresa',
        values='Monto',
        aggfunc='sum',
        fill_value=0.0
    ).reset_index().rename(columns={'Cuenta Completa': 'Cuenta Contable'})

    return pivot_resumen, pivot_detalle