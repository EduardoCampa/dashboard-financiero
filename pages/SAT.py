# ==========================================
# 2. SUBMÓDULO: AMARRE NÓMINAS
# ==========================================
elif submodulo_sat == "👥 Amarre Nóminas":

    st.title("👥 Módulo SAT - Amarre de Nómina vs Contabilidad")

    # --- FUNCIONALIDAD MEJORADA PARA LEER REPORTES HORIZONTALES DE NÓMINA (eToolSAT / SAT) ---
    def cargar_y_procesar_nomina_sat(carpeta_nomina="Nomina", anio_filtro=None, mes_ini=None, mes_fin=None):
        if not os.path.exists(carpeta_nomina):
            return pd.DataFrame()

        archivos = (
            glob.glob(os.path.join(carpeta_nomina, "**", "*.xlsx"), recursive=True) + 
            glob.glob(os.path.join(carpeta_nomina, "*.xlsx"))
        )
        archivos = list(set(archivos))
        dfs_procesados = []

        # Columnas típicas que NO son conceptos numéricos de nómina
        cols_omitir_keywords = [
            'uuid', 'folio', 'estado', 'estatus', 'version', 'serie', 'tipo', 
            'moneda', 'tc', 'cancelac', 'sustituc', 'sello', 'certificado', 
            'versionnomina', 'origenrecurso', 'curp', 'nss', 'numempleado', 
            'departamento', 'puesto', 'riesgopuesto', 'periodicidad', 'banco', 
            'cuentabancaria', 'claveentfed', 'metodopago'
        ]

        for arch in archivos:
            try:
                # Extraer RFC del nombre del archivo si está presente (ej. CIV141222JD5...)
                nombre_base = os.path.basename(arch).upper()
                rfc_archivo = nombre_base.split('-')[0].strip() if '-' in nombre_base else ""

                xls = pd.ExcelFile(arch)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    if df.empty:
                        continue

                    # Limpiar nombres de columnas
                    df.columns = [str(c).strip() for c in df.columns]

                    # 1. Identificar RFC Emisor
                    col_rfc = next((c for c in df.columns if 'rfc' in c.lower() and 'emisor' in c.lower()), None)
                    if not col_rfc:
                        col_rfc = next((c for c in df.columns if 'rfc' in c.lower()), None)

                    if col_rfc:
                        df['Empresa_RFC_Clean'] = df[col_rfc].astype(str).str.strip().str.upper()
                    elif rfc_archivo:
                        df['Empresa_RFC_Clean'] = rfc_archivo
                    else:
                        df['Empresa_RFC_Clean'] = "EMPRESA NÓMINA"

                    # 2. Identificar Fecha para Agrupar por Mes
                    col_fecha = next((c for c in df.columns if 'fecha' in c.lower() and ('pago' in c.lower() or 'emision' in c.lower() or 'certific' in c.lower())), None)
                    if not col_fecha:
                        col_fecha = next((c for c in df.columns if 'fecha' in c.lower()), None)

                    if not col_fecha:
                        continue

                    df['Fecha_Dt'] = pd.to_datetime(df[col_fecha], errors='coerce')
                    df = df[df['Fecha_Dt'].notnull()].copy()

                    # Aplicar filtros de Periodo
                    if anio_filtro:
                        df = df[df['Fecha_Dt'].dt.year == int(anio_filtro)]
                    if mes_ini and mes_fin:
                        df = df[(df['Fecha_Dt'].dt.month >= int(mes_ini)) & (df['Fecha_Dt'].dt.month <= int(mes_fin))]

                    if df.empty:
                        continue

                    df['Mes_Pago'] = df['Fecha_Dt'].dt.strftime('%Y-%m')

                    # 3. Detectar Columnas de Conceptos (Columnas numéricas de Percepciones/Deducciones/Totales)
                    cols_conceptos = []
                    for c in df.columns:
                        c_low = c.lower()
                        # Excluir identificadores y fechas
                        if c in ['Empresa_RFC_Clean', 'Fecha_Dt', 'Mes_Pago', col_fecha]:
                            continue
                        if any(kw in c_low for kw in cols_omitir_keywords):
                            continue
                        # Verificar si la columna contiene valores numéricos o conceptos (ej. 001/001/Sueldo, Total Gravado, etc.)
                        val_num = pd.to_numeric(df[c].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False), errors='coerce')
                        if val_num.notnull().any() and val_num.abs().sum() > 0:
                            df[c] = val_num.fillna(0.0)
                            cols_conceptos.append(c)

                    if not cols_conceptos:
                        continue

                    # Agrupar este archivo/hoja por RFC y Mes
                    df_agg = df.groupby(['Empresa_RFC_Clean', 'Mes_Pago'])[cols_conceptos].sum().reset_index()
                    dfs_procesados.append(df_agg)

            except Exception:
                continue

        if not dfs_procesados:
            return pd.DataFrame()

        # Unir todos los archivos procesados
        df_consolidado = pd.concat(dfs_procesados, ignore_index=True)
        
        # Agrupar nuevamente por si había datos del mismo RFC y Mes en distintos archivos/hojas
        cols_totales = [c for c in df_consolidado.columns if c not in ['Empresa_RFC_Clean', 'Mes_Pago']]
        df_final = df_consolidado.groupby(['Empresa_RFC_Clean', 'Mes_Pago'])[cols_totales].sum().reset_index()

        df_final = df_final.rename(columns={'Empresa_RFC_Clean': 'Empresa / RFC'})
        
        # Opcional: Eliminar columnas que estén totalmente en $0.00
        for c in cols_totales:
            if df_final[c].abs().sum() == 0:
                df_final.drop(columns=[c], inplace=True)

        return df_final


    # --- CONTROLES DE FILTRO ---
    col_a_nom, col_mini_nom, col_mfin_nom = st.columns([1, 1, 1])

    with col_a_nom:
        anio_sel_nom = st.selectbox("Año de Filtro:", [2026, 2025, 2024], index=0, key="nom_anio")
    with col_mini_nom:
        mes_inicial_nom = st.selectbox("Mes Inicial:", list(range(1, 13)), index=0, format_func=lambda x: nombres_meses[x-1], key="nom_mini")
    with col_mfin_nom:
        mes_final_nom = st.selectbox("Mes Final:", list(range(1, 13)), index=7, format_func=lambda x: nombres_meses[x-1], key="nom_mfin")

    st.markdown("---")
    carpeta_nomina_input = st.text_input("Carpeta Raíz de Nómina (XMLs / Excel):", value="Nomina", key="nom_dir")
    ruta_balanzas_nom_input = st.text_input("Carpeta Raíz de Balanzas:", value="Balanzas", key="nom_bal_dir")

    if st.button("🚀 Ejecutar Amarre Nóminas"):
        with st.spinner("Procesando archivos de la carpeta Nómina y consolidando conceptos por columna..."):
            df_nomina_sat = cargar_y_procesar_nomina_sat(
                carpeta_nomina=carpeta_nomina_input,
                anio_filtro=anio_sel_nom,
                mes_ini=mes_inicial_nom,
                mes_fin=mes_final_nom
            )

            st.success("✅ Procesamiento de Amarre de Nóminas completado.")

            subtab_nom_sat, subtab_nom_cont = st.tabs(["📑 SAT", "📊 Contabilidad"])

            with subtab_nom_sat:
                st.markdown(f"#### 📋 Consolidado de Nómina SAT por RFC, Mes y Conceptos ({nombres_meses[mes_inicial_nom-1]} a {nombres_meses[mes_final_nom-1]} {anio_sel_nom})")
                if not df_nomina_sat.empty:
                    num_cols = [col for col in df_nomina_sat.columns if col not in ['Empresa / RFC', 'Mes_Pago']]
                    format_dict = {col: '${:,.2f}' for col in num_cols}
                    
                    st.dataframe(
                        df_nomina_sat.style.format(format_dict),
                        use_container_width=True,
                        hide_index=True
                    )

                    output_nom = io.BytesIO()
                    with pd.ExcelWriter(output_nom, engine='openpyxl') as writer:
                        df_nomina_sat.to_excel(writer, sheet_name='Consolidado Nómina SAT', index=False)
                    excel_data_nom = output_nom.getvalue()

                    st.download_button(
                        label="📥 Descargar Consolidado Nómina SAT en Excel",
                        data=excel_data_nom,
                        file_name=f"Consolidado_Nomina_SAT_{nombres_meses[mes_final_nom-1]}_{anio_sel_nom}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No se encontraron registros de nómina en la carpeta especificada para el periodo seleccionado.")

            with subtab_nom_cont:
                st.markdown("#### 📊 Contabilidad de Nómina")
                st.info("Espacio preparado para la integración con las balanzas/cuentas contables de Nómina.")