# ==========================================
# 3. SUBMÓDULO: REPORTE EJECUTIVO DE PAGOS
# ==========================================
elif submodulo == "📋 Reporte Ejecutivo de Pagos":
    st.title("📋 Reporte Ejecutivo de Pagos")

    df_oc_base = df_ordenes_proc.copy() if df_ordenes_proc is not None and not df_ordenes_proc.empty else pd.DataFrame()
    df_sp_base = df_tesoreria_proc.copy() if df_tesoreria_proc is not None and not df_tesoreria_proc.empty else pd.DataFrame()

    df_rep_list = []
    if not df_oc_base.empty:
        df_rep_list.append(df_oc_base)
    if not df_sp_base.empty:
        df_rep_list.append(df_sp_base)

    if df_rep_list:
        df_rep_total = pd.concat(df_rep_list, ignore_index=True)

        col_emp_name = 'EmpresaOrigen' if 'EmpresaOrigen' in df_rep_total.columns else None
        col_prov_name = 'BusinessEntityName' if 'BusinessEntityName' in df_rep_total.columns else None
        col_folio_name = 'DocFolio' if 'DocFolio' in df_rep_total.columns else None
        col_currency = 'Currency' if 'Currency' in df_rep_total.columns else None

        if col_prov_name and col_folio_name:
            st.markdown("#### ⚙️ Filtros de Selección Independientes")
            c_rf1, c_rf2, c_rf3, c_rf4 = st.columns(4)
            with c_rf1:
                lista_empresas = sorted(df_rep_total[col_emp_name].dropna().unique()) if col_emp_name else []
                if col_emp_name:
                    empresas_seleccionadas = st.multiselect("Filtrar por Empresa Origen:", lista_empresas, default=[], key="rep_emp_f")
                    if empresas_seleccionadas:
                        df_rep_total = df_rep_total[df_rep_total[col_emp_name].isin(empresas_seleccionadas)]
            with c_rf2:
                tipos_disponibles = sorted(df_rep_total['Tipo_Movimiento'].dropna().unique())
                tipos_seleccionados = st.multiselect("Filtrar por Tipo (OC / SP):", tipos_disponibles, default=[], key="rep_tipo_f")
                if tipos_seleccionados:
                    df_rep_total = df_rep_total[df_rep_total['Tipo_Movimiento'].isin(tipos_seleccionados)]
            with c_rf3:
                lista_proveedores = sorted(df_rep_total[col_prov_name].dropna().unique())
                prov_seleccionados = st.multiselect("Filtrar por Proveedor(es):", lista_proveedores, default=[], key="rep_prov_f")
                if prov_seleccionados:
                    df_rep_total = df_rep_total[df_rep_total[col_prov_name].isin(prov_seleccionados)]
            with c_rf4:
                lista_folios = sorted(df_rep_total[col_folio_name].dropna().unique())
                folios_seleccionados = st.multiselect("Filtrar por Folio(s):", lista_folios, default=[], key="rep_folio_f")
                if folios_seleccionados:
                    df_rep_total = df_rep_total[df_rep_total[col_folio_name].isin(folios_seleccionados)]

            # Checkbox para filtrar los no encontrados en FCoG
            no_fcog_chk = st.checkbox("🚫 No se encuentra en FCoG", value=False, key="rep_no_fcog_chk")

            if no_fcog_chk:
                if 'En_FCoG' in df_rep_total.columns:
                    df_rep_total = df_rep_total[df_rep_total['En_FCoG'] == False]
                col_monto_eval = 'Total'
                label_monto = "Total"
                st.info("Mostrando consolidado de Solicitud de Pago y Orden de Compra NO encontradas en FCoG (Mostrando Total en lugar de Saldo Pendiente).")
            else:
                if 'Saldo_Pendiente' in df_rep_total.columns:
                    df_rep_total = df_rep_total[df_rep_total['Saldo_Pendiente'] > 1.0]
                col_monto_eval = 'Saldo_Pendiente'
                label_monto = "Saldo Pendiente"

            st.markdown("---")

            # --- METRICAS / KPIS GENERALES REPORTE ---
            st.markdown("### 📊 Indicadores Clave de Desempeño (KPIs)")
            col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)

            saldo_mxn = df_rep_total[df_rep_total[col_currency].astype(str).str.contains('Peso|MXN', case=False, na=False)][col_monto_eval].sum() if col_currency else df_rep_total[col_monto_eval].sum()
            saldo_usd = df_rep_total[df_rep_total[col_currency].astype(str).str.contains('Dólar|USD', case=False, na=False)][col_monto_eval].sum() if col_currency else 0.0

            with col_kpi1:
                st.metric("Documentos", f"{len(df_rep_total):,}")
            with col_kpi2:
                st.metric(f"{label_monto} (MXN)", formato_mx(saldo_mxn))
            with col_kpi3:
                st.metric(f"{label_monto} (USD)", formato_mx(saldo_usd))
            with col_kpi4:
                st.metric(f"{label_monto} General", formato_mx(df_rep_total[col_monto_eval].sum()))

            st.markdown("---")

            col_fecha = 'DateDocument' if 'DateDocument' in df_rep_total.columns else None
            col_desc = 'Title' if 'Title' in df_rep_total.columns else None

            if not df_rep_total.empty:
                if col_fecha:
                    df_rep_total['Fecha_Fmt'] = pd.to_datetime(df_rep_total[col_fecha], errors='coerce').dt.strftime('%d/%m/%Y')
                else:
                    df_rep_total['Fecha_Fmt'] = ""

                empresas_agrupadas = df_rep_total[col_emp_name].dropna().unique() if col_emp_name else ['General']
                empresa_para_excel = empresas_agrupadas[0] if len(empresas_agrupadas) == 1 else "Consolidado"

                # Preparamos copia para descarga de Excel manteniendo la columna correspondiente
                df_excel = df_rep_total.copy()
                if no_fcog_chk:
                    df_excel['Saldo_Pendiente'] = df_excel['Total']

                st.download_button(
                    label="📥 Descargar Reporte en Excel con UUID",
                    data=generar_excel_ejecutivo(df_excel, empresa_para_excel),
                    file_name="Reporte_Ejecutivo_Pagos_UUID.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.markdown("---")

                for empresa in sorted(empresas_agrupadas):
                    df_emp_subset = df_rep_total[df_rep_total[col_emp_name] == empresa] if col_emp_name else df_rep_total
                    total_empresa = df_emp_subset[col_monto_eval].sum()
                    st.markdown(f"## 🏢 **{empresa}** ({label_monto} Total: {formato_mx(total_empresa)})")

                    resumen_proveedor = df_emp_subset.groupby(col_prov_name)[col_monto_eval].sum().reset_index()
                    resumen_proveedor = resumen_proveedor.sort_values(by=col_monto_eval, ascending=False)

                    for _, prov_row in resumen_proveedor.iterrows():
                        proveedor = prov_row[col_prov_name]
                        subtotal_prov = prov_row[col_monto_eval]

                        with st.expander(f"👤 {proveedor} — {label_monto} Total: {formato_mx(subtotal_prov)}", expanded=True):
                            df_det_prov = df_emp_subset[df_emp_subset[col_prov_name] == proveedor]
                            monedas_del_prov = sorted(df_det_prov[col_currency].dropna().unique()) if col_currency else ['MXN']

                            for moneda in monedas_del_prov:
                                df_det_moneda = df_det_prov[df_det_prov[col_currency] == moneda] if col_currency else df_det_prov

                                st.markdown(f"##### 💱 Moneda: **{moneda}**")

                                data_det_list = []
                                for _, row in df_det_moneda.iterrows():
                                    data_det_list.append({
                                        "Tipo": row['Tipo_Movimiento'],
                                        "Fecha Vencimiento": row['Fecha_Fmt'],
                                        "Folio / Documento": row[col_folio_name],
                                        "DocumentID": row.get('DocumentID', ''),
                                        "UUID": row.get('UUID', ''),
                                        "Moneda": row[col_currency] if col_currency and not pd.isnull(row[col_currency]) else "MXN",
                                        "Descripción": row[col_desc] if col_desc else "",
                                        label_monto: row[col_monto_eval]
                                    })
                                df_tabla_det = pd.DataFrame(data_det_list)

                                mostrar_tabla_con_totales(
                                    df_tabla_det,
                                    [label_monto],
                                    ["Tipo", "Fecha Vencimiento", "Folio / Documento", "DocumentID", "UUID", "Moneda", "Descripción", label_monto]
                                )
                    st.markdown("---")
            else:
                st.warning("No hay registros que coincidan con los filtros seleccionados.")
    else:
        st.warning("No hay datos cargados para generar el reporte.")