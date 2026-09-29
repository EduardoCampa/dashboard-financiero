with subtab_ig_intercos:
                st.markdown("### 📑 Amarre I y G Intercos (Ingresos facturados y Costos Intercos)")
                st.info("Columna 1: Ingresos facturados por otras empresas a esta. Columna 2: Costos base intercos (-00000-0000). Columna 3: Cuenta 054. Columna 4: Costos netos (Costos base - Cuenta 054).")

                todas_empresas_balanza = obtener_lista_empresas(ruta_balanza)

                if todas_empresas_balanza:
                    datos_empresas_recs = {}
                    for emp in todas_empresas_balanza:
                        datos_empresas_recs[emp] = extraer_registros_balanza(ruta_balanza, emp)

                    filas_ig = []
                    tot_ingresos_gen = 0.0
                    tot_costos_base_gen = 0.0
                    tot_cta_054_gen = 0.0
                    tot_costo_neto_gen = 0.0

                    for emp_destino in todas_empresas_balanza:
                        cod_destino = MAPEO_NOMBRE_A_CODIGO.get(emp_destino, "")

                        # 1. Ingresos facturados a esta empresa desde las demás (buscando el código de la empresa destino en las cuentas 4)
                        ingresos_facturados_a_emp = 0.0
                        cuentas_ingresos_proc = set()
                        if cod_destino:
                            for emp_facturadora in todas_empresas_balanza:
                                if emp_facturadora == emp_destino:
                                    continue
                                recs_f = datos_empresas_recs[emp_facturadora]
                                for r in recs_f:
                                    cta = r['cta_raw'].strip()
                                    # Verificar si es cuenta de ingresos (empieza con 4)
                                    if cta.startswith('4'):
                                        segs = cta.split('-')
                                        # Si el código de la empresa destino está en la cuenta de ingresos
                                        if any(seg.strip() == cod_destino for seg in segs):
                                            # Tomar el movimiento neto del mes (Acreedor - Deudor o Abonos - Cargos del mes)
                                            # Usamos acreedor_f - deudor_f asegurando capturar el flujo neto del mes
                                            monto_movimiento = r.get('acreedor_f', 0.0) - r.get('deudor_f', 0.0)
                                            ingresos_facturados_a_emp += monto_movimiento
                                            cuentas_ingresos_proc.add(cta)

                        # 2. Costos base (-00000-0000) y Cuenta 054 (-054-) en las cuentas de gastos/costos (5, 6)
                        costos_base_emp = 0.0
                        cta_054_emp = 0.0
                        cuentas_base_proc = set()
                        cuentas_054_proc = set()
                        
                        recs_d = datos_empresas_recs[emp_destino]
                        for r in recs_d:
                            cta = r['cta_raw'].strip()
                            segs = cta.split('-')
                            
                            if len(segs) >= 4:
                                prefix = segs[0].strip()
                                seg2 = segs[1].strip()
                                seg3 = segs[2].strip()
                                seg4 = segs[3].strip()

                                if prefix.isdigit() and len(prefix) == 3 and prefix.startswith(('5', '6')) and prefix.endswith('1'):
                                    es_cuenta_base = (seg2 in ('00000', '0000', '0') and seg3 in ('000', '0') and seg4 in ('0000', '0'))
                                    es_054 = (seg3 in ('054', '54', '0054'))

                                    if es_cuenta_base and cta not in cuentas_base_proc:
                                        costos_base_emp += r['saldo_final']
                                        cuentas_base_proc.add(cta)
                                    elif es_054 and cta not in cuentas_054_proc:
                                        cta_054_emp += r['saldo_final']
                                        cuentas_054_proc.add(cta)

                        costo_neto_emp = costos_base_emp - cta_054_emp

                        tot_ingresos_gen += ingresos_facturados_a_emp
                        tot_costos_base_gen += costos_base_emp
                        tot_cta_054_gen += cta_054_emp
                        tot_costo_neto_gen += costo_neto_emp

                        filas_ig.append({
                            'EMPRESA': emp_destino,
                            'INGRESOS FACTURADOS': ingresos_facturados_a_emp,
                            'COSTOS BASE (-00000-)': costos_base_emp,
                            'CUENTA 054': cta_054_emp,
                            'COSTO NETO (COSTOS - 054)': costo_neto_emp,
                        })

                    df_ig_resumen = pd.DataFrame(filas_ig)
                    
                    df_ig_resumen.loc[len(df_ig_resumen)] = {
                        'EMPRESA': 'TOTAL',
                        'INGRESOS FACTURADOS': tot_ingresos_gen,
                        'COSTOS BASE (-00000-)': tot_costos_base_gen,
                        'CUENTA 054': tot_cta_054_gen,
                        'COSTO NETO (COSTOS - 054)': tot_costo_neto_gen,
                    }

                    st.dataframe(
                        df_ig_resumen.style.format({
                            'INGRESOS FACTURADOS': '${:,.2f}',
                            'COSTOS BASE (-00000-)': '${:,.2f}',
                            'CUENTA 054': '${:,.2f}',
                            'COSTO NETO (COSTOS - 054)': '${:,.2f}',
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

                    st.markdown("---")
                    excel_todos = exportar_excel_reporte_ejecutivo_ambas(matriz_fact, matriz_prest, df_ig_resumen, anio_sel, mes_sel)
                    st.download_button(
                        label="📥 Descargar Reporte Ejecutivo Completo (Facturación, Préstamos y I y G) en Excel",
                        data=excel_todos,
                        file_name=f"Reporte_Ejecutivo_Intercompañias_Completo_{mes_sel}_{anio_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                    st.success("✅ Cuentas de ingresos y costos intercos barridas y calculadas correctamente.")