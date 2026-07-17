# ============================================================
# PROCESAR REPORTE CON DIAGNÓSTICO PASO A PASO
# ============================================================

with st.spinner("🔄 Procesando datos... Por favor espera"):
    try:
        # ============================================================
        # PASO 1: INICIALIZAR
        # ============================================================
        st.write("### 📋 PASO 1: Inicializando procesador...")
        procesador = ReporteTiemposSystem(fecha_inicio, fecha_fin)
        st.success("✅ Procesador inicializado")
        
        # ============================================================
        # PASO 2: CARGAR ARCHIVOS
        # ============================================================
        st.write("### 📋 PASO 2: Cargando archivos...")
        
        archivos_cargados = {}
        
        for key, file in uploaded_files.items():
            if file is not None:
                st.write(f"📖 Cargando {key}...")
                
                if key == 'novedades_max':
                    # Cargar hoja Novedades
                    try:
                        df_novedades = pd.read_excel(file, sheet_name='Novedades')
                        procesador.df_novedades_max = df_novedades
                        st.success(f"✅ Novedades (hoja 1): {len(df_novedades)} registros")
                        st.write(f"   Columnas: {list(df_novedades.columns)}")
                    except Exception as e:
                        st.error(f"❌ Error cargando Novedades: {e}")
                        st.stop()
                    
                    # Cargar hoja Novedades 2
                    try:
                        df_novedades_2 = pd.read_excel(file, sheet_name='Novedades 2')
                        procesador.df_novedades_max_2 = df_novedades_2
                        st.success(f"✅ Novedades 2 (hoja 2): {len(df_novedades_2)} registros")
                        st.write(f"   Columnas: {list(df_novedades_2.columns)}")
                        
                        # Mostrar muestra
                        if len(df_novedades_2) > 0:
                            st.write("📊 **Muestra de Novedades 2:**")
                            st.dataframe(df_novedades_2.head(5))
                    except Exception as e:
                        st.error(f"❌ Error cargando Novedades 2: {e}")
                        st.stop()
                
                elif key == 'powerbi':
                    try:
                        df = pd.read_excel(file)
                        procesador.df_powerbi = df
                        st.success(f"✅ Power BI: {len(df)} registros")
                        st.write(f"   Columnas: {list(df.columns)}")
                    except Exception as e:
                        st.error(f"❌ Error cargando Power BI: {e}")
                        st.stop()
                
                else:
                    try:
                        df = pd.read_excel(file)
                        if key == 'camp_legal':
                            procesador.df_camp = df
                        elif key == 'smokeball':
                            procesador.df_smokeball = df
                        elif key == 'toggl':
                            procesador.df_toggl = df
                        elif key == 'novedades_clg':
                            procesador.df_novedades_clg = df
                        st.success(f"✅ {key}: {len(df)} registros")
                    except Exception as e:
                        st.error(f"❌ Error cargando {key}: {e}")
                        st.stop()
        
        st.success("✅ Todos los archivos cargados")
        
        # ============================================================
        # PASO 3: CONSTRUIR MAPA DE NOMBRES
        # ============================================================
        st.write("### 📋 PASO 3: Construyendo mapa de nombres...")
        
        if procesador.df_powerbi is None:
            st.error("❌ No hay datos de Power BI")
            st.stop()
        
        cols = COLUMNAS_MAPEO['powerbi']['columnas']
        col_canonico = cols.get('nombre', 'NAME CORRECT')
        col_cl = cols.get('nombre_cl', 'NAME CL')
        col_sb = cols.get('nombre_sb', 'NAME SB')
        col_tg = cols.get('nombre_tg', 'NAME TG')
        col_status = cols.get('status', 'USER STATUS')
        
        st.write(f"🔍 Buscando columna de nombres: '{col_canonico}'")
        
        if col_canonico not in procesador.df_powerbi.columns:
            st.error(f"❌ Columna '{col_canonico}' no encontrada")
            st.write(f"📋 Columnas disponibles: {list(procesador.df_powerbi.columns)}")
            st.stop()
        
        if col_status in procesador.df_powerbi.columns:
            df_activos = procesador.df_powerbi[procesador.df_powerbi[col_status] == 'Active'].copy()
            st.success(f"✅ Usuarios activos: {len(df_activos)}")
        else:
            df_activos = procesador.df_powerbi.copy()
            st.warning(f"⚠️ Columna '{col_status}' no encontrada, usando todos los usuarios")
        
        procesador.mapa_nombres = {}
        procesador.usuarios_con_plataforma = []
        
        contador = 0
        for idx, row in df_activos.iterrows():
            nombre_canonico = str(row[col_canonico]).strip() if pd.notna(row[col_canonico]) else None
            if not nombre_canonico:
                continue
            
            nombre_canonico_limpio = limpiar_nombre(nombre_canonico)
            
            tiene_cl = False
            tiene_sb = False
            tiene_tg = False
            
            if col_cl in df_activos.columns and pd.notna(row[col_cl]):
                valor = str(row[col_cl]).strip()
                if valor and valor.lower() not in ['true', 'false', 'nan', 'none', '']:
                    tiene_cl = True
                    procesador.mapa_nombres[valor] = nombre_canonico_limpio
                    procesador.mapa_nombres[limpiar_nombre(valor)] = nombre_canonico_limpio
            
            if col_sb in df_activos.columns and pd.notna(row[col_sb]):
                valor = str(row[col_sb]).strip()
                if valor and valor.lower() not in ['true', 'false', 'nan', 'none', '']:
                    tiene_sb = True
                    procesador.mapa_nombres[valor] = nombre_canonico_limpio
                    procesador.mapa_nombres[limpiar_nombre(valor)] = nombre_canonico_limpio
            
            if col_tg in df_activos.columns and pd.notna(row[col_tg]):
                valor = str(row[col_tg]).strip()
                if valor and valor.lower() not in ['true', 'false', 'nan', 'none', '']:
                    tiene_tg = True
                    procesador.mapa_nombres[valor] = nombre_canonico_limpio
                    procesador.mapa_nombres[limpiar_nombre(valor)] = nombre_canonico_limpio
            
            procesador.mapa_nombres[nombre_canonico] = nombre_canonico_limpio
            procesador.mapa_nombres[nombre_canonico_limpio] = nombre_canonico_limpio
            
            if tiene_cl or tiene_sb or tiene_tg:
                procesador.usuarios_con_plataforma.append(nombre_canonico_limpio)
            
            contador += 1
            if contador % 50 == 0:
                st.write(f"   Procesados {contador} usuarios...")
        
        st.success(f"✅ Mapa de nombres construido: {len(procesador.mapa_nombres)} entradas")
        st.success(f"✅ Usuarios con plataforma: {len(procesador.usuarios_con_plataforma)}")
        
        # ============================================================
        # PASO 4: PROCESAR NOVEDADES
        # ============================================================
        st.write("### 📋 PASO 4: Procesando novedades...")
        
        novedades_list = []
        
        # Novedades MAX
        if procesador.df_novedades_max is not None:
            st.write("📖 Procesando Novedades MAX...")
            df_max = procesador.df_novedades_max.copy()
            if 'Persona' in df_max.columns and 'Fecha Inicio' in df_max.columns and 'Fecha Fin' in df_max.columns:
                df_max['Usuario_Normalizado'] = df_max['Persona'].apply(procesador.normalizar_nombre)
                df_max['Fecha_Inicio'] = df_max['Fecha Inicio'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_max['Fecha_Fin'] = df_max['Fecha Fin'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_max_validos = df_max[df_max['Fecha_Inicio'].notna() & df_max['Fecha_Fin'].notna()]
                if 'Tipo de Novedad' in df_max.columns:
                    df_max_validos['Tipo'] = df_max_validos['Tipo de Novedad']
                else:
                    df_max_validos['Tipo'] = 'Permiso MAX'
                novedades_list.append(df_max_validos[['Usuario_Normalizado', 'Fecha_Inicio', 'Fecha_Fin', 'Tipo']])
                st.success(f"✅ Novedades MAX: {len(df_max_validos)} registros")
            else:
                st.warning(f"⚠️ Columnas no encontradas en Novedades MAX: {list(df_max.columns)}")
        
        # Novedades MAX 2
        if procesador.df_novedades_max_2 is not None:
            st.write("📖 Procesando Novedades 2...")
            df_max2 = procesador.df_novedades_max_2.copy()
            if 'Persona' in df_max2.columns and 'Fecha' in df_max2.columns:
                df_max2['Fecha_Conv'] = df_max2['Fecha'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_max2_filtrado = df_max2[
                    (df_max2['Fecha_Conv'] >= fecha_inicio) & 
                    (df_max2['Fecha_Conv'] <= fecha_fin)
                ]
                
                st.write(f"   Registros en el rango: {len(df_max2_filtrado)}")
                
                for _, row in df_max2_filtrado.iterrows():
                    nombre = row['Persona']
                    nombre_normalizado = procesador.normalizar_nombre(nombre)
                    if nombre_normalizado:
                        procesador.usuarios_novedades_2.add(nombre_normalizado)
                
                st.success(f"✅ Usuarios en Novedades 2 en el rango: {len(procesador.usuarios_novedades_2)}")
                
                df_max2_validos = df_max2_filtrado[df_max2_filtrado['Fecha_Conv'].notna()]
                if 'Tipo de Novedad' in df_max2.columns:
                    df_max2_validos['Tipo'] = df_max2_validos['Tipo de Novedad']
                else:
                    df_max2_validos['Tipo'] = 'Novedad 2'
                df_max2_validos['Fecha_Inicio'] = df_max2_validos['Fecha_Conv']
                df_max2_validos['Fecha_Fin'] = df_max2_validos['Fecha_Conv']
                df_max2_validos['Usuario_Normalizado'] = df_max2_validos['Persona'].apply(procesador.normalizar_nombre)
                novedades_list.append(df_max2_validos[['Usuario_Normalizado', 'Fecha_Inicio', 'Fecha_Fin', 'Tipo']])
                st.success(f"✅ Novedades 2: {len(df_max2_validos)} registros en el rango")
            else:
                st.warning(f"⚠️ Columnas no encontradas en Novedades 2: {list(df_max2.columns)}")
        
        # Novedades CLG
        if procesador.df_novedades_clg is not None:
            st.write("📖 Procesando Novedades CLG...")
            df_clg = procesador.df_novedades_clg.copy()
            if 'Persona' in df_clg.columns and 'Fecha Inicio' in df_clg.columns and 'Fecha Fin' in df_clg.columns:
                df_clg['Usuario_Normalizado'] = df_clg['Persona'].apply(procesador.normalizar_nombre)
                df_clg['Fecha_Inicio'] = df_clg['Fecha Inicio'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_clg['Fecha_Fin'] = df_clg['Fecha Fin'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_clg_validos = df_clg[df_clg['Fecha_Inicio'].notna() & df_clg['Fecha_Fin'].notna()]
                if 'Tipo de Novedad' in df_clg.columns:
                    df_clg_validos['Tipo'] = df_clg_validos['Tipo de Novedad']
                else:
                    df_clg_validos['Tipo'] = 'Permiso CLG'
                novedades_list.append(df_clg_validos[['Usuario_Normalizado', 'Fecha_Inicio', 'Fecha_Fin', 'Tipo']])
                st.success(f"✅ Novedades CLG: {len(df_clg_validos)} registros")
            else:
                st.warning(f"⚠️ Columnas no encontradas en Novedades CLG: {list(df_clg.columns)}")
        
        if novedades_list:
            procesador.df_novedades_combinadas = pd.concat(novedades_list, ignore_index=True)
            st.success(f"✅ Total novedades combinadas: {len(procesador.df_novedades_combinadas)}")
        else:
            procesador.df_novedades_combinadas = None
            st.warning("⚠️ No se encontraron novedades")
        
        # Mostrar usuarios de Novedades 2
        if procesador.usuarios_novedades_2:
            st.write("📋 **Usuarios en Novedades 2:**")
            for i, usuario in enumerate(sorted(procesador.usuarios_novedades_2)[:20]):
                st.write(f"   {i+1}. {usuario}")
            if len(procesador.usuarios_novedades_2) > 20:
                st.write(f"   ... y {len(procesador.usuarios_novedades_2) - 20} más")
        
        # Verificar usuarios en Novedades 2
        if not procesador.usuarios_novedades_2:
            st.warning("⚠️ No hay usuarios en Novedades 2 para el rango seleccionado.")
            st.info("💡 Todos los usuarios están descansando en este período.")
            st.stop()
        
        # ============================================================
        # PASO 5: PROCESAR PLATAFORMAS
        # ============================================================
        st.write("### 📋 PASO 5: Procesando plataformas...")
        
        # Camp Legal
        st.write("📖 Procesando Camp Legal...")
        df_camp = procesador.procesar_plataforma(procesador.df_camp, 'camp_legal')
        if df_camp is not None:
            st.success(f"✅ Camp Legal: {len(df_camp)} usuarios con tiempo")
            st.dataframe(df_camp.head(5))
        else:
            st.warning("⚠️ Camp Legal: No hay datos en el rango")
        
        # Smokeball
        st.write("📖 Procesando Smokeball...")
        df_sb = procesador.procesar_plataforma(procesador.df_smokeball, 'smokeball')
        if df_sb is not None:
            st.success(f"✅ Smokeball: {len(df_sb)} usuarios con tiempo")
            st.dataframe(df_sb.head(5))
        else:
            st.warning("⚠️ Smokeball: No hay datos en el rango")
        
        # Toggl
        st.write("📖 Procesando Toggl...")
        df_tg = procesador.procesar_plataforma(procesador.df_toggl, 'toggl')
        if df_tg is not None:
            st.success(f"✅ Toggl: {len(df_tg)} usuarios con tiempo")
            st.dataframe(df_tg.head(5))
        else:
            st.warning("⚠️ Toggl: No hay datos en el rango")
        
        # ============================================================
        # PASO 6: CONSOLIDAR
        # ============================================================
        st.write("### 📋 PASO 6: Consolidando resultados...")
        
        usuarios_dict = {}
        for usuario in procesador.usuarios_novedades_2:
            usuarios_dict[usuario] = {
                'Camp Legal': 0.0,
                'Smokeball': 0.0,
                'Toggl': 0.0,
                'Dias_Activos': 0,
                'Actividades': [],
                'Permiso': None
            }
        
        def procesar_plataforma_detalle(df_plat, plataforma):
            if df_plat is None:
                return
            
            for _, row in df_plat.iterrows():
                usuario = row['Usuario_Normalizado']
                if usuario in usuarios_dict:
                    usuarios_dict[usuario][plataforma] = row['Horas']
                    usuarios_dict[usuario]['Dias_Activos'] = max(usuarios_dict[usuario]['Dias_Activos'], row['Dias_Activos'])
                    if row['Actividad'] and row['Actividad'] != 'Sin actividad':
                        usuarios_dict[usuario]['Actividades'].append(f"{plataforma[:4]}: {row['Actividad']}")
        
        procesar_plataforma_detalle(df_camp, 'Camp Legal')
        procesar_plataforma_detalle(df_sb, 'Smokeball')
        procesar_plataforma_detalle(df_tg, 'Toggl')
        
        st.write("🔍 Verificando permisos...")
        for usuario in usuarios_dict:
            fecha_actual = fecha_inicio
            while fecha_actual <= fecha_fin:
                permiso = procesador.verificar_permiso(usuario, fecha_actual)
                if permiso:
                    usuarios_dict[usuario]['Permiso'] = permiso
                    break
                fecha_actual += timedelta(days=1)
        
        datos = []
        for usuario, data in usuarios_dict.items():
            total = data['Camp Legal'] + data['Smokeball'] + data['Toggl']
            porcentaje = (total / procesador.jornada_total_esperada * 100) if procesador.jornada_total_esperada > 0 else 0
            
            if data['Permiso']:
                estado = f"📋 {data['Permiso']}"
            else:
                estado = procesador.calcular_estado(total, data['Camp Legal'], data['Smokeball'], data['Toggl'], porcentaje)
            
            datos.append({
                'Usuario': usuario,
                'Camp Legal': round(data['Camp Legal'], 2),
                'Smokeball': round(data['Smokeball'], 2),
                'Toggl': round(data['Toggl'], 2),
                'Total_Horas': round(total, 2),
                'Dias_Activos': data['Dias_Activos'],
                'Actividades': ' | '.join(data['Actividades']) if data['Actividades'] else 'Sin registro',
                'Permiso': data['Permiso'] if data['Permiso'] else 'Sin permiso',
                'Estado': estado,
                'Porcentaje': round(porcentaje, 1),
                'Plataformas_Activas': sum([1 for x in [data['Camp Legal'], data['Smokeball'], data['Toggl']] if x > 0])
            })
        
        procesador.df_analisis = pd.DataFrame(datos)
        st.success(f"✅ Consolidado: {len(procesador.df_analisis)} usuarios")
        
        # ============================================================
        # PASO 7: MOSTRAR RESULTADOS
        # ============================================================
        
        df_resultados = procesador.df_analisis
        estadisticas = procesador.get_estadisticas()
        
        if df_resultados is None or df_resultados.empty:
            st.warning("⚠️ No se encontraron resultados para el rango seleccionado.")
            st.stop()
        
        st.markdown("---")
        st.markdown("### 📊 Resultados del Reporte")
        
        # Información de Novedades 2
        st.markdown(f"""
        <div class="info-box">
            <strong>📋 Usuarios en Novedades 2:</strong> {len(procesador.usuarios_novedades_2)} personas deben marcar tiempo
            <span style="margin-left:20px;">⛔ <strong>Usuarios NO en Novedades 2:</strong> Están descansando y no aparecen en el reporte</span>
            <br>
            <span style="font-size:12px; color:#7d6608;">📌 Sábados: 4h · Festivos: 0h</span>
        </div>
        """, unsafe_allow_html=True)
        
        # KPIs
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class="card-metric">
                <div class="value">{estadisticas['total_usuarios']}</div>
                <div class="label">👥 Usuarios</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #8e44ad;">
                <div class="value">{estadisticas['con_permiso']}</div>
                <div class="label">📋 Permisos</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #27ae60;">
                <div class="value">{estadisticas['total_horas']:.1f}h</div>
                <div class="label">⏱️ Total Horas</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #e67e22;">
                <div class="value">{estadisticas['promedio']:.1f}h</div>
                <div class="label">📊 Promedio</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Resumen por plataforma
        st.markdown("### 📈 Distribución por Plataforma")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #3498db;">
                <div class="value" style="color: #3498db;">{estadisticas['horas_camp']:.1f}h</div>
                <div class="label">🏛️ Camp Legal</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #2ecc71;">
                <div class="value" style="color: #2ecc71;">{estadisticas['horas_sb']:.1f}h</div>
                <div class="label">📋 Smokeball</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #e67e22;">
                <div class="value" style="color: #e67e22;">{estadisticas['horas_tg']:.1f}h</div>
                <div class="label">⏱️ Toggl</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Tabla de resultados
        st.markdown("### 👥 Detalle por Usuario")
        
        df_mostrar = df_resultados[[
            'Usuario', 'Camp Legal', 'Smokeball', 'Toggl', 
            'Total_Horas', 'Dias_Activos', 'Permiso', 'Estado'
        ]].copy()
        
        df_mostrar.columns = [
            'Usuario', 'Camp Legal', 'Smokeball', 'Toggl',
            'Total Horas', 'Días Activos', 'Permiso', 'Estado'
        ]
        
        st.dataframe(
            df_mostrar,
            use_container_width=True,
            height=400,
            hide_index=True
        )
        
        # Descargar
        st.markdown("### 📥 Exportar")
        
        csv = df_resultados.to_csv(index=False).encode('utf-8')
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="📥 Descargar CSV",
                data=csv,
                file_name=f"reporte_{fecha_inicio.strftime('%Y%m%d')}_{fecha_fin.strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col2:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_resultados.to_excel(writer, sheet_name='Reporte', index=False)
            output.seek(0)
            
            st.download_button(
                label="📥 Descargar Excel",
                data=output,
                file_name=f"reporte_{fecha_inicio.strftime('%Y%m%d')}_{fecha_fin.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.exception(e)
