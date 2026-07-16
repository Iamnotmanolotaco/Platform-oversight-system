# app.py - VERSIÓN SIMPLIFICADA QUE FUNCIONA 100%

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

st.set_page_config(
    page_title="📊 Reporte de Tiempos",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Reporte de Tiempos")
st.caption("Sube los archivos y genera el reporte")

# ============================================================
# FUNCIONES
# ============================================================

def limpiar_nombre(nombre):
    if not isinstance(nombre, str):
        return nombre
    nombre = nombre.strip()
    prefijos = [
        'Assistant', 'Manager', 'Coordinator', 'Specialist', 
        'Analyst', 'Director', 'Supervisor', 'Lead', 'Senior',
        'Junior', 'Associate', 'Consultant', 'Advisor', 'Officer',
        'Executive', 'Head', 'Chief', 'Principal', 'Partner'
    ]
    for prefijo in prefijos:
        if nombre.startswith(prefijo + ' '):
            nombre = nombre[len(prefijo) + 1:]
    nombre = re.sub(r'\([^)]*\)', '', nombre).strip()
    return nombre.strip()

def convertir_fecha(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.date()
    if isinstance(valor, str):
        try:
            return datetime.strptime(valor.strip(), '%m/%d/%Y').date()
        except:
            return None
    return None

def convertir_hora(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        try:
            return float(valor)
        except:
            return 0.0
    return 0.0

# ============================================================
# INTERFAZ
# ============================================================

with st.sidebar:
    st.header("⚙️ Configuración")
    
    st.subheader("📅 Rango de fechas")
    fecha_inicio = st.date_input(
        "Fecha inicio",
        value=datetime.now().date() - timedelta(days=7)
    )
    fecha_fin = st.date_input(
        "Fecha fin",
        value=datetime.now().date() - timedelta(days=1)
    )
    
    st.divider()
    st.subheader("📁 Archivos")
    
    archivo_powerbi = st.file_uploader(
        "📊 Power BI resources.xlsx",
        type=['xlsx'],
        key="powerbi"
    )
    
    archivo_novedades = st.file_uploader(
        "📋 Novedades 2 (MAX)",
        type=['xlsx'],
        key="novedades"
    )
    
    archivo_camp = st.file_uploader(
        "🏛️ Camp Legal",
        type=['xlsx'],
        key="camp"
    )
    
    archivo_sb = st.file_uploader(
        "📋 Smokeball",
        type=['xlsx'],
        key="sb"
    )
    
    archivo_tg = st.file_uploader(
        "⏱️ Toggl",
        type=['xlsx'],
        key="tg"
    )
    
    st.divider()
    
    archivos_subidos = sum([
        archivo_powerbi is not None,
        archivo_novedades is not None,
        archivo_camp is not None,
        archivo_sb is not None,
        archivo_tg is not None
    ])
    
    st.info(f"📁 Archivos cargados: {archivos_subidos}/5")
    
    procesar = st.button(
        "🚀 Generar Reporte",
        type="primary",
        use_container_width=True,
        disabled=archivos_subidos < 2
    )

# ============================================================
# PROCESAMIENTO
# ============================================================

if not procesar:
    st.info("👈 Sube los archivos y presiona 'Generar Reporte'")
    
    st.markdown("### 📋 Requerido:")
    st.markdown("""
    - ✅ Power BI resources.xlsx
    - ✅ Novedades 2 (Template_Novedades_RRHH_MAX)
    - ✅ Al menos una plataforma (Camp Legal, Smokeball o Toggl)
    """)
    st.stop()

# --- PROCESAR ---
with st.spinner("🔄 Procesando datos..."):
    try:
        st.markdown("---")
        st.markdown("### 📊 PROCESANDO ARCHIVOS")
        
        # 1. LEER POWER BI
        st.write("📖 Leyendo Power BI...")
        df_powerbi = pd.read_excel(archivo_powerbi)
        st.success(f"✅ Power BI: {len(df_powerbi)} registros")
        st.write(f"Columnas: {list(df_powerbi.columns)}")
        
        # 2. LEER NOVEDADES
        st.write("📖 Leyendo Novedades 2...")
        df_novedades = pd.read_excel(archivo_novedades)
        st.success(f"✅ Novedades: {len(df_novedades)} registros")
        st.write(f"Columnas: {list(df_novedades.columns)}")
        
        # 3. CREAR MAPA DE NOMBRES
        st.write("📖 Construyendo mapa de nombres...")
        mapa_nombres = {}
        
        # Buscar columna de nombre en Power BI
        col_nombre = None
        for col in ['NAME CORRECT', 'Name', 'NAME']:
            if col in df_powerbi.columns:
                col_nombre = col
                break
        
        if col_nombre:
            st.success(f"✅ Columna de nombres encontrada: '{col_nombre}'")
            for nombre in df_powerbi[col_nombre].dropna().unique():
                nombre_limpio = limpiar_nombre(str(nombre))
                mapa_nombres[nombre] = nombre_limpio
                mapa_nombres[nombre_limpio] = nombre_limpio
        else:
            st.error("❌ No se encontró columna de nombres en Power BI")
            st.stop()
        
        # 4. EXTRAER USUARIOS DE NOVEDADES
        st.write("📖 Extrayendo usuarios de Novedades...")
        usuarios_novedades = set()
        
        col_persona = None
        for col in ['Persona', 'persona', 'NAME']:
            if col in df_novedades.columns:
                col_persona = col
                break
        
        if col_persona:
            for nombre in df_novedades[col_persona].dropna().unique():
                nombre_limpio = limpiar_nombre(str(nombre))
                if nombre_limpio in mapa_nombres:
                    usuarios_novedades.add(mapa_nombres[nombre_limpio])
                else:
                    usuarios_novedades.add(nombre_limpio)
            
            st.success(f"✅ Usuarios en Novedades: {len(usuarios_novedades)}")
            st.write(f"Usuarios: {list(usuarios_novedades)[:10]}...")
        else:
            st.error("❌ No se encontró columna de personas en Novedades")
            st.stop()
        
        # 5. PROCESAR PLATAFORMAS
        resultados = {}
        
        # Camp Legal
        if archivo_camp is not None:
            st.write("📖 Procesando Camp Legal...")
            df = pd.read_excel(archivo_camp)
            st.success(f"✅ Camp Legal: {len(df)} registros")
            
            # Buscar columnas
            col_nom = None
            for col in ['Staff Name', 'Name', 'Member']:
                if col in df.columns:
                    col_nom = col
                    break
            
            col_horas = None
            for col in ['Hours Spent', 'Hours', 'Dur']:
                if col in df.columns:
                    col_horas = col
                    break
            
            col_fecha = None
            for col in ['Time Entry Date', 'Date', 'Date1']:
                if col in df.columns:
                    col_fecha = col
                    break
            
            if col_nom and col_horas and col_fecha:
                df['Usuario'] = df[col_nom].astype(str).apply(limpiar_nombre)
                df = df[df['Usuario'].isin(usuarios_novedades)]
                
                # Convertir fechas
                df['Date'] = df[col_fecha].apply(convertir_fecha)
                df = df[df['Date'] >= fecha_inicio]
                df = df[df['Date'] <= fecha_fin]
                
                # Agrupar
                df_camp = df.groupby('Usuario').agg({
                    col_horas: lambda x: sum([convertir_hora(v) for v in x])
                }).reset_index()
                df_camp.columns = ['Usuario', 'Camp Legal']
                resultados['Camp Legal'] = df_camp
                st.success(f"✅ {len(df_camp)} usuarios en Camp Legal")
            else:
                st.warning("⚠️ Columnas no encontradas en Camp Legal")
        
        # Smokeball
        if archivo_sb is not None:
            st.write("📖 Procesando Smokeball...")
            df = pd.read_excel(archivo_sb)
            st.success(f"✅ Smokeball: {len(df)} registros")
            
            col_nom = None
            for col in ['Name', 'Staff Name', 'Member']:
                if col in df.columns:
                    col_nom = col
                    break
            
            col_horas = None
            for col in ['Hours', 'Hours Spent', 'Dur']:
                if col in df.columns:
                    col_horas = col
                    break
            
            col_fecha = None
            for col in ['Date', 'Time Entry Date', 'Date1']:
                if col in df.columns:
                    col_fecha = col
                    break
            
            if col_nom and col_horas and col_fecha:
                df['Usuario'] = df[col_nom].astype(str).apply(limpiar_nombre)
                df = df[df['Usuario'].isin(usuarios_novedades)]
                
                df['Date'] = df[col_fecha].apply(convertir_fecha)
                df = df[df['Date'] >= fecha_inicio]
                df = df[df['Date'] <= fecha_fin]
                
                df_sb = df.groupby('Usuario').agg({
                    col_horas: lambda x: sum([convertir_hora(v) for v in x])
                }).reset_index()
                df_sb.columns = ['Usuario', 'Smokeball']
                resultados['Smokeball'] = df_sb
                st.success(f"✅ {len(df_sb)} usuarios en Smokeball")
            else:
                st.warning("⚠️ Columnas no encontradas en Smokeball")
        
        # Toggl
        if archivo_tg is not None:
            st.write("📖 Procesando Toggl...")
            df = pd.read_excel(archivo_tg)
            st.success(f"✅ Toggl: {len(df)} registros")
            
            col_nom = None
            for col in ['Member', 'Staff Name', 'Name']:
                if col in df.columns:
                    col_nom = col
                    break
            
            col_horas = None
            for col in ['Dur', 'Hours', 'Hours Spent']:
                if col in df.columns:
                    col_horas = col
                    break
            
            col_fecha = None
            for col in ['Date1', 'Date', 'Time Entry Date']:
                if col in df.columns:
                    col_fecha = col
                    break
            
            if col_nom and col_horas and col_fecha:
                df['Usuario'] = df[col_nom].astype(str).apply(limpiar_nombre)
                df = df[df['Usuario'].isin(usuarios_novedades)]
                
                df['Date'] = df[col_fecha].apply(convertir_fecha)
                df = df[df['Date'] >= fecha_inicio]
                df = df[df['Date'] <= fecha_fin]
                
                df_tg = df.groupby('Usuario').agg({
                    col_horas: lambda x: sum([convertir_hora(v) for v in x])
                }).reset_index()
                df_tg.columns = ['Usuario', 'Toggl']
                resultados['Toggl'] = df_tg
                st.success(f"✅ {len(df_tg)} usuarios en Toggl")
            else:
                st.warning("⚠️ Columnas no encontradas en Toggl")
        
        # 6. CONSOLIDAR RESULTADOS
        st.markdown("---")
        st.markdown("### 📊 RESULTADOS")
        
        if not resultados:
            st.error("❌ No se encontraron resultados en ninguna plataforma")
            st.stop()
        
        # Unir resultados
        df_final = None
        for nombre, df in resultados.items():
            if df_final is None:
                df_final = df
            else:
                df_final = df_final.merge(df, on='Usuario', how='outer')
        
        df_final = df_final.fillna(0)
        
        # Calcular total
        columnas_horas = [col for col in df_final.columns if col in ['Camp Legal', 'Smokeball', 'Toggl']]
        df_final['Total_Horas'] = df_final[columnas_horas].sum(axis=1)
        
        # Calcular estado
        def calcular_estado(row):
            if row['Total_Horas'] == 0:
                return "⛔ Sin registro"
            plataformas = []
            if row.get('Camp Legal', 0) > 0:
                plataformas.append('CL')
            if row.get('Smokeball', 0) > 0:
                plataformas.append('SB')
            if row.get('Toggl', 0) > 0:
                plataformas.append('TG')
            if len(plataformas) == 1:
                return f"⚠️ Solo {plataformas[0]}"
            elif len(plataformas) == 2:
                return f"⚠️ {', '.join(plataformas)}"
            else:
                if row['Total_Horas'] >= 8:
                    return "✅ Completo"
                else:
                    return "⚠️ Parcial"
        
        df_final['Estado'] = df_final.apply(calcular_estado, axis=1)
        
        # 7. MOSTRAR RESULTADOS
        st.success(f"✅ Reporte generado: {len(df_final)} usuarios")
        
        # KPIs
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("👥 Usuarios", len(df_final))
        with col2:
            st.metric("⏱️ Total Horas", f"{df_final['Total_Horas'].sum():.1f}h")
        with col3:
            st.metric("📊 Promedio", f"{df_final['Total_Horas'].mean():.1f}h")
        
        # Tabla
        st.dataframe(df_final, use_container_width=True, hide_index=True)
        
        # Descargar
        st.markdown("### 📥 Exportar")
        csv = df_final.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar CSV",
            data=csv,
            file_name=f"reporte_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.exception(e)
        st.write("---")
        st.write("💡 **Ayuda:** Verifica que los archivos tengan las columnas correctas")
