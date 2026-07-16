# app.py - VERSIÓN DE DIAGNÓSTICO PASO A PASO

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="🔍 Diagnóstico", layout="wide")

st.title("🔍 DIAGNÓSTICO - Reporte de Tiempos")
st.write("Esta app solo muestra qué archivos se cargan y su estructura")

# Sidebar
with st.sidebar:
    st.header("📁 Sube los archivos")
    
    archivo_powerbi = st.file_uploader("1. Power BI resources.xlsx", type=['xlsx'])
    archivo_novedades = st.file_uploader("2. Novedades 2 (MAX)", type=['xlsx'])
    archivo_camp = st.file_uploader("3. Camp Legal", type=['xlsx'])
    archivo_sb = st.file_uploader("4. Smokeball", type=['xlsx'])
    archivo_tg = st.file_uploader("5. Toggl", type=['xlsx'])
    
    procesar = st.button("🔍 Diagnosticar", type="primary", use_container_width=True)

# Área principal
if not procesar:
    st.info("👈 Sube los archivos y presiona 'Diagnosticar'")
    st.stop()

# ============================================================
# DIAGNÓSTICO
# ============================================================

st.markdown("---")
st.markdown("## 📊 RESULTADOS DEL DIAGNÓSTICO")

archivos = {
    'Power BI': archivo_powerbi,
    'Novedades 2': archivo_novedades,
    'Camp Legal': archivo_camp,
    'Smokeball': archivo_sb,
    'Toggl': archivo_tg
}

for nombre, archivo in archivos.items():
    st.markdown(f"### 📄 {nombre}")
    
    if archivo is None:
        st.warning(f"⚠️ No se subió el archivo de {nombre}")
        continue
    
    try:
        # Intentar leer el archivo
        df = pd.read_excel(archivo)
        
        st.success(f"✅ Archivo cargado: {len(df)} registros")
        st.write(f"**Columnas:** {list(df.columns)}")
        
        # Mostrar primeras filas
        st.write("**Muestra de datos:**")
        st.dataframe(df.head(3), use_container_width=True)
        
        # Información específica según el archivo
        if nombre == 'Power BI':
            if 'NAME CORRECT' in df.columns:
                st.success(f"✅ Columna 'NAME CORRECT' encontrada con {df['NAME CORRECT'].nunique()} nombres únicos")
            if 'USER STATUS' in df.columns:
                activos = df[df['USER STATUS'] == 'Active']
                st.success(f"✅ Usuarios activos: {len(activos)}")
        
        if nombre == 'Novedades 2':
            if 'Persona' in df.columns:
                st.success(f"✅ Columna 'Persona' encontrada con {df['Persona'].nunique()} personas")
            if 'Fecha' in df.columns:
                st.success(f"✅ Columna 'Fecha' encontrada")
                # Mostrar rango de fechas
                try:
                    fechas = pd.to_datetime(df['Fecha'])
                    st.write(f"Rango de fechas: {fechas.min()} a {fechas.max()}")
                except:
                    st.warning("⚠️ No se pudieron convertir las fechas")
        
        if nombre in ['Camp Legal', 'Smokeball', 'Toggl']:
            # Buscar columnas comunes
            cols_buscar = ['Staff Name', 'Name', 'Member', 'Hours Spent', 'Hours', 'Dur', 'Time Entry Date', 'Date', 'Date1']
            for col in cols_buscar:
                if col in df.columns:
                    st.success(f"✅ Columna '{col}' encontrada")
        
        st.divider()
        
    except Exception as e:
        st.error(f"❌ Error al leer {nombre}: {e}")
        st.code(str(e))
