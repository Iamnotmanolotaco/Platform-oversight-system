# app.py - VERSIÓN CORREGIDA PARA STREAMLIT CLOUD

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================

st.set_page_config(
    page_title="📊 Reporte de Tiempos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# FUNCIONES DE UTILIDAD
# ============================================================

def limpiar_nombre(nombre):
    """Limpia el nombre eliminando prefijos de cargo"""
    if not isinstance(nombre, str):
        return nombre
    nombre = nombre.strip()
    prefijos = [
        'Assistant', 'Manager', 'Coordinator', 'Specialist', 
        'Analyst', 'Director', 'Supervisor', 'Lead', 'Senior',
        'Junior', 'Associate', 'Consultant', 'Advisor', 'Officer',
        'Executive', 'Head', 'Chief', 'Principal', 'Partner',
        'CL Finance Assistant', 'Finance Assistant', 'Admin',
        'Operations', 'Support', 'Team Lead', 'Team Leader',
        'Legal Agent Support', 'Customer Support Agent'
    ]
    for prefijo in prefijos:
        if nombre.startswith(prefijo + ' '):
            nombre = nombre[len(prefijo) + 1:]
        elif nombre.startswith(prefijo + '-'):
            nombre = nombre[len(prefijo) + 1:]
        elif nombre.startswith(prefijo + ':'):
            nombre = nombre[len(prefijo) + 1:]
    nombre = re.sub(r'\([^)]*\)', '', nombre).strip()
    nombre = re.sub(r'\[[^\]]*\]', '', nombre).strip()
    nombre = re.sub(r'\s+', ' ', nombre)
    return nombre.strip()

def convertir_fecha(valor):
    """Convierte fecha en formato MM/DD/YYYY"""
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.date()
    if isinstance(valor, str):
        try:
            return datetime.strptime(valor.strip(), '%m/%d/%Y').date()
        except:
            try:
                return pd.to_datetime(valor).date()
            except:
                return None
    try:
        return pd.to_datetime(valor).date()
    except:
        return None

def convertir_hora(valor):
    """Convierte formato HH:MM a horas decimales"""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        valor_limpio = valor.strip()
        if ':' in valor_limpio:
            try:
                partes = valor_limpio.split(':')
                if len(partes) == 2:
                    return float(partes[0]) + float(partes[1]) / 60
            except:
                pass
        try:
            valor_limpio = re.sub(r'[^0-9.]', '', valor_limpio)
            if valor_limpio:
                return float(valor_limpio)
        except:
            pass
        return 0.0
    return 0.0

def convertir_hora_decimal(valor):
    """Convierte formato decimal a horas"""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        valor_limpio = valor.strip()
        if ',' in valor_limpio and '.' not in valor_limpio:
            valor_limpio = valor_limpio.replace(',', '.')
        try:
            valor_limpio = re.sub(r'[^0-9.]', '', valor_limpio)
            if valor_limpio:
                return float(valor_limpio)
        except:
            pass
        return 0.0
    return 0.0

def get_jornada_esperada(fecha):
    """Retorna la jornada esperada para un día"""
    if fecha.weekday() == 5:  # Sábado
        return 4.0
    elif fecha.weekday() == 6:  # Domingo
        return 0.0
    else:
        return 8.0

# ============================================================
# PROCESADOR DE REPORTES (VERSIÓN SIMPLIFICADA)
# ============================================================

class ProcesadorReporte:
    def __init__(self, fecha_inicio, fecha_fin):
        self.fecha_inicio = fecha_inicio
        self.fecha_fin = fecha_fin
        
        self.df_powerbi = None
        self.df_camp = None
        self.df_smokeball = None
        self.df_toggl = None
        self.df_novedades_2 = None
        
        self.mapa_nombres = {}
        self.usuarios_novedades_2 = set()
        self.df_resultado = None
    
    def cargar_archivo(self, archivo, key):
        """Carga un archivo desde BytesIO"""
        try:
            df = pd.read_excel(archivo)
            if key == 'powerbi':
                self.df_powerbi = df
            elif key == 'camp_legal':
                self.df_camp = df
            elif key == 'smokeball':
                self.df_smokeball = df
            elif key == 'toggl':
                self.df_toggl = df
            elif key == 'novedades_2':
                self.df_novedades_2 = df
            return True
        except Exception as e:
            st.error(f"Error en {key}: {e}")
            return False
    
    def construir_mapa_nombres(self):
        """Construye el mapa de nombres desde Power BI"""
        if self.df_powerbi is None:
            return False
        
        # Buscar columnas
        col_nombre = None
        for col in ['NAME CORRECT', 'Name', 'NAME']:
            if col in self.df_powerbi.columns:
                col_nombre = col
                break
        
        if col_nombre is None:
            return False
        
        col_cl = None
        for col in ['NAME CL', 'CL Name', 'CL']:
            if col in self.df_powerbi.columns:
                col_cl = col
                break
        
        col_sb = None
        for col in ['NAME SB', 'SB Name', 'SB']:
            if col in self.df_powerbi.columns:
                col_sb = col
                break
        
        col_tg = None
        for col in ['NAME TG', 'TG Name', 'TG']:
            if col in self.df_powerbi.columns:
                col_tg = col
                break
        
        self.mapa_nombres = {}
        
        for _, row in self.df_powerbi.iterrows():
            nombre_canonico = str(row[col_nombre]).strip() if pd.notna(row[col_nombre]) else None
            if not nombre_canonico:
                continue
            
            nombre_limpio = limpiar_nombre(nombre_canonico)
            self.mapa_nombres[nombre_canonico] = nombre_limpio
            self.mapa_nombres[nombre_limpio] = nombre_limpio
            
            if col_cl and pd.notna(row[col_cl]):
                nombre = limpiar_nombre(str(row[col_cl]).strip())
                if nombre:
                    self.mapa_nombres[nombre] = nombre_limpio
            
            if col_sb and pd.notna(row[col_sb]):
                nombre = limpiar_nombre(str(row[col_sb]).strip())
                if nombre:
                    self.mapa_nombres[nombre] = nombre_limpio
            
            if col_tg and pd.notna(row[col_tg]):
                nombre = limpiar_nombre(str(row[col_tg]).strip())
                if nombre:
                    self.mapa_nombres[nombre] = nombre_limpio
        
        return True
    
    def procesar_novedades_2(self):
        """Procesa la hoja Novedades 2"""
        if self.df_novedades_2 is None:
            return False
        
        col_persona = None
        for col in ['Persona', 'persona', 'NAME']:
            if col in self.df_novedades_2.columns:
                col_persona = col
                break
        
        col_fecha = None
        for col in ['Fecha', 'fecha', 'Date']:
            if col in self.df_novedades_2.columns:
                col_fecha = col
                break
        
        if col_persona is None or col_fecha is None:
            return False
        
        for _, row in self.df_novedades_2.iterrows():
            nombre = str(row[col_persona]).strip() if pd.notna(row[col_persona]) else None
            if not nombre:
                continue
            
            nombre_limpio = self.normalizar_nombre(nombre)
            if nombre_limpio:
                self.usuarios_novedades_2.add(nombre_limpio)
        
        return True
    
    def normalizar_nombre(self, nombre):
        """Normaliza un nombre usando el mapa"""
        if not isinstance(nombre, str):
            return nombre
        
        nombre_limpio = limpiar_nombre(nombre.strip())
        if not nombre_limpio:
            return None
        
        if nombre_limpio in self.mapa_nombres:
            return self.mapa_nombres[nombre_limpio]
        
        if nombre.strip() in self.mapa_nombres:
            return self.mapa_nombres[nombre.strip()]
        
        for nombre_plat, nombre_canon in self.mapa_nombres.items():
            if (nombre_limpio.lower() in nombre_plat.lower() or 
                nombre_plat.lower() in nombre_limpio.lower()):
                return nombre_canon
        
        return nombre_limpio
    
    def procesar_plataforma(self, df, col_nombre, col_horas, col_fecha, es_decimal=False):
        """Procesa una plataforma específica"""
        if df is None:
            return None
        
        if col_nombre not in df.columns:
            return None
        
        if col_horas not in df.columns:
            return None
        
        if col_fecha not in df.columns:
            return None
        
        # Procesar
        df_proc = df.copy()
        
        # Normalizar nombres
        df_proc['Usuario'] = df_proc[col_nombre].astype(str).str.strip().apply(
            lambda x: self.normalizar_nombre(x)
        )
        
        # Filtrar solo usuarios en Novedades 2
        df_proc = df_proc[df_proc['Usuario'].isin(self.usuarios_novedades_2)]
        
        if len(df_proc) == 0:
            return None
        
        # Convertir horas
        if es_decimal:
            df_proc['Horas'] = df_proc[col_horas].apply(convertir_hora_decimal)
        else:
            df_proc['Horas'] = df_proc[col_horas].apply(convertir_hora)
        
        # Convertir fechas
        df_proc['Date'] = df_proc[col_fecha].apply(convertir_fecha)
        
        # Filtrar por rango
        df_proc = df_proc[
            (df_proc['Date'] >= self.fecha_inicio) & 
            (df_proc['Date'] <= self.fecha_fin)
        ]
        
        if len(df_proc) == 0:
            return None
        
        # Agrupar por usuario
        df_agrupado = df_proc.groupby('Usuario').agg({
            'Horas': 'sum'
        }).reset_index()
        
        return df_agrupado
    
    def ejecutar(self):
        """Ejecuta todo el procesamiento"""
        # Construir mapa de nombres
        if not self.construir_mapa_nombres():
            st.error("❌ Error al construir mapa de nombres desde Power BI")
            return False
        
        # Procesar Novedades 2
        if not self.procesar_novedades_2():
            st.error("❌ Error al procesar Novedades 2")
            return False
        
        if not self.usuarios_novedades_2:
            st.warning("⚠️ No hay usuarios en Novedades 2 para este rango")
            self.df_resultado = pd.DataFrame()
            return True
        
        # Procesar cada plataforma
        resultados = {}
        
        # Camp Legal
        df_camp = self.procesar_plataforma(
            self.df_camp,
            'Staff Name', 'Hours Spent', 'Time Entry Date',
            es_decimal=False
        )
        if df_camp is not None:
            resultados['Camp Legal'] = df_camp.rename(columns={'Horas': 'Camp Legal'})
        
        # Smokeball
        df_sb = self.procesar_plataforma(
            self.df_smokeball,
            'Name', 'Hours', 'Date',
            es_decimal=True
        )
        if df_sb is not None:
            resultados['Smokeball'] = df_sb.rename(columns={'Horas': 'Smokeball'})
        
        # Toggl
        df_tg = self.procesar_plataforma(
            self.df_toggl,
            'Member', 'Dur', 'Date1',
            es_decimal=False
        )
        if df_tg is not None:
            resultados['Toggl'] = df_tg.rename(columns={'Horas': 'Toggl'})
        
        # Consolidar resultados
        if not resultados:
            st.warning("⚠️ No se encontraron tiempos para los usuarios en Novedades 2")
            self.df_resultado = pd.DataFrame()
            return True
        
        # Unir todos los resultados
        df_consolidado = None
        for nombre, df in resultados.items():
            if df_consolidado is None:
                df_consolidado = df
            else:
                df_consolidado = df_consolidado.merge(df, on='Usuario', how='outer')
        
        # Llenar NaN con 0
        df_consolidado = df_consolidado.fillna(0)
        
        # Calcular total
        columnas_horas = [col for col in df_consolidado.columns if col in ['Camp Legal', 'Smokeball', 'Toggl']]
        df_consolidado['Total_Horas'] = df_consolidado[columnas_horas].sum(axis=1)
        
        # Calcular porcentaje
        jornada_total = 0
        fecha_actual = self.fecha_inicio
        while fecha_actual <= self.fecha_fin:
            jornada_total += get_jornada_esperada(fecha_actual)
            fecha_actual += timedelta(days=1)
        
        df_consolidado['Porcentaje'] = (df_consolidado['Total_Horas'] / jornada_total * 100).round(1)
        
        # Asignar estado
        def calcular_estado(row):
            total = row['Total_Horas']
            if total == 0:
                return "⛔ Sin registro"
            
            plataformas = []
            if row.get('Camp Legal', 0) > 0:
                plataformas.append('CL')
            if row.get('Smokeball', 0) > 0:
                plataformas.append('SB')
            if row.get('Toggl', 0) > 0:
                plataformas.append('TG')
            
            if len(plataformas) == 0:
                return "⛔ Sin registro"
            elif len(plataformas) == 1:
                return f"⚠️ Solo {plataformas[0]} ({row['Porcentaje']:.0f}%)"
            elif len(plataformas) == 2:
                return f"⚠️ {', '.join(plataformas)} ({row['Porcentaje']:.0f}%)"
            else:
                if row['Porcentaje'] >= 90:
                    return f"✅ Completo ({row['Porcentaje']:.0f}%)"
                elif row['Porcentaje'] >= 70:
                    return f"⚠️ Parcial ({row['Porcentaje']:.0f}%)"
                else:
                    return f"❌ Insuficiente ({row['Porcentaje']:.0f}%)"
        
        df_consolidado['Estado'] = df_consolidado.apply(calcular_estado, axis=1)
        
        # Ordenar por porcentaje
        df_consolidado = df_consolidado.sort_values('Porcentaje', ascending=True)
        
        self.df_resultado = df_consolidado
        return True
    
    def get_resultados(self):
        return self.df_resultado

# ============================================================
# INTERFAZ DE USUARIO
# ============================================================

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a3a5c 0%, #2c5f8a 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
    }
    .main-header h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 700;
    }
    .main-header p {
        margin: 0.3rem 0 0 0;
        opacity: 0.85;
    }
    .card-metric {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        text-align: center;
        border-top: 4px solid #2c5f8a;
    }
    .card-metric .value {
        font-size: 28px;
        font-weight: 800;
        color: #1a3a5c;
    }
    .card-metric .label {
        font-size: 11px;
        text-transform: uppercase;
        color: #7a8a9e;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1>📊 Reporte de Tiempos</h1>
    <p>Análisis de tiempos por plataforma · Novedades 2</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("⚙️ Configuración")
    
    st.subheader("📅 Rango de fechas")
    col1, col2 = st.columns(2)
    with col1:
        fecha_inicio = st.date_input(
            "Fecha inicio",
            value=datetime.now().date() - timedelta(days=7),
            format="MM/DD/YYYY"
        )
    with col2:
        fecha_fin = st.date_input(
            "Fecha fin",
            value=datetime.now().date() - timedelta(days=1),
            format="MM/DD/YYYY"
        )
    
    st.divider()
    st.subheader("📁 Archivos")
    
    uploaded_files = {}
    
    uploaded_files['powerbi'] = st.file_uploader(
        "📊 Power BI resources.xlsx (Names)",
        type=['xlsx'],
        key="powerbi"
    )
    
    uploaded_files['novedades_2'] = st.file_uploader(
        "📋 Novedades 2 (Template_Novedades_RRHH_MAX)",
        type=['xlsx'],
        key="novedades_2"
    )
    
    uploaded_files['camp_legal'] = st.file_uploader(
        "🏛️ Camp Legal (Time entries)",
        type=['xlsx'],
        key="camp_legal"
    )
    
    uploaded_files['smokeball'] = st.file_uploader(
        "📋 Smokeball (Entries)",
        type=['xlsx'],
        key="smokeball"
    )
    
    uploaded_files['toggl'] = st.file_uploader(
        "⏱️ Toggl (DataBaseToggl)",
        type=['xlsx'],
        key="toggl"
    )
    
    st.divider()
    
    archivos_cargados = sum([1 for v in uploaded_files.values() if v is not None])
    st.markdown(f"**📁 Archivos cargados:** {archivos_cargados}/5")
    
    procesar = st.button(
        "🚀 Generar Reporte",
        type="primary",
        use_container_width=True,
        disabled=archivos_cargados < 3
    )

# Procesamiento
if not procesar:
    st.info("👈 Sube los archivos requeridos y presiona 'Generar Reporte'")
    
    st.markdown("### 📋 Archivos requeridos")
    st.markdown("""
    **Obligatorios:**
    - ✅ Power BI Resources (Names)
    - ✅ Novedades 2 (Template_Novedades_RRHH_MAX)
    - ✅ Al menos una plataforma (Camp Legal, Smokeball o Toggl)
    
    **El reporte solo mostrará usuarios que están en Novedades 2**
    """)
    
    st.stop()

# Procesar
with st.spinner("🔄 Procesando datos..."):
    try:
        # Inicializar
        procesador = ProcesadorReporte(fecha_inicio, fecha_fin)
        
        # Cargar archivos
        for key, file in uploaded_files.items():
            if file is not None:
                procesador.cargar_archivo(file, key)
        
        # Ejecutar
        if not procesador.ejecutar():
            st.stop()
        
        # Obtener resultados
        df_resultado = procesador.get_resultados()
        
        if df_resultado is None or df_resultado.empty:
            st.warning("⚠️ No hay datos para mostrar")
            st.stop()
        
        # Mostrar resultados
        st.markdown("---")
        st.markdown("### 📊 Resultados")
        
        # Estadísticas
        total_usuarios = len(df_resultado)
        total_horas = df_resultado['Total_Horas'].sum()
        promedio = df_resultado['Total_Horas'].mean()
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"""
            <div class="card-metric">
                <div class="value">{total_usuarios}</div>
                <div class="label">👥 Usuarios</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #27ae60;">
                <div class="value">{total_horas:.1f}h</div>
                <div class="label">⏱️ Total Horas</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #e67e22;">
                <div class="value">{promedio:.1f}h</div>
                <div class="label">📊 Promedio</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Tabla
        st.markdown("### 👥 Detalle por Usuario")
        st.dataframe(df_resultado, use_container_width=True, hide_index=True)
        
        # Descargar
        st.markdown("### 📥 Exportar")
        csv = df_resultado.to_csv(index=False).encode('utf-8')
        
        st.download_button(
            label="📥 Descargar CSV",
            data=csv,
            file_name=f"reporte_{fecha_inicio.strftime('%Y%m%d')}_{fecha_fin.strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.exception(e)
