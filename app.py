# app.py - VERSIÓN COMPLETA Y FUNCIONAL PARA STREAMLIT

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
import re
import warnings
from io import BytesIO
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
# FUNCIONES DE CONVERSIÓN
# ============================================================

def convertir_hora_tiempo(valor):
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
                elif len(partes) == 3:
                    return float(partes[0]) + float(partes[1]) / 60 + float(partes[2]) / 3600
            except:
                pass
        try:
            valor_limpio = re.sub(r'[^0-9.]', '', valor_limpio)
            if valor_limpio:
                return float(valor_limpio)
        except:
            pass
        return 0.0
    if isinstance(valor, pd.Timedelta):
        return valor.total_seconds() / 3600
    return 0.0

def convertir_hora_decimal(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        valor_limpio = valor.strip()
        if ',' in valor_limpio and '.' not in valor_limpio:
            valor_limpio = valor_limpio.replace(',', '.')
        if 'h' in valor_limpio.lower():
            valor_limpio = valor_limpio.lower().replace('h', '').strip()
        try:
            valor_limpio = re.sub(r'[^0-9.]', '', valor_limpio)
            if valor_limpio:
                return float(valor_limpio)
        except:
            pass
        return 0.0
    if isinstance(valor, pd.Timedelta):
        return valor.total_seconds() / 3600
    return 0.0

def convertir_horas_segun_formato(valor, formato):
    if formato == 'tiempo':
        return convertir_hora_tiempo(valor)
    elif formato == 'decimal':
        return convertir_hora_decimal(valor)
    else:
        resultado = convertir_hora_tiempo(valor)
        if resultado == 0:
            resultado = convertir_hora_decimal(valor)
        return resultado

def limpiar_nombre(nombre):
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
        elif nombre.startswith(prefijo):
            if len(nombre) > len(prefijo) and nombre[len(prefijo)] in [' ', '-', ':']:
                nombre = nombre[len(prefijo) + 1:]
    nombre = re.sub(r'\([^)]*\)', '', nombre).strip()
    nombre = re.sub(r'\[[^\]]*\]', '', nombre).strip()
    nombre = re.sub(r'\s+', ' ', nombre)
    return nombre.strip()

def convertir_fecha(valor, formato_fecha='%m/%d/%Y'):
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.date()
    if isinstance(valor, str):
        try:
            return datetime.strptime(valor.strip(), formato_fecha).date()
        except:
            try:
                return pd.to_datetime(valor).date()
            except:
                return None
    try:
        return pd.to_datetime(valor).date()
    except:
        return None

def es_festivo(fecha):
    festivos = [
        (1, 1), (5, 1), (7, 20), (8, 7), (12, 8), (12, 25)
    ]
    return (fecha.month, fecha.day) in festivos

def get_jornada_esperada_por_dia(fecha):
    if fecha.weekday() == 5:
        return 4.0
    elif fecha.weekday() == 6:
        return 0.0
    elif es_festivo(fecha):
        return 0.0
    else:
        return 8.0

# ============================================================
# CONFIGURACIÓN DE COLUMNAS
# ============================================================

COLUMNAS_MAPEO = {
    'camp_legal': {
        'archivo': 'Reporte Diario Camp Legal.xlsx',
        'columnas': {
            'nombre': 'Staff Name',
            'horas': 'Hours Spent',
            'fecha': 'Time Entry Date',
            'actividad': 'Activity'
        },
        'formato_horas': 'tiempo',
        'formato_fecha': '%m/%d/%Y'
    },
    'smokeball': {
        'archivo': 'Reporte_general.xlsx',
        'columnas': {
            'nombre': 'Name',
            'horas': 'Hours',
            'fecha': 'Date',
            'actividad': 'Subject'
        },
        'formato_horas': 'decimal',
        'formato_fecha': '%m/%d/%Y'
    },
    'toggl': {
        'archivo': 'Revision de entradas de tiempo - Toggl.xlsx',
        'columnas': {
            'nombre': 'Member',
            'horas': 'Dur',
            'fecha': 'Date1',
            'actividad': 'Project'
        },
        'formato_horas': 'tiempo',
        'formato_fecha': '%m/%d/%Y'
    },
    'powerbi': {
        'archivo': 'Power BI resources.xlsx',
        'columnas': {
            'nombre': 'NAME CORRECT',
            'nombre_cl': 'NAME CL',
            'nombre_sb': 'NAME SB',
            'nombre_tg': 'NAME TG',
            'status': 'USER STATUS'
        }
    }
}

# ============================================================
# CLASE PROCESADOR
# ============================================================

class ProcesadorReporte:
    def __init__(self, fecha_inicio, fecha_fin):
        self.fecha_inicio = fecha_inicio
        self.fecha_fin = fecha_fin
        self.jornada_esperada = 8.0
        
        self.df_camp = None
        self.df_smokeball = None
        self.df_toggl = None
        self.df_powerbi = None
        self.df_novedades = None
        self.df_analisis = None
        
        self.mapa_nombres = {}
        self.usuarios_novedades = set()
        
        self.dias_habiles = self._calcular_dias_habiles()
        self.jornada_total_esperada = self._calcular_jornada_total()
    
    def _calcular_dias_habiles(self):
        dias = 0
        fecha_actual = self.fecha_inicio
        while fecha_actual <= self.fecha_fin:
            if fecha_actual.weekday() < 5:
                dias += 1
            fecha_actual += timedelta(days=1)
        return max(dias, 1)
    
    def _calcular_jornada_total(self):
        total = 0
        fecha_actual = self.fecha_inicio
        while fecha_actual <= self.fecha_fin:
            total += get_jornada_esperada_por_dia(fecha_actual)
            fecha_actual += timedelta(days=1)
        return max(total, 1)
    
    def cargar_archivo(self, archivo_bytes, key):
        try:
            df = pd.read_excel(archivo_bytes)
            
            if key == 'powerbi':
                self.df_powerbi = df
            elif key == 'camp_legal':
                self.df_camp = df
            elif key == 'smokeball':
                self.df_smokeball = df
            elif key == 'toggl':
                self.df_toggl = df
            elif key == 'novedades':
                self.df_novedades = df
            return True
        except Exception as e:
            st.error(f"Error cargando {key}: {e}")
            return False
    
    def normalizar_nombre(self, nombre):
        if not isinstance(nombre, str):
            return nombre
        nombre_limpio = limpiar_nombre(nombre.strip())
        if not nombre_limpio:
            return nombre_limpio
        if nombre_limpio in self.mapa_nombres:
            return self.mapa_nombres[nombre_limpio]
        if nombre.strip() in self.mapa_nombres:
            return self.mapa_nombres[nombre.strip()]
        for nombre_plat, nombre_canon in self.mapa_nombres.items():
            if (nombre_limpio.lower() in nombre_plat.lower() or 
                nombre_plat.lower() in nombre_limpio.lower()):
                return nombre_canon
        return nombre_limpio
    
    def construir_mapa_nombres(self):
        if self.df_powerbi is None:
            return False
        
        col_canonico = 'NAME CORRECT'
        
        if col_canonico not in self.df_powerbi.columns:
            return False
        
        for nombre in self.df_powerbi[col_canonico].dropna().unique():
            nombre_limpio = limpiar_nombre(str(nombre))
            self.mapa_nombres[nombre] = nombre_limpio
            self.mapa_nombres[nombre_limpio] = nombre_limpio
        
        return True
    
    def procesar_novedades(self):
        if self.df_novedades is None:
            return False
        
        col_persona = None
        for col in ['Persona', 'persona', 'NAME']:
            if col in self.df_novedades.columns:
                col_persona = col
                break
        
        col_fecha = None
        for col in ['Fecha', 'fecha', 'Date']:
            if col in self.df_novedades.columns:
                col_fecha = col
                break
        
        if col_persona is None or col_fecha is None:
            return False
        
        for _, row in self.df_novedades.iterrows():
            nombre = str(row[col_persona]).strip() if pd.notna(row[col_persona]) else None
            if not nombre:
                continue
            nombre_limpio = self.normalizar_nombre(nombre)
            if nombre_limpio:
                self.usuarios_novedades.add(nombre_limpio)
        
        return True
    
    def procesar_plataforma(self, df, config_key):
        if df is None:
            return None
        
        config = COLUMNAS_MAPEO[config_key]
        cols = config['columnas']
        formato_horas = config.get('formato_horas', 'auto')
        
        col_nombre = cols.get('nombre')
        col_horas = cols.get('horas')
        col_fecha = cols.get('fecha')
        col_actividad = cols.get('actividad')
        
        if col_nombre not in df.columns or col_horas not in df.columns or col_fecha not in df.columns:
            return None
        
        df_proc = df.copy()
        df_proc['Usuario_Normalizado'] = df_proc[col_nombre].astype(str).str.strip().apply(self.normalizar_nombre)
        df_proc['Horas'] = df_proc[col_horas].apply(lambda x: convertir_horas_segun_formato(x, formato_horas))
        
        try:
            df_proc['Date'] = df_proc[col_fecha].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
        except:
            df_proc['Date'] = self.fecha_inicio
        
        df_proc = df_proc[
            (df_proc['Date'] >= self.fecha_inicio) & 
            (df_proc['Date'] <= self.fecha_fin)
        ]
        
        if len(df_proc) == 0:
            return None
        
        if col_actividad in df.columns:
            df_proc['Actividad'] = df_proc[col_actividad].astype(str)
        else:
            df_proc['Actividad'] = 'Sin actividad'
        
        df_agrupado = df_proc.groupby('Usuario_Normalizado').agg({
            'Horas': 'sum',
            'Actividad': lambda x: ', '.join(x.unique()[:3])
        }).reset_index()
        
        return df_agrupado
    
    def consolidar(self):
        if not self.usuarios_novedades:
            self.df_analisis = pd.DataFrame()
            return True
        
        df_camp = self.procesar_plataforma(self.df_camp, 'camp_legal')
        df_sb = self.procesar_plataforma(self.df_smokeball, 'smokeball')
        df_tg = self.procesar_plataforma(self.df_toggl, 'toggl')
        
        usuarios_dict = {}
        for usuario in self.usuarios_novedades:
            usuarios_dict[usuario] = {
                'Camp Legal': 0.0,
                'Smokeball': 0.0,
                'Toggl': 0.0,
                'Actividades': []
            }
        
        def procesar_plataforma_detalle(df_plat, plataforma):
            if df_plat is None:
                return
            
            for _, row in df_plat.iterrows():
                usuario = row['Usuario_Normalizado']
                if usuario in usuarios_dict:
                    usuarios_dict[usuario][plataforma] = row['Horas']
                    if row['Actividad'] and row['Actividad'] != 'Sin actividad':
                        usuarios_dict[usuario]['Actividades'].append(f"{plataforma[:4]}: {row['Actividad']}")
        
        procesar_plataforma_detalle(df_camp, 'Camp Legal')
        procesar_plataforma_detalle(df_sb, 'Smokeball')
        procesar_plataforma_detalle(df_tg, 'Toggl')
        
        datos = []
        for usuario, data in usuarios_dict.items():
            total = data['Camp Legal'] + data['Smokeball'] + data['Toggl']
            porcentaje = (total / self.jornada_total_esperada * 100) if self.jornada_total_esperada > 0 else 0
            
            if total == 0:
                estado = "⛔ Sin registro"
            else:
                plataformas = []
                if data['Camp Legal'] > 0:
                    plataformas.append('CL')
                if data['Smokeball'] > 0:
                    plataformas.append('SB')
                if data['Toggl'] > 0:
                    plataformas.append('TG')
                
                if len(plataformas) == 1:
                    estado = f"⚠️ Solo {plataformas[0]} ({porcentaje:.0f}%)"
                elif len(plataformas) == 2:
                    estado = f"⚠️ {', '.join(plataformas)} ({porcentaje:.0f}%)"
                else:
                    if porcentaje >= 90:
                        estado = f"✅ Completo ({porcentaje:.0f}%)"
                    elif porcentaje >= 70:
                        estado = f"⚠️ Parcial ({porcentaje:.0f}%)"
                    else:
                        estado = f"❌ Insuficiente ({porcentaje:.0f}%)"
            
            datos.append({
                'Usuario': usuario,
                'Camp Legal': round(data['Camp Legal'], 2),
                'Smokeball': round(data['Smokeball'], 2),
                'Toggl': round(data['Toggl'], 2),
                'Total_Horas': round(total, 2),
                'Actividades': ' | '.join(data['Actividades']) if data['Actividades'] else 'Sin registro',
                'Estado': estado,
                'Porcentaje': round(porcentaje, 1)
            })
        
        self.df_analisis = pd.DataFrame(datos)
        return True
    
    def obtener_resultados(self):
        return self.df_analisis
    
    def get_estadisticas(self):
        if self.df_analisis is None or self.df_analisis.empty:
            return {
                'total_usuarios': 0,
                'total_horas': 0,
                'promedio': 0,
                'horas_camp': 0,
                'horas_sb': 0,
                'horas_tg': 0
            }
        
        df = self.df_analisis
        return {
            'total_usuarios': len(df),
            'total_horas': df['Total_Horas'].sum(),
            'promedio': df['Total_Horas'].mean(),
            'horas_camp': df['Camp Legal'].sum(),
            'horas_sb': df['Smokeball'].sum(),
            'horas_tg': df['Toggl'].sum()
        }

# ============================================================
# INTERFAZ DE USUARIO
# ============================================================

st.title("📊 Reporte de Tiempos")
st.caption("Análisis de tiempos por plataforma · Control de calidad")

with st.sidebar:
    st.header("⚙️ Configuración")
    
    st.subheader("📅 Rango de fechas")
    fecha_inicio = st.date_input(
        "Fecha inicio",
        value=datetime.now().date() - timedelta(days=7),
        format="MM/DD/YYYY"
    )
    fecha_fin = st.date_input(
        "Fecha fin",
        value=datetime.now().date() - timedelta(days=1),
        format="MM/DD/YYYY"
    )
    
    st.divider()
    st.subheader("📁 Archivos")
    
    archivo_powerbi = st.file_uploader(
        "📊 Power BI resources.xlsx",
        type=['xlsx'],
        key="powerbi"
    )
    
    archivo_novedades = st.file_uploader(
        "📋 Novedades (Template_Novedades_RRHH_MAX)",
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
    
    archivos_cargados = sum([
        archivo_powerbi is not None,
        archivo_novedades is not None,
        archivo_camp is not None,
        archivo_sb is not None,
        archivo_tg is not None
    ])
    
    st.info(f"📁 Archivos cargados: {archivos_cargados}/5")
    
    procesar = st.button(
        "🚀 Generar Reporte",
        type="primary",
        use_container_width=True,
        disabled=archivos_cargados < 2
    )

# ============================================================
# PROCESAR REPORTE
# ============================================================

if not procesar:
    st.info("👈 Sube los archivos requeridos (Power BI y Novedades) y presiona 'Generar Reporte'")
    st.markdown("""
    **✅ Requeridos:**
    - Power BI resources.xlsx
    - Novedades (Template_Novedades_RRHH_MAX)
    
    **📌 Opcionales (al menos una):**
    - Camp Legal
    - Smokeball
    - Toggl
    """)
    st.stop()

# ============================================================
# PROCESAMIENTO
# ============================================================

with st.spinner("🔄 Procesando datos... Por favor espera"):
    try:
        # Inicializar
        procesador = ProcesadorReporte(fecha_inicio, fecha_fin)
        
        # Cargar archivos
        if archivo_powerbi is not None:
            procesador.cargar_archivo(archivo_powerbi, 'powerbi')
            st.success("✅ Power BI cargado")
        
        if archivo_novedades is not None:
            procesador.cargar_archivo(archivo_novedades, 'novedades')
            st.success("✅ Novedades cargado")
        
        if archivo_camp is not None:
            procesador.cargar_archivo(archivo_camp, 'camp_legal')
            st.success("✅ Camp Legal cargado")
        
        if archivo_sb is not None:
            procesador.cargar_archivo(archivo_sb, 'smokeball')
            st.success("✅ Smokeball cargado")
        
        if archivo_tg is not None:
            procesador.cargar_archivo(archivo_tg, 'toggl')
            st.success("✅ Toggl cargado")
        
        # Construir mapa de nombres
        if not procesador.construir_mapa_nombres():
            st.error("❌ Error al construir mapa de nombres")
            st.stop()
        
        # Procesar novedades
        if not procesador.procesar_novedades():
            st.error("❌ Error al procesar novedades")
            st.stop()
        
        # Verificar usuarios
        if not procesador.usuarios_novedades:
            st.warning("⚠️ No hay usuarios en Novedades")
            st.stop()
        
        # Consolidar
        if not procesador.consolidar():
            st.error("❌ Error al consolidar datos")
            st.stop()
        
        # Resultados
        df_resultados = procesador.obtener_resultados()
        estadisticas = procesador.get_estadisticas()
        
        if df_resultados is None or df_resultados.empty:
            st.warning("⚠️ No se encontraron resultados")
            st.stop()
        
        # ============================================================
        # MOSTRAR RESULTADOS
        # ============================================================
        
        st.markdown("---")
        st.markdown("### 📊 Resultados del Reporte")
        
        # KPIs
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("👥 Usuarios", estadisticas['total_usuarios'])
        
        with col2:
            st.metric("⏱️ Total Horas", f"{estadisticas['total_horas']:.1f}h")
        
        with col3:
            st.metric("📊 Promedio", f"{estadisticas['promedio']:.1f}h")
        
        with col4:
            completos = len(df_resultados[df_resultados['Estado'].str.startswith('✅')])
            st.metric("✅ Completos", completos)
        
        # Plataformas
        st.markdown("### 📈 Distribución por Plataforma")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("🏛️ Camp Legal", f"{estadisticas['horas_camp']:.1f}h")
        
        with col2:
            st.metric("📋 Smokeball", f"{estadisticas['horas_sb']:.1f}h")
        
        with col3:
            st.metric("⏱️ Toggl", f"{estadisticas['horas_tg']:.1f}h")
        
        # Tabla
        st.markdown("### 👥 Detalle por Usuario")
        st.dataframe(df_resultados, use_container_width=True, hide_index=True)
        
        # Descargar
        st.markdown("### 📥 Exportar")
        
        csv = df_resultados.to_csv(index=False).encode('utf-8')
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
