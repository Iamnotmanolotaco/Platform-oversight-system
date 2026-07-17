# app.py - Reporte de Tiempos para Streamlit (Versión Final)

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
# FUNCIONES DE CONVERSIÓN (DE TU CÓDIGO ORIGINAL)
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
    """Determina si una fecha es festivo"""
    festivos = [
        (1, 1),   # Año Nuevo
        (5, 1),   # Día del Trabajo
        (7, 20),  # Día de la Independencia
        (8, 7),   # Batalla de Boyacá
        (12, 8),  # Día de la Inmaculada Concepción
        (12, 25), # Navidad
    ]
    return (fecha.month, fecha.day) in festivos

def get_jornada_esperada_por_dia(fecha):
    """Retorna la jornada esperada para un día específico"""
    if fecha.weekday() == 5:  # Sábado
        return 4.0
    elif fecha.weekday() == 6:  # Domingo
        return 0.0
    elif es_festivo(fecha):
        return 0.0
    else:
        return 8.0

# ============================================================
# CONFIGURACIÓN DE COLUMNAS (DE TU CÓDIGO ORIGINAL)
# ============================================================

COLUMNAS_MAPEO = {
    'camp_legal': {
        'archivo': 'Reporte Diario Camp Legal.xlsx',
        'hoja_datos': 'Time entries',
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
        'hoja_datos': 'Entries',
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
        'hoja_datos': 'DataBaseToggl',
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
        'hoja_datos': 'Names',
        'columnas': {
            'nombre': 'NAME CORRECT',
            'nombre_cl': 'NAME CL',
            'nombre_sb': 'NAME SB',
            'nombre_tg': 'NAME TG',
            'status': 'USER STATUS'
        }
    },
    'novedades_max': {
        'archivo': 'Template_Novedades_RRHH_MAX 1 1 2.xlsx',
        'hoja_datos': 'Novedades',
        'columnas': {
            'nombre': 'Persona',
            'fecha_inicio': 'Fecha Inicio',
            'fecha_fin': 'Fecha Fin',
            'tipo': 'Tipo de Novedad'
        },
        'formato_fecha': '%m/%d/%Y'
    },
    'novedades_max_2': {
        'archivo': 'Template_Novedades_RRHH_MAX 1 1 2.xlsx',
        'hoja_datos': 'Novedades 2',
        'columnas': {
            'nombre': 'Persona',
            'fecha': 'Fecha',
            'tipo': 'Tipo de Novedad'
        },
        'formato_fecha': '%m/%d/%Y'
    },
    'novedades_clg': {
        'archivo': 'Template_Novedades_RRHH_CLG - last.xlsx',
        'hoja_datos': 'Novedades',
        'columnas': {
            'nombre': 'Persona',
            'fecha_inicio': 'Fecha Inicio',
            'fecha_fin': 'Fecha Fin',
            'tipo': 'Tipo de Novedad'
        },
        'formato_fecha': '%m/%d/%Y'
    }
}

# ============================================================
# CLASE PROCESADOR (TU CÓDIGO ORIGINAL ADAPTADO PARA STREAMLIT)
# ============================================================

class ReporteTiemposSystem:
    def __init__(self, fecha_inicio, fecha_fin):
        self.fecha_inicio = fecha_inicio
        self.fecha_fin = fecha_fin
        self.jornada_esperada = 8.0
        
        self.df_camp = None
        self.df_smokeball = None
        self.df_toggl = None
        self.df_powerbi = None
        self.df_novedades_max = None
        self.df_novedades_max_2 = None
        self.df_novedades_clg = None
        self.df_analisis = None
        
        self.mapa_nombres = {}
        self.usuarios_con_plataforma = []
        self.usuarios_novedades_2 = set()
        
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
    
    def _obtener_dias_con_jornada(self):
        dias = []
        fecha_actual = self.fecha_inicio
        while fecha_actual <= self.fecha_fin:
            jornada = get_jornada_esperada_por_dia(fecha_actual)
            if jornada > 0:
                dias.append((fecha_actual, jornada))
            fecha_actual += timedelta(days=1)
        return dias
    
    def cargar_archivo(self, archivo_bytes, key):
        """Carga un archivo desde bytes (Streamlit)"""
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
            elif key == 'novedades_max':
                self.df_novedades_max = df
            elif key == 'novedades_max_2':
                self.df_novedades_max_2 = df
            elif key == 'novedades_clg':
                self.df_novedades_clg = df
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
        
        cols = COLUMNAS_MAPEO['powerbi']['columnas']
        col_canonico = cols.get('nombre', 'NAME CORRECT')
        col_cl = cols.get('nombre_cl', 'NAME CL')
        col_sb = cols.get('nombre_sb', 'NAME SB')
        col_tg = cols.get('nombre_tg', 'NAME TG')
        col_status = cols.get('status', 'USER STATUS')
        
        if col_status in self.df_powerbi.columns:
            df_activos = self.df_powerbi[self.df_powerbi[col_status] == 'Active'].copy()
        else:
            df_activos = self.df_powerbi.copy()
        
        self.mapa_nombres = {}
        self.usuarios_con_plataforma = []
        
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
                    self.mapa_nombres[valor] = nombre_canonico_limpio
                    self.mapa_nombres[limpiar_nombre(valor)] = nombre_canonico_limpio
            
            if col_sb in df_activos.columns and pd.notna(row[col_sb]):
                valor = str(row[col_sb]).strip()
                if valor and valor.lower() not in ['true', 'false', 'nan', 'none', '']:
                    tiene_sb = True
                    self.mapa_nombres[valor] = nombre_canonico_limpio
                    self.mapa_nombres[limpiar_nombre(valor)] = nombre_canonico_limpio
            
            if col_tg in df_activos.columns and pd.notna(row[col_tg]):
                valor = str(row[col_tg]).strip()
                if valor and valor.lower() not in ['true', 'false', 'nan', 'none', '']:
                    tiene_tg = True
                    self.mapa_nombres[valor] = nombre_canonico_limpio
                    self.mapa_nombres[limpiar_nombre(valor)] = nombre_canonico_limpio
            
            self.mapa_nombres[nombre_canonico] = nombre_canonico_limpio
            self.mapa_nombres[nombre_canonico_limpio] = nombre_canonico_limpio
            
            if tiene_cl or tiene_sb or tiene_tg:
                self.usuarios_con_plataforma.append(nombre_canonico_limpio)
        
        return True
    
    def procesar_novedades(self):
        novedades_list = []
        
        # Novedades MAX
        if self.df_novedades_max is not None:
            df_max = self.df_novedades_max.copy()
            if 'Persona' in df_max.columns and 'Fecha Inicio' in df_max.columns and 'Fecha Fin' in df_max.columns:
                df_max['Usuario_Normalizado'] = df_max['Persona'].apply(self.normalizar_nombre)
                df_max['Fecha_Inicio'] = df_max['Fecha Inicio'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_max['Fecha_Fin'] = df_max['Fecha Fin'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_max_validos = df_max[df_max['Fecha_Inicio'].notna() & df_max['Fecha_Fin'].notna()]
                if 'Tipo de Novedad' in df_max.columns:
                    df_max_validos['Tipo'] = df_max_validos['Tipo de Novedad']
                else:
                    df_max_validos['Tipo'] = 'Permiso MAX'
                novedades_list.append(df_max_validos[['Usuario_Normalizado', 'Fecha_Inicio', 'Fecha_Fin', 'Tipo']])
        
        # Novedades MAX 2
        if self.df_novedades_max_2 is not None:
            df_max2 = self.df_novedades_max_2.copy()
            if 'Persona' in df_max2.columns and 'Fecha' in df_max2.columns:
                df_max2['Fecha_Conv'] = df_max2['Fecha'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_max2_filtrado = df_max2[
                    (df_max2['Fecha_Conv'] >= self.fecha_inicio) & 
                    (df_max2['Fecha_Conv'] <= self.fecha_fin)
                ]
                
                for _, row in df_max2_filtrado.iterrows():
                    nombre = row['Persona']
                    nombre_normalizado = self.normalizar_nombre(nombre)
                    if nombre_normalizado:
                        self.usuarios_novedades_2.add(nombre_normalizado)
                
                df_max2_validos = df_max2_filtrado[df_max2_filtrado['Fecha_Conv'].notna()]
                if 'Tipo de Novedad' in df_max2.columns:
                    df_max2_validos['Tipo'] = df_max2_validos['Tipo de Novedad']
                else:
                    df_max2_validos['Tipo'] = 'Novedad 2'
                df_max2_validos['Fecha_Inicio'] = df_max2_validos['Fecha_Conv']
                df_max2_validos['Fecha_Fin'] = df_max2_validos['Fecha_Conv']
                df_max2_validos['Usuario_Normalizado'] = df_max2_validos['Persona'].apply(self.normalizar_nombre)
                novedades_list.append(df_max2_validos[['Usuario_Normalizado', 'Fecha_Inicio', 'Fecha_Fin', 'Tipo']])
        
        # Novedades CLG
        if self.df_novedades_clg is not None:
            df_clg = self.df_novedades_clg.copy()
            if 'Persona' in df_clg.columns and 'Fecha Inicio' in df_clg.columns and 'Fecha Fin' in df_clg.columns:
                df_clg['Usuario_Normalizado'] = df_clg['Persona'].apply(self.normalizar_nombre)
                df_clg['Fecha_Inicio'] = df_clg['Fecha Inicio'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_clg['Fecha_Fin'] = df_clg['Fecha Fin'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_clg_validos = df_clg[df_clg['Fecha_Inicio'].notna() & df_clg['Fecha_Fin'].notna()]
                if 'Tipo de Novedad' in df_clg.columns:
                    df_clg_validos['Tipo'] = df_clg_validos['Tipo de Novedad']
                else:
                    df_clg_validos['Tipo'] = 'Permiso CLG'
                novedades_list.append(df_clg_validos[['Usuario_Normalizado', 'Fecha_Inicio', 'Fecha_Fin', 'Tipo']])
        
        if novedades_list:
            self.df_novedades_combinadas = pd.concat(novedades_list, ignore_index=True)
        else:
            self.df_novedades_combinadas = None
        
        return True
    
    def verificar_permiso(self, usuario, fecha):
        if self.df_novedades_combinadas is None:
            return None
        novedades_usuario = self.df_novedades_combinadas[
            self.df_novedades_combinadas['Usuario_Normalizado'] == usuario
        ]
        for _, row in novedades_usuario.iterrows():
            if row['Fecha_Inicio'] <= fecha <= row['Fecha_Fin']:
                return row['Tipo']
        return None
    
    def procesar_plataforma(self, df, config_key):
        if df is None:
            return None
        
        config = COLUMNAS_MAPEO[config_key]
        cols = config['columnas']
        formato_horas = config.get('formato_horas', 'auto')
        formato_fecha = config.get('formato_fecha', '%m/%d/%Y')
        
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
            df_proc['Date'] = df_proc[col_fecha].apply(lambda x: convertir_fecha(x, formato_fecha))
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
        
        df_dias = df_proc.groupby('Usuario_Normalizado')['Date'].nunique().reset_index()
        df_dias.columns = ['Usuario_Normalizado', 'Dias_Activos']
        df_agrupado = df_agrupado.merge(df_dias, on='Usuario_Normalizado', how='left')
        df_agrupado['Dias_Activos'] = df_agrupado['Dias_Activos'].fillna(0).astype(int)
        
        return df_agrupado
    
    def consolidar(self):
        if not self.usuarios_novedades_2:
            self.df_analisis = pd.DataFrame()
            return True
        
        df_camp = self.procesar_plataforma(self.df_camp, 'camp_legal')
        df_sb = self.procesar_plataforma(self.df_smokeball, 'smokeball')
        df_tg = self.procesar_plataforma(self.df_toggl, 'toggl')
        
        usuarios_dict = {}
        for usuario in self.usuarios_novedades_2:
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
        
        for usuario in usuarios_dict:
            fecha_actual = self.fecha_inicio
            while fecha_actual <= self.fecha_fin:
                permiso = self.verificar_permiso(usuario, fecha_actual)
                if permiso:
                    usuarios_dict[usuario]['Permiso'] = permiso
                    break
                fecha_actual += timedelta(days=1)
        
        datos = []
        for usuario, data in usuarios_dict.items():
            total = data['Camp Legal'] + data['Smokeball'] + data['Toggl']
            porcentaje = (total / self.jornada_total_esperada * 100) if self.jornada_total_esperada > 0 else 0
            
            if data['Permiso']:
                estado = f"📋 {data['Permiso']}"
            else:
                estado = self.calcular_estado(total, data['Camp Legal'], data['Smokeball'], data['Toggl'], porcentaje)
            
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
        
        self.df_analisis = pd.DataFrame(datos)
        return True
    
    def calcular_estado(self, total, camp, sb, tg, porcentaje):
        if total == 0:
            return "⛔ Sin registro"
        
        plataformas = []
        if camp > 0:
            plataformas.append('CL')
        if sb > 0:
            plataformas.append('SB')
        if tg > 0:
            plataformas.append('TG')
        
        if len(plataformas) == 0:
            return "⛔ Sin registro"
        elif len(plataformas) == 1:
            return f"⚠️ Solo {plataformas[0]} ({porcentaje:.0f}%)"
        elif len(plataformas) == 2:
            return f"⚠️ {', '.join(plataformas)} ({porcentaje:.0f}%)"
        else:
            if porcentaje >= 90:
                return f"✅ Completo ({porcentaje:.0f}%)"
            elif porcentaje >= 70:
                return f"⚠️ Parcial ({porcentaje:.0f}%)"
            else:
                return f"❌ Insuficiente ({porcentaje:.0f}%)"
    
    def obtener_resultados(self):
        return self.df_analisis
    
    def get_estadisticas(self):
        if self.df_analisis is None or self.df_analisis.empty:
            return {
                'total_usuarios': 0,
                'total_horas': 0,
                'promedio': 0,
                'con_permiso': 0,
                'horas_camp': 0,
                'horas_sb': 0,
                'horas_tg': 0
            }
        
        df = self.df_analisis
        return {
            'total_usuarios': len(df),
            'total_horas': df['Total_Horas'].sum(),
            'promedio': df['Total_Horas'].mean(),
            'con_permiso': len(df[df['Permiso'] != 'Sin permiso']),
            'horas_camp': df['Camp Legal'].sum(),
            'horas_sb': df['Smokeball'].sum(),
            'horas_tg': df['Toggl'].sum()
        }

# ============================================================
# INTERFAZ DE USUARIO - STREAMLIT
# ============================================================

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a3a5c 0%, #2c5f8a 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(26,58,92,0.3);
    }
    .main-header h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 700;
    }
    .main-header p {
        margin: 0.3rem 0 0 0;
        opacity: 0.85;
        font-size: 1rem;
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
        letter-spacing: 0.5px;
    }
    .info-box {
        background: #fef9e7;
        border: 1px solid #f9e79f;
        padding: 12px 16px;
        border-radius: 10px;
        margin: 10px 0;
        color: #7d6608;
    }
    .success-box {
        background: #eafaf1;
        border: 1px solid #a9dfbf;
        padding: 12px 16px;
        border-radius: 10px;
        color: #1a7a42;
    }
    .upload-area {
        border: 2px dashed #bdc3c7;
        border-radius: 10px;
        padding: 0.5rem;
        margin: 0.3rem 0;
        background: #fafbfc;
    }
    .stButton > button {
        background-color: #1a3a5c;
        color: white;
        font-weight: 600;
    }
    .stButton > button:hover {
        background-color: #2c5f8a;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1>📊 Reporte de Tiempos</h1>
    <p>Análisis de tiempos por plataforma · Control de calidad · Novedades 2</p>
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
    st.subheader("⚡ Acceso rápido")
    
    if st.button("📅 Ayer", use_container_width=True):
        fecha_inicio = datetime.now().date() - timedelta(days=1)
        fecha_fin = fecha_inicio
        st.rerun()
    
    if st.button("📅 Semana actual", use_container_width=True):
        hoy = datetime.now().date()
        dias_desde_lunes = hoy.weekday()
        fecha_inicio = hoy - timedelta(days=dias_desde_lunes)
        fecha_fin = hoy
        st.rerun()
    
    if st.button("📅 Mes actual", use_container_width=True):
        hoy = datetime.now().date()
        fecha_inicio = hoy.replace(day=1)
        fecha_fin = hoy
        st.rerun()
    
    st.divider()
    st.subheader("📁 Archivos")
    st.caption("Sube los archivos Excel necesarios")
    
    uploaded_files = {}
    
    # Power BI
    with st.expander("📊 Power BI resources.xlsx", expanded=True):
        uploaded_files['powerbi'] = st.file_uploader(
            "Power BI resources.xlsx (Names)",
            type=['xlsx'],
            key="powerbi",
            label_visibility="collapsed"
        )
    
    # Novedades
    with st.expander("📋 Template_Novedades_RRHH_MAX", expanded=True):
        uploaded_files['novedades_max'] = st.file_uploader(
            "Novedades (MAX)",
            type=['xlsx'],
            key="novedades_max",
            label_visibility="collapsed"
        )
        uploaded_files['novedades_max_2'] = st.file_uploader(
            "Novedades 2 (MAX) - OBLIGATORIO",
            type=['xlsx'],
            key="novedades_max_2",
            label_visibility="collapsed"
        )
    
    # Plataformas
    with st.expander("🏛️ Camp Legal", expanded=False):
        uploaded_files['camp_legal'] = st.file_uploader(
            "Reporte Diario Camp Legal.xlsx",
            type=['xlsx'],
            key="camp_legal",
            label_visibility="collapsed"
        )
    
    with st.expander("📋 Smokeball", expanded=False):
        uploaded_files['smokeball'] = st.file_uploader(
            "Reporte_general.xlsx",
            type=['xlsx'],
            key="smokeball",
            label_visibility="collapsed"
        )
    
    with st.expander("⏱️ Toggl", expanded=False):
        uploaded_files['toggl'] = st.file_uploader(
            "Revision de entradas de tiempo - Toggl.xlsx",
            type=['xlsx'],
            key="toggl",
            label_visibility="collapsed"
        )
    
    # Novedades CLG
    with st.expander("📋 Template_Novedades_RRHH_CLG", expanded=False):
        uploaded_files['novedades_clg'] = st.file_uploader(
            "Novedades CLG",
            type=['xlsx'],
            key="novedades_clg",
            label_visibility="collapsed"
        )
    
    st.divider()
    
    archivos_cargados = sum([1 for v in uploaded_files.values() if v is not None])
    
    # Verificar archivos obligatorios
    obligatorios = ['powerbi', 'novedades_max_2']
    obligatorios_cargados = sum([1 for k in obligatorios if uploaded_files.get(k) is not None])
    
    st.markdown(f"**📁 Archivos cargados:** {archivos_cargados}/7")
    st.markdown(f"**✅ Obligatorios:** {obligatorios_cargados}/2 (Power BI + Novedades 2)")
    
    procesar = st.button(
        "🚀 Generar Reporte",
        type="primary",
        use_container_width=True,
        disabled=obligatorios_cargados < 2
    )

# Área principal
if not procesar:
    st.info("👈 Sube los archivos requeridos (Power BI y Novedades 2) y presiona 'Generar Reporte'")
    
    st.markdown("### 📋 Archivos requeridos")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        **✅ Obligatorios:**
        - 📊 Power BI resources.xlsx
        - 📋 Novedades 2 (MAX)
        """)
    with col2:
        st.markdown("""
        **📌 Plataformas (al menos una):**
        - 🏛️ Camp Legal
        - 📋 Smokeball
        - ⏱️ Toggl
        """)
    with col3:
        st.markdown("""
        **📋 Opcionales:**
        - Novedades MAX
        - Novedades CLG
        """)
    
    st.markdown("---")
    st.markdown("""
    <div class="info-box">
        <strong>💡 Nota:</strong> El reporte solo mostrará usuarios que aparecen en <strong>Novedades 2</strong>.
        Los usuarios que no están en Novedades 2 están descansando y no aparecerán.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Procesar reporte
with st.spinner("🔄 Procesando datos... Por favor espera"):
    try:
        # Inicializar procesador
        procesador = ReporteTiemposSystem(fecha_inicio, fecha_fin)
        
        # Cargar archivos
        for key, file in uploaded_files.items():
            if file is not None:
                procesador.cargar_archivo(file, key)
        
        # Construir mapa de nombres
        if not procesador.construir_mapa_nombres():
            st.error("❌ Error al construir mapa de nombres. Verifica el archivo Power BI Resources.")
            st.stop()
        
        # Procesar novedades
        procesador.procesar_novedades()
        
        # Verificar usuarios en Novedades 2
        if not procesador.usuarios_novedades_2:
            st.warning("⚠️ No hay usuarios en Novedades 2 para el rango seleccionado.")
            st.info("💡 Todos los usuarios están descansando en este período.")
            st.stop()
        
        # Consolidar resultados
        if not procesador.consolidar():
            st.error("❌ Error al consolidar datos.")
            st.stop()
        
        # Obtener resultados
        df_resultados = procesador.obtener_resultados()
        estadisticas = procesador.get_estadisticas()
        
        if df_resultados is None or df_resultados.empty:
            st.warning("⚠️ No se encontraron resultados para el rango seleccionado.")
            st.stop()
        
        # ============================================================
        # MOSTRAR RESULTADOS
        # ============================================================
        
        st.markdown("---")
        st.markdown("### 📊 Resultados del Reporte")
        
        # Información de Novedades 2
        st.markdown(f"""
        <div class="info-box">
            <strong>📋 Usuarios en Novedades 2:</strong> {len(procesador.usuarios_novedades_2)} personas deben marcar tiempo
            <span style="margin-left:20px;">⛔ <strong>Usuarios NO en Novedades 2:</strong> Están descansando y no aparecen en el reporte</span>
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
        
        # Información adicional
        st.markdown("---")
        st.markdown("### 📋 Información Adicional")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"""
            <div class="info-box">
                <strong>📊 Jornada total esperada:</strong> {procesador.jornada_total_esperada:.0f}h<br>
                <strong>📅 Días hábiles:</strong> {procesador.dias_habiles}<br>
                <strong>📋 Usuarios en Novedades 2:</strong> {len(procesador.usuarios_novedades_2)}
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="success-box">
                <strong>✅ Resumen:</strong><br>
                • {estadisticas['total_usuarios']} usuarios con turno<br>
                • {estadisticas['total_horas']:.1f} horas totales<br>
                • Promedio de {estadisticas['promedio']:.1f} horas por usuario
            </div>
            """, unsafe_allow_html=True)
        
        # Descargar resultados
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
        st.error(f"❌ Error al procesar el reporte: {e}")
        st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #95a5a6; font-size: 12px; padding: 20px 0;">
    Reporte generado automáticamente · Datos de Camp Legal, Smokeball y Toggl
</div>
""", unsafe_allow_html=True)
