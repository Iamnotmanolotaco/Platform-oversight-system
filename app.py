# app.py - REPORTE DE TIEMPOS CON DIAGNÓSTICO COMPLETO

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
        'Legal Agent Support', 'Customer Support Agent',
        'Dpto Mail & Records', 'Leads Team'
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
    
    nombre = re.sub(r'^\d+\s+', '', nombre)
    nombre = re.sub(r'\([^)]*\)', '', nombre).strip()
    nombre = re.sub(r'\[[^\]]*\]', '', nombre).strip()
    nombre = re.sub(r'\s+', ' ', nombre)
    
    return nombre.strip()

def normalizar_nombre_flexible(nombre):
    """Limpia y normaliza un nombre, incluyendo variaciones ortográficas comunes"""
    if not isinstance(nombre, str):
        return nombre
    
    nombre_limpio = limpiar_nombre(nombre.strip())
    
    # Reemplazar variaciones comunes
    reemplazos = {
        'medinha': 'medina',
        'cristhian': 'cristian',
    }
    
    nombre_lower = nombre_limpio.lower()
    for incorrecto, correcto in reemplazos.items():
        if incorrecto in nombre_lower:
            nombre_limpio = nombre_limpio.replace(incorrecto, correcto)
    
    return nombre_limpio

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
    festivos = [(1, 1), (5, 1), (7, 20), (8, 7), (12, 8), (12, 25)]
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
# MAPEO DE COLUMNAS
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
            'status': 'USER STATUS',
            'company': 'COMPANY'
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
# CLASE PRINCIPAL
# ============================================================

class ReporteTiemposSystem:
    def __init__(self, fecha_inicio=None, fecha_fin=None):
        self.today = datetime.now().date()
        
        if fecha_inicio is None:
            self.fecha_inicio = self.today - timedelta(days=1)
        else:
            self.fecha_inicio = fecha_inicio
            
        if fecha_fin is None:
            self.fecha_fin = self.fecha_inicio
        else:
            self.fecha_fin = fecha_fin
        
        if self.fecha_inicio > self.fecha_fin:
            self.fecha_inicio, self.fecha_fin = self.fecha_fin, self.fecha_inicio
        
        self.fecha_inicio_str = self.fecha_inicio.strftime('%Y%m%d')
        self.fecha_fin_str = self.fecha_fin.strftime('%Y%m%d')
        self.ruta_base = ""
        self.jornada_esperada = 8.0
        
        self.df_camp = None
        self.df_smokeball = None
        self.df_toggl = None
        self.df_powerbi = None
        self.df_novedades_max = None
        self.df_novedades_max_2 = None
        self.df_novedades_clg = None
        self.df_novedades_combinadas = None
        self.df_analisis = None
        self.df_detalle_diario = None
        
        self.mapa_nombres = {}
        self.mapa_compania = {}
        self.usuarios_con_plataforma = []
        self.usuarios_novedades_2 = set()
        self.df_permisos_por_dia = None
        self.usuarios_manuales = set()
        
        self.dias_totales = (self.fecha_fin - self.fecha_inicio).days + 1
        self.dias_habiles = self._calcular_dias_habiles()
        self.jornada_total_esperada = self._calcular_jornada_total()
        
        print(f"\n📅 Rango: {self.fecha_inicio.strftime('%d/%m/%Y')} - {self.fecha_fin.strftime('%d/%m/%Y')}")
        print(f"📊 Días totales: {self.dias_totales} | Días hábiles: {self.dias_habiles}")
        print(f"📊 Jornada total esperada: {self.jornada_total_esperada:.1f}h")
    
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
    
    def cargar_archivo(self, archivo_bytes, key, sheet_name=None):
        try:
            if sheet_name:
                df = pd.read_excel(archivo_bytes, sheet_name=sheet_name)
            else:
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
        
        nombre_limpio = normalizar_nombre_flexible(nombre.strip())
        if not nombre_limpio:
            return nombre_limpio
        
        # Buscar coincidencia exacta
        if nombre_limpio in self.mapa_nombres:
            return self.mapa_nombres[nombre_limpio]
        
        if nombre.strip() in self.mapa_nombres:
            return self.mapa_nombres[nombre.strip()]
        
        # Buscar coincidencia parcial (más flexible)
        for nombre_plat, nombre_canon in self.mapa_nombres.items():
            # Si el nombre limpio está contenido en el nombre de la plataforma
            if nombre_limpio.lower() in nombre_plat.lower():
                return nombre_canon
            # Si el nombre de la plataforma está contenido en el nombre limpio
            if nombre_plat.lower() in nombre_limpio.lower():
                return nombre_canon
            # Si comparten la primera palabra (nombre)
            if ' ' in nombre_limpio and ' ' in nombre_plat:
                if nombre_limpio.split()[0].lower() == nombre_plat.split()[0].lower():
                    return nombre_canon
                if nombre_limpio.split()[-1].lower() == nombre_plat.split()[-1].lower():
                    return nombre_canon
        
        return nombre_limpio
    
    def obtener_compania(self, nombre):
        nombre_limpio = limpiar_nombre(nombre)
        
        if nombre_limpio in self.mapa_compania:
            return self.mapa_compania[nombre_limpio]
        
        for nombre_map, compania in self.mapa_compania.items():
            if nombre_limpio.lower() in nombre_map.lower() or nombre_map.lower() in nombre_limpio.lower():
                return compania
        
        return 'Sin compañía'
    
    def construir_mapa_nombres(self):
        if self.df_powerbi is None:
            st.error("❌ Power BI no está cargado")
            return False
        
        cols = COLUMNAS_MAPEO['powerbi']['columnas']
        col_canonico = cols.get('nombre', 'NAME CORRECT')
        col_cl = cols.get('nombre_cl', 'NAME CL')
        col_sb = cols.get('nombre_sb', 'NAME SB')
        col_tg = cols.get('nombre_tg', 'NAME TG')
        col_status = cols.get('status', 'USER STATUS')
        col_company = cols.get('company', 'COMPANY')
        
        if col_company not in self.df_powerbi.columns:
            for col in self.df_powerbi.columns:
                if 'COMPANY' in col.upper() or 'EMPRESA' in col.upper():
                    col_company = col
                    break
        
        if col_status in self.df_powerbi.columns:
            df_activos = self.df_powerbi[self.df_powerbi[col_status] == 'Active'].copy()
        else:
            df_activos = self.df_powerbi.copy()
        
        self.mapa_nombres = {}
        self.mapa_compania = {}
        self.usuarios_con_plataforma = []
        
        for idx, row in df_activos.iterrows():
            nombre_canonico = str(row[col_canonico]).strip() if pd.notna(row[col_canonico]) else None
            if not nombre_canonico:
                continue
            
            nombre_canonico_limpio = normalizar_nombre_flexible(nombre_canonico)
            
            if col_company in df_activos.columns and pd.notna(row[col_company]):
                compania = str(row[col_company]).strip()
                if compania:
                    self.mapa_compania[nombre_canonico_limpio] = compania
                    self.mapa_compania[nombre_canonico] = compania
                    partes = nombre_canonico_limpio.split()
                    if len(partes) >= 2:
                        nombre_corto = f"{partes[0]} {partes[-1]}"
                        self.mapa_compania[nombre_corto] = compania
                    if len(partes) >= 1:
                        self.mapa_compania[partes[0]] = compania
                        self.mapa_compania[partes[-1]] = compania
            
            # NAME CL
            if col_cl in df_activos.columns and pd.notna(row[col_cl]):
                valor = str(row[col_cl]).strip()
                if valor and valor.lower() not in ['true', 'false', 'nan', 'none', '']:
                    valor_limpio = normalizar_nombre_flexible(valor)
                    self.mapa_nombres[valor] = nombre_canonico_limpio
                    self.mapa_nombres[valor_limpio] = nombre_canonico_limpio
                    partes = valor_limpio.split()
                    if len(partes) >= 2:
                        nombre_corto = f"{partes[0]} {partes[-1]}"
                        self.mapa_nombres[nombre_corto] = nombre_canonico_limpio
                    if len(partes) >= 1:
                        self.mapa_nombres[partes[0]] = nombre_canonico_limpio
                        self.mapa_nombres[partes[-1]] = nombre_canonico_limpio
                    if col_company in df_activos.columns and pd.notna(row[col_company]):
                        compania = str(row[col_company]).strip()
                        if compania:
                            self.mapa_compania[valor_limpio] = compania
            
            # NAME SB
            if col_sb in df_activos.columns and pd.notna(row[col_sb]):
                valor = str(row[col_sb]).strip()
                if valor and valor.lower() not in ['true', 'false', 'nan', 'none', '']:
                    valor_limpio = normalizar_nombre_flexible(valor)
                    self.mapa_nombres[valor] = nombre_canonico_limpio
                    self.mapa_nombres[valor_limpio] = nombre_canonico_limpio
                    partes = valor_limpio.split()
                    if len(partes) >= 2:
                        nombre_corto = f"{partes[0]} {partes[-1]}"
                        self.mapa_nombres[nombre_corto] = nombre_canonico_limpio
                    if len(partes) >= 1:
                        self.mapa_nombres[partes[0]] = nombre_canonico_limpio
                        self.mapa_nombres[partes[-1]] = nombre_canonico_limpio
                    if col_company in df_activos.columns and pd.notna(row[col_company]):
                        compania = str(row[col_company]).strip()
                        if compania:
                            self.mapa_compania[valor_limpio] = compania
            
            # NAME TG
            if col_tg in df_activos.columns and pd.notna(row[col_tg]):
                valor = str(row[col_tg]).strip()
                if valor and valor.lower() not in ['true', 'false', 'nan', 'none', '']:
                    valor_limpio = normalizar_nombre_flexible(valor)
                    self.mapa_nombres[valor] = nombre_canonico_limpio
                    self.mapa_nombres[valor_limpio] = nombre_canonico_limpio
                    partes = valor_limpio.split()
                    if len(partes) >= 2:
                        nombre_corto = f"{partes[0]} {partes[-1]}"
                        self.mapa_nombres[nombre_corto] = nombre_canonico_limpio
                    if len(partes) >= 1:
                        self.mapa_nombres[partes[0]] = nombre_canonico_limpio
                        self.mapa_nombres[partes[-1]] = nombre_canonico_limpio
                    if col_company in df_activos.columns and pd.notna(row[col_company]):
                        compania = str(row[col_company]).strip()
                        if compania:
                            self.mapa_compania[valor_limpio] = compania
            
            self.mapa_nombres[nombre_canonico] = nombre_canonico_limpio
            self.mapa_nombres[nombre_canonico_limpio] = nombre_canonico_limpio
            
            partes = nombre_canonico_limpio.split()
            if len(partes) >= 2:
                nombre_corto = f"{partes[0]} {partes[-1]}"
                self.mapa_nombres[nombre_corto] = nombre_canonico_limpio
            if len(partes) >= 1:
                self.mapa_nombres[partes[0]] = nombre_canonico_limpio
                self.mapa_nombres[partes[-1]] = nombre_canonico_limpio
            
            tiene_cl = col_cl in df_activos.columns and pd.notna(row[col_cl]) and str(row[col_cl]).strip() not in ['', 'true', 'false', 'nan', 'none']
            tiene_sb = col_sb in df_activos.columns and pd.notna(row[col_sb]) and str(row[col_sb]).strip() not in ['', 'true', 'false', 'nan', 'none']
            tiene_tg = col_tg in df_activos.columns and pd.notna(row[col_tg]) and str(row[col_tg]).strip() not in ['', 'true', 'false', 'nan', 'none']
            
            if tiene_cl or tiene_sb or tiene_tg:
                self.usuarios_con_plataforma.append(nombre_canonico_limpio)
        
        # ============================================================
        # AGREGAR TODOS LOS USUARIOS DE TOGGL AUTOMÁTICAMENTE
        # ============================================================
        
        if self.df_toggl is not None and 'Member' in self.df_toggl.columns:
            st.write("### 🔧 Agregando usuarios de Toggl automáticamente")
            
            usuarios_toggl = self.df_toggl['Member'].dropna().unique()
            st.write(f"**Usuarios encontrados en Toggl:** {len(usuarios_toggl)}")
            
            agregados = 0
            for usuario in usuarios_toggl:
                if isinstance(usuario, str):
                    usuario_limpio = normalizar_nombre_flexible(usuario.strip())
                    if usuario_limpio:
                        if usuario_limpio not in self.usuarios_con_plataforma:
                            # Buscar si hay un nombre similar en el mapa
                            nombre_canon = None
                            for nombre_plat, nombre_canonico in self.mapa_nombres.items():
                                if (usuario_limpio.lower() in nombre_plat.lower() or 
                                    nombre_plat.lower() in usuario_limpio.lower() or
                                    (usuario_limpio.split()[0].lower() == nombre_plat.split()[0].lower() if ' ' in usuario_limpio and ' ' in nombre_plat else False) or
                                    (usuario_limpio.split()[-1].lower() == nombre_plat.split()[-1].lower() if ' ' in usuario_limpio and ' ' in nombre_plat else False)):
                                    nombre_canon = nombre_canonico
                                    break
                            
                            if nombre_canon:
                                self.mapa_nombres[usuario_limpio] = nombre_canon
                                self.mapa_nombres[usuario] = nombre_canon
                                self.usuarios_con_plataforma.append(nombre_canon)
                            else:
                                self.mapa_nombres[usuario_limpio] = usuario_limpio
                                self.mapa_nombres[usuario] = usuario_limpio
                                self.usuarios_con_plataforma.append(usuario_limpio)
                                self.usuarios_manuales.add(usuario_limpio)
                            
                            agregados += 1
                            st.write(f"   ✅ Agregado: {usuario} → {usuario_limpio}")
            
            st.write(f"**Total usuarios agregados de Toggl:** {agregados}")
            st.write(f"**Total usuarios en plataforma:** {len(self.usuarios_con_plataforma)}")
        else:
            st.warning("⚠️ Toggl no está cargado o no tiene la columna 'Member'")
        
        return True
    
    def procesar_novedades(self):
        print("\n" + "="*70)
        print("📋 PROCESANDO NOVEDADES")
        print("="*70)
        
        novedades_list = []
        permisos_por_dia = []
        
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
                print(f"   ✅ MAX (rango - permisos): {len(df_max_validos)} registros")
        
        if self.df_novedades_max_2 is not None:
            df_max2 = self.df_novedades_max_2.copy()
            if 'Persona' in df_max2.columns and 'Fecha' in df_max2.columns:
                df_max2['Fecha_Conv'] = df_max2['Fecha'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                
                if len(df_max2) > 0:
                    df_max2_filtrado = df_max2[
                        (df_max2['Fecha_Conv'] >= self.fecha_inicio) & 
                        (df_max2['Fecha_Conv'] <= self.fecha_fin)
                    ]
                else:
                    df_max2_filtrado = df_max2
                
                for _, row in df_max2_filtrado.iterrows():
                    nombre = row['Persona']
                    nombre_normalizado = self.normalizar_nombre(nombre)
                    fecha = row['Fecha_Conv']
                    tipo = row['Tipo de Novedad'] if 'Tipo de Novedad' in row else 'Trabajo en festivo'
                    
                    if nombre_normalizado and pd.notna(fecha):
                        if es_festivo(fecha):
                            self.usuarios_novedades_2.add(nombre_normalizado)
                            permisos_por_dia.append({
                                'Usuario': nombre_normalizado,
                                'Fecha': fecha,
                                'Tipo': tipo,
                                'Es_Novedad_2': True
                            })
                            print(f"   📌 {nombre_normalizado} debe trabajar en festivo: {fecha}")
                
                print(f"   ✅ MAX 2 (festivos): {len(df_max2_filtrado)} registros en el rango")
                print(f"   ✅ Usuarios que deben trabajar en festivos: {len(self.usuarios_novedades_2)}")
            else:
                print(f"   ⚠️ Columnas 'Persona' o 'Fecha' no encontradas en Novedades 2")
        
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
                print(f"   ✅ CLG (rango - permisos): {len(df_clg_validos)} registros")
        
        if novedades_list:
            self.df_novedades_combinadas = pd.concat(novedades_list, ignore_index=True)
            print(f"   ✅ Total novedades combinadas (rango - permisos): {len(self.df_novedades_combinadas)}")
        else:
            self.df_novedades_combinadas = None
        
        if permisos_por_dia:
            self.df_permisos_por_dia = pd.DataFrame(permisos_por_dia)
            print(f"   ✅ Permisos por día específico (festivos): {len(self.df_permisos_por_dia)}")
        else:
            self.df_permisos_por_dia = None
        
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
    
    def verificar_permiso_dia_especifico(self, usuario, fecha):
        if self.df_permisos_por_dia is None:
            return None
        permisos = self.df_permisos_por_dia[
            (self.df_permisos_por_dia['Usuario'] == usuario) & 
            (self.df_permisos_por_dia['Fecha'] == fecha)
        ]
        if len(permisos) > 0:
            return permisos.iloc[0]['Tipo']
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
        
        df_diario = df_proc.groupby(['Usuario_Normalizado', 'Date']).agg({
            'Horas': 'sum'
        }).reset_index()
        
        df_diario_pivot = df_diario.pivot(index='Usuario_Normalizado', columns='Date', values='Horas').fillna(0)
        df_diario_pivot.columns = [f'Dia_{c.strftime("%d/%m")}' for c in df_diario_pivot.columns]
        df_diario_pivot = df_diario_pivot.reset_index()
        
        df_agrupado = df_agrupado.merge(df_diario_pivot, on='Usuario_Normalizado', how='left')
        
        return df_agrupado
    
    def consolidar_todas_plataformas(self):
        print("\n" + "="*70)
        print("🔄 CONSOLIDANDO DATOS")
        print("="*70)
        print(f"📅 Rango: {self.fecha_inicio.strftime('%d/%m/%Y')} - {self.fecha_fin.strftime('%d/%m/%Y')}")
        print(f"📊 Jornada total esperada: {self.jornada_total_esperada:.1f}h")
        print("-"*70)
        
        if not self.usuarios_con_plataforma:
            print("⚠️ No hay usuarios con plataforma. El reporte estará vacío.")
            self.df_analisis = pd.DataFrame()
            return True
        
        print(f"📋 Usuarios con plataforma: {len(self.usuarios_con_plataforma)}")
        
        df_camp = self.procesar_plataforma(self.df_camp, 'camp_legal')
        df_sb = self.procesar_plataforma(self.df_smokeball, 'smokeball')
        df_tg = self.procesar_plataforma(self.df_toggl, 'toggl')
        
        dias_con_jornada = self._obtener_dias_con_jornada()
        dias_columnas = [f'Dia_{fecha.strftime("%d/%m")}' for fecha, _ in dias_con_jornada]
        
        usuarios_dict = {}
        
        for usuario in self.usuarios_con_plataforma:
            compania = self.obtener_compania(usuario)
            usuarios_dict[usuario] = {
                'Camp Legal': 0.0,
                'Smokeball': 0.0,
                'Toggl': 0.0,
                'Dias_Activos': 0,
                'Actividades': [],
                'Permiso': None,
                'Permiso_Dia_Especifico': {},
                'Novedad_2': 'No',
                'Compañia': compania,
                'Detalle_Diario': {dia: 0.0 for dia in dias_columnas},
                'Jornada_Diaria': {dia: jornada for dia, jornada in [(f'Dia_{fecha.strftime("%d/%m")}', jornada) for fecha, jornada in dias_con_jornada]}
            }
        
        for usuario in self.usuarios_novedades_2:
            if usuario in usuarios_dict:
                usuarios_dict[usuario]['Novedad_2'] = 'Sí'
        
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
                    
                    for col in dias_columnas:
                        if col in row and row[col] > 0:
                            usuarios_dict[usuario]['Detalle_Diario'][col] += row[col]
        
        procesar_plataforma_detalle(df_camp, 'Camp Legal')
        procesar_plataforma_detalle(df_sb, 'Smokeball')
        procesar_plataforma_detalle(df_tg, 'Toggl')
        
        print("   🔍 Verificando permisos por día...")
        for usuario in usuarios_dict:
            fecha_actual = self.fecha_inicio
            while fecha_actual <= self.fecha_fin:
                permiso_rango = self.verificar_permiso(usuario, fecha_actual)
                if permiso_rango:
                    usuarios_dict[usuario]['Permiso'] = permiso_rango
                
                if es_festivo(fecha_actual):
                    permiso_dia = self.verificar_permiso_dia_especifico(usuario, fecha_actual)
                    if permiso_dia:
                        dia_key = f'Dia_{fecha_actual.strftime("%d/%m")}'
                        usuarios_dict[usuario]['Permiso_Dia_Especifico'][dia_key] = permiso_dia
                
                fecha_actual += timedelta(days=1)
        
        datos = []
        for usuario, data in usuarios_dict.items():
            total = data['Camp Legal'] + data['Smokeball'] + data['Toggl']
            porcentaje = (total / self.jornada_total_esperada * 100) if self.jornada_total_esperada > 0 else 0
            
            incumplimiento_diario = []
            dias_incumplidos = []
            
            for dia, horas in data['Detalle_Diario'].items():
                jornada_esperada = data['Jornada_Diaria'].get(dia, 8.0)
                tiene_permiso_dia = dia in data['Permiso_Dia_Especifico']
                es_dia_festivo = False
                
                fecha_str = dia.replace('Dia_', '')
                try:
                    dia_obj = datetime.strptime(fecha_str, '%d/%m').replace(year=self.fecha_inicio.year)
                    es_dia_festivo = es_festivo(dia_obj)
                except:
                    pass
                
                if es_dia_festivo and tiene_permiso_dia:
                    pass
                elif horas < jornada_esperada:
                    try:
                        dia_obj = datetime.strptime(fecha_str, '%d/%m').replace(year=self.fecha_inicio.year)
                        es_sabado = dia_obj.weekday() == 5
                        if es_sabado:
                            incumplimiento_diario.append(f"{fecha_str}: {horas:.1f}h (esperado 4h) ❌")
                            dias_incumplidos.append(dia)
                        else:
                            incumplimiento_diario.append(f"{fecha_str}: {horas:.1f}h (esperado 8h) ❌")
                            dias_incumplidos.append(dia)
                    except:
                        incumplimiento_diario.append(f"{dia}: {horas:.1f}h ❌")
                        dias_incumplidos.append(dia)
                elif es_dia_festivo and horas == 0 and tiene_permiso_dia:
                    incumplimiento_diario.append(f"{fecha_str}: {horas:.1f}h (DEBE trabajar en festivo) 🚨")
                    dias_incumplidos.append(dia)
            
            tiene_incumplimiento = len(dias_incumplidos) > 0
            
            if data['Permiso']:
                estado = f"📋 {data['Permiso']}"
            else:
                estado = self.calcular_estado_rango(total, data['Camp Legal'], data['Smokeball'], data['Toggl'], porcentaje)
            
            detalle_items = []
            for dia, horas in data['Detalle_Diario'].items():
                if horas > 0:
                    jornada_esperada = data['Jornada_Diaria'].get(dia, 8.0)
                    tiene_permiso = dia in data['Permiso_Dia_Especifico']
                    
                    if tiene_permiso:
                        permiso_tipo = data['Permiso_Dia_Especifico'][dia]
                        detalle_items.append(f"{dia}: {horas:.1f}h (Festivo: {permiso_tipo}) 📋")
                    elif horas < jornada_esperada:
                        detalle_items.append(f"{dia}: {horas:.1f}h (esperado {jornada_esperada:.0f}h) ⚠️")
                    else:
                        detalle_items.append(f"{dia}: {horas:.1f}h (esperado {jornada_esperada:.0f}h)")
            detalle_str = ' | '.join(detalle_items)
            
            datos.append({
                'Usuario': usuario,
                'Compañia': data['Compañia'],
                'Camp Legal': round(data['Camp Legal'], 2),
                'Smokeball': round(data['Smokeball'], 2),
                'Toggl': round(data['Toggl'], 2),
                'Total_Horas': round(total, 2),
                'Dias_Activos': data['Dias_Activos'],
                'Actividades': ' | '.join(data['Actividades']) if data['Actividades'] else 'Sin registro',
                'Permiso': data['Permiso'] if data['Permiso'] else 'Sin permiso',
                'Novedad_2': data['Novedad_2'],
                'Estado': estado,
                'Porcentaje': round(porcentaje, 1),
                'Detalle_Diario': detalle_str,
                'Incumplimiento': tiene_incumplimiento,
                'Incumplimiento_Detalle': ' | '.join(incumplimiento_diario) if incumplimiento_diario else '✅ Todo ok',
                'Dias_Incumplidos': len(dias_incumplidos),
                'Plataformas_Activas': sum([1 for x in [data['Camp Legal'], data['Smokeball'], data['Toggl']] if x > 0])
            })
        
        self.df_analisis = pd.DataFrame(datos)
        
        if len(self.df_analisis) > 0:
            self.df_analisis = self.df_analisis.sort_values(['Incumplimiento', 'Porcentaje'], ascending=[False, True])
        
        print(f"\n✅ Usuarios en reporte: {len(self.df_analisis)}")
        
        incumplidores = self.df_analisis[self.df_analisis['Incumplimiento'] == True]
        if len(incumplidores) > 0:
            print(f"⚠️ Usuarios con incumplimiento: {len(incumplidores)}")
            for _, row in incumplidores.iterrows():
                print(f"   • {row['Usuario']} ({row['Compañia']}): {row['Incumplimiento_Detalle']}")
        
        return True
    
    def calcular_estado_rango(self, total, camp, sb, tg, porcentaje):
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
                'horas_tg': 0,
                'incumplidores': 0
            }
        
        df = self.df_analisis
        return {
            'total_usuarios': len(df),
            'total_horas': df['Total_Horas'].sum(),
            'promedio': df['Total_Horas'].mean(),
            'con_permiso': len(df[df['Permiso'] != 'Sin permiso']),
            'horas_camp': df['Camp Legal'].sum(),
            'horas_sb': df['Smokeball'].sum(),
            'horas_tg': df['Toggl'].sum(),
            'incumplidores': len(df[df['Incumplimiento'] == True])
        }

# ============================================================
# INTERFAZ STREAMLIT
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
    .main-header h1 { margin: 0; font-size: 2rem; }
    .main-header p { margin: 0.3rem 0 0 0; opacity: 0.85; }
    .card-metric {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        border-top: 4px solid #2c5f8a;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .card-metric .value { font-size: 28px; font-weight: 800; color: #1a3a5c; }
    .card-metric .label { font-size: 11px; text-transform: uppercase; color: #7a8a9e; }
    .card-metric.danger { border-top-color: #e74c3c; }
    .card-metric.danger .value { color: #e74c3c; }
    .card-metric.success { border-top-color: #27ae60; }
    .card-metric.success .value { color: #27ae60; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📊 Reporte de Tiempos</h1>
    <p>Análisis de tiempos por plataforma · Control de calidad · Festivos</p>
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
    
    archivo_powerbi = st.file_uploader(
        "📊 Power BI resources.xlsx (Names)",
        type=['xlsx'],
        key="powerbi"
    )
    
    st.caption("📌 Un solo archivo con dos hojas:\n- Novedades: Permisos (rango de fechas)\n- Novedades 2: Festivos (días específicos)")
    archivo_novedades_max = st.file_uploader(
        "📋 Template_Novedades_RRHH_MAX 1 1 2.xlsx",
        type=['xlsx'],
        key="novedades_max"
    )
    
    archivo_novedades_clg = st.file_uploader(
        "📋 Template_Novedades_RRHH_CLG - last.xlsx",
        type=['xlsx'],
        key="novedades_clg"
    )
    
    st.divider()
    st.subheader("📌 Plataformas (opcionales)")
    
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
        archivo_novedades_max is not None,
        archivo_novedades_clg is not None,
        archivo_camp is not None,
        archivo_sb is not None,
        archivo_tg is not None
    ])
    
    st.info(f"📁 Archivos cargados: {archivos_cargados}/6")
    
    procesar = st.button(
        "🚀 Generar Reporte",
        type="primary",
        use_container_width=True,
        disabled=archivos_cargados < 2
    )

if not procesar:
    st.info("👈 Sube los archivos requeridos (Power BI y Novedades MAX) y presiona 'Generar Reporte'")
    st.markdown("""
    **✅ Obligatorios:**
    - Power BI resources.xlsx
    - Template_Novedades_RRHH_MAX (2 hojas)
    
    **📌 Opcionales:**
    - Template_Novedades_RRHH_CLG
    - Camp Legal, Smokeball, Toggl
    """)
    st.stop()

# ============================================================
# PROCESAR
# ============================================================

with st.spinner("🔄 Procesando datos... Por favor espera"):
    try:
        sistema = ReporteTiemposSystem(fecha_inicio, fecha_fin)
        
        # Cargar archivos
        if archivo_powerbi is not None:
            sistema.cargar_archivo(archivo_powerbi, 'powerbi')
            st.success("✅ Power BI cargado")
        
        if archivo_novedades_max is not None:
            sistema.cargar_archivo(archivo_novedades_max, 'novedades_max', sheet_name='Novedades')
            sistema.cargar_archivo(archivo_novedades_max, 'novedades_max_2', sheet_name='Novedades 2')
            st.success("✅ Novedades MAX cargado (2 hojas)")
        
        if archivo_novedades_clg is not None:
            sistema.cargar_archivo(archivo_novedades_clg, 'novedades_clg')
            st.success("✅ Novedades CLG cargado")
        
        if archivo_camp is not None:
            sistema.cargar_archivo(archivo_camp, 'camp_legal', sheet_name='Time entries')
            st.success("✅ Camp Legal cargado (Time entries)")
        
        if archivo_sb is not None:
            sistema.cargar_archivo(archivo_sb, 'smokeball', sheet_name='Entries')
            st.success("✅ Smokeball cargado (Entries)")
        
        if archivo_tg is not None:
            sistema.cargar_archivo(archivo_tg, 'toggl', sheet_name='DataBaseToggl')
            st.success("✅ Toggl cargado (DataBaseToggl)")
        
        # Construir mapa de nombres (con inclusión automática de usuarios de Toggl)
        if not sistema.construir_mapa_nombres():
            st.error("❌ Error al construir mapa de nombres. Verifica Power BI.")
            st.stop()
        st.success("✅ Mapa de nombres construido")
        
        # Mostrar usuarios agregados manualmente
        if sistema.usuarios_manuales:
            st.info(f"📋 Usuarios agregados manualmente desde Toggl: {len(sistema.usuarios_manuales)}")
            with st.expander("📋 Usuarios agregados manualmente desde Toggl"):
                for usuario in sorted(sistema.usuarios_manuales):
                    st.write(f"• {usuario}")
        
        # Procesar novedades
        sistema.procesar_novedades()
        st.success(f"✅ Novedades procesadas: {len(sistema.usuarios_novedades_2)} usuarios deben trabajar en festivos")
        
        # Consolidar
        if not sistema.consolidar_todas_plataformas():
            st.error("❌ Error al consolidar datos")
            st.stop()
        st.success("✅ Datos consolidados")
        
        # Resultados
        df_resultados = sistema.obtener_resultados()
        estadisticas = sistema.get_estadisticas()
        
        if df_resultados is None or df_resultados.empty:
            st.warning("⚠️ No se encontraron resultados")
            st.stop()
        
        # ============================================================
        # DIAGNÓSTICO DE USUARIOS ESPECÍFICOS
        # ============================================================
        
        with st.expander("🔍 Diagnóstico de usuarios específicos", expanded=True):
            
            usuarios_a_verificar = [
                'Gustavo Meneses', 'Amber Spelman', 'Angelly Castañeda', 
                'Cristhian Medina', 'Daniela Squires', 'Delmin Salazar',
                'Diana Carolina Moise', 'Diego Londoño', 'Gabriel Taborda',
                'Joseph Hamilton', 'Juan David Quintero', 'Laura Valentina Bonilla',
                'Luisa Segura', 'Maria Camila Acosta', 'Maria Lucia Mage',
                'Marion Garcia', 'Marisol Tinajero', 'Monica Lopez Villamizar',
                'Perlita Avila Casulla', 'Victor Zarate'
            ]
            
            st.write("### 🔍 Verificando usuarios específicos")
            
            for usuario in usuarios_a_verificar:
                st.write(f"---")
                st.write(f"**👤 {usuario}**")
                
                # Buscar en el reporte final
                if df_resultados is not None and len(df_resultados) > 0:
                    usuario_resultado = df_resultados[df_resultados['Usuario'].str.contains(usuario.split()[0], case=False, na=False)]
                    if len(usuario_resultado) > 0:
                        st.success(f"✅ Encontrado en el reporte final")
                        st.dataframe(usuario_resultado[['Usuario', 'Total_Horas', 'Camp Legal', 'Smokeball', 'Toggl', 'Estado']])
                    else:
                        st.warning(f"❌ NO encontrado en el reporte final")
                        
                        # Buscar en Toggl
                        if sistema.df_toggl is not None and 'Member' in sistema.df_toggl.columns:
                            tg_found = sistema.df_toggl[sistema.df_toggl['Member'].str.contains(usuario.split()[0], case=False, na=False)]
                            if len(tg_found) > 0:
                                # Convertir horas
                                tg_found['Horas_Num'] = tg_found['Dur'].apply(convertir_hora_tiempo)
                                horas = tg_found['Horas_Num'].sum()
                                st.write(f"   Toggl: ✅ {horas:.2f} horas")
                                
                                # Mostrar fechas
                                if 'Date1' in tg_found.columns:
                                    tg_found['Date_Conv'] = tg_found['Date1'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                                    fechas = tg_found['Date_Conv'].dropna().unique()
                                    st.write(f"   Fechas en Toggl: {sorted(fechas)[:5]}")
                            else:
                                st.write(f"   Toggl: ❌ No encontrado")
                        
                        # Buscar en Camp Legal
                        if sistema.df_camp is not None and 'Staff Name' in sistema.df_camp.columns:
                            camp_found = sistema.df_camp[sistema.df_camp['Staff Name'].str.contains(usuario.split()[0], case=False, na=False)]
                            if len(camp_found) > 0:
                                camp_found['Horas_Num'] = camp_found['Hours Spent'].apply(convertir_hora_tiempo)
                                horas = camp_found['Horas_Num'].sum()
                                st.write(f"   Camp Legal: ✅ {horas:.2f} horas")
                            else:
                                st.write(f"   Camp Legal: ❌ No encontrado")
                        
                        # Buscar en Smokeball
                        if sistema.df_smokeball is not None and 'Name' in sistema.df_smokeball.columns:
                            sb_found = sistema.df_smokeball[sistema.df_smokeball['Name'].str.contains(usuario.split()[0], case=False, na=False)]
                            if len(sb_found) > 0:
                                sb_found['Horas_Num'] = sb_found['Hours'].apply(convertir_hora_decimal)
                                horas = sb_found['Horas_Num'].sum()
                                st.write(f"   Smokeball: ✅ {horas:.2f} horas")
                            else:
                                st.write(f"   Smokeball: ❌ No encontrado")
        
        # ============================================================
        # DIAGNÓSTICO ESPECÍFICO DE TOGGL
        # ============================================================
        
        with st.expander("🔍 DIAGNÓSTICO ESPECÍFICO - TOGGL", expanded=True):
            
            st.write("### ⏱️ ANÁLISIS COMPLETO DE TOGGL")
            
            if sistema.df_toggl is not None:
                st.write(f"**Total registros en Toggl:** {len(sistema.df_toggl)}")
                
                col_nombre = 'Member'
                col_horas = 'Dur'
                col_fecha = 'Date1'
                
                if col_nombre in sistema.df_toggl.columns:
                    # Convertir horas y fechas
                    sistema.df_toggl['Horas_Num'] = sistema.df_toggl[col_horas].apply(convertir_hora_tiempo)
                    sistema.df_toggl['Date_Conv'] = sistema.df_toggl[col_fecha].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                    
                    # Filtrar por rango
                    df_rango = sistema.df_toggl[
                        (sistema.df_toggl['Date_Conv'] >= fecha_inicio) & 
                        (sistema.df_toggl['Date_Conv'] <= fecha_fin)
                    ]
                    
                    st.write(f"**Registros en el rango del reporte:** {len(df_rango)}")
                    
                    if len(df_rango) > 0:
                        # Agrupar por usuario
                        df_agg = df_rango.groupby(col_nombre).agg({
                            'Horas_Num': 'sum'
                        }).reset_index().sort_values('Horas_Num', ascending=False)
                        
                        st.write("**Top 20 usuarios por horas en Toggl:**")
                        st.dataframe(df_agg.head(20))
                        
                        # Verificar qué usuarios de Toggl NO están en el reporte
                        usuarios_tg = set(df_agg[col_nombre].unique())
                        usuarios_final = set(df_resultados['Usuario'].tolist()) if df_resultados is not None else set()
                        
                        # Normalizar para comparación
                        usuarios_tg_norm = set()
                        for u in usuarios_tg:
                            if isinstance(u, str):
                                u_norm = normalizar_nombre_flexible(u)
                                usuarios_tg_norm.add(u_norm)
                        
                        usuarios_final_norm = set()
                        for u in usuarios_final:
                            if isinstance(u, str):
                                u_norm = normalizar_nombre_flexible(u)
                                usuarios_final_norm.add(u_norm)
                        
                        usuarios_faltantes = usuarios_tg_norm - usuarios_final_norm
                        
                        if usuarios_faltantes:
                            st.warning(f"⚠️ {len(usuarios_faltantes)} usuarios de Toggl NO están en el reporte final:")
                            
                            for usuario_norm in sorted(usuarios_faltantes)[:20]:
                                # Encontrar nombre original
                                nombre_orig = None
                                for u in usuarios_tg:
                                    if isinstance(u, str) and normalizar_nombre_flexible(u) == usuario_norm:
                                        nombre_orig = u
                                        break
                                
                                if nombre_orig:
                                    registros = df_rango[df_rango[col_nombre] == nombre_orig]
                                    horas = registros['Horas_Num'].sum()
                                    dias = registros['Date_Conv'].nunique()
                                    st.write(f"• {nombre_orig} → {usuario_norm}: {horas:.2f}h ({dias} días)")
                        else:
                            st.success("✅ Todos los usuarios de Toggl están en el reporte final")
        # ============================================================
        # MOSTRAR RESULTADOS
        # ============================================================
        
        st.markdown("---")
        st.markdown("### 📊 Resultados del Reporte")
        
        incumplidores = df_resultados[df_resultados['Incumplimiento'] == True]
        if len(incumplidores) > 0:
            st.error(f"🚨 {len(incumplidores)} usuarios NO cumplieron con la jornada mínima")
        else:
            st.success("✅ Todos los usuarios cumplieron con la jornada mínima")
        
        usuarios_festivos = df_resultados[df_resultados['Novedad_2'] == 'Sí']
        if len(usuarios_festivos) > 0:
            st.info(f"📋 {len(usuarios_festivos)} usuarios deben trabajar en festivos (Novedades 2)")
        
        st.info(f"👥 Total de usuarios: {len(df_resultados)}")
        
        # KPIs
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
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
        
        with col5:
            color = "#e74c3c" if estadisticas['incumplidores'] > 0 else "#27ae60"
            st.markdown(f"""
            <div class="card-metric danger" style="border-top-color: {color};">
                <div class="value" style="color: {color};">{estadisticas['incumplidores']}</div>
                <div class="label">🚨 Incumplen</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col6:
            companias = df_resultados['Compañia'].unique()
            st.markdown(f"""
            <div class="card-metric" style="border-top-color: #9b59b6;">
                <div class="value" style="font-size: 16px;">{', '.join(companias)}</div>
                <div class="label">🏢 Compañías</div>
            </div>
            """, unsafe_allow_html=True)
        
        # ============================================================
        # TABLA
        # ============================================================
        
        st.markdown("### 👥 Detalle por Usuario")
        
        df_mostrar = df_resultados[[
            'Usuario', 'Compañia', 'Camp Legal', 'Smokeball', 'Toggl', 
            'Total_Horas', 'Dias_Activos', 'Permiso', 'Novedad_2', 'Estado', 'Incumplimiento', 'Incumplimiento_Detalle'
        ]].copy()
        
        df_mostrar.columns = [
            'Usuario', 'Compañia', 'Camp Legal', 'Smokeball', 'Toggl',
            'Total Horas', 'Días Activos', 'Permiso', '📋 Festivo', 'Estado', '⚠️ Incumple', 'Detalle Incumplimiento'
        ]
        
        df_mostrar['⚠️ Incumple'] = df_mostrar['⚠️ Incumple'].apply(lambda x: '🚨 SÍ' if x else '✅ NO')
        
        st.dataframe(
            df_mostrar,
            use_container_width=True,
            height=400,
            hide_index=True
        )
        
        # ============================================================
        # FILTROS AVANZADOS
        # ============================================================
        
        with st.expander("🔍 Filtros avanzados"):
            col_filtro1, col_filtro2 = st.columns(2)
            
            with col_filtro1:
                compania_seleccionada = st.selectbox(
                    "Filtrar por Compañía",
                    ['Todas'] + sorted(df_resultados['Compañia'].unique().tolist())
                )
            
            with col_filtro2:
                mostrar_solo_incumplidores = st.checkbox("Mostrar solo usuarios que incumplen")
            
            df_filtrado = df_mostrar.copy()
            
            if compania_seleccionada != 'Todas':
                df_filtrado = df_filtrado[df_filtrado['Compañia'] == compania_seleccionada]
            
            if mostrar_solo_incumplidores:
                df_filtrado = df_filtrado[df_filtrado['⚠️ Incumple'] == '🚨 SÍ']
            
            if len(df_filtrado) > 0:
                st.dataframe(
                    df_filtrado,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No hay usuarios que coincidan con los filtros seleccionados")
        
        # ============================================================
        # RESUMEN POR COMPAÑÍA
        # ============================================================
        
        st.markdown("### 📊 Resumen por Compañía")
        
        resumen_compania = df_resultados.groupby('Compañia').agg({
            'Usuario': 'count',
            'Total_Horas': 'sum',
            'Incumplimiento': lambda x: sum(x)
        }).reset_index()
        
        resumen_compania.columns = ['Compañía', 'Usuarios', 'Total Horas', 'Incumplidores']
        resumen_compania['Promedio'] = (resumen_compania['Total Horas'] / resumen_compania['Usuarios']).round(1)
        
        st.dataframe(
            resumen_compania,
            use_container_width=True,
            hide_index=True
        )
        
        # ============================================================
        # TARJETAS POR USUARIO
        # ============================================================
        
        st.markdown("### 📋 Detalle por Usuario (Tarjetas)")
        
        cols = st.columns(3)
        
        for idx, (_, row) in enumerate(df_resultados.iterrows()):
            col = cols[idx % 3]
            
            if row['Incumplimiento']:
                bg_color = '#fdedec'
                border_color = '#e74c3c'
                estado_emoji = '🚨'
                estado_text = 'Incumple'
            elif row['Total_Horas'] == 0:
                bg_color = '#f4f6f7'
                border_color = '#95a5a6'
                estado_emoji = '⛔'
                estado_text = 'Sin registro'
            else:
                bg_color = '#eafaf1'
                border_color = '#27ae60'
                estado_emoji = '✅'
                estado_text = 'Cumple'
            
            festivo_tag = ''
            if row['Novedad_2'] == 'Sí':
                festivo_tag = '<span style="background: #f39c12; color: white; padding: 1px 8px; border-radius: 10px; font-size: 9px; font-weight: 600; margin-left: 5px;">📋 Festivo</span>'
            
            with col:
                st.markdown(f"""
                <div style="
                    background-color: {bg_color};
                    border: 2px solid {border_color};
                    border-radius: 12px;
                    padding: 14px 16px;
                    margin-bottom: 12px;
                ">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <div>
                            <span style="font-size: 15px; font-weight: 700;">{estado_emoji} {row['Usuario']}</span>
                            <br>
                            <span style="font-size: 12px; color: #666;">🏢 {row['Compañia']}</span>
                            {festivo_tag}
                        </div>
                        <div style="font-size: 22px; font-weight: 800; color: {border_color};">
                            {row['Total_Horas']:.1f}h
                        </div>
                    </div>
                    <div style="font-size: 13px; color: #555; margin-top: 6px;">
                        CL: {row['Camp Legal']:.1f}h · SB: {row['Smokeball']:.1f}h · TG: {row['Toggl']:.1f}h
                    </div>
                    <div style="font-size: 12px; color: #888; margin-top: 4px; display: flex; justify-content: space-between;">
                        <span>📅 {row['Dias_Activos']} días activos</span>
                        <span style="font-weight: 600; color: {border_color};">{estado_text}</span>
                    </div>
                    {f'<div style="font-size: 11px; color: #c0392b; margin-top: 6px; border-top: 1px solid #f5c6cb; padding-top: 6px;">⚠️ {row["Incumplimiento_Detalle"]}</div>' if row['Incumplimiento'] else ''}
                </div>
                """, unsafe_allow_html=True)
        
        # ============================================================
        # DESCARGAR RESULTADOS
        # ============================================================
        
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

st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #95a5a6; font-size: 12px;">
    Reporte generado automáticamente · Datos de Camp Legal, Smokeball y Toggl
</div>
""", unsafe_allow_html=True)
