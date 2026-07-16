# sistema_reporte_tiempos_v33.py
# VERSIÓN CON FILTRO CORRECTO DE NOVEDADES 2

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
import re
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# CONFIGURACIÓN - RUTA UNIFICADA
# ============================================================

RUTA_BASE = r"C:\Users\MSI\OneDrive - Community Law Group\Automation\Data Automation Platforms"

# ============================================================
# MAPEO DE COLUMNAS POR ARCHIVO
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
    """Determina si una fecha es festivo"""
    # Festivos de Colombia (puedes modificarlos según tu país)
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
        self.ruta_base = RUTA_BASE
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
        self.usuarios_con_plataforma = []
        
        # Lista de usuarios que aparecen en Novedades 2
        self.usuarios_novedades_2 = set()
        
        # Calcular días en el rango
        self.dias_totales = (self.fecha_fin - self.fecha_inicio).days + 1
        self.dias_habiles = self._calcular_dias_habiles()
        
        # Calcular jornada total esperada considerando sábados y festivos
        self.jornada_total_esperada = self._calcular_jornada_total()
        
        print(f"\n📅 Rango: {self.fecha_inicio.strftime('%d/%m/%Y')} - {self.fecha_fin.strftime('%d/%m/%Y')}")
        print(f"📊 Días totales: {self.dias_totales} | Días hábiles: {self.dias_habiles}")
        print(f"📊 Jornada total esperada: {self.jornada_total_esperada:.1f}h")
    
    def _calcular_dias_habiles(self):
        """Calcula los días hábiles (lunes a viernes) en el rango"""
        dias = 0
        fecha_actual = self.fecha_inicio
        while fecha_actual <= self.fecha_fin:
            if fecha_actual.weekday() < 5:
                dias += 1
            fecha_actual += timedelta(days=1)
        return max(dias, 1)
    
    def _calcular_jornada_total(self):
        """Calcula la jornada total esperada considerando sábados, domingos y festivos"""
        total = 0
        fecha_actual = self.fecha_inicio
        while fecha_actual <= self.fecha_fin:
            total += get_jornada_esperada_por_dia(fecha_actual)
            fecha_actual += timedelta(days=1)
        return max(total, 1)
    
    def _obtener_dias_con_jornada(self):
        """Obtiene la lista de días en el rango con su jornada esperada"""
        dias = []
        fecha_actual = self.fecha_inicio
        while fecha_actual <= self.fecha_fin:
            jornada = get_jornada_esperada_por_dia(fecha_actual)
            if jornada > 0:
                dias.append((fecha_actual, jornada))
            fecha_actual += timedelta(days=1)
        return dias
    
    # ============================================================
    # LECTURA DE ARCHIVOS
    # ============================================================
    
    def leer_archivos(self):
        print("\n" + "="*70)
        print("📂 LEYENDO ARCHIVOS")
        print("="*70)
        
        for key, config in COLUMNAS_MAPEO.items():
            nombre = config['archivo']
            ruta = os.path.join(self.ruta_base, nombre)
            print(f"\n📄 Buscando: {nombre}")
            
            if not os.path.exists(ruta):
                print(f"   ❌ Archivo NO ENCONTRADO")
                if key == 'camp_legal':
                    self.df_camp = None
                elif key == 'smokeball':
                    self.df_smokeball = None
                elif key == 'toggl':
                    self.df_toggl = None
                elif key == 'powerbi':
                    self.df_powerbi = None
                elif key == 'novedades_max':
                    self.df_novedades_max = None
                elif key == 'novedades_max_2':
                    self.df_novedades_max_2 = None
                elif key == 'novedades_clg':
                    self.df_novedades_clg = None
                continue
            
            try:
                if key == 'powerbi':
                    hoja_datos = config.get('hoja_datos', 'Names')
                    xl = pd.ExcelFile(ruta)
                    hojas_disponibles = xl.sheet_names
                    if hoja_datos in hojas_disponibles:
                        df = pd.read_excel(ruta, sheet_name=hoja_datos)
                        print(f"   ✅ Hoja '{hoja_datos}' cargada: {len(df)} registros")
                        self.df_powerbi = df
                    else:
                        print(f"   ❌ Hoja '{hoja_datos}' no encontrada")
                        self.df_powerbi = None
                    continue
                
                hoja_datos = config.get('hoja_datos', 'Hoja1')
                try:
                    xl = pd.ExcelFile(ruta)
                except PermissionError:
                    print(f"   ⚠️ Archivo bloqueado - cierra el archivo y vuelve a ejecutar")
                    continue
                
                hojas_disponibles = xl.sheet_names
                if hoja_datos in hojas_disponibles:
                    df = pd.read_excel(ruta, sheet_name=hoja_datos)
                    print(f"   ✅ Hoja '{hoja_datos}' cargada: {len(df)} registros")
                    
                    if key == 'camp_legal':
                        self.df_camp = df
                    elif key == 'smokeball':
                        self.df_smokeball = df
                    elif key == 'toggl':
                        self.df_toggl = df
                    elif key == 'novedades_max':
                        self.df_novedades_max = df
                    elif key == 'novedades_max_2':
                        self.df_novedades_max_2 = df
                        print(f"\n   📊 NOVEDADES 2 - Muestra:")
                        if 'Persona' in df.columns and 'Fecha' in df.columns:
                            print(df[['Persona', 'Fecha', 'Tipo de Novedad']].head(5))
                    elif key == 'novedades_clg':
                        self.df_novedades_clg = df
                else:
                    print(f"   ❌ Hoja '{hoja_datos}' no encontrada")
                    if key == 'camp_legal':
                        self.df_camp = None
                    elif key == 'smokeball':
                        self.df_smokeball = None
                    elif key == 'toggl':
                        self.df_toggl = None
                    elif key == 'novedades_max':
                        self.df_novedades_max = None
                    elif key == 'novedades_max_2':
                        self.df_novedades_max_2 = None
                    elif key == 'novedades_clg':
                        self.df_novedades_clg = None
                
            except Exception as e:
                print(f"   ❌ Error: {e}")
                if key == 'camp_legal':
                    self.df_camp = None
                elif key == 'smokeball':
                    self.df_smokeball = None
                elif key == 'toggl':
                    self.df_toggl = None
                elif key == 'powerbi':
                    self.df_powerbi = None
                elif key == 'novedades_max':
                    self.df_novedades_max = None
                elif key == 'novedades_max_2':
                    self.df_novedades_max_2 = None
                elif key == 'novedades_clg':
                    self.df_novedades_clg = None
        
        cargados = sum([
            self.df_camp is not None,
            self.df_smokeball is not None,
            self.df_toggl is not None,
            self.df_powerbi is not None,
            self.df_novedades_max_2 is not None
        ])
        
        print(f"\n📊 Archivos cargados: {cargados}/7")
        return cargados >= 4
    
    # ============================================================
    # CONSTRUIR MAPA DE NOMBRES
    # ============================================================
    
    def construir_mapa_nombres(self):
        print("\n" + "="*70)
        print("🔄 CONSTRUYENDO MAPA DE NOMBRES")
        print("="*70)
        
        if self.df_powerbi is None:
            print("⚠️ No hay datos de Power BI")
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
        
        print(f"   ✅ Usuarios a incluir: {len(self.usuarios_con_plataforma)}")
        return True
    
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
    
    # ============================================================
    # PROCESAR NOVEDADES
    # ============================================================
    
    def procesar_novedades(self):
        print("\n" + "="*70)
        print("📋 PROCESANDO NOVEDADES")
        print("="*70)
        
        novedades_list = []
        
        # Procesar Novedades MAX (rango de fechas)
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
                print(f"   ✅ MAX: {len(df_max_validos)} registros")
        
        # Procesar Novedades MAX 2 (fecha específica) - IMPORTANTE: extraer usuarios
        if self.df_novedades_max_2 is not None:
            df_max2 = self.df_novedades_max_2.copy()
            if 'Persona' in df_max2.columns and 'Fecha' in df_max2.columns:
                # Filtrar por el rango de fechas del reporte
                df_max2['Fecha_Conv'] = df_max2['Fecha'].apply(lambda x: convertir_fecha(x, '%m/%d/%Y'))
                df_max2_filtrado = df_max2[
                    (df_max2['Fecha_Conv'] >= self.fecha_inicio) & 
                    (df_max2['Fecha_Conv'] <= self.fecha_fin)
                ]
                
                # Extraer usuarios de Novedades 2
                for _, row in df_max2_filtrado.iterrows():
                    nombre = row['Persona']
                    nombre_normalizado = self.normalizar_nombre(nombre)
                    if nombre_normalizado:
                        self.usuarios_novedades_2.add(nombre_normalizado)
                
                print(f"   ✅ Usuarios en Novedades 2 en el rango: {len(self.usuarios_novedades_2)}")
                
                # Procesar como novedad regular también
                df_max2_validos = df_max2_filtrado[df_max2_filtrado['Fecha_Conv'].notna()]
                if 'Tipo de Novedad' in df_max2.columns:
                    df_max2_validos['Tipo'] = df_max2_validos['Tipo de Novedad']
                else:
                    df_max2_validos['Tipo'] = 'Novedad 2'
                df_max2_validos['Fecha_Inicio'] = df_max2_validos['Fecha_Conv']
                df_max2_validos['Fecha_Fin'] = df_max2_validos['Fecha_Conv']
                df_max2_validos['Usuario_Normalizado'] = df_max2_validos['Persona'].apply(self.normalizar_nombre)
                novedades_list.append(df_max2_validos[['Usuario_Normalizado', 'Fecha_Inicio', 'Fecha_Fin', 'Tipo']])
                print(f"   ✅ MAX 2: {len(df_max2_validos)} registros en el rango")
            else:
                print(f"   ⚠️ Columnas 'Persona' o 'Fecha' no encontradas en Novedades 2")
        
        # Procesar Novedades CLG
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
                print(f"   ✅ CLG: {len(df_clg_validos)} registros")
        
        if novedades_list:
            self.df_novedades_combinadas = pd.concat(novedades_list, ignore_index=True)
            print(f"   ✅ Total novedades combinadas: {len(self.df_novedades_combinadas)}")
        else:
            self.df_novedades_combinadas = None
        
        # Mostrar usuarios de Novedades 2
        if self.usuarios_novedades_2:
            print(f"\n   📋 Usuarios en Novedades 2:")
            for usuario in sorted(self.usuarios_novedades_2):
                print(f"      • {usuario}")
        
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
    
    def verificar_novedad_2(self, usuario):
        """Verifica si el usuario aparece en Novedades 2 en el rango de fechas"""
        return usuario in self.usuarios_novedades_2
    
    # ============================================================
    # PROCESAR PLATAFORMA CON RANGO DE FECHAS Y DETALLE DIARIO
    # ============================================================
    
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
        
        # Filtrar por RANGO de fechas
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
        
        # Agrupar por usuario (total del rango)
        df_agrupado = df_proc.groupby('Usuario_Normalizado').agg({
            'Horas': 'sum',
            'Actividad': lambda x: ', '.join(x.unique()[:3])
        }).reset_index()
        
        # Agregar conteo de días con actividad
        df_dias = df_proc.groupby('Usuario_Normalizado')['Date'].nunique().reset_index()
        df_dias.columns = ['Usuario_Normalizado', 'Dias_Activos']
        df_agrupado = df_agrupado.merge(df_dias, on='Usuario_Normalizado', how='left')
        df_agrupado['Dias_Activos'] = df_agrupado['Dias_Activos'].fillna(0).astype(int)
        
        # Detalle diario
        df_diario = df_proc.groupby(['Usuario_Normalizado', 'Date']).agg({
            'Horas': 'sum'
        }).reset_index()
        
        df_diario_pivot = df_diario.pivot(index='Usuario_Normalizado', columns='Date', values='Horas').fillna(0)
        df_diario_pivot.columns = [f'Dia_{c.strftime("%d/%m")}' for c in df_diario_pivot.columns]
        df_diario_pivot = df_diario_pivot.reset_index()
        
        df_agrupado = df_agrupado.merge(df_diario_pivot, on='Usuario_Normalizado', how='left')
        
        return df_agrupado
    
    # ============================================================
    # CONSOLIDAR
    # ============================================================
    
    def consolidar_todas_plataformas(self):
        print("\n" + "="*70)
        print("🔄 CONSOLIDANDO DATOS")
        print("="*70)
        print(f"📅 Rango: {self.fecha_inicio.strftime('%d/%m/%Y')} - {self.fecha_fin.strftime('%d/%m/%Y')}")
        print(f"📊 Jornada total esperada: {self.jornada_total_esperada:.1f}h")
        print(f"📋 Usuarios en Novedades 2: {len(self.usuarios_novedades_2)}")
        print("-"*70)
        
        df_camp = self.procesar_plataforma(self.df_camp, 'camp_legal')
        df_sb = self.procesar_plataforma(self.df_smokeball, 'smokeball')
        df_tg = self.procesar_plataforma(self.df_toggl, 'toggl')
        
        # Obtener días con jornada (solo días con jornada > 0)
        dias_con_jornada = self._obtener_dias_con_jornada()
        dias_columnas = [f'Dia_{fecha.strftime("%d/%m")}' for fecha, _ in dias_con_jornada]
        
        # Crear diccionario con todos los usuarios de Novedades 2
        usuarios_dict = {}
        
        # Solo incluir usuarios que están en Novedades 2
        if not self.usuarios_novedades_2:
            print("⚠️ No hay usuarios en Novedades 2. El reporte estará vacío.")
            self.df_analisis = pd.DataFrame()
            return True
        
        for usuario in self.usuarios_novedades_2:
            usuarios_dict[usuario] = {
                'Camp Legal': 0.0,
                'Smokeball': 0.0,
                'Toggl': 0.0,
                'Dias_Activos': 0,
                'Actividades': [],
                'Permiso': None,
                'Novedad_2': 'Sí',
                'Detalle_Diario': {dia: 0.0 for dia in dias_columnas},
                'Jornada_Diaria': {dia: jornada for dia, jornada in [(f'Dia_{fecha.strftime("%d/%m")}', jornada) for fecha, jornada in dias_con_jornada]}
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
                    
                    for col in dias_columnas:
                        if col in row and row[col] > 0:
                            usuarios_dict[usuario]['Detalle_Diario'][col] += row[col]
        
        procesar_plataforma_detalle(df_camp, 'Camp Legal')
        procesar_plataforma_detalle(df_sb, 'Smokeball')
        procesar_plataforma_detalle(df_tg, 'Toggl')
        
        print("   🔍 Verificando permisos...")
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
            
            # Calcular porcentaje sobre jornada total esperada
            porcentaje = (total / self.jornada_total_esperada * 100) if self.jornada_total_esperada > 0 else 0
            
            # Determinar estado
            if data['Permiso']:
                estado = f"📋 {data['Permiso']}"
            else:
                estado = self.calcular_estado_rango(total, data['Camp Legal'], data['Smokeball'], data['Toggl'], porcentaje)
            
            # Construir detalle diario
            detalle_items = []
            for dia, horas in data['Detalle_Diario'].items():
                if horas > 0:
                    jornada_esperada = data['Jornada_Diaria'].get(dia, 8.0)
                    detalle_items.append(f"{dia}: {horas:.1f}h (esperado {jornada_esperada:.0f}h)")
            detalle_str = ' | '.join(detalle_items)
            
            datos.append({
                'Usuario': usuario,
                'Camp Legal': round(data['Camp Legal'], 2),
                'Smokeball': round(data['Smokeball'], 2),
                'Toggl': round(data['Toggl'], 2),
                'Total_Horas': round(total, 2),
                'Dias_Activos': data['Dias_Activos'],
                'Actividades': ' | '.join(data['Actividades']) if data['Actividades'] else 'Sin registro',
                'Permiso': data['Permiso'] if data['Permiso'] else 'Sin permiso',
                'Novedad_2': 'Sí',
                'Estado': estado,
                'Porcentaje': round(porcentaje, 1),
                'Detalle_Diario': detalle_str,
                'Plataformas_Activas': sum([1 for x in [data['Camp Legal'], data['Smokeball'], data['Toggl']] if x > 0])
            })
        
        self.df_analisis = pd.DataFrame(datos)
        
        print(f"\n✅ Usuarios en reporte: {len(self.df_analisis)}")
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
    
    # ============================================================
    # GENERAR HTML CON DESGLOSE DIARIO
    # ============================================================
    
    def generar_html_reporte(self):
        if self.df_analisis is None or self.df_analisis.empty:
            return self._generar_html_vacio()
        
        df = self.df_analisis
        fecha_inicio_str = self.fecha_inicio.strftime('%d/%m/%Y')
        fecha_fin_str = self.fecha_fin.strftime('%d/%m/%Y')
        
        total_usuarios = len(df)
        total_horas = df['Total_Horas'].sum()
        promedio = df['Total_Horas'].mean()
        con_permiso = len(df[df['Permiso'] != 'Sin permiso'])
        
        horas_camp = df['Camp Legal'].sum()
        horas_sb = df['Smokeball'].sum()
        horas_tg = df['Toggl'].sum()
        
        # Generar tarjetas con detalle diario
        tarjetas_html = ""
        df_ordenado = df.sort_values('Porcentaje', ascending=True)
        
        for _, row in df_ordenado.iterrows():
            if row['Estado'].startswith('✅'):
                bg_color = '#eafaf1'
                icono = '✅'
            elif row['Estado'].startswith('⚠️'):
                bg_color = '#fef9e7'
                icono = '⚠️'
            elif row['Estado'].startswith('❌'):
                bg_color = '#fdedec'
                icono = '❌'
            elif row['Estado'].startswith('⛔'):
                bg_color = '#f4f6f7'
                icono = '⛔'
            elif row['Estado'].startswith('📋'):
                bg_color = '#f4ecf7'
                icono = '📋'
            else:
                bg_color = '#f8f9fa'
                icono = '📊'
            
            tags = []
            if row['Camp Legal'] > 0:
                tags.append('<span class="tag tag-cl">CL</span>')
            if row['Smokeball'] > 0:
                tags.append('<span class="tag tag-sb">SB</span>')
            if row['Toggl'] > 0:
                tags.append('<span class="tag tag-tg">TG</span>')
            if row['Permiso'] != 'Sin permiso':
                tags.append('<span class="tag tag-permiso">🔒</span>')
            if row['Novedad_2'] == 'Sí':
                tags.append('<span class="tag tag-novedad">📋</span>')
            if not tags:
                tags.append('<span class="tag tag-sin">—</span>')
            
            horas_detalle = []
            if row['Camp Legal'] > 0:
                horas_detalle.append(f'CL {row["Camp Legal"]:.1f}h')
            if row['Smokeball'] > 0:
                horas_detalle.append(f'SB {row["Smokeball"]:.1f}h')
            if row['Toggl'] > 0:
                horas_detalle.append(f'TG {row["Toggl"]:.1f}h')
            horas_detalle_str = ' · '.join(horas_detalle) if horas_detalle else 'Sin registro'
            
            # Detalle diario
            detalle_diario_html = ""
            if row['Detalle_Diario']:
                items = row['Detalle_Diario'].split(' | ')
                for item in items:
                    # Extraer hora y esperado del formato: "Dia: X.Xh (esperado Yh)"
                    match = re.search(r'(\d{2}/\d{2}): (\d+\.?\d*)h \(esperado (\d+)h\)', item)
                    if match:
                        dia, horas_str, esperado = match.groups()
                        horas_float = float(horas_str)
                        esperado_int = int(esperado)
                        # Color según cumplimiento de la jornada esperada para ese día
                        if horas_float >= esperado_int:
                            color = '#27ae60'
                        elif horas_float >= esperado_int * 0.5:
                            color = '#f39c12'
                        elif horas_float > 0:
                            color = '#e74c3c'
                        else:
                            color = '#bdc3c7'
                        detalle_diario_html += f'<span class="dia-item" style="background:{color};">{dia} {horas_float:.1f}h</span>'
            
            dias_info = f'📅 {row["Dias_Activos"]} días activos'
            jornada_text = f'Jornada: {self.jornada_total_esperada:.0f}h'
            
            tarjetas_html += f"""
                <div class="card" style="background: {bg_color};">
                    <div class="card-header">
                        <div class="card-nombre">
                            <span class="card-icono">{icono}</span>
                            {row['Usuario']}
                        </div>
                        <div class="card-tags">{' '.join(tags)}</div>
                    </div>
                    <div class="card-body">
                        <div class="card-horas">
                            <span class="total">{row['Total_Horas']:.1f}h</span>
                            <span class="detalle">{horas_detalle_str}</span>
                        </div>
                        <div class="card-estado">
                            <span class="estado-text">{row['Estado']}</span>
                        </div>
                    </div>
                    <div class="card-footer">
                        <span class="dias-info">{dias_info}</span>
                        <span class="porcentaje-info">{row['Porcentaje']:.0f}% · {jornada_text}</span>
                    </div>
                    {f'<div class="card-detalle">{detalle_diario_html}</div>' if detalle_diario_html else ''}
                </div>
            """
        
        return f"""
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Reporte Tiempos - {fecha_inicio_str}</title>
            <style>
                * {{ margin: 0; padding: 0; box-sizing: border-box; }}
                body {{
                    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
                    background: #eef2f7;
                    padding: 20px;
                    color: #1a2332;
                }}
                .container {{ max-width: 1400px; margin: 0 auto; }}
                
                .header {{
                    background: linear-gradient(135deg, #1a3a5c 0%, #2c5f8a 100%);
                    color: white;
                    padding: 18px 25px;
                    border-radius: 12px;
                    margin-bottom: 18px;
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    flex-wrap: wrap;
                    gap: 12px;
                }}
                .header h1 {{
                    font-size: 20px;
                    font-weight: 700;
                    display: flex;
                    align-items: center;
                    gap: 10px;
                }}
                .header h1 span {{
                    background: rgba(255,255,255,0.2);
                    padding: 1px 12px;
                    border-radius: 20px;
                    font-size: 13px;
                }}
                .header .fecha {{
                    font-size: 14px;
                    font-weight: 600;
                    opacity: 0.9;
                }}
                .header .rango {{
                    font-size: 13px;
                    opacity: 0.75;
                    margin-top: 2px;
                }}
                .header .info {{
                    font-size: 12px;
                    opacity: 0.8;
                    margin-top: 2px;
                }}
                
                .kpi-row {{
                    display: grid;
                    grid-template-columns: repeat(4, 1fr);
                    gap: 12px;
                    margin-bottom: 18px;
                }}
                .kpi {{
                    background: white;
                    padding: 14px 18px;
                    border-radius: 10px;
                    text-align: center;
                    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
                }}
                .kpi .num {{ font-size: 28px; font-weight: 800; }}
                .kpi .label {{ font-size: 10px; text-transform: uppercase; color: #7a8a9e; font-weight: 600; letter-spacing: 0.3px; }}
                .kpi .sub {{ font-size: 11px; color: #95a5a6; margin-top: 2px; }}
                .kpi.blue .num {{ color: #2c5f8a; }}
                .kpi.purple .num {{ color: #8e44ad; }}
                .kpi.green .num {{ color: #27ae60; }}
                .kpi.orange .num {{ color: #e67e22; }}
                
                .plat-row {{
                    display: grid;
                    grid-template-columns: repeat(3, 1fr);
                    gap: 12px;
                    margin-bottom: 18px;
                }}
                .plat {{
                    background: white;
                    padding: 10px 16px;
                    border-radius: 10px;
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
                }}
                .plat .name {{ font-weight: 600; font-size: 13px; }}
                .plat .horas {{ font-size: 20px; font-weight: 800; }}
                .plat.cl .name {{ color: #2471a3; }}
                .plat.cl .horas {{ color: #2471a3; }}
                .plat.sb .name {{ color: #1e8449; }}
                .plat.sb .horas {{ color: #1e8449; }}
                .plat.tg .name {{ color: #ca6f1e; }}
                .plat.tg .horas {{ color: #ca6f1e; }}
                
                .cards-grid {{
                    display: grid;
                    grid-template-columns: repeat(auto-fill, minmax(360px, 1fr));
                    gap: 12px;
                }}
                
                .card {{
                    border-radius: 12px;
                    padding: 14px 18px;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
                    transition: transform 0.15s, box-shadow 0.15s;
                }}
                .card:hover {{
                    transform: translateY(-2px);
                    box-shadow: 0 4px 16px rgba(0,0,0,0.10);
                }}
                
                .card-header {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 8px;
                    flex-wrap: wrap;
                    gap: 4px;
                }}
                .card-nombre {{
                    font-size: 14px;
                    font-weight: 700;
                    color: #1a2332;
                    display: flex;
                    align-items: center;
                    gap: 6px;
                }}
                .card-icono {{ font-size: 16px; }}
                .card-tags {{ display: flex; gap: 4px; flex-wrap: wrap; }}
                
                .tag {{
                    display: inline-block;
                    padding: 1px 10px;
                    border-radius: 12px;
                    font-size: 10px;
                    font-weight: 700;
                    color: white;
                    letter-spacing: 0.5px;
                }}
                .tag-cl {{ background: #3498db; }}
                .tag-sb {{ background: #2ecc71; }}
                .tag-tg {{ background: #e67e22; }}
                .tag-permiso {{ background: #8e44ad; }}
                .tag-novedad {{ background: #e74c3c; }}
                .tag-sin {{ background: #bdc3c7; color: #2c3e50; }}
                
                .card-body {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    flex-wrap: wrap;
                    gap: 6px;
                    margin-bottom: 6px;
                }}
                .card-horas {{
                    display: flex;
                    align-items: baseline;
                    gap: 8px;
                    flex-wrap: wrap;
                }}
                .card-horas .total {{ font-size: 20px; font-weight: 800; color: #1a2332; }}
                .card-horas .detalle {{ font-size: 11px; color: #5a6a7e; font-weight: 500; }}
                .card-estado {{
                    display: flex;
                    align-items: center;
                    gap: 6px;
                    flex-wrap: wrap;
                }}
                .estado-text {{ font-size: 12px; font-weight: 600; }}
                .permiso-badge {{
                    background: #8e44ad;
                    color: white;
                    padding: 1px 10px;
                    border-radius: 12px;
                    font-size: 10px;
                    font-weight: 700;
                }}
                
                .card-footer {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    padding-top: 6px;
                    border-top: 1px solid rgba(0,0,0,0.05);
                    font-size: 11px;
                    color: #7a8a9e;
                }}
                .dias-info {{ font-weight: 600; }}
                .porcentaje-info {{ font-weight: 600; color: #2c5f8a; }}
                
                .card-detalle {{
                    display: flex;
                    flex-wrap: wrap;
                    gap: 4px;
                    margin-top: 8px;
                    padding-top: 8px;
                    border-top: 1px solid rgba(0,0,0,0.05);
                }}
                .dia-item {{
                    display: inline-block;
                    padding: 1px 8px;
                    border-radius: 10px;
                    font-size: 10px;
                    font-weight: 600;
                    color: white;
                }}
                
                .leyenda-calidad {{
                    background: white;
                    padding: 10px 16px;
                    border-radius: 10px;
                    margin-top: 15px;
                    display: flex;
                    flex-wrap: wrap;
                    gap: 15px;
                    font-size: 11px;
                    align-items: center;
                    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
                }}
                .leyenda-calidad .item {{
                    display: flex;
                    align-items: center;
                    gap: 5px;
                }}
                .leyenda-calidad .color {{
                    display: inline-block;
                    width: 14px;
                    height: 14px;
                    border-radius: 4px;
                }}
                
                .info-novedades {{
                    background: #fef9e7;
                    border: 1px solid #f9e79f;
                    padding: 10px 16px;
                    border-radius: 10px;
                    margin-bottom: 15px;
                    font-size: 12px;
                    color: #7d6608;
                    display: flex;
                    flex-wrap: wrap;
                    gap: 15px;
                    align-items: center;
                }}
                .info-novedades strong {{
                    color: #5a3e0a;
                }}
                
                @media (max-width: 768px) {{
                    body {{ padding: 10px; }}
                    .header {{ flex-direction: column; text-align: center; padding: 14px 16px; }}
                    .header h1 {{ font-size: 17px; justify-content: center; }}
                    .kpi-row {{ grid-template-columns: repeat(2, 1fr); gap: 8px; }}
                    .plat-row {{ grid-template-columns: 1fr; }}
                    .cards-grid {{ grid-template-columns: 1fr; }}
                    .card {{ padding: 12px 14px; }}
                    .kpi .num {{ font-size: 22px; }}
                }}
                @media (max-width: 480px) {{
                    .kpi-row {{ grid-template-columns: 1fr 1fr; }}
                }}
            </style>
        </head>
        <body>
            <div class="container">
                
                <div class="header">
                    <div>
                        <h1>
                            📊 Reporte de Tiempos
                            <span>{total_usuarios}</span>
                        </h1>
                        <div class="rango">📅 {fecha_inicio_str} - {fecha_fin_str}</div>
                        <div class="info">📋 Jornada total: {self.jornada_total_esperada:.0f}h · Sábados: 4h · Festivos: 0h</div>
                    </div>
                    <div class="fecha">👥 Usuarios en Novedades 2: {len(self.usuarios_novedades_2)}</div>
                </div>
                
                <div class="info-novedades">
                    <span>📋 <strong>Usuarios en Novedades 2:</strong> {len(self.usuarios_novedades_2)} personas deben marcar tiempo</span>
                    <span>⛔ <strong>Usuarios NO en Novedades 2:</strong> Están descansando y no aparecen en el reporte</span>
                </div>
                
                <div class="kpi-row">
                    <div class="kpi blue">
                        <div class="num">{total_usuarios}</div>
                        <div class="label">👥 Usuarios</div>
                        <div class="sub">con turno</div>
                    </div>
                    <div class="kpi purple">
                        <div class="num">{con_permiso}</div>
                        <div class="label">📋 Permiso</div>
                    </div>
                    <div class="kpi green">
                        <div class="num">{total_horas:.1f}h</div>
                        <div class="label">⏱️ Total Horas</div>
                    </div>
                    <div class="kpi orange">
                        <div class="num">{promedio:.1f}h</div>
                        <div class="label">📊 Promedio</div>
                    </div>
                </div>
                
                <div class="plat-row">
                    <div class="plat cl">
                        <span class="name">🏛️ Camp Legal</span>
                        <span class="horas">{horas_camp:.1f}h</span>
                    </div>
                    <div class="plat sb">
                        <span class="name">📋 Smokeball</span>
                        <span class="horas">{horas_sb:.1f}h</span>
                    </div>
                    <div class="plat tg">
                        <span class="name">⏱️ Toggl</span>
                        <span class="horas">{horas_tg:.1f}h</span>
                    </div>
                </div>
                
                <div class="cards-grid">
                    {tarjetas_html}
                </div>
                
                <div class="leyenda-calidad">
                    <span style="font-weight:600; color:#1a2332;">📊 Control de Calidad - Desglose Diario:</span>
                    <span class="item"><span class="color" style="background:#27ae60;"></span> ≥ Jornada (Completo)</span>
                    <span class="item"><span class="color" style="background:#f39c12;"></span> 50-100% (Parcial)</span>
                    <span class="item"><span class="color" style="background:#e74c3c;"></span> 1-50% (Bajo)</span>
                    <span class="item"><span class="color" style="background:#bdc3c7;"></span> 0h (Sin registro)</span>
                    <span style="color:#7a8a9e; font-size:10px; margin-left:auto;">📋 Novedad 2 = Debe marcar tiempo</span>
                </div>
                
            </div>
        </body>
        </html>
        """
    
    def _generar_html_vacio(self):
        fecha_inicio_str = self.fecha_inicio.strftime('%d/%m/%Y')
        fecha_fin_str = self.fecha_fin.strftime('%d/%m/%Y')
        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Reporte Tiempos - {fecha_inicio_str}</title>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #eef2f7; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px; }}
                .empty {{ background: white; padding: 40px 50px; border-radius: 16px; text-align: center; max-width: 500px; box-shadow: 0 2px 12px rgba(0,0,0,0.06); }}
                .empty .icon {{ font-size: 56px; }}
                .empty h1 {{ color: #1a3a5c; font-size: 24px; margin: 12px 0; }}
                .empty p {{ color: #7a8a9e; font-size: 16px; }}
                .empty .fecha {{ color: #2c5f8a; font-weight: 600; font-size: 16px; margin-top: 10px; }}
                .empty .info {{ color: #7a8a9e; font-size: 14px; margin-top: 10px; }}
            </style>
        </head>
        <body>
            <div class="empty">
                <div class="icon">📋</div>
                <h1>No hay usuarios con turno</h1>
                <p>No hay personas registradas en <strong>Novedades 2</strong> para este rango de fechas.</p>
                <div class="fecha">📅 {fecha_inicio_str} - {fecha_fin_str}</div>
                <div class="info">💡 Los usuarios que no aparecen en Novedades 2 están descansando.</div>
            </div>
        </body>
        </html>
        """
    
    # ============================================================
    # EJECUTAR
    # ============================================================
    
    def ejecutar(self):
        print("="*70)
        print("📊 SISTEMA DE REPORTE DE TIEMPOS v33.0")
        print("📌 FILTRO POR NOVEDADES 2 (SOLO QUIENES TRABAJAN)")
        print("="*70)
        print(f"📅 Rango: {self.fecha_inicio.strftime('%d/%m/%Y')} - {self.fecha_fin.strftime('%d/%m/%Y')}")
        print("="*70)
        
        if not self.leer_archivos():
            print("❌ Error al leer archivos")
            return False
        
        if not self.construir_mapa_nombres():
            print("❌ Error al construir mapa de nombres")
            return False
        
        self.procesar_novedades()
        
        # Verificar si hay usuarios en Novedades 2
        if not self.usuarios_novedades_2:
            print("\n⚠️ No hay usuarios en Novedades 2. El reporte estará vacío.")
            print("📋 Esto significa que todos los usuarios están descansando en este rango.")
        
        if not self.consolidar_todas_plataformas():
            print("❌ Error al consolidar datos")
            return False
        
        html_body = self.generar_html_reporte()
        
        preview_dir = os.path.join(self.ruta_base, "reporte_previews")
        os.makedirs(preview_dir, exist_ok=True)
        
        filename = os.path.join(preview_dir, f"reporte_novedades2_{self.fecha_inicio_str}_{self.fecha_fin_str}.html")
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_body)
        
        print(f"\n📄 Reporte guardado: {filename}")
        
        import webbrowser
        webbrowser.open(filename)
        print("   🌐 Reporte abierto en el navegador")
        
        return True


# ============================================================
# SELECCIONAR RANGO DE FECHAS
# ============================================================

def seleccionar_rango_fechas():
    print("\n" + "="*70)
    print("📅 SELECCIÓN DE RANGO DE FECHAS")
    print("="*70)
    print("📌 NOTA: Las fechas en los archivos están en formato MM/DD/YYYY")
    print()
    
    hoy = datetime.now().date()
    
    print("Opciones disponibles:")
    print("  1. Ayer (un día)")
    print("  2. Semana actual (Lunes - Hoy)")
    print("  3. Semana pasada (Lunes - Domingo)")
    print("  4. Mes actual (1er día - Hoy)")
    print("  5. Rango personalizado")
    print()
    
    opcion = input("Selecciona (1-5): ")
    
    if opcion == "1":
        fecha_fin = hoy - timedelta(days=1)
        fecha_inicio = fecha_fin
    
    elif opcion == "2":
        dias_desde_lunes = hoy.weekday()
        fecha_inicio = hoy - timedelta(days=dias_desde_lunes)
        fecha_fin = hoy
    
    elif opcion == "3":
        dias_desde_lunes = hoy.weekday()
        lunes_actual = hoy - timedelta(days=dias_desde_lunes)
        fecha_inicio = lunes_actual - timedelta(days=7)
        fecha_fin = fecha_inicio + timedelta(days=6)
    
    elif opcion == "4":
        fecha_inicio = hoy.replace(day=1)
        fecha_fin = hoy
    
    elif opcion == "5":
        fecha_inicio_str = input("Fecha inicio (MM/DD/YYYY): ")
        fecha_fin_str = input("Fecha fin (MM/DD/YYYY): ")
        try:
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%m/%d/%Y").date()
            fecha_fin = datetime.strptime(fecha_fin_str, "%m/%d/%Y").date()
            if fecha_inicio > fecha_fin:
                fecha_inicio, fecha_fin = fecha_fin, fecha_inicio
        except:
            print("❌ Formato incorrecto. Usando fecha de ayer.")
            fecha_fin = hoy - timedelta(days=1)
            fecha_inicio = fecha_fin
    
    else:
        print("❌ Opción no válida. Usando fecha de ayer.")
        fecha_fin = hoy - timedelta(days=1)
        fecha_inicio = fecha_fin
    
    print(f"\n✅ Rango seleccionado: {fecha_inicio.strftime('%d/%m/%Y')} - {fecha_fin.strftime('%d/%m/%Y')}")
    return fecha_inicio, fecha_fin


# ============================================================
# MAIN
# ============================================================

def main():
    try:
        fecha_inicio, fecha_fin = seleccionar_rango_fechas()
        sistema = ReporteTiemposSystem(fecha_inicio=fecha_inicio, fecha_fin=fecha_fin)
        sistema.ejecutar()
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "="*70)
    input("Presiona Enter para salir...")


if __name__ == "__main__":
    main()