#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Procesador de Causación
Módulo para análisis de archivos DIAN y contables
"""

import pandas as pd
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import openpyxl
import numpy as np
from datetime import datetime, date, timedelta
import re
from itertools import combinations
from difflib import SequenceMatcher

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('causacion.log'),
        logging.StreamHandler()
    ]
)

class CausacionProcessor:
    """Clase para procesar archivos de causación DIAN y contables"""
    
    def __init__(self):
        """Inicializar el procesador de causación"""
        self.logger = logging.getLogger(__name__)
        self.dian_data: Optional[pd.DataFrame] = None
        self.contable_data: Optional[pd.DataFrame] = None
        self.dian_file_path: Optional[Path] = None
        self.contable_file_path: Optional[Path] = None
        
        self.logger.info("CausacionProcessor inicializado")
    
    def load_dian_file(self, file_path: str | Path) -> pd.DataFrame:
        """
        Cargar archivo DIAN desde Excel
        
        Args:
            file_path: Ruta del archivo DIAN (str o Path)
            
        Returns:
            DataFrame con los datos del archivo DIAN
            
        Raises:
            FileNotFoundError: Si el archivo no existe
            Exception: Si hay error al leer el archivo
        """
        file_path = Path(file_path)
        
        try:
            # Validar que el archivo existe
            if not file_path.exists():
                raise FileNotFoundError(f"El archivo DIAN no existe: {file_path}")
            
            # Validar extensión
            if file_path.suffix.lower() not in ['.xlsx', '.xls']:
                raise ValueError(f"El archivo debe ser Excel (.xlsx o .xls): {file_path}")
            
            self.logger.info(f"Cargando archivo DIAN: {file_path}")
            
            # Leer el archivo Excel
            df = pd.read_excel(file_path)
            
            # Validar que el DataFrame no esté vacío
            if df.empty:
                raise ValueError("El archivo DIAN está vacío")
            
            # Limpiar datos básicos
            df = self._clean_dataframe(df)
            
            # Aplicar limpieza específica para DIAN
            df = self.clean_dian_data(df)
            
            # Validar calidad de datos
            quality_report = self.validate_data_quality(df, 'DIAN')
            if not quality_report['is_valid']:
                self.logger.warning(f"Problemas de calidad en archivo DIAN: Score {quality_report['overall_score']:.1f}")
            
            # Guardar referencia
            self.dian_data = df
            self.dian_file_path = file_path
            
            self.logger.info(f"Archivo DIAN cargado exitosamente: {len(df)} filas, {len(df.columns)} columnas")
            
            return df
            
        except FileNotFoundError:
            self.logger.error(f"Archivo DIAN no encontrado: {file_path}")
            raise
        except Exception as e:
            self.logger.error(f"Error al cargar archivo DIAN: {e}")
            raise Exception(f"Error al cargar archivo DIAN: {e}")
    
    def load_contable_file(self, file_path: str | Path) -> pd.DataFrame:
        """
        Cargar archivo contable desde Excel
        
        Args:
            file_path: Ruta del archivo contable (str o Path)
            
        Returns:
            DataFrame con los datos del archivo contable
            
        Raises:
            FileNotFoundError: Si el archivo no existe
            Exception: Si hay error al leer el archivo
        """
        file_path = Path(file_path)
        
        try:
            # Validar que el archivo existe
            if not file_path.exists():
                raise FileNotFoundError(f"El archivo contable no existe: {file_path}")
            
            # Validar extensión
            if file_path.suffix.lower() not in ['.xlsx', '.xls']:
                raise ValueError(f"El archivo debe ser Excel (.xlsx o .xls): {file_path}")
            
            self.logger.info(f"Cargando archivo contable: {file_path}")
            
            # Leer el archivo Excel
            df = pd.read_excel(file_path)
            
            # Validar que el DataFrame no esté vacío
            if df.empty:
                raise ValueError("El archivo contable está vacío")
            
            # Limpiar datos básicos
            df = self._clean_dataframe(df)
            
            # Aplicar limpieza específica para contable
            df = self.clean_contable_data(df)
            
            # Validar calidad de datos
            quality_report = self.validate_data_quality(df, 'contable')
            if not quality_report['is_valid']:
                self.logger.warning(f"Problemas de calidad en archivo contable: Score {quality_report['overall_score']:.1f}")
            
            # Guardar referencia
            self.contable_data = df
            self.contable_file_path = file_path
            
            self.logger.info(f"Archivo contable cargado exitosamente: {len(df)} filas, {len(df.columns)} columnas")
            
            return df
            
        except FileNotFoundError:
            self.logger.error(f"Archivo contable no encontrado: {file_path}")
            raise
        except Exception as e:
            self.logger.error(f"Error al cargar archivo contable: {e}")
            raise Exception(f"Error al cargar archivo contable: {e}")
    
    def validate_files(self) -> Tuple[bool, List[str]]:
        """
        Validar que los archivos existen y son válidos
        
        Returns:
            Tuple con (es_válido, lista_de_errores)
        """
        errors = []
        
        # Validar archivo DIAN
        if self.dian_data is None:
            errors.append("Archivo DIAN no ha sido cargado")
        elif self.dian_data.empty:
            errors.append("Archivo DIAN está vacío")
        
        # Validar archivo contable
        if self.contable_data is None:
            errors.append("Archivo contable no ha sido cargado")
        elif self.contable_data.empty:
            errors.append("Archivo contable está vacío")
        
        # Validar que ambos archivos estén cargados
        if self.dian_data is not None and self.contable_data is not None:
            self.logger.info("Validación de archivos completada exitosamente")
        else:
            self.logger.warning(f"Validación de archivos falló: {errors}")
        
        return len(errors) == 0, errors
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Limpiar DataFrame básico
        
        Args:
            df: DataFrame a limpiar
            
        Returns:
            DataFrame limpio
        """
        # Crear copia para no modificar el original
        clean_df = df.copy()
        
        # Eliminar filas completamente vacías
        clean_df = clean_df.dropna(how='all')
        
        # Eliminar espacios en blanco de columnas de texto
        for col in clean_df.select_dtypes(include=['object']).columns:
            clean_df[col] = clean_df[col].astype(str).str.strip()
        
        # Resetear índices
        clean_df = clean_df.reset_index(drop=True)
        
        return clean_df
    
    def clean_dian_data(self, df_dian: pd.DataFrame) -> pd.DataFrame:
        """
        Limpiar y procesar datos del archivo DIAN
        
        Args:
            df_dian: DataFrame con datos DIAN sin procesar
            
        Returns:
            DataFrame con datos DIAN limpios y procesados
        """
        self.logger.info("Iniciando limpieza de datos DIAN")
        
        try:
            # Crear copia para no modificar el original
            clean_df = df_dian.copy()
            
            # 1. Limpiar espacios en blanco de todas las columnas de texto
            self.logger.info("Limpiando espacios en blanco...")
            for col in clean_df.select_dtypes(include=['object']).columns:
                clean_df[col] = clean_df[col].astype(str).str.strip()
            
            # 2. Convertir 'Folio' a string (asumiendo que existe la columna)
            folio_columns = [col for col in clean_df.columns if 'folio' in col.lower()]
            for col in folio_columns:
                self.logger.info(f"Convirtiendo columna '{col}' a string...")
                clean_df[col] = clean_df[col].astype(str).str.strip()
            
            # 3. Formatear fechas (buscar columnas que contengan 'fecha' o 'date')
            date_columns = [col for col in clean_df.columns 
                          if any(keyword in col.lower() for keyword in ['fecha', 'date', 'dia', 'mes', 'año'])]
            
            for col in date_columns:
                self.logger.info(f"Procesando columna de fecha: {col}")
                clean_df[col] = self._format_date_column(clean_df[col])
            
            # 4. Validar campos críticos
            critical_fields = self._identify_critical_fields(clean_df)
            validation_result = self._validate_critical_fields(clean_df, critical_fields)
            
            if not validation_result['is_valid']:
                self.logger.warning(f"Campos críticos con problemas: {validation_result['errors']}")
            
            # 5. Eliminar filas completamente vacías
            initial_rows = len(clean_df)
            clean_df = clean_df.dropna(how='all')
            final_rows = len(clean_df)
            
            if initial_rows != final_rows:
                self.logger.info(f"Eliminadas {initial_rows - final_rows} filas vacías")
            
            # 6. Resetear índices
            clean_df = clean_df.reset_index(drop=True)
            
            self.logger.info(f"Limpieza DIAN completada: {len(clean_df)} filas, {len(clean_df.columns)} columnas")
            
            return clean_df
            
        except Exception as e:
            self.logger.error(f"Error en limpieza de datos DIAN: {e}")
            raise Exception(f"Error en limpieza de datos DIAN: {e}")
    
    def clean_contable_data(self, df_contable: pd.DataFrame) -> pd.DataFrame:
        """
        Limpiar y procesar datos del archivo contable
        
        Args:
            df_contable: DataFrame con datos contables sin procesar
            
        Returns:
            DataFrame con datos contables limpios y procesados
        """
        self.logger.info("Iniciando limpieza de datos contables")
        
        try:
            # Crear copia para no modificar el original
            clean_df = df_contable.copy()
            
            # 1. Saltar primeras 4 filas de metadatos
            self.logger.info("Eliminando filas de metadatos...")
            if len(clean_df) > 4:
                clean_df = clean_df.iloc[4:].reset_index(drop=True)
                self.logger.info("Eliminadas 4 filas de metadatos")
            
            # 2. Mapear columnas 'Unnamed' a nombres descriptivos
            self.logger.info("Mapeando columnas sin nombre...")
            clean_df = self._map_unnamed_columns(clean_df)
            
            # 3. Combinar Año/Mes/Día en fecha única
            clean_df = self._combine_date_columns(clean_df)
            
            # 4. Limpiar datos numéricos
            clean_df = self._clean_numeric_data(clean_df)
            
            # 5. Limpiar espacios en blanco
            for col in clean_df.select_dtypes(include=['object']).columns:
                try:
                    # Verificar si la columna existe y es una Serie
                    if col in clean_df.columns and isinstance(clean_df[col], pd.Series):
                        clean_df[col] = clean_df[col].astype(str).str.strip()
                    else:
                        # Si hay columnas duplicadas, manejar cada una individualmente
                        if col in clean_df.columns:
                            clean_df[col] = clean_df[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
                except Exception as e:
                    self.logger.warning(f"No se pudo limpiar la columna '{col}': {e}")
                    continue
            
            # 6. Eliminar filas completamente vacías
            initial_rows = len(clean_df)
            clean_df = clean_df.dropna(how='all')
            final_rows = len(clean_df)
            
            if initial_rows != final_rows:
                self.logger.info(f"Eliminadas {initial_rows - final_rows} filas vacías")
            
            # 7. Resetear índices
            clean_df = clean_df.reset_index(drop=True)
            
            self.logger.info(f"Limpieza contable completada: {len(clean_df)} filas, {len(clean_df.columns)} columnas")
            
            return clean_df
            
        except Exception as e:
            self.logger.error(f"Error en limpieza de datos contables: {e}")
            raise Exception(f"Error en limpieza de datos contables: {e}")
    
    def validate_data_quality(self, df: pd.DataFrame, source: str) -> Dict[str, Any]:
        """
        Validar calidad de datos del DataFrame
        
        Args:
            df: DataFrame a validar
            source: Fuente de datos ('DIAN' o 'contable')
            
        Returns:
            Diccionario con resultados de validación
        """
        self.logger.info(f"Iniciando validación de calidad para {source}")
        
        validation_results = {
            'source': source,
            'total_rows': len(df),
            'total_columns': len(df.columns),
            'missing_values': {},
            'date_format_issues': [],
            'numeric_format_issues': [],
            'critical_field_issues': [],
            'overall_score': 0.0,
            'is_valid': True
        }
        
        try:
            # 1. Verificar integridad de datos
            for col in df.columns:
                missing_count = df[col].isna().sum()
                missing_percentage = (missing_count / len(df)) * 100
                
                validation_results['missing_values'][col] = {
                    'count': missing_count,
                    'percentage': missing_percentage
                }
                
                if missing_percentage > 50:  # Más del 50% de valores faltantes
                    validation_results['critical_field_issues'].append(
                        f"Columna '{col}' tiene {missing_percentage:.1f}% de valores faltantes"
                    )
            
            # 2. Validar formatos de fecha
            date_columns = [col for col in df.columns 
                          if any(keyword in col.lower() for keyword in ['fecha', 'date', 'dia', 'mes', 'año'])]
            
            for col in date_columns:
                date_issues = self._validate_date_format(df[col], col)
                validation_results['date_format_issues'].extend(date_issues)
            
            # 3. Validar formatos numéricos
            numeric_columns = df.select_dtypes(include=[np.number]).columns
            for col in numeric_columns:
                numeric_issues = self._validate_numeric_format(df[col], col)
                validation_results['numeric_format_issues'].extend(numeric_issues)
            
            # 4. Calcular score general
            total_issues = (len(validation_results['critical_field_issues']) + 
                          len(validation_results['date_format_issues']) + 
                          len(validation_results['numeric_format_issues']))
            
            validation_results['overall_score'] = max(0, 100 - (total_issues * 10))
            validation_results['is_valid'] = validation_results['overall_score'] >= 70
            
            # 5. Logging de resultados
            if validation_results['is_valid']:
                self.logger.info(f"Validación {source} exitosa - Score: {validation_results['overall_score']:.1f}")
            else:
                self.logger.warning(f"Validación {source} con problemas - Score: {validation_results['overall_score']:.1f}")
                self.logger.warning(f"Problemas encontrados: {total_issues}")
            
            return validation_results
            
        except Exception as e:
            self.logger.error(f"Error en validación de calidad {source}: {e}")
            validation_results['is_valid'] = False
            validation_results['error'] = str(e)
            return validation_results
    
    def _format_date_column(self, date_series: pd.Series) -> pd.Series:
        """
        Formatear columna de fecha a formato DD-MM-YYYY
        
        Args:
            date_series: Serie con fechas
            
        Returns:
            Serie con fechas formateadas
        """
        try:
            # Intentar convertir a datetime con formato específico
            # Primero intentar con formato DD-MM-YYYY
            try:
                formatted_dates = pd.to_datetime(date_series, format='%d-%m-%Y', errors='coerce')
            except:
                # Si falla, intentar con formato DD-MM-YYYY HH:MM:SS
                try:
                    formatted_dates = pd.to_datetime(date_series, format='%d-%m-%Y %H:%M:%S', errors='coerce')
                except:
                    # Si falla, usar inferencia automática con dayfirst=True
                    formatted_dates = pd.to_datetime(date_series, dayfirst=True, errors='coerce')
            
            # Formatear como DD-MM-YYYY
            formatted_dates = formatted_dates.dt.strftime('%d-%m-%Y')
            
            # Reemplazar NaT con None
            formatted_dates = formatted_dates.replace('NaT', None)
            
            return formatted_dates
            
        except Exception as e:
            self.logger.warning(f"Error al formatear fechas: {e}")
            return date_series
    
    def _identify_critical_fields(self, df: pd.DataFrame) -> List[str]:
        """
        Identificar campos críticos basado en nombres de columnas
        
        Args:
            df: DataFrame a analizar
            
        Returns:
            Lista de nombres de columnas críticas
        """
        critical_keywords = ['folio', 'numero', 'identificacion', 'nit', 'ruc', 'fecha', 'valor', 'monto']
        critical_fields = []
        
        for col in df.columns:
            if any(keyword in col.lower() for keyword in critical_keywords):
                critical_fields.append(col)
        
        return critical_fields
    
    def _validate_critical_fields(self, df: pd.DataFrame, critical_fields: List[str]) -> Dict[str, Any]:
        """
        Validar que campos críticos no estén vacíos
        
        Args:
            df: DataFrame a validar
            critical_fields: Lista de campos críticos
            
        Returns:
            Diccionario con resultados de validación
        """
        errors = []
        
        for field in critical_fields:
            if field in df.columns:
                missing_count = df[field].isna().sum()
                if missing_count > 0:
                    errors.append(f"Campo '{field}' tiene {missing_count} valores faltantes")
        
        return {
            'is_valid': len(errors) == 0,
            'errors': errors
        }
    
    def _map_unnamed_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Mapear columnas 'Unnamed' a nombres descriptivos basado en contenido
        
        Args:
            df: DataFrame con columnas sin nombre
            
        Returns:
            DataFrame con columnas renombradas
        """
        unnamed_columns = [col for col in df.columns if 'unnamed' in col.lower()]
        
        if unnamed_columns:
            self.logger.info(f"Mapeando {len(unnamed_columns)} columnas sin nombre")
            
            mapping = {}
            
            # Mapeo inteligente basado en análisis de contenido
            for col in unnamed_columns:
                try:
                    sample_values = df[col].dropna().head(10)
                    if len(sample_values) == 0:
                        continue
                        
                    # Analizar el contenido para determinar el tipo de columna
                    first_vals = sample_values.head(3).astype(str).tolist()
                    
                    # Verificar si son números de documento (números largos)
                    if self._looks_like_document_numbers(sample_values):
                        mapping[col] = 'numero_documento'
                        self.logger.info(f"Columna {col} mapeada como 'numero_documento' - valores ejemplo: {first_vals}")
                        
                    # Verificar si son valores monetarios (números con decimales o grandes)
                    elif self._looks_like_monetary_values(sample_values):
                        mapping[col] = 'valor'
                        self.logger.info(f"Columna {col} mapeada como 'valor' - valores ejemplo: {first_vals}")
                        
                    # Verificar si son fechas
                    elif self._looks_like_dates(sample_values):
                        mapping[col] = 'fecha'
                        self.logger.info(f"Columna {col} mapeada como 'fecha' - valores ejemplo: {first_vals}")
                        
                    # Verificar si son códigos de cuenta (números de cuenta contable)
                    elif self._looks_like_account_codes(sample_values):
                        mapping[col] = 'cuenta_contable'
                        self.logger.info(f"Columna {col} mapeada como 'cuenta_contable' - valores ejemplo: {first_vals}")
                        
                    # Verificar si son descripciones (texto largo)
                    elif self._looks_like_descriptions(sample_values):
                        mapping[col] = 'descripcion'
                        self.logger.info(f"Columna {col} mapeada como 'descripcion' - valores ejemplo: {first_vals}")
                        
                except Exception as e:
                    self.logger.warning(f"Error al analizar columna {col}: {e}")
                    continue
            
            # Aplicar mapeo evitando duplicados
            if mapping:
                # Evitar nombres duplicados agregando sufijos
                final_mapping = {}
                used_names = set()
                
                for original_col, target_name in mapping.items():
                    if target_name not in used_names:
                        final_mapping[original_col] = target_name
                        used_names.add(target_name)
                    else:
                        # Agregar sufijo para evitar duplicados
                        counter = 2
                        new_name = f"{target_name}_{counter}"
                        while new_name in used_names:
                            counter += 1
                            new_name = f"{target_name}_{counter}"
                        final_mapping[original_col] = new_name
                        used_names.add(new_name)
                
                df = df.rename(columns=final_mapping)
                self.logger.info(f"Columnas renombradas exitosamente: {list(final_mapping.values())}")
            else:
                self.logger.warning("No se pudo mapear ninguna columna automáticamente")
        
        return df
    
    def _looks_like_document_numbers(self, values: pd.Series) -> bool:
        """Verificar si los valores parecen números de documento"""
        try:
            for val in values.head(5):
                val_str = str(val).strip()
                if val_str.isdigit() and len(val_str) >= 6:
                    return True
            return False
        except:
            return False
    
    def _looks_like_monetary_values(self, values: pd.Series) -> bool:
        """Verificar si los valores parecen valores monetarios"""
        try:
            numeric_count = 0
            for val in values.head(5):
                try:
                    num_val = float(val)
                    if num_val > 1000 or (num_val > 0 and num_val < 1000000000):  # Rango razonable para valores monetarios
                        numeric_count += 1
                except:
                    continue
            return numeric_count >= len(values.head(5)) * 0.6
        except:
            return False
    
    def _looks_like_dates(self, values: pd.Series) -> bool:
        """Verificar si los valores parecen fechas"""
        try:
            date_count = 0 
            for val in values.head(5):
                val_str = str(val).strip()
                # Verificar patrones comunes de fecha
                if (len(val_str) >= 8 and 
                    (('/' in val_str) or ('-' in val_str) or 
                     (val_str.isdigit() and len(val_str) == 8))):  # YYYYMMDD
                    date_count += 1
            return date_count >= len(values.head(5)) * 0.6
        except:
            return False
    
    def _looks_like_account_codes(self, values: pd.Series) -> bool:
        """Verificar si los valores parecen códigos de cuenta contable"""
        try:
            for val in values.head(5):
                val_str = str(val).strip()
                # Códigos de cuenta suelen ser números de 4-10 dígitos
                if val_str.isdigit() and 4 <= len(val_str) <= 10:
                    return True
            return False
        except:
            return False
    
    def _looks_like_descriptions(self, values: pd.Series) -> bool:
        """Verificar si los valores parecen descripciones de texto"""
        try:
            text_count = 0
            for val in values.head(5):
                val_str = str(val).strip()
                # Descripciones suelen tener espacios y más de 10 caracteres
                if len(val_str) > 10 and (' ' in val_str or len(val_str) > 20):
                    text_count += 1
            return text_count >= len(values.head(5)) * 0.6
        except:
            return False
    
    def _get_contable_column_mapping(self) -> List[str]:
        """
        Obtener mapeo sugerido para columnas contables
        
        Returns:
            Lista de nombres sugeridos para columnas
        """
        return [
            'fecha_transaccion',
            'numero_documento',
            'descripcion',
            'debito',
            'credito',
            'saldo',
            'cuenta_contable',
            'centro_costo'
        ]
    
    def _combine_date_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Combinar columnas de Año/Mes/Día en fecha única
        
        Args:
            df: DataFrame con columnas separadas de fecha
            
        Returns:
            DataFrame con fecha combinada
        """
        # Buscar columnas de año, mes, día
        year_cols = [col for col in df.columns if 'año' in col.lower() or 'year' in col.lower()]
        month_cols = [col for col in df.columns if 'mes' in col.lower() or 'month' in col.lower()]
        day_cols = [col for col in df.columns if 'dia' in col.lower() or 'day' in col.lower()]
        
        if year_cols and month_cols and day_cols:
            self.logger.info("Combinando columnas de fecha...")
            
            year_col = year_cols[0]
            month_col = month_cols[0]
            day_col = day_cols[0]
            
            # Crear fecha combinada
            df['fecha_combinada'] = pd.to_datetime(
                df[year_col].astype(str) + '-' + 
                df[month_col].astype(str).str.zfill(2) + '-' + 
                df[day_col].astype(str).str.zfill(2),
                errors='coerce'
            )
            
            # Formatear como DD-MM-YYYY
            df['fecha_combinada'] = df['fecha_combinada'].dt.strftime('%d-%m-%Y')
        
        return df
    
    def _clean_numeric_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Limpiar datos numéricos
        
        Args:
            df: DataFrame con datos numéricos
            
        Returns:
            DataFrame con datos numéricos limpios
        """
        # Identificar columnas numéricas
        numeric_columns = df.select_dtypes(include=[np.number]).columns
        
        for col in numeric_columns:
            # Reemplazar valores infinitos con NaN
            df[col] = df[col].replace([np.inf, -np.inf], np.nan)
            
            # Redondear a 2 decimales si es necesario
            if df[col].dtype in ['float64', 'float32']:
                df[col] = df[col].round(2)
        
        return df
    
    def _validate_date_format(self, date_series: pd.Series, column_name: str) -> List[str]:
        """
        Validar formato de fechas en una columna
        
        Args:
            date_series: Serie con fechas
            column_name: Nombre de la columna
            
        Returns:
            Lista de errores encontrados
        """
        issues = []
        
        try:
            # Intentar convertir a datetime con formato específico
            try:
                pd.to_datetime(date_series, format='%d-%m-%Y', errors='raise')
            except:
                try:
                    pd.to_datetime(date_series, format='%d-%m-%Y %H:%M:%S', errors='raise')
                except:
                    pd.to_datetime(date_series, dayfirst=True, errors='raise')
        except Exception as e:
            issues.append(f"Columna '{column_name}': Error en formato de fecha - {str(e)}")
        
        return issues
    
    def _validate_numeric_format(self, numeric_series: pd.Series, column_name: str) -> List[str]:
        """
        Validar formato numérico en una columna
        
        Args:
            numeric_series: Serie numérica
            column_name: Nombre de la columna
            
        Returns:
            Lista de errores encontrados
        """
        issues = []
        
        # Verificar valores infinitos
        if np.isinf(numeric_series).any():
            issues.append(f"Columna '{column_name}': Contiene valores infinitos")
        
        # Verificar valores muy grandes (posibles errores)
        if numeric_series.max() > 1e12:
            issues.append(f"Columna '{column_name}': Contiene valores muy grandes")
        
        return issues
    
    def perform_data_matching(self, df_dian: pd.DataFrame, df_contable: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """
        Realizar cruce de datos entre archivos DIAN y contables
        
        Args:
            df_dian: DataFrame con datos DIAN limpios
            df_contable: DataFrame con datos contables limpios
            
        Returns:
            Diccionario con DataFrames de coincidencias y no coincidencias
        """
        self.logger.info("Iniciando cruce de datos DIAN vs Contable")
        
        try:
            # Validar que los DataFrames no estén vacíos
            if df_dian.empty or df_contable.empty:
                raise ValueError("Uno o ambos DataFrames están vacíos")
            
            self.logger.info(f"Procesando {len(df_dian)} registros DIAN vs {len(df_contable)} registros contables")
            
            # Identificar columnas de cruce
            dian_doc_col = self._find_document_column(df_dian, 'DIAN')
            contable_doc_col = self._find_document_column(df_contable, 'contable')
            
            self.logger.info(f"Columnas de cruce: DIAN='{dian_doc_col}' vs Contable='{contable_doc_col}'")
            
            # Realizar matching
            matches, non_matches = self.identify_matches(df_dian, df_contable, dian_doc_col, contable_doc_col)
            
            # Generar reporte
            report = self.generate_matching_report(matches, non_matches)
            
            self.logger.info(f"Cruce completado: {len(matches)} coincidencias, {len(non_matches)} no coincidencias")
            
            return {
                'matches': matches,
                'non_matches': non_matches,
                'report': report
            }
            
        except Exception as e:
            self.logger.error(f"Error en cruce de datos: {e}")
            raise Exception(f"Error en cruce de datos: {e}")
    
    def identify_matches(self, df_dian: pd.DataFrame, df_contable: pd.DataFrame, 
                        dian_doc_col: str, contable_doc_col: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Identificar coincidencias entre registros DIAN y contables
        
        Args:
            df_dian: DataFrame DIAN
            df_contable: DataFrame contable
            dian_doc_col: Columna de documento en DIAN
            contable_doc_col: Columna de documento en contable
            
        Returns:
            Tuple con (DataFrame de coincidencias, DataFrame de no coincidencias)
        """
        self.logger.info("Iniciando identificación de coincidencias")
        
        try:
            # Crear copias para no modificar los originales
            dian_df = df_dian.copy()
            contable_df = df_contable.copy()
            
            # Agregar columnas de scoring
            dian_df['match_score'] = 0.0
            dian_df['match_type'] = 'no_match'
            dian_df['matched_contable_id'] = None
            
            contable_df['match_score'] = 0.0
            contable_df['match_type'] = 'no_match'
            contable_df['matched_dian_id'] = None
            
            # 1. Cruce primario por número de documento (exacto)
            self.logger.info("Realizando cruce primario por documento...")
            exact_matches = self._find_exact_document_matches(dian_df, contable_df, dian_doc_col, contable_doc_col)
            
            # 2. Cruce secundario por valor y fecha
            self.logger.info("Realizando cruce secundario por valor y fecha...")
            secondary_matches = self._find_secondary_matches(dian_df, contable_df, exact_matches)
            
            # 3. Cruce por similitud de texto
            self.logger.info("Realizando cruce por similitud...")
            similarity_matches = self._find_similarity_matches(dian_df, contable_df, exact_matches + secondary_matches)
            
            # Combinar todos los matches
            all_matches = exact_matches + secondary_matches + similarity_matches
            
            # Crear DataFrame de coincidencias
            matches_df = self._create_matches_dataframe(dian_df, contable_df, all_matches)
            
            # Crear DataFrame de no coincidencias
            non_matches_df = self._create_non_matches_dataframe(dian_df, contable_df, all_matches)
            
            self.logger.info(f"Identificadas {len(matches_df)} coincidencias y {len(non_matches_df)} no coincidencias")
            
            return matches_df, non_matches_df
            
        except Exception as e:
            self.logger.error(f"Error en identificación de coincidencias: {e}")
            raise Exception(f"Error en identificación de coincidencias: {e}")
    
    def generate_matching_report(self, matches: pd.DataFrame, non_matches: pd.DataFrame) -> Dict[str, Any]:
        """
        Generar reporte de análisis de matching
        
        Args:
            matches: DataFrame con coincidencias
            non_matches: DataFrame con no coincidencias
            
        Returns:
            Diccionario con estadísticas del matching
        """
        self.logger.info("Generando reporte de matching")
        
        try:
            report = {
                'total_dian_records': len(matches) + len(non_matches[non_matches['source'] == 'DIAN']),
                'total_contable_records': len(matches) + len(non_matches[non_matches['source'] == 'contable']),
                'total_matches': len(matches),
                'total_non_matches': len(non_matches),
                'match_rate': 0.0,
                'match_breakdown': {},
                'quality_metrics': {},
                'discrepancies': []
            }
            
            # Calcular tasa de matching
            if report['total_dian_records'] > 0:
                report['match_rate'] = (report['total_matches'] / report['total_dian_records']) * 100
            
            # Desglose por tipo de match
            if not matches.empty and 'match_type' in matches.columns:
                match_types = matches['match_type'].value_counts()
                report['match_breakdown'] = match_types.to_dict()
            
            # Métricas de calidad
            if not matches.empty:
                report['quality_metrics'] = {
                    'avg_match_score': matches['match_score'].mean() if 'match_score' in matches.columns else 0.0,
                    'high_confidence_matches': len(matches[matches['match_score'] >= 0.8]) if 'match_score' in matches.columns else 0,
                    'medium_confidence_matches': len(matches[(matches['match_score'] >= 0.6) & (matches['match_score'] < 0.8)]) if 'match_score' in matches.columns else 0,
                    'low_confidence_matches': len(matches[matches['match_score'] < 0.6]) if 'match_score' in matches.columns else 0
                }
            
            # Análisis de discrepancias
            report['discrepancies'] = self._analyze_discrepancies(matches)
            
            # Logging de resultados
            self.logger.info(f"Tasa de matching: {report['match_rate']:.2f}%")
            self.logger.info(f"Coincidencias de alta confianza: {report['quality_metrics'].get('high_confidence_matches', 0)}")
            self.logger.info(f"Discrepancias encontradas: {len(report['discrepancies'])}")
            
            return report
            
        except Exception as e:
            self.logger.error(f"Error en generación de reporte: {e}")
            return {'error': str(e)}
    
    def _find_document_column(self, df: pd.DataFrame, source: str) -> str:
        """
        Encontrar columna de documento en el DataFrame
        
        Args:
            df: DataFrame a analizar
            source: Fuente de datos ('DIAN' o 'contable')
            
        Returns:
            Nombre de la columna de documento
        """
        if source == 'DIAN':
            # Para archivos DIAN, buscar específicamente 'Folio' primero
            if 'Folio' in df.columns:
                self.logger.info(f"Usando columna 'Folio' para documento DIAN")
                return 'Folio'
            
            # Buscar otras palabras clave
            keywords = ['folio', 'numero', 'documento', 'factura']
            for col in df.columns:
                col_lower = col.lower()
                if any(keyword in col_lower for keyword in keywords):
                    # Evitar 'Tipo de documento' que no contiene números
                    if 'tipo' not in col_lower:
                        self.logger.info(f"Usando columna '{col}' para documento DIAN")
                        return col
                        
        else:  # contable
            # Para archivos contables, buscar columna que contenga valores DIAN conocidos
            if hasattr(self, 'dian_data') and self.dian_data is not None and 'Folio' in self.dian_data.columns:
                # Obtener algunos folios DIAN para buscar coincidencias
                folios_dian = self.dian_data['Folio'].dropna().astype(str).head(20).tolist()
                
                best_match_col = None
                max_matches = 0
                
                for col in df.columns:
                    try:
                        # Convertir la columna a string para búsqueda
                        col_values = df[col].astype(str)
                        matches_found = 0
                        
                        # Contar cuántos folios DIAN se encuentran en esta columna
                        for folio in folios_dian:
                            if col_values.str.contains(folio, na=False).any():
                                matches_found += 1
                        
                        # Si encontramos coincidencias, esta podría ser la columna correcta
                        if matches_found > max_matches:
                            max_matches = matches_found
                            best_match_col = col
                            
                    except Exception as e:
                        continue
                
                # Si encontramos una columna con coincidencias, usarla
                if best_match_col and max_matches > 0:
                    self.logger.info(f"Usando columna '{best_match_col}' para documento contable (encontradas {max_matches} coincidencias con folios DIAN)")
                    return best_match_col
            
            # Método de respaldo: buscar por contenido numérico
            numeric_cols = []
            for col in df.columns:
                try:
                    # Verificar si la columna contiene valores que parecen números de documento
                    sample_values = df[col].dropna().head(10)
                    if len(sample_values) > 0:
                        # Convertir a string y verificar si son números largos (posibles documentos)
                        str_values = sample_values.astype(str)
                        numeric_count = 0
                        for val in str_values:
                            val_clean = val.strip()
                            if val_clean.isdigit() and len(val_clean) >= 6:  # Números de al menos 6 dígitos
                                numeric_count += 1
                        
                        if numeric_count >= len(str_values) * 0.8:  # Al menos 80% son números largos
                            numeric_cols.append((col, numeric_count))
                except:
                    continue
            
            # Si encontramos columnas con números largos, usar la primera
            if numeric_cols:
                # Ordenar por cantidad de números válidos
                numeric_cols.sort(key=lambda x: x[1], reverse=True)
                col_name = numeric_cols[0][0]
                self.logger.info(f"Usando columna '{col_name}' para documento contable (detectada por contenido numérico)")
                return col_name
            
            # Buscar por palabras clave como fallback
            keywords = ['numero', 'documento', 'cruce', 'factura', 'comprobante']
            for col in df.columns:
                col_lower = col.lower()
                if any(keyword in col_lower for keyword in keywords):
                    self.logger.info(f"Usando columna '{col}' para documento contable")
                    return col
            
            # Como último recurso, buscar columnas 'Unnamed' con contenido numérico
            unnamed_cols = [col for col in df.columns if 'unnamed' in col.lower()]
            for col in unnamed_cols:
                try:
                    sample_values = df[col].dropna().head(5)
                    if len(sample_values) > 0:
                        # Verificar si parecen números de documento
                        first_val = str(sample_values.iloc[0]).strip()
                        if first_val.isdigit() and len(first_val) >= 6:
                            self.logger.info(f"Usando columna '{col}' para documento contable (Unnamed con números)")
                            return col
                except:
                    continue
        
        # Fallback: usar la primera columna disponible
        if len(df.columns) > 0:
            fallback_col = df.columns[0]
            self.logger.warning(f"No se encontró columna de documento obvia para {source}, usando '{fallback_col}' como fallback")
            return fallback_col
            
        return None
    
    def _find_exact_document_matches(self, dian_df: pd.DataFrame, contable_df: pd.DataFrame, 
                                   dian_col: str, contable_col: str) -> List[Dict[str, Any]]:
        """
        Encontrar coincidencias exactas por número de documento
        
        Args:
            dian_df: DataFrame DIAN
            contable_df: DataFrame contable
            dian_col: Columna de documento en DIAN
            contable_col: Columna de documento en contable
            
        Returns:
            Lista de coincidencias exactas
        """
        matches = []
        
        # Normalizar columnas de documento
        dian_docs = dian_df[dian_col].astype(str).str.strip().str.upper()
        contable_docs = contable_df[contable_col].astype(str).str.strip().str.upper()
        
        # Crear índice para búsqueda eficiente
        contable_doc_index = contable_docs.to_dict()
        
        for dian_idx, dian_doc in dian_docs.items():
            if dian_doc in contable_doc_index.values():
                # Encontrar todos los índices contables que coinciden
                contable_indices = [idx for idx, doc in contable_doc_index.items() if doc == dian_doc]
                
                for contable_idx in contable_indices:
                    match = {
                        'dian_idx': dian_idx,
                        'contable_idx': contable_idx,
                        'match_type': 'exact_document',
                        'match_score': 1.0,
                        'match_reason': f'Documento exacto: {dian_doc}'
                    }
                    matches.append(match)
        
        self.logger.info(f"Encontradas {len(matches)} coincidencias exactas por documento")
        return matches
    
    def _find_secondary_matches(self, dian_df: pd.DataFrame, contable_df: pd.DataFrame, 
                              existing_matches: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Encontrar coincidencias secundarias por valor y fecha (optimizado)
        
        Args:
            dian_df: DataFrame DIAN
            contable_df: DataFrame contable
            existing_matches: Lista de coincidencias existentes
            
        Returns:
            Lista de coincidencias secundarias
        """
        matches = []
        
        # Obtener índices ya emparejados
        matched_dian_indices = {match['dian_idx'] for match in existing_matches}
        matched_contable_indices = {match['contable_idx'] for match in existing_matches}
        
        # Encontrar columnas de valor y fecha
        dian_value_col = self._find_value_column(dian_df, 'DIAN')
        contable_value_col = self._find_value_column(contable_df, 'contable')
        dian_date_col = self._find_date_column(dian_df, 'DIAN')
        contable_date_col = self._find_date_column(contable_df, 'contable')
        
        if not (dian_value_col and contable_value_col):
            self.logger.info("No se encontraron columnas de valor para cruce secundario")
            return matches
        
        # Filtrar registros no emparejados
        dian_unmatched = dian_df[~dian_df.index.isin(matched_dian_indices)]
        contable_unmatched = contable_df[~contable_df.index.isin(matched_contable_indices)]
        
        if dian_unmatched.empty or contable_unmatched.empty:
            self.logger.info("No hay registros sin emparejar para cruce secundario")
            return matches
        
        # Tolerancia para diferencias en valores (5%)
        tolerance = 0.05
        
        # Crear índices por valor para optimizar búsqueda
        self.logger.info("Creando índices de valor para optimización...")
        
        # Agrupar contables por valor (aproximado)
        contable_by_value = {}
        for idx, row in contable_unmatched.iterrows():
            value = row[contable_value_col]
            numeric_value = self._safe_to_numeric(value)
            if numeric_value is not None:
                # Redondear a 2 decimales para agrupar valores similares
                rounded_value = round(numeric_value, 2)
                if rounded_value not in contable_by_value:
                    contable_by_value[rounded_value] = []
                contable_by_value[rounded_value].append(idx)
        
        self.logger.info(f"Índices creados para {len(contable_by_value)} valores únicos")
        
        # Buscar coincidencias
        processed_count = 0
        total_dian = len(dian_unmatched)
        
        for dian_idx, dian_row in dian_unmatched.iterrows():
            processed_count += 1
            if processed_count % 100 == 0:
                self.logger.info(f"Procesando DIAN: {processed_count}/{total_dian}")
            
            dian_value = dian_row[dian_value_col]
            dian_date = dian_row[dian_date_col] if dian_date_col else None
            
            # Convertir valor DIAN a numérico de forma segura
            numeric_dian_value = self._safe_to_numeric(dian_value)
            if numeric_dian_value is None:
                continue
            
            # Buscar valores similares en contable
            rounded_dian_value = round(numeric_dian_value, 2)
            
            # Buscar en un rango de valores (±10%)
            min_value = rounded_dian_value * 0.9
            max_value = rounded_dian_value * 1.1
            
            candidates = []
            for value, indices in contable_by_value.items():
                if min_value <= value <= max_value:
                    candidates.extend(indices)
            
            if not candidates:
                continue
            
            # Verificar candidatos
            for contable_idx in candidates:
                if contable_idx in matched_contable_indices:
                    continue
                
                contable_row = contable_df.loc[contable_idx]
                contable_value = contable_row[contable_value_col]
                contable_date = contable_row[contable_date_col] if contable_date_col else None
                
                if pd.isna(contable_value):
                    continue
                
                # Verificar coincidencia de valor
                value_match = self._check_value_match(dian_value, contable_value, tolerance)
                
                if value_match:
                    # Verificar coincidencia de fecha
                    date_match = self._check_date_match(dian_date, contable_date)
                    
                    if date_match:
                        match_score = 0.8
                        match_reason = f'Valor y fecha coinciden: {dian_value}'
                    else:
                        match_score = 0.6
                        match_reason = f'Solo valor coincide: {dian_value}'
                    
                    match = {
                        'dian_idx': dian_idx,
                        'contable_idx': contable_idx,
                        'match_type': 'secondary_value_date',
                        'match_score': match_score,
                        'match_reason': match_reason
                    }
                    matches.append(match)
                    
                    # Marcar como emparejados
                    matched_dian_indices.add(dian_idx)
                    matched_contable_indices.add(contable_idx)
                    break
        
        self.logger.info(f"Encontradas {len(matches)} coincidencias secundarias")
        return matches
    
    def _find_similarity_matches(self, dian_df: pd.DataFrame, contable_df: pd.DataFrame, 
                               existing_matches: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Encontrar coincidencias por similitud de texto
        
        Args:
            dian_df: DataFrame DIAN
            contable_df: DataFrame contable
            existing_matches: Lista de coincidencias existentes
            
        Returns:
            Lista de coincidencias por similitud
        """
        matches = []
        
        # Obtener índices ya emparejados
        matched_dian_indices = {match['dian_idx'] for match in existing_matches}
        matched_contable_indices = {match['contable_idx'] for match in existing_matches}
        
        # Encontrar columnas de descripción
        dian_desc_col = self._find_description_column(dian_df, 'DIAN')
        contable_desc_col = self._find_description_column(contable_df, 'contable')
        
        if dian_desc_col and contable_desc_col:
            similarity_threshold = 0.7
            
            for dian_idx in dian_df.index:
                if dian_idx in matched_dian_indices:
                    continue
                
                dian_desc = str(dian_df.loc[dian_idx, dian_desc_col]).strip()
                
                if not dian_desc or dian_desc == 'nan':
                    continue
                
                best_match = None
                best_score = 0.0
                
                for contable_idx in contable_df.index:
                    if contable_idx in matched_contable_indices:
                        continue
                    
                    contable_desc = str(contable_df.loc[contable_idx, contable_desc_col]).strip()
                    
                    if not contable_desc or contable_desc == 'nan':
                        continue
                    
                    # Calcular similitud
                    similarity = SequenceMatcher(None, dian_desc.lower(), contable_desc.lower()).ratio()
                    
                    if similarity > best_score and similarity >= similarity_threshold:
                        best_score = similarity
                        best_match = contable_idx
                
                if best_match is not None:
                    match = {
                        'dian_idx': dian_idx,
                        'contable_idx': best_match,
                        'match_type': 'similarity',
                        'match_score': best_score,
                        'match_reason': f'Similitud de texto: {best_score:.2f}'
                    }
                    matches.append(match)
                    
                    # Marcar como emparejados
                    matched_dian_indices.add(dian_idx)
                    matched_contable_indices.add(best_match)
        
        self.logger.info(f"Encontradas {len(matches)} coincidencias por similitud")
        return matches
    
    def _find_value_column(self, df: pd.DataFrame, source: str) -> str:
        """Encontrar columna de valor/monto"""
        value_keywords = ['valor', 'monto', 'importe', 'total', 'debito', 'credito']
        
        for col in df.columns:
            if any(keyword in col.lower() for keyword in value_keywords):
                return col
        
        return None
    
    def _find_date_column(self, df: pd.DataFrame, source: str) -> str:
        """Encontrar columna de fecha"""
        date_keywords = ['fecha', 'date', 'dia', 'mes', 'año']
        
        for col in df.columns:
            if any(keyword in col.lower() for keyword in date_keywords):
                return col
        
        return None
    
    def _find_description_column(self, df: pd.DataFrame, source: str) -> str:
        """Encontrar columna de descripción"""
        desc_keywords = ['descripcion', 'concepto', 'detalle', 'observacion']
        
        for col in df.columns:
            if any(keyword in col.lower() for keyword in desc_keywords):
                return col
        
        return None
    
    def _safe_to_numeric(self, value) -> float:
        """Convertir valor a numérico de forma segura"""
        if pd.isna(value):
            return None
        
        try:
            # Si ya es numérico, retornarlo
            if isinstance(value, (int, float)):
                return float(value)
            
            # Si es string, intentar convertir
            if isinstance(value, str):
                # Limpiar el string (eliminar espacios, comas, etc.)
                cleaned = str(value).strip().replace(',', '').replace('$', '')
                
                # Si el string limpio está vacío o contiene solo letras, retornar None
                if not cleaned or cleaned.isalpha():
                    return None
                
                return float(cleaned)
            
            return float(value)
        except (ValueError, TypeError):
            return None
    
    def _check_value_match(self, value1: float, value2: float, tolerance: float) -> bool:
        """Verificar si dos valores coinciden dentro de la tolerancia"""
        if pd.isna(value1) or pd.isna(value2):
            return False
        
        try:
            val1 = float(value1)
            val2 = float(value2)
            
            if val1 == 0 and val2 == 0:
                return True
            
            if val1 == 0 or val2 == 0:
                return False
            
            difference = abs(val1 - val2) / max(abs(val1), abs(val2))
            return difference <= tolerance
            
        except (ValueError, TypeError):
            return False
    
    def _check_date_match(self, date1: str, date2: str) -> bool:
        """Verificar si dos fechas coinciden"""
        if pd.isna(date1) or pd.isna(date2):
            return False
        
        try:
            # Intentar convertir a datetime con formato específico
            def parse_date(date_str):
                try:
                    return pd.to_datetime(date_str, format='%d-%m-%Y', errors='coerce')
                except:
                    try:
                        return pd.to_datetime(date_str, format='%d-%m-%Y %H:%M:%S', errors='coerce')
                    except:
                        return pd.to_datetime(date_str, dayfirst=True, errors='coerce')
            
            dt1 = parse_date(date1)
            dt2 = parse_date(date2)
            
            if pd.isna(dt1) or pd.isna(dt2):
                return False
            
            # Tolerancia de 3 días
            date_diff = abs((dt1 - dt2).days)
            return date_diff <= 3
            
        except Exception:
            return False
    
    def _create_matches_dataframe(self, dian_df: pd.DataFrame, contable_df: pd.DataFrame, 
                                matches: List[Dict[str, Any]]) -> pd.DataFrame:
        """Crear DataFrame con las coincidencias"""
        if not matches:
            return pd.DataFrame()
        
        match_records = []
        
        for match in matches:
            dian_idx = match['dian_idx']
            contable_idx = match['contable_idx']
            
            dian_record = dian_df.loc[dian_idx].to_dict()
            contable_record = contable_df.loc[contable_idx].to_dict()
            
            # Combinar registros
            combined_record = {
                'source': 'match',
                'dian_idx': dian_idx,
                'contable_idx': contable_idx,
                'match_type': match['match_type'],
                'match_score': match['match_score'],
                'match_reason': match['match_reason']
            }
            
            # Agregar columnas DIAN
            for col, value in dian_record.items():
                combined_record[f'dian_{col}'] = value
            
            # Agregar columnas contables
            for col, value in contable_record.items():
                combined_record[f'contable_{col}'] = value
            
            match_records.append(combined_record)
        
        return pd.DataFrame(match_records)
    
    def _create_non_matches_dataframe(self, dian_df: pd.DataFrame, contable_df: pd.DataFrame, 
                                    matches: List[Dict[str, Any]]) -> pd.DataFrame:
        """Crear DataFrame con las no coincidencias"""
        non_matches = []
        
        # Obtener índices emparejados
        matched_dian_indices = {match['dian_idx'] for match in matches}
        matched_contable_indices = {match['contable_idx'] for match in matches}
        
        # Agregar registros DIAN no emparejados
        for idx in dian_df.index:
            if idx not in matched_dian_indices:
                record = dian_df.loc[idx].to_dict()
                record['source'] = 'DIAN'
                record['unmatched_idx'] = idx
                non_matches.append(record)
        
        # Agregar registros contables no emparejados
        for idx in contable_df.index:
            if idx not in matched_contable_indices:
                record = contable_df.loc[idx].to_dict()
                record['source'] = 'contable'
                record['unmatched_idx'] = idx
                non_matches.append(record)
        
        return pd.DataFrame(non_matches)
    
    def _analyze_discrepancies(self, matches: pd.DataFrame) -> List[Dict[str, Any]]:
        """Analizar discrepancias en las coincidencias"""
        discrepancies = []
        
        if matches.empty:
            return discrepancies
        
        # Verificar discrepancias en valores
        value_cols = [col for col in matches.columns if 'valor' in col.lower() or 'monto' in col.lower()]
        
        for _, match in matches.iterrows():
            dian_value_cols = [col for col in value_cols if col.startswith('dian_')]
            contable_value_cols = [col for col in value_cols if col.startswith('contable_')]
            
            for dian_col in dian_value_cols:
                for contable_col in contable_value_cols:
                    dian_value = match[dian_col]
                    contable_value = match[contable_col]
                    
                    if not pd.isna(dian_value) and not pd.isna(contable_value):
                        try:
                            numeric_dian = self._safe_to_numeric(dian_value)
                            numeric_contable = self._safe_to_numeric(contable_value)
                            if numeric_dian is not None and numeric_contable is not None:
                                diff = abs(numeric_dian - numeric_contable)
                            else:
                                continue
                            if diff > 0:
                                discrepancy = {
                                    'type': 'value_discrepancy',
                                    'dian_idx': match.get('dian_idx'),
                                    'contable_idx': match.get('contable_idx'),
                                    'dian_value': dian_value,
                                    'contable_value': contable_value,
                                    'difference': diff,
                                    'match_score': match.get('match_score', 0)
                                }
                                discrepancies.append(discrepancy)
                        except (ValueError, TypeError):
                            continue
        
        return discrepancies
    
    def get_file_info(self) -> Dict[str, Any]:
        """
        Obtener información de los archivos cargados
        
        Returns:
            Diccionario con información de los archivos
        """
        info = {
            'dian_loaded': self.dian_data is not None,
            'contable_loaded': self.contable_data is not None,
            'dian_rows': len(self.dian_data) if self.dian_data is not None else 0,
            'contable_rows': len(self.contable_data) if self.contable_data is not None else 0,
            'dian_columns': len(self.dian_data.columns) if self.dian_data is not None else 0,
            'contable_columns': len(self.contable_data.columns) if self.contable_data is not None else 0,
            'dian_file_path': str(self.dian_file_path) if self.dian_file_path else None,
            'contable_file_path': str(self.contable_file_path) if self.contable_file_path else None
        }
        
        return info
    
    def get_sheet_names(self, file_path: str | Path) -> List[str]:
        """
        Obtener nombres de todas las hojas del archivo Excel
        
        Args:
            file_path: Ruta del archivo Excel
            
        Returns:
            Lista con nombres de las hojas
        """
        file_path = Path(file_path)
        
        try:
            if not file_path.exists():
                raise FileNotFoundError(f"El archivo no existe: {file_path}")
            
            workbook = openpyxl.load_workbook(file_path)
            sheet_names = workbook.sheetnames
            
            self.logger.info(f"Hojas encontradas en {file_path}: {sheet_names}")
            
            return sheet_names
            
        except Exception as e:
            self.logger.error(f"Error al obtener nombres de hojas: {e}")
            raise Exception(f"Error al obtener nombres de hojas: {e}")
    
    def reset(self):
        """Limpiar todos los datos cargados"""
        self.dian_data = None
        self.contable_data = None
        self.dian_file_path = None
        self.contable_file_path = None
        self.logger.info("Datos del procesador limpiados")

    def create_coincidencias_dataframe(self, matches: pd.DataFrame) -> pd.DataFrame:
        """
        Crear DataFrame de coincidencias con estructura específica para Excel
        
        Args:
            matches: DataFrame con los registros que coinciden entre DIAN y contable
            
        Returns:
            DataFrame estructurado para la hoja "Coincidencias"
        """
        try:
            self.logger.info("Creando DataFrame de coincidencias")
            
            if matches.empty:
                self.logger.warning("No hay coincidencias para procesar")
                return pd.DataFrame()
            
            # Crear DataFrame de coincidencias con estructura específica
            coincidencias = pd.DataFrame()
            
            # Extraer columnas de DIAN (usando prefijo dian_)
            dian_columns = [col for col in matches.columns if col.startswith('dian_')]
            contable_columns = [col for col in matches.columns if col.startswith('contable_')]
            
            # Columnas de DIAN - mapear dinámicamente
            coincidencias['FOLIO DIAN'] = matches.get('dian_Folio', matches.get('dian_folio', ''))
            coincidencias['FECHA DIAN'] = matches.get('dian_Fecha Emisión', matches.get('dian_fecha', ''))
            coincidencias['VALOR DIAN'] = pd.to_numeric(matches.get('dian_Total', matches.get('dian_valor', 0.0)), errors='coerce').fillna(0.0)
            coincidencias['DESCRIPCIÓN DIAN'] = matches.get('dian_Descripción', matches.get('dian_descripcion', ''))
            coincidencias['TIPO DOCUMENTO DIAN'] = matches.get('dian_Tipo de documento', matches.get('dian_tipo_documento', ''))
            
            # Columnas de Contable - mapear dinámicamente
            coincidencias['NÚMERO DOCUMENTO CRUCE'] = matches.get('contable_valor_2', matches.get('contable_numero_documento', ''))
            coincidencias['FECHA CONTABLE'] = matches.get('contable_fecha', '')
            coincidencias['VALOR CONTABLE'] = pd.to_numeric(matches.get('contable_valor', 0.0), errors='coerce').fillna(0.0)
            coincidencias['DESCRIPCIÓN CONTABLE'] = matches.get('contable_descripcion', '')
            coincidencias['CUENTA CONTABLE'] = matches.get('contable_cuenta', '')
            
            # Columnas de validación y diferencias (ya convertidas a float arriba)
            coincidencias['DIFERENCIA VALOR'] = (
                coincidencias['VALOR DIAN'] - coincidencias['VALOR CONTABLE']
            ).round(2)
            
            # Calcular diferencia de fechas de forma segura
            fecha_dian = pd.to_datetime(coincidencias['FECHA DIAN'], errors='coerce')
            fecha_contable = pd.to_datetime(coincidencias['FECHA CONTABLE'], errors='coerce')
            coincidencias['DIFERENCIA FECHA'] = (fecha_dian - fecha_contable).dt.days.fillna(0).astype(int)
            
            # Columna de validación
            coincidencias['ESTADO VALIDACIÓN'] = coincidencias.apply(
                lambda row: self._evaluate_match_quality(row), axis=1
            )
            
            # Columna de tipo de coincidencia
            coincidencias['TIPO COINCIDENCIA'] = matches.get('match_type', 'Exacta')
            
            # Columna de confianza
            coincidencias['NIVEL CONFIANZA'] = matches.get('confidence', 1.0)
            
            # Ordenar por folio DIAN
            coincidencias = coincidencias.sort_values('FOLIO DIAN').reset_index(drop=True)
            
            self.logger.info(f"DataFrame de coincidencias creado: {len(coincidencias)} registros")
            
            return coincidencias
            
        except Exception as e:
            self.logger.error(f"Error al crear DataFrame de coincidencias: {e}")
            raise Exception(f"Error al crear DataFrame de coincidencias: {e}")

    def create_no_coincidencias_dataframe(self, non_matches: pd.DataFrame) -> pd.DataFrame:
        """
        Crear DataFrame de no coincidencias con estructura específica para Excel
        
        Args:
            non_matches: DataFrame con los registros que no coinciden
            
        Returns:
            DataFrame estructurado para la hoja "No coincidencias"
        """
        try:
            self.logger.info("Creando DataFrame de no coincidencias")
            
            if non_matches.empty:
                self.logger.warning("No hay no coincidencias para procesar")
                # Crear DataFrame vacío con estructura correcta
                return pd.DataFrame(columns=[
                    'FOLIO DIAN', 'FECHA DIAN', 'VALOR DIAN', 'DESCRIPCIÓN DIAN', 'TIPO DOCUMENTO DIAN',
                    'NÚMERO DOCUMENTO CRUCE', 'FECHA CONTABLE', 'VALOR CONTABLE', 'DESCRIPCIÓN CONTABLE',
                    'CUENTA CONTABLE', 'MOTIVO NO COINCIDENCIA', 'ORIGEN'
                ])
            
            # Crear DataFrame de no coincidencias con estructura específica
            no_coincidencias = pd.DataFrame()
            
            # Identificar registros DIAN sin contraparte
            dian_only = non_matches[non_matches['source'] == 'DIAN'].copy()
            contable_only = non_matches[non_matches['source'] == 'contable'].copy()
            
            # Procesar registros DIAN sin contraparte
            if not dian_only.empty:
                # Usar nombres de columnas reales del DataFrame DIAN
                dian_records = []
                for idx, row in dian_only.iterrows():
                    record = {
                        'FOLIO DIAN': row.get('Folio', ''),
                        'FECHA DIAN': row.get('Fecha Emisión', ''),
                        'VALOR DIAN': row.get('Valor Total', 0.0),
                        'DESCRIPCIÓN DIAN': row.get('Descripción', ''),
                        'TIPO DOCUMENTO DIAN': row.get('Tipo de documento', ''),
                        'NÚMERO DOCUMENTO CRUCE': '',
                        'FECHA CONTABLE': '',
                        'VALOR CONTABLE': 0.0,
                        'DESCRIPCIÓN CONTABLE': '',
                        'CUENTA CONTABLE': '',
                        'MOTIVO NO COINCIDENCIA': 'Registro DIAN sin contraparte contable',
                        'ORIGEN': 'DIAN'
                    }
                    dian_records.append(record)
                
                if dian_records:
                    dian_df = pd.DataFrame(dian_records)
                    no_coincidencias = pd.concat([no_coincidencias, dian_df], ignore_index=True)
            
            # Procesar registros contables sin contraparte
            if not contable_only.empty:
                # Usar nombres de columnas reales del DataFrame contable
                contable_records = []
                for idx, row in contable_only.iterrows():
                    record = {
                        'FOLIO DIAN': '',
                        'FECHA DIAN': '',
                        'VALOR DIAN': 0.0,
                        'DESCRIPCIÓN DIAN': '',
                        'TIPO DOCUMENTO DIAN': '',
                        'NÚMERO DOCUMENTO CRUCE': row.get('numero_documento', ''),
                        'FECHA CONTABLE': row.get('fecha', ''),
                        'VALOR CONTABLE': row.get('valor', 0.0),
                        'DESCRIPCIÓN CONTABLE': row.get('descripcion', row.get('detalle', '')),
                        'CUENTA CONTABLE': row.get('cuenta', ''),
                        'MOTIVO NO COINCIDENCIA': 'Registro contable sin contraparte DIAN',
                        'ORIGEN': 'CONTABLE'
                    }
                    contable_records.append(record)
                
                if contable_records:
                    contable_df = pd.DataFrame(contable_records)
                    no_coincidencias = pd.concat([no_coincidencias, contable_df], ignore_index=True)
            
            # Si no hay registros, crear DataFrame vacío con estructura correcta
            if no_coincidencias.empty:
                no_coincidencias = pd.DataFrame(columns=[
                    'FOLIO DIAN', 'FECHA DIAN', 'VALOR DIAN', 'DESCRIPCIÓN DIAN', 'TIPO DOCUMENTO DIAN',
                    'NÚMERO DOCUMENTO CRUCE', 'FECHA CONTABLE', 'VALOR CONTABLE', 'DESCRIPCIÓN CONTABLE',
                    'CUENTA CONTABLE', 'MOTIVO NO COINCIDENCIA', 'ORIGEN'
                ])
            else:
                # Agregar análisis de motivos más específicos
                no_coincidencias = self._add_detailed_non_match_reasons(no_coincidencias)
                
                # Ordenar por origen y luego por valor
                no_coincidencias = no_coincidencias.sort_values(
                    ['ORIGEN', 'VALOR DIAN', 'VALOR CONTABLE'], 
                    ascending=[True, False, False]
                ).reset_index(drop=True)
            
            self.logger.info(f"DataFrame de no coincidencias creado: {len(no_coincidencias)} registros")
            
            return no_coincidencias
            
        except Exception as e:
            self.logger.error(f"Error al crear DataFrame de no coincidencias: {e}")
            raise Exception(f"Error al crear DataFrame de no coincidencias: {e}")

    def calculate_statistics(self, coincidencias: pd.DataFrame, no_coincidencias: pd.DataFrame) -> Dict[str, Any]:
        """
        Calcular estadísticas completas del proceso de causación
        
        Args:
            coincidencias: DataFrame de registros que coinciden
            no_coincidencias: DataFrame de registros que no coinciden
            
        Returns:
            Diccionario con estadísticas detalladas
        """
        try:
            self.logger.info("Calculando estadísticas del proceso de causación")
            
            stats = {}
            
            # Totales por categoría
            stats['total_coincidencias'] = len(coincidencias)
            stats['total_no_coincidencias'] = len(no_coincidencias)
            stats['total_registros'] = stats['total_coincidencias'] + stats['total_no_coincidencias']
            
            # Porcentajes de matching
            if stats['total_registros'] > 0:
                stats['porcentaje_coincidencias'] = (stats['total_coincidencias'] / stats['total_registros']) * 100
                stats['porcentaje_no_coincidencias'] = (stats['total_no_coincidencias'] / stats['total_registros']) * 100
            else:
                stats['porcentaje_coincidencias'] = 0.0
                stats['porcentaje_no_coincidencias'] = 0.0
            
            # Análisis de valores
            if not coincidencias.empty:
                # Convertir columnas a numérico de forma segura
                valor_dian_numeric = pd.to_numeric(coincidencias['VALOR DIAN'], errors='coerce').fillna(0)
                valor_contable_numeric = pd.to_numeric(coincidencias['VALOR CONTABLE'], errors='coerce').fillna(0)
                
                stats['valor_total_coincidencias'] = valor_dian_numeric.sum()
                stats['valor_total_contable_coincidencias'] = valor_contable_numeric.sum()
                stats['diferencia_total_valores'] = stats['valor_total_coincidencias'] - stats['valor_total_contable_coincidencias']
                stats['porcentaje_diferencia_valores'] = (
                    abs(stats['diferencia_total_valores']) / stats['valor_total_coincidencias'] * 100
                    if stats['valor_total_coincidencias'] > 0 else 0.0
                )
                
                # Análisis de diferencias por tipo de coincidencia
                stats['coincidencias_exactas'] = len(coincidencias[coincidencias['TIPO COINCIDENCIA'] == 'Exacta'])
                stats['coincidencias_secundarias'] = len(coincidencias[coincidencias['TIPO COINCIDENCIA'] == 'Secundaria'])
                stats['coincidencias_similitud'] = len(coincidencias[coincidencias['TIPO COINCIDENCIA'] == 'Similitud'])
            else:
                # Valores por defecto cuando no hay coincidencias
                stats['valor_total_coincidencias'] = 0.0
                stats['valor_total_contable_coincidencias'] = 0.0
                stats['diferencia_total_valores'] = 0.0
                stats['porcentaje_diferencia_valores'] = 0.0
                stats['coincidencias_exactas'] = 0
                stats['coincidencias_secundarias'] = 0
                stats['coincidencias_similitud'] = 0
            
            # Análisis de no coincidencias
            if not no_coincidencias.empty:
                dian_only = no_coincidencias[no_coincidencias['ORIGEN'] == 'DIAN']
                contable_only = no_coincidencias[no_coincidencias['ORIGEN'] == 'CONTABLE']
                
                stats['registros_dian_sin_contraparte'] = len(dian_only)
                stats['registros_contable_sin_contraparte'] = len(contable_only)
                
                # Convertir a numérico de forma segura para no coincidencias
                if len(dian_only) > 0 and 'VALOR DIAN' in dian_only.columns:
                    valor_dian_no_match = pd.to_numeric(dian_only['VALOR DIAN'], errors='coerce').fillna(0)
                    stats['valor_dian_sin_contraparte'] = valor_dian_no_match.sum()
                else:
                    stats['valor_dian_sin_contraparte'] = 0.0
                    
                if len(contable_only) > 0 and 'VALOR CONTABLE' in contable_only.columns:
                    valor_contable_no_match = pd.to_numeric(contable_only['VALOR CONTABLE'], errors='coerce').fillna(0)
                    stats['valor_contable_sin_contraparte'] = valor_contable_no_match.sum()
                else:
                    stats['valor_contable_sin_contraparte'] = 0.0
            else:
                # Valores por defecto cuando no hay no coincidencias
                stats['registros_dian_sin_contraparte'] = 0
                stats['registros_contable_sin_contraparte'] = 0
                stats['valor_dian_sin_contraparte'] = 0.0
                stats['valor_contable_sin_contraparte'] = 0.0
            
            # Métricas de calidad
            if not coincidencias.empty:
                stats['coincidencias_con_diferencia_valor'] = len(
                    coincidencias[abs(coincidencias['DIFERENCIA VALOR']) > 0.01]
                )
                stats['coincidencias_con_diferencia_fecha'] = len(
                    coincidencias[abs(coincidencias['DIFERENCIA FECHA']) > 0]
                )
                stats['coincidencias_perfectas'] = len(
                    coincidencias[
                        (abs(coincidencias['DIFERENCIA VALOR']) <= 0.01) & 
                        (abs(coincidencias['DIFERENCIA FECHA']) == 0)
                    ]
                )
            else:
                # Valores por defecto cuando no hay coincidencias
                stats['coincidencias_con_diferencia_valor'] = 0
                stats['coincidencias_con_diferencia_fecha'] = 0
                stats['coincidencias_perfectas'] = 0
            
            # Resumen ejecutivo
            stats['resumen_ejecutivo'] = {
                'total_procesado': stats['total_registros'],
                'coincidencias_encontradas': stats['total_coincidencias'],
                'tasa_exito': f"{stats['porcentaje_coincidencias']:.1f}%",
                'valor_total_coincidencias': f"${stats['valor_total_coincidencias']:,.2f}",
                'diferencia_total': f"${stats['diferencia_total_valores']:,.2f}",
                'calidad_general': self._calculate_overall_quality(stats)
            }
            
            self.logger.info("Estadísticas calculadas exitosamente")
            
            return stats
            
        except Exception as e:
            self.logger.error(f"Error al calcular estadísticas: {e}")
            raise Exception(f"Error al calcular estadísticas: {e}")

    def _evaluate_match_quality(self, row: pd.Series) -> str:
        """
        Evaluar la calidad de una coincidencia basada en diferencias de valor y fecha
        
        Args:
            row: Fila del DataFrame de coincidencias
            
        Returns:
            Estado de validación: 'Perfecta', 'Buena', 'Regular', 'Revisar'
        """
        try:
            diff_valor = abs(row['DIFERENCIA VALOR'])
            diff_fecha = abs(row['DIFERENCIA FECHA'])
            
            if diff_valor <= 0.01 and diff_fecha == 0:
                return 'Perfecta'
            elif diff_valor <= 1.0 and diff_fecha <= 1:
                return 'Buena'
            elif diff_valor <= 10.0 and diff_fecha <= 7:
                return 'Regular'
            else:
                return 'Revisar'
                
        except Exception:
            return 'Revisar'

    def _add_detailed_non_match_reasons(self, no_coincidencias: pd.DataFrame) -> pd.DataFrame:
        """
        Agregar motivos detallados de no coincidencia basados en análisis de datos
        
        Args:
            no_coincidencias: DataFrame de no coincidencias
            
        Returns:
            DataFrame con motivos detallados agregados
        """
        try:
            # Crear copia para no modificar el original
            df = no_coincidencias.copy()
            
            # Analizar motivos específicos
            for idx, row in df.iterrows():
                if row['ORIGEN'] == 'DIAN':
                    # Analizar posibles motivos para registros DIAN sin contraparte
                    motivo = self._analyze_dian_non_match_reason(row)
                    df.at[idx, 'MOTIVO NO COINCIDENCIA'] = motivo
                elif row['ORIGEN'] == 'CONTABLE':
                    # Analizar posibles motivos para registros contables sin contraparte
                    motivo = self._analyze_contable_non_match_reason(row)
                    df.at[idx, 'MOTIVO NO COINCIDENCIA'] = motivo
            
            return df
            
        except Exception as e:
            self.logger.warning(f"Error al agregar motivos detallados: {e}")
            return no_coincidencias

    def _analyze_dian_non_match_reason(self, row: pd.Series) -> str:
        """
        Analizar motivo específico para registro DIAN sin contraparte
        
        Args:
            row: Fila del DataFrame de no coincidencias
            
        Returns:
            Motivo específico de no coincidencia
        """
        try:
            folio = str(row['FOLIO DIAN']).strip()
            valor = self._safe_to_numeric(row['VALOR DIAN'])
            fecha = str(row['FECHA DIAN'])
            
            # Verificar si el folio tiene formato válido
            if not folio or folio == 'nan':
                return 'Folio DIAN vacío o inválido'
            
            # Verificar si el valor es muy alto o muy bajo (posible error)
            if valor > 1000000000:  # Más de 1 billón
                return 'Valor DIAN extremadamente alto (posible error)'
            elif valor < 0:
                return 'Valor DIAN negativo'
            
            # Verificar formato de fecha
            if not fecha or fecha == 'nan':
                return 'Fecha DIAN vacía o inválida'
            
            # Motivo genérico si no se encuentra causa específica
            return 'Registro DIAN sin contraparte contable'
            
        except Exception:
            return 'Registro DIAN sin contraparte contable'

    def _analyze_contable_non_match_reason(self, row: pd.Series) -> str:
        """
        Analizar motivo específico para registro contable sin contraparte
        
        Args:
            row: Fila del DataFrame de no coincidencias
            
        Returns:
            Motivo específico de no coincidencia
        """
        try:
            numero_doc = str(row['NÚMERO DOCUMENTO CRUCE']).strip()
            valor = self._safe_to_numeric(row['VALOR CONTABLE'])
            fecha = str(row['FECHA CONTABLE'])
            
            # Verificar si el número de documento tiene formato válido
            if not numero_doc or numero_doc == 'nan':
                return 'Número de documento contable vacío o inválido'
            
            # Verificar si el valor es muy alto o muy bajo (posible error)
            if valor > 1000000000:  # Más de 1 billón
                return 'Valor contable extremadamente alto (posible error)'
            elif valor < 0:
                return 'Valor contable negativo'
            
            # Verificar formato de fecha
            if not fecha or fecha == 'nan':
                return 'Fecha contable vacía o inválida'
            
            # Motivo genérico si no se encuentra causa específica
            return 'Registro contable sin contraparte DIAN'
            
        except Exception:
            return 'Registro contable sin contraparte DIAN'

    def _calculate_overall_quality(self, stats: Dict[str, Any]) -> str:
        """
        Calcular calidad general del proceso basado en estadísticas
        
        Args:
            stats: Diccionario con estadísticas del proceso
            
        Returns:
            Calificación de calidad: 'Excelente', 'Buena', 'Regular', 'Mala'
        """
        try:
            # Calcular score basado en múltiples factores
            score = 0
            
            # Factor 1: Porcentaje de coincidencias (40% del score)
            if stats['porcentaje_coincidencias'] >= 90:
                score += 40
            elif stats['porcentaje_coincidencias'] >= 80:
                score += 35
            elif stats['porcentaje_coincidencias'] >= 70:
                score += 30
            elif stats['porcentaje_coincidencias'] >= 60:
                score += 25
            else:
                score += 20
            
            # Factor 2: Calidad de coincidencias (30% del score)
            if stats['total_coincidencias'] > 0:
                porcentaje_perfectas = (stats['coincidencias_perfectas'] / stats['total_coincidencias']) * 100
                if porcentaje_perfectas >= 80:
                    score += 30
                elif porcentaje_perfectas >= 60:
                    score += 25
                elif porcentaje_perfectas >= 40:
                    score += 20
                else:
                    score += 15
            
            # Factor 3: Diferencia de valores (20% del score)
            if stats['porcentaje_diferencia_valores'] <= 1:
                score += 20
            elif stats['porcentaje_diferencia_valores'] <= 5:
                score += 15
            elif stats['porcentaje_diferencia_valores'] <= 10:
                score += 10
            else:
                score += 5
            
            # Factor 4: Balance entre DIAN y contable (10% del score)
            if stats['total_no_coincidencias'] > 0:
                balance = abs(stats['registros_dian_sin_contraparte'] - stats['registros_contable_sin_contraparte'])
                balance_ratio = balance / stats['total_no_coincidencias']
                if balance_ratio <= 0.2:
                    score += 10
                elif balance_ratio <= 0.4:
                    score += 8
                elif balance_ratio <= 0.6:
                    score += 6
                else:
                    score += 4
            
            # Determinar calificación final
            if score >= 85:
                return 'Excelente'
            elif score >= 70:
                return 'Buena'
            elif score >= 50:
                return 'Regular'
            else:
                return 'Mala'
                
        except Exception:
            return 'Regular'

    def create_excel_file(self, coincidencias_df: pd.DataFrame, no_coincidencias_df: pd.DataFrame, 
                         output_path: str | Path, stats: Dict[str, Any] = None) -> str:
        """
        Crear archivo Excel profesional con formato de tabla
        
        Args:
            coincidencias_df: DataFrame de coincidencias
            no_coincidencias_df: DataFrame de no coincidencias
            output_path: Ruta donde guardar el archivo Excel
            stats: Estadísticas del proceso (opcional)
            
        Returns:
            Ruta del archivo Excel creado
            
        Raises:
            Exception: Si hay error al crear el archivo
        """
        try:
            output_path = Path(output_path)
            
            # Generar nombre de archivo dinámico si no se especifica
            if output_path.is_dir() or not output_path.suffix:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"Reporte_Causacion_{timestamp}.xlsx"
                output_path = output_path / filename
            
            self.logger.info(f"Creando archivo Excel: {output_path}")
            
            # Crear el archivo Excel con openpyxl (más estable)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Crear hojas básicas con openpyxl (más simple y estable)
                try:
                    # Crear hoja de coincidencias
                    if not coincidencias_df.empty:
                        coincidencias_df.to_excel(writer, sheet_name='Coincidencias', index=False)
                    
                    # Crear hoja de no coincidencias
                    if not no_coincidencias_df.empty:
                        no_coincidencias_df.to_excel(writer, sheet_name='No_coincidencias', index=False)
                    
                    # Crear hoja de resumen básica
                    if stats:
                        self._create_simple_summary_sheet(writer, stats)
                    
                    # Crear hoja de metadatos básica
                    self._create_simple_metadata_sheet(writer)
                    
                    self.logger.info("Excel creado con formato básico usando openpyxl")
                    
                except Exception as e:
                    self.logger.error(f"Error creando Excel con openpyxl: {e}")
                    # Si falla con openpyxl, intentar con xlsxwriter básico
                    raise e
            
            self.logger.info(f"Archivo Excel creado exitosamente: {output_path}")
            return str(output_path)
            
        except Exception as e:
            self.logger.error(f"Error al crear archivo Excel: {e}")
            # Intentar crear un archivo básico como último recurso
            try:
                self.logger.info("Intentando crear archivo Excel básico como fallback...")
                self._create_basic_excel_emergency(coincidencias_df, no_coincidencias_df, output_path)
                return str(output_path)
            except Exception as fallback_error:
                self.logger.error(f"Error en fallback: {fallback_error}")
                # No hacer raise, continuar con el proceso
                raise Exception(f"Error al crear archivo Excel: {e}")

    def _create_excel_formats(self, workbook) -> Dict[str, Any]:
        """
        Crear formatos para el archivo Excel
        
        Args:
            workbook: Objeto workbook de xlsxwriter
            
        Returns:
            Diccionario con los formatos creados
        """
        formats = {}
        
        try:
            # Formato de título principal mejorado
            formats['title'] = workbook.add_format({
                'bold': True,
                'font_size': 18,
                'font_color': '#ffffff',
                'font_name': 'Calibri',
                'align': 'center',
                'valign': 'vcenter',
                'border': 2,
                'border_color': '#1f4e79',
                'bg_color': '#1f4e79'
            })
        
            # Formato de subtítulo
            formats['subtitle'] = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'font_color': '#2e5984',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#e6f3ff'
            })
        
            # Formato de encabezados de tabla mejorado
            formats['header'] = workbook.add_format({
                'bold': True,
                'font_size': 11,
                'font_color': '#ffffff',
                'font_name': 'Calibri',
                'align': 'center',
                'valign': 'vcenter',
                'border': 2,
                'border_color': '#2e5984',
                'bg_color': '#2e5984',
                'text_wrap': True
            })
        
            # Formato de datos normales mejorado
            formats['data'] = workbook.add_format({
                'font_size': 10,
                'font_name': 'Calibri',
                'align': 'left',
                'valign': 'vcenter',
                'border': 1,
                'border_color': '#d0d0d0',
                'text_wrap': True,
                'shrink': True  # Ajustar texto automáticamente
            })
        
            # Formato de datos numéricos mejorado
            formats['number'] = workbook.add_format({
                'font_size': 10,
                'font_name': 'Calibri',
                'align': 'right',
                'valign': 'vcenter',
                'border': 1,
                'border_color': '#d0d0d0',
                'num_format': '$#,##0.00',  # Formato de moneda
                'shrink': True
            })
        
            # Formato de fechas mejorado
            formats['date'] = workbook.add_format({
                'font_size': 10,
                'font_name': 'Calibri',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'border_color': '#d0d0d0',
                'num_format': 'dd/mm/yyyy',
                'shrink': True
            })
        
            # Formato de estado de validación
            formats['status_perfecta'] = workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#c6efce',
                'font_color': '#006100'
            })
        
            formats['status_buena'] = workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#ffeb9c',
                'font_color': '#9c6500'
            })
        
            formats['status_regular'] = workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#ffc7ce',
                'font_color': '#9c0006'
            })
        
            formats['status_revisar'] = workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#ffeb9c',
                'font_color': '#9c6500'
            })
        
            # Formato de información
            formats['info'] = workbook.add_format({
                'font_size': 10,
                'align': 'left',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#f2f2f2'
            })
        
        except Exception as e:
            self.logger.warning(f"Error creando formatos avanzados: {e}. Usando formatos básicos.")
            # Formatos básicos como fallback
            formats = {
                'title': workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'}),
                'subtitle': workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center'}),
                'header': workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center'}),
                'data': workbook.add_format({'font_size': 10, 'align': 'left'}),
                'number': workbook.add_format({'font_size': 10, 'align': 'right', 'num_format': '$#,##0.00'}),
                'date': workbook.add_format({'font_size': 10, 'align': 'center', 'num_format': 'dd/mm/yyyy'}),
                'status_perfecta': workbook.add_format({'align': 'center', 'bg_color': '#c6efce'}),
                'status_buena': workbook.add_format({'align': 'center', 'bg_color': '#ffeb9c'}),
                'status_regular': workbook.add_format({'align': 'center', 'bg_color': '#ffc7ce'}),
                'status_revisar': workbook.add_format({'align': 'center', 'bg_color': '#ffcccc'}),
                'info': workbook.add_format({'font_size': 10, 'align': 'left'})
            }
        
        return formats

    def _create_basic_excel_fallback(self, writer, coincidencias_df: pd.DataFrame, no_coincidencias_df: pd.DataFrame, stats: Dict[str, Any]):
        """
        Crear Excel básico sin formatos avanzados como fallback
        
        Args:
            writer: ExcelWriter object
            coincidencias_df: DataFrame de coincidencias
            no_coincidencias_df: DataFrame de no coincidencias
            stats: Estadísticas del proceso
        """
        try:
            self.logger.info("Creando Excel con formato básico...")
            
            # Crear hojas básicas sin formatos avanzados
            if not coincidencias_df.empty:
                coincidencias_df.to_excel(writer, sheet_name='Coincidencias', index=False)
                
            if not no_coincidencias_df.empty:
                no_coincidencias_df.to_excel(writer, sheet_name='No coincidencias', index=False)
                
            # Crear hoja de resumen básica
            if stats:
                import pandas as pd
                summary_data = {
                    'Metric': ['Total Coincidencias', 'Total No Coincidencias', 'Tasa de Matching'],
                    'Value': [
                        len(coincidencias_df),
                        len(no_coincidencias_df),
                        f"{(len(coincidencias_df) / (len(coincidencias_df) + len(no_coincidencias_df)) * 100):.2f}%" if (len(coincidencias_df) + len(no_coincidencias_df)) > 0 else "0%"
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Resumen', index=False)
                
            self.logger.info("Excel básico creado exitosamente")
            
        except Exception as e:
            self.logger.error(f"Error en fallback básico: {e}")
            raise

    def _create_basic_excel_emergency(self, coincidencias_df: pd.DataFrame, no_coincidencias_df: pd.DataFrame, output_path: Path):
        """
        Crear Excel de emergencia muy básico
        
        Args:
            coincidencias_df: DataFrame de coincidencias
            no_coincidencias_df: DataFrame de no coincidencias
            output_path: Ruta de salida
        """
        try:
            self.logger.info("Creando Excel de emergencia...")
            
            # Cambiar a openpyxl que es más estable
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                if not coincidencias_df.empty:
                    coincidencias_df.to_excel(writer, sheet_name='Coincidencias', index=False)
                    
                if not no_coincidencias_df.empty:
                    no_coincidencias_df.to_excel(writer, sheet_name='No_coincidencias', index=False)
                    
            self.logger.info("Excel de emergencia creado exitosamente")
            
        except Exception as e:
            self.logger.error(f"Error en Excel de emergencia: {e}")
            raise

    def _create_simple_excel_fallback(self, writer, coincidencias_df: pd.DataFrame, no_coincidencias_df: pd.DataFrame, stats: Dict[str, Any]):
        """
        Crear Excel muy simple usando solo pandas como último recurso
        
        Args:
            writer: ExcelWriter object
            coincidencias_df: DataFrame de coincidencias
            no_coincidencias_df: DataFrame de no coincidencias
            stats: Estadísticas del proceso
        """
        try:
            self.logger.info("Creando Excel simple como último recurso...")
            
            # Solo escribir datos sin formatos
            if not coincidencias_df.empty:
                coincidencias_df.to_excel(writer, sheet_name='Coincidencias', index=False, startrow=0)
                
            if not no_coincidencias_df.empty:
                no_coincidencias_df.to_excel(writer, sheet_name='No_coincidencias', index=False, startrow=0)
                
            self.logger.info("Excel simple creado exitosamente")
            
        except Exception as e:
            self.logger.error(f"Error en Excel simple: {e}")
            # Este es el último recurso, si falla aquí no hay más opciones

    def _create_simple_summary_sheet(self, writer, stats: Dict[str, Any]):
        """
        Crear hoja de resumen simple
        
        Args:
            writer: ExcelWriter object
            stats: Estadísticas del proceso
        """
        try:
            # Crear datos de resumen básicos
            total_coincidencias = stats.get('total_coincidencias', 0)
            total_no_coincidencias = stats.get('total_no_coincidencias', 0)
            total_registros = total_coincidencias + total_no_coincidencias
            
            tasa_matching = (total_coincidencias / total_registros * 100) if total_registros > 0 else 0
            
            summary_data = {
                'Métrica': [
                    'Total de Coincidencias',
                    'Total de No Coincidencias',
                    'Total de Registros',
                    'Tasa de Matching (%)',
                    'Fecha de Proceso'
                ],
                'Valor': [
                    total_coincidencias,
                    total_no_coincidencias,
                    total_registros,
                    f"{tasa_matching:.2f}%",
                    datetime.now().strftime('%d-%m-%Y %H:%M:%S')
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Resumen', index=False)
            
            self.logger.info("Hoja de resumen simple creada exitosamente")
            
        except Exception as e:
            self.logger.error(f"Error creando resumen simple: {e}")

    def _create_simple_metadata_sheet(self, writer):
        """
        Crear hoja de metadatos simple
        
        Args:
            writer: ExcelWriter object
        """
        try:
            # Crear datos de metadatos básicos
            metadata_data = {
                'Propiedad': [
                    'Fecha de Generación',
                    'Procesador',
                    'Versión',
                    'Motor Excel',
                    'Sistema'
                ],
                'Valor': [
                    datetime.now().strftime('%d-%m-%Y %H:%M:%S'),
                    'CausacionProcessor',
                    '1.0',
                    'openpyxl',
                    'Windows'
                ]
            }
            
            metadata_df = pd.DataFrame(metadata_data)
            metadata_df.to_excel(writer, sheet_name='Metadatos', index=False)
            
            self.logger.info("Hoja de metadatos simple creada exitosamente")
            
        except Exception as e:
            self.logger.error(f"Error creando metadatos simples: {e}")

    def _create_coincidencias_sheet(self, writer, coincidencias_df: pd.DataFrame, formats: Dict[str, Any]):
        """
        Crear hoja de coincidencias con formato profesional
        
        Args:
            writer: ExcelWriter object
            coincidencias_df: DataFrame de coincidencias
            formats: Diccionario de formatos
        """
        worksheet = writer.book.add_worksheet('Coincidencias')
        
        # Aplicar formato básico
        self._apply_basic_formatting(worksheet, 'REPORTE DE COINCIDENCIAS', 
                                   'Registros que coinciden entre DIAN y Contable', formats)
        
        # Definir columnas y sus formatos
        columns = [
            ('FOLIO DIAN', 'data'),
            ('FECHA DIAN', 'date'),
            ('VALOR DIAN', 'number'),
            ('DESCRIPCIÓN DIAN', 'data'),
            ('TIPO DOCUMENTO DIAN', 'data'),
            ('NÚMERO DOCUMENTO CRUCE', 'data'),
            ('FECHA CONTABLE', 'date'),
            ('VALOR CONTABLE', 'number'),
            ('DESCRIPCIÓN CONTABLE', 'data'),
            ('CUENTA CONTABLE', 'data'),
            ('DIFERENCIA VALOR', 'number'),
            ('DIFERENCIA FECHA', 'number'),
            ('ESTADO VALIDACIÓN', 'data'),
            ('TIPO COINCIDENCIA', 'data'),
            ('NIVEL CONFIANZA', 'number')
        ]
        
        # Escribir encabezados
        start_row = 4
        for col_idx, (col_name, format_type) in enumerate(columns):
            worksheet.write(start_row, col_idx, col_name, formats['header'])
        
        # Escribir datos
        for row_idx, (_, row) in enumerate(coincidencias_df.iterrows(), start_row + 1):
            for col_idx, (col_name, format_type) in enumerate(columns):
                value = row[col_name]
                
                # Aplicar formato específico según el tipo de columna (temporalmente simplificado)
                try:
                    if format_type == 'date':
                        worksheet.write(row_idx, col_idx, value, formats['date'])
                    elif format_type == 'number':
                        worksheet.write(row_idx, col_idx, value, formats['number'])
                    elif col_name == 'ESTADO VALIDACIÓN':
                        # Aplicar formato de color según el estado
                        status_format = self._get_status_format(value, formats)
                        worksheet.write(row_idx, col_idx, value, status_format)
                    else:
                        worksheet.write(row_idx, col_idx, value, formats['data'])
                except Exception as format_error:
                    # Fallback sin formato si hay error
                    worksheet.write(row_idx, col_idx, value)
        
        # Ajustar ancho de columnas de forma inteligente
        for col_idx, (col_name, _) in enumerate(columns):
            if not coincidencias_df.empty:
                # Calcular ancho basado en contenido
                content_width = coincidencias_df[col_name].astype(str).str.len().max()
                header_width = len(col_name)
                
                # Ancho mínimo y máximo inteligente
                min_width = 12
                max_width = 50 if col_name in ['DESCRIPCIÓN DIAN', 'DESCRIPCIÓN CONTABLE'] else 25
                
                # Calcular ancho óptimo
                optimal_width = max(header_width + 3, content_width + 2, min_width)
                final_width = min(optimal_width, max_width)
                
                worksheet.set_column(col_idx, col_idx, final_width)
            else:
                # Ancho por defecto para columnas vacías
                worksheet.set_column(col_idx, col_idx, 15)
        
        # Agregar información de resumen
        self._add_sheet_summary(worksheet, coincidencias_df, start_row + len(coincidencias_df) + 2, formats)
        
        # Ajustar altura de filas para mejor visualización
        for row_idx in range(start_row, start_row + len(coincidencias_df) + 1):
            worksheet.set_row(row_idx, 20)  # Altura óptima para lectura
        
        # Aplicar formato avanzado
        data_range = f'A{start_row}:O{start_row + len(coincidencias_df)}'
        self.apply_conditional_formatting(writer.book, worksheet, data_range, 'coincidencias')
        self.add_filters_and_sorting(writer.book, worksheet, start_row, data_range)
        self.add_contador_tools(writer.book, worksheet, data_range, 'coincidencias')

    def _create_no_coincidencias_sheet(self, writer, no_coincidencias_df: pd.DataFrame, formats: Dict[str, Any]):
        """
        Crear hoja de no coincidencias con formato profesional
        
        Args:
            writer: ExcelWriter object
            no_coincidencias_df: DataFrame de no coincidencias
            formats: Diccionario de formatos
        """
        worksheet = writer.book.add_worksheet('No coincidencias')
        
        # Aplicar formato básico
        self._apply_basic_formatting(worksheet, 'REPORTE DE NO COINCIDENCIAS', 
                                   'Registros que no coinciden entre DIAN y Contable', formats)
        
        # Definir columnas y sus formatos
        columns = [
            ('FOLIO DIAN', 'data'),
            ('FECHA DIAN', 'date'),
            ('VALOR DIAN', 'number'),
            ('DESCRIPCIÓN DIAN', 'data'),
            ('TIPO DOCUMENTO DIAN', 'data'),
            ('NÚMERO DOCUMENTO CRUCE', 'data'),
            ('FECHA CONTABLE', 'date'),
            ('VALOR CONTABLE', 'number'),
            ('DESCRIPCIÓN CONTABLE', 'data'),
            ('CUENTA CONTABLE', 'data'),
            ('MOTIVO NO COINCIDENCIA', 'data'),
            ('ORIGEN', 'data')
        ]
        
        # Escribir encabezados
        start_row = 4
        for col_idx, (col_name, format_type) in enumerate(columns):
            worksheet.write(start_row, col_idx, col_name, formats['header'])
        
        # Escribir datos
        for row_idx, (_, row) in enumerate(no_coincidencias_df.iterrows(), start_row + 1):
            for col_idx, (col_name, format_type) in enumerate(columns):
                value = row[col_name]
                
                # Aplicar formato específico según el tipo de columna (temporalmente simplificado)
                try:
                    if format_type == 'date':
                        worksheet.write(row_idx, col_idx, value, formats['date'])
                    elif format_type == 'number':
                        worksheet.write(row_idx, col_idx, value, formats['number'])
                    else:
                        worksheet.write(row_idx, col_idx, value, formats['data'])
                except Exception as format_error:
                    # Fallback sin formato si hay error
                    worksheet.write(row_idx, col_idx, value)
        
        # Ajustar ancho de columnas de forma inteligente
        for col_idx, (col_name, _) in enumerate(columns):
            if not no_coincidencias_df.empty:
                # Calcular ancho basado en contenido
                content_width = no_coincidencias_df[col_name].astype(str).str.len().max()
                header_width = len(col_name)
                
                # Ancho mínimo y máximo inteligente
                min_width = 12
                max_width = 50 if col_name in ['DESCRIPCIÓN DIAN', 'DESCRIPCIÓN CONTABLE', 'MOTIVO NO COINCIDENCIA'] else 25
                
                # Calcular ancho óptimo
                optimal_width = max(header_width + 3, content_width + 2, min_width)
                final_width = min(optimal_width, max_width)
                
                worksheet.set_column(col_idx, col_idx, final_width)
            else:
                # Ancho por defecto para columnas vacías
                worksheet.set_column(col_idx, col_idx, 15)
        
        # Agregar información de resumen
        self._add_sheet_summary(worksheet, no_coincidencias_df, start_row + len(no_coincidencias_df) + 2, formats)
        
        # Ajustar altura de filas para mejor visualización
        for row_idx in range(start_row, start_row + len(no_coincidencias_df) + 1):
            worksheet.set_row(row_idx, 20)  # Altura óptima para lectura
        
        # Aplicar formato avanzado
        data_range = f'A{start_row}:L{start_row + len(no_coincidencias_df)}'
        self.apply_conditional_formatting(writer.book, worksheet, data_range, 'no_coincidencias')
        self.add_filters_and_sorting(writer.book, worksheet, start_row, data_range)
        self.add_contador_tools(writer.book, worksheet, data_range, 'no_coincidencias')

    def _create_summary_sheet(self, writer, stats: Dict[str, Any], formats: Dict[str, Any]):
        """
        Crear hoja de resumen con estadísticas
        
        Args:
            writer: ExcelWriter object
            stats: Diccionario con estadísticas
            formats: Diccionario de formatos
        """
        worksheet = writer.book.add_worksheet('Resumen')
        
        # Aplicar formato básico
        self._apply_basic_formatting(worksheet, 'RESUMEN EJECUTIVO', 
                                   'Estadísticas del proceso de causación', formats)
        
        # Escribir estadísticas principales
        start_row = 4
        summary_data = [
            ('Total de registros procesados', stats['total_registros']),
            ('Coincidencias encontradas', stats['total_coincidencias']),
            ('No coincidencias', stats['total_no_coincidencias']),
            ('Porcentaje de coincidencias', f"{stats['porcentaje_coincidencias']:.1f}%"),
            ('Porcentaje de no coincidencias', f"{stats['porcentaje_no_coincidencias']:.1f}%"),
            ('', ''),
            ('VALORES', ''),
            ('Valor total coincidencias DIAN', f"${stats['valor_total_coincidencias']:,.2f}"),
            ('Valor total coincidencias Contable', f"${stats['valor_total_contable_coincidencias']:,.2f}"),
            ('Diferencia total de valores', f"${stats['diferencia_total_valores']:,.2f}"),
            ('Porcentaje de diferencia', f"{stats['porcentaje_diferencia_valores']:.2f}%"),
            ('', ''),
            ('CALIDAD', ''),
            ('Coincidencias exactas', stats['coincidencias_exactas']),
            ('Coincidencias perfectas', stats['coincidencias_perfectas']),
            ('Calidad general del proceso', stats['resumen_ejecutivo']['calidad_general'])
        ]
        
        for row_idx, (label, value) in enumerate(summary_data, start_row):
            worksheet.write(row_idx, 0, label, formats['info'])
            worksheet.write(row_idx, 1, value, formats['data'])
        
        # Ajustar ancho de columnas
        worksheet.set_column(0, 0, 35)
        worksheet.set_column(1, 1, 25)

    def _create_metadata_sheet(self, writer, formats: Dict[str, Any]):
        """
        Crear hoja de metadatos del reporte
        
        Args:
            writer: ExcelWriter object
            formats: Diccionario de formatos
        """
        worksheet = writer.book.add_worksheet('Metadatos')
        
        # Aplicar formato básico
        self._apply_basic_formatting(worksheet, 'METADATOS DEL REPORTE', 
                                   'Información de generación y archivos procesados', formats)
        
        # Obtener información de archivos
        file_info = self.get_file_info()
        
        # Escribir metadatos
        start_row = 4
        metadata = [
            ('Fecha de generación', datetime.now().strftime('%d-%m-%Y %H:%M:%S')),
            ('Procesador', 'CausacionProcessor v1.0'),
            ('', ''),
            ('ARCHIVOS PROCESADOS', ''),
            ('Archivo DIAN', file_info.get('dian_file', 'No cargado')),
            ('Registros DIAN', file_info.get('dian_rows', 0)),
            ('Columnas DIAN', file_info.get('dian_columns', 0)),
            ('', ''),
            ('Archivo Contable', file_info.get('contable_file', 'No cargado')),
            ('Registros Contable', file_info.get('contable_rows', 0)),
            ('Columnas Contable', file_info.get('contable_columns', 0)),
            ('', ''),
            ('CONFIGURACIÓN', ''),
            ('Tolerancia de valores', '0.01'),
            ('Tolerancia de fechas', '1 día'),
            ('Nivel de similitud mínimo', '0.8')
        ]
        
        for row_idx, (label, value) in enumerate(metadata, start_row):
            try:
                worksheet.write(row_idx, 0, label, formats['info'])
                worksheet.write(row_idx, 1, value, formats['data'])
            except Exception:
                # Fallback sin formato si hay error
                worksheet.write(row_idx, 0, label)
                worksheet.write(row_idx, 1, value)
        
        # Ajustar ancho de columnas
        worksheet.set_column(0, 0, 35)
        worksheet.set_column(1, 1, 40)

    def _apply_basic_formatting(self, worksheet, title: str, subtitle: str, formats: Dict[str, Any]):
        """
        Aplicar formato básico a una hoja de Excel
        
        Args:
            worksheet: Objeto worksheet
            title: Título principal
            subtitle: Subtítulo
            formats: Diccionario de formatos
        """
        # Título principal
        worksheet.merge_range('A1:O1', title, formats['title'])
        
        # Subtítulo
        worksheet.merge_range('A2:O2', subtitle, formats['subtitle'])
        
        # Línea separadora
        worksheet.merge_range('A3:O3', '', formats['data'])

    def _get_status_format(self, status: str, formats: Dict[str, Any]):
        """
        Obtener formato específico según el estado de validación
        
        Args:
            status: Estado de validación
            formats: Diccionario de formatos
            
        Returns:
            Formato específico para el estado
        """
        status_formats = {
            'Perfecta': formats['status_perfecta'],
            'Buena': formats['status_buena'],
            'Regular': formats['status_regular'],
            'Revisar': formats['status_revisar']
        }
        return status_formats.get(status, formats['data'])

    def _add_sheet_summary(self, worksheet, df: pd.DataFrame, start_row: int, formats: Dict[str, Any]):
        """
        Agregar resumen de la hoja
        
        Args:
            worksheet: Objeto worksheet
            df: DataFrame de la hoja
            start_row: Fila donde empezar el resumen
            formats: Diccionario de formatos
        """
        if df.empty:
            return
        
        # Línea separadora
        worksheet.merge_range(f'A{start_row}:O{start_row}', '', formats['data'])
        
        # Información de resumen
        summary_row = start_row + 1
        worksheet.write(summary_row, 0, 'RESUMEN DE LA HOJA:', formats['subtitle'])
        worksheet.write(summary_row + 1, 0, f'Total de registros: {len(df)}', formats['info'])
        
        # Si es la hoja de coincidencias, agregar estadísticas adicionales
        if 'ESTADO VALIDACIÓN' in df.columns:
            perfectas = len(df[df['ESTADO VALIDACIÓN'] == 'Perfecta'])
            buenas = len(df[df['ESTADO VALIDACIÓN'] == 'Buena'])
            regulares = len(df[df['ESTADO VALIDACIÓN'] == 'Regular'])
            revisar = len(df[df['ESTADO VALIDACIÓN'] == 'Revisar'])
            
            worksheet.write(summary_row + 2, 0, f'Coincidencias perfectas: {perfectas}', formats['info'])
            worksheet.write(summary_row + 3, 0, f'Coincidencias buenas: {buenas}', formats['info'])
            worksheet.write(summary_row + 4, 0, f'Coincidencias regulares: {regulares}', formats['info'])
            worksheet.write(summary_row + 5, 0, f'Coincidencias a revisar: {revisar}', formats['info'])

    def apply_conditional_formatting(self, workbook, worksheet, data_range: str, sheet_type: str = 'coincidencias'):
        """
        Aplicar formato condicional avanzado a la hoja
        
        Args:
            workbook: Objeto workbook
            worksheet: Objeto worksheet
            data_range: Rango de datos (ej: 'A5:O100')
            sheet_type: Tipo de hoja ('coincidencias' o 'no_coincidencias')
        """
        try:
            self.logger.info(f"Aplicando formato condicional a hoja {sheet_type}")
            
            # Validar que el workbook y worksheet son válidos
            if not workbook or not worksheet:
                self.logger.warning("Workbook o worksheet no válidos, omitiendo formato condicional")
                return
            
            # Validar que el data_range es válido
            if not data_range or ':' not in data_range:
                self.logger.warning(f"Rango de datos inválido: {data_range}, omitiendo formato condicional")
                return
            
            # Validar que el workbook tiene el método add_format (xlsxwriter)
            if not hasattr(workbook, 'add_format'):
                self.logger.warning("Workbook no soporta add_format, omitiendo formato condicional")
                return
                
            if sheet_type == 'coincidencias':
                self._apply_coincidencias_conditional_formatting(workbook, worksheet, data_range)
            elif sheet_type == 'no_coincidencias':
                self._apply_no_coincidencias_conditional_formatting(workbook, worksheet, data_range)
            else:
                self._apply_general_conditional_formatting(workbook, worksheet, data_range)
            
            self.logger.info("Formato condicional aplicado exitosamente")
            
        except Exception as e:
            self.logger.warning(f"Error al aplicar formato condicional: {e}")
            # Continuar sin formato condicional en lugar de fallar completamente
            self.logger.info("Continuando sin formato condicional...")
            # No hacer raise, solo continuar sin formato condicional

    def _apply_coincidencias_conditional_formatting(self, workbook, worksheet, data_range: str):
        """
        Aplicar formato condicional específico para hoja de coincidencias
        
        Args:
            worksheet: Objeto worksheet
            data_range: Rango de datos
        """
        try:
            # Obtener rango de filas
            start_row = int(data_range.split(":")[0].replace("A", ""))
            end_row = int(data_range.split(":")[1].replace("O", ""))
            
            # Crear formatos una sola vez para reutilizar
            format_green = workbook.add_format({
                'bg_color': '#c6efce',
                'font_color': '#006100',
                'border': 1
            })
            
            format_yellow = workbook.add_format({
                'bg_color': '#ffeb9c',
                'font_color': '#9c6500',
                'border': 1
            })
            
            format_red = workbook.add_format({
                'bg_color': '#ffc7ce',
                'font_color': '#9c0006',
                'border': 1,
                'bold': True
            })
            
            format_highlight = workbook.add_format({
                'bg_color': '#fff2cc',
                'font_color': '#7f6000',
                'border': 1,
                'bold': True
            })
            
            format_confidence_high = workbook.add_format({
                'bg_color': '#d4edda',
                'font_color': '#155724',
                'border': 1
            })
            
            format_confidence_medium = workbook.add_format({
                'bg_color': '#fff3cd',
                'font_color': '#856404',
                'border': 1
            })
            
            format_confidence_low = workbook.add_format({
                'bg_color': '#f8d7da',
                'font_color': '#721c24',
                'border': 1,
                'bold': True
            })
            
            # Aplicar formato condicional para diferencias de valor (columna K)
            # Verde para diferencias menores a 1
            worksheet.conditional_format(
                f'K{start_row}:K{end_row}',
                {
                    'type': 'cell',
                    'criteria': 'between',
                    'minimum': -1,
                    'maximum': 1,
                    'format': format_green
                }
            )
            
            # Amarillo para diferencias moderadas (1-100)
            worksheet.conditional_format(
                f'K{start_row}:K{end_row}',
                {
                    'type': 'cell',
                    'criteria': 'between',
                    'minimum': 1.01,
                    'maximum': 100,
                    'format': format_yellow
                }
            )
            
            # Rojo para diferencias grandes (>100)
            worksheet.conditional_format(
                f'K{start_row}:K{end_row}',
                {
                    'type': 'cell',
                    'criteria': '>',
                    'value': 100,
                    'format': format_red
                }
            )
            
            # Formato para valores altos (>1,000,000) en columnas de valor
            worksheet.conditional_format(
                f'C{start_row}:C{end_row}',
                {
                    'type': 'cell',
                    'criteria': '>',
                    'value': 1000000,
                    'format': format_highlight
                }
            )
            
            # Formato para nivel de confianza (columna O)
            # Verde para alta confianza (>=0.9)
            worksheet.conditional_format(
                f'O{start_row}:O{end_row}',
                {
                    'type': 'cell',
                    'criteria': '>=',
                    'value': 0.9,
                    'format': format_confidence_high
                }
            )
            
            # Amarillo para confianza media (0.7-0.89)
            worksheet.conditional_format(
                f'O{start_row}:O{end_row}',
                {
                    'type': 'cell',
                    'criteria': 'between',
                    'minimum': 0.7,
                    'maximum': 0.89,
                    'format': format_confidence_medium
                }
            )
            
            # Rojo para baja confianza (<0.7)
            worksheet.conditional_format(
                f'O{start_row}:O{end_row}',
                {
                    'type': 'cell',
                    'criteria': '<',
                    'value': 0.7,
                    'format': format_confidence_low
                }
            )
            
            self.logger.info("Formato condicional aplicado exitosamente")
            
        except Exception as e:
            self.logger.warning(f"Error al aplicar formato condicional: {e}")
        
        # CÓDIGO ORIGINAL COMENTADO TEMPORALMENTE
        # try:
        #     # Formato condicional para diferencias de valor
        #     worksheet.conditional_format(
        #         f'K{data_range.split(":")[0].replace("A", "")}:K{data_range.split(":")[1].replace("O", "")}',
        #         {
        #             'type': 'cell',
        #             'criteria': 'between',
        #             'minimum': -0.01,
        #             'maximum': 0.01,
        #             'format': self._get_conditional_format(workbook, 'perfect_match')
        #         }
        #     )
        # except Exception as e:
        #     self.logger.warning(f"Error aplicando formato perfect_match: {e}")
        # 
        # try:


    def _apply_no_coincidencias_conditional_formatting(self, workbook, worksheet, data_range: str):
        """
        Aplicar formato condicional específico para hoja de no coincidencias
        
        Args:
            worksheet: Objeto worksheet
            data_range: Rango de datos
        """
        try:
            # Obtener rango de filas
            start_row = int(data_range.split(":")[0].replace("A", ""))
            end_row = int(data_range.split(":")[1].replace("L", ""))
            
            # Crear formatos una sola vez para reutilizar
            format_dian = workbook.add_format({
                'bg_color': '#e6f3ff',
                'font_color': '#0c5460',
                'border': 1
            })
            
            format_contable = workbook.add_format({
                'bg_color': '#fff2e6',
                'font_color': '#8b4513',
                'border': 1
            })
            
            format_high_value = workbook.add_format({
                'bg_color': '#fff2cc',
                'font_color': '#7f6000',
                'border': 1,
                'bold': True
            })
            
            format_blank = workbook.add_format({
                'bg_color': '#f2f2f2',
                'border': 1
            })
            
            # Aplicar formatos condicionales usando los formatos creados
            # Formato para registros solo DIAN (azul claro)
            worksheet.conditional_format(
                f'A{start_row}:A{end_row}',
                {
                    'type': 'text',
                    'criteria': 'containing',
                    'value': 'DIAN',
                    'format': format_dian
                }
            )
            
            # Formato para registros solo CONTABLE (naranja claro)
            worksheet.conditional_format(
                f'A{start_row}:A{end_row}',
                {
                    'type': 'text',
                    'criteria': 'containing',
                    'value': 'CONTABLE',
                    'format': format_contable
                }
            )
            
            # Formato para valores altos en no coincidencias
            worksheet.conditional_format(
                f'C{start_row}:C{end_row}',
                {
                    'type': 'cell',
                    'criteria': '>',
                    'value': 1000000,
                    'format': format_high_value
                }
            )
            
            # Formato para celdas vacías (gris claro)
            worksheet.conditional_format(
                f'B{start_row}:L{end_row}',
                {
                    'type': 'blanks',
                    'format': format_blank
                }
            )
            
            self.logger.info("Formato condicional para no coincidencias aplicado exitosamente")
            
        except Exception as e:
            self.logger.warning(f"Error al aplicar formato condicional no coincidencias: {e}")

    def _apply_general_conditional_formatting(self, workbook, worksheet, data_range: str):
        """
        Aplicar formato condicional general
        
        Args:
            worksheet: Objeto worksheet
            data_range: Rango de datos
        """
        try:
            # Obtener rango de filas para formato general
            start_row = int(data_range.split(":")[0].replace("A", ""))
            end_row = int(data_range.split(":")[1].split(":")[0][-1])
            
            # Crear formato para celdas vacías
            format_blank = workbook.add_format({
                'bg_color': '#f2f2f2',
                'border': 1
            })
            
            # Aplicar formato general para celdas vacías
            worksheet.conditional_format(
                data_range,
                {
                    'type': 'blanks',
                    'format': format_blank
                }
            )
            
            self.logger.info("Formato condicional general aplicado exitosamente")
            
        except Exception as e:
            self.logger.warning(f"Error al aplicar formato condicional general: {e}")

    def _get_conditional_format(self, workbook, format_type: str):
        """
        Obtener formato condicional específico con manejo de errores mejorado
        
        Args:
            workbook: Objeto workbook
            format_type: Tipo de formato
            
        Returns:
            Formato condicional
        """
        try:
            if format_type == 'perfect_match':
                return workbook.add_format({
                    'bg_color': '#c6efce',
                    'font_color': '#006100',
                    'border': 1
                })
            elif format_type == 'minor_difference':
                return workbook.add_format({
                    'bg_color': '#ffeb9c',
                    'font_color': '#9c6500',
                    'border': 1
                })
            elif format_type == 'major_difference':
                return workbook.add_format({
                    'bg_color': '#ffc7ce',
                    'font_color': '#9c0006',
                    'border': 1
                })
            elif format_type == 'high_value':
                return workbook.add_format({
                    'bg_color': '#ffd966',
                    'font_color': '#7c4a00',
                    'border': 1,
                    'bold': True
                })
            elif format_type == 'high_confidence':
                return workbook.add_format({
                    'bg_color': '#d4edda',
                    'font_color': '#155724',
                    'border': 1
                })
            elif format_type == 'medium_confidence':
                return workbook.add_format({
                    'bg_color': '#fff3cd',
                    'font_color': '#856404',
                    'border': 1
                })
            elif format_type == 'low_confidence':
                return workbook.add_format({
                    'bg_color': '#f8d7da',
                    'font_color': '#721c24',
                    'border': 1
                })
            elif format_type == 'dian_only':
                return workbook.add_format({
                    'bg_color': '#e6f3ff',
                    'font_color': '#1f4e79',
                    'border': 1
                })
            elif format_type == 'contable_only':
                return workbook.add_format({
                    'bg_color': '#fff2e6',
                    'font_color': '#cc6600',
                    'border': 1
                })
            elif format_type == 'empty_cell':
                return workbook.add_format({
                    'bg_color': '#f2f2f2',
                    'font_color': '#666666',
                    'border': 1,
                    'italic': True
                })
            else:
                return workbook.add_format({'border': 1})
                
        except Exception as e:
            self.logger.warning(f"Error creando formato {format_type}: {e}")
            # Retornar formato básico como fallback
            return workbook.add_format({'border': 1})

    def add_filters_and_sorting(self, workbook, worksheet, header_row: int, data_range: str):
        """
        Agregar filtros automáticos y configuración de ordenamiento
        
        Args:
            worksheet: Objeto worksheet
            header_row: Fila de encabezados
            data_range: Rango de datos
        """
        try:
            self.logger.info("Agregando filtros y ordenamiento")
            
            # Configurar ordenamiento por defecto (por folio DIAN)
            worksheet.set_row(header_row, None, {'hidden': False})
            
            # Agregar configuración de tabla dinámica (incluye autofilter)
            self._add_pivot_table_config(workbook, worksheet, header_row, data_range)
            
            self.logger.info("Filtros y ordenamiento configurados exitosamente")
            
        except Exception as e:
            self.logger.error(f"Error al agregar filtros y ordenamiento: {e}")

    def _add_pivot_table_config(self, workbook, worksheet, header_row: int, data_range: str):
        """
        Agregar configuración para tabla dinámica
        
        Args:
            worksheet: Objeto worksheet
            header_row: Fila de encabezados
            data_range: Rango de datos
        """
        try:
            # Crear hoja de tabla dinámica si es necesario
            
            # Agregar configuración de datos para tabla dinámica
            worksheet.set_row(header_row, None, {
                'outline_level': 1,
                'hidden': False,
                'collapsed': False
            })
            
            # Detectar tipo de hoja para configurar tabla apropiada
            sheet_name = worksheet.get_name().lower()
            
            # Configurar tabla según el tipo de hoja
            if 'coincidencias' in sheet_name and 'no' not in sheet_name:
                # Tabla para coincidencias - estilo azul
                table_style = 'Table Style Medium 9'
                table_columns = [
                    {'header': 'FOLIO DIAN'},
                    {'header': 'FECHA DIAN'},
                    {'header': 'VALOR DIAN'},
                    {'header': 'DESCRIPCIÓN DIAN'},
                    {'header': 'TIPO DOCUMENTO DIAN'},
                    {'header': 'NÚMERO DOCUMENTO CRUCE'},
                    {'header': 'FECHA CONTABLE'},
                    {'header': 'VALOR CONTABLE'},
                    {'header': 'DESCRIPCIÓN CONTABLE'},
                    {'header': 'CUENTA CONTABLE'},
                    {'header': 'DIFERENCIA VALOR'},
                    {'header': 'DIFERENCIA FECHA'},
                    {'header': 'ESTADO VALIDACIÓN'},
                    {'header': 'TIPO COINCIDENCIA'},
                    {'header': 'NIVEL CONFIANZA'}
                ]
            elif 'no' in sheet_name and 'coincidencias' in sheet_name:
                # Tabla para no coincidencias - estilo naranja
                table_style = 'Table Style Medium 7'
                table_columns = [
                    {'header': 'FOLIO DIAN'},
                    {'header': 'FECHA DIAN'},
                    {'header': 'VALOR DIAN'},
                    {'header': 'DESCRIPCIÓN DIAN'},
                    {'header': 'TIPO DOCUMENTO DIAN'},
                    {'header': 'NÚMERO DOCUMENTO CRUCE'},
                    {'header': 'FECHA CONTABLE'},
                    {'header': 'VALOR CONTABLE'},
                    {'header': 'DESCRIPCIÓN CONTABLE'},
                    {'header': 'CUENTA CONTABLE'},
                    {'header': 'MOTIVO NO COINCIDENCIA'},
                    {'header': 'ORIGEN'}
                ]
            else:
                # Tabla genérica
                table_style = 'Table Style Medium 2'
                table_columns = [{'header': f'Columna {i+1}'} for i in range(12)]
            
            # Configurar tabla profesional
            worksheet.add_table(
                f'A{header_row}:{data_range.split(":")[1]}',
                {
                    'style': table_style,
                    'first_column': True,     # Resaltar primera columna
                    'last_column': True,      # Resaltar última columna
                    'banded_rows': True,      # Filas alternadas
                    'banded_columns': False,  # Sin columnas alternadas
                    'columns': table_columns
                }
            )
            
        except Exception as e:
            self.logger.warning(f"No se pudo configurar tabla dinámica: {e}")

    def add_contador_tools(self, workbook, worksheet, data_range: str, sheet_type: str = 'coincidencias'):
        """
        Agregar herramientas útiles para contadores
        
        Args:
            workbook: Objeto workbook
            worksheet: Objeto worksheet
            data_range: Rango de datos
            sheet_type: Tipo de hoja
        """
        try:
            self.logger.info("Agregando herramientas para contadores")
            
            # Calcular posición para herramientas
            start_row = int(data_range.split(":")[1].replace("O", "").replace("L", "")) + 10
            
            # Agregar fórmulas de suma automática
            self._add_summary_formulas(workbook, worksheet, data_range, start_row)
            
            # Agregar validaciones de datos
            self._add_data_validations(workbook, worksheet, data_range)
            
            # Agregar alertas para discrepancias
            self._add_discrepancy_alerts(workbook, worksheet, data_range, start_row + 5)
            
            # Agregar herramientas de análisis
            self._add_analysis_tools(workbook, worksheet, data_range, start_row + 10)
            
            self.logger.info("Herramientas para contadores agregadas exitosamente")
            
        except Exception as e:
            self.logger.error(f"Error al agregar herramientas para contadores: {e}")

    def _add_summary_formulas(self, workbook, worksheet, data_range: str, start_row: int):
        """
        Agregar fórmulas de suma automática
        
        Args:
            worksheet: Objeto worksheet
            data_range: Rango de datos
            start_row: Fila donde empezar las fórmulas
        """
        # Fórmula para suma total de valores DIAN
        worksheet.write(start_row, 0, 'SUMA TOTAL VALORES DIAN:', self._get_info_format(workbook))
        worksheet.write_formula(start_row, 1, f'=SUM(C{data_range.split(":")[0].replace("A", "")}:C{data_range.split(":")[1].replace("O", "").replace("L", "")})', 
                               self._get_number_format(workbook))
        
        # Fórmula para suma total de valores contables
        worksheet.write(start_row + 1, 0, 'SUMA TOTAL VALORES CONTABLES:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 1, 1, f'=SUM(H{data_range.split(":")[0].replace("A", "")}:H{data_range.split(":")[1].replace("O", "").replace("L", "")})', 
                               self._get_number_format(workbook))
        
        # Fórmula para diferencia total
        worksheet.write(start_row + 2, 0, 'DIFERENCIA TOTAL:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 2, 1, f'=B{start_row + 1}-B{start_row + 2}', 
                               self._get_number_format(workbook))
        
        # Fórmula para promedio de valores
        worksheet.write(start_row + 3, 0, 'PROMEDIO VALORES:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 3, 1, f'=AVERAGE(C{data_range.split(":")[0].replace("A", "")}:C{data_range.split(":")[1].replace("O", "").replace("L", "")})', 
                               self._get_number_format(workbook))
        
        # Fórmula para conteo de registros
        worksheet.write(start_row + 4, 0, 'TOTAL REGISTROS:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 4, 1, f'=COUNTA(A{data_range.split(":")[0].replace("A", "")}:A{data_range.split(":")[1].replace("O", "").replace("L", "")})', 
                               self._get_number_format(workbook))

    def _add_data_validations(self, workbook, worksheet, data_range: str):
        """
        Agregar validaciones de datos
        
        Args:
            worksheet: Objeto worksheet
            data_range: Rango de datos
        """
        try:
            # Validación para valores numéricos positivos
            worksheet.data_validation(
                f'C{data_range.split(":")[0].replace("A", "")}:C{data_range.split(":")[1].replace("O", "").replace("L", "")}',
                {
                    'validate': 'decimal',
                    'criteria': '>=',
                    'value': 0,
                    'input_title': 'Validación de Valor',
                    'input_message': 'El valor debe ser un número positivo',
                    'error_title': 'Error de Validación',
                    'error_message': 'Por favor ingrese un valor numérico positivo'
                }
            )
            
            # Validación para fechas
            worksheet.data_validation(
                f'B{data_range.split(":")[0].replace("A", "")}:B{data_range.split(":")[1].replace("O", "").replace("L", "")}',
                {
                    'validate': 'date',
                    'criteria': 'between',
                    'minimum': '1900-01-01',
                    'maximum': '2100-12-31',
                    'input_title': 'Validación de Fecha',
                    'input_message': 'La fecha debe estar entre 1900 y 2100',
                    'error_title': 'Error de Fecha',
                    'error_message': 'Por favor ingrese una fecha válida'
                }
            )
            
        except Exception as e:
            self.logger.warning(f"No se pudieron agregar todas las validaciones: {e}")

    def _add_discrepancy_alerts(self, workbook, worksheet, data_range: str, start_row: int):
        """
        Agregar alertas para discrepancias
        
        Args:
            worksheet: Objeto worksheet
            data_range: Rango de datos
            start_row: Fila donde empezar las alertas
        """
        # Alerta para diferencias de valor mayores a 10%
        worksheet.write(start_row, 0, 'ALERTAS DE DISCREPANCIAS:', self._get_alert_format(workbook))
        
        # Contar registros con diferencias significativas
        worksheet.write(start_row + 1, 0, 'Registros con diferencia > 10%:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 1, 1, 
                               f'=COUNTIF(K{data_range.split(":")[0].replace("A", "")}:K{data_range.split(":")[1].replace("O", "").replace("L", "")},">10")', 
                               self._get_number_format(workbook))
        
        # Contar registros con diferencias de fecha > 7 días
        worksheet.write(start_row + 2, 0, 'Registros con diferencia fecha > 7 días:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 2, 1, 
                               f'=COUNTIF(L{data_range.split(":")[0].replace("A", "")}:L{data_range.split(":")[1].replace("O", "").replace("L", "")},">7")', 
                               self._get_number_format(workbook))
        
        # Contar registros con nivel de confianza bajo
        worksheet.write(start_row + 3, 0, 'Registros con confianza < 0.7:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 3, 1, 
                               f'=COUNTIF(O{data_range.split(":")[0].replace("A", "")}:O{data_range.split(":")[1].replace("O", "").replace("L", "")},"<0.7")', 
                               self._get_number_format(workbook))

    def _add_analysis_tools(self, workbook, worksheet, data_range: str, start_row: int):
        """
        Agregar herramientas de análisis
        
        Args:
            worksheet: Objeto worksheet
            data_range: Rango de datos
            start_row: Fila donde empezar las herramientas
        """
        # Herramientas de análisis estadístico
        worksheet.write(start_row, 0, 'HERRAMIENTAS DE ANÁLISIS:', self._get_alert_format(workbook))
        
        # Valor máximo
        worksheet.write(start_row + 1, 0, 'Valor máximo DIAN:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 1, 1, 
                               f'=MAX(C{data_range.split(":")[0].replace("A", "")}:C{data_range.split(":")[1].replace("O", "").replace("L", "")})', 
                               self._get_number_format(workbook))
        
        # Valor mínimo
        worksheet.write(start_row + 2, 0, 'Valor mínimo DIAN:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 2, 1, 
                               f'=MIN(C{data_range.split(":")[0].replace("A", "")}:C{data_range.split(":")[1].replace("O", "").replace("L", "")})', 
                               self._get_number_format(workbook))
        
        # Desviación estándar
        worksheet.write(start_row + 3, 0, 'Desviación estándar DIAN:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 3, 1, 
                               f'=STDEV(C{data_range.split(":")[0].replace("A", "")}:C{data_range.split(":")[1].replace("O", "").replace("L", "")})', 
                               self._get_number_format(workbook))
        
        # Mediana
        worksheet.write(start_row + 4, 0, 'Mediana DIAN:', self._get_info_format(workbook))
        worksheet.write_formula(start_row + 4, 1, 
                               f'=MEDIAN(C{data_range.split(":")[0].replace("A", "")}:C{data_range.split(":")[1].replace("O", "").replace("L", "")})', 
                               self._get_number_format(workbook))

    def _get_info_format(self, workbook):
        """Obtener formato para información"""
        return workbook.add_format({
            'bold': True,
            'font_size': 10,
            'font_color': '#1f4e79',
            'bg_color': '#e6f3ff',
            'border': 1
        })

    def _get_alert_format(self, workbook):
        """Obtener formato para alertas"""
        return workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': '#721c24',
            'bg_color': '#f8d7da',
            'border': 1
        })

    def _get_number_format(self, workbook):
        """Obtener formato para números"""
        return workbook.add_format({
            'font_size': 10,
            'num_format': '#,##0.00',
            'border': 1,
            'align': 'right'
        })