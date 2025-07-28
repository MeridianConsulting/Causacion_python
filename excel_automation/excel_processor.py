#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Procesador de Archivos Excel
"""

import pandas as pd
import openpyxl
from pathlib import Path
from typing import Dict, List, Any

class ExcelProcessor:
    """Clase para procesar archivos Excel"""
    
    def __init__(self):
        """Inicializar el procesador"""
        self.workbook = None
        self.worksheet = None
        
    def read_excel(self, file_path: Path, sheet_name: str = None) -> pd.DataFrame:
        """
        Leer archivo Excel y retornar DataFrame
        
        Args:
            file_path: Ruta del archivo Excel
            sheet_name: Nombre de la hoja (opcional)
            
        Returns:
            DataFrame con los datos del Excel
        """
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"Archivo leído exitosamente: {len(df)} filas")
            return df
        except Exception as e:
            raise Exception(f"Error al leer el archivo Excel: {e}")
    
    def write_excel(self, df: pd.DataFrame, file_path: Path, sheet_name: str = "Sheet1"):
        """
        Escribir DataFrame a archivo Excel
        
        Args:
            df: DataFrame a escribir
            file_path: Ruta de destino
            sheet_name: Nombre de la hoja
        """
        try:
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado exitosamente: {file_path}")
        except Exception as e:
            raise Exception(f"Error al escribir el archivo Excel: {e}")
    
    def process_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Procesar los datos del DataFrame
        
        Args:
            df: DataFrame a procesar
            
        Returns:
            DataFrame procesado
        """
        # Aquí puedes agregar tu lógica de procesamiento
        # Ejemplo básico:
        processed_df = df.copy()
        
        # Eliminar filas vacías
        processed_df = processed_df.dropna(how='all')
        
        # Eliminar espacios en blanco de las columnas de texto
        for col in processed_df.select_dtypes(include=['object']).columns:
            processed_df[col] = processed_df[col].astype(str).str.strip()
        
        print(f"Datos procesados: {len(processed_df)} filas")
        return processed_df
    
    def process_file(self, input_path: Path, output_path: Path):
        """
        Procesar archivo completo
        
        Args:
            input_path: Ruta del archivo de entrada
            output_path: Ruta del archivo de salida
        """
        # Leer el archivo
        df = self.read_excel(input_path)
        
        # Procesar los datos
        processed_df = self.process_data(df)
        
        # Guardar el resultado
        self.write_excel(processed_df, output_path)
    
    def get_sheet_names(self, file_path: Path) -> List[str]:
        """
        Obtener nombres de todas las hojas del archivo Excel
        
        Args:
            file_path: Ruta del archivo Excel
            
        Returns:
            Lista con nombres de las hojas
        """
        try:
            workbook = openpyxl.load_workbook(file_path)
            return workbook.sheetnames
        except Exception as e:
            raise Exception(f"Error al obtener nombres de hojas: {e}") 