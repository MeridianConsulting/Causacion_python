#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configuración del Proyecto
"""

import sys
from pathlib import Path

def _get_base_path():
    """
    Obtener la ruta base del proyecto.
    Si está ejecutándose como ejecutable, usa la carpeta del ejecutable.
    Si está ejecutándose como script, usa la carpeta del proyecto.
    """
    # Detectar si está ejecutándose como ejecutable (PyInstaller)
    if getattr(sys, 'frozen', False):
        # Ejecutándose como ejecutable compilado
        # sys.executable apunta al .exe
        return Path(sys.executable).parent
    else:
        # Ejecutándose como script Python
        return Path(__file__).parent

class Config:
    """Configuraciones generales del proyecto"""
    
    # Rutas del proyecto
    BASE_PATH = _get_base_path()
    
    # Si está ejecutándose como ejecutable, usar la carpeta del ejecutable para salida
    # Si está ejecutándose como script, usar la estructura de carpetas normal
    if getattr(sys, 'frozen', False):
        # Ejecutándose como ejecutable: guardar reportes en la misma carpeta del .exe
        OUTPUT_PATH = BASE_PATH
        INPUT_PATH = BASE_PATH / "input"  # Opcional: carpeta input junto al ejecutable
    else:
        # Ejecutándose como script: usar estructura de carpetas del proyecto
        DATA_PATH = BASE_PATH / "data"
        INPUT_PATH = DATA_PATH / "input"
        OUTPUT_PATH = DATA_PATH / "output"
    
    # Configuraciones de Excel
    DEFAULT_SHEET_NAME = "Sheet1"
    DATE_FORMAT = "%d/%m/%Y"
    
    # Crear directorios si no existen
    INPUT_PATH.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.mkdir(parents=True, exist_ok=True) 