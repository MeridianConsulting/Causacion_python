#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configuración del Proyecto
"""

from pathlib import Path

class Config:
    """Configuraciones generales del proyecto"""
    
    # Rutas del proyecto
    BASE_PATH = Path(__file__).parent
    DATA_PATH = BASE_PATH / "data"
    INPUT_PATH = DATA_PATH / "input"
    OUTPUT_PATH = DATA_PATH / "output"
    
    # Configuraciones de Excel
    DEFAULT_SHEET_NAME = "Sheet1"
    DATE_FORMAT = "%d/%m/%Y"
    
    # Crear directorios si no existen
    INPUT_PATH.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.mkdir(parents=True, exist_ok=True) 