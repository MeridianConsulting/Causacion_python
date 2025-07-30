#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Paquete de Automatización Excel
"""

__version__ = "1.0.0"
__author__ = "Tu Nombre"

# Importar clases principales
from .excel_processor import ExcelProcessor
from .causacion_processor import CausacionProcessor

__all__ = ['ExcelProcessor', 'CausacionProcessor'] 