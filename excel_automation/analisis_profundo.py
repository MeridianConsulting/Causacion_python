#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Análisis profundo de la estructura de datos para encontrar la columna correcta
"""

import pandas as pd
from pathlib import Path
import re

def analizar_datos_profundo():
    """Análisis profundo de los datos"""
    
    print("=" * 70)
    print("🔬 ANÁLISIS PROFUNDO DE DATOS")
    print("=" * 70)
    
    # Cargar archivos raw
    dian_file = Path("../data/input/17_Julio_2025_Dian.xlsx")
    contable_file = Path("../data/input/movimientocontable.xlsx")
    
    # Analizar DIAN
    print(f"\n📊 ANÁLISIS DETALLADO - ARCHIVO DIAN")
    print("-" * 50)
    
    df_dian = pd.read_excel(dian_file)
    print(f"Folio ejemplos: {df_dian['Folio'].dropna().head(10).tolist()}")
    print(f"Total ejemplos únicos: {df_dian['Folio'].nunique()}")
    
    # Buscar patrones en DIAN
    folio_lengths = df_dian['Folio'].astype(str).str.len().value_counts()
    print(f"Longitudes de Folio: {dict(folio_lengths)}")
    
    # Analizar contable
    print(f"\n📊 ANÁLISIS DETALLADO - ARCHIVO CONTABLE")
    print("-" * 50)
    
    df_contable_raw = pd.read_excel(contable_file)
    df_contable = df_contable_raw.iloc[4:].reset_index(drop=True)  # Skip metadata
    
    print(f"Total columnas contable: {len(df_contable.columns)}")
    
    # Analizar todas las columnas que podrían contener números de documento
    print(f"\n🔍 ANÁLISIS DE COLUMNAS CANDIDATAS:")
    
    for i, col in enumerate(df_contable.columns):
        if i > 20:  # Analizar solo las primeras 20 columnas
            break
            
        sample_values = df_contable[col].dropna().head(10)
        if len(sample_values) == 0:
            continue
            
        # Convertir a string para análisis
        str_values = sample_values.astype(str).tolist()
        
        # Buscar números que podrían ser documentos
        numeric_values = []
        for val in str_values:
            val_clean = str(val).strip()
            if val_clean.isdigit():
                numeric_values.append(val_clean)
        
        if numeric_values:
            lengths = [len(v) for v in numeric_values]
            unique_lengths = set(lengths)
            
            print(f"\n   Columna {i:2d} ({col}):")
            print(f"      Valores: {str_values[:5]}")
            print(f"      Longitudes: {unique_lengths}")
            
            # Verificar si algún valor podría estar relacionado con DIAN
            for folio in ['3020045', '3020044', '3020068']:
                for val in numeric_values:
                    if folio in val or val in folio:
                        print(f"      ⭐ POSIBLE COINCIDENCIA: {folio} <-> {val}")
                    elif folio[-4:] in val or val[-4:] in folio:
                        print(f"      🤔 POSIBLE RELACIÓN: {folio} <-> {val} (últimos dígitos)")
    
    # Buscar en todas las columnas valores que contengan patrones de DIAN
    print(f"\n🎯 BÚSQUEDA DE PATRONES DIAN EN CONTABLE:")
    
    folios_dian = df_dian['Folio'].astype(str).tolist()[:20]  # Primeros 20 folios
    
    for folio in folios_dian:
        print(f"\n   Buscando '{folio}' en archivo contable...")
        found = False
        
        for col in df_contable.columns:
            try:
                col_values = df_contable[col].astype(str)
                matches = col_values[col_values.str.contains(folio, na=False)]
                if len(matches) > 0:
                    print(f"      ✅ Encontrado en {col}: {matches.tolist()[:3]}")
                    found = True
            except:
                continue
        
        if not found:
            # Buscar patrones parciales (últimos 4 dígitos)
            partial = folio[-4:]
            for col in df_contable.columns:
                try:
                    col_values = df_contable[col].astype(str)
                    matches = col_values[col_values.str.contains(partial, na=False)]
                    if len(matches) > 0:
                        print(f"      🔍 Patrón parcial '{partial}' en {col}: {matches.tolist()[:2]}")
                        break
                except:
                    continue
    
    print(f"\n" + "=" * 70)
    print("💡 RECOMENDACIONES:")
    print("=" * 70)
    print("1. Si no hay coincidencias exactas, los archivos podrían usar diferentes numeraciones")
    print("2. Considera matching por valor monetario y fecha")
    print("3. Verifica si hay un mapeo/tabla de conversión entre sistemas")
    print("4. Posiblemente los sistemas DIAN y contable usen numeraciones independientes")

if __name__ == "__main__":
    analizar_datos_profundo()