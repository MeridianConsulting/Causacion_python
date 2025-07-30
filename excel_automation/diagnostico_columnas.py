#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de diagnóstico para verificar las columnas de los archivos DIAN y contable
"""

import pandas as pd
import sys
from pathlib import Path

def diagnosticar_archivos():
    """Diagnosticar estructura de archivos DIAN y contable"""
    
    # Rutas de los archivos (ajusta según tus archivos)
    dian_file = Path("../data/input/17_Julio_2025_Dian.xlsx")
    contable_file = Path("../data/input/movimientocontable.xlsx")
    
    print("=" * 60)
    print("🔍 DIAGNÓSTICO DE COLUMNAS DE ARCHIVOS")
    print("=" * 60)
    
    # Verificar archivos DIAN
    if dian_file.exists():
        print(f"\n📊 ARCHIVO DIAN: {dian_file.name}")
        print("-" * 40)
        
        try:
            df_dian = pd.read_excel(dian_file)
            print(f"   Filas: {len(df_dian)}")
            print(f"   Columnas: {len(df_dian.columns)}")
            print(f"\n   COLUMNAS DISPONIBLES:")
            for i, col in enumerate(df_dian.columns, 1):
                print(f"   {i:2d}. {col}")
                
            # Buscar columnas candidatas para documento
            print(f"\n   🎯 CANDIDATOS PARA COLUMNA DE DOCUMENTO:")
            candidatos_dian = []
            for col in df_dian.columns:
                col_lower = col.lower()
                if any(keyword in col_lower for keyword in ['folio', 'numero', 'documento', 'factura']):
                    candidatos_dian.append(col)
                    # Mostrar algunos valores de ejemplo
                    valores_ejemplo = df_dian[col].dropna().head(3).tolist()
                    print(f"      ✓ {col} -> Ejemplos: {valores_ejemplo}")
                    
            if not candidatos_dian:
                print("      ❌ No se encontraron candidatos obvios")
                print("      🔍 Todas las columnas:")
                for col in df_dian.columns[:10]:  # Primeras 10 columnas
                    valores_ejemplo = df_dian[col].dropna().head(2).tolist()
                    print(f"         - {col} -> {valores_ejemplo}")
                    
        except Exception as e:
            print(f"   ❌ Error al leer archivo DIAN: {e}")
    else:
        print(f"\n❌ ARCHIVO DIAN NO ENCONTRADO: {dian_file}")
    
    # Verificar archivo contable
    if contable_file.exists():
        print(f"\n📊 ARCHIVO CONTABLE: {contable_file.name}")
        print("-" * 40)
        
        try:
            df_contable = pd.read_excel(contable_file)
            print(f"   Filas: {len(df_contable)}")
            print(f"   Columnas: {len(df_contable.columns)}")
            
            # Saltar las primeras 4 filas como hace el procesador
            if len(df_contable) > 4:
                df_contable_clean = df_contable.iloc[4:].reset_index(drop=True)
                print(f"   Filas después de limpiar metadatos: {len(df_contable_clean)}")
                
                print(f"\n   COLUMNAS DISPONIBLES (después de limpiar):")
                for i, col in enumerate(df_contable_clean.columns, 1):
                    print(f"   {i:2d}. {col}")
                    
                # Buscar columnas candidatas para documento
                print(f"\n   🎯 CANDIDATOS PARA COLUMNA DE DOCUMENTO:")
                candidatos_contable = []
                for col in df_contable_clean.columns:
                    col_lower = col.lower()
                    if any(keyword in col_lower for keyword in ['numero', 'documento', 'cruce', 'factura', 'comprobante']):
                        candidatos_contable.append(col)
                        # Mostrar algunos valores de ejemplo
                        valores_ejemplo = df_contable_clean[col].dropna().head(3).tolist()
                        print(f"      ✓ {col} -> Ejemplos: {valores_ejemplo}")
                        
                if not candidatos_contable:
                    print("      ❌ No se encontraron candidatos obvios")
                    print("      🔍 Primeras columnas:")
                    for col in df_contable_clean.columns[:10]:
                        valores_ejemplo = df_contable_clean[col].dropna().head(2).tolist()
                        print(f"         - {col} -> {valores_ejemplo}")
                        
        except Exception as e:
            print(f"   ❌ Error al leer archivo contable: {e}")
    else:
        print(f"\n❌ ARCHIVO CONTABLE NO ENCONTRADO: {contable_file}")
    
    print("\n" + "=" * 60)
    print("💡 RECOMENDACIONES:")
    print("=" * 60)
    print("1. Identifica cuál columna de DIAN contiene números de documento")
    print("2. Identifica cuál columna contable contiene números de documento")
    print("3. Ajusta el mapeo de columnas en el código")
    print("\n📧 Comparte este resultado para obtener ayuda específica")

if __name__ == "__main__":
    diagnosticar_archivos()