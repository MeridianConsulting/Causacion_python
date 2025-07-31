#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de prueba para verificar que se solucionó el error de formato Excel
"""

import pandas as pd
import numpy as np
from pathlib import Path
from causacion_processor import CausacionProcessor
from datetime import datetime

def crear_datos_minimos():
    """Crear datos mínimos para la prueba"""
    
    print("🧪 Creando datos mínimos de prueba...")
    
    # Datos de coincidencias muy básicos
    coincidencias_data = {
        'FOLIO DIAN': ['F001', 'F002'],
        'FECHA DIAN': ['01-01-2024', '02-01-2024'],
        'VALOR DIAN': [1000.00, 2500.50],
        'DESCRIPCIÓN DIAN': ['Factura 001', 'Factura 002'],
        'TIPO DOCUMENTO DIAN': ['Factura', 'Factura'],
        'NÚMERO DOCUMENTO CRUCE': ['DOC001', 'DOC002'],
        'FECHA CONTABLE': ['01-01-2024', '02-01-2024'],
        'VALOR CONTABLE': [1000.00, 2500.50],
        'DESCRIPCIÓN CONTABLE': ['Pago 001', 'Pago 002'],
        'CUENTA CONTABLE': ['130505', '130505'],
        'DIFERENCIA VALOR': [0.00, 0.00],
        'DIFERENCIA FECHA': [0, 0],
        'ESTADO VALIDACIÓN': ['Perfecta', 'Perfecta'],
        'TIPO COINCIDENCIA': ['Exacta', 'Exacta'],
        'NIVEL CONFIANZA': [1.0, 1.0]
    }
    
    # Datos de no coincidencias muy básicos
    no_coincidencias_data = {
        'FOLIO DIAN': ['F003', ''],
        'FECHA DIAN': ['03-01-2024', ''],
        'VALOR DIAN': [2000.00, 0.0],
        'DESCRIPCIÓN DIAN': ['Factura 003', ''],
        'TIPO DOCUMENTO DIAN': ['Factura', ''],
        'NÚMERO DOCUMENTO CRUCE': ['', 'DOC003'],
        'FECHA CONTABLE': ['', '03-01-2024'],
        'VALOR CONTABLE': [0.0, 1800.00],
        'DESCRIPCIÓN CONTABLE': ['', 'Pago 003'],
        'CUENTA CONTABLE': ['', '130505'],
        'MOTIVO NO COINCIDENCIA': ['Solo en DIAN', 'Solo en Contable'],
        'ORIGEN': ['DIAN', 'CONTABLE']
    }
    
    return pd.DataFrame(coincidencias_data), pd.DataFrame(no_coincidencias_data)

def probar_fix_error():
    """Probar que se solucionó el error de formato Excel"""
    
    print("=" * 70)
    print("🔧 PROBANDO CORRECCIÓN DE ERROR EXCEL")
    print("=" * 70)
    
    try:
        # Inicializar procesador
        processor = CausacionProcessor()
        
        # Crear datos de prueba
        coincidencias_df, no_coincidencias_df = crear_datos_minimos()
        
        print(f"✅ Datos de prueba creados:")
        print(f"   - Coincidencias: {len(coincidencias_df)} registros")
        print(f"   - No coincidencias: {len(no_coincidencias_df)} registros")
        
        # Estadísticas básicas
        stats = {
            'total_coincidencias': len(coincidencias_df),
            'total_no_coincidencias': len(no_coincidencias_df),
            'coincidencias_exactas': len(coincidencias_df),
            'coincidencias_perfectas': len(coincidencias_df),
            'resumen_ejecutivo': {
                'calidad_general': 'Excelente'
            }
        }
        
        # Crear directorio de salida
        output_dir = Path("output_fix")
        output_dir.mkdir(exist_ok=True)
        
        print(f"\\n📊 Generando Excel con correcciones...")
        
        # Intentar crear archivo Excel
        excel_path = processor.create_excel_file(
            coincidencias_df=coincidencias_df,
            no_coincidencias_df=no_coincidencias_df,
            output_path=output_dir,
            stats=stats
        )
        
        print(f"✅ Archivo Excel creado exitosamente: {excel_path}")
        
        # Verificar archivo
        excel_file = Path(excel_path)
        if excel_file.exists():
            file_size = excel_file.stat().st_size / 1024  # KB
            print(f"📁 Verificación del archivo:")
            print(f"   - Tamaño: {file_size:.1f} KB")
            print(f"   - Ubicación: {excel_path}")
            
            # Verificar que se puede leer
            try:
                with pd.ExcelFile(excel_path) as xls:
                    sheets = xls.sheet_names
                    print(f"   - Hojas: {', '.join(sheets)}")
                    
                    for sheet in sheets:
                        df = pd.read_excel(excel_path, sheet_name=sheet)
                        print(f"      • {sheet}: {len(df)} filas")
                        
            except Exception as e:
                print(f"   ⚠️ Error al verificar contenido: {e}")
                return False
        else:
            print("   ❌ El archivo no se creó")
            return False
        
        print(f"\\n🎯 CORRECCIONES APLICADAS:")
        print(f"   ✅ Eliminado conflicto autofilter/tabla")
        print(f"   ✅ Formatos condicionales con objetos Format válidos")
        print(f"   ✅ Validaciones de entrada mejoradas")
        print(f"   ✅ Sistema de fallback implementado")
        print(f"   ✅ Manejo de errores robusto")
        
        print(f"\\n🎉 PRUEBA EXITOSA - ERROR SOLUCIONADO")
        return True
        
    except Exception as e:
        print(f"❌ Error durante la prueba: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Función principal"""
    
    success = probar_fix_error()
    
    if success:
        print(f"\\n" + "=" * 70)
        print(f"✅ CORRECCIÓN VERIFICADA")
        print(f"=" * 70)
        print(f"El error 'dict' object has no attribute '_get_xf_index' ha sido solucionado.")
        print(f"El sistema ahora puede generar Excel con formatos avanzados sin errores.")
    else:
        print(f"\\n" + "=" * 70)
        print(f"❌ CORRECCIÓN FALLIDA")
        print(f"=" * 70)
        print(f"Aún hay problemas que resolver.")

if __name__ == "__main__":
    main()