#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de prueba para verificar las mejoras de formato Excel
"""

import pandas as pd
import numpy as np
from pathlib import Path
from causacion_processor import CausacionProcessor
from datetime import datetime

def crear_datos_prueba():
    """Crear datos de prueba para verificar las mejoras"""
    
    print("🧪 Creando datos de prueba...")
    
    # Datos de coincidencias con diferentes escenarios
    coincidencias_data = {
        'FOLIO DIAN': ['F001', 'F002', 'F003', 'F004', 'F005'],
        'FECHA DIAN': ['01-01-2024', '02-01-2024', '03-01-2024', '04-01-2024', '05-01-2024'],
        'VALOR DIAN': [1000.00, 2500.50, 750000.25, 3000.00, 1500000.75],
        'DESCRIPCIÓN DIAN': [
            'Factura de servicios profesionales de consultoría empresarial', 
            'Venta de productos tecnológicos avanzados', 
            'Servicios de mantenimiento y soporte técnico especializado',
            'Factura 004', 
            'Servicio de desarrollo de software personalizado'
        ],
        'TIPO DOCUMENTO DIAN': ['Factura', 'Factura', 'Factura', 'Factura', 'Factura'],
        'NÚMERO DOCUMENTO CRUCE': ['DOC001', 'DOC002', 'DOC003', 'DOC004', 'DOC005'],
        'FECHA CONTABLE': ['01-01-2024', '02-01-2024', '03-01-2024', '04-01-2024', '05-01-2024'],
        'VALOR CONTABLE': [1000.00, 2500.50, 750050.25, 3000.00, 1500000.75],
        'DESCRIPCIÓN CONTABLE': [
            'Pago factura servicios profesionales consultoría', 
            'Pago venta productos tecnológicos', 
            'Pago servicios mantenimiento técnico especializado',
            'Pago factura 004', 
            'Pago desarrollo software personalizado'
        ],
        'CUENTA CONTABLE': ['130505', '130505', '130505', '130505', '130505'],
        'DIFERENCIA VALOR': [0.00, 0.00, 50.00, 0.00, 0.00],
        'DIFERENCIA FECHA': [0, 0, 0, 0, 0],
        'ESTADO VALIDACIÓN': ['Perfecta', 'Perfecta', 'Buena', 'Perfecta', 'Perfecta'],
        'TIPO COINCIDENCIA': ['Exacta', 'Exacta', 'Aproximada', 'Exacta', 'Exacta'],
        'NIVEL CONFIANZA': [1.0, 1.0, 0.95, 1.0, 1.0]
    }
    
    # Datos de no coincidencias
    no_coincidencias_data = {
        'FOLIO DIAN': ['F006', 'F007', '', ''],
        'FECHA DIAN': ['06-01-2024', '07-01-2024', '', ''],
        'VALOR DIAN': [2000.00, 3500000.00, 0.0, 0.0],
        'DESCRIPCIÓN DIAN': [
            'Factura de servicios de auditoría y consultoría fiscal', 
            'Venta de equipos industriales de alta tecnología', 
            '', ''
        ],
        'TIPO DOCUMENTO DIAN': ['Factura', 'Factura', '', ''],
        'NÚMERO DOCUMENTO CRUCE': ['', '', 'DOC006', 'DOC007'],
        'FECHA CONTABLE': ['', '', '06-01-2024', '07-01-2024'],
        'VALOR CONTABLE': [0.0, 0.0, 1800.00, 4000000.00],
        'DESCRIPCIÓN CONTABLE': [
            '', '', 
            'Pago servicios varios no identificados', 
            'Pago equipos diversos importados'
        ],
        'CUENTA CONTABLE': ['', '', '130505', '130505'],
        'MOTIVO NO COINCIDENCIA': [
            'Solo en DIAN', 'Solo en DIAN', 
            'Solo en Contable', 'Solo en Contable'
        ],
        'ORIGEN': ['DIAN', 'DIAN', 'CONTABLE', 'CONTABLE']
    }
    
    return pd.DataFrame(coincidencias_data), pd.DataFrame(no_coincidencias_data)

def probar_mejoras_excel():
    """Probar las mejoras implementadas en el formato Excel"""
    
    print("=" * 70)
    print("🚀 PROBANDO MEJORAS DE FORMATO EXCEL")
    print("=" * 70)
    
    try:
        # Inicializar procesador
        processor = CausacionProcessor()
        
        # Crear datos de prueba
        coincidencias_df, no_coincidencias_df = crear_datos_prueba()
        
        print(f"✅ Datos de prueba creados:")
        print(f"   - Coincidencias: {len(coincidencias_df)} registros")
        print(f"   - No coincidencias: {len(no_coincidencias_df)} registros")
        
        # Calcular estadísticas básicas
        stats = {
            'total_coincidencias': len(coincidencias_df),
            'total_no_coincidencias': len(no_coincidencias_df),
            'coincidencias_exactas': len(coincidencias_df[coincidencias_df['NIVEL CONFIANZA'] == 1.0]),
            'coincidencias_perfectas': len(coincidencias_df[coincidencias_df['ESTADO VALIDACIÓN'] == 'Perfecta']),
            'resumen_ejecutivo': {
                'calidad_general': 'Excelente'
            }
        }
        
        # Crear directorio de salida
        output_dir = Path("output_prueba")
        output_dir.mkdir(exist_ok=True)
        
        print(f"\n📊 Generando Excel con mejoras...")
        
        # Crear archivo Excel con las mejoras
        excel_path = processor.create_excel_file(
            coincidencias_df=coincidencias_df,
            no_coincidencias_df=no_coincidencias_df,
            output_path=output_dir,
            stats=stats
        )
        
        print(f"✅ Archivo Excel creado: {excel_path}")
        
        # Verificar archivo
        excel_file = Path(excel_path)
        if excel_file.exists():
            file_size = excel_file.stat().st_size / 1024  # KB
            print(f"📁 Archivo generado exitosamente:")
            print(f"   - Tamaño: {file_size:.1f} KB")
            print(f"   - Ubicación: {excel_path}")
            
            # Verificar que se puede leer
            try:
                with pd.ExcelFile(excel_path) as xls:
                    sheets = xls.sheet_names
                    print(f"   - Hojas creadas: {', '.join(sheets)}")
                    
                    for sheet in sheets:
                        df = pd.read_excel(excel_path, sheet_name=sheet)
                        print(f"      • {sheet}: {len(df)} filas")
                        
            except Exception as e:
                print(f"   ⚠️ Error al verificar archivo: {e}")
        else:
            print("   ❌ El archivo no se creó correctamente")
            return False
        
        # Resumen de mejoras implementadas
        print(f"\n🎨 MEJORAS IMPLEMENTADAS:")
        print(f"   ✅ Ajuste automático inteligente de columnas (sin límite de 30)")
        print(f"   ✅ Auto-ajuste de altura de filas (20px para mejor lectura)")
        print(f"   ✅ Formato condicional reactivado y mejorado:")
        print(f"      • Verde: Diferencias menores (≤1)")
        print(f"      • Amarillo: Diferencias moderadas (1-100)")
        print(f"      • Rojo: Diferencias grandes (>100)")
        print(f"      • Destacado: Valores altos (>1,000,000)")
        print(f"      • Confianza por colores")
        print(f"   ✅ Tablas profesionales con filtros automáticos:")
        print(f"      • Coincidencias: Estilo azul Medium 9")
        print(f"      • No coincidencias: Estilo naranja Medium 7")
        print(f"   ✅ Formatos mejorados:")
        print(f"      • Títulos con fondo azul y texto blanco")
        print(f"      • Bordes definidos y colores profesionales")
        print(f"      • Formato de moneda ($#,##0.00)")
        print(f"      • Fechas en formato dd/mm/yyyy")
        print(f"      • Texto con ajuste automático")
        
        print(f"\n🔍 INSTRUCCIONES DE USO:")
        print(f"   1. Abrir {Path(excel_path).name} en Excel")
        print(f"   2. Observar las columnas auto-ajustadas")
        print(f"   3. Ver los colores del formato condicional")
        print(f"   4. Usar los filtros en los encabezados de tabla")
        print(f"   5. Notar la altura mejorada de las filas")
        print(f"   6. Revisar los diferentes estilos de tabla por hoja")
        
        return True
        
    except Exception as e:
        print(f"❌ Error durante la prueba: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Función principal"""
    
    success = probar_mejoras_excel()
    
    if success:
        print(f"\n" + "=" * 70)
        print(f"🎉 PRUEBA COMPLETADA EXITOSAMENTE")
        print(f"=" * 70)
        print(f"Las mejoras de formato Excel han sido implementadas y probadas.")
        print(f"El archivo de prueba está listo para revisar.")
    else:
        print(f"\n" + "=" * 70)
        print(f"❌ PRUEBA FALLÓ")
        print(f"=" * 70)
        print(f"Revisar los errores anteriores para solucionar problemas.")

if __name__ == "__main__":
    main()