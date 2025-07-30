#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ejemplo de uso: Integración de UI con Procesador de Causación
Demuestra cómo usar la interfaz gráfica integrada con el sistema de causación
"""

import sys
import pandas as pd
import numpy as np
from pathlib import Path
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtCore import QThread, Signal

def create_sample_files():
    """Crear archivos de ejemplo para demostrar la funcionalidad"""
    
    print("=== CREANDO ARCHIVOS DE EJEMPLO ===\n")
    
    # Crear directorio de ejemplo
    example_dir = Path("example_files")
    example_dir.mkdir(exist_ok=True)
    
    # 1. Crear archivo DIAN de ejemplo
    print("1. Creando archivo DIAN de ejemplo...")
    
    dian_data = {
        'Folio': ['F001', 'F002', 'F003', 'F004', 'F005', 'F006', 'F007', 'F008'],
        'Fecha': ['01-01-2024', '02-01-2024', '03-01-2024', '04-01-2024', 
                 '05-01-2024', '06-01-2024', '07-01-2024', '08-01-2024'],
        'Valor': [1000.00, 2500.50, 750.25, 3000.00, 1500.75, 2200.00, 1800.50, 2000.00],
        'Descripción': ['Factura 001', 'Factura 002', 'Factura 003', 'Factura 004',
                       'Factura 005', 'Factura 006', 'Factura 007', 'Factura 008'],
        'Tipo Documento': ['Factura', 'Factura', 'Factura', 'Factura',
                          'Factura', 'Factura', 'Factura', 'Factura']
    }
    
    dian_df = pd.DataFrame(dian_data)
    dian_file = example_dir / "archivo_dian_ejemplo.xlsx"
    dian_df.to_excel(dian_file, index=False)
    print(f"   ✓ Archivo DIAN creado: {dian_file}")
    
    # 2. Crear archivo contable de ejemplo
    print("\n2. Creando archivo contable de ejemplo...")
    
    contable_data = {
        'NÚMERO DE DOCUMENTO CRUCE': ['DOC001', 'DOC002', 'DOC003', 'DOC004', 'DOC005', 'DOC006', 'DOC007'],
        'Año': [2024, 2024, 2024, 2024, 2024, 2024, 2024],
        'Mes': [1, 1, 1, 1, 1, 1, 1],
        'Día': [1, 2, 3, 4, 5, 6, 7],
        'Valor': [1000.00, 2500.50, 750.25, 3000.00, 1500.75, 2200.00, 1800.50],
        'Descripción': ['Pago factura 001', 'Pago factura 002', 'Pago factura 003', 'Pago factura 004',
                       'Pago factura 005', 'Pago factura 006', 'Pago factura 007'],
        'Cuenta': ['130505', '130505', '130505', '130505', '130505', '130505', '130505']
    }
    
    contable_df = pd.DataFrame(contable_data)
    contable_file = example_dir / "archivo_contable_ejemplo.xlsx"
    contable_df.to_excel(contable_file, index=False)
    print(f"   ✓ Archivo contable creado: {contable_file}")
    
    print(f"\n✅ Archivos de ejemplo creados en: {example_dir}")
    print("   - archivo_dian_ejemplo.xlsx")
    print("   - archivo_contable_ejemplo.xlsx")
    
    return dian_file, contable_file

def demonstrate_ui_features():
    """Demostrar características de la interfaz integrada"""
    
    print("\n=== CARACTERÍSTICAS DE LA INTERFAZ INTEGRADA ===")
    
    features = [
        ("🎯 Procesamiento Completo", "Flujo completo de causación desde la UI"),
        ("📊 Progreso Detallado", "Mensajes de progreso en tiempo real"),
        ("🔍 Validación de Archivos", "Verificación automática de archivos de entrada"),
        ("📈 Estadísticas Visuales", "Mostrar estadísticas del proceso en la UI"),
        ("🎨 Interfaz Moderna", "Diseño limpio y profesional"),
        ("🔄 Drag & Drop", "Arrastrar archivos directamente a la interfaz"),
        ("⚡ Procesamiento Asíncrono", "No bloquea la interfaz durante el procesamiento"),
        ("📋 Log Detallado", "Registro completo de todas las operaciones"),
        ("✅ Manejo de Errores", "Mensajes de error claros y útiles"),
        ("📁 Salida Automática", "Archivos Excel generados automáticamente")
    ]
    
    for feature, description in features:
        print(f"   {feature}: {description}")

def demonstrate_processing_flow():
    """Demostrar el flujo de procesamiento"""
    
    print("\n=== FLUJO DE PROCESAMIENTO ===")
    
    steps = [
        ("1. Inicialización", "🔧 Inicializando procesador de causación..."),
        ("2. Carga DIAN", "📄 Cargando archivo DIAN..."),
        ("3. Carga Contable", "📄 Cargando archivo contable..."),
        ("4. Validación", "🔍 Validando archivos..."),
        ("5. Cruce de Datos", "🔗 Realizando cruce de datos..."),
        ("6. Generación", "📊 Generando DataFrames de resultado..."),
        ("7. Estadísticas", "📈 Calculando estadísticas..."),
        ("8. Excel", "📋 Creando archivo Excel profesional..."),
        ("9. Finalización", "✅ Procesamiento completado")
    ]
    
    for step, message in steps:
        print(f"   {step}: {message}")

def show_usage_instructions():
    """Mostrar instrucciones de uso"""
    
    print("\n=== INSTRUCCIONES DE USO ===")
    
    instructions = [
        "1. Ejecutar la aplicación: python -m excel_automation.ui_main",
        "2. Arrastrar archivo DIAN a la zona correspondiente",
        "3. Arrastrar archivo contable a la zona correspondiente",
        "4. Hacer clic en '🚀 Iniciar Causación'",
        "5. Observar el progreso en tiempo real",
        "6. Revisar las estadísticas finales",
        "7. Abrir el archivo Excel generado en la carpeta output/"
    ]
    
    for instruction in instructions:
        print(f"   {instruction}")

def test_processor_integration():
    """Probar la integración del procesador directamente"""
    
    print("\n=== PRUEBA DE INTEGRACIÓN DEL PROCESADOR ===")
    
    try:
        from causacion_processor import CausacionProcessor
        
        # Crear archivos de ejemplo
        dian_file, contable_file = create_sample_files()
        
        # Inicializar procesador
        print("\n🔧 Probando procesador de causación...")
        processor = CausacionProcessor()
        
        # Cargar archivos
        print("📄 Cargando archivos de ejemplo...")
        dian_df = processor.load_dian_file(dian_file)
        contable_df = processor.load_contable_file(contable_file)
        
        print(f"   ✓ DIAN: {len(dian_df)} registros")
        print(f"   ✓ Contable: {len(contable_df)} registros")
        
        # Validar archivos
        is_valid, errors = processor.validate_files()
        if is_valid:
            print("   ✅ Archivos validados correctamente")
        else:
            print(f"   ❌ Errores de validación: {errors}")
            return False
        
        # Realizar matching
        print("🔗 Realizando cruce de datos...")
        matching_result = processor.perform_data_matching(dian_df, contable_df)
        matches_df = matching_result['matches']
        non_matches_df = matching_result['non_matches']
        
        print(f"   ✓ Coincidencias: {len(matches_df)}")
        print(f"   ✓ No coincidencias: {len(non_matches_df)}")
        
        # Generar DataFrames finales
        print("📊 Generando DataFrames de resultado...")
        coincidencias_df = processor.create_coincidencias_dataframe(matches_df)
        no_coincidencias_df = processor.create_no_coincidencias_dataframe(non_matches_df)
        
        # Calcular estadísticas
        print("📈 Calculando estadísticas...")
        stats = processor.calculate_statistics(coincidencias_df, no_coincidencias_df)
        
        print(f"   ✓ Calidad general: {stats['resumen_ejecutivo']['calidad_general']}")
        print(f"   ✓ Porcentaje coincidencias: {stats['porcentaje_coincidencias']:.1f}%")
        
        # Crear archivo Excel
        print("📋 Creando archivo Excel...")
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        
        excel_path = processor.create_excel_file(
            coincidencias_df=coincidencias_df,
            no_coincidencias_df=no_coincidencias_df,
            output_path=output_dir,
            stats=stats
        )
        
        print(f"   ✅ Archivo Excel creado: {Path(excel_path).name}")
        
        print("\n🎉 ¡Prueba de integración completada exitosamente!")
        return True
        
    except Exception as e:
        print(f"❌ Error en la prueba de integración: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Función principal del ejemplo"""
    
    print("=== EJEMPLO: INTEGRACIÓN DE UI CON PROCESADOR DE CAUSACIÓN ===\n")
    
    # Mostrar características
    demonstrate_ui_features()
    
    # Mostrar flujo de procesamiento
    demonstrate_processing_flow()
    
    # Mostrar instrucciones
    show_usage_instructions()
    
    # Crear archivos de ejemplo
    print("\n" + "="*60)
    create_sample_files()
    
    # Probar integración
    print("\n" + "="*60)
    success = test_processor_integration()
    
    if success:
        print("\n✅ Todo listo para usar la interfaz gráfica!")
        print("\n📋 PRÓXIMOS PASOS:")
        print("1. Ejecutar: python -m excel_automation.ui_main")
        print("2. Usar los archivos de ejemplo creados")
        print("3. Observar el procesamiento completo")
        print("4. Revisar el archivo Excel generado")
    else:
        print("\n❌ Hay problemas que resolver antes de usar la interfaz")
    
    print("\n=== EJEMPLO COMPLETADO ===")

if __name__ == "__main__":
    main() 