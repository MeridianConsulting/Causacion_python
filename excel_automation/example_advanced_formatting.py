#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ejemplo de uso: Formato Avanzado de Excel
Demuestra las funcionalidades de formato condicional, filtros y herramientas para contadores
"""

import pandas as pd
import numpy as np
from pathlib import Path
from causacion_processor import CausacionProcessor

def main():
    """Ejemplo completo de formato avanzado de Excel"""
    
    print("=== EJEMPLO: FORMATO AVANZADO DE EXCEL ===\n")
    
    # Inicializar el procesador
    processor = CausacionProcessor()
    
    try:
        # 1. Crear datos de ejemplo con variaciones para demostrar formato condicional
        print("1. Creando datos de ejemplo con variaciones...")
        
        # Datos de coincidencias con diferentes escenarios
        matches_data = {
            'dian_folio': ['F001', 'F002', 'F003', 'F004', 'F005', 'F006', 'F007', 'F008', 'F009', 'F010'],
            'dian_fecha': ['01-01-2024', '02-01-2024', '03-01-2024', '04-01-2024', '05-01-2024', 
                          '06-01-2024', '07-01-2024', '08-01-2024', '09-01-2024', '10-01-2024'],
            'dian_valor': [1000.00, 2500.50, 750.25, 3000.00, 1500.75, 2200.00, 1800.50, 1500000.00, 500.00, 3000000.00],
            'dian_descripcion': ['Factura 001', 'Factura 002', 'Factura 003', 'Factura 004', 'Factura 005',
                                'Factura 006', 'Factura 007', 'Factura 008', 'Factura 009', 'Factura 010'],
            'dian_tipo_documento': ['Factura', 'Factura', 'Factura', 'Factura', 'Factura',
                                   'Factura', 'Factura', 'Factura', 'Factura', 'Factura'],
            'contable_numero_documento': ['DOC001', 'DOC002', 'DOC003', 'DOC004', 'DOC005',
                                         'DOC006', 'DOC007', 'DOC008', 'DOC009', 'DOC010'],
            'contable_fecha': ['01-01-2024', '02-01-2024', '03-01-2024', '04-01-2024', '05-01-2024',
                              '06-01-2024', '07-01-2024', '08-01-2024', '09-01-2024', '10-01-2024'],
            'contable_valor': [1000.00, 2500.50, 750.25, 3000.00, 1500.75, 2200.00, 1800.50, 1500000.00, 500.00, 3000000.00],
            'contable_descripcion': ['Pago factura 001', 'Pago factura 002', 'Pago factura 003', 'Pago factura 004', 'Pago factura 005',
                                    'Pago factura 006', 'Pago factura 007', 'Pago factura 008', 'Pago factura 009', 'Pago factura 010'],
            'contable_cuenta': ['130505', '130505', '130505', '130505', '130505',
                               '130505', '130505', '130505', '130505', '130505'],
            'match_type': ['Exacta', 'Exacta', 'Exacta', 'Exacta', 'Exacta',
                          'Exacta', 'Exacta', 'Exacta', 'Exacta', 'Exacta'],
            'confidence': [1.0, 1.0, 1.0, 1.0, 1.0, 0.95, 0.85, 0.65, 0.75, 0.90]
        }
        
        # Agregar algunas diferencias para demostrar formato condicional
        matches_data['contable_valor'][5] = 2250.00  # Diferencia de 50
        matches_data['contable_valor'][6] = 1850.00  # Diferencia de 49.5
        matches_data['contable_valor'][8] = 520.00   # Diferencia de 20
        matches_data['contable_valor'][9] = 3050000.00  # Diferencia de 50000
        
        # Agregar diferencias de fecha
        matches_data['contable_fecha'][6] = '10-01-2024'  # 3 días de diferencia
        matches_data['contable_fecha'][8] = '15-01-2024'  # 6 días de diferencia
        
        matches_df = pd.DataFrame(matches_data)
        
        # Datos de no coincidencias
        non_matches_data = {
            'source': ['DIAN', 'DIAN', 'CONTABLE', 'CONTABLE', 'DIAN', 'CONTABLE', 'DIAN'],
            'dian_folio': ['F011', 'F012', '', '', 'F013', '', 'F014'],
            'dian_fecha': ['11-01-2024', '12-01-2024', '', '', '13-01-2024', '', '14-01-2024'],
            'dian_valor': [2000.00, 3500.00, 0.0, 0.0, 1200.00, 0.0, 2500000.00],
            'dian_descripcion': ['Factura 011', 'Factura 012', '', '', 'Factura 013', '', 'Factura 014'],
            'dian_tipo_documento': ['Factura', 'Factura', '', '', 'Factura', '', 'Factura'],
            'contable_numero_documento': ['', '', 'DOC011', 'DOC012', '', 'DOC013', ''],
            'contable_fecha': ['', '', '11-01-2024', '12-01-2024', '', '13-01-2024', ''],
            'contable_valor': [0.0, 0.0, 1800.00, 4000.00, 0.0, 2800.00, 0.0],
            'contable_descripcion': ['', '', 'Pago factura 011', 'Pago factura 012', '', 'Pago factura 013', ''],
            'contable_cuenta': ['', '', '130505', '130505', '', '130505', '']
        }
        
        non_matches_df = pd.DataFrame(non_matches_data)
        
        print(f"   - Coincidencias: {len(matches_df)} registros (con variaciones)")
        print(f"   - No coincidencias: {len(non_matches_df)} registros")
        
        # 2. Generar DataFrames estructurados
        print("\n2. Generando DataFrames estructurados...")
        
        coincidencias_df = processor.create_coincidencias_dataframe(matches_df)
        no_coincidencias_df = processor.create_no_coincidencias_dataframe(non_matches_df)
        
        print("   ✓ DataFrames estructurados creados")
        
        # 3. Calcular estadísticas
        print("\n3. Calculando estadísticas...")
        
        stats = processor.calculate_statistics(coincidencias_df, no_coincidencias_df)
        
        print("   ✓ Estadísticas calculadas")
        
        # 4. Crear archivo Excel con formato avanzado
        print("\n4. Creando archivo Excel con formato avanzado...")
        
        # Definir ruta de salida
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        
        # Crear archivo Excel
        excel_path = processor.create_excel_file(
            coincidencias_df=coincidencias_df,
            no_coincidencias_df=no_coincidencias_df,
            output_path=output_dir,
            stats=stats
        )
        
        print(f"   ✓ Archivo Excel con formato avanzado creado: {excel_path}")
        
        # 5. Verificar funcionalidades aplicadas
        print("\n5. Verificando funcionalidades aplicadas...")
        
        excel_file = Path(excel_path)
        if excel_file.exists():
            file_size = excel_file.stat().st_size / 1024  # KB
            print(f"   ✓ Archivo existe y tiene {file_size:.1f} KB")
            
            # Verificar que se puede abrir con pandas
            try:
                with pd.ExcelFile(excel_path) as xls:
                    sheets = xls.sheet_names
                    print(f"   ✓ Hojas creadas: {sheets}")
                    
                    # Mostrar información de cada hoja
                    for sheet in sheets:
                        df = pd.read_excel(excel_path, sheet_name=sheet)
                        print(f"      - {sheet}: {len(df)} filas, {len(df.columns)} columnas")
                        
            except Exception as e:
                print(f"   ⚠ Error al verificar archivo: {e}")
        else:
            print("   ❌ El archivo no se creó correctamente")
        
        # 6. Mostrar características del formato avanzado
        print("\n6. Características del formato avanzado aplicado:")
        
        print("   🎨 FORMATO CONDICIONAL:")
        print("      ✓ Diferencias de valor con colores (Verde/Amarillo/Rojo)")
        print("      ✓ Diferencias de fecha con indicadores visuales")
        print("      ✓ Valores altos resaltados en amarillo")
        print("      ✓ Niveles de confianza con colores")
        print("      ✓ Estados de validación con formato específico")
        print("      ✓ Celdas vacías con formato especial")
        
        print("\n   🔍 FILTROS Y ORDENAMIENTO:")
        print("      ✓ Filtros automáticos en todas las columnas")
        print("      ✓ Configuración de tabla dinámica")
        print("      ✓ Ordenamiento por defecto por folio DIAN")
        print("      ✓ Búsqueda rápida en filtros")
        print("      ✓ Estructura de tabla profesional")
        
        print("\n   🧮 HERRAMIENTAS PARA CONTADORES:")
        print("      ✓ Fórmulas de suma automática")
        print("      ✓ Validaciones de datos (valores positivos, fechas)")
        print("      ✓ Alertas para discrepancias significativas")
        print("      ✓ Herramientas de análisis estadístico")
        print("      ✓ Conteo de registros problemáticos")
        print("      ✓ Cálculos de máximo, mínimo, promedio, mediana")
        
        # 7. Mostrar ejemplos de fórmulas aplicadas
        print("\n7. Fórmulas y herramientas aplicadas:")
        
        formulas = [
            "=SUM(C5:C14) - Suma total valores DIAN",
            "=SUM(H5:H14) - Suma total valores contables",
            "=B16-B17 - Diferencia total",
            "=AVERAGE(C5:C14) - Promedio de valores",
            "=COUNTA(A5:A14) - Total de registros",
            "=COUNTIF(K5:K14,\">10\") - Registros con diferencia > 10",
            "=COUNTIF(L5:L14,\">7\") - Registros con diferencia fecha > 7 días",
            "=COUNTIF(O5:O14,\"<0.7\") - Registros con confianza < 0.7",
            "=MAX(C5:C14) - Valor máximo DIAN",
            "=MIN(C5:C14) - Valor mínimo DIAN",
            "=STDEV(C5:C14) - Desviación estándar",
            "=MEDIAN(C5:C14) - Mediana"
        ]
        
        for formula in formulas:
            print(f"      • {formula}")
        
        # 8. Información sobre validaciones
        print("\n8. Validaciones de datos configuradas:")
        validations = [
            "Valores numéricos positivos (>= 0)",
            "Fechas entre 1900-01-01 y 2100-12-31",
            "Mensajes de error personalizados",
            "Títulos de validación informativos"
        ]
        
        for validation in validations:
            print(f"      ✓ {validation}")
        
        print("\n=== EJEMPLO COMPLETADO EXITOSAMENTE ===")
        print(f"Archivo Excel con formato avanzado: {excel_path}")
        print("\n📋 INSTRUCCIONES PARA USAR EL ARCHIVO:")
        print("1. Abrir el archivo en Excel")
        print("2. Usar los filtros en los encabezados para buscar datos")
        print("3. Observar los colores del formato condicional")
        print("4. Revisar las herramientas de contador al final de cada hoja")
        print("5. Utilizar las fórmulas para análisis adicional")
        
    except Exception as e:
        print(f"Error en el ejemplo: {e}")
        import traceback
        traceback.print_exc()

def demonstrate_conditional_formatting():
    """Demostrar tipos de formato condicional aplicados"""
    
    print("\n=== TIPOS DE FORMATO CONDICIONAL ===")
    
    formats = [
        ("Perfect Match", "Verde claro (#c6efce)", "Diferencias ≤ 0.01, fechas iguales"),
        ("Minor Difference", "Amarillo claro (#ffeb9c)", "Diferencias 0.01-10, fechas 1-7 días"),
        ("Major Difference", "Rojo claro (#ffc7ce)", "Diferencias > 10, fechas > 7 días"),
        ("High Value", "Amarillo oscuro (#ffd966)", "Valores > 1,000,000"),
        ("High Confidence", "Verde (#d4edda)", "Confianza ≥ 0.9"),
        ("Medium Confidence", "Amarillo (#fff3cd)", "Confianza 0.7-0.89"),
        ("Low Confidence", "Rojo (#f8d7da)", "Confianza < 0.7"),
        ("DIAN Only", "Azul claro (#e6f3ff)", "Registros solo DIAN"),
        ("Contable Only", "Naranja claro (#fff2e6)", "Registros solo contables"),
        ("Empty Cell", "Gris (#f2f2f2)", "Celdas vacías")
    ]
    
    for name, color, description in formats:
        print(f"   🎨 {name}: {color} - {description}")

def demonstrate_tools():
    """Demostrar herramientas para contadores"""
    
    print("\n=== HERRAMIENTAS PARA CONTADORES ===")
    
    tools = [
        ("📊 Fórmulas de Suma", "Cálculos automáticos de totales y promedios"),
        ("🔍 Alertas de Discrepancias", "Conteo de registros problemáticos"),
        ("📈 Análisis Estadístico", "Máximo, mínimo, desviación estándar, mediana"),
        ("✅ Validaciones de Datos", "Control de entrada de datos válidos"),
        ("🎯 Indicadores de Calidad", "Métricas de confianza y precisión"),
        ("📋 Filtros Avanzados", "Búsqueda y filtrado por múltiples criterios"),
        ("🔄 Ordenamiento Inteligente", "Organización automática de datos"),
        ("📊 Tabla Dinámica", "Configuración para análisis dinámico")
    ]
    
    for tool, description in tools:
        print(f"   {tool}: {description}")

if __name__ == "__main__":
    main()
    demonstrate_conditional_formatting()
    demonstrate_tools() 