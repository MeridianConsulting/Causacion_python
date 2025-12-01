# Lista de Validaciones del Script de Causación

Este documento lista las principales validaciones que realiza el script `causacion_processor.py`.

## 1. Validaciones de Archivos

### 1.1 Validación de Existencia
- ✅ Verifica que el archivo DIAN existe antes de cargarlo
- ✅ Verifica que el archivo contable existe antes de cargarlo
- ✅ Lanza `FileNotFoundError` si el archivo no existe

### 1.2 Validación de Extensión
- ✅ Verifica que los archivos sean de tipo Excel (`.xlsx` o `.xls`)
- ✅ Lanza `ValueError` si la extensión no es válida

### 1.3 Validación de Contenido
- ✅ Verifica que el archivo DIAN no esté vacío después de cargar
- ✅ Verifica que el archivo contable no esté vacío después de cargar
- ✅ Lanza `ValueError` si el DataFrame está vacío

### 1.4 Validación de Archivos Cargados
- ✅ Verifica que ambos archivos (DIAN y contable) estén cargados antes de procesar
- ✅ Verifica que ninguno de los DataFrames esté vacío

## 2. Validaciones de Calidad de Datos

### 2.1 Validación de Integridad de Datos
- ✅ Calcula el porcentaje de valores faltantes por columna
- ✅ Identifica columnas con más del 50% de valores faltantes como críticas
- ✅ Genera reporte de valores faltantes con conteo y porcentaje

### 2.2 Validación de Campos Críticos
- ✅ Identifica campos críticos basándose en palabras clave: `folio`, `numero`, `identificacion`, `nit`, `ruc`, `fecha`, `valor`, `monto`
- ✅ Verifica que los campos críticos no estén vacíos
- ✅ Reporta campos críticos con valores faltantes

### 2.3 Validación de Formatos de Fecha
- ✅ Identifica columnas de fecha por palabras clave: `fecha`, `date`, `dia`, `mes`, `año`
- ✅ Valida formato de fecha `DD-MM-YYYY`
- ✅ Valida formato de fecha con hora `DD-MM-YYYY HH:MM:SS`
- ✅ Valida fechas usando inferencia automática con `dayfirst=True`
- ✅ Reporta errores de formato de fecha por columna

### 2.4 Validación de Formatos Numéricos
- ✅ Detecta columnas numéricas automáticamente
- ✅ Verifica valores infinitos (`inf` o `-inf`)
- ✅ Verifica valores extremadamente grandes (mayores a 1 billón - 1e12)
- ✅ Reporta problemas de formato numérico

### 2.5 Cálculo de Score de Calidad
- ✅ Calcula un score general de calidad (0-100)
- ✅ Considera problemas de campos críticos, fechas y números
- ✅ Marca como válido si el score es >= 70
- ✅ Genera reporte detallado de calidad

## 3. Validaciones de Datos DIAN

### 3.1 Validación de Folio
- ✅ Verifica que el folio no esté vacío o sea 'nan'
- ✅ Valida formato del folio

### 3.2 Validación de Valores DIAN
- ✅ Verifica que los valores no sean extremadamente altos (> 1 billón)
- ✅ Verifica que los valores no sean negativos
- ✅ Identifica posibles errores de captura

### 3.3 Validación de Fechas DIAN
- ✅ Verifica que las fechas no estén vacías o sean 'nan'
- ✅ Valida formato de fecha DIAN

## 4. Validaciones de Datos Contables

### 4.1 Validación de Número de Documento
- ✅ Verifica que el número de documento no esté vacío o sea 'nan'
- ✅ Valida formato del número de documento

### 4.2 Validación de Valores Contables
- ✅ Verifica que los valores no sean extremadamente altos (> 1 billón)
- ✅ Verifica que los valores no sean negativos
- ✅ Identifica posibles errores de captura

### 4.3 Validación de Fechas Contables
- ✅ Verifica que las fechas no estén vacías o sean 'nan'
- ✅ Valida formato de fecha contable

## 5. Validaciones de Coincidencias

### 5.1 Validación de Estado de Coincidencia
- ✅ **Perfecta**: Diferencia de valor ≤ 0.01 y diferencia de fecha = 0 días
- ✅ **Buena**: Diferencia de valor ≤ 1.0 y diferencia de fecha ≤ 1 día
- ✅ **Regular**: Diferencia de valor ≤ 10.0 y diferencia de fecha ≤ 7 días
- ✅ **Revisar**: Cualquier otra combinación de diferencias

### 5.2 Validación de Coincidencia de Valores
- ✅ Verifica coincidencia exacta de valores (con tolerancia de redondeo)
- ✅ Verifica coincidencia aproximada de valores (dentro de tolerancia)
- ✅ Calcula diferencias de valor entre registros DIAN y contables

### 5.3 Validación de Coincidencia de Fechas
- ✅ Verifica coincidencia exacta de fechas
- ✅ Verifica coincidencia aproximada de fechas (dentro de tolerancia de días)
- ✅ Calcula diferencias de fecha entre registros DIAN y contables

### 5.4 Validación de Coincidencia de Documentos
- ✅ Verifica coincidencia de números de documento
- ✅ Verifica coincidencia de NIT/identificación
- ✅ Verifica coincidencia de folios

## 6. Validaciones de No Coincidencias

### 6.1 Análisis de Motivos DIAN sin Contraparte
- ✅ Folio DIAN vacío o inválido
- ✅ Valor DIAN extremadamente alto (posible error)
- ✅ Valor DIAN negativo
- ✅ Fecha DIAN vacía o inválida
- ✅ Registro DIAN sin contraparte contable

### 6.2 Análisis de Motivos Contable sin Contraparte
- ✅ Número de documento contable vacío o inválido
- ✅ Valor contable extremadamente alto (posible error)
- ✅ Valor contable negativo
- ✅ Fecha contable vacía o inválida
- ✅ Registro contable sin contraparte DIAN

## 7. Validaciones de Excel (Validaciones de Datos en Excel)

### 7.1 Validación de Valores Numéricos
- ✅ Valida que los valores sean números decimales
- ✅ Valida que los valores sean mayores o iguales a 0 (positivos)
- ✅ Muestra mensaje de error si el valor no es válido
- ✅ Aplica validación en columnas de valores del archivo de salida

### 7.2 Validación de Fechas en Excel
- ✅ Valida que las fechas estén en formato de fecha válido
- ✅ Valida que las fechas estén entre 1900-01-01 y 2100-12-31
- ✅ Muestra mensaje de error si la fecha no es válida
- ✅ Aplica validación en columnas de fechas del archivo de salida

## 8. Validaciones de Estructura de Datos

### 8.1 Validación de DataFrames
- ✅ Verifica que los DataFrames no estén vacíos antes del cruce
- ✅ Verifica que existan columnas necesarias para el procesamiento
- ✅ Valida estructura de datos antes de operaciones críticas

### 8.2 Validación de Columnas
- ✅ Identifica columnas críticas automáticamente
- ✅ Mapea columnas sin nombre basándose en el contenido
- ✅ Verifica existencia de columnas antes de acceder a ellas

## 9. Validaciones de Calidad General del Proceso

### 9.1 Cálculo de Calidad General
- ✅ Calcula porcentaje de coincidencias (40% del score)
- ✅ Calcula calidad de coincidencias basada en porcentaje de "Perfectas" (30% del score)
- ✅ Evalúa completitud de datos (30% del score)
- ✅ Genera calificación: 'Excelente', 'Buena', 'Regular', 'Mala'

### 9.2 Validación de Estadísticas
- ✅ Verifica que existan coincidencias antes de calcular estadísticas
- ✅ Verifica que existan no coincidencias antes de calcular estadísticas
- ✅ Valida cálculos de porcentajes y totales

## 10. Validaciones de Limpieza de Datos

### 10.1 Validación de Metadatos
- ✅ Identifica y elimina filas de metadatos al inicio de los archivos
- ✅ Verifica si la primera fila contiene metadatos o encabezados
- ✅ Limpia filas vacías completamente

### 10.2 Validación de Espacios y Formato
- ✅ Elimina espacios en blanco al inicio y final de valores de texto
- ✅ Normaliza valores 'nan', 'None', '' a valores nulos apropiados
- ✅ Limpia caracteres especiales problemáticos

## Resumen de Métodos de Validación

| Método | Propósito |
|--------|-----------|
| `validate_files()` | Valida que los archivos existan y estén cargados |
| `validate_data_quality()` | Valida calidad general de datos del DataFrame |
| `_validate_critical_fields()` | Valida que campos críticos no estén vacíos |
| `_validate_date_format()` | Valida formato de fechas en columnas |
| `_validate_numeric_format()` | Valida formato numérico en columnas |
| `_evaluate_match_quality()` | Evalúa calidad de una coincidencia |
| `_analyze_dian_non_match_reason()` | Analiza motivo de no coincidencia DIAN |
| `_analyze_contable_non_match_reason()` | Analiza motivo de no coincidencia contable |
| `_calculate_overall_quality()` | Calcula calidad general del proceso |
| `_add_data_validations()` | Agrega validaciones de datos en Excel de salida |

## Notas Importantes

- Todas las validaciones generan logs informativos, de advertencia o de error
- Las validaciones críticas lanzan excepciones que detienen el proceso
- Las validaciones de calidad generan scores y reportes detallados
- Las validaciones en Excel se aplican al archivo de salida para prevenir errores de entrada
- El sistema es tolerante a fallos: algunas validaciones solo generan advertencias sin detener el proceso

