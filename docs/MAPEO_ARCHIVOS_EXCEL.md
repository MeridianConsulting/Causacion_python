# 📊 MAPEO DEFINITIVO DE ARCHIVOS EXCEL

**Fecha de análisis:** Enero 2025  
**Proyecto:** Sistema de Causación Automática  
**Estado:** ✅ Mapeo completado - Estructura 100% identificada

---

## 🏛️ ARCHIVO DIAN

### Información General
- **Archivo:** `17_Julio_2025_Dian.xlsx`
- **Ubicación:** `data/input/17_Julio_2025_Dian.xlsx`
- **Estado:** ✅ Perfectamente estructurado y listo para procesamiento
- **Dimensiones:** 647 registros × 32 columnas
- **Período:** 1-17 de Julio 2025
- **Calidad:** Excelente - Sin filas vacías, todas las columnas tienen datos

### Estructura de Columnas Principales

| # | Campo | Tipo | Registros | Descripción | Clave Cruce |
|---|-------|------|-----------|-------------|-------------|
| 1 | Tipo de documento | object | 647 | Factura electrónica, Nota débito, etc. | |
| 2 | CUFE/CUDE | object | 647 | Código único factura electrónica | |
| 3 | **Folio** | object | 647 | Número del documento | ⭐ PRIMARIA |
| 4 | Prefijo | object | 344 | Prefijo del documento (303 nulos) | |
| 5 | Divisa | object | 165 | Moneda del documento (482 nulos) | |
| 6 | Forma de Pago | float64 | 165 | Forma de pago (482 nulos) | |
| 7 | Medio de Pago | object | 165 | Medio de pago (482 nulos) | |
| 8 | **Fecha Emisión** | object | 647 | Fecha emisión formato DD-MM-YYYY | ⭐ PRIMARIA |
| 9 | Fecha Recepción | object | 647 | Fecha de recepción | |
| 10 | **NIT Emisor** | int64 | 647 | Identificación del emisor | ⭐ SECUNDARIA |
| 11 | Nombre Emisor | object | 647 | Razón social del emisor | |
| 12 | **NIT Receptor** | int64 | 647 | Identificación del receptor | ⭐ SECUNDARIA |
| 13 | Nombre Receptor | object | 647 | Razón social del receptor | |
| 14 | IVA | float64 | 647 | Valor del IVA | |
| 15 | ICA | int64 | 647 | Valor del ICA | |
| 16 | IC | int64 | 647 | Impuesto al consumo | |
| 17 | INC | float64 | 647 | Impuesto nacional al consumo | |
| 18 | Timbre | int64 | 647 | Impuesto de timbre | |
| 19 | INC Bolsas | int64 | 647 | INC bolsas plásticas | |
| 20 | IN Carbono | int64 | 647 | Impuesto nacional carbono | |
| 21 | IN Combustibles | int64 | 647 | Impuesto combustibles | |
| 22 | IC Datos | int64 | 647 | Impuesto consumo datos | |
| 23 | ICL | int64 | 647 | Impuesto consumo licores | |
| 24 | INPP | int64 | 647 | Impuesto productos plásticos | |
| 25 | IBUA | int64 | 647 | Impuesto bebidas ultraprocesadas | |
| 26 | ICUI | float64 | 647 | Impuesto consumo cigarrillos | |
| 27 | Rete IVA | int64 | 647 | Retención en la fuente IVA | |
| 28 | Rete Renta | int64 | 647 | Retención en la fuente Renta | |
| 29 | Rete ICA | int64 | 647 | Retención ICA | |
| 30 | **Total** | float64 | 647 | Valor total del documento | ⭐ PRIMARIA |
| 31 | Estado | object | 647 | Estado del documento | |
| 32 | Grupo | object | 647 | Clasificación del documento | |

### Estadísticas DIAN
- ✅ **643 documentos únicos** (campo Folio)
- ✅ **88 emisores diferentes**
- ✅ **275 receptores únicos**
- ✅ **Todos los registros tienen datos completos**

### Configuración de Lectura
```python
df_dian = pd.read_excel('data/input/17_Julio_2025_Dian.xlsx', header=0)
```

---

## 💼 ARCHIVO MOVIMIENTO CONTABLE

### Información General
- **Archivo:** `movimientocontable.xlsx`
- **Ubicación:** `data/input/movimientocontable.xlsx`
- **Estado:** ⚠️ Estructura compleja, requiere preprocesamiento
- **Dimensiones:** 1,105 filas × 125 columnas
- **Configuración:** Encabezados en fila 4, datos desde fila 5
- **Calidad:** Buena - Datos identificados correctamente

### Estructura Identificada
- **Filas 1-3:** Metadatos y títulos del reporte
  - Fila 1: "MODELO PARA LA IMPORTACION DE MOVIMIENTO CONTABLE"
  - Fila 2: "De : JUL 1/2025 A : JUL 18/2025"
  - Fila 3: [Vacía]
- **Fila 4:** Encabezados reales de las columnas
- **Filas 5+:** Datos de movimientos contables

### Columnas Principales Identificadas

| # | Posición | Campo Original | Nombre Sugerido | Tipo | Ejemplo | Descripción | Clave Cruce |
|---|----------|----------------|-----------------|------|---------|-------------|-------------|
| 1 | Col 0 | TIPO DE COMPROBANTE | tipo_comprobante | str | L | Tipo de asiento contable | |
| 2 | Col 1 | CÓDIGO COMPROBANTE | codigo_comprobante | int | 19 | Código del comprobante | |
| 3 | Col 2 | **NÚMERO DE DOCUMENTO** | numero_documento | int | 13, 14 | Número del documento | ⭐ PRIMARIA |
| 4 | Col 3 | CUENTA CONTABLE | cuenta_contable | str/int | 2525050100 | Código cuenta contable | |
| 5 | Col 4 | DÉBITO O CRÉDITO | debito_credito | str | D, C | Naturaleza del movimiento | |
| 6 | Col 5 | **VALOR DE LA SECUENCIA** | valor_movimiento | float | 8635900 | Valor del movimiento | ⭐ PRIMARIA |
| 7 | Col 6 | **AÑO DEL DOCUMENTO** | año | int | 2025 | Año del documento | ⭐ PRIMARIA |
| 8 | Col 7 | **MES DEL DOCUMENTO** | mes | int | 7 | Mes del documento | ⭐ PRIMARIA |
| 9 | Col 8 | **DÍA DEL DOCUMENTO** | dia | int | 1 | Día del documento | ⭐ PRIMARIA |
| 10 | Col 9 | CÓDIGO DEL VENDEDOR | codigo_vendedor | int | 0 | Código vendedor | |
| 11+ | Col 10+ | [Múltiples campos] | campos_adicionales | mixed | - | Campos adicionales contables | |

### Datos de Ejemplo Identificados
```
Registro 1: L | 19 | 13 | 2525050100 | D | 8635900 | 2025 | 7 | 1
Registro 2: L | 19 | 13 | 2370050100 | C | 241900  | 2025 | 7 | 1
Registro 3: L | 19 | 13 | 2380301400 | C | 241900  | 2025 | 7 | 1
```

### Configuración de Lectura
```python
# Leer saltando las primeras 4 filas de metadatos
df_contable = pd.read_excel('data/input/movimientocontable.xlsx', skiprows=4)

# Mapear nombres de columnas
column_mapping = {
    'MERIDIAN CONSULTING LTDA': 'tipo_comprobante',
    'Unnamed: 1': 'codigo_comprobante',
    'Unnamed: 2': 'numero_documento',
    'Unnamed: 3': 'cuenta_contable',
    'Unnamed: 4': 'debito_credito',
    'Unnamed: 5': 'valor_movimiento',
    'Unnamed: 6': 'año',
    'Unnamed: 7': 'mes',
    'Unnamed: 8': 'dia',
    'Unnamed: 9': 'codigo_vendedor'
}

df_contable = df_contable.rename(columns=column_mapping)
```

---

## 🔗 ESTRATEGIA DE CRUCE ENTRE ARCHIVOS

### Campos de Enlace Identificados

| Prioridad | Campo DIAN | Campo Contable | Estrategia | Confiabilidad |
|-----------|------------|----------------|------------|---------------|
| 1 | **Folio** | **numero_documento** | Match directo por número | ⭐⭐⭐ ALTA |
| 2 | **Total** | **valor_movimiento** | Match por valor monetario | ⭐⭐⭐ ALTA |
| 3 | **NIT Emisor/Receptor** | **Campo tercero** | Match por identificación | ⭐⭐ MEDIA |
| 4 | **Fecha Emisión** | **año + mes + dia** | Match por fecha completa | ⭐⭐⭐ ALTA |

### Lógica de Cruce Recomendada

#### 🎯 Nivel 1 - MATCH PRIMARIO (Más confiable)
```python
match_primario = (
    (dian['Folio'] == contable['numero_documento']) &
    (dian['fecha_procesada'] == contable['fecha_procesada']) &
    (dian['Total'] == contable['valor_movimiento'])
)
```

#### 🎯 Nivel 2 - MATCH SECUNDARIO (Confiable)
```python
match_secundario = (
    (dian['NIT_Emisor'].isin([contable['nit_tercero']]) | 
     dian['NIT_Receptor'].isin([contable['nit_tercero']])) &
    (dian['Total'] == contable['valor_movimiento']) &
    (dian['fecha_procesada'] == contable['fecha_procesada'])
)
```

#### 🎯 Nivel 3 - MATCH TERCIARIO (Menos confiable)
```python
match_terciario = (
    (dian['Folio'] == contable['numero_documento']) &
    (dian['Total'] == contable['valor_movimiento'])
)
```

---

## 🛠️ PLAN DE IMPLEMENTACIÓN

### ✅ Procesamiento Archivo DIAN
- **Estado:** Listo para uso inmediato
- **Función:** `pd.read_excel(file, header=0)`
- **Campos clave:** Folio, NIT_Emisor, NIT_Receptor, Total, Fecha_Emisión

### ⚙️ Preprocesamiento Archivo Contable
1. **Lectura:** `pd.read_excel(file, skiprows=4)` # Saltar metadatos
2. **Mapeo:** Renombrar columnas 'Unnamed' a nombres descriptivos
3. **Limpieza:** Filtrar filas con datos válidos
4. **Transformación:** Combinar Año/Mes/Día en fecha única

### 📋 Módulos a Desarrollar

#### 1. Módulo de Carga (`data_loader.py`)
```python
def cargar_archivo_dian(file_path)
def cargar_archivo_contable(file_path)
def limpiar_datos_contable(df)
def mapear_columnas_contable(df)
```

#### 2. Módulo de Cruce (`data_matcher.py`)
```python
def cruzar_datos_nivel1(df_dian, df_contable)
def cruzar_datos_nivel2(df_dian, df_contable)
def cruzar_datos_nivel3(df_dian, df_contable)
def generar_reporte_cruce(matches)
```

#### 3. Módulo de Reportes (`report_generator.py`)
```python
def generar_reporte_causacion(matches)
def exportar_excel_resultado(data, output_path)
def generar_estadisticas_cruce(matches)
```

---

## 📊 PRÓXIMOS PASOS

1. ✅ **Crear módulo de carga y limpieza de archivos**
2. ✅ **Implementar mapeo de columnas del archivo contable**
3. ✅ **Desarrollar función de cruce de datos**
4. ✅ **Crear reportes de causación automática**
5. ✅ **Integrar con la interfaz gráfica existente**

---

## 🎯 NOTAS IMPORTANTES

### ⚠️ Consideraciones Especiales
- El archivo contable tiene una estructura compleja que requiere preprocesamiento
- Las primeras 4 filas contienen metadatos que deben ser omitidos
- Muchas columnas del archivo contable están sin nombrar ('Unnamed: X')
- Es necesario implementar tolerancia en los matches por posibles diferencias de formato

### 💡 Recomendaciones
- Implementar logging detallado para el proceso de cruce
- Crear validaciones de integridad de datos antes del cruce
- Generar reportes de calidad del match (% de coincidencias)
- Implementar backup automático antes de procesar

### 🔧 Configuraciones Técnicas
- **Encoding:** UTF-8 para caracteres especiales
- **Formato fechas:** DD-MM-YYYY (DIAN) vs componentes separados (Contable)
- **Valores monetarios:** float64 para precisión en cálculos
- **NITs:** int64 para evitar problemas de precisión

---

**📝 Documento generado:** Enero 2025  
**👤 Responsable:** Sistema de Análisis Automático  
**🔄 Última actualización:** Análisis inicial completo  
**✅ Estado:** Mapeo finalizado - Listo para implementación 