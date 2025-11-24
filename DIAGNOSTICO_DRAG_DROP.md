# Diagnóstico de Drag & Drop

## Problema: Icono de prohibido al arrastrar archivos

Si al intentar arrastrar archivos Excel a la aplicación aparece un icono de prohibido (🚫) y no se pueden soltar los archivos, sigue estos pasos para diagnosticar el problema.

## Paso 1: Verificar si los eventos se están recibiendo

1. Ejecuta la aplicación principal
2. Abre la consola/terminal donde se ejecuta la aplicación
3. Intenta arrastrar un archivo Excel sobre el área de drop
4. **Observa la consola:**

   - ✅ **Si ves mensajes `[DEBUG] DRAG ENTER`**: Los eventos SÍ están llegando, el problema está en el manejo del código
   - ❌ **Si NO ves ningún mensaje**: Los eventos NO están llegando, el problema es de permisos/entorno de Windows

## Paso 2: Ejecutar el test mínimo

Se ha creado un archivo `test_dnd.py` para verificar si el drag & drop funciona en tu entorno.

### Cómo ejecutarlo:

```bash
python test_dnd.py
```

**IMPORTANTE:** Ejecuta esto **SIN permisos de administrador**.

### Qué esperar:

- ✅ **Si el test funciona**: Verás mensajes `[TEST]` en la consola y podrás soltar archivos. Esto significa que el problema está en la aplicación principal.
- ❌ **Si el test NO funciona**: No verás ningún mensaje. Esto confirma que el problema es de permisos/entorno de Windows.

## Paso 3: Verificar permisos de administrador

El problema más común es que la aplicación se está ejecutando con permisos de administrador mientras el Explorador de Windows no, lo que bloquea el drag & drop por seguridad.

### Cómo verificar y solucionar:

#### Si ejecutas como script Python:

1. **Cierra cualquier terminal/IDE que esté ejecutándose como administrador**
2. Abre un **cmd o PowerShell normal** (sin "Ejecutar como administrador")
3. Navega a la carpeta del proyecto
4. Ejecuta: `python main.py` o `python -m excel_automation.ui_main`

#### Si usas un IDE (VSCode, PyCharm, etc.):

1. **Cierra completamente el IDE**
2. **Abre el IDE normalmente** (sin "Ejecutar como administrador")
3. Ejecuta la aplicación desde el IDE

#### Si usas un .exe compilado:

1. Clic derecho en el `.exe` → **Propiedades**
2. Pestaña **Compatibilidad**
3. **Asegúrate de que NO esté marcada** la casilla "Ejecutar este programa como administrador"
4. Aplica los cambios y ejecuta de nuevo

#### Si compilaste con PyInstaller:

Verifica que NO hayas usado:
- `--uac-admin`
- Un manifest con `requireAdministrator`

## Paso 4: Verificar que estás soltando en el área correcta

Asegúrate de soltar el archivo **dentro del recuadro rayado** (DropArea), no sobre:
- El título de la tarjeta
- La descripción
- Fuera del área de drop

## Resumen de diagnóstico

| Situación | Causa | Solución |
|-----------|-------|----------|
| No ves `[DEBUG]` en consola | Permisos de admin | Ejecutar sin admin |
| Ves `[DEBUG]` pero no acepta | Problema en código | Revisar lógica de mimeData |
| Test funciona pero app no | Widgets interfiriendo | Verificar overlays/widgets padre |
| Test NO funciona | Entorno Windows | Verificar permisos y nivel de integridad |

## Mensajes de debug esperados

Cuando el drag & drop funciona correctamente, deberías ver en la consola:

```
============================================================
[DEBUG] DRAG ENTER - Evento recibido
[DEBUG] mime formats: ['text/uri-list', 'text/plain']
[DEBUG] hasUrls: True
[DEBUG] URLs encontradas: 1
[DEBUG]   URL 1: file:///C:/ruta/al/archivo.xlsx
[DEBUG]   Local file: C:\ruta\al\archivo.xlsx
[DEBUG] Archivo detectado en dragEnter: C:\ruta\al\archivo.xlsx
[OK] ARCHIVO EXCEL VÁLIDO - Aceptando drag
============================================================
[DEBUG] DROP EVENT - Evento recibido
[DEBUG] Archivo para procesar en drop: C:\ruta\al\archivo.xlsx
[OK] PROCESANDO ARCHIVO EXCEL
[OK] DROP COMPLETADO EXITOSAMENTE
============================================================
```

Si NO ves estos mensajes, el problema es de permisos/entorno, no del código.

