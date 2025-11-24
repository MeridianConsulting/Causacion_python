# Configuración del Icono de la Aplicación

## Ubicación del archivo de icono

Coloca tu archivo de icono con el nombre `app_icon.ico` en una de estas ubicaciones (en orden de prioridad):

1. `resources/app_icon.ico` (recomendado)
2. `assets/app_icon.ico`
3. `icon.ico` (en la raíz del proyecto)
4. `app_icon.ico` (en la raíz del proyecto)

## Formato del icono

- **Formato recomendado**: `.ico` (formato nativo de Windows)
- **Tamaños recomendados**: 
  - 16x16 píxeles (icono pequeño en barra de tareas)
  - 32x32 píxeles (icono mediano)
  - 48x48 píxeles (icono grande)
  - 256x256 píxeles (alta resolución)

Un archivo `.ico` puede contener múltiples tamaños en un solo archivo.

## Cómo crear un archivo .ico

### Opción 1: Usar herramientas online
- [ICO Convert](https://icoconvert.com/)
- [ConvertICO](https://convertico.com/)
- [Favicon Generator](https://www.favicon-generator.org/)

### Opción 2: Usar herramientas de diseño
- **GIMP** (gratis): Exportar como `.ico`
- **Photoshop**: Usar plugin para exportar `.ico`
- **Paint.NET**: Con plugin ICO
- **Inkscape**: Exportar SVG a PNG y luego convertir a ICO

### Opción 3: Convertir desde PNG
Si tienes un archivo PNG, puedes convertirlo a ICO usando herramientas online o:
- Python con `Pillow`: `pip install Pillow` y usar scripts de conversión
- Herramientas de línea de comandos

## Pasos para agregar tu icono

1. Crea o descarga tu icono en formato `.ico`
2. Crea la carpeta `resources` en la raíz del proyecto (si no existe)
3. Coloca el archivo como `resources/app_icon.ico`
4. Ejecuta la aplicación - el icono aparecerá automáticamente

## Verificación

Una vez agregado el icono, deberías ver:
- El icono personalizado en la barra de tareas de Windows
- El icono en la barra de título de la ventana
- El icono en el administrador de tareas (Ctrl+Shift+Esc)

Si no ves el icono, verifica:
- Que el archivo existe en una de las rutas mencionadas
- Que el archivo tiene extensión `.ico`
- Que el archivo no está corrupto
- Reinicia la aplicación después de agregar el icono

