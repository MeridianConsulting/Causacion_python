# Instrucciones para mostrar el icono en la barra de tareas de Windows

## Problema

El icono personalizado aparece en la ventana (barra de título y Alt+Tab), pero en la **barra de tareas de Windows** sigue apareciendo el icono de Python.

## Explicación

- **Icono en ventana/Alt+Tab**: Lo controla Qt con `setWindowIcon()` ✅ (ya funciona)
- **Icono en barra de tareas**: Lo controla Windows usando el icono del **ejecutable** que se está ejecutando

Si ejecutas `python.exe main.py`, Windows ve que el programa es `python.exe`, así que muestra el logo de Python en la barra de tareas.

## Solución: Generar un .exe con tu icono

### Opción 1: Usar el script automático (Recomendado)

#### En Windows:

1. Abre una terminal en la carpeta del proyecto
2. Ejecuta:
   ```bash
   build_exe.bat
   ```

3. Espera a que termine (puede tardar unos minutos)
4. El ejecutable estará en: `dist\CausacionDIAN.exe`
5. Ejecuta `dist\CausacionDIAN.exe` y verás tu icono en la barra de tareas

#### En Linux/Mac:

1. Abre una terminal en la carpeta del proyecto
2. Ejecuta:
   ```bash
   chmod +x build_exe.sh
   ./build_exe.sh
   ```

### Opción 2: Comando manual de PyInstaller

Si prefieres ejecutar el comando manualmente:

```bash
pyinstaller --name CausacionDIAN --windowed --icon resources\app_icon.ico --onefile --clean main.py
```

**Parámetros explicados:**
- `--name CausacionDIAN`: Nombre del ejecutable
- `--windowed`: No mostrar consola (solo ventana gráfica)
- `--icon resources\app_icon.ico`: Usar tu icono personalizado
- `--onefile`: Generar un solo archivo .exe
- `--clean`: Limpiar archivos temporales antes de construir
- `main.py`: Archivo principal de entrada

### Opción 3: Usar un archivo .spec (Para configuraciones avanzadas)

Si necesitas más control, puedes crear un archivo `CausacionDIAN.spec`:

```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('resources', 'resources')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CausacionDIAN',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/app_icon.ico',
)
```

Luego ejecuta:
```bash
pyinstaller CausacionDIAN.spec
```

## Verificación

Después de generar el .exe:

1. **Ejecuta** `dist\CausacionDIAN.exe`
2. **Verifica** que el icono aparece en:
   - ✅ Barra de tareas de Windows
   - ✅ Barra de título de la ventana
   - ✅ Alt+Tab (selector de ventanas)
   - ✅ Administrador de tareas

3. **Ancla a la barra de tareas**:
   - Clic derecho en el icono de la barra de tareas
   - Selecciona "Anclar a la barra de tareas"
   - A partir de ahí siempre usará tu icono

## Requisitos previos

Asegúrate de tener PyInstaller instalado:

```bash
pip install pyinstaller
```

## Solución alternativa: Acceso directo (Menos recomendado)

Si no quieres generar un .exe todavía:

1. Crea un **acceso directo** en el escritorio
2. Destino: `"C:\ruta\a\pythonw.exe" "C:\ruta\a\main.py"`
3. Clic derecho → **Propiedades** → **Cambiar icono...**
4. Selecciona `resources\app_icon.ico`
5. Ancla ese acceso directo a la barra de tareas

**Nota**: Esta solución es menos fiable porque Windows a veces sigue mostrando el icono del ejecutable real (pythonw.exe).

## Resumen

- ✅ Tu código de iconos en Qt está correcto
- ✅ El icono aparece en la ventana y Alt+Tab
- ⚠️ Para la barra de tareas necesitas un .exe propio con el icono
- 🎯 Usa `build_exe.bat` para generar el ejecutable con tu icono

