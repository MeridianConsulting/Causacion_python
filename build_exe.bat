@echo off
REM Script para generar el ejecutable con icono personalizado
REM Asegúrate de tener PyInstaller instalado: pip install pyinstaller

echo ========================================
echo Generando ejecutable con icono personalizado
echo ========================================
echo.

REM Verificar que existe el icono
if not exist "resources\app_icon.ico" (
    echo [ERROR] No se encuentra el archivo resources\app_icon.ico
    echo Por favor, coloca tu icono en resources\app_icon.ico
    pause
    exit /b 1
)

echo [OK] Icono encontrado: resources\app_icon.ico
echo.

REM Generar el ejecutable
REM Usar python -m pyinstaller para evitar problemas con PATH
python -m PyInstaller ^
    --name CausacionDIAN ^
    --windowed ^
    --icon resources\app_icon.ico ^
    --onefile ^
    --clean ^
    --noconfirm ^
    --add-data "resources;resources" ^
    main.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo [OK] Ejecutable generado exitosamente
    echo ========================================
    echo.
    echo El ejecutable se encuentra en: dist\CausacionDIAN.exe
    echo.
    echo Ahora puedes ejecutar dist\CausacionDIAN.exe y veras tu icono
    echo personalizado en la barra de tareas de Windows.
    echo.
) else (
    echo.
    echo [ERROR] Hubo un error al generar el ejecutable
    echo.
)

pause

