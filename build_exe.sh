#!/bin/bash
# Script para generar el ejecutable con icono personalizado (Linux/Mac)
# Asegúrate de tener PyInstaller instalado: pip install pyinstaller

echo "========================================"
echo "Generando ejecutable con icono personalizado"
echo "========================================"
echo ""

# Verificar que existe el icono
if [ ! -f "resources/app_icon.ico" ]; then
    echo "[ERROR] No se encuentra el archivo resources/app_icon.ico"
    echo "Por favor, coloca tu icono en resources/app_icon.ico"
    exit 1
fi

echo "[OK] Icono encontrado: resources/app_icon.ico"
echo ""

# Generar el ejecutable
pyinstaller \
    --name CausacionDIAN \
    --windowed \
    --icon resources/app_icon.ico \
    --onefile \
    --clean \
    --noconfirm \
    --add-data "resources:resources" \
    main.py

if [ $? -eq 0 ]; then
    echo ""
    echo "========================================"
    echo "[OK] Ejecutable generado exitosamente"
    echo "========================================"
    echo ""
    echo "El ejecutable se encuentra en: dist/CausacionDIAN"
    echo ""
else
    echo ""
    echo "[ERROR] Hubo un error al generar el ejecutable"
    echo ""
fi

