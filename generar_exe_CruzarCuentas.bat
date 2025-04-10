@echo off
title Generar .exe de CruzarCuentas

echo Instalando dependencias necesarias...
pip install pyinstaller

echo Generando el ejecutable...
pyinstaller --noconfirm --onefile --windowed --icon=raul_icono.ico principal.py

echo.
echo ✅ El archivo .exe ha sido generado correctamente en la carpeta "dist"
pause
