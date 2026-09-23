@echo off
cd /d "%~dp0"

:: 1. Ejecutar tus scripts de Python
py actualizar_reportes.py
py consolidar.py

:: 2. Agregar los archivos modificados a Git
git add .

:: 3. Crear el commit con la fecha y hora actual
git commit -m "Actualización automática de reportes - %date% %time%"

:: 4. Subir los cambios a GitHub
git push origin main

pause