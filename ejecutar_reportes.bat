@echo off
cd /d "%~dp0"

:: 1. Ejecutar tus scripts de Python
py actualizar_reportes.py
py consolidar.py

pause