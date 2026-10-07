@echo off
chcp 65001 >nul
title Sistema de Auditoria de Correlativos
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python. Instalelo desde https://www.python.org/downloads/
  echo Marque la opcion "Add Python to PATH" durante la instalacion.
  pause
  exit /b 1
)
python -c "import openpyxl, reportlab" 2>nul || python -m pip install --user -r requirements.txt
start "" pythonw SistemaAuditoria.py %*
