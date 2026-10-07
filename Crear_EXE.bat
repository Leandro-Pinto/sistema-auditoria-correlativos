@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Creando SistemaAuditoria.exe (tarda 1 o 2 minutos)...
python -m pip install --user -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name SistemaAuditoria --hidden-import reporte_pdf --hidden-import reportlab.graphics.barcode.common --collect-submodules reportlab.graphics.barcode SistemaAuditoria.py
if exist dist\SistemaAuditoria.exe (
  echo.
  echo Listo: dist\SistemaAuditoria.exe
  echo Puede copiar ese .exe a cualquier PC con Windows; no necesita Python.
) else (
  echo No se pudo crear el ejecutable. Revise los mensajes de arriba.
)
pause
