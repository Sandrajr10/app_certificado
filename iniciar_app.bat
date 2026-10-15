@echo off
REM Arranca la app de Academia Horizonte con doble clic.
REM Usa la carpeta donde está este archivo, aunque la ruta tenga espacios.
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python. Instala Python 3 desde python.org y vuelve a intentarlo.
  pause
  exit /b 1
)
python -m pip install -r requirements.txt
python app.py
pause
