@echo off
setlocal
cd /d "%~dp0\.."
REM 1) Build frontend
cd frontend
call npm run build || goto :eof
cd ..

REM 2) Crear venv (si hace falta) e instalar reqs
if not exist .venv (
  py -3 -m venv .venv
)
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r backend\requirements.txt
pip install pyinstaller

REM 3) PyInstaller
pyinstaller backend\CobranzaTray.spec

echo.
echo Build listo en: dist\CobranzaTray\CobranzaTray.exe
endlocal