@echo off
setlocal
cd /d "%~dp0\.."

REM 1) Build frontend
cd frontend
call npm run build || goto :fatal
cd ..

REM 2) Crear venv (si hace falta) e instalar reqs
if not exist .venv (
  py -3 -m venv .venv || goto :fatal
)
call .venv\Scripts\activate.bat || goto :fatal
python -m pip install --upgrade pip wheel "setuptools<81" || goto :fatal
pip install -r requirements.txt || goto :fatal
pip install pyinstaller || goto :fatal

REM 3) PyInstaller
pyinstaller backend\CobranzaTray.spec --clean --noconfirm || goto :fatal

echo.
echo Build listo en: dist\CobranzaTray\CobranzaTray.exe
exit /b 0

:fatal
echo.
echo ERROR: Fallo la compilacion del EXE (make_tray_exe.bat).
exit /b 1
