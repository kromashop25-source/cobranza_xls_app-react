@echo off
setlocal
cd /d "%~dp0\.."

cd frontend
call npm run build
if errorlevel 1 goto :fatal
cd ..

call .venv\Scripts\activate.bat
if errorlevel 1 goto :fatal

python -m pip install -U pip wheel "setuptools<81" pyinstaller pyinstaller-hooks-contrib
if errorlevel 1 goto :fatal

python -m PyInstaller backend\CobranzaTray.spec --clean --noconfirm
if errorlevel 1 goto :fatal

echo.
echo Build listo en: dist\CobranzaTray\CobranzaTray.exe
exit /b 0

:fatal
echo.
echo ERROR: Fallo la compilacion del EXE (frontend\COMPILADOR.BAT).
exit /b 1
