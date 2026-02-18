@echo off
setlocal
title Build .EXE sin consola - CobranzaApp

cd /d "%~dp0"
set "ROOT_DIR=%~dp0.."
for %%I in ("%ROOT_DIR%") do set "ROOT_DIR=%%~fI"

echo ===============================================
echo   Build .EXE sin consola - CobranzaApp
echo ===============================================

set "VENV_DIR=%CD%\.buildvenv"

if exist "%VENV_DIR%\Scripts\python.exe" (
  set "VENV_SCRIPTS=%VENV_DIR%\Scripts"
  set "PYTHON=%VENV_SCRIPTS%\python.exe"
  "%PYTHON%" -c "import sys; raise SystemExit(0 if (3,11) <= sys.version_info[:2] < (3,14) else 1)" >nul 2>&1
  if errorlevel 1 (
    echo Venv de build invalido o con Python no soportado, recreando...
    rmdir /s /q "%VENV_DIR%"
  )
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
  if defined VIRTUAL_ENV (
    set "BASE_PY=%VIRTUAL_ENV%\Scripts\python.exe"
    if exist "%BASE_PY%" (
      "%BASE_PY%" -m venv "%VENV_DIR%" >nul 2>&1
    )
  )
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
  py -3.13 -m venv "%VENV_DIR%" >nul 2>&1
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
  py -3.12 -m venv "%VENV_DIR%" >nul 2>&1
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
  python -m venv "%VENV_DIR%" >nul 2>&1
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
  py -3 -m venv "%VENV_DIR%" || (echo Error creando venv & exit /b 1)
)

set "VENV_SCRIPTS=%VENV_DIR%\Scripts"
set "PYTHON=%VENV_SCRIPTS%\python.exe"

if not exist "%PYTHON%" (
  echo Python no encontrado en "%PYTHON%"
  exit /b 1
)
"%PYTHON%" -c "import sys" >nul 2>&1 || (
  echo Python de build no funciona en "%PYTHON%"
  exit /b 1
)

echo [2/6] Actualizando pip/setuptools/wheel...
"%PYTHON%" -m pip install --upgrade pip wheel "setuptools<81" || (
  echo Error actualizando herramientas base de Python.
  exit /b 1
)

echo [3/6] Instalando dependencias del proyecto...
"%PYTHON%" -m pip install -r "%ROOT_DIR%\requirements.txt" || (
  echo Error instalando dependencias del proyecto.
  exit /b 1
)

echo [4/6] Instalando herramientas de build...
"%PYTHON%" -m pip install pyinstaller pywin32 || (
  echo Error instalando herramientas de build.
  exit /b 1
)

echo [5/6] Limpieza de build/dist previos...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [6/6] Compilando con PyInstaller...
"%PYTHON%" -m PyInstaller ^
  --name "CobranzaApp" ^
  --noconsole ^
  --onefile ^
  --hidden-import win32com ^
  --hidden-import win32timezone ^
  --collect-submodules win32com ^
  --collect-submodules pythoncom ^
  --hidden-import app ^
  --hidden-import app.main ^
  --hidden-import app.services.excel_copy ^
  --add-data "app\static;app\static" ^
  --add-data "app\data;app\data" ^
  "%CD%\run_app.py"

if errorlevel 1 (
  echo Error durante la compilacion.
  exit /b 1
)

echo Hecho! EXE en .\dist\CobranzaApp.exe
exit /b 0
