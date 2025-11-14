@echo off
setlocal
cd /d "%~dp0\.."

REM 1) Asegura que el exe está listo
if not exist dist\CobranzaTray\CobranzaTray.exe (
  echo No existe dist\CobranzaTray\CobranzaTray.exe. Ejecuta scripts\make_tray_exe.bat primero.
  exit /b 1
)

REM 2) Compilar Inno Setup (requiere Inno instalado y en PATH)
iscc installer\CobranzaXLS.iss

echo.
echo Instalador generado en: dist\installer\CobranzaXLSSetup.exe
endlocal
