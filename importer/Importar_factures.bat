@echo off
setlocal
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Importar_factures_v2.ps1" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
  echo La importacio ha acabat amb errors.
) else (
  echo La importacio ha acabat correctament.
)
echo Prem una tecla per tancar aquesta finestra.
pause >nul
exit /b %EXIT_CODE%
