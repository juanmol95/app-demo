@echo off
title Registro de Implantes - servidor (no cierres esta ventana mientras uses la app)
cd /d "%~dp0servidor"

rem Si el servidor ya esta encendido, solo abre la app
powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort 8001 -State Listen -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"
if %errorlevel%==0 (
    start "" "http://127.0.0.1:8001/"
    exit /b
)

rem Abre la app en el navegador cuando el servidor termine de arrancar
start "" /min cmd /c "timeout /t 4 /nobreak >nul & start "" http://127.0.0.1:8001/"

echo Encendiendo el servidor de Registro de Implantes...
echo La app se abrira sola en el navegador:  http://127.0.0.1:8001/
echo Para apagarlo, cierra esta ventana.
echo.
python "App (1).py"
pause
