@echo off
title Registro de Implantes - MODO CELULAR (no cierres esta ventana)
cd /d "%~dp0servidor"
echo ==================================================================
echo   MODO CELULAR: el servidor queda visible para la red Wi-Fi.
echo   Usalo SOLO en una red privada, por ejemplo la ZONA WI-FI de tu
echo   celular. NO lo uses en eduroam ni en redes publicas: la app
echo   todavia no tiene contrasena y cualquiera en la red podria entrar.
echo ==================================================================
echo.

rem Avisar si Windows considera publica la red actual
powershell -NoProfile -Command "$r = Get-NetConnectionProfile -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'eduroam' }; if ($r) { Write-Host 'ATENCION: estas conectado a eduroam. Conectate a la zona Wi-Fi de tu celular.' -ForegroundColor Red; exit 1 } else { exit 0 }"
if errorlevel 1 (
    choice /c SN /m "Seguro que quieres abrir el servidor en esta red"
    if errorlevel 2 exit /b
)

echo.
echo En el celular, conectado a la MISMA red, abre una de estas direcciones:
powershell -NoProfile -Command "Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } | ForEach-Object { '     https://' + $_.IPAddress + ':8001/' }"
echo.
echo Si Windows pregunta si Python puede usar la red, permite SOLO "Redes privadas".
echo La primera vez el celular dira "La conexion no es privada": pulsa
echo "Configuracion avanzada" y luego "Continuar". Despues la camara funciona.
echo Para apagarlo, cierra esta ventana.
echo.
set HOST=0.0.0.0
set HTTPS=1
python "App (1).py"
pause
