@echo off
setlocal enabledelayedexpansion
title Instalando Office 365 + Activando con MAS (Ohook)...
pushd "%~dp0"

:: Verificar permisos de administrador
fltmc >nul 2>&1 || (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Detectar arquitectura
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" set "CPU=64"
if "%PROCESSOR_ARCHITECTURE%"=="x86" set "CPU=32"

:: Crear archivo de configuración XML
set "XML=Office365_Config.xml"
(
echo ^<Configuration^>
echo   ^<Add OfficeClientEdition="%CPU%" Channel="Current"^>
echo     ^<Product ID="O365ProPlusRetail"^>
echo       ^<Language ID="MatchOS" Fallback="en-US" /^>
echo     ^</Product^>
echo   ^</Add^>
echo   ^<Display Level="None" AcceptEULA="TRUE" /^>
echo   ^<Property Name="ForceAppShutdown" Value="TRUE" /^>
echo ^</Configuration^>
) > "%XML%"

:: Verificar que setup.exe esté presente
if not exist "setup.exe" (
    echo [ERROR] No se encontro setup.exe en la carpeta actual.
    echo Asegurate de descargar el Office Deployment Tool y extraer setup.exe aqui.
    pause
    exit /b
)

:: Instalar Office
echo Instalando Office 365 ProPlus (%CPU%-bit)...
start /wait setup.exe /configure "%XML%"

:: Desactivar telemetría
reg add "HKLM\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" /v "DisableTelemetry" /t REG_DWORD /d "1" /f >nul 2>&1

:: Activar Office con MAS (Ohook)
echo Activando Office con MAS (Ohook)...
powershell -Command "irm https://get.activated.win | iex; MAS; ohook"

:: Limpiar
del "%XML%" 2>nul
echo.
echo [OK] Instalacion y activacion finalizadas.
timeout /t 5 >nul
exit