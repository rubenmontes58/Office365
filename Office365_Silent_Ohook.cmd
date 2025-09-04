@echo off
setlocal enabledelayedexpansion
title Instalando Office 365 + Activando con MAS (Ohook)...
pushd "%~dp0"

:: LOG: carpeta actual
echo [INFO] Carpeta actual: "%~dp0"
echo [INFO] Buscando setup.exe...

if not exist "setup.exe" (
    echo [ERROR] No se encontro setup.exe en esta carpeta.
    pause
    exit /b
)

:: Detectar arquitectura
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" set "CPU=64"
if "%PROCESSOR_ARCHITECTURE%"=="x86" set "CPU=32"

:: Crear XML SIEMPRE en esta carpeta
set "XML=%~dp0Office365_Config.xml"
echo [INFO] Creando configuracion XML: "%XML%"

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

:: Verificar que se haya creado
if not exist "%XML%" (
    echo [ERROR] No se pudo crear el archivo XML.
    pause
    exit /b
)

:: Instalar Office
echo [INFO] Instalando Office 365 ProPlus (%CPU%-bit)...
start /wait setup.exe /configure "%XML%"

:: Desactivar telemetría
reg add "HKLM\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" /v "DisableTelemetry" /t REG_DWORD /d "1" /f >nul 2>&1

:: Activar con MAS
echo [INFO] Activando Office con MAS (Ohook)...
powershell -Command "irm https://get.activated.win | iex"

:: Fin
echo [OK] Instalacion y activacion finalizadas.
pause
exit