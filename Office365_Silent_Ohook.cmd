@echo off
setlocal enabledelayedexpansion
title Instalando Office 365 + Activando con MAS (Ohook)...
pushd "%~dp0"

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

:: Crear XML con ruta absoluta
set "XML=%~dp0Office365_Config.xml"
echo [INFO] Creando XML: "%XML%"

> "%XML%" (
    echo ^<Configuration^>
    echo   ^<Add OfficeClientEdition="%CPU%" Channel="Current"^>
    echo     ^<Product ID="O365ProPlusRetail"^>
    echo       ^<Language ID="MatchOS" Fallback="en-US" /^>
    echo     ^</Product^>
    echo   ^</Add^>
    echo   ^<Display Level="Full" AcceptEULA="TRUE" /^>
    echo   ^<Property Name="ForceAppShutdown" Value="TRUE" /^>
    echo ^</Configuration^>
)

:: Verificar que se haya creado
if not exist "%XML%" (
    echo [ERROR] No se pudo crear el archivo XML.
    pause
    exit /b
)

:: Mostrar contenido del XML (para debug)
echo [INFO] Contenido del XML:
type "%XML%"

:: Instalar Office
echo [INFO] Lanzando setup.exe con /configure...
start /wait setup.exe /configure "%XML%"

:: Activar con MAS
echo [INFO] Activando Office con MAS (Ohook)...
powershell -Command "irm https://get.activated.win | iex"

echo [OK] Proceso finalizado.
pause
exit