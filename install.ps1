# install.ps1 (actualizado con descarga de setup.exe)
Write-Host "Descargando Office365 desde tu repo..." -ForegroundColor Cyan

$repoUrl = "https://github.com/rubenmontes58/Office365/archive/refs/heads/main.zip"
$temp = "$env:TEMP\Office365"
$output = "$temp\Office365.zip"

# Crear carpeta temporal
New-Item -ItemType Directory -Force -Path $temp | Out-Null

# Descargar repo
Invoke-WebRequest -Uri $repoUrl -OutFile $output
Expand-Archive -Path $output -DestinationPath $temp -Force

# Ruta al instalador
$installer = "$temp\Office365-main\Office365_Silent_Ohook.cmd"

# Verificar si falta setup.exe
$setupPath = "$temp\Office365-main\setup.exe"
if (-not (Test-Path $setupPath)) {
    Write-Host "Descargando Office Deployment Tool..." -ForegroundColor Yellow
    $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20158.exe"
    $odtInstaller = "$temp\odt.exe"
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtInstaller

    # Extraer ODT en la misma carpeta
    Start-Process -FilePath $odtInstaller -ArgumentList "/quiet /extract:`"$temp\Office365-main`"" -Wait
    Remove-Item $odtInstaller -Force
}

# Ejecutar instalador
if (Test-Path $installer) {
    Write-Host "Ejecutando instalador..." -ForegroundColor Green
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$installer`"" -Wait -Verb RunAs
} else {
    Write-Host "[ERROR] No se encontro el instalador." -ForegroundColor Red
}

Write-Host "Instalacion y activacion finalizadas." -ForegroundColor Green