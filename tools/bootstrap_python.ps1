# =====================================================================
#  bootstrap_python.ps1
#  Instala un Python PORTABLE dentro del proyecto (.\python\), sin
#  permisos de administrador y sin instalador MSI (no usa winget ni
#  la Microsoft Store). Ejecutar una sola vez por equipo nuevo.
#
#  Requiere acceso a internet la primera vez:
#    - https://www.python.org        (paquete embebido de Python)
#    - https://bootstrap.pypa.io     (get-pip)
#    - https://pypi.org              (dependencias de requirements.txt)
#
#  Uso:
#    Clic derecho sobre este archivo  ->  "Ejecutar con PowerShell"
#  o desde una terminal, en la carpeta del proyecto:
#    powershell -ExecutionPolicy Bypass -File .\tools\bootstrap_python.ps1
# =====================================================================

$ErrorActionPreference = 'Stop'

# Version de Python a fijar. El paquete "embeddable" es solo archivos:
# se descomprime, no se instala, no toca el registro ni el sistema.
$PyVersion = '3.12.10'

$root   = Split-Path -Parent $PSScriptRoot
$target = Join-Path $root 'python'
$req    = Join-Path $root 'requirements.txt'

if (Test-Path (Join-Path $target 'python.exe')) {
    Write-Host "Ya existe .\python\python.exe -- nada que reinstalar." -ForegroundColor Green
    & (Join-Path $target 'python.exe') --version
    exit 0
}

if (-not (Test-Path $req)) {
    throw "No se encontro requirements.txt en $root"
}

# ---- 1. Descargar el paquete embebido -------------------------------
$zip = Join-Path $env:TEMP ("python-embed-{0}.zip" -f (Get-Random))
$url = "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-embed-amd64.zip"
Write-Host "Descargando  $url"
Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing

# ---- 2. Extraer en .\python\ --------------------------------------
Write-Host "Extrayendo   $target"
New-Item -ItemType Directory -Force -Path $target | Out-Null
Expand-Archive -Path $zip -DestinationPath $target -Force
Remove-Item $zip -Force

# ---- 3. Habilitar site-packages en el archivo ._pth ---------------
# El paquete embebido trae 'import site' comentado y no busca en
# Lib\site-packages. Sin este ajuste, pip instala pero nada se importa.
$pth = Get-ChildItem -Path $target -Filter 'python*._pth' | Select-Object -First 1
if (-not $pth) { throw "No se encontro el archivo python*._pth en $target" }

$lines = Get-Content $pth.FullName | ForEach-Object {
    $_ -replace '^\s*#\s*import\s+site\s*$', 'import site'
}
if ($lines -notcontains 'Lib\site-packages') { $lines += 'Lib\site-packages' }
if ($lines -notcontains 'import site')       { $lines += 'import site' }
Set-Content -Path $pth.FullName -Value $lines -Encoding ASCII

# ---- 4. Bootstrap de pip ----------------------------------------
$getpip = Join-Path $target 'get-pip.py'
Write-Host "Instalando   pip"
Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $getpip -UseBasicParsing
& (Join-Path $target 'python.exe') $getpip --no-warn-script-location --no-cache-dir
Remove-Item $getpip -Force

# ---- 5. Dependencias del proyecto (versiones fijadas) ------------
Write-Host "Instalando   dependencias de requirements.txt"
& (Join-Path $target 'python.exe') -m pip install --no-warn-script-location --no-cache-dir -r $req

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host " Python portable listo en:  $target"                   -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
& (Join-Path $target 'python.exe') --version
& (Join-Path $target 'python.exe') -m pip list --format=columns
Write-Host ""
Write-Host "Ya puede usar  generar.bat  (usara este Python automaticamente)."
