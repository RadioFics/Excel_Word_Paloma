# =====================================================================
#  hacer_paquete.ps1
#  Empaqueta el proyecto COMPLETO (incluido .\python\) en un unico .zip
#  que se ejecuta en cualquier Windows x64 sin instalar nada y sin
#  internet. Ese .zip es "el descargable" para la prueba externa.
#
#  Requisito: haber corrido antes tools\bootstrap_python.ps1.
#  Uso:  powershell -ExecutionPolicy Bypass -File .\tools\hacer_paquete.ps1
# =====================================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$py   = Join-Path $root 'python\python.exe'
if (-not (Test-Path $py)) {
    throw "Falta .\python\  ->  ejecute primero  tools\bootstrap_python.ps1"
}

$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force -Path $dist | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd'
$zip   = Join-Path $dist "GeneradorFS_portable_$stamp.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }

$incluir = @(
    'generar.bat', 'refrescar.bat', 'config.json', 'requirements.txt',
    'README.md',
    'src', 'docs', 'plantillas', 'ejemplos',
    'python', 'tools'
) | ForEach-Object { Join-Path $root $_ } | Where-Object { Test-Path $_ }

$tmp = Join-Path $env:TEMP ("pkg-" + [guid]::NewGuid().ToString())
New-Item -ItemType Directory -Force -Path (Join-Path $tmp 'salidas') | Out-Null
Copy-Item -Path $incluir -Destination $tmp -Recurse -Force

# quitar caches de Python para aligerar
Get-ChildItem -Path $tmp -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# tar.exe (incluido en Windows 10 1803+ / 11) genera un .zip con separador
# "/" que abre bien en cualquier extractor. Compress-Archive de Windows
# PowerShell usa "\" y algunos extractores fuera de Windows fallan.
$tar = Join-Path $env:SystemRoot 'System32\tar.exe'
if (Test-Path $tar) {
    & $tar -a -c -f $zip -C $tmp .
    if ($LASTEXITCODE -ne 0) { throw "tar.exe devolvio $LASTEXITCODE" }
} else {
    Write-Warning "tar.exe no disponible; usando Compress-Archive (zip con backslashes)."
    Compress-Archive -Path (Join-Path $tmp '*') -DestinationPath $zip -Force
}
Remove-Item $tmp -Recurse -Force

$mb = [math]::Round((Get-Item $zip).Length / 1MB, 1)
Write-Host ""
Write-Host "Paquete creado:  $zip   ($mb MB)" -ForegroundColor Green
Write-Host "Subalo a OneDrive o como 'release' del repositorio."
Write-Host "En el equipo destino: descomprimir y doble clic en generar.bat"
Write-Host "(o ejecutar  tools\verificar.ps1  para la comprobacion automatica)."
