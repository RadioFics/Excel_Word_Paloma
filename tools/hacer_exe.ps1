# =====================================================================
#  hacer_exe.ps1
#  Empaqueta la aplicacion en UN SOLO archivo: dist\GeneradorFS.exe
#  El usuario final solo descarga ese .exe y lo ejecuta (o arrastra su
#  Excel encima). No hay que descomprimir nada ni instalar Python.
#
#  Requisitos (equipo de desarrollo): tools\bootstrap_python.ps1 ya
#  ejecutado y  .\python\python.exe -m pip install pyinstaller
#  Uso:  powershell -ExecutionPolicy Bypass -File .\tools\hacer_exe.ps1
# =====================================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$py   = Join-Path $root 'python\python.exe'
if (-not (Test-Path $py)) { throw "Falta .\python\  ->  tools\bootstrap_python.ps1" }

Push-Location $root
try {
    & $py -m PyInstaller `
        --onefile `
        --name GeneradorFS `
        --console `
        --noconfirm `
        --clean `
        --add-data "plantilla_estado_situacion_financiera.docx;." `
        --add-data "config.json;." `
        --collect-submodules docxtpl `
        --collect-submodules docx `
        generador_fs.py
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller devolvio $LASTEXITCODE" }
}
finally {
    Pop-Location
}

$exe = Join-Path $root 'dist\GeneradorFS.exe'
$mb  = [math]::Round((Get-Item $exe).Length / 1MB, 1)
$hash = (Get-FileHash -Algorithm SHA256 $exe).Hash
Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host " Listo:  $exe   ($mb MB)"                              -ForegroundColor Green
Write-Host " SHA-256: $hash"
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "Publiquelo como release del repositorio junto con ese hash."
