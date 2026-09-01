# =====================================================================
#  hacer_exe.ps1
#  Empaqueta DOS ejecutables independientes en dist\:
#
#    GeneradorFS.exe   crea un Word NUEVO en salidas\   (comportamiento clasico)
#    RefrescarFS.exe   ACTUALIZA el documento base conservando la redaccion
#
#  Ambos aceptan que se les arrastre el Excel encima. RefrescarFS necesita
#  un config.json junto al .exe con la clave "documento_base".
#
#  Requisitos (equipo de desarrollo): tools\bootstrap_python.ps1 ya
#  ejecutado y  .\python\python.exe -m pip install pyinstaller
#  Uso:  powershell -ExecutionPolicy Bypass -File .\tools\hacer_exe.ps1
# =====================================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$py   = Join-Path $root 'python\python.exe'
if (-not (Test-Path $py)) { throw "Falta .\python\  ->  tools\bootstrap_python.ps1" }

# nombre de .exe  ->  script de entrada en src\
$objetivos = [ordered]@{
    'GeneradorFS' = 'generador_fs.py'
    'RefrescarFS' = 'refrescar_fs.py'
}

Push-Location $root
try {
    foreach ($nombre in $objetivos.Keys) {
        $entrada = Join-Path 'src' $objetivos[$nombre]
        Write-Host ""
        Write-Host "Empaquetando $nombre  ($entrada)" -ForegroundColor Cyan

        # --paths src : para que encuentre fs_contrato / fs_documento, que
        # se importan por nombre y no cuelgan de un paquete.
        & $py -m PyInstaller `
            --onefile `
            --name $nombre `
            --console `
            --noconfirm `
            --clean `
            --paths 'src' `
            --add-data "plantillas\plantilla_estado_situacion_financiera.docx;." `
            --add-data "config.json;." `
            --collect-submodules docxtpl `
            --collect-submodules docx `
            $entrada
        if ($LASTEXITCODE -ne 0) { throw "PyInstaller devolvio $LASTEXITCODE para $nombre" }
    }
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
foreach ($nombre in $objetivos.Keys) {
    $exe = Join-Path $root "dist\$nombre.exe"
    if (-not (Test-Path $exe)) { throw "No se genero $exe" }
    $mb   = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    $hash = (Get-FileHash -Algorithm SHA256 $exe).Hash
    Write-Host " $nombre.exe   ($mb MB)" -ForegroundColor Green
    Write-Host "   SHA-256: $hash"
}
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "Publiquelos como release del repositorio junto con esos hashes."
Write-Host "RefrescarFS.exe necesita config.json a su lado (documento_base)."
