# =====================================================================
#  hacer_exe.ps1
#  Empaqueta TRES ejecutables independientes en dist\:
#
#    GeneradorFS.exe   crea un Word NUEVO en salidas\   (comportamiento clasico)
#    RefrescarFS.exe   ACTUALIZA el documento base conservando la redaccion
#    EstadosFinancieros.exe  UN SOLO icono que pregunta cual de los dos
#
#  Los tres aceptan que se les arrastre el Excel encima. RefrescarFS necesita
#  un config.json junto al .exe con la clave "documento_base".
#
#  Requisitos (equipo de desarrollo): tools\bootstrap_python.ps1 ya
#  ejecutado y  .\python\python.exe -m pip install pyinstaller
#  Uso:  powershell -ExecutionPolicy Bypass -File .\tools\hacer_exe.ps1
# =====================================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
# Python portable del proyecto; si no esta, se acepta uno del sistema. Antes
# se abortaba en seco, y en un equipo sin .\python\ no habia forma de
# reconstruir los .exe (que es justo cuando hace falta).
$py = Join-Path $root 'python\python.exe'
if (-not (Test-Path $py)) {
    $py = $null
    foreach ($cand in @('py', 'python')) {
        $c = Get-Command $cand -ErrorAction SilentlyContinue
        if ($c) { $py = $c.Source; break }
    }
    if (-not $py) {
        throw "No hay Python. Ejecute tools\bootstrap_python.ps1, o instale Python 3."
    }
    Write-Host "Sin .\python\ portable; uso $py" -ForegroundColor Yellow
}

# nombre de .exe  ->  script de entrada en src\
$objetivos = [ordered]@{
    'GeneradorFS' = 'generador_fs.py'
    'RefrescarFS' = 'refrescar_fs.py'
    'EstadosFinancieros' = 'fs_menu.py'
}

# La carpeta de trabajo de PyInstaller va FUERA de OneDrive: si se deja
# dentro, la sincronizacion bloquea archivos a medio escribir y el
# empaquetado falla con "Acceso denegado".
# --specpath mueve la base de las rutas relativas, asi que los recursos
# se pasan en absoluto.
$plantilla     = Join-Path $root 'plantillas\plantilla_estado_situacion_financiera.docx'
$trabajo = Join-Path $env:TEMP ("fsbuild-" + [guid]::NewGuid().ToString())
New-Item -ItemType Directory -Force -Path $trabajo | Out-Null

# La config que se EMBEBE en el .exe se limpia primero: 'documento_base' es
# una ruta absoluta de la maquina que compila, y quedaba congelada dentro
# del binario. En cuanto el .exe cambiaba de manos (o de perfil de usuario)
# apuntaba a una carpeta inexistente, y TODAS las opciones que dependen del
# documento morian a la vez con el mismo error. El .exe debe salir sin
# documento y pedirlo la primera vez; el config.json de al lado manda.
$configuracion = Join-Path $trabajo 'config.json'
$cfg = Get-Content (Join-Path $root 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$cfg.documento_base = ''
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($configuracion, ($cfg | ConvertTo-Json -Depth 20), $utf8)
Write-Host "Config embebida sin 'documento_base' (se pedira en el primer uso)." -ForegroundColor DarkGray

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
            --workpath $trabajo `
            --specpath $trabajo `
            --paths 'src' `
            --add-data "$plantilla;." `
            --add-data "$configuracion;." `
            --collect-submodules docxtpl `
            --collect-submodules docx `
            $entrada
        if ($LASTEXITCODE -ne 0) { throw "PyInstaller devolvio $LASTEXITCODE para $nombre" }
    }
}
finally {
    Pop-Location
    Remove-Item $trabajo -Recurse -Force -ErrorAction SilentlyContinue
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
