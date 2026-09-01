# =====================================================================
#  verificar.ps1
#  Comprueba que este equipo puede ejecutar el generador y que produce
#  el resultado esperado sobre el Excel de ejemplo.
#  Verde = todo bien. Rojo = algo falta.
#
#  Uso:  clic derecho -> Ejecutar con PowerShell
#    o:  powershell -ExecutionPolicy Bypass -File .\tools\verificar.ps1
# =====================================================================
$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $PSScriptRoot
$py      = Join-Path $root 'python\python.exe'
$ejemplo = Join-Path $root 'ejemplos\Copia_Editable_con_columna_Tipo.xlsx'
$src     = Join-Path $root 'src'
$script:fallos = 0

function ok($m) { Write-Host "  [OK]    $m" -ForegroundColor Green }
function no($m) { Write-Host "  [FALLA] $m" -ForegroundColor Red; $script:fallos++ }

Write-Host ""
Write-Host "Verificando el entorno..." -ForegroundColor Cyan

if (Test-Path $py) {
    ok "Python portable presente (.\python\python.exe)"
} else {
    no "Falta .\python\python.exe  ->  ejecute  tools\bootstrap_python.ps1"
}

if ($script:fallos -eq 0) {
    $v = (& $py --version 2>&1) -join ' '
    if ($v -match '3\.1[0-9]') { ok "Version: $v" } else { no "Version inesperada: $v" }

    $mods = (& $py -c "import openpyxl, docxtpl, jinja2, docx; print('mods-ok')" 2>&1) -join ' '
    if ($mods -match 'mods-ok') { ok "Dependencias importan (openpyxl, docxtpl, jinja2, python-docx)" }
    else { no "Dependencias no importan: $mods" }

    if (Test-Path $ejemplo) { ok "Excel de ejemplo presente (ejemplos\)" }
    else { no "Falta ejemplos\Copia_Editable_con_columna_Tipo.xlsx" }

    if (Test-Path (Join-Path $src 'generador_fs.py')) { ok "Codigo presente (src\)" }
    else { no "Falta src\generador_fs.py" }
}

if ($script:fallos -eq 0) {
    $prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    $salida = & $py (Join-Path $src 'generador_fs.py') $ejemplo '--revisar' 2>&1
    $rc = $LASTEXITCODE
    $ErrorActionPreference = $prev
    if ($rc -ne 0) { no "El generador termino con codigo $rc"; $salida | ForEach-Object { Write-Host "    $_" } }
    $linea  = ($salida | Select-String 'trasladadas')
    if ($linea -and "$linea" -match '33') {
        ok ("Genera 33 lineas sobre el ejemplo  ->  " + ("$linea").Trim())
    } else {
        no "Conteo de lineas inesperado"
        $salida | ForEach-Object { Write-Host "    $_" }
    }
    $csv = Join-Path $root 'salidas\revisar_tipos.csv'
    if (Test-Path $csv) { ok "Escribio salidas\revisar_tipos.csv" } else { no "No escribio el CSV de revision" }
}

# ---- motor de documento vivo ---------------------------------------
if ($script:fallos -eq 0) {
    $prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    $tmp = Join-Path $env:TEMP ("fsverif-" + [guid]::NewGuid().ToString() + ".docx")
    $salida = & $py (Join-Path $src 'fs_documento.py') 'plantilla' $tmp '--excel' $ejemplo 2>&1
    $rc = $LASTEXITCODE
    if ($rc -eq 0 -and (Test-Path $tmp)) { ok "Construye un documento base con sus regiones" }
    else { no "No pudo construir el documento base"; $salida | ForEach-Object { Write-Host "    $_" } }

    if (Test-Path $tmp) {
        $salida = & $py (Join-Path $src 'fs_documento.py') 'refrescar' $tmp $ejemplo 2>&1
        $rc = $LASTEXITCODE
        $linea = ($salida | Select-String "Tabla 'principal'")
        if ($rc -eq 0 -and $linea -and "$linea" -match '33') {
            ok ("Refresca el documento  ->  " + ("$linea").Trim())
        } else {
            no "El refresco no escribio las 33 filas esperadas"
            $salida | ForEach-Object { Write-Host "    $_" }
        }
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        Remove-Item "$tmp.bak" -Force -ErrorAction SilentlyContinue
    }
    $ErrorActionPreference = $prev
}

Write-Host ""
if ($script:fallos -eq 0) {
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host " TODO CORRECTO. La aplicacion funciona en este equipo." -ForegroundColor Green
    Write-Host "========================================================" -ForegroundColor Green
    exit 0
} else {
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host " $($script:fallos) comprobacion(es) fallaron. Ver arriba." -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Red
    exit 1
}
