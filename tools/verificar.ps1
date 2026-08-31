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
$ejemplo = Join-Path $root 'Copia_Editable_con_columna_Tipo.xlsx'
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

    $mods = (& $py -c "import openpyxl, docxtpl, jinja2; print('mods-ok')" 2>&1) -join ' '
    if ($mods -match 'mods-ok') { ok "Dependencias importan (openpyxl, docxtpl, jinja2)" }
    else { no "Dependencias no importan: $mods" }

    if (Test-Path $ejemplo) { ok "Excel de ejemplo presente" }
    else { no "Falta Copia_Editable_con_columna_Tipo.xlsx" }
}

if ($script:fallos -eq 0) {
    $prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    $salida = & $py (Join-Path $root 'generador_fs.py') $ejemplo '--revisar' 2>&1
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
