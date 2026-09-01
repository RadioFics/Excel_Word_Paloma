@echo off
REM ============================================================
REM  generar.bat
REM  Crea un Word NUEVO en salidas\ a partir de la plantilla.
REM  NO actualiza el documento base: para eso use refrescar.bat.
REM
REM  Uso 1: arrastre el archivo Excel sobre este icono y sueltelo.
REM  Uso 2: haga doble clic sin arrastrar nada; buscara el Excel
REM         mas reciente en esta misma carpeta cuyo nombre
REM         contenga "FS" (ver BUSCAR_POR_CONVENCION en el script).
REM
REM  Usa el Python PORTABLE de .\python\ si existe (lo crea
REM  tools\bootstrap_python.ps1). Si no, cae al 'python' del PATH.
REM ============================================================
cd /d "%~dp0"

set "PY=%~dp0python\python.exe"
if not exist "%PY%" set "PY=python"

if "%~1"=="" (
    "%PY%" src\generador_fs.py
) else (
    "%PY%" src\generador_fs.py "%~1"
)
