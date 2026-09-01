@echo off
REM ============================================================
REM  refrescar.bat
REM  ACTUALIZA el documento base (el que vive en OneDrive)
REM  conservando toda la redaccion. NO crea un archivo nuevo.
REM
REM  Uso 1: arrastre el Excel sobre este icono y sueltelo.
REM  Uso 2: doble clic; busca el Excel por convencion de nombre.
REM
REM  Que documento actualiza: el de config.json -> "documento_base".
REM  La primera vez, o si anadio zonas nuevas al documento, ejecute
REM  una vez:   refrescar.bat --preparar
REM
REM  IMPORTANTE: cierre el documento en Word antes de refrescar.
REM
REM  Para crear un Word NUEVO en salidas\ use generar.bat.
REM ============================================================
cd /d "%~dp0"

set "PY=%~dp0python\python.exe"
if not exist "%PY%" set "PY=python"

if "%~1"=="" (
    "%PY%" src\refrescar_fs.py
) else (
    "%PY%" src\refrescar_fs.py "%~1" %2 %3
)
