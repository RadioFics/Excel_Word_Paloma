@echo off
REM ============================================================
REM  refrescar.bat
REM  ACTUALIZA el documento base (el que vive en OneDrive)
REM  conservando toda la redaccion. NO crea un archivo nuevo.
REM
REM  Uso 1: arrastre el Excel sobre este icono y sueltelo.
REM  Uso 2: doble clic; busca el Excel por convencion de nombre.
REM
REM  Que documento actualiza: el de config.local.json (o, si no,
REM  el de config.json) -> "documento_base".
REM  Si al documento le faltan las regiones, se le anaden solas
REM  antes de volcar las cifras: en blanco se usa de base, y con
REM  redaccion encima el estado entra como un apartado aparte.
REM  Para impedirlo:  refrescar.bat --no-preparar
REM
REM  IMPORTANTE: cierre el documento en Word antes de refrescar.
REM ============================================================

cd /d "%~dp0"

REM Busca un Python usable. Si no lo hay, buscar_python.bat ya ha
REM explicado que falta: aqui solo hay que dejar la ventana abierta
REM para que se pueda leer.
call "%~dp0tools\buscar_python.bat"
if errorlevel 1 (
    echo.
    pause
    exit /b 1
)

if "%~1"=="" (
    "%PY%" src\refrescar_fs.py
) else (
    "%PY%" src\refrescar_fs.py "%~1" %2 %3
)
if errorlevel 1 pause
