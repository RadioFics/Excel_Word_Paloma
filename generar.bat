@echo off
REM ============================================================
REM  generar.bat
REM  Crea un Word NUEVO en salidas\ a partir de la plantilla.
REM  NO actualiza el documento base: para eso use refrescar.bat.
REM
REM  Uso 1: arrastre el archivo Excel sobre este icono y sueltelo.
REM  Uso 2: haga doble clic sin arrastrar nada; buscara el Excel
REM         mas reciente en esta misma carpeta cuyo nombre
REM         contenga "FS".
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
    "%PY%" src\generador_fs.py
) else (
    "%PY%" src\generador_fs.py "%~1"
)
if errorlevel 1 pause
