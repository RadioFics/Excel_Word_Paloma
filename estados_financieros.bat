@echo off
REM ============================================================
REM  estados_financieros.bat
REM  UN SOLO icono que pregunta que hacer:
REM    1) ACTUALIZAR el documento de siempre (conserva la redaccion)
REM    2) CREAR un documento nuevo en salidas\
REM    3) CAMBIAR el documento que se actualiza
REM    4) Candado de las cifras
REM
REM  Arrastre el Excel encima, o haga doble clic.
REM
REM  Si no tiene Python, use dist\EstadosFinancieros.exe: es lo
REM  mismo y no necesita nada instalado.
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

"%PY%" src\fs_menu.py %*
if errorlevel 1 pause
