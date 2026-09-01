@echo off
REM ============================================================
REM  estados_financieros.bat
REM  UN SOLO icono que pregunta que hacer:
REM    1) ACTUALIZAR el documento de siempre (conserva la redaccion)
REM    2) CREAR un documento nuevo en salidas\
REM    3) VER el estado del proyecto
REM
REM  Arrastre el Excel encima, o haga doble clic.
REM
REM  Los dos caminos siguen existiendo por separado:
REM    generar.bat    = opcion 2 directa
REM    refrescar.bat  = opcion 1 directa
REM ============================================================
cd /d "%~dp0"

set "PY=%~dp0python\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" src\fs_menu.py %*
