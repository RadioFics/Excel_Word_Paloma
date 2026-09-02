@echo off
REM ============================================================
REM  probar.bat
REM  Bateria de pruebas del refresco Excel -> Word.
REM  Trabaja sobre copias temporales: no toca ni el documento
REM  ni el libro reales.
REM
REM  Arrastrele un .xlsx encima para anadir una pasada con ese
REM  libro de verdad.
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
    "%PY%" tools\probar_refresco.py
) else (
    "%PY%" tools\probar_refresco.py --libro "%~1"
)
echo.
pause
