@echo off
REM ============================================================
REM  buscar_python.bat
REM  Deja en la variable PY la ruta de un Python USABLE (con las
REM  dependencias puestas), o explica que falta y devuelve error.
REM
REM  Se llama con:   call "%~dp0tools\buscar_python.bat"
REM
REM  Existe porque "si no esta el portable, usa python" no basta en
REM  Windows: el 'python' del PATH suele ser el ATAJO de la Microsoft
REM  Store, que no es Python. Imprime un aviso, devuelve 9009, y la
REM  ventana se cierra antes de que nadie lea nada.
REM ============================================================
set "PY="
set "FS_MOTIVO="

REM --- 1) el Python portable del proyecto ---------------------
if exist "%~dp0..\python\python.exe" (
    set "PY=%~dp0..\python\python.exe"
    goto :comprobar
)

REM --- 2) el lanzador 'py' ------------------------------------
REM  Es lo que instala python.org. No lo tapa el alias de la Store.
for /f "delims=" %%i in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%i"
if defined PY goto :comprobar

REM --- 3) el 'python' del PATH --------------------------------
REM  Se le pregunta por sys.executable a proposito: el atajo de la
REM  Store falla aqui y deja PY vacia, en vez de colarse.
for /f "delims=" %%i in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%i"
if not defined PY goto :sin_python

:comprobar
"%PY%" -c "import openpyxl, docx, docxtpl" >nul 2>&1
if errorlevel 1 goto :sin_dependencias
exit /b 0

:sin_python
echo.
echo ============================================================
echo  NO HAY PYTHON EN ESTE EQUIPO
echo ============================================================
echo.
echo  Lo que hay en el PATH como "python" es el atajo de la
echo  Microsoft Store, que no es Python.
echo.
echo  Tiene dos salidas:
echo.
echo   A) Use los ejecutables, que no necesitan Python:
echo        dist\EstadosFinancieros.exe
echo.
echo   B) Instale el Python portable del proyecto:
echo        powershell -ExecutionPolicy Bypass -File tools\bootstrap_python.ps1
echo.
echo ============================================================
exit /b 1

:sin_dependencias
echo.
echo ============================================================
echo  FALTAN LAS DEPENDENCIAS
echo ============================================================
echo.
echo  Python si esta:
echo    %PY%
echo  pero le faltan openpyxl / python-docx / docxtpl.
echo.
echo  Instalelas con:
echo    "%PY%" -m pip install -r requirements.txt
echo.
echo  O use los ejecutables, que las llevan dentro:
echo    dist\EstadosFinancieros.exe
echo.
echo ============================================================
exit /b 2
