@echo off
REM Inventario Bajo Soppalco - Instalacion para el usuario actual (doble clic).
REM Crea el icono del escritorio y el arranque al iniciar sesion. Registro: sistema\instalacion.log
REM La primera vez obtiene Python oficial (python.org) y lo deja en sistema\python: no se instala en Windows.
cd /d "%~dp0"
title Inventario Bajo Soppalco - Instalacion
if exist "sistema\python\python.exe" goto instalar

echo Preparando Python (solo la primera vez)...
set "ZIP="
for %%Z in (python-3.*-embed-amd64.zip) do set "ZIP=%%~fZ"
if defined ZIP goto descomprimir
echo   Descargando de python.org...
curl.exe -L --fail --silent --show-error -o "sistema\python.zip" "https://www.python.org/ftp/python/3.14.10/python-3.14.10-embed-amd64.zip"
if errorlevel 1 goto sinpython
set "ZIP=%~dp0sistema\python.zip"

:descomprimir
echo   Descomprimiendo %ZIP%
mkdir "sistema\python" 2>nul
tar -xf "%ZIP%" -C "sistema\python"
if errorlevel 1 goto sinpython
if not exist "sistema\python\python.exe" goto sinpython
if exist "sistema\python\python314._pth" copy /y "sistema\python314_pth.txt" "sistema\python\python314._pth" >nul
if exist "sistema\python.zip" del "sistema\python.zip"
echo   OK
echo.

:instalar
"sistema\python\python.exe" "sistema\accesos.py" instalar
echo.
pause
exit /b

:sinpython
if exist "sistema\python.zip" del "sistema\python.zip"
echo.
echo   No se pudo obtener Python automaticamente (sin internet o descarga bloqueada).
echo   1. En python.org, Downloads, Windows: bajar "Windows embeddable package (64-bit)" de Python 3.
echo      Es un archivo python-3.xx.x-embed-amd64.zip
echo   2. Ponerlo en esta carpeta, al lado de Instalar.bat (sin descomprimir).
echo   3. Volver a ejecutar Instalar.bat.
echo.
pause
exit /b 1
