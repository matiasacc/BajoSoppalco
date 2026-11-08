@echo off
setlocal EnableExtensions
REM ==================================================================
REM  INVENTARIO BAJO SOPPALCO - Desinstalacion (doble clic)
REM  Detiene el programa y quita el icono y el arranque al iniciar sesion.
REM  No borra los datos (carpeta "datos").
REM ==================================================================
cd /d "%~dp0"
title Inventario Bajo Soppalco - Desinstalacion
set "PY="
if exist "sistema\python_path.txt" set /p PY=<"sistema\python_path.txt"
if defined PY if not exist "%PY%" set "PY="
if not defined PY if exist "%~dp0sistema\python\python.exe" set "PY=%~dp0sistema\python\python.exe"
if not defined PY set "PY=python"
echo.
"%PY%" "%~dp0sistema\accesos.py" desinstalar
echo.
echo  Listo. Los datos siguen en la carpeta "datos".
echo.
pause
