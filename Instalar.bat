@echo off
REM Inventario Bajo Soppalco - Instalacion para el usuario actual (doble clic).
REM Crea el icono del escritorio y el arranque al iniciar sesion. Registro: sistema\instalacion.log
cd /d "%~dp0"
title Inventario Bajo Soppalco - Instalacion
if not exist "sistema\python\python.exe" (
  echo Falta la carpeta sistema\python: volver a descargar el programa completo.
  pause
  exit /b 1
)
"sistema\python\python.exe" "sistema\accesos.py" instalar
echo.
pause
