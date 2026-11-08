@echo off
REM Inventario Bajo Soppalco - Desinstalacion (doble clic).
REM Cierra el programa y quita el icono y el arranque al iniciar sesion. No borra la carpeta "datos".
cd /d "%~dp0"
title Inventario Bajo Soppalco - Desinstalacion
"sistema\python\python.exe" "sistema\accesos.py" desinstalar
echo.
pause
