@echo off
REM Carga herramientas\movimientos.csv al Excel (carga alternativa sin navegador).
REM Formato: fecha;tipo;codigo;cantidad;unidad;persona;obs   (tipo E o S)
cd /d "%~dp0..\sistema"
title Inventario Bajo Soppalco - carga por CSV
"python\python.exe" cargar_movimientos.py
echo.
pause
