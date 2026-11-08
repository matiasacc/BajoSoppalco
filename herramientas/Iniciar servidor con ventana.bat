@echo off
REM Arranque manual del servidor con ventana (diagnostico).
REM Si el servidor ya esta corriendo, solo abre el navegador.
cd /d "%~dp0..\sistema"
title Inventario Bajo Soppalco - servidor
"python\python.exe" servidor_inventario.py
if errorlevel 1 pause
