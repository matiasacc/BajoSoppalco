@echo off
REM Arranque manual del servidor con ventana (diagnostico o uso sin instalar).
REM Si el servidor ya esta corriendo, solo abre el navegador.
cd /d "%~dp0..\sistema"
title Inventario Bajo Soppalco - servidor
set "PY=python"
if exist python_path.txt set /p PY=<python_path.txt
"%PY%" servidor_inventario.py
if errorlevel 1 pause
