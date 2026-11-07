@echo off
setlocal EnableExtensions
REM ==================================================================
REM  INVENTARIO BAJO SOPPALCO - Desinstalacion (doble clic)
REM  No borra los datos (carpeta "datos").
REM ==================================================================
cd /d "%~dp0"
title Inventario Bajo Soppalco - Desinstalacion
echo.
echo  [1/3] Quitando arranque al iniciar sesion e icono del escritorio...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sistema\quitar_accesos.ps1"
echo  [2/3] Deteniendo el programa...
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name like 'python%%'\" | Where-Object { $_.CommandLine -like '*servicio.pyw*' -or $_.CommandLine -like '*servidor_inventario.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
echo  [3/3] Limpiando...
del /q "sistema\python_path.txt" >nul 2>&1
echo.
echo  Listo. Los datos siguen en la carpeta "datos".
echo.
pause
