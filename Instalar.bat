@echo off
setlocal EnableExtensions
REM ==================================================================
REM  INVENTARIO BAJO SOPPALCO - Instalacion (doble clic)
REM  Instala para el usuario actual: icono y arranque al iniciar sesion.
REM  Registro: sistema\instalacion.log
REM ==================================================================
cd /d "%~dp0"
title Inventario Bajo Soppalco - Instalacion
set "LOG=%~dp0sistema\instalacion.log"
echo ==== %date% %time% ==== > "%LOG%"

call :paso "[1/3] Python..."
set "PY="
REM 1) el Python portatil incluido en la carpeta (no se instala nada)
if exist "%~dp0sistema\python\python.exe" set "PY=%~dp0sistema\python\python.exe"
REM 2) si no esta, uno ya instalado en la PC
if not defined PY for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do if not defined PY set "PY=%%P"
if not defined PY for /f "delims=" %%P in ('where python 2^>nul ^| findstr /v /i "WindowsApps"') do if not defined PY set "PY=%%P"
if not defined PY for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do if not defined PY set "PY=%%P"
if not defined PY (
  call :paso "      ERROR: falta la carpeta sistema\python."
  call :paso "      Volver a descargar el programa completo."
  goto fin
)
call :paso "      %PY%"

call :paso "[2/3] Verificando..."
"%PY%" -c "import sys;sys.path.insert(0,r'%~dp0sistema');import cargar_movimientos" >> "%LOG%" 2>&1
if errorlevel 1 (
  call :paso "      ERROR: el programa no arranca (ver instalacion.log)."
  goto fin
)
> "sistema\python_path.txt" echo %PY%
call :paso "      OK"

call :paso "[3/3] Icono en el escritorio y arranque al iniciar sesion..."
"%PY%" "%~dp0sistema\accesos.py" instalar >> "%LOG%" 2>&1
if errorlevel 1 ( call :paso "      AVISO: no se pudieron crear los accesos (ver instalacion.log)." ) else ( call :paso "      OK" )

echo.
echo  ==============================================================
echo   INSTALACION TERMINADA
echo   En esta PC: icono "Inventario Bajo Soppalco" del escritorio
echo.
for /f "tokens=2 delims=:" %%I in ('ipconfig ^| findstr /c:"IPv4"') do for /f "tokens=*" %%J in ("%%I") do echo   Desde otra PC: http://%%J:8080
echo.
echo   Contrasena inicial de Administracion: soppalco
echo  ==============================================================

:fin
echo.
echo  Registro: sistema\instalacion.log
pause
exit /b

:paso
echo %~1
>> "%LOG%" echo %~1
exit /b
