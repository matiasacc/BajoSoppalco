@echo off
REM ==================================================================
REM  Para IT, solo si otra PC no puede entrar al programa.
REM  Ejecutar con clic derecho > Ejecutar como administrador.
REM  Crea una regla de ENTRADA en el Firewall de Windows:
REM    "Inventario Bajo Soppalco" - TCP 8080 - perfiles Dominio y Privado
REM  Para quitarla:
REM    netsh advfirewall firewall delete rule name="Inventario Bajo Soppalco"
REM ==================================================================
net session >nul 2>&1
if errorlevel 1 ( echo. & echo  Ejecutar con clic derecho ^> Ejecutar como administrador. & echo. & pause & exit /b 1 )
netsh advfirewall firewall delete rule name="Inventario Bajo Soppalco" >nul 2>&1
netsh advfirewall firewall add rule name="Inventario Bajo Soppalco" dir=in action=allow protocol=TCP localport=8080 profile=domain,private
echo.
echo  Regla creada: TCP 8080 entrante (Dominio / Privado).
pause
