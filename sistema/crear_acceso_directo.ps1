# Crea los accesos del usuario actual:
#  - Escritorio: abre el programa
#  - Inicio (shell:startup): arranca el servidor oculto al iniciar sesion
param([string]$Sistema, [string]$Pythonw)
$Sistema = $Sistema.TrimEnd('\')
$ws  = New-Object -ComObject WScript.Shell
$ico = (Join-Path $Sistema "inventario.ico") + ",0"

$lnk = $ws.CreateShortcut((Join-Path ([Environment]::GetFolderPath("Desktop")) "Inventario Bajo Soppalco.lnk"))
$lnk.TargetPath = $Pythonw
$lnk.Arguments = '"' + (Join-Path $Sistema "abrir_inventario.pyw") + '"'
$lnk.WorkingDirectory = $Sistema
$lnk.IconLocation = $ico
$lnk.Description = "Inventario Bajo Soppalco"
$lnk.Save()

$ini = $ws.CreateShortcut((Join-Path ([Environment]::GetFolderPath("Startup")) "Inventario Bajo Soppalco.lnk"))
$ini.TargetPath = $Pythonw
$ini.Arguments = '"' + (Join-Path $Sistema "servicio.pyw") + '"'
$ini.WorkingDirectory = $Sistema
$ini.IconLocation = $ico
$ini.Description = "Inventario Bajo Soppalco (servidor)"
$ini.Save()
Write-Host "OK"
