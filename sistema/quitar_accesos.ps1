# Quita los accesos directos del programa (icono del escritorio y arranque al iniciar sesion).
# Busca en el escritorio del usuario (incluido OneDrive), el escritorio publico y las carpetas
# de Inicio, y borra todo acceso con el nombre del programa o que apunte a sus archivos.
$nombre = 'Inventario Bajo Soppalco.lnk'
$carpetas = @(
  [Environment]::GetFolderPath('Desktop'),
  [Environment]::GetFolderPath('CommonDesktopDirectory'),
  [Environment]::GetFolderPath('Startup'),
  [Environment]::GetFolderPath('CommonStartup'),
  (Join-Path $env:USERPROFILE 'Desktop'),
  (Join-Path $env:USERPROFILE 'Escritorio')
)
foreach ($od in (Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -Filter 'OneDrive*' -ErrorAction SilentlyContinue)) {
  $carpetas += (Join-Path $od.FullName 'Desktop')
  $carpetas += (Join-Path $od.FullName 'Escritorio')
}
$ws = New-Object -ComObject WScript.Shell
$n = 0
foreach ($c in ($carpetas | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique)) {
  foreach ($f in (Get-ChildItem -LiteralPath $c -Filter '*.lnk' -ErrorAction SilentlyContinue)) {
    $l = $ws.CreateShortcut($f.FullName)
    if ($f.Name -eq $nombre -or $l.Arguments -match 'abrir_inventario\.pyw|servicio\.pyw') {
      Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
      if (Test-Path -LiteralPath $f.FullName) { Write-Host "        No se pudo quitar: $($f.FullName)" }
      else { $n++; Write-Host "        Quitado: $($f.FullName)" }
    }
  }
}
if ($n -eq 0) { Write-Host "        No habia accesos para quitar." }
