# Inventario Bajo Soppalco

Programa de control de inventario para el depósito Bajo Soppalco.

- Carga de entradas y salidas desde el navegador: tocando la foto del insumo, buscándolo o con lectora de códigos de barras.
- Ingreso por **legajo y contraseña**; cada movimiento queda registrado a nombre de quien ingresó, con fecha y hora automáticas.
- Stock en unidades y en empaques (cajas cerradas + sueltas), alertas de stock mínimo e informe de consumo.
- Administración con contraseña: maestro de insumos, personas, corrección y anulación de movimientos.
- Los datos se guardan en un Excel (`datos/Inventario_Bajo_Soppalco.xlsx`).

Una PC funciona como servidor web interno (HTTP, puerto 8080) y las demás PCs de la red acceden con el navegador, sin instalar nada.

## Instalación

1. Copiar esta carpeta a `C:\Inventario Bajo Soppalco`.
2. Doble clic en `Instalar.bat`.

No hace falta instalar Python: el programa trae uno portátil (`sistema/python`) y la librería de Excel (`sistema/lib`).

La primera vez el programa crea la planilla con **datos de ejemplo** para probar: 4 insumos y 2 usuarios
(legajos `1001` y `1002`, contraseña `1234`). Se borran desde Administración antes de empezar a usarlo.
Contraseña inicial de Administración: `soppalco` (cambiarla).

Instrucciones completas: [`LEEME - Instrucciones.pdf`](LEEME%20-%20Instrucciones.pdf).

## Estructura

| Carpeta / archivo | Contenido |
|---|---|
| `Instalar.bat`, `Desinstalar.bat` | Instalación en la PC servidor |
| `datos/` | Excel con los datos reales (**no se sube al repositorio**) |
| `herramientas/` | Carga por CSV, arranque manual con ventana, regla de firewall para IT |
| `sistema/` | Programa (Python + página web); `python/` portátil y `lib/` con openpyxl |
| `docs/` | Fuente del PDF de instrucciones |

## Actualizar una instalación

1. `Desinstalar.bat`.
2. Reemplazar todo **excepto la carpeta `datos`**.
3. `Instalar.bat`.
