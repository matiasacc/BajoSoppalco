# Inventario Bajo Soppalco

Programa de control de inventario para el depósito Bajo Soppalco.

- Carga de entradas y salidas desde el navegador, a mano o con lectora de códigos de barras.
- Ingreso por **legajo y contraseña**; cada movimiento queda registrado a nombre de quien ingresó, con fecha y hora automáticas.
- Stock en unidades y en empaques (cajas cerradas + sueltas), alertas de stock mínimo e informe de consumo.
- Administración con contraseña: maestro de insumos, personas, corrección y anulación de movimientos.
- Los datos se guardan en un Excel (`datos/Inventario_Bajo_Soppalco.xlsx`).

Una PC funciona como servidor web interno (HTTP, puerto 8080) y las demás PCs de la red acceden con el navegador, sin instalar nada.

## Instalación

1. Instalar Python 3.8 o superior (Microsoft Store o python.org).
2. Copiar esta carpeta a `C:\Inventario Bajo Soppalco`.
3. Doble clic en `Instalar.bat`.

Instrucciones completas: [`LEEME - Instrucciones.pdf`](LEEME%20-%20Instrucciones.pdf).

## Estructura

| Carpeta / archivo | Contenido |
|---|---|
| `Instalar.bat`, `Desinstalar.bat` | Instalación en la PC servidor |
| `datos/` | Excel con los datos reales (**no se sube al repositorio**) |
| `herramientas/` | Carga por CSV, arranque manual con ventana, regla de firewall para IT |
| `sistema/` | Programa (Python + página web) |
| `docs/` | Fuente del PDF de instrucciones |

## Actualizar una instalación

1. `Desinstalar.bat`.
2. Reemplazar todo **excepto la carpeta `datos`**.
3. `Instalar.bat`.
