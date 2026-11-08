# -*- coding: utf-8 -*-
"""
Instalación y desinstalación para el usuario actual (lo llaman Instalar.bat y Desinstalar.bat).

  python accesos.py instalar      icono del escritorio + arranque al iniciar sesión, y abre el programa
  python accesos.py desinstalar   cierra el programa y quita esos dos accesos

Los accesos directos se crean con la interfaz estándar de Windows (IShellLink), la misma que
usa el Explorador. No se modifica el registro ni se usan permisos especiales.
Registro: sistema/instalacion.log
"""
import ctypes
import datetime as dt
import os
import socket
import subprocess
import sys
import time
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("INVENTARIO_PORT", "8080"))
NOMBRE = "Inventario Bajo Soppalco.lnk"
LOG = os.path.join(BASE, "instalacion.log")
CLAVE_APAGAR = os.path.join(BASE, "apagar.key")   # la escribe el servidor al arrancar


def paso(txt=""):
    print(txt)
    try:
        with open(LOG, "a", encoding="utf-8") as fh:
            fh.write(txt + "\n")
    except OSError:
        pass


# ---------------------------------------------------------------- carpetas y accesos (Windows)
class GUID(ctypes.Structure):
    _fields_ = [("d1", ctypes.c_uint32), ("d2", ctypes.c_uint16), ("d3", ctypes.c_uint16), ("d4", ctypes.c_ubyte * 8)]

    def __init__(self, texto):
        h = texto.strip("{}").replace("-", "")
        super().__init__(int(h[:8], 16), int(h[8:12], 16), int(h[12:16], 16),
                         (ctypes.c_ubyte * 8)(*bytes.fromhex(h[16:])))


ESCRITORIO = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
INICIO = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"


def carpeta(folderid):
    """Ruta de una carpeta del usuario (respeta el Escritorio redirigido a OneDrive)."""
    p = ctypes.c_wchar_p()
    if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(GUID(folderid)), 0, None, ctypes.byref(p)) != 0:
        raise OSError("No se encontró la carpeta del usuario.")
    ruta = p.value
    ctypes.windll.ole32.CoTaskMemFree(p)
    return ruta


def _metodo(ptr, indice, *argtypes):
    vtbl = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    return ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)(vtbl[indice])


def crear_acceso(destino, programa, argumentos, descripcion):
    ole32 = ctypes.windll.ole32
    ole32.CoInitialize(None)
    link = ctypes.c_void_p()
    hr = ole32.CoCreateInstance(ctypes.byref(GUID("{00021401-0000-0000-C000-000000000046}")), None, 1,
                                ctypes.byref(GUID("{000214F9-0000-0000-C000-000000000046}")), ctypes.byref(link))
    if hr != 0:
        raise OSError(f"No se pudo crear el acceso ({hr & 0xFFFFFFFF:#x}).")
    W = ctypes.c_wchar_p
    try:
        # IShellLinkW: 20 SetPath, 11 SetArguments, 9 SetWorkingDirectory, 7 SetDescription, 17 SetIconLocation
        _metodo(link, 20, W)(link, programa)
        _metodo(link, 11, W)(link, argumentos)
        _metodo(link, 9, W)(link, BASE)
        _metodo(link, 7, W)(link, descripcion)
        _metodo(link, 17, W, ctypes.c_int)(link, os.path.join(BASE, "inventario.ico"), 0)
        pf = ctypes.c_void_p()
        hr = _metodo(link, 0, ctypes.c_void_p, ctypes.c_void_p)(
            link, ctypes.byref(GUID("{0000010B-0000-0000-C000-000000000046}")), ctypes.byref(pf))
        if hr != 0:
            raise OSError(f"No se pudo guardar el acceso ({hr & 0xFFFFFFFF:#x}).")
        try:
            hr = _metodo(pf, 6, W, ctypes.c_int)(pf, destino, 1)        # IPersistFile.Save
            if hr != 0:
                raise OSError(f"No se pudo guardar {destino} ({hr & 0xFFFFFFFF:#x}).")
        finally:
            _metodo(pf, 2)(pf)
    finally:
        _metodo(link, 2)(link)


def pythonw():
    w = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    return w if os.path.isfile(w) else sys.executable


def accesos():
    """Las dos rutas de acceso del programa: escritorio e inicio de sesión."""
    return [os.path.join(carpeta(ESCRITORIO), NOMBRE), os.path.join(carpeta(INICIO), NOMBRE)]


def crear_accesos():
    esc, ini = accesos()
    crear_acceso(esc, pythonw(), f'"{os.path.join(BASE, "abrir_inventario.pyw")}"', "Inventario Bajo Soppalco")
    paso(f"      Escritorio: {esc}")
    crear_acceso(ini, pythonw(), f'"{os.path.join(BASE, "servidor_inventario.py")}" --sin-navegador',
                 "Inventario Bajo Soppalco (servidor)")
    paso(f"      Inicio de sesion: {ini}")


def quitar_accesos():
    for ruta in accesos():
        if os.path.exists(ruta):
            os.remove(ruta)
            paso(f"      Quitado: {ruta}")


# ---------------------------------------------------------------- servidor
def servidor_activo():
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=1):
            return True
    except OSError:
        return False


def detener():
    """Pide al servidor que se cierre en orden (termina de escribir el Excel antes)."""
    if not servidor_activo():
        return True
    try:
        with open(CLAVE_APAGAR, encoding="utf-8") as fh:
            clave = fh.read().strip()
        req = urllib.request.Request(f"http://127.0.0.1:{PORT}/api/apagar", data=b"{}", method="POST",
                                     headers={"X-Apagar": clave})
        urllib.request.urlopen(req, timeout=5).read()
    except Exception:
        pass
    for _ in range(30):
        if not servidor_activo():
            return True
        time.sleep(0.5)
    return False


def ip_local():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))
        return s.getsockname()[0]
    except OSError:
        return None
    finally:
        s.close()


# ---------------------------------------------------------------- acciones
def instalar():
    with open(LOG, "w", encoding="utf-8") as fh:
        fh.write(f"==== {dt.datetime.now():%d/%m/%Y %H:%M} ====\n")
    paso("[1/3] Verificando el programa...")
    sys.path.insert(0, BASE)
    try:
        import cargar_movimientos  # noqa: F401  (incluye openpyxl desde sistema/lib)
    except Exception as e:
        paso(f"      ERROR: {e}")
        return 1
    paso("      OK")

    paso("[2/3] Icono en el escritorio y arranque al iniciar sesion...")
    try:
        detener()                      # si había una versión anterior abierta
        quitar_accesos()
        crear_accesos()
    except Exception as e:
        paso(f"      ERROR: {e}")
        return 1

    paso("[3/3] Abriendo el programa...")
    subprocess.Popen([pythonw(), os.path.join(BASE, "abrir_inventario.pyw")], cwd=BASE)
    ip = ip_local()
    paso()
    paso("  ==============================================================")
    paso("   INSTALACION TERMINADA")
    paso('   En esta PC: icono "Inventario Bajo Soppalco" del escritorio')
    if ip:
        paso(f"   Desde otra PC: http://{ip}:{PORT}")
    paso("   Contrasena inicial de Administracion: soppalco")
    paso("  ==============================================================")
    return 0


def desinstalar():
    paso("[1/2] Cerrando el programa...")
    paso("      OK" if detener() else "      No se pudo cerrar: reiniciar la PC antes de borrar la carpeta.")
    paso("[2/2] Quitando el icono y el arranque al iniciar sesion...")
    try:
        quitar_accesos()
    except Exception as e:
        paso(f"      ERROR: {e}")
        return 1
    paso("      OK")
    paso()
    paso('  Listo. Los datos siguen en la carpeta "datos".')
    return 0


if __name__ == "__main__":
    accion = sys.argv[1] if len(sys.argv) > 1 else ""
    sys.exit(instalar() if accion == "instalar" else desinstalar() if accion == "desinstalar" else 2)
