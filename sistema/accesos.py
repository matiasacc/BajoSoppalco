# -*- coding: utf-8 -*-
"""
Instalación y desinstalación para el usuario actual (lo llaman Instalar.bat y Desinstalar.bat).

  python accesos.py instalar      icono del escritorio + arranque al iniciar sesión, y abre el programa
  python accesos.py desinstalar   detiene el programa y quita los accesos

Los accesos directos se crean con la interfaz estándar de Windows (IShellLink), la misma que
usa el Explorador. No se modifica el registro ni se usan permisos especiales.
"""
import ctypes
import os
import subprocess
import sys
import time
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("INVENTARIO_PORT", "8080"))
NOMBRE = "Inventario Bajo Soppalco.lnk"
DETENER = os.path.join(BASE, "detener.flag")      # el vigilante no reinicia el servidor mientras exista
CLAVE_APAGAR = os.path.join(BASE, "apagar.key")   # la escribe el servidor al arrancar


# ---------------------------------------------------------------- utilidades de Windows (ctypes)
class GUID(ctypes.Structure):
    _fields_ = [("d1", ctypes.c_uint32), ("d2", ctypes.c_uint16), ("d3", ctypes.c_uint16), ("d4", ctypes.c_ubyte * 8)]

    def __init__(self, texto):
        h = texto.strip("{}").replace("-", "")
        b = bytes.fromhex(h[16:])
        super().__init__(int(h[:8], 16), int(h[8:12], 16), int(h[12:16], 16), (ctypes.c_ubyte * 8)(*b))


FOLDERID = {
    "Desktop": "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}",
    "PublicDesktop": "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}",
    "Startup": "{B97D20BB-F46A-4C97-BA10-5E3608430854}",
    "CommonStartup": "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}",
}


def carpeta(nombre):
    """Ruta real de una carpeta especial (respeta el Escritorio redirigido a OneDrive)."""
    p = ctypes.c_wchar_p()
    try:
        r = ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(GUID(FOLDERID[nombre])), 0, None, ctypes.byref(p))
        if r == 0 and p.value:
            ruta = p.value
            ctypes.windll.ole32.CoTaskMemFree(p)
            return ruta
    except Exception:
        pass
    return None


def _metodo(ptr, indice, *argtypes):
    """Función de la tabla de métodos de un objeto COM."""
    vtbl = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    return ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)(vtbl[indice])


def crear_acceso(destino_lnk, programa, argumentos, carpeta_trabajo, icono, descripcion):
    ole32 = ctypes.windll.ole32
    ole32.CoInitialize(None)
    link = ctypes.c_void_p()
    hr = ole32.CoCreateInstance(ctypes.byref(GUID("{00021401-0000-0000-C000-000000000046}")), None, 1,
                                ctypes.byref(GUID("{000214F9-0000-0000-C000-000000000046}")), ctypes.byref(link))
    if hr != 0:
        raise OSError(f"CoCreateInstance ShellLink: {hr & 0xFFFFFFFF:#x}")
    W = ctypes.c_wchar_p
    try:
        # IShellLinkW: 7 SetDescription, 9 SetWorkingDirectory, 11 SetArguments, 17 SetIconLocation, 20 SetPath
        for idx, val in ((20, programa), (11, argumentos), (9, carpeta_trabajo), (7, descripcion)):
            hr = _metodo(link, idx, W)(link, val)
            if hr != 0:
                raise OSError(f"IShellLinkW[{idx}]: {hr & 0xFFFFFFFF:#x}")
        _metodo(link, 17, W, ctypes.c_int)(link, icono, 0)
        pf = ctypes.c_void_p()
        hr = _metodo(link, 0, ctypes.c_void_p, ctypes.c_void_p)(
            link, ctypes.byref(GUID("{0000010B-0000-0000-C000-000000000046}")), ctypes.byref(pf))   # QueryInterface IPersistFile
        if hr != 0:
            raise OSError(f"IPersistFile: {hr & 0xFFFFFFFF:#x}")
        try:
            hr = _metodo(pf, 6, W, ctypes.c_int)(pf, destino_lnk, 1)                                # IPersistFile.Save
            if hr != 0:
                raise OSError(f"Guardar {destino_lnk}: {hr & 0xFFFFFFFF:#x}")
        finally:
            _metodo(pf, 2)(pf)
    finally:
        _metodo(link, 2)(link)                                                                       # Release


def refrescar_iconos():
    try:
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)   # SHCNE_ASSOCCHANGED
    except Exception:
        pass


# ---------------------------------------------------------------- accesos del programa
def carpetas_accesos():
    home = os.environ.get("USERPROFILE", "")
    cs = [carpeta(n) for n in FOLDERID]
    cs += [os.path.join(home, "Desktop"), os.path.join(home, "Escritorio")]
    try:
        for d in os.listdir(home):
            if d.lower().startswith("onedrive"):
                cs += [os.path.join(home, d, "Desktop"), os.path.join(home, d, "Escritorio")]
    except OSError:
        pass
    vistos, res = set(), []
    for c in cs:
        if c and os.path.isdir(c) and os.path.normcase(c) not in vistos:
            vistos.add(os.path.normcase(c)); res.append(c)
    return res


def es_nuestro(ruta):
    if os.path.basename(ruta).lower() == NOMBRE.lower():
        return True
    try:
        with open(ruta, "rb") as fh:
            datos = fh.read(65536)
    except OSError:
        return False
    return any(t.encode("utf-16-le") in datos or t.encode("latin-1") in datos
               for t in ("abrir_inventario.pyw", "servicio.pyw"))


def quitar_accesos():
    n = 0
    for c in carpetas_accesos():
        for f in os.listdir(c):
            ruta = os.path.join(c, f)
            if f.lower().endswith(".lnk") and es_nuestro(ruta):
                try:
                    os.remove(ruta); n += 1
                    print(f"      Quitado: {ruta}")
                except OSError as e:
                    print(f"      No se pudo quitar {ruta}: {e}")
    if not n:
        print("      No habia accesos para quitar.")
    refrescar_iconos()


def pythonw():
    w = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    return w if os.path.isfile(w) else sys.executable


def crear_accesos():
    pw, ico = pythonw(), os.path.join(BASE, "inventario.ico")
    escritorio = carpeta("Desktop") or os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
    inicio = carpeta("Startup")
    crear_acceso(os.path.join(escritorio, NOMBRE), pw, f'"{os.path.join(BASE, "abrir_inventario.pyw")}"',
                 BASE, ico, "Inventario Bajo Soppalco")
    print(f"      Escritorio: {escritorio}")
    if inicio:
        crear_acceso(os.path.join(inicio, NOMBRE), pw, f'"{os.path.join(BASE, "servicio.pyw")}"',
                     BASE, ico, "Inventario Bajo Soppalco (servidor)")
        print(f"      Inicio de sesion: {inicio}")
    refrescar_iconos()


# ---------------------------------------------------------------- arrancar / detener
def servidor_activo():
    import socket
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=1):
            return True
    except OSError:
        return False


def detener():
    """Pide al servidor que se cierre en orden (termina de escribir el Excel antes)."""
    with open(DETENER, "w") as fh:
        fh.write("1")
    if not servidor_activo():
        print("      No estaba en marcha.")
        return True
    try:
        with open(CLAVE_APAGAR, encoding="utf-8") as fh:
            clave = fh.read().strip()
        req = urllib.request.Request(f"http://127.0.0.1:{PORT}/api/apagar", data=b"{}", method="POST",
                                     headers={"X-Apagar": clave})
        urllib.request.urlopen(req, timeout=5).read()
    except Exception as e:
        print(f"      Aviso: {e}")
    for _ in range(30):
        if not servidor_activo():
            print("      Detenido.")
            return True
        time.sleep(0.5)
    print("      No se pudo detener: reiniciar la PC antes de borrar la carpeta.")
    return False


def arrancar():
    try:
        os.remove(DETENER)
    except OSError:
        pass
    pw = pythonw()
    flags = 0x00000008 | 0x00000200 if os.name == "nt" else 0   # DETACHED_PROCESS | NEW_PROCESS_GROUP
    subprocess.Popen([pw, os.path.join(BASE, "servicio.pyw")], cwd=BASE, creationflags=flags, close_fds=True)
    for _ in range(40):
        if servidor_activo():
            break
        time.sleep(0.5)
    subprocess.Popen([pw, os.path.join(BASE, "abrir_inventario.pyw")], cwd=BASE, creationflags=flags, close_fds=True)


def main():
    accion = sys.argv[1] if len(sys.argv) > 1 else ""
    if accion == "instalar":
        print("      Quitando accesos anteriores...")
        quitar_accesos()
        print("      Creando accesos...")
        try:
            crear_accesos()
        except Exception as e:
            print(f"      ERROR al crear los accesos: {e}")
            return 1
        print("      Iniciando el programa...")
        arrancar()
        return 0
    if accion == "desinstalar":
        print("[1/2] Deteniendo el programa...")
        detener()
        print("[2/2] Quitando el icono y el arranque al iniciar sesion...")
        quitar_accesos()
        for f in ("python_path.txt",):
            try:
                os.remove(os.path.join(BASE, f))
            except OSError:
                pass
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main())
