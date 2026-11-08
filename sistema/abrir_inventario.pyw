# -*- coding: utf-8 -*-
"""
Lanzador del Inventario Bajo Soppalco (lo ejecuta el icono del escritorio).
- Si el servidor no está corriendo, lo arranca OCULTO (sin ventana).
- Abre el programa en una ventana tipo aplicación (Edge/Chrome en modo app).
Se ejecuta con pythonw.exe, así que no muestra consola.
"""
import os
import socket
import subprocess
import sys
import time
import webbrowser

BASE = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("INVENTARIO_PORT", "8080"))
URL = f"http://localhost:{PORT}"


def aviso(texto):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, texto, "Inventario Bajo Soppalco", 0x30)
    except Exception:
        print(texto)


def servidor_activo():
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=1):
            return True
    except OSError:
        return False


def arrancar_servidor():
    exe = sys.executable                              # pythonw.exe si se lanzó desde el icono
    log = open(os.path.join(BASE, "servidor.log"), "a", encoding="utf-8")
    subprocess.Popen([exe, "-u", os.path.join(BASE, "servidor_inventario.py"), "--sin-navegador"],
                     cwd=BASE, stdout=log, stderr=log, stdin=subprocess.DEVNULL)
    for _ in range(120):                              # espera hasta 60 s (la primera vez puede tardar)
        time.sleep(0.5)
        if servidor_activo():
            return True
    return False


def abrir_ventana():
    candidatos = [
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    ]
    for exe in candidatos:
        if os.path.isfile(exe):
            subprocess.Popen([exe, f"--app={URL}", "--window-size=1280,900"])
            return
    webbrowser.open(URL)


def main():
    if not servidor_activo():
        if not arrancar_servidor():
            aviso("No se pudo iniciar el programa de inventario.\n\n"
                  f"Revisar el archivo servidor.log en:\n{BASE}")
            return
    abrir_ventana()


if __name__ == "__main__":
    main()
