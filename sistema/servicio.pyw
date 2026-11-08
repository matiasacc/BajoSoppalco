# -*- coding: utf-8 -*-
"""
Vigilante del servidor (lo ejecuta el inicio automático de Windows, sin ventana).
Arranca servidor_inventario.py oculto y, si se cae, lo vuelve a levantar a los 10 s.
Registro en sistema/servidor.log
"""
import datetime as dt
import os
import subprocess
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(BASE, "servidor.log")
NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def log(txt):
    with open(LOG, "a", encoding="utf-8") as fh:
        fh.write(f"{dt.datetime.now():%d/%m/%Y %H:%M:%S} - {txt}\n")


PORT = int(os.environ.get("INVENTARIO_PORT", "8080"))
DETENER = os.path.join(BASE, "detener.flag")   # lo crea Desinstalar.bat: no reiniciar


def puerto_ocupado():
    import socket
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=1):
            return True
    except OSError:
        return False


def main():
    if puerto_ocupado() or os.path.exists(DETENER):   # ya hay uno corriendo, o se pidió detener
        return
    log("vigilante iniciado")
    while True:
        with open(LOG, "a", encoding="utf-8") as fh:
            p = subprocess.Popen([sys.executable, "-u", os.path.join(BASE, "servidor_inventario.py"), "--sin-navegador"],
                                 cwd=BASE, stdout=fh, stderr=fh, stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
            code = p.wait()
        if os.path.exists(DETENER):
            log("servidor detenido a pedido, vigilante finalizado")
            return
        log(f"servidor detenido (código {code}), reinicio en 10 s")
        time.sleep(10)
        if puerto_ocupado():          # otro proceso tomó el puerto: este vigilante termina
            log("otro servidor activo, vigilante finalizado")
            return


if __name__ == "__main__":
    main()
