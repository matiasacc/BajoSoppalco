# -*- coding: utf-8 -*-
"""
INVENTARIO BAJO SOPPALCO - servidor de carga (web local)
=========================================================
Arranque:   automático (Instalar.bat) o "herramientas/Iniciar servidor con ventana.bat"
Después abrir en el navegador la dirección que muestra la consola:
    - en la misma PC:            http://localhost:8080
    - desde otra PC / celular:   http://<IP de esta PC>:8080   (misma red)

Todo se guarda en datos/Inventario_Bajo_Soppalco.xlsx.
Si el Excel está abierto y no se puede guardar, los movimientos quedan en
datos/pendientes.json y se pasan solos al Excel apenas se cierre.

Contraseña inicial de administración: soppalco   (cambiarla en Administración > Cuenta)

Requiere: Python 3.8+ y openpyxl.
"""
import datetime as dt
import hashlib
import json
import os
import re
import secrets
import shutil
import socket
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))   # Python portátil: no agrega la carpeta del script
import cargar_movimientos as cm
import openpyxl

BASE = cm.BASE
CONFIG = os.path.join(BASE, "config_inventario.json")
PEND = os.path.join(cm.DATOS, "pendientes.json")
APP = os.path.join(BASE, "app.html")
PORT = int(os.environ.get("INVENTARIO_PORT", "8080"))
# Por defecto atiende a toda la red interna (como siempre). Si existe sistema/solo_local.on
# atiende solo a esta PC (crear ese archivo vacío a mano si se quisiera restringir).
ACCESO_RED = not os.path.exists(os.path.join(BASE, "solo_local.on"))
HOST = "0.0.0.0" if ACCESO_RED else "127.0.0.1"
LOCK = threading.RLock()
INACTIVIDAD = int(os.environ.get("INVENTARIO_INACTIVIDAD", "300"))  # seg. sin actividad -> se cierra la sesión
INACT_ADMIN = int(os.environ.get("INVENTARIO_INACTIVIDAD_ADMIN", "120"))  # Administración: más estricto
TOKENS = {}                 # sesiones de administración: token -> última actividad (epoch)
SES = {}                    # sesiones de personas: token -> {legajo, nombre, cambiar, ultimo}
FALLOS = {}                 # legajo -> [intentos fallidos, bloqueado hasta (epoch)]
MAX_FALLOS, BLOQUEO = 5, 5 * 60
CLAVE_INICIAL = "soppalco"


# ---------------------------------------------------------------- configuración / contraseña
def _hash(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), 120000).hex()


def load_config():
    if not os.path.exists(CONFIG):
        salt = secrets.token_hex(16)
        save_config({"salt": salt, "hash": _hash(CLAVE_INICIAL, salt), "clave_inicial": True})
    with open(CONFIG, encoding="utf-8") as fh:
        return json.load(fh)


def save_config(cfg):
    with open(CONFIG, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=1)


def check_password(pw):
    cfg = load_config()
    return secrets.compare_digest(_hash(pw, cfg["salt"]), cfg["hash"])


def token_ok(tok):
    """Sesión de administración válida (y renueva la actividad)."""
    t = TOKENS.get(tok or "")
    if t and time.time() - t < INACT_ADMIN:
        TOKENS[tok] = time.time()
        return True
    TOKENS.pop(tok or "", None)
    return False


def sesion(tok):
    """Sesión de persona válida (y renueva la actividad). Devuelve el dict o None."""
    s = SES.get(tok or "")
    if s and time.time() - s["ultimo"] < INACTIVIDAD:
        s["ultimo"] = time.time()
        return s
    SES.pop(tok or "", None)
    return None


def cerrar_sesiones(legajo):
    for t in [t for t, s in SES.items() if s["legajo"] == legajo]:
        SES.pop(t, None)


def norm_leg(x):
    x = cm.norm(x)
    return x[:-2] if x.endswith(".0") else x


# ---------------------------------------------------------------- pendientes (Excel abierto)
def read_pend():
    if not os.path.exists(PEND):
        return []
    with open(PEND, encoding="utf-8") as fh:
        return json.load(fh)


def write_pend(lst):
    if lst:
        with open(PEND, "w", encoding="utf-8") as fh:
            json.dump(lst, fh, ensure_ascii=False, indent=1)
    elif os.path.exists(PEND):
        os.remove(PEND)


def append_movs(wb, movs):
    ws = wb[cm.SH_MOV]
    r = cm.last_row(ws)
    for m in movs:
        r += 1
        f = m["fecha"] if isinstance(m["fecha"], dt.date) else dt.date.fromisoformat(m["fecha"])
        cm.write_movement(ws, r, f, m["tipo"], m["cod"], m["cant"], m["unidad"], m["persona"],
                          m["origen"], m["obs"] or None)
    ws.auto_filter.ref = f"A1:J{r}"


def save_wb(wb):
    """Guarda; si el Excel está abierto (bloqueado) lanza PermissionError."""
    cm.guardar(wb)


def regenerar_html(wb=None):
    try:
        wb = wb or openpyxl.load_workbook(cm.XLSX)
        items = cm.read_maestro(wb)
        movs = cm.read_movements(wb, items)
        cm.build_html(items, movs, cm.compute_stock(items, movs), cm.HTML_OUT)
    except Exception as e:                     # el HTML estático es secundario
        print("Aviso: no se pudo regenerar el HTML:", e)


def flush_pend():
    """Intenta pasar los pendientes al Excel. Devuelve cuántos quedan."""
    with LOCK:
        pend = read_pend()
        if not pend:
            return 0
        wb = openpyxl.load_workbook(cm.XLSX)
        items = cm.read_maestro(wb)
        pend_ok = [m for m in pend if m["cod"] in items]
        append_movs(wb, pend_ok)
        cm.rebuild_stock(wb)
        try:
            save_wb(wb)
        except PermissionError:
            return len(pend)
        write_pend([m for m in pend if m["cod"] not in items])
        print(f"Se pasaron {len(pend_ok)} movimientos pendientes al Excel.")
        regenerar_html(wb)
        return len(read_pend())


def snapshot():
    """Lee Excel + pendientes y calcula stock."""
    with LOCK:
        npend = flush_pend()
        wb = openpyxl.load_workbook(cm.XLSX)
        items = cm.read_maestro(wb)
        movs = cm.read_movements(wb, items)
        for m in read_pend():
            if m["cod"] in items:
                movs.append(dict(fecha=dt.date.fromisoformat(m["fecha"]), tipo=m["tipo"], cod=m["cod"],
                                 cant=m["cant"], unidad=m["unidad"],
                                 base=cm.base_qty(items[m["cod"]], m["cant"], m["unidad"]),
                                 persona=m["persona"], origen=m["origen"] + " (pendiente)",
                                 obs=m["obs"], fila=None))
        st = cm.compute_stock(items, movs)
        return wb, items, movs, st, npend


# ---------------------------------------------------------------- fotos de productos
FOTOS = os.path.join(cm.DATOS, "fotos")
RE_COD = re.compile(r"^[A-Z0-9][A-Z0-9\-]{0,11}$")


def foto_path(cod):
    return os.path.join(FOTOS, f"{cod}.jpg")


def foto_version(cod):
    """0 si no tiene foto; si tiene, su fecha (sirve para refrescar la imagen en el navegador)."""
    p = foto_path(cod)
    return int(os.path.getmtime(p)) if os.path.isfile(p) else 0


def api_admin_foto(body):
    import base64
    cod = cm.norm(body.get("cod")).upper()
    if not RE_COD.match(cod):
        return 400, dict(error="Código inválido.")
    with LOCK:
        items = cm.read_maestro(openpyxl.load_workbook(cm.XLSX))
    if cod not in items:
        return 400, dict(error=f"No existe {cod}.")
    if body.get("borrar"):
        try:
            os.remove(foto_path(cod))
        except OSError:
            pass
        return 200, dict(ok=True, msg=f"Foto de {cod} quitada.")
    data = cm.norm(body.get("data"))
    if not data.startswith("data:image/jpeg;base64,"):
        return 400, dict(error="Formato de imagen no válido.")
    try:
        raw = base64.b64decode(data.split(",", 1)[1], validate=True)
    except Exception:
        return 400, dict(error="Imagen dañada.")
    if raw[:3] != b"\xff\xd8\xff" or len(raw) > 1_500_000:
        return 400, dict(error="La imagen no es un JPG válido o es demasiado grande.")
    os.makedirs(FOTOS, exist_ok=True)
    with open(foto_path(cod), "wb") as fh:
        fh.write(raw)
    return 200, dict(ok=True, msg=f"Foto de {cod} guardada.")


# ---------------------------------------------------------------- copias de seguridad
BACKUPS = os.path.join(cm.DATOS, "backups")
MAX_AUTO, MAX_MANUAL = 30, 10
RE_BACKUP = re.compile(r"^Inventario_\d{4}-\d{2}-\d{2}(_\d{4}_manual)?\.xlsx$")


def lista_backups():
    if not os.path.isdir(BACKUPS):
        return []
    out = []
    for f in os.listdir(BACKUPS):
        if RE_BACKUP.match(f):
            p = os.path.join(BACKUPS, f)
            out.append(dict(nombre=f, fecha=dt.datetime.fromtimestamp(os.path.getmtime(p)).strftime("%d/%m/%Y %H:%M"),
                            kb=round(os.path.getsize(p) / 1024), manual="_manual" in f, t=os.path.getmtime(p)))
    return sorted(out, key=lambda x: -x["t"])


def hacer_backup(manual=False):
    """Copia el Excel a datos/backups. Automática: una por día. Manual: con hora."""
    with LOCK:
        if not os.path.exists(cm.XLSX):
            return None
        os.makedirs(BACKUPS, exist_ok=True)
        ahora = dt.datetime.now()
        nombre = f"Inventario_{ahora:%Y-%m-%d}" + (f"_{ahora:%H%M}_manual" if manual else "") + ".xlsx"
        destino = os.path.join(BACKUPS, nombre)
        if manual or not os.path.exists(destino):
            shutil.copy(cm.XLSX, destino)       # copy (no copy2): la fecha de la copia es la de hoy
    for tipo, maximo in ((False, MAX_AUTO), (True, MAX_MANUAL)):
        for b in [b for b in lista_backups() if b["manual"] == tipo][maximo:]:
            try:
                os.remove(os.path.join(BACKUPS, b["nombre"]))
            except OSError:
                pass
    return nombre


def ciclo_backup():
    while True:
        try:
            hacer_backup()
        except Exception as e:
            print("Aviso: no se pudo hacer la copia de seguridad:", e)
        time.sleep(3600)


def api_admin_backups_get():
    return 200, dict(backups=[{k: v for k, v in b.items() if k != "t"} for b in lista_backups()])


def api_admin_backup_post():
    n = hacer_backup(manual=True)
    return (200, dict(ok=True, msg=f"Copia guardada: {n}")) if n else (400, dict(error="No hay planilla para copiar."))


# ---------------------------------------------------------------- API
def api_datos():
    wb, items, movs, st, npend = snapshot()
    lst = []
    for c in sorted(items):
        it, s = items[c], st[c]
        lst.append(dict(cod=c, ins=it["ins"], sec=it["sec"], ubi=it["ubi"], ub=it["ub"], emp=it["emp"],
                        fac=it["fac"], emp2=it["emp2"], fac2=it["fac2"], bar=it["bar"],
                        ini=it["ini"] if not isinstance(it["ini"], str) else None,
                        act=None if s["falta"] else s["act"], estado=s["estado"], min=it["min"],
                        foto=foto_version(c)))
    ult = []
    for m in list(reversed(movs))[:15]:
        ult.append(dict(fecha=m["fecha"].strftime("%d/%m/%Y"), tipo=m["tipo"], cod=m["cod"],
                        ins=items[m["cod"]]["ins"], cant=m["cant"], unidad=m["unidad"],
                        persona=m["persona"], origen=m["origen"]))
    return 200, dict(items=lst, ultimos=ult, pendientes=npend,
                     clave_inicial=load_config().get("clave_inicial", False))


def api_movimientos(body, ses):
    if not ses:
        return 401, dict(error="La sesión se cerró. Volver a ingresar con legajo y contraseña.")
    if ses["cambiar"]:
        return 403, dict(error="Primero hay que cambiar la contraseña.")
    lineas = body.get("lineas") or []
    if not lineas:
        return 400, dict(error="No hay movimientos en la lista.")
    if len(lineas) > 300:
        return 400, dict(error="Máximo 300 movimientos por carga.")
    ahora = dt.datetime.now()
    hoy = ahora.strftime("%d/%m/%Y")          # la fecha la pone el servidor, no el usuario
    rows = []
    for i, l in enumerate(lineas, start=1):
        rows.append((i, [hoy] + [cm.norm(l.get(k)) for k in ("tipo", "cod", "cant", "unidad")] + [ses["nombre"], cm.norm(l.get("obs"))]))
    with LOCK:
        flush_pend()
        wb = openpyxl.load_workbook(cm.XLSX)
        items = cm.read_maestro(wb)
        ok, err = cm.validate(rows, items, ahora.date())
        if err:
            return 400, dict(error="Hay errores, no se cargó nada.", detalle=sorted({e.strip() for e in err}))
        origen = f"Web {ahora:%d/%m/%Y %H:%M} · legajo {ses['legajo']}"
        for m in ok:
            m["origen"] = origen
        append_movs(wb, ok)
        cm.rebuild_stock(wb)
        try:
            save_wb(wb)
        except PermissionError:
            pend = read_pend()
            for m in ok:
                d = dict(m)
                d["fecha"] = m["fecha"].isoformat()
                pend.append(d)
            write_pend(pend)
            return 200, dict(ok=len(ok), pendiente=True)
        regenerar_html(wb)
        return 200, dict(ok=len(ok), pendiente=False)


def api_ingresar(body):
    leg = norm_leg(body.get("legajo"))
    clave = cm.norm(body.get("clave"))
    if not leg or not clave:
        return 400, dict(error="Completar legajo y contraseña.")
    f = FALLOS.get(leg, [0, 0])
    if f[1] > time.time():
        mins = int((f[1] - time.time()) // 60) + 1
        return 429, dict(error=f"Demasiados intentos. Probar de nuevo en {mins} min.")
    with LOCK:
        wb = openpyxl.load_workbook(cm.XLSX)
        per = {norm_leg(p["legajo"]): p for p in cm.read_personas(wb) if p["legajo"]}
    p = per.get(leg)
    if p and not p["clave"]:
        return 403, dict(error="Ese legajo no tiene contraseña asignada. Pedirla a un administrador.")
    if not p or not cm.clave_ok(clave, p["clave"]):
        f = [f[0] + 1, 0]
        if f[0] >= MAX_FALLOS:
            f = [0, time.time() + BLOQUEO]
        FALLOS[leg] = f
        time.sleep(1.0)
        return 401, dict(error="Legajo o contraseña incorrectos.")
    FALLOS.pop(leg, None)
    tok = secrets.token_urlsafe(24)
    SES[tok] = dict(legajo=leg, nombre=p["nombre"], cambiar=p["cambiar"], ultimo=time.time())
    return 200, dict(token=tok, nombre=p["nombre"], legajo=leg, cambiar=p["cambiar"], inactividad=INACTIVIDAD)


def api_mi_clave(body, ses):
    if not ses:
        return 401, dict(error="La sesión se cerró. Volver a ingresar.")
    nueva = cm.norm(body.get("nueva"))
    if len(nueva) < 4:
        return 400, dict(error="La contraseña debe tener al menos 4 caracteres.")
    with LOCK:
        wb = openpyxl.load_workbook(cm.XLSX)
        ws = cm.personas_sheet(wb)
        p = next((x for x in cm.read_personas(wb) if norm_leg(x["legajo"]) == ses["legajo"]), None)
        if not p:
            return 400, dict(error="Legajo no encontrado.")
        if not ses["cambiar"] and not cm.clave_ok(cm.norm(body.get("actual")), p["clave"]):
            return 400, dict(error="La contraseña actual no es correcta.")
        if cm.clave_ok(nueva, p["clave"]):
            return 400, dict(error="La nueva contraseña tiene que ser distinta de la anterior.")
        ws.cell(row=p["fila"], column=cm.P_CLA).value = cm.hash_clave(nueva)
        ws.cell(row=p["fila"], column=cm.P_CAM).value = None
        try:
            save_wb(wb)
        except PermissionError:
            return 409, dict(error="El Excel está abierto en la PC del depósito: cerralo y reintentá.")
    ses["cambiar"] = False
    return 200, dict(ok=True)


def api_login(body):
    if check_password(cm.norm(body.get("password"))):
        tok = secrets.token_urlsafe(24)
        TOKENS[tok] = time.time()
        return 200, dict(token=tok, clave_inicial=load_config().get("clave_inicial", False), inactividad=INACT_ADMIN)
    time.sleep(1.0)
    return 401, dict(error="Contraseña incorrecta.")


def api_password(body):
    if not check_password(cm.norm(body.get("actual"))):
        return 400, dict(error="La contraseña actual no es correcta.")
    nueva = cm.norm(body.get("nueva"))
    if len(nueva) < 6:
        return 400, dict(error="La nueva contraseña debe tener al menos 6 caracteres.")
    cfg = load_config()
    cfg["salt"] = secrets.token_hex(16)
    cfg["hash"] = _hash(nueva, cfg["salt"])
    cfg["clave_inicial"] = False
    save_config(cfg)
    return 200, dict(ok=True)


def api_admin_maestro_get():
    _, items, movs, st, _ = snapshot()
    nmov = {}
    for m in movs:
        nmov[m["cod"]] = nmov.get(m["cod"], 0) + 1
    out = []
    for c in sorted(items):
        it, s = items[c], st[c]
        out.append(dict(cod=c, ins=it["ins"], sec=it["sec"], ubi=it["ubi"], ub=it["ub"], emp=it["emp"],
                        fac=it["fac"], emp2=it["emp2"], fac2=it["fac2"], bar=it["bar"], sie=it["sie"], sis=it["sis"], min=it["min"], obs=it["obs"],
                        act=None if s["falta"] else s["act"], estado=s["estado"], nmov=nmov.get(c, 0),
                        foto=foto_version(c)))
    return 200, dict(items=out)


CAMPOS = {"ins": cm.M_INS, "sec": cm.M_SEC, "ubi": cm.M_UBI, "ub": cm.M_UB, "emp": cm.M_EMP,
          "fac": cm.M_FAC, "emp2": cm.M_EMP2, "fac2": cm.M_FAC2, "sie": cm.M_SIE, "sis": cm.M_SIS,
          "min": cm.M_MIN, "obs": cm.M_OBS, "bar": cm.M_BAR}
NUMERICOS = {"fac", "fac2", "sie", "sis", "min"}


def _limpiar_campos(campos):
    out, err = {}, []
    for k, v in campos.items():
        if k not in CAMPOS:
            continue
        if k in NUMERICOS:
            if v in (None, ""):
                out[k] = None
            else:
                n = cm.num(v)
                if n is None or n < 0:
                    err.append(f"{k}: '{v}' no es un número válido")
                elif k in ("fac", "fac2") and n == 0:
                    err.append("Cant. por empaque no puede ser 0 (dejar vacío si no se sabe)")
                else:
                    out[k] = int(n) if float(n).is_integer() else n
        else:
            out[k] = cm.norm(v) or None
    if "ins" in out and not out["ins"]:
        err.append("El nombre del insumo no puede quedar vacío")
    if "ub" in out and not out["ub"]:
        err.append("La unidad base no puede quedar vacía")
    nombres = [out.get(k) for k in ("ub", "emp", "emp2") if out.get(k)]
    if len({x.lower() for x in nombres}) < len(nombres):
        err.append("Unidad base, empaque 1 y empaque 2 deben tener nombres distintos")
    return out, err


def _reescribir_formulas_maestro(mae, desde):
    for r in range(desde, cm.last_row(mae) + 1):
        mae.cell(row=r, column=cm.M_SIT).value = cm.maestro_total_formula(r)


def _reescribir_formulas_mov(ws, desde):
    for r in range(desde, cm.last_row(ws) + 1):
        ins_f, base_f = cm.mov_row_formulas(r)
        ws.cell(row=r, column=4).value = ins_f
        ws.cell(row=r, column=7).value = base_f


def api_admin_maestro_post(body):
    with LOCK:
        flush_pend()
        if read_pend():
            return 409, dict(error="Hay movimientos pendientes porque el Excel está abierto. Cerrá el Excel y reintentá.")
        wb = openpyxl.load_workbook(cm.XLSX)
        mae, mov = wb[cm.SH_MAE], wb[cm.SH_MOV]
        items = cm.read_maestro(wb)
        accion = body.get("accion")
        msg = ""
        campos_in = body.get("campos") or {}
        if cm.norm(campos_in.get("bar")):
            nb = cm.norm_bar(campos_in["bar"])
            cod_actual = cm.norm(body.get("cod")).upper()
            for c2, it2 in items.items():
                if c2 != cod_actual and nb in (cm.norm_bar(it2["bar"]), cm.norm_bar(c2)):
                    return 400, dict(error=f"El código de barras '{campos_in['bar']}' ya corresponde a {c2}.")
        if accion == "editar":
            cod = cm.norm(body.get("cod")).upper()
            if cod not in items:
                return 400, dict(error=f"No existe {cod}")
            campos, err = _limpiar_campos(body.get("campos") or {})
            if err:
                return 400, dict(error="; ".join(err))
            r, it = items[cod]["fila"], items[cod]
            # si cambia el nombre de la unidad base o del empaque, se actualizan los movimientos
            ren = {}
            if "ub" in campos and campos["ub"] and it["ub"] and campos["ub"] != it["ub"]:
                ren[it["ub"].lower()] = campos["ub"]
            for k in ("emp", "emp2"):
                if k in campos and it[k] and campos[k] and campos[k] != it[k]:
                    ren[it[k].lower()] = campos[k]
            if ren:
                for rr in range(2, cm.last_row(mov) + 1):
                    if cm.norm(mov.cell(row=rr, column=3).value).upper() == cod:
                        u = cm.norm(mov.cell(row=rr, column=6).value).lower()
                        if u in ren:
                            mov.cell(row=rr, column=6).value = ren[u]
            for k, v in campos.items():
                mae.cell(row=r, column=CAMPOS[k]).value = v
            msg = f"{cod} actualizado."
        elif accion == "nuevo":
            cod = cm.norm(body.get("cod")).upper()
            if not re.fullmatch(r"[A-Z0-9][A-Z0-9\-]{1,11}", cod):
                return 400, dict(error="Código inválido (letras, números y guion; ej. CIN-02).")
            if cod in items:
                return 400, dict(error=f"El código {cod} ya existe.")
            campos, err = _limpiar_campos(body.get("campos") or {})
            if not campos.get("ins"):
                err.append("Falta el nombre del insumo")
            if not campos.get("ub"):
                err.append("Falta la unidad base")
            if err:
                return 400, dict(error="; ".join(err))
            r = cm.last_row(mae) + 1
            mae.cell(row=r, column=cm.M_COD).value = cod
            for k, v in campos.items():
                mae.cell(row=r, column=CAMPOS[k]).value = v
            mae.cell(row=r, column=cm.M_SIT).value = cm.maestro_total_formula(r)
            mae.cell(row=r, column=cm.M_ORI).value = f"Alta web {dt.date.today():%d/%m/%Y}"
            cm.style_maestro_row(mae, r)
            mae.auto_filter.ref = f"A1:P{r}"
            msg = f"{cod} agregado."
        elif accion == "borrar":
            cod = cm.norm(body.get("cod")).upper()
            if cod not in items:
                return 400, dict(error=f"No existe {cod}")
            if any(cm.norm(mov.cell(row=rr, column=3).value).upper() == cod for rr in range(2, cm.last_row(mov) + 1)):
                return 400, dict(error=f"{cod} tiene movimientos: no se puede borrar.")
            r = items[cod]["fila"]
            mae.delete_rows(r)
            _reescribir_formulas_maestro(mae, r)
            try:
                os.remove(foto_path(cod))
            except OSError:
                pass
            msg = f"{cod} borrado."
        else:
            return 400, dict(error="Acción desconocida")
        cm.rebuild_stock(wb)
        try:
            save_wb(wb)
        except PermissionError:
            return 409, dict(error="El Excel está abierto: cerralo y volvé a guardar.")
        regenerar_html(wb)
        return 200, dict(ok=True, msg=msg)


def api_admin_movs_get():
    _, items, movs, _, _ = snapshot()
    out = []
    for m in reversed(movs):
        out.append(dict(fila=m["fila"], fecha=m["fecha"].strftime("%d/%m/%Y"), tipo=m["tipo"], cod=m["cod"],
                        ins=items[m["cod"]]["ins"], cant=m["cant"], unidad=m["unidad"], persona=m["persona"],
                        origen=m["origen"], obs=m["obs"]))
    return 200, dict(movs=out)


def api_admin_personas_get():
    with LOCK:
        wb = openpyxl.load_workbook(cm.XLSX)
        per = cm.read_personas(wb, solo_activas=False)
        _, items, movs, _, _ = snapshot()
    usos = {}
    for m in movs:
        usos[m["persona"].lower()] = usos.get(m["persona"].lower(), 0) + 1
    out = []
    for p in per:
        out.append(dict(fila=p["fila"], legajo=p["legajo"], nombre=p["nombre"], activo=p["activo"], obs=p["obs"],
                        tiene_clave=bool(p["clave"]), cambiar=p["cambiar"], usos=usos.get(p["nombre"].lower(), 0)))
    return 200, dict(personas=sorted(out, key=lambda p: p["nombre"].lower()))


def api_admin_personas_post(body):
    with LOCK:
        flush_pend()
        wb = openpyxl.load_workbook(cm.XLSX)
        ws = cm.personas_sheet(wb)
        per = cm.read_personas(wb, solo_activas=False)
        accion = body.get("accion")
        fila_nueva = None
        nombre = re.sub(r"\s+", " ", cm.norm(body.get("nombre")))
        leg = norm_leg(body.get("legajo"))
        clave = cm.norm(body.get("clave"))
        obs = cm.norm(body.get("obs"))

        def chequeos(fila_propia=None):
            if not nombre or len(nombre) > 60:
                return "Nombre inválido."
            if not re.fullmatch(r"[A-Za-z0-9\-]{1,12}", leg or ""):
                return "Legajo inválido (letras/números, hasta 12)."
            for x in per:
                if x["fila"] == fila_propia:
                    continue
                if norm_leg(x["legajo"]).lower() == leg.lower():
                    return f"El legajo {leg} ya corresponde a {x['nombre']}."
                if x["nombre"].lower() == nombre.lower():
                    return f"'{nombre}' ya está en la lista."
            return None

        if accion == "nuevo":
            e = chequeos()
            if e:
                return 400, dict(error=e)
            fila_nueva = cm.last_row(ws, cm.P_NOM) + 1
            cm.write_persona(ws, fila_nueva, leg, nombre, True, obs)
            msg = f"{nombre} (legajo {leg}) agregado. Falta que defina su contraseña."
        else:
            try:
                fila = int(body.get("fila"))
            except (TypeError, ValueError):
                return 400, dict(error="Fila inválida.")
            p = next((x for x in per if x["fila"] == fila), None)
            if not p or p["nombre"].lower() != cm.norm(body.get("nombre_actual")).lower():
                return 409, dict(error="La lista cambió. Actualizá y reintentá.")
            if accion == "editar":
                e = chequeos(fila)
                if e:
                    return 400, dict(error=e)
                ws.cell(row=fila, column=cm.P_LEG).value = leg
                ws.cell(row=fila, column=cm.P_NOM).value = nombre
                ws.cell(row=fila, column=cm.P_OBS).value = obs or None
                if norm_leg(p["legajo"]) != leg:
                    cerrar_sesiones(norm_leg(p["legajo"]))
                msg = f"{nombre} actualizado."
            elif accion == "clave":
                # la escribe la propia persona en la pantalla de Administración (campos ocultos)
                if len(clave) < 4:
                    return 400, dict(error="La contraseña debe tener al menos 4 caracteres.")
                ws.cell(row=fila, column=cm.P_CLA).value = cm.hash_clave(clave)
                ws.cell(row=fila, column=cm.P_CAM).value = None
                cerrar_sesiones(norm_leg(p["legajo"]))
                FALLOS.pop(norm_leg(p["legajo"]), None)
                msg = f"Contraseña de {p['nombre']} guardada."
            elif accion in ("activar", "desactivar"):
                ws.cell(row=fila, column=cm.P_ACT).value = "Sí" if accion == "activar" else "No"
                if accion == "desactivar":
                    cerrar_sesiones(norm_leg(p["legajo"]))
                msg = f"{p['nombre']} {'activado' if accion == 'activar' else 'desactivado'}."
            elif accion == "borrar":
                # se quita el acceso; los movimientos conservan el nombre en el historial
                ws.delete_rows(fila)
                cerrar_sesiones(norm_leg(p["legajo"]))
                msg = f"{p['nombre']} borrado. Sus movimientos se conservan en el historial."
            else:
                return 400, dict(error="Acción desconocida")
        try:
            save_wb(wb)
        except PermissionError:
            return 409, dict(error="El Excel está abierto: cerralo y reintentá.")
        return 200, dict(ok=True, msg=msg, fila=fila_nueva if accion == "nuevo" else None)


def api_admin_editar_mov(body):
    with LOCK:
        flush_pend()
        wb = openpyxl.load_workbook(cm.XLSX)
        ws = wb[cm.SH_MOV]
        items = cm.read_maestro(wb)
        try:
            r = int(body.get("fila"))
        except (TypeError, ValueError):
            return 400, dict(error="Fila inválida (los pendientes no se pueden editar hasta que pasen al Excel).")
        if r < 2 or r > cm.last_row(ws):
            return 400, dict(error="Fila inválida.")
        cod = cm.norm(ws.cell(row=r, column=3).value).upper()
        if cod != cm.norm(body.get("cod")).upper() or cm.num(ws.cell(row=r, column=5).value) != cm.num(body.get("cant_orig")):
            return 409, dict(error="El Excel cambió desde que se abrió la lista. Actualizá y reintentá.")
        c = body.get("campos") or {}
        row = [c.get("fecha"), c.get("tipo"), cod, str(c.get("cant", "")), c.get("unidad"), c.get("persona"), c.get("obs")]
        ok, err = cm.validate([(1, [cm.norm(x) for x in row])], items, dt.date.today())
        if err:
            return 400, dict(error=err[0].replace("Línea 1: ", "").strip())
        m = ok[0]
        antes = f"{ws.cell(row=r, column=5).value} {ws.cell(row=r, column=6).value}"
        obs = m["obs"]
        marca = f"[editado {dt.date.today():%d/%m/%Y}, antes {antes}]"
        obs = f"{obs} {marca}".strip() if marca not in (obs or "") else obs
        for col, v in ((1, m["fecha"]), (2, m["tipo"]), (5, m["cant"]), (6, m["unidad"]), (8, m["persona"]), (10, obs)):
            ws.cell(row=r, column=col).value = v
        cm.rebuild_stock(wb)
        try:
            save_wb(wb)
        except PermissionError:
            return 409, dict(error="El Excel está abierto: cerralo y reintentá.")
        regenerar_html(wb)
        return 200, dict(ok=True)


def api_admin_anular(body):
    with LOCK:
        flush_pend()
        wb = openpyxl.load_workbook(cm.XLSX)
        ws = wb[cm.SH_MOV]
        items = cm.read_maestro(wb)
        try:
            r = int(body.get("fila"))
        except (TypeError, ValueError):
            return 400, dict(error="Fila inválida (los pendientes no se pueden anular hasta que pasen al Excel).")
        if r < 2 or r > cm.last_row(ws):
            return 400, dict(error="Fila inválida.")
        cod = cm.norm(ws.cell(row=r, column=3).value).upper()
        if cod != cm.norm(body.get("cod")).upper() or cm.num(ws.cell(row=r, column=5).value) != cm.num(body.get("cant")):
            return 409, dict(error="El Excel cambió desde que se abrió la lista. Actualizá y reintentá.")
        motivo = cm.norm(body.get("motivo"))
        if not motivo:
            return 400, dict(error="Indicar el motivo.")
        if "Anulados" not in wb.sheetnames:
            an = wb.create_sheet("Anulados")
            cm.style_header(an, cm.MOV_HEAD + ["Anulado el", "Motivo"], [11, 9, 9, 30, 9, 9, 10, 13, 22, 40, 16, 40])
        an = wb["Anulados"]
        ra = cm.last_row(an) + 1
        vals = [ws.cell(row=r, column=c).value for c in range(1, 11)]
        vals[3] = items.get(cod, {}).get("ins", "")
        vals[6] = cm.base_qty(items[cod], cm.num(vals[4]) or 0, cm.norm(vals[5])) if cod in items else None
        for c, v in enumerate(vals + [dt.datetime.now(), motivo], start=1):
            an.cell(row=ra, column=c).value = v
        cm.style_row(an, ra, 12)
        an.cell(row=ra, column=1).number_format = "dd/mm/yyyy"
        an.cell(row=ra, column=11).number_format = "dd/mm/yyyy hh:mm"
        ws.delete_rows(r)
        _reescribir_formulas_mov(ws, r)
        ws.auto_filter.ref = f"A1:J{max(cm.last_row(ws), 2)}"
        cm.rebuild_stock(wb)
        try:
            save_wb(wb)
        except PermissionError:
            return 409, dict(error="El Excel está abierto: cerralo y reintentá.")
        regenerar_html(wb)
        return 200, dict(ok=True)


# ---------------------------------------------------------------- HTTP
class H(BaseHTTPRequestHandler):
    server_version = "InventarioBS/1.0"

    def log_message(self, fmt, *args):
        if "/api/" in (args[0] if args else "") and "POST" in (args[0] if args else ""):
            sys.stdout.write(f"{dt.datetime.now():%d/%m %H:%M:%S} {self.client_address[0]} {args[0]}\n")

    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        data = body if isinstance(body, bytes) else (
            json.dumps(body, ensure_ascii=False, default=str).encode("utf-8") if not isinstance(body, str) else body.encode("utf-8"))
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        if "Cache-Control" not in (extra or {}):
            self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _token(self):
        return self.headers.get("X-Token") or parse_qs(urlparse(self.path).query).get("token", [""])[0]

    def _guard(self, fn, *a):
        try:
            code, body = fn(*a)
        except Exception as e:  # error inesperado: se informa sin tumbar el servidor
            import traceback
            traceback.print_exc()
            import zipfile, zlib
            if isinstance(e, (zipfile.BadZipFile, zlib.error)):
                code, body = 500, dict(error="El Excel datos\\Inventario_Bajo_Soppalco.xlsx está dañado (no se puede leer). "
                                             "Restaurar una copia de datos\\backups (ver instrucciones) o, si es una instalación nueva, "
                                             "borrarlo: el programa crea uno nuevo.")
            else:
                code, body = 500, dict(error=f"Error interno: {e}")
        self._send(code, body)

    def do_GET(self):
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            with open(APP, encoding="utf-8") as fh:
                return self._send(200, fh.read(), "text/html; charset=utf-8")
        if path.startswith("/foto/"):
            cod = path[len("/foto/"):].upper()
            p = foto_path(cod)
            if not RE_COD.match(cod) or not os.path.isfile(p):
                return self._send(404, dict(error="Sin foto"))
            with open(p, "rb") as fh:
                return self._send(200, fh.read(), "image/jpeg", {"Cache-Control": "max-age=86400"})
        if path in ("/favicon.ico", "/icono.png"):
            f = os.path.join(BASE, "inventario.ico" if path.endswith(".ico") else "icono.png")
            if os.path.exists(f):
                with open(f, "rb") as fh:
                    return self._send(200, fh.read(), "image/x-icon" if f.endswith(".ico") else "image/png",
                                      {"Cache-Control": "max-age=86400"})
            return self._send(404, dict(error="sin icono"))
        if path == "/inventario":
            if not token_ok(self._token()):
                return self._send(401, "<p style='font:14px Arial'>Sesión de Administración cerrada. Volver a ingresar.</p>",
                                  "text/html; charset=utf-8")
            try:
                q = parse_qs(urlparse(self.path).query)
                g = lambda k: q.get(k, [""])[0]
                _, items, movs, st, _ = snapshot()
                return self._send(200, cm.build_html(items, movs, st, tema=g("tema"), desde=g("desde") or None,
                                                     hasta=g("hasta") or None, sector=g("sector") or None),
                                  "text/html; charset=utf-8")
            except Exception as e:
                return self._send(500, f"Error: {e}", "text/plain; charset=utf-8")
        if path == "/api/datos":
            return self._guard(api_datos)
        if path.startswith("/api/admin/") or path == "/excel":
            if not token_ok(self._token()):
                return self._send(401, dict(error="Sesión vencida: volver a ingresar."))
            if path == "/api/admin/maestro":
                return self._guard(api_admin_maestro_get)
            if path == "/api/admin/movimientos":
                return self._guard(api_admin_movs_get)
            if path == "/api/admin/personas":
                return self._guard(api_admin_personas_get)
            if path == "/api/admin/backups":
                return self._guard(api_admin_backups_get)
            if path == "/api/admin/exportar":
                q = parse_qs(urlparse(self.path).query)
                g = lambda k: q.get(k, [""])[0]
                _, items, movs, st, _ = snapshot()
                data = cm.exportar_xlsx(items, movs, st, g("desde") or None, g("hasta") or None, g("sector") or None)
                nombre = "Informe_Inventario_" + (g("desde") or "inicio") + "_" + (g("hasta") or "hoy") + ".xlsx"
                return self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": f'attachment; filename="{re.sub(r"[^A-Za-z0-9_.-]", "", nombre)}"'})
            if path == "/api/admin/backup_descargar":
                nombre = parse_qs(urlparse(self.path).query).get("nombre", [""])[0]
                p = os.path.join(BACKUPS, nombre)
                if not RE_BACKUP.match(nombre) or not os.path.isfile(p):
                    return self._send(404, dict(error="No encontrado"))
                with open(p, "rb") as fh:
                    data = fh.read()
                return self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": f'attachment; filename="{nombre}"'})
            if path == "/excel":
                with LOCK:
                    with open(cm.XLSX, "rb") as fh:
                        data = fh.read()
                name = f"Inventario_Bajo_Soppalco_{dt.date.today():%Y%m%d}.xlsx"
                return self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": f'attachment; filename="{name}"'})
        self._send(404, dict(error="No encontrado"))

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            n = int(self.headers.get("Content-Length", "0"))
            if n > 2_000_000:
                return self._send(413, dict(error="Demasiado grande"))
            body = json.loads(self.rfile.read(n).decode("utf-8") or "{}")
        except Exception:
            return self._send(400, dict(error="Datos inválidos"))
        ses = sesion(self.headers.get("X-Sesion"))
        if path == "/api/movimientos":
            return self._guard(api_movimientos, body, ses)
        if path == "/api/ingresar":
            return self._guard(api_ingresar, body)
        if path == "/api/mi_clave":
            return self._guard(api_mi_clave, body, ses)
        if path == "/api/salir":
            SES.pop(self.headers.get("X-Sesion") or "", None)
            return self._send(200, dict(ok=True))
        if path == "/api/sesion":
            return self._send(200 if ses else 401, dict(ok=bool(ses)) if ses else dict(error="Sesión cerrada."))
        if path == "/api/login":
            return self._guard(api_login, body)
        if path == "/api/apagar":
            # lo usa Desinstalar.bat: solo desde esta misma PC y con la clave de apagar.key
            if self.client_address[0] not in ("127.0.0.1", "::1") or not secrets.compare_digest(
                    self.headers.get("X-Apagar", ""), CLAVE_APAGAR):
                return self._send(403, dict(error="No permitido"))
            self._send(200, dict(ok=True))
            threading.Thread(target=apagar, daemon=True).start()
            return
        if path.startswith("/api/admin/"):
            if not token_ok(self._token()):
                return self._send(401, dict(error="Sesión vencida: volver a ingresar."))
            if path == "/api/admin/maestro":
                return self._guard(api_admin_maestro_post, body)
            if path == "/api/admin/anular":
                return self._guard(api_admin_anular, body)
            if path == "/api/admin/editar_mov":
                return self._guard(api_admin_editar_mov, body)
            if path == "/api/admin/personas":
                return self._guard(api_admin_personas_post, body)
            if path == "/api/admin/password":
                return self._guard(api_password, body)
            if path == "/api/admin/ping":
                return self._send(200, dict(ok=True))
            if path == "/api/admin/backup":
                return self._guard(api_admin_backup_post)
            if path == "/api/admin/foto":
                return self._guard(api_admin_foto, body)
            if path == "/api/admin/salir":
                TOKENS.pop(self._token(), None)
                return self._send(200, dict(ok=True))
        self._send(404, dict(error="No encontrado"))


class Servidor(ThreadingHTTPServer):
    # En Windows, SO_REUSEADDR deja abrir dos servidores en el mismo puerto: se desactiva
    allow_reuse_address = os.name != "nt"
    daemon_threads = True


def ip_local():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


CLAVE_APAGAR = secrets.token_hex(16)
SRV = None


def apagar():
    """Cierre ordenado: espera a que termine cualquier escritura del Excel."""
    time.sleep(0.3)
    with LOCK:
        if SRV:
            SRV.shutdown()


def main():
    global SRV
    if sys.stdout is None or sys.stderr is None:   # pythonw (arranque al iniciar sesión): sin consola
        fh = open(os.path.join(BASE, "servidor.log"), "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stdout or fh
        sys.stderr = sys.stderr or fh
    cfg = load_config()
    # Primero se toma el puerto: así nunca hay dos servidores tocando el mismo Excel.
    try:
        srv = SRV = Servidor((HOST, PORT), H)
    except OSError:
        # puerto ocupado: ya hay un servidor corriendo (p. ej. el de inicio automático)
        print(f"El servidor ya está corriendo en esta PC (puerto {PORT}). Abriendo el navegador...")
        if "--sin-navegador" in sys.argv:
            return
        webbrowser.open(f"http://localhost:{PORT}")
        time.sleep(3)
        return
    cm.asegurar_excel()
    threading.Thread(target=ciclo_backup, daemon=True).start()
    print("=" * 60)
    print(" INVENTARIO BAJO SOPPALCO - servidor de carga")
    print(f"   En esta PC:        http://localhost:{PORT}")
    if ACCESO_RED:
        print(f"   Desde otra PC:     http://{ip_local()}:{PORT}")
    else:
        print("   Desde otra PC:     deshabilitado (modo 'solo esta PC')")
    if cfg.get("clave_inicial"):
        print(f"   Contraseña admin inicial: {CLAVE_INICIAL}  (cambiarla)")
    print("   Para cerrar: Ctrl+C o cerrar esta ventana")
    print("=" * 60)
    if "--sin-navegador" not in sys.argv:
        threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    try:
        with open(os.path.join(BASE, "apagar.key"), "w", encoding="utf-8") as fh:
            fh.write(CLAVE_APAGAR)
    except OSError:
        pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    print("Servidor detenido.")


if __name__ == "__main__":
    main()
