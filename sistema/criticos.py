# -*- coding: utf-8 -*-
"""
Insumos críticos: parámetros de reposición, órdenes de compra y panel de alertas.

Solo LEE los movimientos y el stock que ya calcula el programa (cargar_movimientos.py);
no los modifica. Los datos propios se guardan en el mismo Excel, en tres hojas:
  - "Críticos"         parámetros por insumo
  - "OC"               órdenes de compra
  - "Config críticos"  parámetros generales
Todas las cantidades van en la UNIDAD BASE del insumo (la misma del stock).
"""
import datetime as dt
import io
import math

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import cargar_movimientos as cm

SH_CRI, SH_OC, SH_CFG = "Críticos", "OC", "Config críticos"
MAX_CRITICOS = 50

# ---------------------------------------------------------------- estructura de las hojas
# Críticos
(C_COD, C_INS, C_PROV, C_LT, C_MIN, C_SEG, C_MULT, C_INI, C_DCUB, C_OBS) = range(1, 11)
CRI_HEAD = ["Código", "Insumo", "Proveedor", "Lead time\n(días háb.)", "Stock mínimo fijo\n(u. base)",
            "Seguridad\n(días de consumo)", "Múltiplo de compra\n(u. base)", "Consumo semanal\ninicial estimado\n(u. base)",
            "Días a cubrir\npor pedido\n(días háb.)", "Observaciones"]
CRI_W = [12, 32, 18, 10, 12, 12, 12, 14, 12, 36]

# OC
(O_NRO, O_FEC, O_COD, O_INS, O_CANT, O_PROM, O_ANU, O_OBS, O_REG) = range(1, 10)
OC_HEAD = ["Nº OC / Requi", "Fecha emisión", "Código", "Insumo", "Cantidad pedida\n(u. base)",
           "Fecha prometida\n(opcional)", "Anulada", "Observaciones", "Registrada"]
OC_W = [14, 12, 12, 32, 12, 13, 9, 36, 24]

# Config críticos: (clave, título, valor por defecto, explicación)
CFG = [
    ("semanas", "Semanas de historial", 4,
     "El consumo semanal es el promedio de las salidas de las últimas N semanas."),
    ("dias_semana", "Días hábiles por semana", 5, "Convierte el consumo semanal en diario."),
    ("margen", "Margen de pre-alerta (%)", 20,
     "Si el stock está hasta este % por encima del punto de pedido, el estado es ATENCIÓN."),
    ("dias_cubrir", "Días a cubrir por defecto (días háb.)", 15,
     "Para los insumos que no tienen un valor propio de días a cubrir por pedido."),
]

ESTADOS = ["QUIEBRE", "COMPRAR YA", "COMPRAR", "EN CAMINO", "ATENCIÓN", "OK"]
URGENCIA = {"QUIEBRE": 0, "COMPRAR YA": 1, "COMPRAR": 2, "ATENCIÓN": 3, "EN CAMINO": 4, "OK": 5, "—": 6}
COLOR = {   # fondo, letra  (Excel)
    "QUIEBRE": ("7B1E1E", "FFFFFF"), "COMPRAR YA": ("E06666", "FFFFFF"), "COMPRAR": ("F6B26B", "000000"),
    "EN CAMINO": ("6FA8DC", "000000"), "ATENCIÓN": ("FFE599", "000000"), "OK": ("93C47D", "000000"),
}


def _hoja(wb, nombre, head, widths):
    if nombre in wb.sheetnames:
        return wb[nombre], False
    ws = wb.create_sheet(nombre)
    cm.style_header(ws, head, widths)
    ws.auto_filter.ref = f"A1:{cm.L(len(head))}1"
    return ws, True


def asegurar_hojas(wb):
    """Crea las hojas de críticos si faltan y el título de la columna Referencia en Movimientos.
    No modifica datos existentes. Devuelve True si cambió algo."""
    cambio = cm.asegurar_columna_ref(wb)
    _, n1 = _hoja(wb, SH_CRI, CRI_HEAD, CRI_W)
    _, n2 = _hoja(wb, SH_OC, OC_HEAD, OC_W)
    ws, n3 = _hoja(wb, SH_CFG, ["Parámetro", "Valor", "Qué hace"], [36, 10, 80])
    if n3:
        for i, (_, tit, val, exp) in enumerate(CFG, start=2):
            for c, v in enumerate((tit, val, exp), start=1):
                ws.cell(row=i, column=c).value = v
            cm.style_row(ws, i, 3, wide_cols=(3,))
            ws.cell(row=i, column=2).fill = cm.FILL_IN
    return cambio or n1 or n2 or n3


def _insumo_f(col_cod, r):
    return f'=IFERROR(INDEX({cm.SH_MAE}!$B:$B,MATCH(${col_cod}{r},{cm.SH_MAE}!$A:$A,0)),"CÓDIGO?")'


# ---------------------------------------------------------------- días hábiles (función aislada)
FERIADOS = set()   # fechas (datetime.date) no laborables. Por ahora vacío: lunes a viernes.


def es_habil(d):
    return d.weekday() < 5 and d not in FERIADOS


def sumar_habiles(desde, n):
    """Fecha que resulta de avanzar n días hábiles desde 'desde' (como WORKDAY de Excel)."""
    d, n = desde, int(n)
    while n > 0:
        d += dt.timedelta(days=1)
        if es_habil(d):
            n -= 1
    return d


def habiles_entre(a, b):
    """Días hábiles después de 'a' hasta 'b' inclusive (0 si b <= a)."""
    n, d = 0, a
    while d < b:
        d += dt.timedelta(days=1)
        if es_habil(d):
            n += 1
    return n


# ---------------------------------------------------------------- lectura
def _num(v):
    return cm.num(v)


def _fecha(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = cm.norm(v)
    for f in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def leer_config(wb):
    cfg = {k: v for k, _, v, _ in CFG}
    if SH_CFG in wb.sheetnames:
        ws = wb[SH_CFG]
        titulos = {tit: k for k, tit, _, _ in CFG}
        for r in range(2, cm.last_row(ws) + 1):
            k = titulos.get(cm.norm(ws.cell(row=r, column=1).value))
            v = _num(ws.cell(row=r, column=2).value)
            if k and v is not None:
                cfg[k] = v
    return cfg


def leer_criticos(wb):
    out = {}
    if SH_CRI not in wb.sheetnames:
        return out
    ws = wb[SH_CRI]
    for r in range(2, cm.last_row(ws) + 1):
        g = lambda c: ws.cell(row=r, column=c).value
        cod = cm.norm(g(C_COD)).upper()
        if not cod:
            continue
        out[cod] = dict(cod=cod, prov=cm.norm(g(C_PROV)), lt=_num(g(C_LT)), min=_num(g(C_MIN)),
                        seg=_num(g(C_SEG)), mult=_num(g(C_MULT)), ini=_num(g(C_INI)), dcub=_num(g(C_DCUB)),
                        obs=cm.norm(g(C_OBS)), fila=r)
    return out


def leer_oc(wb):
    out = []
    if SH_OC not in wb.sheetnames:
        return out
    ws = wb[SH_OC]
    for r in range(2, cm.last_row(ws) + 1):
        g = lambda c: ws.cell(row=r, column=c).value
        nro = cm.norm(g(O_NRO))
        if not nro:
            continue
        out.append(dict(nro=nro, fecha=_fecha(g(O_FEC)), cod=cm.norm(g(O_COD)).upper(), cant=_num(g(O_CANT)) or 0,
                        prom=_fecha(g(O_PROM)), anulada=cm.norm(g(O_ANU)).lower() in ("sí", "si", "s", "x", "1"),
                        obs=cm.norm(g(O_OBS)), reg=cm.norm(g(O_REG)), fila=r))
    return out


def norm_ref(s):
    return " ".join(cm.norm(s).upper().split())


# ---------------------------------------------------------------- cálculos
def calcular(items, movs, st, cri, ocs, cfg, hoy=None):
    """Devuelve dict(filas, ocs, contadores, hoy). Solo lectura de movimientos y stock."""
    hoy = hoy or dt.date.today()
    N = max(1, int(cfg["semanas"]))
    dsem = max(1, float(cfg["dias_semana"]))
    margen = float(cfg["margen"]) / 100.0
    dcub_def = float(cfg["dias_cubrir"])

    # --- órdenes de compra
    recibido = {}
    for m in movs:
        if m["tipo"] == "ENTRADA" and m.get("ref") and m["base"] is not None:
            k = (norm_ref(m["ref"]), m["cod"])
            recibido[k] = recibido.get(k, 0) + m["base"]
    oc_out = []
    for o in ocs:
        p = cri.get(o["cod"], {})
        lt = int(p.get("lt") or 0)
        prev = o["prom"] or (sumar_habiles(o["fecha"], lt) if o["fecha"] else None)
        rec = recibido.get((norm_ref(o["nro"]), o["cod"]), 0)
        pend = 0 if o["anulada"] else max(o["cant"] - rec, 0)
        if o["anulada"]:
            est = "Anulada"
        elif pend <= 0:
            est = "Recibida"
        elif prev and prev < hoy:
            est = "Atrasada"
        else:
            est = "En curso"
        atraso = habiles_entre(prev, hoy) if est == "Atrasada" else None
        oc_out.append(dict(o, prev=prev, recibido=rec, pendiente=pend, estado=est, atraso=atraso,
                           ins=items.get(o["cod"], {}).get("ins", "")))

    # --- consumo: salidas de las últimas N semanas
    desde = hoy - dt.timedelta(days=7 * N)
    fechas = [m["fecha"] for m in movs if isinstance(m["fecha"], dt.date)]
    primer = min(fechas) if fechas else hoy
    # Con historial corto (programa recién empezado) se divide por las semanas que realmente pasaron,
    # para no subestimar el consumo. Mínimo 1 semana, máximo N.
    semanas = min(float(N), max(1.0, (hoy - primer).days / 7.0))
    salidas, incompleto = {}, set()
    for m in movs:
        if m["tipo"] != "SALIDA" or not isinstance(m["fecha"], dt.date) or not (desde < m["fecha"] <= hoy):
            continue
        if m["base"] is None:
            incompleto.add(m["cod"])
            continue
        salidas[m["cod"]] = salidas.get(m["cod"], 0) + m["base"]

    filas = []
    for cod, p in cri.items():
        it, s = items.get(cod), st.get(cod)
        f = dict(cod=cod, ins=it["ins"] if it else "", ub=it["ub"] if it else "", prov=p["prov"],
                 lt=p["lt"], min=p["min"], seg=p["seg"], mult=p["mult"], ini=p["ini"], dcub=p["dcub"], obs=p["obs"])
        alertas = []
        if not it:
            f.update(stock=None, estado="—", alertas="No está en el Maestro", sugerida=0)
            filas.append(f)
            continue
        stock = s["act"]
        if s["falta"] or cod in incompleto:
            alertas.append("Revisar Maestro: falta cant. por empaque")
        if it["ini"] is None:
            alertas.append("Sin stock inicial en el Maestro")
        lt = float(p["lt"] or 0)
        seg = float(p["seg"] or 0)
        mult = max(1.0, float(p["mult"] or 1))
        minfijo = float(p["min"] or 0)
        dcub = float(p["dcub"]) if p["dcub"] else dcub_def

        sem = salidas.get(cod, 0) / semanas
        if sem <= 0 and p["ini"]:
            sem = float(p["ini"])
        diario = sem / dsem
        cob = stock / diario if diario > 0 else None
        pp = max(minfijo, diario * (lt + seg))
        ocp = sum(o["pendiente"] for o in oc_out if o["cod"] == cod and o["estado"] in ("En curso", "Atrasada"))
        disp = stock + ocp

        if stock <= 0:
            est = "QUIEBRE"
        elif stock <= pp:
            if ocp > 0 and disp > pp:
                est = "EN CAMINO"
            elif cob is not None and cob <= lt:
                est = "COMPRAR YA"
            else:
                est = "COMPRAR"
        elif stock <= pp * (1 + margen):
            est = "ATENCIÓN"
        else:
            est = "OK"

        sug = 0
        if est in ("QUIEBRE", "COMPRAR YA", "COMPRAR") and disp <= pp:
            x = max(pp + diario * dcub - disp, mult)
            sug = round(math.ceil(x / mult - 1e-9) * mult, 3)
            sug = int(sug) if sug == int(sug) else sug

        if est == "EN CAMINO":
            antes = "OC en curso"
        elif cob is None:
            antes = ""
        elif cob <= lt:
            antes = "YA"
        elif cob > 365:
            antes = ""
        else:
            antes = sumar_habiles(hoy, math.floor(cob - lt)).strftime("%d/%m/%Y")

        if stock <= 0:
            quiebre = "En quiebre"
        elif cob is None or cob > 365:
            quiebre = ""
        else:
            quiebre = sumar_habiles(hoy, math.floor(cob)).strftime("%d/%m/%Y")

        if any(o["cod"] == cod and o["estado"] == "Atrasada" for o in oc_out):
            alertas.insert(0, "OC atrasada")
        f.update(stock=stock, sem=sem, diario=diario, cob=cob, pp=pp, ocp=ocp, disp=disp, estado=est,
                 sugerida=sug, antes=antes, quiebre=quiebre, alertas="; ".join(alertas), dcub_usado=dcub)
        filas.append(f)

    filas.sort(key=lambda f: (URGENCIA.get(f["estado"], 9),
                              f.get("cob") if f.get("cob") is not None else float("inf"), f["cod"]))
    cont = {e: sum(1 for f in filas if f["estado"] == e) for e in ESTADOS}
    cont["OC atrasadas"] = sum(1 for o in oc_out if o["estado"] == "Atrasada")
    return dict(filas=filas, ocs=oc_out, contadores=cont, hoy=hoy, semanas_usadas=semanas)


# ---------------------------------------------------------------- escritura (la llama el servidor con el LOCK tomado)
def _val_num(v, nombre, minimo=0, entero=False, opcional=True):
    s = cm.norm(v).replace(",", ".")
    if s == "":
        if opcional:
            return None
        raise ValueError(f"Falta {nombre}.")
    try:
        x = float(s)
    except ValueError:
        raise ValueError(f"{nombre}: número inválido.")
    if x < minimo:
        raise ValueError(f"{nombre}: no puede ser menor que {minimo:g}.")
    if entero and x != int(x):
        raise ValueError(f"{nombre}: tiene que ser un número entero.")
    return int(x) if x == int(x) else x


def guardar_critico(wb, items, cod, campos):
    """Alta o edición de un insumo crítico."""
    cod = cm.norm(cod).upper()
    if cod not in items:
        raise ValueError(f"El código {cod or '(vacío)'} no está en el Maestro: darlo de alta ahí primero.")
    ws = wb[SH_CRI]
    cri = leer_criticos(wb)
    vals = dict(prov=cm.norm(campos.get("prov"))[:60],
                lt=_val_num(campos.get("lt"), "Lead time", 0, entero=True, opcional=False),
                min=_val_num(campos.get("min"), "Stock mínimo fijo", 0),
                seg=_val_num(campos.get("seg"), "Seguridad", 0),
                mult=_val_num(campos.get("mult"), "Múltiplo de compra", 1) or 1,
                ini=_val_num(campos.get("ini"), "Consumo semanal inicial", 0),
                dcub=_val_num(campos.get("dcub"), "Días a cubrir", 1),
                obs=cm.norm(campos.get("obs"))[:200])
    if cod in cri:
        r = cri[cod]["fila"]
    else:
        if len(cri) >= MAX_CRITICOS:
            raise ValueError(f"Máximo {MAX_CRITICOS} insumos críticos.")
        r = cm.last_row(ws) + 1
    for c, v in ((C_COD, cod), (C_INS, _insumo_f("A", r)), (C_PROV, vals["prov"] or None), (C_LT, vals["lt"]),
                 (C_MIN, vals["min"]), (C_SEG, vals["seg"]), (C_MULT, vals["mult"]), (C_INI, vals["ini"]),
                 (C_DCUB, vals["dcub"]), (C_OBS, vals["obs"] or None)):
        ws.cell(row=r, column=c).value = v
    cm.style_row(ws, r, len(CRI_HEAD), wide_cols=(C_OBS,))
    for c in range(C_LT, C_DCUB + 1):
        ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")
    ws.auto_filter.ref = f"A1:{cm.L(len(CRI_HEAD))}{max(cm.last_row(ws), 2)}"


def borrar_critico(wb, items, movs, st, cod):
    cod = cm.norm(cod).upper()
    cri = leer_criticos(wb)
    if cod not in cri:
        raise ValueError("Ese insumo no está en la lista de críticos.")
    calc = calcular(items, movs, st, cri, leer_oc(wb), leer_config(wb))
    if any(o["cod"] == cod and o["estado"] in ("En curso", "Atrasada") for o in calc["ocs"]):
        raise ValueError("Tiene órdenes de compra abiertas: anularlas o esperar a recibirlas antes de quitarlo.")
    ws = wb[SH_CRI]
    r = cri[cod]["fila"]
    ws.delete_rows(r)
    for rr in range(r, cm.last_row(ws) + 1):             # la fórmula de Insumo depende de la fila
        if ws.cell(row=rr, column=C_COD).value:
            ws.cell(row=rr, column=C_INS).value = _insumo_f("A", rr)


def guardar_config(wb, campos):
    ws = wb[SH_CFG]
    lim = dict(semanas=(1, 52, True), dias_semana=(1, 7, True), margen=(0, 500, False), dias_cubrir=(1, 365, False))
    nuevos = {}
    for k, tit, _, _ in CFG:
        a, b, ent = lim[k]
        v = _val_num(campos.get(k), tit, a, entero=ent, opcional=False)
        if v > b:
            raise ValueError(f"{tit}: máximo {b}.")
        nuevos[tit] = v
    for r in range(2, cm.last_row(ws) + 1):
        tit = cm.norm(ws.cell(row=r, column=1).value)
        if tit in nuevos:
            ws.cell(row=r, column=2).value = nuevos.pop(tit)
    for tit, v in nuevos.items():                        # por si alguien borró una fila de la hoja
        r = cm.last_row(ws) + 1
        ws.cell(row=r, column=1).value, ws.cell(row=r, column=2).value = tit, v


def guardar_oc(wb, items, campos, quien, fila=None, orig=None):
    """Alta (fila=None) o edición de una orden de compra."""
    cri = leer_criticos(wb)
    nro = cm.norm(campos.get("nro"))[:30]
    cod = cm.norm(campos.get("cod")).upper()
    if not nro:
        raise ValueError("Falta el Nº de OC o requisición.")
    if cod not in cri:
        raise ValueError("Elegir un insumo de la lista de críticos.")
    fec = _fecha(campos.get("fecha"))
    if not fec:
        raise ValueError("Fecha de emisión inválida.")
    if fec > dt.date.today():
        raise ValueError("La fecha de emisión no puede ser futura.")
    cant = _val_num(campos.get("cant"), "Cantidad pedida", 0, opcional=False)
    if cant <= 0:
        raise ValueError("Cantidad pedida: tiene que ser mayor que 0.")
    prom = None
    if cm.norm(campos.get("prom")):
        prom = _fecha(campos.get("prom"))
        if not prom:
            raise ValueError("Fecha prometida inválida.")
        if prom < fec:
            raise ValueError("La fecha prometida no puede ser anterior a la de emisión.")
    anulada = bool(campos.get("anulada"))
    ws = wb[SH_OC]
    existentes = leer_oc(wb)
    if fila is None:
        if any(norm_ref(o["nro"]) == norm_ref(nro) and o["cod"] == cod for o in existentes):
            raise ValueError(f"La OC {nro} ya tiene cargado {cod}: editar esa fila.")
        r = cm.last_row(ws) + 1
        reg = f"{dt.datetime.now():%d/%m/%Y %H:%M} {quien}".strip()
    else:
        r = int(fila)
        actual = next((o for o in existentes if o["fila"] == r), None)
        if not actual or (orig and (norm_ref(actual["nro"]) != norm_ref(orig.get("nro")) or actual["cod"] != cm.norm(orig.get("cod")).upper())):
            raise LookupError("El Excel cambió desde que se abrió la lista. Actualizá y reintentá.")
        if any(o["fila"] != r and norm_ref(o["nro"]) == norm_ref(nro) and o["cod"] == cod for o in existentes):
            raise ValueError(f"La OC {nro} ya tiene cargado {cod} en otra fila.")
        reg = actual["reg"]
    for c, v in ((O_NRO, nro), (O_FEC, fec), (O_COD, cod), (O_INS, _insumo_f("C", r)), (O_CANT, cant),
                 (O_PROM, prom), (O_ANU, "Sí" if anulada else "No"), (O_OBS, cm.norm(campos.get("obs"))[:200] or None),
                 (O_REG, reg)):
        ws.cell(row=r, column=c).value = v
    cm.style_row(ws, r, len(OC_HEAD), wide_cols=(O_OBS,))
    for c in (O_FEC, O_PROM):
        ws.cell(row=r, column=c).number_format = "dd/mm/yyyy"
    for c in (O_NRO, O_FEC, O_COD, O_CANT, O_PROM, O_ANU):
        ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")
    ws.auto_filter.ref = f"A1:{cm.L(len(OC_HEAD))}{max(cm.last_row(ws), 2)}"


# ---------------------------------------------------------------- descarga del panel en Excel
def exportar_xlsx(calc, cfg):
    hoy = calc["hoy"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Panel"
    ws["A1"] = f"Insumos críticos · {hoy:%d/%m/%Y}"
    ws["A1"].font = Font(name=cm.FONT, size=13, bold=True)
    for i, e in enumerate(ESTADOS + ["OC atrasadas"], start=1):
        h = ws.cell(row=3, column=i, value=e)
        v = ws.cell(row=4, column=i, value=calc["contadores"][e])
        h.font, h.border, v.border = cm.F_B, cm.BRD, cm.BRD
        h.alignment = v.alignment = Alignment(horizontal="center")
        v.font = Font(name=cm.FONT, size=12, bold=True)
        if e in COLOR:
            h.fill = PatternFill("solid", fgColor=COLOR[e][0])
            h.font = Font(name=cm.FONT, size=10, bold=True, color=COLOR[e][1])
    head = ["Código", "Descripción", "Unidad", "Proveedor", "Stock actual", "Consumo semanal", "Cobertura\n(días háb.)",
            "Punto de pedido", "OC pendientes", "Disponible\nproyectado", "Estado", "Cantidad\nsugerida",
            "Pedir antes de", "Quiebre estimado", "Alertas"]
    R0 = 6
    for c, h in enumerate(head, start=1):
        cell = ws.cell(row=R0, column=c, value=h)
        cell.font, cell.fill, cell.border = cm.F_H, cm.FILL_H, cm.BRD
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[R0].height = 30
    r1 = lambda x: None if x is None else round(x, 1)
    for i, f in enumerate(calc["filas"], start=R0 + 1):
        cob = f.get("cob")
        vals = [f["cod"], f["ins"], f["ub"], f["prov"], r1(f.get("stock")), r1(f.get("sem")),
                ("sin consumo" if cob is None and f.get("stock") is not None else r1(cob)),
                r1(f.get("pp")), r1(f.get("ocp")), r1(f.get("disp")), f["estado"], f.get("sugerida") or 0,
                f.get("antes", ""), f.get("quiebre", ""), f.get("alertas", "")]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font, cell.border = cm.F_N, cm.BRD
            if 5 <= c <= 12:
                cell.alignment = Alignment(horizontal="right" if c != 11 else "center")
        e = ws.cell(row=i, column=11)
        if f["estado"] in COLOR:
            e.fill = PatternFill("solid", fgColor=COLOR[f["estado"]][0])
            e.font = Font(name=cm.FONT, size=10, bold=True, color=COLOR[f["estado"]][1])
    for c, w in enumerate([13, 32, 9, 16, 10, 11, 11, 11, 11, 11, 13, 11, 13, 14, 34], start=1):
        ws.column_dimensions[cm.L(c)].width = w
    ws.freeze_panes = ws.cell(row=R0 + 1, column=3)
    ws.auto_filter.ref = f"A{R0}:{cm.L(len(head))}{max(R0 + len(calc['filas']), R0 + 1)}"

    oc = wb.create_sheet("OC")
    oh = ["Nº OC / Requi", "Fecha emisión", "Código", "Insumo", "Cantidad pedida", "Fecha prometida", "Fecha prevista",
          "Recibido", "Pendiente", "Anulada", "Estado", "Días de atraso (háb.)", "Observaciones"]
    cm.style_header(oc, oh, [14, 12, 12, 30, 11, 12, 12, 10, 10, 9, 11, 12, 30])
    for i, o in enumerate(calc["ocs"], start=2):
        vals = [o["nro"], o["fecha"], o["cod"], o["ins"], o["cant"], o["prom"], o["prev"], o["recibido"],
                o["pendiente"], "Sí" if o["anulada"] else "No", o["estado"], o["atraso"], o["obs"]]
        for c, v in enumerate(vals, start=1):
            oc.cell(row=i, column=c, value=v)
        cm.style_row(oc, i, len(oh))
        for c in (2, 6, 7):
            oc.cell(row=i, column=c).number_format = "dd/mm/yyyy"

    pa = wb.create_sheet("Parámetros")
    ph = ["Código", "Insumo", "Lead time (días háb.)", "Stock mínimo fijo", "Seguridad (días)", "Múltiplo de compra",
          "Consumo semanal inicial", "Días a cubrir"]
    cm.style_header(pa, ph, [13, 32, 12, 12, 12, 12, 13, 12])
    for i, f in enumerate(sorted(calc["filas"], key=lambda f: f["cod"]), start=2):
        for c, v in enumerate([f["cod"], f["ins"], f["lt"], f["min"], f["seg"], f["mult"], f["ini"], f["dcub"]], start=1):
            pa.cell(row=i, column=c, value=v)
        cm.style_row(pa, i, len(ph))
    r = len(calc["filas"]) + 3
    pa.cell(row=r, column=1, value="Parámetros generales").font = cm.F_B
    for j, (k, tit, _, _) in enumerate(CFG, start=1):
        pa.cell(row=r + j, column=1, value=tit)
        pa.cell(row=r + j, column=2, value=cfg[k])
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
