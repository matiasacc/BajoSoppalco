# -*- coding: utf-8 -*-
"""
INVENTARIO BAJO SOPPALCO - carga de movimientos
================================================
Uso: doble clic en "herramientas/Cargar movimientos desde CSV.bat"
    (o: python sistema/cargar_movimientos.py [--solo-html])
Lee herramientas/movimientos.csv y escribe en datos/Inventario_Bajo_Soppalco.xlsx.

CSV (separador ; o ,). Columnas:  fecha;tipo;codigo;cantidad;unidad;persona;obs
  fecha    dd/mm/aaaa, dd/mm/aa o dd/mm (año actual). Vacía = hoy.
  tipo     E = entrada / S = salida  (también acepta ENTRADA / SALIDA)
  codigo   el código de la hoja Maestro (ej. CIN-07)
  cantidad número (acepta coma decimal)
  unidad   vacía = unidad base. Si se pone el empaque (ej. Caja) se convierte
           con "Cant. por empaque" del Maestro.
  persona  quién retira / ingresa
  obs      opcional
Las filas que empiezan con # se ignoran.

Si hay algún error en el CSV NO se carga nada y se indica la línea.
Cargado OK: el CSV se mueve a la carpeta "cargados" y queda uno vacío listo.

Requiere: Python 3.8+ y openpyxl  (pip install openpyxl)
"""
import csv
import re
import datetime as dt
import html
import os
import shutil
import sys
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except ImportError:
    sys.exit("Falta openpyxl. Instalar con:  pip install openpyxl")

# Estructura de carpetas:
#   <raíz>/datos/         Excel (base de datos), informe HTML, pendientes
#   <raíz>/herramientas/  carga por CSV y arranque manual
#   <raíz>/sistema/       programa (este archivo)
BASE = os.path.dirname(os.path.abspath(__file__))          # .../sistema
RAIZ = os.path.dirname(BASE)
DATOS = os.path.join(RAIZ, "datos")
HERR = os.path.join(RAIZ, "herramientas")
XLSX = os.path.join(DATOS, "Inventario_Bajo_Soppalco.xlsx")
CSV_IN = os.path.join(HERR, "movimientos.csv")
CARGADOS = os.path.join(HERR, "cargados")
HTML_OUT = os.path.join(DATOS, "Informe_Inventario.html")
CSV_HEADER = ["fecha", "tipo", "codigo", "cantidad", "unidad", "persona", "obs"]

SH_MAE, SH_MOV, SH_STK = "Maestro", "Movimientos", "Stock"

# Maestro: columnas
(M_COD, M_INS, M_SEC, M_UBI, M_UB, M_EMP, M_FAC, M_EMP2, M_FAC2,
 M_SIE, M_SIS, M_SIT, M_MIN, M_OBS, M_ORI, M_BAR) = range(1, 17)
MAESTRO_HEAD = ["Código", "Insumo", "Sector", "Ubicación", "Unidad base",
                "Empaque 1\n(ej. Caja)", "Cant. por\nempaque 1\n(u. base)",
                "Empaque 2\n(ej. Bulto)", "Cant. por\nempaque 2\n(u. base)",
                "Stock inicial\n(empaques 1)", "Stock inicial\n(sueltas, u. base)",
                "Stock inicial\ntotal (u. base)", "Stock mínimo\n(u. base)", "Observaciones", "Origen",
                "Código de barras\n(vacío = usa Código)"]
MAESTRO_W = [9, 32, 10, 10, 10, 11, 10, 11, 10, 11, 12, 12, 11, 46, 18, 18]
L = openpyxl.utils.get_column_letter
# Movimientos: columnas
MOV_HEAD = ["Fecha", "Tipo", "Código", "Insumo", "Cantidad", "Unidad", "Cant.\n(u. base)",
            "Persona", "Origen del dato", "Observaciones"]
STK_HEAD = ["Código", "Insumo", "Unidad base", "Stock inicial", "Entradas", "Salidas",
            "Stock actual\n(u. base)", "Empaque", "Cant. por\nempaque", "Empaques\ncerrados", "Sueltas\n(u. base)",
            "Stock mínimo", "Estado"]

FONT = "Arial"
F_N = Font(name=FONT, size=10)
F_B = Font(name=FONT, size=10, bold=True)
F_H = Font(name=FONT, size=10, bold=True, color="FFFFFF")
FILL_H = PatternFill("solid", fgColor="404040")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")   # celdas a completar / editables
THIN = Side(style="thin", color="BFBFBF")
BRD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FALTA = "FALTA FACTOR"


# ---------------------------------------------------------------- utilidades Excel
def style_header(ws, headers, widths):
    for c, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font, cell.fill, cell.border = F_H, FILL_H, BRD
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"


def style_row(ws, r, ncols, wide_cols=()):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=r, column=c)
        cell.font, cell.border = F_N, BRD
        cell.alignment = Alignment(vertical="center", wrap_text=(c in wide_cols))


def mov_row_formulas(r):
    """Fórmulas de una fila de Movimientos (Insumo y Cant. u. base).
    Unidad = empaque 1 -> x cant. empaque 1; = empaque 2 -> x cant. empaque 2; otra -> unidad base."""
    m = f"MATCH($C{r},{SH_MAE}!$A:$A,0)"
    col = lambda c: f"INDEX({SH_MAE}!${L(c)}:${L(c)},{m})"
    insumo = f'=IFERROR({col(M_INS)},"CÓDIGO?")'
    e1, f1, e2, f2 = col(M_EMP), col(M_FAC), col(M_EMP2), col(M_FAC2)
    base = (f'=IFERROR(IF(AND($F{r}<>"",$F{r}={e1}),IF({f1}="","{FALTA}",$E{r}*{f1}),'
            f'IF(AND($F{r}<>"",$F{r}={e2}),IF({f2}="","{FALTA}",$E{r}*{f2}),$E{r})),"CÓDIGO?")')
    return insumo, base


def maestro_total_formula(r):
    """Stock inicial total (u. base) = empaques 1 x cant. por empaque 1 + sueltas."""
    e, s, f = f"{L(M_SIE)}{r}", f"{L(M_SIS)}{r}", f"{L(M_FAC)}{r}"
    return (f'=IF(AND({e}="",{s}=""),"",IF(AND({e}<>"",{f}=""),"{FALTA}",'
            f'N({e})*N({f})+N({s})))')


def style_maestro_row(ws, r):
    style_row(ws, r, len(MAESTRO_HEAD), wide_cols=(M_OBS,))
    for c in (M_COD, M_UBI, M_UB, M_EMP, M_FAC, M_EMP2, M_FAC2, M_SIE, M_SIS, M_SIT, M_MIN):
        ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")
    for c in (M_FAC, M_FAC2, M_SIE, M_SIS, M_MIN, M_BAR):
        ws.cell(row=r, column=c).fill = FILL_IN
    ws.cell(row=r, column=M_BAR).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=r, column=M_BAR).number_format = "@"


def write_movement(ws, r, fecha, tipo, cod, cant, unidad, persona, origen, obs):
    insumo_f, base_f = mov_row_formulas(r)
    vals = [fecha, tipo, cod, insumo_f, cant, unidad, base_f, persona, origen, obs]
    for c, v in enumerate(vals, start=1):
        ws.cell(row=r, column=c).value = v
    style_row(ws, r, len(MOV_HEAD), wide_cols=(10,))
    ws.cell(row=r, column=1).number_format = "dd/mm/yyyy"
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for c in (2, 3, 5, 6, 7):
        ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")


def last_row(ws, col=1):
    r = ws.max_row
    while r > 1 and ws.cell(row=r, column=col).value in (None, ""):
        r -= 1
    return r


def rebuild_stock(wb):
    """Rehace la hoja Stock con una fila por código del Maestro (todo con fórmulas)."""
    mae = wb[SH_MAE]
    if SH_STK in wb.sheetnames:
        idx = wb.sheetnames.index(SH_STK)
        del wb[SH_STK]
        ws = wb.create_sheet(SH_STK, idx)
    else:
        ws = wb.create_sheet(SH_STK)
    style_header(ws, STK_HEAD, [10, 34, 11, 12, 10, 10, 13, 9, 10, 11, 11, 12, 26])
    n = last_row(mae)
    mv = SH_MOV
    for i, mr in enumerate(range(2, n + 1), start=2):
        r = i
        ref = lambda col: f"{SH_MAE}!{L(col)}{mr}"
        okfac = f"AND($H{r}<>\"\",ISNUMBER($I{r}),$G{r}>=0)"
        f = {
            1: f"={ref(M_COD)}",
            2: f"={ref(M_INS)}",
            3: f"={ref(M_UB)}",
            4: f'=IF({ref(M_SIT)}="","A relevar",{ref(M_SIT)})',
            5: f'=SUMIFS({mv}!$G:$G,{mv}!$C:$C,$A{r},{mv}!$B:$B,"ENTRADA")',
            6: f'=SUMIFS({mv}!$G:$G,{mv}!$C:$C,$A{r},{mv}!$B:$B,"SALIDA")',
            7: f"=IF(ISNUMBER($D{r}),$D{r},0)+$E{r}-$F{r}",
            8: f'=IF({ref(M_EMP)}="","",{ref(M_EMP)})',
            9: f'=IF(AND(ISNUMBER({ref(M_FAC)}),{ref(M_FAC)}>0),{ref(M_FAC)},"")',
            10: f'=IF({okfac},INT($G{r}/$I{r}),"")',
            11: f'=IF({okfac},$G{r}-$J{r}*$I{r},"")',
            12: f'=IF({ref(M_MIN)}="","",{ref(M_MIN)})',
            13: (f'=IF(OR($D{r}="{FALTA}",COUNTIFS({mv}!$C:$C,$A{r},{mv}!$G:$G,"{FALTA}")>0),'
                 f'"* Falta cant. por empaque",'
                 f'IF($D{r}="A relevar","* Sin stock inicial",'
                 f'IF($G{r}<0,"* Negativo",'
                 f'IF(AND(ISNUMBER($L{r}),$G{r}<=$L{r}),"* Bajo mínimo","OK"))))'),
        }
        for c, v in f.items():
            ws.cell(row=r, column=c).value = v
        style_row(ws, r, len(STK_HEAD))
        for c in range(3, 13):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=7).font = F_B
    ws.auto_filter.ref = f"A1:M{n}"
    return ws


# ---------------------------------------------------------------- Personas
SH_PER = "Personas"
PER_HEAD = ["Legajo", "Nombre", "Activo (Sí/No)", "Observaciones", "Clave (cifrada)", "Cambiar clave\nal ingresar"]
P_LEG, P_NOM, P_ACT, P_OBS, P_CLA, P_CAM = range(1, 7)


def hash_clave(clave, salt=None):
    """PBKDF2-SHA256 con salt. Se guarda 'salt$hash' (nunca la clave en texto)."""
    import hashlib
    import secrets
    salt = salt or secrets.token_hex(8)
    h = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), bytes.fromhex(salt), 100000).hex()
    return f"{salt}${h}"


def clave_ok(clave, guardada):
    import secrets
    if not guardada or "$" not in guardada:
        return False
    salt = guardada.split("$", 1)[0]
    return secrets.compare_digest(hash_clave(clave, salt), guardada)


def personas_sheet(wb):
    """Devuelve la hoja Personas con el formato actual (la crea o la actualiza si hace falta)."""
    if SH_PER not in wb.sheetnames:
        ws = wb.create_sheet(SH_PER)
        style_header(ws, PER_HEAD, [11, 28, 13, 36, 22, 13])
        ws.column_dimensions["E"].hidden = True
        ws.auto_filter.ref = "A1:F1"
        return ws
    ws = wb[SH_PER]
    if norm(ws.cell(row=1, column=1).value) != "Legajo":           # formato viejo: Nombre | Activo | Obs
        old = [[ws.cell(row=r, column=c).value for c in (1, 2, 3)] for r in range(2, last_row(ws) + 1)]
        idx = wb.sheetnames.index(SH_PER)
        del wb[SH_PER]
        ws = wb.create_sheet(SH_PER, idx)
        style_header(ws, PER_HEAD, [11, 28, 13, 36, 22, 13])
        ws.column_dimensions["E"].hidden = True
        for r, (n, a, o) in enumerate(old, start=2):
            write_persona(ws, r, None, n, norm(a).lower() not in ("no", "n"), o)
    return ws


def read_personas(wb, solo_activas=True):
    if SH_PER not in wb.sheetnames:
        return []
    ws = wb[SH_PER]
    if norm(ws.cell(row=1, column=1).value) != "Legajo":
        ws = personas_sheet(wb)
    out = []
    for r in range(2, last_row(ws, P_NOM) + 1):
        g = lambda c: ws.cell(row=r, column=c).value
        n = norm(g(P_NOM))
        if not n:
            continue
        leg = norm(g(P_LEG))
        if leg.endswith(".0"):
            leg = leg[:-2]
        act = norm(g(P_ACT)).lower() not in ("no", "n", "0", "false")
        if act or not solo_activas:
            out.append(dict(legajo=leg, nombre=n, activo=act, obs=norm(g(P_OBS)), clave=norm(g(P_CLA)),
                            cambiar=norm(g(P_CAM)).lower() in ("sí", "si", "s"), fila=r))
    return out


def write_persona(ws, r, legajo, nombre, activo=True, obs=None, clave=None, cambiar=False):
    vals = {P_LEG: legajo or None, P_NOM: nombre, P_ACT: "Sí" if activo else "No", P_OBS: obs or None,
            P_CLA: clave or None, P_CAM: "Sí" if cambiar else None}
    for c, v in vals.items():
        ws.cell(row=r, column=c).value = v
    style_row(ws, r, len(PER_HEAD))
    ws.cell(row=r, column=P_LEG).number_format = "@"
    for c in (P_LEG, P_ACT, P_CAM):
        ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")
    ws.auto_filter.ref = f"A1:F{r}"


def crear_excel_vacio(path):
    """Crea la planilla con la estructura completa y sin datos (primera instalación)."""
    wb = openpyxl.Workbook()
    ins = wb.active
    ins.title = "Instrucciones"
    wb.create_sheet(SH_STK)
    mov = wb.create_sheet(SH_MOV)
    mae = wb.create_sheet(SH_MAE)
    style_header(mae, MAESTRO_HEAD, MAESTRO_W)
    mae.row_dimensions[1].height = 45
    style_header(mov, MOV_HEAD, [11, 9, 9, 30, 9, 9, 10, 13, 26, 52])
    personas_sheet(wb)
    rebuild_stock(wb)
    ins.column_dimensions["A"].width = 22
    ins.column_dimensions["B"].width = 95
    lineas = [
        ("INVENTARIO BAJO SOPPALCO", None), ("", None),
        ("Stock", "Stock actual por insumo (fórmulas, no se edita)."),
        ("Movimientos", "Entradas y salidas. Se cargan desde el programa."),
        ("Maestro", "Insumos, unidades, empaques y stock inicial. Se edita desde Administración."),
        ("Personas", "Legajo, nombre y contraseña cifrada de quienes cargan. Se administra desde Administración."),
        ("", None), ("Instrucciones completas", "LEEME - Instrucciones.pdf"),
    ]
    for r, (a, b) in enumerate(lineas, start=1):
        ins.cell(row=r, column=1).value = a
        ins.cell(row=r, column=2).value = b
        ins.cell(row=r, column=1).font = F_B if b is None else F_N
        ins.cell(row=r, column=2).font = F_N
    ins["A1"].font = Font(name=FONT, size=13, bold=True)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def asegurar_excel():
    if not os.path.exists(XLSX):
        crear_excel_vacio(XLSX)
        print(f"No había planilla: se creó una vacía en {XLSX}")


# ---------------------------------------------------------------- lectura de datos (Python)
def norm(s):
    return (str(s).strip() if s is not None else "")


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def read_maestro(wb):
    ws = wb[SH_MAE]
    items = {}
    for r in range(2, last_row(ws) + 1):
        g = lambda c: ws.cell(row=r, column=c).value
        cod = norm(g(M_COD)).upper()
        if not cod:
            continue
        fac, fac2, sie, sis = num(g(M_FAC)), num(g(M_FAC2)), num(g(M_SIE)), num(g(M_SIS))
        if sie is None and sis is None:
            ini = None                      # a relevar
        elif sie is not None and not fac:
            ini = FALTA
        else:
            ini = (sie or 0) * (fac or 0) + (sis or 0)
        items[cod] = dict(cod=cod, ins=norm(g(M_INS)), sec=norm(g(M_SEC)), ubi=norm(g(M_UBI)),
                          ub=norm(g(M_UB)), emp=norm(g(M_EMP)), fac=fac, emp2=norm(g(M_EMP2)), fac2=fac2,
                          ini=ini, sie=sie, sis=sis,
                          min=num(g(M_MIN)), obs=norm(g(M_OBS)), bar=norm(g(M_BAR)), fila=r)
    return items


def norm_bar(s):
    """Normaliza un código leído: solo letras y números, en mayúsculas.
    (Evita problemas de la lectora con '-' u otros signos según la distribución de teclado.)"""
    return re.sub(r"[^A-Z0-9]", "", norm(s).upper())


def base_qty(item, cant, unidad):
    u = (unidad or "").lower()
    if item["emp"] and u == item["emp"].lower():
        return cant * item["fac"] if item["fac"] else None
    if item.get("emp2") and u == item["emp2"].lower():
        return cant * item["fac2"] if item["fac2"] else None
    return cant


def unidades(item):
    """Unidades válidas para cargar: base, empaque 1, empaque 2."""
    return [u for u in (item["ub"], item["emp"], item.get("emp2")) if u]


def read_movements(wb, items):
    ws = wb[SH_MOV]
    out = []
    for r in range(2, last_row(ws) + 1):
        g = lambda c: ws.cell(row=r, column=c).value
        f, cod = g(1), norm(g(3)).upper()
        if not cod or cod not in items:
            continue
        if isinstance(f, dt.datetime):
            f = f.date()
        cant = num(g(5)) or 0
        out.append(dict(fecha=f, tipo=norm(g(2)).upper(), cod=cod, cant=cant, unidad=norm(g(6)),
                        base=base_qty(items[cod], cant, norm(g(6))), persona=norm(g(8)),
                        origen=norm(g(9)), obs=norm(g(10)), fila=r))
    return out


def compute_stock(items, movs):
    st = {}
    for cod, it in items.items():
        st[cod] = dict(ent=0.0, sal=0.0, falta=(it["ini"] == FALTA))
    for m in movs:
        s = st[m["cod"]]
        if m["base"] is None:
            s["falta"] = True
            continue
        if m["tipo"] == "ENTRADA":
            s["ent"] += m["base"]
        elif m["tipo"] == "SALIDA":
            s["sal"] += m["base"]
    for cod, it in items.items():
        s = st[cod]
        ini = it["ini"] if isinstance(it["ini"], (int, float)) else 0
        s["act"] = ini + s["ent"] - s["sal"]
        if s["falta"]:
            s["estado"] = "Falta cant. por empaque"
        elif it["ini"] is None:
            s["estado"] = "Sin stock inicial"
        elif s["act"] < 0:
            s["estado"] = "Negativo"
        elif it["min"] is not None and s["act"] <= it["min"]:
            s["estado"] = "Bajo mínimo"
        else:
            s["estado"] = "OK"
    return st


# ---------------------------------------------------------------- CSV
def parse_fecha(s, hoy):
    s = norm(s)
    if not s:
        return hoy
    s = s.replace("-", "/").replace(".", "/")
    p = s.split("/")
    try:
        if len(p) == 2:
            return dt.date(hoy.year, int(p[1]), int(p[0]))
        if len(p) == 3:
            if len(p[0]) == 4:                      # aaaa/mm/dd
                return dt.date(int(p[0]), int(p[1]), int(p[2]))
            y = int(p[2])
            return dt.date(y + 2000 if y < 100 else y, int(p[1]), int(p[0]))
    except ValueError:
        return None
    return None


def read_csv_rows(path):
    raw = None
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, encoding=enc) as fh:
                raw = fh.read()
            break
        except UnicodeDecodeError:
            continue
    lines = raw.splitlines()
    first = next((l for l in lines if l.strip() and not l.lstrip().startswith("#")), "")
    delim = ";" if first.count(";") >= first.count(",") else ","
    rows = []
    for n, row in enumerate(csv.reader(lines, delimiter=delim), start=1):
        if not row or not "".join(row).strip() or row[0].lstrip().startswith("#"):
            continue
        if row[0].strip().lower() == "fecha":
            continue
        rows.append((n, row + [""] * (7 - len(row))))
    return rows


def same_unit(entrada, unidad):
    """Acepta singular/plural y mayúsculas: caja = Cajas, unidad = unidades."""
    e, u = entrada.strip().lower(), unidad.strip().lower()
    return e in (u, u + "s", u + "es")


def validate(rows, items, hoy):
    ok, err = [], []
    for n, row in rows:
        fecha_s, tipo_s, cod_s, cant_s, uni_s, pers_s, obs_s = [norm(x) for x in row[:7]]
        e = []
        fecha = parse_fecha(fecha_s, hoy)
        if fecha is None:
            e.append(f"fecha '{fecha_s}' inválida")
        elif fecha > hoy:
            e.append(f"fecha {fecha:%d/%m/%Y} es futura")
        tipo = {"E": "ENTRADA", "ENTRADA": "ENTRADA", "S": "SALIDA", "SALIDA": "SALIDA"}.get(tipo_s.upper())
        if not tipo:
            e.append(f"tipo '{tipo_s}' (usar E o S)")
        cod = cod_s.upper()
        it = items.get(cod)
        if not it:
            e.append(f"código '{cod_s}' no existe en Maestro")
        cant = num(cant_s)
        if cant is None or cant <= 0:
            e.append(f"cantidad '{cant_s}' inválida")
        unidad = ""
        if it:
            match = [u for u in unidades(it) if same_unit(uni_s, u)] if uni_s else [it["ub"]]
            if match:
                unidad = match[0]
            else:
                opciones = " / ".join(unidades(it))
                e.append(f"unidad '{uni_s}' no corresponde a {cod} (usar {opciones} o dejar vacío)")
        if not pers_s:
            e.append("falta persona")
        if e:
            err.append(f"  Línea {n}: " + "; ".join(e))
        else:
            ok.append(dict(fecha=fecha, tipo=tipo, cod=cod, cant=cant, unidad=unidad,
                           persona=pers_s, obs=obs_s))
    return ok, err


def reset_csv():
    with open(CSV_IN, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write(";".join(CSV_HEADER) + "\n")
        fh.write("# ejemplo (las filas con # no se cargan):\n")
        fh.write("# 19/10/2026;S;CIN-07;8;;Juan Perez;\n")
        fh.write("# 20/10/2026;E;CIN-08;10;Caja;Ana Gomez;compra\n")


# ---------------------------------------------------------------- HTML
def fmt(v):
    if v is None or v == "":
        return "–"
    if isinstance(v, str):
        return v
    return f"{v:,.0f}".replace(",", ".") if abs(v - round(v)) < 1e-9 else f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def plural(w, n):
    w = (w or "").lower()
    if not w or n == 1:
        return w
    return w + ("s" if w[-1] in "aeiou" else "es")


def en_empaques(act, it):
    """'15 cajas + 1 rollo' (solo si hay empaque con cantidad y stock >= 0)."""
    if not it.get("emp") or not it.get("fac") or act is None or act < 0:
        return ""
    cer = int(act // it["fac"])
    sue = act - cer * it["fac"]
    t = f"{cer} {plural(it['emp'], cer)}"
    if sue:
        t += f" + {fmt(sue)} {plural(it['ub'], sue)}"
    return t


def sparkline(values, w=120, h=22):
    if not values or max(values) <= 0:
        return f'<svg width="{w}" height="{h}"></svg>'
    mx, n = max(values), len(values)
    bw = w / n
    bars = []
    for i, v in enumerate(values):
        bh = 0 if v <= 0 else max(1.5, (v / mx) * (h - 2))
        bars.append(f'<rect x="{i * bw + 1:.1f}" y="{h - bh:.1f}" width="{max(bw - 2, 1):.1f}" height="{bh:.1f}"><title>{v:g}</title></rect>')
    return f'<svg class="spark" width="{w}" height="{h}" viewBox="0 0 {w} {h}">{"".join(bars)}</svg>'


def build_html(items, movs, st, path=None, tema=None):
    hoy = dt.date.today()
    fechas = [m["fecha"] for m in movs if isinstance(m["fecha"], dt.date)]
    f0 = min(fechas) if fechas else hoy
    f1 = max(fechas) if fechas else hoy
    dias = (f1 - f0).days + 1
    lunes0 = f0 - dt.timedelta(days=f0.weekday())
    nsem = ((f1 - lunes0).days // 7) + 1
    sem_lbl = [(lunes0 + dt.timedelta(weeks=i)) for i in range(nsem)]

    sal_sem = defaultdict(lambda: [0.0] * nsem)
    n_ret = defaultdict(int)
    pers = defaultdict(lambda: defaultdict(int))
    tot_sem = [0] * nsem
    for m in movs:
        if m["tipo"] != "SALIDA" or not isinstance(m["fecha"], dt.date):
            continue
        k = (m["fecha"] - lunes0).days // 7
        if m["base"] is not None:
            sal_sem[m["cod"]][k] += m["base"]
        n_ret[m["cod"]] += 1
        tot_sem[k] += 1
        pers[m["persona"]][m["cod"]] += 1

    alert = [c for c in items if st[c]["estado"] != "OK"]
    n_sal = sum(1 for m in movs if m["tipo"] == "SALIDA")
    n_ent = sum(1 for m in movs if m["tipo"] == "ENTRADA")

    # tabla consumo
    rows_c = []
    orden = sorted(items, key=lambda c: (-n_ret[c], items[c]["ins"]))
    for c in orden:
        it, s = items[c], st[c]
        if n_ret[c] == 0:
            continue
        sal = s["sal"]
        prom_sem = sal / dias * 7 if dias else 0
        cob = ""
        if isinstance(it["ini"], (int, float)) and not s["falta"] and sal > 0 and s["act"] > 0:
            cob = f"{s['act'] / (sal / dias):.0f}"
        rows_c.append(
            f"<tr><td>{c}</td><td>{html.escape(it['ins'])}</td><td class=c>{html.escape(it['ub'])}</td>"
            f"<td class=n>{n_ret[c]}</td><td class=n>{'*' if s['falta'] else fmt(sal)}</td>"
            f"<td class=n>{'*' if s['falta'] else fmt(round(prom_sem, 1))}</td>"
            f"<td class=n>{cob or '–'}</td></tr>")

    # tabla stock
    rows_s = []
    for c in sorted(items, key=lambda c: (st[c]["estado"] == "OK", c)):
        it, s = items[c], st[c]
        ini = "A relevar" if it["ini"] is None else ("*" if it["ini"] == FALTA else fmt(it["ini"]))
        act = "*" if s["falta"] else fmt(s["act"])
        emp = "" if s["falta"] else en_empaques(s["act"], it)
        est = s["estado"]
        cls = " class=neg" if est == "Negativo" else ""
        mark = "" if est == "OK" else "* "
        rows_s.append(
            f"<tr><td>{c}</td><td>{html.escape(it['ins'])}</td><td>{html.escape(it['sec'])}</td>"
            f"<td class=c>{html.escape(it['ubi'])}</td><td class=c>{html.escape(it['ub'])}</td>"
            f"<td class=n>{ini}</td><td class=n>{fmt(s['ent'])}</td><td class=n>{'*' if s['falta'] else fmt(s['sal'])}</td>"
            f"<td class=n><b{cls}>{act}</b></td><td class=n>{emp or '–'}</td>"
            f"<td class=n>{fmt(it['min']) if it['min'] is not None else '–'}</td><td>{mark}{est}</td></tr>")

    # personas
    rows_p = []
    for p in sorted(pers, key=lambda p: -sum(pers[p].values())):
        top = sorted(pers[p].items(), key=lambda kv: -kv[1])[:3]
        det = ", ".join(f"{html.escape(items[c]['ins'])} ({n})" for c, n in top)
        rows_p.append(f"<tr><td>{html.escape(p)}</td><td class=n>{sum(pers[p].values())}</td><td>{det}</td></tr>")

    # últimos movimientos
    ult = sorted([m for m in movs if isinstance(m["fecha"], dt.date)], key=lambda m: m["fecha"], reverse=True)[:25]
    rows_u = "".join(
        f"<tr><td class=c>{m['fecha']:%d/%m/%Y}</td><td class=c>{'E' if m['tipo'] == 'ENTRADA' else 'S'}</td>"
        f"<td>{m['cod']}</td><td>{html.escape(items[m['cod']]['ins'])}</td><td class=n>{fmt(m['cant'])}</td>"
        f"<td class=c>{html.escape(m['unidad'])}</td><td>{html.escape(m['persona'])}</td></tr>" for m in ult)

    # barras retiros por semana
    mx = max(tot_sem) if tot_sem and max(tot_sem) else 1
    W, H, pad = 640, 150, 28
    bw = (W - pad) / max(nsem, 1)
    bars = []
    for i, v in enumerate(tot_sem):
        bh = (v / mx) * (H - 40)
        x = pad + i * bw
        bars.append(f'<rect x="{x + 6:.1f}" y="{H - 20 - bh:.1f}" width="{bw - 12:.1f}" height="{bh:.1f}"/>'
                    f'<text x="{x + bw / 2:.1f}" y="{H - 24 - bh:.1f}" class="lb">{v}</text>'
                    f'<text x="{x + bw / 2:.1f}" y="{H - 6}" class="ax">{sem_lbl[i]:%d/%m}</text>')
    chart = (f'<svg viewBox="0 0 {W} {H}" class="chart" role="img" aria-label="Retiros por semana">'
             f'<line x1="{pad}" y1="{H - 20}" x2="{W}" y2="{H - 20}" class="base"/>{"".join(bars)}</svg>')

    doc = f"""<!doctype html>
<html lang="es"{' data-theme="dark"' if tema == "oscuro" else ' data-theme="light"' if tema == "claro" else ""}><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Inventario Bajo Soppalco</title>
<link rel="icon" href="/favicon.ico" sizes="any">
<style>
:root{{--bg:#fff;--fg:#1d1d1d;--mut:#666;--ln:#d6d6d6;--hd:#f2f2f2;--ac:#3a3a3a;--neg:#b00020}}
@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{--bg:#161616;--fg:#e8e8e8;--mut:#9a9a9a;--ln:#383838;--hd:#222;--ac:#cfcfcf;--neg:#ff6b7f}}}}
:root[data-theme="dark"]{{--bg:#161616;--fg:#e8e8e8;--mut:#9a9a9a;--ln:#383838;--hd:#222;--ac:#cfcfcf;--neg:#ff6b7f}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--bg);color:var(--fg);font:13px/1.45 Arial,Helvetica,sans-serif}}
main{{max-width:1100px;margin:0 auto;padding:24px 16px 48px}}
h1{{font-size:20px;margin:0 0 2px}} h2{{font-size:14px;margin:28px 0 8px;text-transform:uppercase;letter-spacing:.04em;border-bottom:2px solid var(--ac);padding-bottom:4px}}
.sub{{color:var(--mut);margin:0 0 16px}}
.kpi{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));border:1px solid var(--ln)}}
.kpi div{{padding:10px 12px;border-right:1px solid var(--ln)}} .kpi div:last-child{{border-right:0}}
.kpi b{{display:block;font-size:22px}} .kpi span{{color:var(--mut);font-size:12px}}
.tw{{overflow-x:auto}}
table{{border-collapse:collapse;width:100%;min-width:640px}}
th,td{{border:1px solid var(--ln);padding:4px 6px;vertical-align:middle}}
th{{background:var(--hd);font-weight:bold;text-align:left;font-size:12px}}
td.n{{text-align:right;font-variant-numeric:tabular-nums}} td.c{{text-align:center}}
.neg{{color:var(--neg)}}
.spark rect{{fill:var(--ac)}}
.chart{{width:100%;max-width:700px;height:auto}} .chart rect{{fill:var(--ac)}}
.chart .lb,.chart .ax{{font-size:10px;text-anchor:middle;fill:var(--mut)}} .chart .base{{stroke:var(--ln)}}
.note{{color:var(--mut);font-size:12px;margin-top:6px}}
</style></head><body><main>
<h1>Inventario Bajo Soppalco</h1>
<p class="sub">Movimientos del {f0:%d/%m/%Y} al {f1:%d/%m/%Y} ({dias} días) · generado {dt.datetime.now():%d/%m/%Y %H:%M}</p>
<div class="kpi">
<div><b>{len(items)}</b><span>insumos</span></div>
<div><b>{n_sal}</b><span>salidas registradas</span></div>
<div><b>{n_ent}</b><span>entradas registradas</span></div>
<div><b>{len(alert)}</b><span>insumos con observación (*)</span></div>
</div>

<h2>Retiros por semana</h2>
{chart}
<p class="note">Cantidad de retiros (renglones) por semana, semana que empieza el lunes indicado.</p>

<h2>Consumo por insumo</h2>
<div class="tw"><table><thead><tr><th>Código</th><th>Insumo</th><th>Unidad</th><th>Retiros</th><th>Total salido</th>
<th>Prom. semanal</th><th>Cobertura (días)</th></tr></thead><tbody>{"".join(rows_c)}</tbody></table></div>
<p class="note">Cantidades en unidad base. Cobertura = stock actual / consumo diario promedio del período (solo si hay stock inicial).
* = no se puede calcular hasta completar "Cant. por empaque" en el Maestro.</p>

<h2>Stock actual</h2>
<div class="tw"><table><thead><tr><th>Código</th><th>Insumo</th><th>Sector</th><th>Ubic.</th><th>Unidad</th><th>Inicial</th>
<th>Entradas</th><th>Salidas</th><th>Actual</th><th>Empaques cerrados + sueltas</th><th>Mínimo</th><th>Estado</th></tr></thead><tbody>{"".join(rows_s)}</tbody></table></div>
<p class="note">* Sin stock inicial: insumo no relevado, el saldo arranca en 0 y da negativo. * Negativo: salió más de lo relevado (revisar conteo o conversión).</p>

<h2>Retiros por persona</h2>
<div class="tw"><table><thead><tr><th>Persona</th><th>Retiros</th><th>Insumos más retirados</th></tr></thead><tbody>{"".join(rows_p)}</tbody></table></div>

<h2>Últimos movimientos</h2>
<div class="tw"><table><thead><tr><th>Fecha</th><th>E/S</th><th>Código</th><th>Insumo</th><th>Cant.</th><th>Unidad</th><th>Persona</th></tr></thead><tbody>{rows_u}</tbody></table></div>
</main></body></html>"""
    if path:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(doc)
    return doc


# ---------------------------------------------------------------- main
def main():
    solo_html = "--solo-html" in sys.argv
    asegurar_excel()
    wb = openpyxl.load_workbook(XLSX)
    items = read_maestro(wb)
    hoy = dt.date.today()

    if not solo_html:
        if not os.path.exists(CSV_IN):
            reset_csv()
            sys.exit("No había movimientos.csv: se creó uno vacío. Completarlo y volver a correr.")
        rows = read_csv_rows(CSV_IN)
        if not rows:
            print("movimientos.csv está vacío: no hay nada para cargar. Se regenera solo el HTML.")
        else:
            ok, err = validate(rows, items, hoy)
            if err:
                print("ERRORES en movimientos.csv (no se cargó nada):")
                print("\n".join(err))
                sys.exit(1)
            ws = wb[SH_MOV]
            r = last_row(ws)
            for m in ok:
                r += 1
                write_movement(ws, r, m["fecha"], m["tipo"], m["cod"], m["cant"], m["unidad"],
                               m["persona"], f"CSV {hoy:%d/%m/%Y}", m["obs"])
            ws.auto_filter.ref = f"A1:J{r}"
            rebuild_stock(wb)
            try:
                wb.save(XLSX)
            except PermissionError:
                sys.exit("No se pudo guardar: CERRÁ el Excel y volvé a correr (no se cargó nada).")
            os.makedirs(CARGADOS, exist_ok=True)
            shutil.move(CSV_IN, os.path.join(CARGADOS, f"movimientos_{dt.datetime.now():%Y%m%d_%H%M%S}.csv"))
            reset_csv()
            print(f"Cargados {len(ok)} movimientos. CSV archivado en herramientas/cargados.")

    movs = read_movements(wb, items)
    st = compute_stock(items, movs)
    neg = [c for c in items if st[c]["estado"] == "Negativo"]
    if neg:
        print("Atención, stock negativo: " + ", ".join(f"{c} {items[c]['ins']}" for c in neg))
    build_html(items, movs, st, HTML_OUT)
    print(f"HTML generado: {HTML_OUT}")


if __name__ == "__main__":
    main()
