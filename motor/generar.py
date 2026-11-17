"""Motor del plan de email: marca.json -> Excel con fórmulas vivas + cifras.json + defensa.md.

Uso:  python generar.py <marca.json> [carpeta_salida]

Las cifras de cifras.json se calculan en Python con la MISMA lógica que las fórmulas del Excel;
test_motor.py comprueba que coinciden después de recalcular el Excel con LibreOffice.
Formato de marca.json: references/formato-marca-json.md.
"""
import json
import re
import sys
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

TIPOS = ["Oferta", "Inspiracional", "Producto", "Fidelización", "Trigger", "Contenido"]
PROG = [t for t in TIPOS if t != "Trigger"]  # los triggers no se programan: van en "triggers"
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
OFERTA_PAL = ("%", "descuento", "rebaja", "oferta", "promo", "outlet", "black friday", "cyber")


def fecha(d):
    return f"{d.day}-{MESES[d.month - 1]}-{d.year}"


def xround(x):  # ROUND(x;0) de Excel (mitad hacia arriba), no el redondeo bancario de Python
    return int(x + 0.5)


def recibe_de(m, k):
    return m["sin_perfil"]["recibe"] if k == "S" else k


def n_semana(cal, w, k):
    return sum(1 for r in cal if r[0] == w and r[1] in (k, "Todos"))


def aplica(t, k):
    return t["perfiles"] != "Nuevos" and k in [x.strip() for x in t["perfiles"].split(",")]


def meses(m):
    """Bloques de semanas por mes (el mes del jueves de cada semana): [[etiqueta, sem_ini, sem_fin], ...]."""
    d0, out = date.fromisoformat(m["inicio"]), []
    for w in range(1, m["semanas"] + 1):
        mes = (d0 + timedelta(days=7 * (w - 1) + 3)).month
        if out and out[-1][0] == mes:
            out[-1][2] = w
        else:
            out.append([mes, w, w])
    return [[f"{MESES[me - 1].upper()} · sem {a}" + (f"-{b}" if b != a else ""), a, b] for me, a, b in out]


# ------------------------------------------------------------------ validación
def validar(m):
    """Devuelve (errores, avisos). Con errores no se genera nada."""
    err, av = [], []
    W, P, cal, S = m["semanas"], m["perfiles"], m["calendario"], m["supuestos"]
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,30}", str(m.get("slug", ""))):
        err.append("slug: solo letras, números, - y _ (se usa en los nombres de archivo)")
    textos = [p["nombre"] for p in P] + list(m["eventos"]) + [t["nombre"] for t in m["triggers"]] \
        + [str(x) for r in cal for x in r[4:7]] + [str(x) for r in m["medicion"] for x in r] + [str(v) for v in m["textos"].values()]
    for t in textos:  # el Excel convertiría en fórmula un texto que empiece por '='
        if t.startswith("="):
            err.append(f"texto que empieza por '=' (el Excel lo tomaría por fórmula): «{t[:40]}»")
    e = m.get("deck", {}).get("estilo", {})
    for k, v in e.items():
        if k == "tipos":
            malos = [c for c in v.values() if not re.fullmatch(r"[0-9A-Fa-f]{6}", c)]
        elif k.startswith("fuente"):
            malos = [] if re.fullmatch(r"[\w .-]+", v) else [v]
        else:
            malos = [] if re.fullmatch(r"[0-9A-Fa-f]{6}", v) else [v]
        if malos:
            err.append(f"deck.estilo.{k}: {malos} (colores en hex de 6 cifras sin #; fuentes, solo letras y espacios)")
    codes = [p["k"] for p in P]
    if date.fromisoformat(m["inicio"]).weekday() != 0:
        err.append(f"inicio {m['inicio']} no es lunes")
    if not 1 <= len(P) <= 4:
        err.append(f"{len(P)} perfiles: el motor admite de 1 a 4 (la rueda no cabe con más)")
    for k in codes:
        if not (len(k) == 1 and k.isalpha() and k.isupper() and k != "S"):
            err.append(f"código de perfil '{k}': tiene que ser una letra mayúscula distinta de S")
    if len(set(codes)) != len(codes):
        err.append("hay códigos de perfil repetidos")
    tot = sum(p["pct"] for p in P) + m["sin_perfil"]["pct"]
    if abs(tot - 1) > 1e-9:
        err.append(f"los % de perfiles + sin perfil suman {tot:.1%}, no 100 %")
    if m["sin_perfil"]["recibe"] not in codes:
        err.append(f"sin_perfil.recibe = '{m['sin_perfil']['recibe']}' no es ningún perfil")
    if len(m["eventos"]) != W:
        err.append(f"hay {len(m['eventos'])} eventos y {W} semanas: uno por semana")
    if sorted(t["tipo"] for t in m["tipos"]) != sorted(TIPOS):
        err.append(f"'tipos' tiene que definir exactamente estos 6: {', '.join(TIPOS)}")
    for i, r in enumerate(cal, 1):
        if len(r) != 7:
            err.append(f"calendario fila {i}: necesita 7 campos [semana, perfil, tipo, día, asunto, objetivo, test]")
            continue
        if not (isinstance(r[0], int) and 1 <= r[0] <= W):
            err.append(f"calendario fila {i}: semana {r[0]} fuera de 1-{W}")
        if r[1] not in codes and r[1] != "Todos":
            err.append(f"calendario fila {i}: perfil '{r[1]}' no existe")
        if r[2] not in PROG:
            err.append(f"calendario fila {i}: tipo '{r[2]}' no programable (los Trigger van en 'triggers')")
        if r[3] not in DIAS:
            err.append(f"calendario fila {i}: día '{r[3]}' (usa {'/'.join(DIAS)})")
    if err:
        return err, av
    for k in codes:
        for w in range(1, W + 1):
            n = n_semana(cal, w, k)
            if n > S["tope_semanal"]:
                err.append(f"perfil {k}, semana {w}: {n} programados (tope {S['tope_semanal']})")
            if n == 0:
                av.append(f"perfil {k}, semana {w}: sin email programado")
    for p in P:  # la contradicción que destapó la auditoría de TNF: el perfil «sin descuentos» recibía descuentos
        k, suyos = p["k"], [r for r in cal if r[1] in (p["k"], "Todos")]
        if "max_ofertas" in p:
            n = sum(1 for r in suyos if r[2] == "Oferta")
            if n > p["max_ofertas"]:
                err.append(f"perfil {k} admite {p['max_ofertas']} ofertas y el calendario le manda {n}: contradice su barrera")
            for r in suyos:  # descuento colado en un email que no es de tipo Oferta
                if r[2] != "Oferta" and any(w in r[4].lower() for w in OFERTA_PAL):
                    av.append(f"perfil {k}, semana {r[0]}: «{r[4]}» suena a oferta y su barrera las limita")
        tit = p.get("tarjeta", {}).get("titular")
        if tit and not any(tit.lower() in r[4].lower() for r in suyos):
            av.append(f"perfil {k}: su titular «{tit}» no está en ningún asunto del calendario (el perfil no se llega a probar)")
        if not p.get("refutacion"):
            av.append(f"perfil {k}: falta 'refutacion' (qué resultado tumbaría el perfil)")
    vistos = {}
    for r in cal:
        if r[4] in vistos and vistos[r[4]] != r[1]:
            av.append(f"asunto idéntico en {vistos[r[4]]} y {r[1]}: «{r[4]}» (prueba del mensaje único)")
        vistos.setdefault(r[4], r[1])
    for t in m["triggers"]:
        if t["perfiles"] != "Nuevos":
            malos = [x.strip() for x in t["perfiles"].split(",") if x.strip() not in codes]
            if malos:
                err.append(f"trigger «{t['nombre']}»: perfiles inexistentes {malos}")
        if not 0 <= t["tasa"] <= 1:
            err.append(f"trigger «{t['nombre']}»: tasa {t['tasa']} fuera de 0-1")
    if "deck" in m and len(m["deck"].get("meses", [])) != len(meses(m)):
        err.append(f"deck.meses tiene {len(m['deck'].get('meses', []))} textos y el periodo abarca {len(meses(m))} meses: "
                   + ", ".join(x[0] for x in meses(m)))
    if "notas" in m and len(m["notas"]) != 7:
        err.append(f"notas: {len(m['notas'])} y el deck tiene 7 diapositivas")
    return err, av


# ------------------------------------------------------------------ cálculo (espejo de las fórmulas)
def calcular(m):
    S, W, cal, P = m["supuestos"], m["semanas"], m["calendario"], m["perfiles"]
    prog = {p["k"]: sum(1 for r in cal if r[1] in (p["k"], "Todos")) for p in P}
    pico = {k: sum(1 for w in range(1, W + 1) if n_semana(cal, w, k) >= S["tope_semanal"]) for k in prog}

    def trig_pp(k):  # triggers esperados por persona del perfil k en todo el periodo
        return sum(t["tasa"] * (t["emails"] - (pico[k] / W if t.get("tope") else 0)) for t in m["triggers"] if aplica(t, k))

    filas = [(p["k"], p["nombre"], p["pct"]) for p in P] + [("S", m["sin_perfil"]["nombre"], m["sin_perfil"]["pct"])]
    perfiles = []
    for k, nombre, pct in filas:
        kp = recibe_de(m, k)
        aud = S["base"] * pct
        rec = xround(aud * (1 - S["control"]))
        ip, it = prog[kp] * rec, aud * trig_pp(kp)
        perfiles.append(dict(k=k, nombre=nombre, pct=pct, programa_de=kp, audiencia=aud, recibe=rec, programados=prog[kp],
                             picos=pico[kp], imp_programados=ip, imp_triggers=it, imp_total=ip + it,
                             emails_por_persona=prog[kp] + (it / aud if aud else 0)))
    e_bien = sum(t["emails"] for t in m["triggers"] if t["perfiles"] == "Nuevos")
    bien = S["nuevos"] * e_bien
    imp_p = sum(p["imp_programados"] for p in perfiles)
    imp_t = sum(p["imp_triggers"] for p in perfiles) + bien
    piezas = dict(unicos=sum(1 for r in cal if r[1] != "Todos"), variantes=sum(1 for r in cal if r[6]),
                  comunes=sum(1 for r in cal if r[1] == "Todos"), plantillas=sum(t["plantillas"] for t in m["triggers"]))
    piezas["total"] = sum(piezas.values())
    return dict(marca=m["marca"], semanas=W, emails_programados=len(cal), perfiles=perfiles,
                bienvenida=dict(altas=S["nuevos"], emails=e_bien, impactos=bien),
                impactos=dict(programados=imp_p, triggers=imp_t, total=imp_p + imp_t), piezas=piezas,
                tests=[[r[0], r[1], r[6]] for r in cal if r[6]], meses=meses(m),
                por_tipo={k: {t: sum(1 for r in cal if r[1] in (k, "Todos") and r[2] == t) for t in PROG} for k in prog})


# ------------------------------------------------------------------ Excel
NAVY = "1F3A68"
HF, HFONT = PatternFill("solid", fgColor=NAVY), Font(bold=True, color="FFFFFF")
YEL, NOTE, KEY, WARN = (PatternFill("solid", fgColor=c) for c in ("FFF2CC", "EEF3FA", "E2EFDA", "FCE4D6"))
TCOL = {"Oferta": "F8CBAD", "Inspiracional": "D9E1F2", "Producto": "FFE699", "Fidelización": "E2EFDA", "Trigger": "E4DFEC", "Contenido": "DDEBF7"}
_th = Side(style="thin", color="BFBFBF")
BOX = Border(left=_th, right=_th, top=_th, bottom=_th)
WRAP = Alignment(wrap_text=True, vertical="top")
N0, P0 = "#,##0", "0%"


def title(ws, t, expl, n=6):
    ws["A1"], ws["A1"].font = t, Font(bold=True, size=14, color=NAVY)
    ws["A2"], ws["A2"].alignment, ws["A2"].fill = expl, WRAP, NOTE
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[2].height = 15 * max(2, len(expl) // 110 + 1)


def header(ws, r, labels):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=r, column=i, value=lab)
        c.fill, c.font, c.border = HF, HFONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put(ws, r, vals, fmt=None, fill=None):
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=i, value=v if v != "" else None)
        c.border, c.alignment = BOX, WRAP
        if fmt and i in fmt:
            c.number_format = fmt[i]
        if fill and i in fill:
            c.fill = fill[i]


def bold(ws, r, n):
    for c in range(1, n + 1):
        ws.cell(row=r, column=c).font = Font(bold=True)


def widths(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v


def note(ws, cell, text, fill=NOTE, n=6):
    ws[cell], ws[cell].alignment, ws[cell].fill = text, WRAP, fill
    r, c = ws[cell].row, ws[cell].column
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + n - 1)
    ws.row_dimensions[r].height = 15 * max(2, len(text) // 105 + 1)


def excel(m, out):
    S_, W, cal, P, TX = m["supuestos"], m["semanas"], m["calendario"], m["perfiles"], m["textos"]
    N, T = len(P), len(m["triggers"])
    codes = [p["k"] for p in P]
    d0 = date.fromisoformat(m["inicio"])
    dfin = d0 + timedelta(days=7 * W - 1)
    filas = [(p["k"], p["nombre"]) for p in P] + [("S", m["sin_perfil"]["nombre"])]
    wb = Workbook()
    R = wb.active
    R.title = "Resumen"

    # Supuestos: B4 base, B5 control, B6 tope; perfiles desde la fila 8; triggers desde TR0
    DR = m.get("datos_reales")  # si hay datos de la marca, base y reparto salen de la hoja Datos_reales
    S = wb.create_sheet("Supuestos")
    if DR:
        expl = (f"Base y reparto entre perfiles: DATOS REALES de {DR['fuente']} (importados el {DR['fecha']}; hoja Datos_reales). "
                "Lo que sigue en amarillo es supuesto y se puede editar. ")
    else:
        expl = (f"{TX['fuente_supuestos']}: el tamaño de la base y el reparto entre perfiles son SUPUESTOS para poder calcular impactos. "
                "Con datos reales se cambian aquí y nada más (o con motor/importar.py). ")
    title(S, "Supuestos y datos (celdas amarillas = editables; todo el libro se recalcula)",
          expl + f"El grupo de control ({S_['control']:.0%}) se aparta ANTES de segmentar y no recibe los envíos programados "
          "(skill medicion-incremental, cap. 5-6).", 6)
    RTD = 6 + N  # fila del total en Datos_reales
    put(S, 4, ["Suscriptores activos de email", f"=Datos_reales!C{RTD}" if DR else S_["base"]], {2: N0}, None if DR else {2: YEL})
    put(S, 5, ["Grupo de control (no recibe el programa)", S_["control"]], {2: P0}, {2: YEL})
    put(S, 6, ["Tope de emails programados por semana y perfil", S_["tope_semanal"]], None, {2: YEL})
    header(S, 7, ["Perfil", "Código", "% de la base", "Audiencia total", "Audiencia que recibe (sin control)", "Origen"])
    pcts = [p["pct"] for p in P] + [m["sin_perfil"]["pct"]]
    for i, ((k, nombre), pct) in enumerate(zip(filas, pcts)):
        r = 8 + i
        put(S, r, [nombre, k, f"=Datos_reales!D{5 + i}" if DR else pct, f"=$B$4*C{r}", f"=ROUND(D{r}*(1-$B$5),0)",
                   "Dato real" if DR else "Supuesto"], {3: P0, 4: N0, 5: N0}, None if DR else {3: YEL})
    RT = 9 + N  # fila del total
    put(S, RT, ["Total", "", f"=SUM(C8:C{RT - 1})", f"=SUM(D8:D{RT - 1})", f"=SUM(E8:E{RT - 1})"], {3: P0, 4: N0, 5: N0})
    bold(S, RT, 5)
    S.conditional_formatting.add(f"C{RT}", CellIsRule(operator="notEqual", formula=["1"], fill=PatternFill("solid", fgColor="FFC7CE")))
    if DR:
        ant = DR["antes"]["pct"]
        note(S, f"A{RT + 1}", f"Reparto real: {DR['regla']}. Supuesto anterior: "
             + " / ".join(f"{k} {ant[k]:.0%}" for k, _ in filas) + ". La diferencia es lo que el dato corrige de la hipótesis.", KEY, 6)
    else:
        note(S, f"A{RT + 1}", f"Reparto {' / '.join(f'{x * 100:g}' for x in pcts)} %. {TX['nota_reparto']}", WARN, 6)
    header(S, 12 + N, ["Trigger (always-on)", "Perfiles", f"Tasa de activación en {W} semanas", "Emails por activación",
                       "Plantillas (piezas)", "Tope (−1 email en semanas pico)", "Origen de la tasa"])
    TR0 = 13 + N
    for j, t in enumerate(m["triggers"]):
        put(S, TR0 + j, [t["nombre"], t["perfiles"], t["tasa"], t["emails"], t["plantillas"], "sí" if t.get("tope") else "",
                         "Dato real" if t.get("real") else ("No se usa" if t["perfiles"] == "Nuevos" else "Supuesto")],
            {3: P0}, {3: YEL, 4: YEL})
    if DR:  # hoja con el conteo real, para que se vea de dónde sale cada %
        X = wb.create_sheet("Datos_reales")
        title(X, f"Datos reales de la marca · {DR['fuente']}",
              f"Importados el {DR['fecha']} con motor/importar.py. Regla: {DR['regla']}. Las hojas Supuestos e Impactos leen de aquí.", 5)
        header(X, 4, ["Perfil", "Código", "Clientes", "% real", "% supuesto antes"])
        for i, (k, nombre) in enumerate(filas):
            put(X, 5 + i, [nombre, k, DR["conteo"][k], f"=C{5 + i}/$C${RTD}", DR["antes"]["pct"][k]], {3: N0, 4: "0.0%", 5: P0})
        put(X, RTD, ["Total", "", f"=SUM(C5:C{RTD - 1})", f"=SUM(D5:D{RTD - 1})", ""], {3: N0, 4: "0.0%"})
        bold(X, RTD, 5)
        if DR["antes"].get("triggers"):
            header(X, RTD + 2, ["Trigger", "", "Tasa real", "Tasa supuesta antes", ""])
            for j, (n, v) in enumerate(DR["antes"]["triggers"].items()):
                put(X, RTD + 3 + j, [n, "", next(t["tasa"] for t in m["triggers"] if t["nombre"] == n), v, ""], {3: "0.0%", 4: P0})
        widths(X, {"A": 40, "B": 9, "C": 14, "D": 12, "E": 18})
    RN = TR0 + T  # nuevos suscriptores
    put(S, RN, [f"Nuevos suscriptores en las {W} semanas", "", S_["nuevos"]], {3: N0}, {3: YEL})
    note(S, f"A{RN + 2}", "Las tasas de activación de los triggers son supuestos de orden de magnitud (no benchmarks): con la plataforma de "
         "envíos se sustituyen por las reales. Los triggers con perfil 'Nuevos' los recibe cada alta (su tasa no se usa). Los triggers "
         "se envían también al grupo de control, porque responden a una acción del propio cliente (como los transaccionales); lo que "
         "se mide con el control es el programa planificado.", WARN, 5)
    widths(S, {"A": 46, "B": 14, "C": 18, "D": 16, "E": 22, "F": 16, "G": 16})

    # Tipos_Email
    Tw = wb.create_sheet("Tipos_Email")
    title(Tw, f"Pregunta 1 · Los 6 tipos de email y cómo los usa {m['marca']}", TX["tipos_intro"], 7)
    header(Tw, 4, ["Tipo", "Qué es", "Variables a cuidar", "KPI principal", "Uso principal en el plan", "Cuándo NO", f"Ejemplo {m['slug']}"])
    for i, t in enumerate(m["tipos"]):
        put(Tw, 5 + i, [t["tipo"], t["que"], t["variables"], t["kpi"], t["uso"], t["cuando_no"], t["ejemplo"]],
            fill={1: PatternFill("solid", fgColor=TCOL[t["tipo"]])})
        Tw.cell(row=5 + i, column=1).font = Font(bold=True)
    widths(Tw, {"A": 14, "B": 30, "C": 42, "D": 20, "E": 34, "F": 34, "G": 30})

    # Calendario
    C = wb.create_sheet("Calendario")
    title(C, f"Pregunta 2 · Programa de comunicación: {W} semanas ({fecha(d0)} a {fecha(dfin)}), email por perfil",
          TX["reglas_calendario"], 10)
    header(C, 4, ["Semana", "Lunes", "Evento", "Perfil", "Tipo", "Día", "Asunto / tema", "Objetivo / KPI", "Test A/B", "Nº piezas"])
    for i, (w, p, tipo, dia, asunto, obj, ab) in enumerate(cal):
        r = 5 + i
        put(C, r, [w, d0 + timedelta(days=7 * (w - 1)), m["eventos"][w - 1], p, tipo, dia, asunto, obj, ab, f'=1+IF(I{r}<>"",1,0)'],
            {2: "dd-mmm-yy"}, {5: PatternFill("solid", fgColor=TCOL[tipo])})
    LAST = 4 + len(cal)
    widths(C, {"A": 8, "B": 11, "C": 36, "D": 7, "E": 13, "F": 6, "G": 52, "H": 26, "I": 26, "J": 9})
    C.freeze_panes = "A5"
    CR, CP, CT, CI = (f"Calendario!${c}$5:${c}${LAST}" for c in "ADEI")

    # Matriz_Semanal: perfiles en columnas D.. ; semanas en filas 5..4+W
    M = wb.create_sheet("Matriz_Semanal")
    col = {k: L(4 + i) for i, k in enumerate(codes)}
    title(M, "Vista semana a semana: emails programados que recibe cada perfil",
          "Cuenta con fórmulas sobre la hoja Calendario (los 'Todos' suman a todos los perfiles).", 3 + N)
    header(M, 4, ["Semana", "Lunes", "Evento"] + [f"{p['k']} · {p['corto']}" for p in P])
    for w in range(1, W + 1):
        r = 4 + w
        put(M, r, [w, d0 + timedelta(days=7 * (w - 1)), m["eventos"][w - 1]]
            + [f'=COUNTIFS({CR},$A{r},{CP},"{k}")+COUNTIFS({CR},$A{r},{CP},"Todos")' for k in codes], {2: "dd-mmm-yy"})
    MT = 5 + W
    put(M, MT, ["Total", "", f"{W} semanas"] + [f"=SUM({col[k]}5:{col[k]}{MT - 1})" for k in codes])
    put(M, MT + 1, ["Media semanal", "", ""] + [f"={col[k]}{MT}/{W}" for k in codes], {4 + i: "0.00" for i in range(N)})
    bold(M, MT, 3 + N)
    bold(M, MT + 1, 3 + N)
    header(M, MT + 3, ["Tipo de email", "", ""] + [f"{p['k']} · {p['corto']}" for p in P])
    for i, t in enumerate(PROG):
        r = MT + 4 + i
        put(M, r, [t, "", ""] + [f'=COUNTIFS({CP},"{k}",{CT},$A{r})+COUNTIFS({CP},"Todos",{CT},$A{r})' for k in codes],
            fill={1: PatternFill("solid", fgColor=TCOL[t])})
    if TX.get("lectura_matriz"):
        note(M, f"A{MT + 4 + len(PROG) + 1}", TX["lectura_matriz"], KEY, 3 + N)
    widths(M, {"A": 14, "B": 11, "C": 44, **{col[k]: 15 for k in codes}})

    # Impactos: filas 5..5+N (perfiles + sin perfil), bienvenida, total; luego emails por persona
    I = wb.create_sheet("Impactos")
    title(I, "Pregunta 3 · Impactos: cuántos emails recibe la audiencia",
          "Impacto = un email entregado a una persona. Se calculan por perfil: emails programados por persona × audiencia que los recibe "
          "(sin el control), más los triggers según su tasa de activación.", 6)
    header(I, 4, ["Perfil", "Emails programados por persona", "Audiencia que recibe", "Impactos programados", "Impactos por triggers", "Impactos totales"])
    for i, (k, nombre) in enumerate(filas):
        r, kp = 5 + i, recibe_de(m, k)
        c = col[kp]
        partes = []
        for j, t in enumerate(m["triggers"]):
            if t["perfiles"] == "Nuevos":
                continue
            tr = TR0 + j
            emails = f"Supuestos!$D${tr}"
            if t.get("tope"):  # en las semanas que ya llegan al tope de programados, este trigger manda un email menos
                emails = f'({emails}-COUNTIF(Matriz_Semanal!{c}5:{c}{MT - 1},">="&Supuestos!$B$6)/{W})'
            partes.append(f'IF(ISNUMBER(SEARCH(",{kp},",","&SUBSTITUTE(Supuestos!$B${tr}," ","")&",")),Supuestos!$C${tr}*{emails},0)')
        etiqueta = f"{k} · {nombre}" if k != "S" else f"{nombre}: recibe el programa de {kp}"
        put(I, r, [etiqueta, f"=Matriz_Semanal!{c}{MT}", f"=Supuestos!E{8 + i}", f"=B{r}*C{r}",
                   f"=Supuestos!D{8 + i}*({'+'.join(partes) or '0'})", f"=D{r}+E{r}"], {3: N0, 4: N0, 5: N0, 6: N0})
    RB, RTI = 6 + N, 7 + N
    put(I, RB, ["Bienvenida (nuevos suscriptores)", "", f"=Supuestos!C{RN}", "",
                f'=Supuestos!C{RN}*SUMIF(Supuestos!$B${TR0}:$B${RN - 1},"Nuevos",Supuestos!$D${TR0}:$D${RN - 1})', f"=E{RB}"],
        {3: N0, 5: N0, 6: N0})
    put(I, RTI, ["TOTAL", "", "", f"=SUM(D5:D{RB - 1})", f"=SUM(E5:E{RB})", f"=SUM(F5:F{RB})"], {4: N0, 5: N0, 6: N0})
    bold(I, RTI, 6)
    header(I, RTI + 2, ["Perfil", f"Emails por persona en {W} semanas (programados + triggers esperados)", "Por semana", "", "", ""])
    RPP = RTI + 3
    for i in range(N + 1):
        r = RPP + i
        put(I, r, [f"=A{5 + i}", f"=IF(Supuestos!D{8 + i}=0,B{5 + i},B{5 + i}+E{5 + i}/Supuestos!D{8 + i})", f"=B{r}/{W}"], {2: "0.0", 3: "0.00"})
    note(I, f"A{RPP + N + 2}", "Cómo se lee: cada persona que recibe el programa recibe los emails programados de la hoja Calendario, más los "
         "triggers que le tocan de media (solo llegan a quien hace la acción: carrito, navegación, compra; por eso se calculan con una tasa).", KEY, 6)
    widths(I, {"A": 40, "B": 22, "C": 18, "D": 18, "E": 18, "F": 18})

    # Piezas
    Pz = wb.create_sheet("Piezas")
    TC = L(2 + N)
    title(Pz, "Pregunta 4 · Total de piezas que hay que producir",
          "Pieza = una creatividad distinta (diseño + copy). Un email del calendario = 1 pieza; si lleva test A/B = 2 piezas (cambia solo "
          "la primera línea o el asunto, cap. 14). Un email 'Todos' es 1 pieza para todos los perfiles. Cada trigger cuenta sus plantillas "
          "(una por perfil al que aplica), no sus envíos.", 2 + N)
    header(Pz, 4, ["Concepto"] + codes + ["Total"])
    put(Pz, 5, ["Emails programados (únicos)"] + [f'=COUNTIFS({CP},"{k}")' for k in codes] + [f"=SUM(B5:{L(1 + N)}5)"])
    put(Pz, 6, ["Variantes extra por test A/B"] + [f'=COUNTIFS({CP},"{k}",{CI},"<>")' for k in codes] + [f'=COUNTIF({CI},"<>")'])
    put(Pz, 7, ["Emails comunes ('Todos'): 1 pieza para todos"] + [""] * N + [f'=COUNTIFS({CP},"Todos")'])
    put(Pz, 8, ["Plantillas de triggers"] + [""] * N + [f"=SUM(Supuestos!E{TR0}:E{RN - 1})"])
    put(Pz, 9, ["TOTAL PIEZAS"] + [""] * N + [f"=SUM({TC}5:{TC}8)"])
    bold(Pz, 9, 2 + N)
    header(Pz, 11, ["Piezas por tipo de email"] + [""] * N + ["Total"])
    for i, t in enumerate(PROG):
        put(Pz, 12 + i, [t] + [""] * N + [f'=COUNTIFS({CT},"{t}")+COUNTIFS({CT},"{t}",{CI},"<>")'],
            fill={1: PatternFill("solid", fgColor=TCOL[t])})
    put(Pz, 17, ["Trigger (plantillas)"] + [""] * N + [f"={TC}8"], fill={1: PatternFill("solid", fgColor=TCOL["Trigger"])})
    put(Pz, 18, ["Comprobación (debe coincidir con el total)"] + [""] * N + [f"=SUM({TC}12:{TC}17)"])
    Pz.conditional_formatting.add(f"{TC}18", CellIsRule(operator="notEqual", formula=[f"${TC}$9"], fill=PatternFill("solid", fgColor="FFC7CE")))
    widths(Pz, {"A": 44, **{L(2 + i): 8 for i in range(N)}, TC: 12})

    # Medicion
    D = wb.create_sheet("Medicion")
    title(D, "Cómo se mide si el programa funciona (no se presenta nada sin esto)",
          "Skill medicion-incremental (cap. 5-8). Sin grupo de control, cualquier venta de los picos se apuntaría como mérito del email, "
          "cuando gran parte iba a ocurrir igual.", 3)
    filas_med = m["medicion"] + [[f"Refutación {p['k']} · {p['corto']}", p["refutacion"]] for p in P if p.get("refutacion")]
    for i, (a, b) in enumerate(filas_med):
        put(D, 4 + i, [a, b])
    widths(D, {"A": 22, "B": 110})

    # Resumen
    per = f"{MESES[d0.month - 1]}-{d0.year} a {MESES[dfin.month - 1]}-{dfin.year}"
    title(R, f"{m['marca']} {m['mercado']} · Plan de email {per} (respuestas)",
          "Cifras calculadas con fórmulas desde las hojas siguientes. Cambia los supuestos (hoja Supuestos) y todo se recalcula.", 3)
    header(R, 4, ["Pregunta", "Respuesta", "Hoja"])
    sp = m["sin_perfil"]
    j = lambda xs: '&" / "&'.join(xs)
    rows = [
        ("Buyer personas", " · ".join(f"{p['k']} · {p['nombre']} ({p['pct']:.0%})" for p in P)
         + f" · sin perfil ({sp['pct']:.0%}, recibe el programa de {sp['recibe']}). Variable que discrimina: {TX['variable_discrimina']}.",
         "Supuestos + documento de buyer personas"),
        ("1 · Propuesta de contenidos", TX["resumen_contenidos"], "Tipos_Email"),
        (f"2 · Programa {W} semanas", f'=COUNTA(Calendario!$A$5:$A${LAST})&" emails programados en {W} semanas ({fecha(d0)} a {fecha(dfin)}), '
         f'semana a semana, con {TX["resumen_eventos"].replace(chr(34), chr(34) * 2)}"', "Calendario · Matriz_Semanal"),
        (f"Emails programados por persona ({' / '.join(codes)})", "=" + j([f"Matriz_Semanal!{col[k]}{MT}" for k in codes]), "Matriz_Semanal"),
        ("3 · Impactos totales", f'=TEXT(Impactos!F{RTI},"#,##0")&" impactos ("&TEXT(Impactos!D{RTI},"#,##0")&" programados + "'
         f'&TEXT(Impactos!E{RTI},"#,##0")&" por triggers)"', "Impactos"),
        (f"Emails por persona en {W} semanas ({' / '.join(codes)} / sin perfil)", "=" + j([f"ROUND(Impactos!B{RPP + i},1)" for i in range(N + 1)]), "Impactos"),
        ("4 · Total de piezas", f'=Piezas!{TC}9&" piezas ("&Piezas!{TC}5&" emails únicos + "&Piezas!{TC}6&" variantes A/B + "&Piezas!{TC}7'
         f'&" comunes + "&Piezas!{TC}8&" plantillas de trigger)"', "Piezas"),
        ("Medición", TX["resumen_medicion"], "Medicion"),
    ]
    for i, row in enumerate(rows):
        put(R, 5 + i, list(row))
        R.row_dimensions[5 + i].height = 45
    widths(R, {"A": 34, "B": 100, "C": 26})
    wb.save(out)
    # celdas que test_motor.py compara con calcular() tras recalcular con LibreOffice
    return {"impactos_total": f"Impactos!F{RTI}", "impactos_programados": f"Impactos!D{RTI}", "piezas_total": f"Piezas!{TC}9",
            "piezas_check": f"Piezas!{TC}18", "suma_pct": f"Supuestos!C{RT}",
            **{f"programados_{k}": f"Matriz_Semanal!{col[k]}{MT}" for k in codes},
            **{f"imp_total_{k}": f"Impactos!F{5 + i}" for i, (k, _) in enumerate(filas)},
            **{f"emails_pp_{k}": f"Impactos!B{RPP + i}" for i, (k, _) in enumerate(filas)}}


# ------------------------------------------------------------------ chuleta de defensa
def es(x, dec=0):
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def defensa(m, c):
    S, W = m["supuestos"], m["semanas"]
    p = c["perfiles"][0]
    trig = [t for t in m["triggers"] if aplica(t, p["k"])]
    tz = c["piezas"]
    pp = [x["emails_por_persona"] for x in c["perfiles"]]
    DR = m.get("datos_reales")
    if DR:
        ant = DR["antes"]["pct"]
        origen = (f"- **¿De dónde sale la base de {es(S['base'])}?** De los datos de la marca ({DR['fuente']}, importados el {DR['fecha']}).\n"
                  f"- **¿Por qué ese reparto entre perfiles?** Sale de {DR['regla']}. Suponíamos "
                  + " / ".join(f"{x['k']} {ant[x['k']]:.0%}" for x in c["perfiles"]) + "; el dato dice "
                  + " / ".join(f"{x['k']} {x['pct']:.0%}" for x in c["perfiles"]) + ". Esa diferencia es lo que hemos aprendido de la base.")
        demo = "Cambia un conteo en la hoja Datos_reales:"
    else:
        origen = (f"- **¿De dónde sale la base de {es(S['base'])}?** Es un supuesto; sin el CRM de la marca no hay dato público. Se cambia una celda (o se importan los datos con motor/importar.py) y todo se recalcula.\n"
                  "- **¿Por qué ese reparto entre perfiles?** También es supuesto, y está marcado como tal. El método no cambia con el dato real.")
        demo = f"Cambia la base de {es(S['base'])} a {es(S['base'] / 2)} en Supuestos:"
    return f"""# Cómo defender el Excel · {m['marca']}

## La idea en una frase
Todo sale de dos sitios: la hoja **Supuestos** (celdas amarillas) y la hoja **Calendario** (una fila por email). El resto son fórmulas.

## Ejemplo trabajado: {p['k']} · {p['nombre']}
- Recibe **{p['programados']} emails programados**. Su audiencia es el {p['pct']:.0%} de {es(S['base'])} = {es(p['audiencia'])}; sin el {S['control']:.0%} de control quedan **{es(p['recibe'])}**.
- Programados: {p['programados']} × {es(p['recibe'])} = **{es(p['imp_programados'])} impactos**.
- Triggers (solo a quien hace la acción, por eso van con tasa): {', '.join(f"{t['nombre']} {t['tasa']:.0%}" for t in trig)} → **{es(p['imp_triggers'])}**.
- Total del perfil: **{es(p['imp_total'])}**. Sumando todos los perfiles y la bienvenida ({es(c['bienvenida']['impactos'])}): **{es(c['impactos']['total'])} impactos**.

## Piezas: {tz['total']}
Pieza = creatividad distinta, no envío. {tz['unicos']} emails únicos + {tz['variantes']} variantes A/B + {tz['comunes']} {'común' if tz['comunes'] == 1 else 'comunes'} ('Todos') + {tz['plantillas']} plantillas de trigger.

## Preguntas trampa
{origen}
- **¿Por qué {es(min(pp), 1)}–{es(max(pp), 1)} emails por persona si hay {min(x['programados'] for x in c['perfiles'])}–{max(x['programados'] for x in c['perfiles'])} programados?** Se suman los triggers que le tocan de media.
- **¿El control no recibe nada?** No recibe programados; sí triggers y transaccionales, porque responden a lo que hace el cliente. Lo que se mide es el efecto del plan.
- **¿Y la saturación en los picos?** Tope de {S['tope_semanal']} programados por semana y perfil; en esas semanas el trigger con tope manda un email menos, y la fórmula lo aplica.

## Demostración en directo
{demo} los impactos cambian y las piezas siguen en {tz['total']}, porque producir no depende de cuánta gente lo recibe.
"""


def main(path, outdir=None):
    sys.stdout.reconfigure(encoding="utf-8")  # «·» y tildes legibles también con la salida redirigida en Windows
    m = json.loads(Path(path).read_text(encoding="utf-8"))
    err, av = validar(m)
    for a in av:
        print("AVISO:", a)
    if err:
        print("\n".join("ERROR: " + e for e in err))
        sys.exit(1)
    outdir = Path(outdir or Path(path).parent)
    outdir.mkdir(parents=True, exist_ok=True)
    xlsx = outdir / f"plan_email_{m['slug']}.xlsx"
    c = calcular(m)
    c["celdas"] = excel(m, xlsx)
    (outdir / "cifras.json").write_text(json.dumps(c, ensure_ascii=False, indent=1), encoding="utf-8")
    (outdir / f"defensa_{m['slug']}.md").write_text(defensa(m, c), encoding="utf-8")
    print(f"OK {xlsx} | {c['emails_programados']} emails · {es(c['impactos']['total'])} impactos · {c['piezas']['total']} piezas")
    return c


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(*sys.argv[1:3])
