"""Datos reales de la marca -> marca.json. Sustituye los supuestos por datos y guarda los supuestos de antes
(la diferencia es lo que el dato corrige de la hipótesis). Lee Excel (.xlsx) o CSV; la primera fila son las cabeceras.

A) Una fila por cliente con su perfil ya asignado (o filas agregadas, con --n):
   python importar.py marca.json clientes.xlsx --hoja Clientes --perfil Segmento [--n Clientes] [--suscrito Optin] [--base 120000]
B) Una fila por línea de compra: el perfil sale de la familia con más del --umbral del gasto (variable que discrimina):
   python importar.py marca.json ventas.csv --cliente ID --familia Familia --importe Importe [--fecha Fecha --meses 24]
   Necesita en marca.json  "familias": {"<familia>": "<código de perfil>"}.
En las dos: --suscrito filtra a quien acepta email (1/sí/true/x); --base fija el total de suscriptores (los que no
aparecen en la tabla cuentan como sin perfil); --triggers HOJA (o archivo) con columnas nombre y tasa;
--saltar N salta N filas de título por encima de la cabecera real (típico de un export tal cual, sin limpiar).
El perfil puede venir como código (A) o como nombre del perfil. Valor vacío o desconocido = sin perfil.
Filas con dato roto (cliente sin id, importe o fecha no reconocibles) se descartan con un AVISO, no rompen el resto.
Celdas combinadas (una familia repartida visualmente sobre varias filas) no se rellenan solas: hay que
deshacerlas en el Excel antes de importar, o esas filas quedan sin familia y cuentan como sin perfil.
"""
import argparse
import csv
import json
import math
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

SI = {"1", "si", "sí", "true", "verdadero", "x", "yes", "y", "s"}


def leer(path, hoja=None, saltar=0):
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(p, read_only=True, data_only=True)
        if hoja and hoja not in wb.sheetnames:
            sys.exit(f"La hoja '{hoja}' no existe en {p.name}. Hojas: {wb.sheetnames}")
        it = (wb[hoja] if hoja else wb.worksheets[0]).iter_rows(values_only=True)
        for _ in range(saltar):  # filas de título por encima de la cabecera real, típicas de un export de verdad
            next(it)
        cab = [str(c).strip() if c is not None else "" for c in next(it)]
        return [dict(zip(cab, r)) for r in it if any(v not in (None, "") for v in r)]
    with open(p, encoding="utf-8-sig", newline="") as f:
        lineas = f.readlines()[saltar:]
        dialecto = csv.Sniffer().sniff("".join(lineas[:20]), delimiters=",;\t")
        return list(csv.DictReader(lineas, dialect=dialecto))


def columnas(filas, *cols):
    faltan = [c for c in cols if c and filas and c not in filas[0]]
    if faltan:
        sys.exit(f"No encuentro las columnas {faltan}. Columnas del archivo: {list(filas[0])}")


def num(v):
    """Lanza ValueError si v no es un número (texto suelto, celda vacía de verdad no debería llegar aquí)."""
    if isinstance(v, (int, float)):
        if not math.isfinite(v):
            raise ValueError(v)
        return float(v)
    s = str(v or "0").strip().replace("€", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    x = float(s or 0)  # float(str) lanza ValueError solo si s no está vacío y no es numérico
    if not math.isfinite(x):
        raise ValueError(v)
    return x


def fecha(v):
    """Lanza ValueError si v no es una fecha reconocible."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(v).strip()[:10], f).date()
        except ValueError:
            pass
    raise ValueError(v)


def si(v):
    if isinstance(v, bool):  # checkbox de verdad (openpyxl lo da como bool, no como texto)
        return v
    if isinstance(v, (int, float)):  # 1 o 1.0 numérico, no el texto "1"
        return v == 1
    return str(v if v is not None else "").strip().lower() in SI


def contar_clientes(filas, a, m):
    columnas(filas, a.perfil, a.n, a.suscrito)
    buscar = {}
    for p in m["perfiles"]:
        for clave in (p["k"], p["nombre"], p["corto"]):
            buscar[clave.strip().lower()] = p["k"]
    c, raros, malas = defaultdict(float), defaultdict(int), 0
    for f in filas:
        if a.suscrito and not si(f.get(a.suscrito)):
            continue
        v = str(f.get(a.perfil) or "").strip()
        k = buscar.get(v.lower(), "S")
        if v and k == "S":
            raros[v] += 1
        if a.n:
            try:
                c[k] += num(f[a.n])
            except ValueError:  # fila con un valor no numérico en --n (texto suelto, "N/D"...): se descarta, no rompe el resto
                malas += 1
                continue
        else:
            c[k] += 1
    if malas:
        print(f"AVISO: {malas} fila(s) con un valor no numérico en «{a.n}», descartadas")
    regla = f"columna «{a.perfil}» de {Path(a.datos).name}" + (f", ponderada por «{a.n}»" if a.n else "")
    return c, raros, regla


def contar_transacciones(filas, a, m):
    columnas(filas, a.cliente, a.familia, a.importe, a.fecha, a.suscrito)
    fam_perfil = {k.strip().lower(): v for k, v in (m.get("familias") or {}).items()}
    if not fam_perfil:
        sys.exit('Para asignar perfil por familia hace falta "familias": {"<familia>": "<código>"} en marca.json')
    sin_cliente = sin_fecha = importe_malo = 0
    fechas_ok = []
    for f in filas:
        if not a.fecha:
            continue
        try:
            fechas_ok.append(fecha(f[a.fecha]))
        except ValueError:
            pass
    corte = max(fechas_ok) - timedelta(days=round(a.meses * 30.44)) if fechas_ok else None
    gasto, acepta = defaultdict(lambda: defaultdict(float)), defaultdict(bool)
    for f in filas:
        if not str(f.get(a.cliente) or "").strip():  # sin id de cliente no se puede agregar su gasto: no es "un cliente más"
            sin_cliente += 1
            continue
        if a.fecha:
            try:
                d = fecha(f[a.fecha])
            except ValueError:
                sin_fecha += 1
                continue
            if corte and d <= corte:
                continue
        try:
            imp = num(f[a.importe])
        except ValueError:
            importe_malo += 1
            continue
        cli = str(f[a.cliente]).strip()
        gasto[cli][str(f[a.familia] or "").strip().lower()] += imp
        if a.suscrito:
            acepta[cli] |= si(f.get(a.suscrito))
    for n, etiqueta in [(sin_cliente, "sin identificador de cliente"), (sin_fecha, f"con fecha no reconocible en «{a.fecha}»"),
                        (importe_malo, f"con un importe no numérico en «{a.importe}»")]:
        if n:
            print(f"AVISO: {n} fila(s) {etiqueta}, descartadas")
    c, raros = defaultdict(float), defaultdict(int)
    for cli, g in gasto.items():
        if a.suscrito and not acepta[cli]:
            continue
        total = sum(g.values())
        fam, v = max(g.items(), key=lambda x: x[1])
        if fam not in fam_perfil:
            raros[fam] += 1
        dominante = total > 0 and v / total > a.umbral
        c[fam_perfil.get(fam, "S") if dominante else "S"] += 1
    regla = (f"familia con más del {a.umbral:.0%} del gasto" + (f" en los últimos {a.meses} meses" if a.fecha else "")
             + f" ({Path(a.datos).name})")
    return c, raros, regla


def triggers(a, m):
    origen = a.triggers if Path(a.triggers).suffix.lower() in (".csv", ".xlsx") else a.datos
    filas = leer(origen, None if origen == a.triggers else a.triggers)
    cab = {k.lower(): k for k in filas[0]}
    if "nombre" not in cab or "tasa" not in cab:
        sys.exit(f"La tabla de triggers necesita columnas 'nombre' y 'tasa'. Tiene: {list(filas[0])}")
    try:
        reales = {str(f[cab["nombre"]]).strip().lower(): num(f[cab["tasa"]]) for f in filas}
    except ValueError as e:
        sys.exit(f"Tasa no numérica en la tabla de triggers: {e}")
    antes = {}
    for t in m["triggers"]:
        v = reales.pop(t["nombre"].strip().lower(), None)
        if v is not None:
            v = v / 100 if v > 1 else v  # admite 8 o 0,08
            antes[t["nombre"]] = t["tasa"]
            t["tasa"], t["real"] = v, True
    for n in reales:
        print(f"AVISO: el trigger «{n}» de los datos no existe en marca.json (los nombres tienen que coincidir)")
    return antes


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("marca")
    ap.add_argument("datos")
    ap.add_argument("--hoja")
    ap.add_argument("--saltar", type=int, default=0, help="filas de título por encima de la cabecera real")
    ap.add_argument("--perfil")
    ap.add_argument("--n")
    ap.add_argument("--cliente")
    ap.add_argument("--familia")
    ap.add_argument("--importe")
    ap.add_argument("--fecha")
    ap.add_argument("--meses", type=int, default=24)
    ap.add_argument("--umbral", type=float, default=0.5)
    ap.add_argument("--suscrito")
    ap.add_argument("--base", type=int)
    ap.add_argument("--triggers")
    a = ap.parse_args(argv)
    if bool(a.perfil) == bool(a.familia):
        sys.exit("Indica --perfil (forma A) o --familia con --cliente y --importe (forma B), no las dos.")
    if hasattr(sys.stdout, "reconfigure"):  # no StringIO (tests capturan la salida para comprobar los avisos)
        sys.stdout.reconfigure(encoding="utf-8")
    ruta = Path(a.marca)
    m = json.loads(ruta.read_text(encoding="utf-8"))
    filas = leer(a.datos, a.hoja, a.saltar)
    c, raros, regla = (contar_clientes if a.perfil else contar_transacciones)(filas, a, m)
    total = sum(c.values())
    if a.base:
        if a.base < total:
            sys.exit(f"--base {a.base} es menor que los {total:,.0f} clientes contados")
        c["S"] += a.base - total
        regla += f"; base de {a.base:,} suscriptores (los que no compran cuentan como sin perfil)".replace(",", ".")
        total = a.base
    if not total:
        sys.exit("No he contado a nadie: revisa el filtro --suscrito y las columnas.")
    codes = [p["k"] for p in m["perfiles"]]
    previo = m.get("datos_reales", {}).get("antes", {})  # si se reimporta, «antes» sigue siendo el supuesto original
    antes = {"base": previo.get("base", m["supuestos"]["base"]),
             "pct": previo.get("pct", {**{p["k"]: p["pct"] for p in m["perfiles"]}, "S": m["sin_perfil"]["pct"]}),
             "triggers": previo.get("triggers", {})}
    m["supuestos"]["base"] = round(total)
    for p in m["perfiles"]:
        p["pct"] = c[p["k"]] / total
    m["sin_perfil"]["pct"] = c["S"] / total
    if a.triggers:
        antes["triggers"] = {**triggers(a, m), **antes["triggers"]}
    m["datos_reales"] = {"fuente": Path(a.datos).name, "fecha": date.today().isoformat(), "regla": regla,
                         "conteo": {k: c[k] for k in codes + ["S"]}, "antes": antes}
    ruta.write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Datos de {Path(a.datos).name} → {ruta.name}  ({regla})")
    print(f"  base: {antes['base']:,} supuesta → {round(total):,} real".replace(",", "."))
    for k in codes + ["S"]:
        print(f"  {k}: {antes['pct'][k]:.0%} supuesto → {c[k] / total:.1%} real ({c[k]:,.0f} clientes)".replace(",", "."))
    for t, v in antes["triggers"].items():
        print(f"  trigger «{t}»: {v:.0%} supuesto → {next(x['tasa'] for x in m['triggers'] if x['nombre'] == t):.1%} real")
    if raros:
        print("AVISO: valores sin perfil asignado (cuentan como sin perfil):", dict(raros))
    return m


if __name__ == "__main__":
    main()
