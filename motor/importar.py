"""Datos reales de la marca -> marca.json. Sustituye los supuestos por datos y guarda los supuestos de antes
(la diferencia es lo que el dato corrige de la hipótesis). Lee Excel (.xlsx) o CSV; la primera fila son las cabeceras.

A) Una fila por cliente con su perfil ya asignado (o filas agregadas, con --n):
   python importar.py marca.json clientes.xlsx --hoja Clientes --perfil Segmento [--n Clientes] [--suscrito Optin] [--base 120000]
B) Una fila por línea de compra: el perfil sale de la familia con más del --umbral del gasto (variable que discrimina):
   python importar.py marca.json ventas.csv --cliente ID --familia Familia --importe Importe [--fecha Fecha --meses 24]
   Necesita en marca.json  "familias": {"<familia>": "<código de perfil>"}.
En las dos: --suscrito filtra a quien acepta email (1/sí/true/x); --base fija el total de suscriptores (los que no
aparecen en la tabla cuentan como sin perfil); --triggers HOJA (o archivo) con columnas nombre y tasa.
El perfil puede venir como código (A) o como nombre del perfil. Valor vacío o desconocido = sin perfil.
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


def leer(path, hoja=None):
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(p, read_only=True, data_only=True)
        if hoja and hoja not in wb.sheetnames:
            sys.exit(f"La hoja '{hoja}' no existe en {p.name}. Hojas: {wb.sheetnames}")
        it = (wb[hoja] if hoja else wb.worksheets[0]).iter_rows(values_only=True)
        cab = [str(c).strip() if c is not None else "" for c in next(it)]
        return [dict(zip(cab, r)) for r in it if any(v not in (None, "") for v in r)]
    with open(p, encoding="utf-8-sig", newline="") as f:
        dialecto = csv.Sniffer().sniff(f.read(4096), delimiters=",;\t")
        f.seek(0)
        return list(csv.DictReader(f, dialect=dialecto))


def columnas(filas, *cols):
    faltan = [c for c in cols if c and filas and c not in filas[0]]
    if faltan:
        sys.exit(f"No encuentro las columnas {faltan}. Columnas del archivo: {list(filas[0])}")


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "0").strip().replace("€", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    x = float(s or 0)
    if not math.isfinite(x):
        sys.exit(f"Importe no válido: {v!r}")
    return x


def fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(v).strip()[:10], f).date()
        except ValueError:
            pass
    sys.exit(f"Fecha que no entiendo: {v!r} (usa AAAA-MM-DD o DD/MM/AAAA)")


def si(v):
    return str(v if v is not None else "").strip().lower() in SI


def contar_clientes(filas, a, m):
    columnas(filas, a.perfil, a.n, a.suscrito)
    buscar = {}
    for p in m["perfiles"]:
        for clave in (p["k"], p["nombre"], p["corto"]):
            buscar[clave.strip().lower()] = p["k"]
    c, raros = defaultdict(float), defaultdict(int)
    for f in filas:
        if a.suscrito and not si(f.get(a.suscrito)):
            continue
        v = str(f.get(a.perfil) or "").strip()
        k = buscar.get(v.lower(), "S")
        if v and k == "S":
            raros[v] += 1
        c[k] += num(f[a.n]) if a.n else 1
    regla = f"columna «{a.perfil}» de {Path(a.datos).name}" + (f", ponderada por «{a.n}»" if a.n else "")
    return c, raros, regla


def contar_transacciones(filas, a, m):
    columnas(filas, a.cliente, a.familia, a.importe, a.fecha, a.suscrito)
    fam_perfil = m.get("familias")
    if not fam_perfil:
        sys.exit('Para asignar perfil por familia hace falta "familias": {"<familia>": "<código>"} en marca.json')
    corte = None
    if a.fecha:
        corte = max(fecha(f[a.fecha]) for f in filas) - timedelta(days=round(a.meses * 30.44))
    gasto, acepta = defaultdict(lambda: defaultdict(float)), defaultdict(bool)
    for f in filas:
        if corte and fecha(f[a.fecha]) <= corte:
            continue
        cli = str(f[a.cliente]).strip()
        gasto[cli][str(f[a.familia]).strip()] += num(f[a.importe])
        if a.suscrito:
            acepta[cli] |= si(f.get(a.suscrito))
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
    reales = {str(f[cab["nombre"]]).strip().lower(): num(f[cab["tasa"]]) for f in filas}
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
    sys.stdout.reconfigure(encoding="utf-8")
    ruta = Path(a.marca)
    m = json.loads(ruta.read_text(encoding="utf-8"))
    filas = leer(a.datos, a.hoja)
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
