"""Comprobación del motor. Uso: python motor/test_motor.py

1. The North Face reproduce las cifras entregadas (50 emails, 1.661.240 impactos, 67 piezas).
2. Para cada ejemplos/*.json, el Excel recalculado por LibreOffice da lo mismo que calcular() en Python.
3. La validación para lo que la auditoría de TNF tuvo que corregir a mano.
4. importar.py cuenta bien las dos formas de tabla y, con datos reales, el Excel sigue cuadrando con Python.
"""
import csv
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generar as g  # noqa: E402
import importar  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
SOFFICE = Path(shutil.which("soffice") or r"C:\Program Files\LibreOffice\program\soffice.exe")


def cargar(nombre):
    return json.loads((RAIZ / "ejemplos" / nombre).read_text(encoding="utf-8"))


def esperado(c, clave):
    por = {p["k"]: p for p in c["perfiles"]}
    if clave in ("impactos_total", "impactos_programados"):
        return c["impactos"][clave.split("_")[1]]
    if clave in ("piezas_total", "piezas_check"):
        return c["piezas"]["total"]
    if clave == "suma_pct":
        return 1
    campo, k = clave.rsplit("_", 1)
    return por[k][{"programados": "programados", "imp_total": "imp_total", "emails_pp": "emails_por_persona"}[campo]]


def test_tnf_reproduce_lo_entregado():
    c = g.calcular(cargar("north-face.json"))
    assert c["emails_programados"] == 50
    assert abs(c["impactos"]["total"] - 1_661_240) < 0.5
    assert c["piezas"] == {"unicos": 49, "variantes": 4, "comunes": 1, "plantillas": 13, "total": 67}
    assert [round(p["imp_total"]) for p in c["perfiles"]] == [179_160, 601_920, 533_760, 334_400]
    assert [round(p["emails_por_persona"], 1) for p in c["perfiles"]] == [16.5, 18.5, 18.5, 18.5]


def comparar(m, nombre):
    """Genera el Excel, lo recalcula con LibreOffice y compara cada celda clave con calcular()."""
    if not SOFFICE.exists():
        print(f"  {nombre}: sin LibreOffice, se omite la comparación con el Excel recalculado")
        return
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        xlsx = tmp / "plan.xlsx"
        celdas = g.excel(m, xlsx)
        c = g.calcular(m)
        subprocess.run([str(SOFFICE), "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp / "r"), str(xlsx)],
                       check=True, capture_output=True)
        wb = load_workbook(tmp / "r" / "plan.xlsx", data_only=True)
        for clave, ref in celdas.items():
            hoja, celda = ref.split("!")
            v = wb[hoja][celda].value
            assert isinstance(v, (int, float)) and abs(v - esperado(c, clave)) < 0.01, f"{nombre} {clave} {ref}: Excel {v} vs Python {esperado(c, clave)}"
        print(f"  {nombre}: {len(celdas)} celdas iguales en Excel y Python")


def test_excel_igual_que_python():
    for f in sorted((RAIZ / "ejemplos").glob("*.json")):
        comparar(json.loads(f.read_text(encoding="utf-8")), f.name)


def errores(cambio):
    m = cargar("north-face.json")
    cambio(m)
    return " | ".join(g.validar(m)[0])


def test_validacion():
    assert g.validar(cargar("north-face.json"))[0] == []
    assert g.validar(cargar("lumen-cafe.json")) == ([], ["perfil G, semana 2: sin email programado"])
    # el descuento colado en un email que no es de tipo Oferta (en TNF: el take-back del 10 % al alpinista) se avisa
    assert any("suena a oferta" in a and "Take-back" in a for a in g.validar(cargar("north-face.json"))[1])
    # el titular del perfil tiene que enviarse alguna vez
    m = cargar("lumen-cafe.json")
    m["perfiles"][0]["tarjeta"]["titular"] = "Un titular que no se envía"
    assert any("no está en ningún asunto" in a for a in g.validar(m)[1])
    # el perfil «sin descuentos» no puede recibir descuentos
    assert "contradice su barrera" in errores(lambda m: m["calendario"].append([2, "A", "Oferta", "Vie", "x", "y", ""]))
    # tope semanal
    assert "tope 2" in errores(lambda m: m["calendario"].append([4, "B", "Producto", "Jue", "x", "y", ""]))
    # el reparto tiene que sumar 100 %
    assert "no 100 %" in errores(lambda m: m["sin_perfil"].update(pct=0.30))
    # los triggers no se programan
    assert "no programable" in errores(lambda m: m["calendario"].append([2, "A", "Trigger", "Vie", "x", "y", ""]))
    # trigger a un perfil que no existe
    assert "inexistentes" in errores(lambda m: m["triggers"][1].update(perfiles="A,B,Z"))
    # los textos de meses del deck tienen que cuadrar con el periodo
    assert "deck.meses" in errores(lambda m: m["deck"]["meses"].pop())


def escribir_csv(ruta, filas):
    with open(ruta, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerows(filas)


def test_importar_clientes():
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        shutil.copy(RAIZ / "ejemplos" / "lumen-cafe.json", tmp / "marca.json")
        # perfil por código o por nombre; vacío o desconocido = sin perfil; los que no aceptan email no cuentan
        escribir_csv(tmp / "clientes.csv", [["ID", "Segmento", "Optin"], [1, "R", "sí"], [2, "Recurrente", "1"], [3, "R", "x"],
                                            [4, "G", "si"], [5, "", "si"], [6, "VIP", "si"], [7, "R", "no"]])
        m = importar.main([str(tmp / "marca.json"), str(tmp / "clientes.csv"), "--perfil", "Segmento", "--suscrito", "Optin"])
        dr = m["datos_reales"]
        assert dr["conteo"] == {"R": 3, "G": 1, "S": 2} and m["supuestos"]["base"] == 6
        assert dr["antes"]["pct"]["R"] == 0.55 and abs(m["perfiles"][0]["pct"] - 0.5) < 1e-12
        assert g.validar(m)[0] == []
        comparar(m, "lumen-cafe con datos reales")
        # reimportar no pierde el supuesto original
        m = importar.main([str(tmp / "marca.json"), str(tmp / "clientes.csv"), "--perfil", "Segmento", "--base", "10"])
        assert m["datos_reales"]["antes"]["pct"]["R"] == 0.55 and m["datos_reales"]["conteo"]["S"] == 10 - 5


def test_importar_datos_sucios():
    """Un Excel real: fila de título, mayúsculas/espacios en la familia, cliente sin id, importe y fecha
    rotos, y opt-in como bool/número en vez de texto. Ninguna fila mala debe tumbar el import entero."""
    import io
    from contextlib import redirect_stdout
    from datetime import date
    from openpyxl import Workbook

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        m = cargar("north-face.json")
        m["familias"] = {"Mountain": "A", "Outdoor": "B", "Urban": "C"}
        (tmp / "marca.json").write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")
        wb = Workbook()
        ws = wb.active
        ws.append(["Export interno de ventas — no editar a mano"])
        ws.append(["Cliente", "Familia", "Importe", "Fecha", "Optin"])
        for row in [
            ["c1", " mountain ", 300.0, date(2026, 10, 1), True],   # A: familia con espacios y minúsculas
            ["c1", "URBAN", 50, date(2026, 10, 2), True],           # mismo cliente, sigue siendo A (300/350 > 50 %)
            [None, "Mountain", 999, date(2026, 10, 1), True],       # sin cliente: no se puede agregar, se descarta
            ["c2", "Outdoor", "no es un número", date(2026, 9, 1), 1.0],  # importe roto: se descarta la fila
            ["c3", "Outdoor", 500, "fecha-rara", 1],                # fecha rota: se descarta la fila
            ["c4", "Outdoor", 1200.5, date(2026, 7, 1), 1],         # B: opt-in como entero 1, no como texto
            ["c5", "Calcetines", 20, date(2026, 6, 1), True],       # familia sin perfil -> sin perfil
            ["c6", "Mountain", 80, date(2024, 1, 1), True],         # fuera de los 24 meses
            ["c7", "Outdoor", 500, date(2026, 8, 1), False],        # no acepta email (bool False): se excluye
        ]:
            ws.append(row)
        xlsx = tmp / "ventas.xlsx"
        wb.save(xlsx)
        salida = io.StringIO()
        with redirect_stdout(salida):
            m = importar.main([str(tmp / "marca.json"), str(xlsx), "--saltar", "1", "--cliente", "Cliente",
                               "--familia", "Familia", "--importe", "Importe", "--fecha", "Fecha", "--meses", "24",
                               "--suscrito", "Optin"])
        log = salida.getvalue()
        assert m["datos_reales"]["conteo"] == {"A": 1, "B": 1, "C": 0, "S": 1}, m["datos_reales"]["conteo"]
        assert "1 fila(s) sin identificador de cliente" in log
        assert "1 fila(s) con un importe no numérico" in log
        assert "1 fila(s) con fecha no reconocible" in log
        assert g.validar(m)[0] == []
        comparar(m, "north-face con un Excel real y sucio")


def test_importar_transacciones():
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        m = cargar("north-face.json")
        m["familias"] = {"Mountain": "A", "Outdoor": "B", "Urban": "C"}
        (tmp / "marca.json").write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")
        escribir_csv(tmp / "ventas.csv", [["Cliente", "Familia", "Importe", "Fecha"],
                                          ["c1", "Mountain", "300,00", "2026-10-01"],                              # A
                                          ["c2", "Mountain", "60", "2026-09-01"], ["c2", "Urban", "40", "2026-09-02"],  # A: 60 % > 50 %
                                          ["c3", "Mountain", "50", "2026-08-01"], ["c3", "Urban", "50", "2026-08-02"],  # sin familia dominante
                                          ["c4", "Outdoor", "1.200,50", "2026-07-01"],                            # B (importe con miles)
                                          ["c5", "Urban", "80", "2024-01-01"],                                    # fuera de los 24 meses
                                          ["c6", "Calcetines", "20", "2026-06-01"]])                              # familia sin perfil
        m = importar.main([str(tmp / "marca.json"), str(tmp / "ventas.csv"), "--cliente", "Cliente", "--familia", "Familia",
                           "--importe", "Importe", "--fecha", "Fecha", "--meses", "24"])
        assert m["datos_reales"]["conteo"] == {"A": 2, "B": 1, "C": 0, "S": 2}
        assert g.validar(m)[0] == []
        comparar(m, "north-face con transacciones")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
