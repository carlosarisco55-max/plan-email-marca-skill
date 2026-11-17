"""Todo de una vez y listo para enseñar: marca.json -> carpeta con el Excel, el deck (.pptx y .pdf), las vistas previas,
la chuleta de defensa en PDF y el guion en PDF (si hay guion.md).

Uso (desde PowerShell):  python entrega.py <marca.json> [carpeta_salida] [--guion guion.md] [--abrir]
Por defecto la carpeta es entrega_<SLUG>/ junto a marca.json, y el guion es guion.md junto a marca.json si existe.
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

MOTOR = Path(__file__).resolve().parent
sys.path.insert(0, str(MOTOR))
import generar  # noqa: E402
import md_pdf  # noqa: E402
import render  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("marca")
    ap.add_argument("carpeta", nargs="?")
    ap.add_argument("--guion")
    ap.add_argument("--abrir", action="store_true", help="abre la carpeta al terminar")
    a = ap.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    marca = Path(a.marca).resolve()
    slug = json.loads(marca.read_text(encoding="utf-8"))["slug"]
    out = Path(a.carpeta or marca.parent / f"entrega_{slug}").resolve()

    generar.main(marca, out)  # valida (sale si hay errores) y escribe Excel + cifras.json + defensa .md
    if not (MOTOR / "node_modules").exists():
        subprocess.run([shutil.which("npm") or "npm", "ci", "--ignore-scripts", "--silent"], cwd=MOTOR, check=True)
    pptx = out / f"{slug}-plan-email.pptx"
    subprocess.run(["node", str(MOTOR / "deck.js"), str(marca), str(out / "cifras.json"), str(pptx)], check=True)
    render.main(pptx, out / "vista_previa")
    shutil.move(out / "vista_previa" / f"{pptx.stem}.pdf", out / f"{pptx.stem}.pdf")
    md_pdf.main(out / f"defensa_{slug}.md", marca)
    guion = Path(a.guion) if a.guion else marca.parent / "guion.md"
    if guion.exists():
        md_pdf.main(guion, marca, out / f"guion_{slug}.pdf")

    print(f"\nEntrega lista en {out}")
    for f in sorted(out.rglob("*")):
        if f.is_file():
            print(f"  {f.relative_to(out)}  ({f.stat().st_size / 1024:,.0f} KB)")
    if a.abrir and hasattr(os, "startfile"):
        os.startfile(out)
    return out


if __name__ == "__main__":
    main()
