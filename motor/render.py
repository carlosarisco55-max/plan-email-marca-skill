"""QA visual: .pptx -> PDF (LibreOffice) -> hojas PNG de 2x2 diapositivas para revisarlas a ojo.

Uso: python render.py <deck.pptx> [carpeta_salida]
Requiere LibreOffice y pymupdf (pip install pymupdf).
"""
import shutil
import subprocess
import sys
from pathlib import Path

import pymupdf as fitz

SOFFICE = shutil.which("soffice") or r"C:\Program Files\LibreOffice\program\soffice.exe"


def main(pptx, outdir=None):
    pptx = Path(pptx).resolve()
    outdir = Path(outdir or pptx.parent / "qa").resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(pptx)], check=True, capture_output=True)
    doc = fitz.open(outdir / (pptx.stem + ".pdf"))
    pages = [p.get_pixmap(dpi=75) for p in doc]
    w, h = pages[0].width, pages[0].height
    for i in range(0, len(pages), 4):
        sheet = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 2 * w, 2 * h), False)
        sheet.clear_with(255)
        for j, pm in enumerate(pages[i:i + 4]):
            pm.set_origin((j % 2) * w, (j // 2) * h)
            sheet.copy(pm, pm.irect)
        out = outdir / f"{pptx.stem}_{i // 4 + 1}.png"
        sheet.save(out)
        print("OK", out)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(*sys.argv[1:3])
