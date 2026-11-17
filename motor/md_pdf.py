"""Markdown -> PDF con los colores de la marca (guion de los ponentes, chuleta de defensa).

Uso: python md_pdf.py <archivo.md> <marca.json> [salida.pdf]
Admite: # título, ## sección, ### sección, párrafos, listas con "- ", **negrita**, "---" y la cabecera de diapositiva
del guion:  ### 3 · ANDRÉS · Título de la diapositiva (1 min 15 s)
Imprime con Microsoft Edge (mejor resultado); si no puede, con LibreOffice.
"""
import html
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

EDGE = [Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"), Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe")]
SOFFICE = shutil.which("soffice") or r"C:\Program Files\LibreOffice\program\soffice.exe"


def inline(t):
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", html.escape(t, quote=False))


def a_html(md, m):
    e = m["deck"]["estilo"]
    colores = [e["acento"]] + list(e["tipos"].values())[1:]  # un color por ponente, en orden de aparición
    ponentes, body, lista = {}, [], False
    for line in md.splitlines():
        if line.startswith("- "):
            if not lista:
                body.append("<ul>")
                lista = True
            body.append(f"<li>{inline(line[2:])}</li>")
            continue
        if lista:
            body.append("</ul>")
            lista = False
        if not line.strip():
            continue
        h = re.match(r"### (\d+) · (.+?) · (.+?) \((.+)\)$", line)
        if h:
            n, quien, tit, dur = h.groups()
            bg = colores[ponentes.setdefault(quien, len(ponentes)) % len(colores)]
            r, g, b = (int(bg[j:j + 2], 16) for j in (0, 2, 4))
            fg = "141414" if 0.299 * r + 0.587 * g + 0.114 * b > 150 else "FFFFFF"  # texto legible sobre el color del ponente
            body.append(f'<div class="slide"><span class="num">{n}</span><span class="who" style="background:#{bg};color:#{fg}">{html.escape(quien)}</span>'
                        f"<h3>{inline(tit)}</h3><span class='dur'>{html.escape(dur)}</span></div>")
        elif line.startswith("### "):
            body.append(f"<h3 class='sec'>{inline(line[4:])}</h3>")
        elif line.startswith("## "):
            body.append(f"<h2>{inline(line[3:])}</h2>")
        elif line.startswith("# "):
            body.append(f"<h1>{inline(line[2:])}</h1>")
        elif line.startswith("---"):
            body.append("<hr>")
        else:
            body.append(f"<p>{inline(line)}</p>")
    if lista:
        body.append("</ul>")
    H, B, A, K = e["fuente_titulo"], e["fuente_texto"], e["acento"], e["oscuro"]
    css = f"""@page{{size:A4;margin:18mm 18mm 16mm}}body{{font-family:'{B}','Segoe UI',sans-serif;color:#1A1A1A;font-size:11.5pt;line-height:1.5}}
h1{{font-family:'{H}','{B}',sans-serif;font-size:24pt;text-transform:uppercase;border-bottom:4px solid #{A};padding-bottom:6px;margin:0 0 10px}}
h2,.sec{{font-family:'{H}','{B}',sans-serif;font-size:15pt;text-transform:uppercase;color:#{K};border-left:6px solid #{A};padding-left:8px;margin:18px 0 8px;break-after:avoid}}
p{{margin:0 0 8px}}ul{{margin:0 0 8px 18px;padding:0}}li{{margin:0 0 4px}}hr{{border:0;border-top:1px solid #CFCAC0;margin:14px 0}}
.slide{{display:flex;align-items:baseline;gap:10px;margin:18px 0 8px;padding:8px 12px;background:#{K};color:#EDEBE6;border-left:6px solid #{A};break-after:avoid}}
.slide h3{{margin:0;font-family:'{H}','{B}',sans-serif;font-size:15pt;text-transform:uppercase;flex:1;font-weight:600}}
.num{{font-family:'{H}';font-size:18pt;color:#{A}}}.who{{font-size:9pt;letter-spacing:2px;padding:2px 8px;border-radius:10px;font-weight:700}}
.dur{{font-size:9.5pt;color:#A8A49C}}"""
    return (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><title>{html.escape(m["marca"])}</title>'
            f"<style>{css}</style></head><body>{''.join(body)}</body></html>")


def main(md_path, marca_path, out=None):
    md_path = Path(md_path).resolve()
    out = Path(out or md_path.with_suffix(".pdf")).resolve()
    m = json.loads(Path(marca_path).read_text(encoding="utf-8"))
    out.unlink(missing_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        page = Path(tmp) / f"{out.stem}.html"
        page.write_text(a_html(md_path.read_text(encoding="utf-8"), m), encoding="utf-8")
        edge = next((p for p in EDGE if p.exists()), None)
        if edge:
            subprocess.run([str(edge), "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--user-data-dir={tmp}\\edge",
                            f"--print-to-pdf={out}", page.as_uri()], capture_output=True, timeout=120)
        if not out.exists() and Path(SOFFICE).exists():  # p. ej. lanzado desde Git Bash: Edge sale con 0 sin escribir nada
            subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf:writer_web_pdf_Export", "--outdir", tmp, str(page)],
                           capture_output=True, timeout=180)
            if (Path(tmp) / f"{out.stem}.pdf").exists():
                shutil.move(Path(tmp) / f"{out.stem}.pdf", out)
        if not out.exists():
            keep = out.with_suffix(".html")
            shutil.copy(page, keep)
            sys.exit(f"No he podido imprimir el PDF: abre {keep} en el navegador e imprímelo.")
    print("OK", out)
    return out


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(*sys.argv[1:4])
