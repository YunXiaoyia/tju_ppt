#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_render.py — render a deck to PNGs for visual inspection.

    python tju_render.py deck.pptx out_dir [--pages 3,4,9] [--dpi 144]

Linux/macOS: LibreOffice headless → PDF → pdftoppm.  Windows: PowerPoint COM (exact fidelity) → PDF → PyMuPDF.
Prints the PNG paths.  LibreOffice wraps CJK slightly wider than PowerPoint, so a line that just fits in
LibreOffice will fit in PowerPoint; the reverse is not guaranteed.
"""

import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile


def _pdf_windows(pptx, pdf):
    import gc
    import win32com.client  # type: ignore
    app = win32com.client.DispatchEx("PowerPoint.Application")
    deck = app.Presentations.Open(os.path.abspath(pptx), WithWindow=False)
    try:
        deck.SaveAs(os.path.abspath(pdf), 32)
    finally:
        deck.Close()
        app.Quit()
        gc.collect()


def _pdf_libreoffice(pptx, outdir):
    exe = shutil.which("libreoffice") or shutil.which("soffice")
    if not exe:
        sys.exit("libreoffice/soffice not found")
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx],
                   check=True, capture_output=True, timeout=300)
    return os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")


def render(pptx, out_dir, pages=None, dpi=144):
    os.makedirs(out_dir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="tju_render_")
    staged = os.path.join(tmp, "deck.pptx")
    shutil.copyfile(pptx, staged)               # never lock the original
    if sys.platform.startswith("win"):
        pdf = os.path.join(tmp, "deck.pdf")
        _pdf_windows(staged, pdf)
        import fitz  # type: ignore  (PyMuPDF)
        doc = fitz.open(pdf)
        out = []
        for i, page in enumerate(doc, start=1):
            if pages and i not in pages:
                continue
            p = os.path.join(out_dir, f"slide-{i:02d}.png")
            page.get_pixmap(dpi=dpi).save(p)
            out.append(p)
        return out
    pdf = _pdf_libreoffice(staged, tmp)
    for old in glob.glob(os.path.join(out_dir, "slide-*.png")):
        os.remove(old)
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf, os.path.join(out_dir, "slide")], check=True)
    out = []
    for p in sorted(glob.glob(os.path.join(out_dir, "slide-*.png"))):
        n = int(os.path.splitext(p)[0].rsplit("-", 1)[1])
        dst = os.path.join(out_dir, f"slide-{n:02d}.png")
        if p != dst:
            os.replace(p, dst)
        if pages and n not in pages:
            os.remove(dst)
            continue
        out.append(dst)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("out_dir")
    ap.add_argument("--pages", default="")
    ap.add_argument("--dpi", type=int, default=144)
    a = ap.parse_args()
    pages = {int(x) for x in a.pages.split(",") if x.strip()} or None
    for p in render(a.pptx, a.out_dir, pages, a.dpi):
        print(p)


if __name__ == "__main__":
    main()
