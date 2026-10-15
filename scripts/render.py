# -*- coding: utf-8 -*-
"""逐页导出 PNG 供目视核对。用法: python render.py <pptx> <outdir> [页码...]
Windows 优先用 PowerPoint COM（与最终放映一致）；否则用 LibreOffice 转 PDF 再用 PyMuPDF 栅格化。"""
import os
import subprocess
import sys
import tempfile


def by_com(pptx, outdir, pages):
    import win32com.client as win32
    app = win32.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(pptx, WithWindow=False, ReadOnly=True)
    try:
        for n in pages or range(1, pres.Slides.Count + 1):
            p = os.path.join(outdir, "s%02d.png" % n)
            pres.Slides(n).Export(p, "PNG", 1672, 941)
            print(p)
    finally:
        pres.Close()
        app.Quit()


def by_soffice(pptx, outdir, pages):
    import fitz
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, pptx], check=True)
    doc = fitz.open(os.path.join(tmp, os.path.splitext(os.path.basename(pptx))[0] + ".pdf"))
    for n in pages or range(1, doc.page_count + 1):
        p = os.path.join(outdir, "s%02d.png" % n)
        doc[n - 1].get_pixmap(dpi=150).save(p)
        print(p)


def main():
    pptx, outdir = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
    pages = [int(x) for x in sys.argv[3:]]
    os.makedirs(outdir, exist_ok=True)
    try:
        by_com(pptx, outdir, pages)
    except ImportError:
        by_soffice(pptx, outdir, pages)


if __name__ == "__main__":
    main()
