# -*- coding: utf-8 -*-
"""docx2png.py · 用 WPS/Word COM 把 docx 转 PDF，再把每页渲染成 PNG，便于检查排版。"""
import os
import sys

import pythoncom
import win32com.client as win32


def convert(docx_path, pdf_path):
    pythoncom.CoInitialize()
    last = None
    for prog in ("KWPS.Application", "Word.Application"):
        try:
            app = win32.DispatchEx(prog)
            break
        except Exception as e:  # noqa: BLE001
            last = e
            app = None
    if app is None:
        raise RuntimeError("无法启动 WPS/Word COM: %s" % last)

    app.Visible = True                      # WPS 自动化在隐藏窗口下会报“发生意外”
    try:
        app.DisplayAlerts = 0
    except Exception:  # noqa: BLE001
        pass
    doc = app.Documents.Open(os.path.abspath(docx_path))
    ok = False
    try:
        doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17, False, 0, 0, 1, 9999)
        ok = True
    except Exception as e:                # noqa: BLE001
        print("ExportAsFixedFormat 失败：", e)
    try:
        doc.Close(False)
    finally:
        app.Quit()
    pythoncom.CoUninitialize()
    if not ok or not os.path.exists(pdf_path):
        raise RuntimeError("PDF 转换失败")


def render(pdf_path, out_dir, zoom=2.0):
    import fitz
    os.makedirs(out_dir, exist_ok=True)
    d = fitz.open(pdf_path)
    paths = []
    for i, page in enumerate(d, start=1):
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        p = os.path.join(out_dir, "page_%02d.png" % i)
        pix.save(p)
        paths.append(p)
    print("共 %d 页" % len(paths))
    return paths


if __name__ == "__main__":
    docx_in, pdf_out, png_dir = sys.argv[1], sys.argv[2], sys.argv[3]
    convert(docx_in, pdf_out)
    print("PDF ->", pdf_out)
    for p in render(pdf_out, png_dir):
        print(p)
