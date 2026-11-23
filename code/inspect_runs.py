# -*- coding: utf-8 -*-
"""inspect_runs.py · 查看关键段落的 run 级格式与表格属性。"""
import sys
import docx
from docx.shared import Emu

d = docx.Document(sys.argv[1])
want = set(range(0, 30)) | {56, 57, 58, 64, 65, 66, 67}
for i, p in enumerate(d.paragraphs):
    if i not in want:
        continue
    pf = p.paragraph_format
    print("P%03d style=%s align=%s ind=%s/%s spacing=%s/%s" %
          (i, p.style.name, p.alignment,
           pf.first_line_indent, pf.left_indent, pf.line_spacing, pf.space_after))
    for j, r in enumerate(p.runs):
        f = r.font
        print("   run%d %r font=%s size=%s bold=%s underline=%s eastasia=%s" % (
            j, r.text[:60], f.name, f.size.pt if f.size else None, f.bold, f.underline,
            r._element.rPr.rFonts.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia")
            if r._element.rPr is not None and r._element.rPr.rFonts is not None else None))

sec = d.sections[0]
print("margins L%.2f R%.2f T%.2f B%.2f cm" % (
    Emu(sec.left_margin).cm, Emu(sec.right_margin).cm,
    Emu(sec.top_margin).cm, Emu(sec.bottom_margin).cm))
print("page width cm", Emu(sec.page_width).cm)

for ti, t in enumerate(d.tables):
    print("TBL%d style=%s align=%s width=%s" % (ti, t.style.name if t.style else None,
                                                t.alignment, t._tbl.tblPr.xml[:300]))
    cell = t.rows[0].cells[0]
    print("   cell width=%s" % (cell.width,))
    for p in cell.paragraphs:
        print("   cellp style=%s align=%s first=%s spacing=%s text=%r" % (
            p.style.name, p.alignment, p.paragraph_format.first_line_indent,
            p.paragraph_format.line_spacing, p.text[:40]))
        for r in p.runs:
            print("      run %r size=%s bold=%s font=%s" % (
                r.text[:30], r.font.size.pt if r.font.size else None, r.font.bold, r.font.name))
