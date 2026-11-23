# -*- coding: utf-8 -*-
"""inspect_docx.py · 查看模板结构（段落/表格/分节顺序），用于确定报告填充位置。"""
import sys
import docx
from docx.oxml.ns import qn

path = sys.argv[1]
d = docx.Document(path)
body = d.element.body
idx_p = idx_t = 0
out = []
for child in body.iterchildren():
    tag = child.tag.split("}")[-1]
    if tag == "p":
        txt = "".join(n.text or "" for n in child.iter(qn("w:t")))
        style = ""
        pPr = child.find(qn("w:pPr"))
        if pPr is not None:
            ps = pPr.find(qn("w:pStyle"))
            if ps is not None:
                style = ps.get(qn("w:val"))
        out.append("P%03d [%s] %s" % (idx_p, style, txt[:120]))
        idx_p += 1
    elif tag == "tbl":
        rows = child.findall(qn("w:tr"))
        cells = "".join("".join(n.text or "" for n in tc.iter(qn("w:t"))) for tc in child.iter(qn("w:tc")))
        out.append("TBL%03d rows=%d text=%s" % (idx_t, len(rows), cells[:100]))
        idx_t += 1
    elif tag == "sectPr":
        out.append("<<sectPr>>")
    else:
        out.append("<<%s>>" % tag)
print("\n".join(out))
print("\nsections:", len(d.sections))
for i, s in enumerate(d.sections):
    print(i, "start_type", s.start_type, "w", s.page_width, "h", s.page_height)
