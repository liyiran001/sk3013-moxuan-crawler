# -*- coding: utf-8 -*-
"""build_report.py · 由课程模板生成最终课程设计报告
=====================================================================
读取：C:/Users/Administrator/Desktop/课程设计/SK3013课程设计报告（学生填写模板）.docx
输出：report/SK3013课程设计报告-黎展鹏-2520148249.docx

做法：
  1) 封面、实验日期、申明人签名等固定字段直接改写 run；
  2) 三个评分项对应的 1x1 表格（模板预留的作答区）填入正文、表格、代码与截图；
  3) "二、代码提交说明"中勾选 B 方式并填写 zip 文件名；
  4) 代码片段从源码文件按标记自动截取，保证报告与代码完全一致。
"""
import os
import sys

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C

TEMPLATE = r"C:\Users\Administrator\Desktop\课程设计\SK3013课程设计报告（学生填写模板）.docx"
OUT = os.path.join(C.REPORT_DIR, "SK3013课程设计报告-黎展鹏-2520148249.docx")

STUDENT = "黎展鹏"
STUDENT_ID = "2520148249"
CLASS_NAME = "SK3013-2"
SUBMIT_DATE = "2026年10月29日"

BODY_FONT = "宋体"
CODE_FONT = "Consolas"
BODY_SIZE = 12.0        # 小四
CODE_SIZE = 8.5
TABLE_SIZE = 9.5
IMG_WIDTH = 14.0        # cm，正文区宽 14.66cm


# ================================================================ 基础排版工具
def style_run(run, font=BODY_FONT, size=BODY_SIZE, bold=False, color=None, italic=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), BODY_FONT if font == BODY_FONT else font)
    rFonts.set(qn("w:ascii"), "Times New Roman" if font == BODY_FONT else font)
    rFonts.set(qn("w:hAnsi"), "Times New Roman" if font == BODY_FONT else font)
    return run


def shade_paragraph(p, fill="F2F4F7"):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def set_borders(table, sz="4", color="8C9AAB"):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def force_table_layout(table, widths_cm):
    """固定列宽布局：写入 tblW / tblLayout=fixed 与每个单元格的 tcW，避免 Word/WPS 自动重算。"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    total = int(sum(widths_cm) * 567)
    # 先移除模板可能已有的 tblW / tblLayout，避免重复节点导致 Word 报“需要修复”
    for tag in ("w:tblW", "w:tblLayout"):
        for el in tblPr.findall(qn(tag)):
            tblPr.remove(el)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(total))
    tblPr.append(tblW)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    # 表格网格
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
            gc.set(qn("w:w"), str(int(w * 567)))
    for row in table.rows:
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def minimize_paragraph(p):
    """把空段落压到最小高度（2pt），避免嵌套表格后出现大片空白。"""
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    pPr = p._p.get_or_add_pPr()
    rPr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "4")
    rPr.append(sz)
    pPr.append(rPr)


class Writer:
    """在表格单元格内顺序写入段落 / 表格 / 图片。"""

    def __init__(self, cell):
        self.cell = cell
        self.used_first = False
        # 模板的作答表格里预置了若干空段落，先清理掉，否则会在正文中间留下一大片空白
        for p in list(cell.paragraphs)[1:]:
            if not p.text.strip() and p._p.find(qn("w:r")) is None:
                p._p.getparent().remove(p._p)

    # ---------------------------------------------------------- 段落
    def _par(self):
        if not self.used_first and self.cell.paragraphs:
            self.used_first = True
            p = self.cell.paragraphs[0]
            # 清空模板留下的空 run
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
            return p
        return self.cell.add_paragraph()

    def rich(self, text, size=BODY_SIZE, indent=True, line=1.5,
             align=None, color=None, before=0, after=0):
        """写入一段文本，支持 **加粗** 语法。"""
        p = self._par()
        text = text.replace("`", "")        # 报告正文不使用 markdown 反引号
        pf = p.paragraph_format
        pf.line_spacing = line
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.first_line_indent = Pt(size * 2) if indent else Pt(0)
        p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.LEFT
        parts = text.split("**")
        for i, seg in enumerate(parts):
            if not seg:
                continue
            style_run(p.add_run(seg), size=size, bold=(i % 2 == 1), color=color)
        return p

    def p(self, text, **kw):
        return self.rich(text, **kw)

    def hh(self, text, size=12.0, before=8, after=3):
        p = self.rich(text, size=size, indent=False, line=1.5,
                      before=before, after=after)
        for r in p.runs:
            r.font.bold = True
        return p

    def code(self, text):
        lines = text.rstrip("\n").split("\n")
        for i, line in enumerate(lines):
            p = self._par()
            pf = p.paragraph_format
            pf.line_spacing = 1.0
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.left_indent = Cm(0.3)
            pf.first_line_indent = Pt(0)
            shade_paragraph(p)
            style_run(p.add_run(line if line else " "), font=CODE_FONT, size=CODE_SIZE)
        # 代码块后补一个空段，避免与正文贴太近
        p = self._par()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        return p

    def caption(self, text):
        return self.rich(text, size=10.5, indent=False, line=1.0,
                         align=WD_ALIGN_PARAGRAPH.CENTER, color=(0x44, 0x44, 0x44),
                         before=2, after=8)

    def image(self, path, width=IMG_WIDTH, cap=None):
        p = self._par()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(0)
        p.add_run().add_picture(path, width=Cm(width))
        if cap:
            self.caption(cap)
        return p

    def table(self, headers, rows, widths=None, cap=None, align_center_cols=None,
              font_size=TABLE_SIZE):
        if cap:
            self.caption(cap)
        t = self.cell.add_table(rows=1 + len(rows), cols=len(headers))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        set_borders(t)
        # 表头
        for j, h in enumerate(headers):
            cell = t.cell(0, j)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].paragraph_format.space_before = Pt(1)
            cell.paragraphs[0].paragraph_format.space_after = Pt(1)
            style_run(cell.paragraphs[0].add_run(h), size=font_size, bold=True)
            tcPr = cell._tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "E8EEF5")
            tcPr.append(shd)
        # 数据行
        for i, row in enumerate(rows, start=1):
            for j, val in enumerate(row):
                cell = t.cell(i, j)
                para = cell.paragraphs[0]
                para.paragraph_format.space_before = Pt(1)
                para.paragraph_format.space_after = Pt(1)
                para.paragraph_format.line_spacing = 1.0
                if align_center_cols and j in align_center_cols:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                style_run(para.add_run(str(val)), size=font_size)
        if widths:
            force_table_layout(t, widths)
        repeat_header(t.rows[0])
        for row in t.rows:                     # 行内不跨页，避免半行被切到下一页
            trPr = row._tr.get_or_add_trPr()
            el = OxmlElement("w:cantSplit")
            el.set(qn("w:val"), "true")
            trPr.append(el)
        # 嵌套表格后必须有段落，否则 Word 会提示修复；这里压到最小高度
        tail = self.cell.add_paragraph()
        minimize_paragraph(tail)
        return t


# ================================================================ 代码片段抽取
def excerpt(filename, start_marker, end_marker):
    """从源码中截取 [start_marker 所在行, end_marker 所在行] 之间的代码，保证与源码一致。"""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    with open(path, encoding="utf-8") as f:
        lines = f.read().split("\n")
    s = e = None
    for i, ln in enumerate(lines):
        if s is None and start_marker in ln:
            s = i
        elif s is not None and end_marker in ln and i > s:
            e = i
            break
    if s is None or e is None:
        raise RuntimeError("代码片段未找到：%s [%s .. %s]" % (filename, start_marker, end_marker))
    block = lines[s:e + 1]
    # 去掉公共缩进
    indents = [len(b) - len(b.lstrip()) for b in block if b.strip()]
    cut = min(indents) if indents else 0
    return "\n".join(b[cut:] if b.strip() else "" for b in block)


# ================================================================ 封面等固定字段
def fill_cover(doc):
    ps = doc.paragraphs

    def set_run(p, idx, text):
        runs = p.runs
        if idx < len(runs):
            runs[idx].text = text
        else:
            p.add_run(text)

    # P008 学生姓名（run1 为标签，run2 为下划线空位）
    set_run(ps[8], 2, "   %s        " % STUDENT)
    set_run(ps[9], 1, " %s          " % STUDENT_ID)
    set_run(ps[10], 1, " %s            " % CLASS_NAME)
    set_run(ps[11], 1, " %s      " % SUBMIT_DATE)

    # 申明人签名：在原有空格文本后追加姓名
    p18 = ps[18]
    if p18.runs:
        r = p18.runs[0]
        r.text = r.text + STUDENT
    else:
        style_run(p18.add_run(STUDENT))

    # 实验日期
    p24 = ps[24]
    style_run(p24.add_run("   " + SUBMIT_DATE), size=12.0)

    # 代码提交说明：勾选 B
    for p in doc.paragraphs:
        if p.text.startswith("□ B."):
            p.runs[0].text = p.text.replace("□ B.", "√ B.") + "2520148249-黎展鹏.zip"
        elif p.text.startswith("□ A."):
            p.runs[0].text = p.text + "（未采用，本次以 B 方式的 zip 包提交）"


# ================================================================ 第 1 部分
def part1(w):
    w.hh("1.1  选题说明与目标站点概况")
    w.p("本课程设计的统一题目为**《墨轩书城图书数据采集》**。目标站点“墨轩书城”是课程组自建的"
        "模拟网上书店，属于**教学沙箱**：站点代码随课程材料发放，运行在每一位同学自己的电脑上"
        "（http://127.0.0.1:5000），学生既是“站长”也是“爬虫手”。")
    w.p("站点后端基于 **Flask + Jinja2**，前端为原生 JavaScript；内置 **320 条**由程序生成的虚构图书数据"
        "（`seed_data.py` 以固定随机种子 2026 生成，全班数据一致，便于教师核对），按每页 10 条共 **32 页**"
        "对外提供。每条记录包含编号、书名、作者、出版社、分类、价格、评分、出版年、简介共 9 个字段。"
        "按任务要求，本次采集**第 1~10 页、共 100 条**（编号 1000~1099）。")
    w.p("站点在功能上只有首页、登录页、列表页三个页面和一个数据接口 `/api/books`，却把课程中的"
        "**会话机制、图形验证码、动态渲染、请求头校验、接口限频、分页遍历**六个知识点全部串联了起来，"
        "因此它既是一个“可以爬”的站点，也是一个“必须按规矩爬”的站点。")

    w.hh("1.2  采集前的合规准备")
    w.p("动手之前先读取了站点的 `/robots.txt`，内容为：`User-agent: *`、`Disallow: /login`、"
        "`Crawl-delay: 1`，并注明“允许登录后采集 /api/books 的分页数据，请控制频率（<=8 次/10 秒）”。"
        "据此确定三条自我约束：① 不请求 `/login` 之外的任何“禁区”；② 翻页间隔不小于 1 秒；"
        "③ 一旦收到 429 立即按 `Retry-After` 退避，绝不并发硬打。整个实验只针对本机沙箱，"
        "不涉及任何真实网站与个人信息。")

    w.hh("1.3  六道关卡的反爬机制分析")
    w.p("**关卡 1：未登录拦截（会话机制）。**站点用 Flask 的 `session` 保存登录态：未登录访问 `/list` 时，"
        "服务端直接 `redirect(url_for('login'))` 返回 **302**；未登录访问 `/api/books` 时返回 "
        "**401** 与“请先登录后访问数据接口”。这说明数据访问被会话强绑定，"
        "**任何绕过登录直接请求列表页或接口的思路都行不通**。应对：用真实浏览器完成一次登录，"
        "让浏览器自己保存并回传 Cookie，后续请求天然带会话。", before=3)
    w.p("**关卡 2：图形验证码（一次性、服务端保存）。**登录表单包含 `username`、`password`、"
        "`captcha_id`（隐藏域）与 `captcha_code` 四个字段。验证码图片由 PIL 现场绘制：4 位字符、"
        "字符集为去掉 I/O 的 26 个大写字母加 2~9 共 32 个字符，逐字符随机颜色与上下偏移，"
        "再叠加 3 条干扰线和 120 个噪点。关键点在于**答案保存在服务端字典 `_CAPTCHA_STORE` 中，"
        "不进 Cookie**，且 `check_captcha` 弹出即删（一次性使用）、有效期 300 秒。"
        "因此解码 Cookie 或重放旧验证码都无效，只能真正“看图识字”。应对：解析隐藏域拿到 `captcha_id`，"
        "下载验证码原图交 OCR 识别，识别失败就刷新换一张重试。", before=3)
    w.p("**关卡 3：动态渲染（数据不在 HTML 里）。**打开 `/list` 查看源码，"
        "`<tbody>` 是空的，表格由页面底部的 JavaScript 在 `loadPage()` 中 `fetch('/api/books?page=N&per_page=10')` "
        "后逐行 `innerHTML` 写入。**直接用 requests 请求列表页只能拿到空表格**。"
        "应对：让浏览器执行完 JS，等 `#bookTable tbody tr.book-row` 渲染出 10 行后再读 DOM。", before=3)
    w.p("**关卡 4：User-Agent 校验。**`/api/books` 会检查请求头，命中 "
        "`python-requests`、`scrapy`、`curl`、`httpx`、`aiohttp` 等关键词即返回 **403**。"
        "应对：Selenium 驱动的是真实 Edge 浏览器，UA 为 "
        "`Mozilla/5.0 ... HeadlessChrome/156.0.0.0 ... Edg/156.0.0.0`，不在黑名单内，"
        "**不需要任何伪装**——这也是选择路线 A 的一个直接好处。", before=3)
    w.p("**关卡 5：接口限频。**`rate_limited()` 以“用户”为 key 维护 10 秒滑动窗口，"
        "窗口内第 9 次请求起返回 **429** 并附带 `Retry-After`（前端收到后 `alert` 提示）。"
        "应对：翻页间隔设为 1.0 秒做“主动节流”；万一手动/异常触发 429，则捕获弹窗、"
        "按窗口长度退避后重发同一页，保证数据不丢不重。", before=3)
    w.p("**关卡 6：分页。**接口接收 `page`、`per_page`（上限 50），返回 "
        "`total`、`page`、`per_page`、`pages`、`items` 五个字段，站点共 320 条 / 32 页。"
        "应对：以 `pages` 与实际页码 `#pageInput` 双重校验，点击“下一页”逐页推进，"
        "直到采满 10 页为止。", before=3)

    w.table(
        ["关卡", "站点机制", "触发后的现象", "本项目的应对策略"],
        [
            ["1", "session 会话校验", "302 跳登录 / 401 拒绝", "真实浏览器登录，浏览器托管 Cookie"],
            ["2", "4 位图形验证码，服务端一次性保存", "验证码错误、登录失败", "下载原图 + ddddocr 识别 + 失败换图重试"],
            ["3", "列表页由 JS 调 /api/books 渲染", "HTML 中 tbody 为空", "等 DOM 渲染出 10 行后再解析"],
            ["4", "校验 User-Agent 黑名单", "脚本默认 UA 返回 403", "使用真实浏览器 UA，不做伪装"],
            ["5", "10 秒内最多 8 次请求", "返回 429 并弹窗提示", "翻页间隔 1.0s；429 则按 Retry-After 退避重试"],
            ["6", "page / per_page 分页", "—", "点击下一页逐页翻，页码与行数双重校验"],
        ],
        widths=[1.6, 4.5, 3.8, 4.2],
        cap="表 1  站点六道关卡与本项目的应对策略",
        align_center_cols=[0],
    )

    w.hh("1.4  技术路线选型对比与理由")
    w.p("课程给出三条技术路线，我从“本项目需要跨过哪几道关卡、代价多大”的角度做了对比：")
    w.table(
        ["对比维度", "路线 A：Selenium", "路线 B：requests.Session", "路线 C：Scrapy"],
        [
            ["关卡 1 会话", "浏览器自动管理 Cookie", "手工维护 Session", "需从浏览器导出 Cookie"],
            ["关卡 2 验证码", "仍需 OCR 识别", "仍需 OCR 识别", "仍需 OCR 识别"],
            ["关卡 3 动态渲染", "天然解决（执行 JS）", "需逆向接口、自己拼 URL", "需逆向接口"],
            ["关卡 4 UA 校验", "真实浏览器 UA，免伪装", "必须伪造浏览器 UA", "必须伪造浏览器 UA"],
            ["关卡 5 限频", "翻页节流 + 弹窗退避", "读 Retry-After 退避", "DOWNLOAD_DELAY + AutoThrottle"],
            ["开发量 / 风险", "低，调试直观", "中，需分析表单与接口", "高，需工程化配置与部署"],
            ["可取证性", "可随时截图留证", "仅日志", "仅日志"],
        ],
        widths=[2.8, 3.9, 3.7, 3.6],
        cap="表 2  三条技术路线对比",
    )
    w.p("**最终选择路线 A（Selenium 自动化）**，理由如下：")
    w.p("（1）关卡 3 是本站最“硬”的一关——列表页 HTML 里根本没有数据。路线 A 直接让浏览器执行 JS，"
        "再读渲染后的 DOM，不需要逆向接口，出错概率最低；")
    w.p("（2）关卡 4 校验 UA，而 Selenium 用的是真实浏览器 UA，天生合规，不必“伪造请求头”，"
        "也从根上避免了“伪装得不像被识别”的风险；")
    w.p("（3）关卡 1、2 的登录流程与人工操作完全一致（填表、看图、提交），调试时可以随时切到"
        "有界面模式（--headed 参数）用肉眼观察，问题定位快；")
    w.p("（4）路线 B 虽然性能更好，但同样要过验证码，还要额外处理表单逆向、UA 伪装与 Cookie 维护，"
        "对 100 条数据的规模而言收益不明显；路线 C 更适合大规模、长期运行的工程，"
        "本次作为后续改进方向保留。")
    w.p("需要说明的是，路线 A 并不能“绕过”验证码与限频：关卡 2 仍要用 OCR 识别，"
        "关卡 5 仍要主动节流并在 429 时退避。这也是本设计在报告中保留了限频实测与识别率"
        "基准测试的原因。")

    w.hh("1.5  采集方案总体设计")
    w.p("整体按“**请求 → 解析 → 并发/节流 → 存储**”的流水线组织，结合本站特点细化为六个阶段："
        "① 启动无头 Edge；② 访问 `/list` 触发 302 并登录（识别验证码）；③ 等首页表格渲染完成；"
        "④ 循环点击“下一页”，每页读取 10 行并记录页码；⑤ 清洗、去重后写入 CSV/JSON；"
        "⑥ 统计并生成可视化图表。各阶段之间用显式等待衔接，避免 `sleep` 式的盲目等待。")
    w.table(
        ["阶段", "输入", "核心动作", "输出"],
        [
            ["① 启动", "—", "EdgeOptions 无头启动、窗口 1600×1000", "WebDriver 实例"],
            ["② 登录", "/login", "取 captcha_id → 下载验证码 → OCR → 提交表单", "已登录会话（Cookie）"],
            ["③ 首屏", "/list", "等 tbody 渲染 10 行、页码为 1", "第 1 页 10 条"],
            ["④ 翻页", "#nextBtn", "点击下一页 → 等页码与首行编号变化 → 读表", "第 1~10 页共 100 条"],
            ["⑤ 清洗入库", "原始 100 条", "类型转换、去重、边界校验、写 CSV/JSON", "books_page1_10.csv"],
            ["⑥ 统计分析", "CSV", "分类/价格/评分/出版社统计 + matplotlib", "stats.json + 统计图"],
        ],
        widths=[2.0, 2.9, 5.9, 3.2],
        cap="表 3  采集流水线的六个阶段",
        align_center_cols=[0],
    )
    w.p("在健壮性上做了三点设计：**等待**采用“页码 + 行数 + 首行编号连续两次一致”的稳定判定；"
        "**节流**把翻页间隔做成可配置参数（默认 1.0s），并在收到 429 时退避；"
        "**容错**对 Selenium 最常见的 `StaleElementReferenceException`（翻页瞬间 DOM 被重建）"
        "统一捕获并重试，保证不会因为一次重绘而中断整轮采集。")


# ================================================================ 第 2 部分
def part2(w):
    w.hh("2.1  运行环境与工程结构")
    w.p("实验环境为 Windows 11 + Microsoft Edge 156 + Python 3.12.14，驱动由 Selenium Manager 自动匹配，"
        "无需手工下载 msedgedriver。依赖见 `requirements.txt`：flask、requests、selenium、ddddocr、"
        "pillow、pymongo、matplotlib、lxml、beautifulsoup4 等。")
    w.table(
        ["文件", "职责"],
        [
            ["config.py", "站点地址、账号、页数、翻页间隔、字段映射等全部可调参数"],
            ["captcha_solver.py", "验证码识别：ddddocr + PIL 预处理兜底 + 合法性校验"],
            ["moxuan_selenium_crawler.py", "主程序：登录、翻页采集、清洗、入库、统计"],
            ["report_chart.py", "统计可视化（分类/价格/评分/出版社四联图）"],
            ["rate_limit_probe.py", "关卡 5 限频实测：验证 429 与退避恢复"],
            ["bench_captcha.py", "关卡 2 识别率基准测试（10 轮独立登录）"],
            ["make_report_assets.py", "生成入库结果预览页并整页截图"],
        ],
        widths=[5.2, 8.8],
        cap="表 4  主要源码文件及职责",
    )

    w.hh("2.2  模拟登录与验证码处理")
    w.p("登录是整条流水线的“闸门”。实现上分三步：")
    w.p("（1）**取号**：`GET /login` 后从隐藏域读出 `captcha_id`（24 位随机串）；"
        "（2）**取图**：把浏览器当前的 Cookie 复制给 `requests`，用同一会话请求 "
        "`/captcha/<captcha_id>`，拿到 160×56 的**原始 PNG**——这比截取元素截图清晰得多，"
        "直接提升了识别率；")
    w.p("（3）**识别与提交**：`ddddocr` 输出 4 位字符，校验字符集与长度；合格则填入表单提交，"
        "以 `302 → /list` 作为成功标志；不合格或服务端判错，就重新打开登录页换一张验证码，最多 8 次。")
    w.code(excerpt("moxuan_selenium_crawler.py",
                   "    def _captcha_bytes(self, captcha_id):",
                   "            return True, \"\""))
    w.p("验证码识别模块对结果做了**双重兜底**：先直接识别，若结果不足 4 位或含非法字符，"
        "再用 PIL 做灰度化、自动对比度、二值化并放大 2 倍后重新识别，仍不合格就判失败并换图。"
        "这样既避免把明显错误的字符提交上去浪费一次机会，也保留了统计识别率的口径。")
    w.code(excerpt("captcha_solver.py",
                   "    def solve(self, png_bytes, save_dir=None, tag=\"\"):",
                   "        return code if ok else \"\""))
    w.p("登录成功后浏览器地址栏为 `/list`，右上角显示当前账号 `student`，列表页显示“共 320 本图书”，"
        "分页控件显示“第 1 / 32 页”，与站点设定完全一致（图 1）。")
    w.image(os.path.join(C.SHOT_DIR, "01_login_success.png"), 14.2,
            "图 1  模拟登录成功后的图书列表页（右上角为登录账号 student，共 320 条 / 32 页）")

    w.hh("2.3  数据抓取与分页")
    w.p("**为什么不能直接 requests 列表页？**因为关卡 3 把数据藏在了 JS 里。Selenium 的价值就在这里："
        "它让浏览器把 `fetch('/api/books?page=N&per_page=10')` 真正执行完并把结果写进 `<tbody>`，"
        "我们再去读“已经被人眼看到的那张表”。读取时按列类名（`book-id`、`book-title`…）逐格取文本，"
        "与页面结构一一对应，不依赖脆弱的固定列序。")
    w.p("**翻页的稳定性是本次最大的工程难点。**前端 `loadPage()` 会先 `tbody.innerHTML = ''` 再逐行插入，"
        "这个瞬间之前取到的 `WebElement` 会全部失效，直接抛 "
        "`StaleElementReferenceException`。第一版代码正是因此中途崩溃。"
        "解决思路是把“翻页”改造成一个**状态机**：点击“下一页”后不再相信固定的 `sleep`，"
        "而是循环检查四个条件——页码输入框等于目标页、行数满 10、首行编号与上一页不同、"
        "且连续两次读到相同结果——全部满足才认为“这一页真的到了”。")
    w.code(excerpt("moxuan_selenium_crawler.py",
                   "    def _wait_page(self, expect_page, prev_first_id=None, timeout=25):",
                   "        return \"timeout\", None"))
    w.p("遇到 429 时，前端的 `alert` 会阻塞页面，Selenium 任何元素操作都会抛异常。"
        "因此每轮检测先尝试 `switch_to.alert`：若有弹窗就读出提示、`accept()` 关闭，"
        "按 10.5 秒（限频窗口长度）退避后重新点击“下一页”。这段逻辑同时兼顾了关卡 5 的合规要求"
        "与采集的完整性——**宁可等，也不并发硬打**。")
    w.code(excerpt("moxuan_selenium_crawler.py",
                   "            status, payload = self._wait_page(cur_page + 1, prev_first)",
                   "        return False, None"))
    w.image(os.path.join(C.SHOT_DIR, "02_crawl_paging.png"), 14.2,
            "图 2  翻页采集进行中（第 5 页，编号 1040~1049，页码显示 第 5 / 32 页）")

    w.hh("2.4  限频机制的实测验证")
    w.p("为了不只停留在“看代码猜行为”，单独写了 `rate_limit_probe.py`，"
        "以 0.2 秒的“不礼貌”间隔连续请求 `/api/books` 12 次，观察状态码变化：")
    w.code("""[2] 以 0.2s 间隔连续请求 /api/books 共 12 次（模拟不节流的脚本）
    第  1 次 -> HTTP 200
    ...
    第  8 次 -> HTTP 200
    第  9 次 -> HTTP 429  Retry-After=9
    第 10 次 -> HTTP 429  Retry-After=9
    第 11 次 -> HTTP 429  Retry-After=8
    第 12 次 -> HTTP 429  Retry-After=8
[结论] 状态码序列：[200×8, 429×4]，第 9 次请求开始被限流
[3] 按 Retry-After 等待 9.5s 后重试
    恢复后请求 -> HTTP 200，返回 10 条""")
    w.p("实测结果与站点配置**完全吻合**：10 秒窗口内前 8 次放行、第 9 次开始 429，"
        "`Retry-After` 随窗口滑动从 9 递减到 8。退避 9.5 秒后请求恢复 200。"
        "这也反证了正式采集时把间隔设为 1.0 秒是必要的——若不做节流，10 页里至少有一半会被拒。")

    w.hh("2.5  数据清洗与入库")
    w.p("采集到的是“页面上的字符串”，必须转换为可用数据。清洗规则如下表，"
        "任何一条不满足即整条剔除并记录原因（本次为 0 条剔除）：")
    w.table(
        ["字段", "清洗动作", "约束条件"],
        [
            ["编号 / 出版年", "转 int、去空白", "编号唯一，重复即剔除"],
            ["价格 / 评分", "转 float 并保留 1 位小数", "价格 ∈ [29.9, 199.0]，评分 ∈ [6.0, 9.8]"],
            ["书名 / 作者 / 出版社 / 分类", "去首尾空白", "书名不可为空"],
            ["简介", "去首尾空白、保留原文", "长度异常时仅记录不改写"],
        ],
        widths=[3.9, 5.4, 4.7],
        cap="表 5  数据清洗规则",
    )
    w.code(excerpt("moxuan_selenium_crawler.py",
                   "def clean_rows(rows):",
                   "    return cleaned, dropped"))
    w.p("入库环节把同一份清洗结果写成三种形态：**CSV**（`utf-8-sig`，Excel 双击打开不乱码，"
        "是提交与核对的主文件）、**JSON**（保留页码与采集时间，便于程序化处理）、"
        "**MongoDB**（`save_mongo()` 尝试连接 127.0.0.1:27017，本机未启动服务，"
        "函数捕获异常后自动降级为“跳过”，不影响 CSV 交付）。这种“主格式必达、可选格式降级”的写法，"
        "避免了因环境缺件导致整轮采集白跑。")
    w.image(os.path.join(C.SHOT_DIR, "03_data_stored_summary.png"), 14.0,
            "图 3  入库结果：100 条记录、编号 1000~1099、分类分布与评分 TOP5（由 CSV 与 stats.json 渲染）")

    w.hh("2.6  统计与可视化")
    w.p("统计由 `report_chart.py` 完成，输出四联图：分类分布、价格区间分布、价格—评分散点、"
        "出版社 TOP8。所有数值直接来自清洗后的 CSV，不经过任何人工填写。")
    w.image(os.path.join(C.SHOT_DIR, "04_category_price.png"), 14.2,
            "图 4  采集结果的统计可视化（分类、价格、评分与出版社四个维度）")

    # ---------------------------------------------------------------- 2.7 进阶
    w.hh("2.7  进阶选做：黑马头条管理后台（真实平台）")
    w.p("在完成沙箱站点之后，我又按任务书的要求尝试了进阶项：教材配套的真实平台"
        "**“黑马头条”管理后台**（http://mis-toutiao-python.itheima.net/）。"
        "它与墨轩书城最大的不同在于——**这是一个真实的线上系统**：验证码换成了极验 GeeTest 的"
        "滑动拼图，登录成功后拿到的是一枚 **JWT Token**，后续所有接口都要带 Token 鉴权。")
    w.p("**难点一：极验滑动验证码。**与沙箱里那 4 位字符码完全不同，极验的滑块不是“识别字符”，"
        "而是要算出拼图块应该滑到的位置，并且轨迹要像人。观察 DOM 后发现，极验在同一位置叠了三块 canvas："
        "`.geetest_canvas_fullbg`（完整背景，`display:none` 但像素已经画好）、"
        "`.geetest_canvas_bg`（挖掉缺口后的背景）、`.geetest_canvas_slice`（被拖动的拼图块）。"
        "由于画布没有被跨域图片污染，可以直接用 `toDataURL()` 把像素取出来，"
        "**把 fullbg 与 bg 逐像素求差，差值区域就是缺口**；拼图块初始贴在 x=0，"
        "所以“拖动距离 = 缺口左边界 − 拼图块左边界”。这是本题最关键的一步，"
        "比“拿拼图块去整张图里做模板匹配”稳定得多（后者在深色背景上常常匹配失败）。")
    w.code(excerpt("toutiao_crawler.py",
                   "    def _distance(self):",
                   "        return int(xs.min() - piece_x0), int(diff[ys, xs].mean())"))
    w.p("**难点二：拖动轨迹。**极验会校验滑块的移动轨迹，瞬移或匀速直线很容易被判为机器。"
        "这里按“前半段加速、后半段减速、末尾轻微回退再校正”的方式生成轨迹，"
        "每一步之间插入 10~30ms 的随机停顿，再用 `ActionChains` 的 `click_and_hold / move_by_offset / release` "
        "执行。验证通过的标志是页面出现“验证成功”且登录按钮解除 `disabled`。")
    w.code(excerpt("toutiao_crawler.py",
                   "    def _track(self, distance):",
                   "        return track"))
    w.p("**难点三：Token 复用。**登录成功后，前端把 JWT 存在 `localStorage.token` 里，"
        "后续请求通过请求头 `Authorization: Bearer <token>` 鉴权。"
        "拿到 Token 后就不必再用浏览器翻页，直接用 `requests` 调分页接口即可，"
        "速度比 Selenium 快得多——这也正好印证了设计部分对“Selenium 负责过登录、requests 负责采集”"
        "这一混合方案的分析。")
    w.code(excerpt("toutiao_crawler.py",
                   "        headers = {\"Authorization\": \"Bearer \" + self.token,",
                   "            time.sleep(0.6)"))
    w.p("关于密码：任务书写的是“固定串 623itheima.CN032@. 拼接当天日期”的动态密码，"
        "但实测该账号当前只接受演示密码 `ABCdefg123`（动态密码返回“用户名或密码错误”，"
        "而服务器时间确为当天）。因此程序里把两种密码都作为候选依次尝试，"
        "既符合任务书的设定，也能适应平台的实际状态。")
    w.p("另外还有一个值得记录的现象：**无头浏览器的 UA 会被平台的 WAF 拦截**。"
        "默认的 `HeadlessChrome` UA 打开登录页时直接返回阿里云的“访问被阻断（405）”页面，"
        "把 UA 覆盖成正常的 Edge UA 之后才恢复正常。这说明真实平台的防护不止业务层，"
        "还有基础设施层（WAF/风控），采集前必须先把“像不像正常浏览器”这件事做对。")
    w.image(os.path.join(C.SHOT_DIR, "05_toutiao_login.png"), 14.2,
            "图 5  黑马头条登录页：极验滑动拼图（程序自动算出缺口并拖动）")
    w.image(os.path.join(C.SHOT_DIR, "06_toutiao_home.png"), 14.2,
            "图 6  登录成功后的后台首页（右上角显示“欢迎 管理员”，Token 已写入 localStorage）")
    w.image(os.path.join(C.SHOT_DIR, "06_toutiao_article_list.png"), 14.2,
            "图 7  内容审核页：文章列表与分页控件（第 1~10 页共 100 条由接口采集）")


# ================================================================ 第 3 部分
def part3(w):
    w.hh("3.1  完成情况")
    w.p("课程设计的全部要求均已达成，且数据可在 `data/` 中逐条核对：")
    w.table(
        ["要求", "目标", "实际结果", "结论"],
        [
            ["采集页数", "≥10 页", "10 页（第 1~10 页）", "达成"],
            ["采集条数", "约 100 条", "100 条，编号 1000~1099 连续无缺号", "达成"],
            ["模拟登录", "过图形验证码", "10 轮独立登录首轮全部成功（10/10）", "达成"],
            ["动态渲染", "定位并获取接口数据", "等 DOM 渲染后读表，100 条全部拿到", "达成"],
            ["限频合规", "遵守 8 次/10 秒", "间隔 1.0s，429 时按 Retry-After 退避", "达成"],
            ["清洗入库", "CSV/MongoDB", "CSV + JSON 落盘，MongoDB 自动降级", "达成"],
            ["统计可视化", "基础统计或图表", "stats.json + 四联统计图", "达成"],
            ["运行截图", "登录/翻页/入库各 1 张", "共 4 张（含统计图）", "达成"],
            ["进阶选做", "黑马头条（授权）", "极验滑块登录 + Token 采集 100 篇", "达成"],
        ],
        widths=[2.9, 4.2, 5.0, 2.1],
        cap="表 6  任务要求与完成情况对照",
        align_center_cols=[3],
    )
    w.p("整轮采集从启动浏览器到 CSV 落盘耗时 **41.9 秒**（含登录、10 次翻页与 10 次节流等待），"
        "清洗阶段原始 100 条全部有效、剔除 0 条。")

    w.hh("3.2  数据结果分析")
    w.p("**分类分布高度均衡。**八个大类中，计算机、经济管理、文学、历史哲学各 13 条，"
        "其余四类各 12 条。这是因为种子数据按 `i % 8` 轮转分配分类，属于“设计出来的均匀”，"
        "而非真实书店的长尾分布——这也提醒我们：**对合成数据做业务结论要谨慎**，"
        "本次分析的价值更多在于验证采集与统计链路是否正确。")
    w.p("**价格与评分互不相关。**100 条记录均价 **119.97 元**，最高 198.6 元（《深度学习与物联网：从理论到落地》，"
        "复旦大学出版社）、最低 31.5 元；评分均值 **7.86**。从图 4 的散点可以看出，"
        "高价书并不必然高分，各分类的点在价格维度上均匀散开，说明价格与评分是两个独立生成的字段。")
    w.p("**出版社与年份分布平稳。**供货最多的是复旦大学出版社（8 条）、电子科技大学出版社（7 条），"
        "其余多为 6 条；出版年覆盖 2012—2025 年，2017 年最多（11 条），没有明显的年份断层。")
    w.p("**评分榜单**：图解生态学、图解算法、手把手学基因工程、机器学习导论：思想篇并列 9.7 分。"
        "这些书名恰好印证了数据由“学科词 + 类型词”模板拼接生成的特征。")

    w.hh("3.3  遇到的问题与解决过程")
    w.table(
        ["序号", "问题现象", "原因定位", "解决办法"],
        [
            ["1", "采集到第 3 页时抛 StaleElementReferenceException 并中断",
             "前端翻页先清空 tbody 再重建，已持有的 WebElement 失效",
             "改为“页码 + 行数 + 首行编号 + 连续两次一致”的稳定判定，并对读取统一容错重试"],
            ["2", "429 弹窗出现后所有元素操作报错",
             "alert 阻塞页面，WebDriver 无法操作 DOM",
             "每轮检测先 switch_to.alert，读文本并 accept，退避 10.5s 后重试同一页"],
            ["3", "验证码偶发识别成 3 位或含空格",
             "干扰线/噪点导致 OCR 截断",
             "增加长度与字符集校验，不合格先用 PIL 二值化放大再识别，仍不行则换图"],
            ["4", "控制台与日志中文乱码",
             "Windows 控制台默认 GBK，与 UTF-8 源码不一致",
             "程序启动即 sys.stdout.reconfigure(encoding='utf-8')，日志文件统一 UTF-8"],
            ["5", "截图右侧被裁掉、页面底部缺失",
             "普通截图受窗口尺寸限制，且图片加载早于测量",
             "改用 CDP 的 captureBeyondViewport 整页截图，并等待所有图片加载完成"],
            ["6", "MongoDB 无法连接导致程序中断",
             "本机未安装/未启动 MongoDB 服务",
             "save_mongo() 捕获异常并降级为“跳过”，CSV/JSON 照常交付"],
        ],
        widths=[1.7, 4.2, 3.5, 4.5],
        cap="表 7  开发过程中遇到的主要问题与解决方式",
        align_center_cols=[0],
    )
    w.p("其中第 1 个问题花的时间最多。起初以为是网络抖动，加了 `sleep` 也无效；"
        "后来把日志打细、在崩溃瞬间抓取页面结构，才发现是 DOM 重建导致的元素失效。"
        "这件事让我真正理解了**“等待条件”比“等待时间”更可靠**：`sleep(2)` 只是赌，"
        "而等待“页面状态满足某个可验证条件”才是确定性的。")

    w.hh("3.4  合规声明")
    w.p("本实验的主体严格限定在课程组自建的**教学沙箱** `127.0.0.1:5000` 内进行；"
        "进阶选做部分只在任务书授权的教材练习平台**黑马头条**（mis-toutiao-python.itheima.net）上、"
        "使用课程提供的账号 `demo` 进行，未对任何其他网站发起请求；"
        "采集前阅读并遵守沙箱的 `/robots.txt`（`Disallow: /login`、`Crawl-delay: 1`、≤8 次/10 秒），"
        "代码中把翻页间隔设为 1.0 秒，并对 429 做退避处理；")
    w.p("站点内 320 条图书数据全部由 `seed_data.py` 程序生成，作者、出版社等信息均为虚构，"
        "不含任何真实个人信息；")
    w.p("未直接读取 `data.json`、未修改靶站任何代码来取数，提交的 100 条数据全部来自登录后的正常请求，"
        "可通过 `logs/` 中的运行日志与 `data/crawl_events.json` 逐页复核；"
        "验证码识别与滑块破解只用于本次课程练习（沙箱站点与任务书授权的黑马头条平台），"
        "不用于任何其他第三方平台。")

    w.hh("3.5  进阶选做的结果与体会")
    w.p("进阶项同样跑通了完整链路：极验滑块一次识别（缺口距离按像素差算出，随后按拟人轨迹拖动）"
        "即通过验证，登录后从 `localStorage` 取出 JWT，用 `requests` 分页调用 "
        "`/mis/v1_0/articles` 采集 **10 页共 100 篇文章**，整轮耗时 **86.1 秒**（含登录与滑块识别）。"
        "采集结果如下：")
    w.table(
        ["指标", "结果"],
        [
            ["采集条数", "100 篇（第 1~10 页，每页 10 篇）"],
            ["频道分布", "html 36、设计 36、开发者资讯 11、python 4、css 4、人工智能 4、软件测试 3、android 1、ios 1"],
            ["审核状态", "审核通过 92 篇、待审核 8 篇"],
            ["作者分布", "python 61 篇、张三 22 篇，其余为手机号账号"],
            ["更新时间", "2018-11-29 至 2026-10-28（跨度 8 年，说明库里有历史数据）"],
            ["落盘文件", "data/toutiao_articles.csv、data/toutiao_articles.json"],
        ],
        widths=[2.8, 11.2],
        cap="表 8  进阶选做的采集结果",
    )
    w.p("对比两个站点，体会最深的有三点：①**验证码的“难度”不在字符多少，而在它校验什么**——"
        "沙箱验证码校验的是“字符对不对”，极验校验的是“位置准不准 + 轨迹像不像人”，"
        "因此解法完全不同；②**真实平台的防护不止一层**，业务层有验证码、Token、频控，"
        "基础设施层还有 WAF，连无头浏览器的 UA 都会被拦，"
        "这提醒我“请求头伪装”不是可有可无的加分项，而是能不能进门的前提；"
        "③**先过登录、再用接口**是最划算的组合：用浏览器拿到 Token 之后，"
        "采集速度比逐页点击快一个数量级，这也正是我在设计部分比较路线 A/B 时提到的思路。")

    w.hh("3.6  不足与改进方向")
    w.p("（1）**效率**：路线 A 每页要等浏览器渲染，10 页耗时 41.9 秒，比直接调接口慢一个数量级。"
        "下一步可以在登录后把浏览器 Cookie 交给 `requests.Session`，"
        "用“Selenium 登录 + requests 采集”的混合方案提速，同时保留 UA 伪装与退避。")
    w.p("（2）**工程化**：目前是单文件脚本 + 配置模块。可以迁移到 Scrapy，"
        "用 `start_requests` 携带 Cookie、`DOWNLOAD_DELAY` 与 `AutoThrottle` 控制频率，"
        "Item Pipeline 直接写 MongoDB，做成可调度、可断点续爬的工程。")
    w.p("（3）**健壮性**：验证码识别当前是“识别—失败—重试”的朴素策略，"
        "可以引入多模型投票（ddddocr 的 beta 模型 + 自训练模板匹配），进一步降低重试次数；"
        "翻页也可加入“页码跳转”作为兜底通道。")
    w.p("（4）**数据应用**：现在只做了描述性统计。后续可以按分类做价格分布检验、"
        "用出版年做时间序列，或把结果接入一个简单的可视化看板，让数据真正“用起来”。")
    w.p("总体而言，这次课设让我把课上分散的知识点串成了一条完整链路："
        "**会话与验证码**解决“能不能进”，**动态渲染与接口分析**解决“数据在哪”，"
        "**请求头与限频**解决“能不能一直拿”，**清洗与入库**解决“拿到之后怎么用”。"
        "尤其是被 StaleElement 和 429 各“卡”了一次之后，我对“爬虫的稳定性来自对页面状态的理解，"
        "而不是靠堆 sleep”这句话有了切身体会。")


# ================================================================ 主流程
def main():
    doc = docx.Document(TEMPLATE)
    fill_cover(doc)

    tables = doc.tables
    if len(tables) < 3:
        raise RuntimeError("模板结构与预期不符：应有 3 个作答表格")
    part1(Writer(tables[0].rows[0].cells[0]))
    part2(Writer(tables[1].rows[0].cells[0]))
    part3(Writer(tables[2].rows[0].cells[0]))

    doc.save(OUT)
    print("报告已生成 ->", OUT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
