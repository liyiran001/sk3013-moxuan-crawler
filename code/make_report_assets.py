# -*- coding: utf-8 -*-
"""make_report_assets.py · 生成"入库结果"预览页并截图
=================================================
1) 读取 data/books_page1_10.csv 与 data/stats.json；
2) 渲染 report/storage_preview.html（记录概况 + 明细表 + 统计卡片 + 统计图）；
3) 用 Selenium 打开该页面整页截图 -> screenshots/03_data_stored.png。

这样"入库结果"截图展示的是**真实落盘的数据**，而不是伪造的界面。
"""
import base64
import csv
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C

HTML_HEAD = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>墨轩书城采集 · 数据入库结果</title>
<style>
  * { box-sizing: border-box; }
  body { margin:0; padding:28px 40px; width:1540px; background:#eef2f6;
         font-family:"Microsoft YaHei",-apple-system,"Segoe UI",sans-serif; color:#1f2d3d; }
  h1 { font-size:24px; margin:0 0 6px; }
  .sub { color:#6b7c93; font-size:13px; margin-bottom:18px; }
  .cards { display:flex; gap:14px; margin-bottom:18px; }
  .card { flex:1; background:#fff; border-radius:10px; padding:14px 18px;
          box-shadow:0 1px 4px rgba(20,40,80,.08); }
  .card .k { font-size:12px; color:#8492a6; }
  .card .v { font-size:22px; font-weight:700; margin-top:4px; color:#2c5d8f; }
  .panel { background:#fff; border-radius:10px; padding:18px 20px; margin-bottom:16px;
           box-shadow:0 1px 4px rgba(20,40,80,.08); }
  .panel h2 { font-size:15px; margin:0 0 12px; color:#2c3e50; }
  table { width:100%; border-collapse:collapse; font-size:12.5px; table-layout:fixed; }
  th, td { padding:6px 8px; border-bottom:1px solid #eef1f5; text-align:left; vertical-align:top;
           overflow-wrap:anywhere; }
  th { background:#f5f8fb; color:#5a6b80; font-weight:600; white-space:nowrap; }
  td.num { text-align:right; font-variant-numeric:tabular-nums; }
  .badge { background:#e8f0fa; color:#2c5d8f; border-radius:10px; padding:1px 8px; font-size:11.5px; }
  .tag { display:inline-block; background:#eaf6ee; color:#2f7d4f; border-radius:4px;
         padding:2px 8px; font-size:12px; margin-right:8px; }
  .mono { font-family:Consolas,monospace; font-size:12px; color:#3c4858; }
  img { display:block; width:100%; max-width:100%; height:auto; border-radius:8px; }
</style></head><body>
"""


def build_html(rows, stats, include_chart=True, page_width=1540):
    head = HTML_HEAD.replace("width:1540px", "width:%dpx" % page_width)
    cats = "".join('<span class="tag">%s %d 条</span>' % (k, v)
                   for k, v in stats["分类分布"].items())
    trs = []
    for r in rows:
        trs.append(
            "<tr><td>%d</td><td>%s</td><td>%s</td><td>%s</td>"
            "<td><span class='badge'>%s</span></td><td class='num'>%.1f</td>"
            "<td class='num'>%.1f</td><td class='num'>%d</td><td>%s</td></tr>"
            % (r["编号"], r["书名"], r["作者"], r["出版社"], r["分类"],
               r["价格"], r["评分"], r["出版年"], r["简介"][:34] + ("…" if len(r["简介"]) > 34 else "")))

    body = [
        "<h1>墨轩书城 · 图书数据采集结果</h1>",
        '<div class="sub">采集范围：第 1~%d 页 ｜ 站点共 320 条 / 32 页 ｜ 落盘格式：CSV(utf-8-sig) + JSON ｜ 采集时间：%s</div>'
        % (stats["页数"], rows[0].get("采集时间", "")),
        '<div class="cards">',
        '<div class="card"><div class="k">有效记录数</div><div class="v">%d 条</div></div>' % stats["记录数"],
        '<div class="card"><div class="k">编号范围</div><div class="v">%d ~ %d</div></div>'
        % (stats["编号范围"][0], stats["编号范围"][1]),
        '<div class="card"><div class="k">平均价格</div><div class="v">￥%.1f</div></div>' % stats["均价"],
        '<div class="card"><div class="k">平均评分</div><div class="v">%.2f</div></div>' % stats["平均评分"],
        "</div>",
        '<div class="panel"><h2>分类分布</h2><div>%s</div></div>' % cats,
        '<div class="panel"><h2>数据明细（前 12 条，完整数据见 data/books_page1_10.csv）</h2>',
        '<table><colgroup><col style="width:52px"><col style="width:170px"><col style="width:76px">'
        '<col style="width:132px"><col style="width:78px"><col style="width:66px"><col style="width:52px">'
        '<col style="width:60px"><col></colgroup>'
        "<thead><tr><th>ID</th><th>书名</th><th>作者</th><th>出版社</th><th>分类</th>"
        "<th>价格(¥)</th><th>评分</th><th>出版年</th><th>简介</th></tr></thead><tbody>",
        "".join(trs[:12]),
        "</tbody></table></div>",
        '<div class="panel"><h2>评分 TOP5</h2><table><thead><tr><th>书名</th><th>作者</th>'
        '<th>评分</th></tr></thead><tbody>',
        "".join("<tr><td>%s</td><td>%s</td><td>%s</td></tr>" % (b["书名"], b["作者"], b["评分"])
                for b in stats["评分TOP5"]),
        "</tbody></table></div>",
    ]
    if include_chart:
        body.append('<div class="panel"><h2>统计可视化</h2>'
                    '<img src="../screenshots/04_category_price.png"></div>')
    body += [
        '<div class="panel"><h2>落盘文件</h2><div class="mono">data/books_page1_10.csv（%d 行 + 1 行表头）<br>'
        "data/books_page1_10.json（%d 条）<br>data/stats.json</div></div>" % (stats["记录数"], stats["记录数"]),
        "</body></html>",
    ]
    return head + "\n".join(body)


def render_screenshot(html_path, out_png):
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1600,1200")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    d = webdriver.Edge(options=opts)
    try:
        d.get("file:///" + html_path.replace("\\", "/"))
        # 等所有图片加载完成再截图，否则统计图未撑开会导致整页截图被截断
        deadline = time.time() + 10
        while time.time() < deadline:
            ready = d.execute_script(
                "return document.readyState==='complete' && "
                "Array.from(document.images).every(i => i.complete && i.naturalHeight > 0)")
            if ready:
                break
            time.sleep(0.3)
        time.sleep(0.8)
        # 用 CDP 整页截图（captureBeyondViewport），避免普通截图受窗口高度限制
        metrics = d.execute_cdp_cmd("Page.getLayoutMetrics", {})
        size = metrics.get("cssContentSize") or metrics.get("contentSize")
        w = int(size["width"])
        h = int(size["height"])
        png = d.execute_cdp_cmd("Page.captureScreenshot", {
            "format": "png",
            "captureBeyondViewport": True,
            "clip": {"x": 0, "y": 0, "width": w, "height": h, "scale": 1},
        })
        with open(out_png, "wb") as f:
            f.write(base64.b64decode(png["data"]))
    finally:
        d.quit()


def main():
    with open(C.CSV_PATH, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["编号"] = int(r["编号"]); r["价格"] = float(r["价格"])
        r["评分"] = float(r["评分"]); r["出版年"] = int(r["出版年"])
    with open(C.STATS_PATH, encoding="utf-8") as f:
        stats = json.load(f)

    html = build_html(rows, stats, include_chart=True)
    with open(C.PREVIEW_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print("预览页 ->", C.PREVIEW_PATH)

    out = os.path.join(C.SHOT_DIR, "03_data_stored.png")
    render_screenshot(C.PREVIEW_PATH, out)
    print("入库结果截图（整页）->", out)

    # 报告正文只用"汇总部分"（不入图表，避免与图 4 重复且缩小后看不清）
    summary_path = os.path.join(C.REPORT_DIR, "storage_summary.html")
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(build_html(rows, stats, include_chart=False, page_width=1160))
    out2 = os.path.join(C.SHOT_DIR, "03_data_stored_summary.png")
    render_screenshot(summary_path, out2)
    print("入库结果截图（汇总）->", out2)


if __name__ == "__main__":
    main()
