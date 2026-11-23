# -*- coding: utf-8 -*-
"""report_chart.py · 采集结果的统计可视化
=================================================
输入：data/books_page1_10.csv（清洗后的数据）
输出：screenshots/04_category_price.png（四合一统计图）

四个子图分别回答：
  ① 分类分布          —— 这 100 本书集中在哪些学科？
  ② 价格区间分布      —— 书价落在什么区间，是否符合 29.9~199 元的设定？
  ③ 评分-价格散点     —— 价格高的书是否评分也高（本数据为虚构，期望无明显相关）
  ④ 出版社 TOP8       —— 哪些出版社供货最多？
"""
import csv
import os
import sys
from collections import Counter

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C

# ---------------------------------------------------------------- 中文字体
def setup_chinese_font():
    for name in ("Microsoft YaHei", "SimHei", "SimSun", "Noto Sans CJK SC"):
        try:
            font_manager.findfont(name, fallback_to_default=False)
            plt.rcParams["font.sans-serif"] = [name]
            plt.rcParams["axes.unicode_minus"] = False
            return name
        except Exception:                                  # noqa: BLE001
            continue
    plt.rcParams["axes.unicode_minus"] = False
    return "default"


def load_rows(path=C.CSV_PATH):
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["价格"] = float(r["价格"])
        r["评分"] = float(r["评分"])
        r["出版年"] = int(r["出版年"])
    return rows


def main():
    font = setup_chinese_font()
    rows = load_rows()
    print("读取 %d 条记录，字体：%s" % (len(rows), font))

    fig, axes = plt.subplots(2, 2, figsize=(14, 9.5))
    fig.suptitle("墨轩书城图书数据采集结果分析（第 1~10 页，共 %d 条）" % len(rows),
                 fontsize=17, fontweight="bold")

    # ① 分类分布（横向条形图，按数量降序）
    ax = axes[0][0]
    cats = Counter(r["分类"] for r in rows).most_common()
    names = [c for c, _ in cats][::-1]
    vals = [v for _, v in cats][::-1]
    bars = ax.barh(names, vals, color="#3b6ea5")
    ax.bar_label(bars, padding=2)
    ax.set_title("① 图书分类分布")
    ax.set_xlabel("数量（条）")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    # ② 价格区间分布
    ax = axes[0][1]
    bins = [0, 50, 80, 110, 140, 170, 200]
    ax.hist([r["价格"] for r in rows], bins=bins, color="#5aa469",
            edgecolor="white", rwidth=0.9)
    ax.set_title("② 价格区间分布（均价 %.1f 元）" % (sum(r["价格"] for r in rows) / len(rows)))
    ax.set_xlabel("价格（元）")
    ax.set_ylabel("数量（条）")
    ax.grid(axis="y", alpha=0.3)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    # ③ 评分-价格散点
    ax = axes[1][0]
    cmap = plt.get_cmap("tab10")
    cat_list = [c for c, _ in cats]
    for i, c in enumerate(cat_list):
        xs = [r["价格"] for r in rows if r["分类"] == c]
        ys = [r["评分"] for r in rows if r["分类"] == c]
        ax.scatter(xs, ys, s=42, alpha=0.85, color=cmap(i % 10), label=c)
    ax.set_title("③ 价格与评分关系")
    ax.set_xlabel("价格（元）")
    ax.set_ylabel("评分")
    ax.legend(fontsize=8, ncol=2, frameon=False, loc="lower right")
    ax.grid(alpha=0.25)

    # ④ 出版社 TOP8
    ax = axes[1][1]
    pubs = Counter(r["出版社"] for r in rows).most_common(8)[::-1]
    pn = [p for p, _ in pubs]
    pv = [v for _, v in pubs]
    bars = ax.barh(pn, pv, color="#c46a3f")
    ax.bar_label(bars, padding=2)
    ax.set_title("④ 出版社供货数量 TOP8")
    ax.set_xlabel("数量（条）")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    fig.tight_layout(rect=(0, 0, 1, 0.95))
    fig.savefig(C.CHART_PATH, dpi=150, bbox_inches="tight", facecolor="white")
    print("统计图已保存 ->", C.CHART_PATH)


if __name__ == "__main__":
    main()
