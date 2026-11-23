# -*- coding: utf-8 -*-
"""run_all.py · 一键跑完整采集流程
=====================================================================
依次执行：
  1) Selenium 登录 + 采集第 1~10 页（含截图 01、02）
  2) 统计可视化（生成 04_category_price.png）
  3) 入库结果预览页 + 截图（03_data_stored.png）

用法：python run_all.py           # 默认 10 页
      python run_all.py 3         # 只采 3 页
前置：靶站已在 http://127.0.0.1:5000 运行（python app.py）
"""
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable


def run(name, args):
    print("\n" + "=" * 70)
    print(">>> %s" % name)
    print("=" * 70)
    t = time.time()
    r = subprocess.run([PY, "-X", "utf8"] + args, cwd=HERE)
    print("<<< %s 结束，退出码 %d，耗时 %.1fs" % (name, r.returncode, time.time() - t))
    return r.returncode


def main():
    pages = sys.argv[1] if len(sys.argv) > 1 else "10"
    steps = [
        ("第一步：Selenium 采集", ["moxuan_selenium_crawler.py", "--pages", pages]),
        ("第二步：统计可视化", ["report_chart.py"]),
        ("第三步：入库结果截图", ["make_report_assets.py"]),
    ]
    for name, args in steps:
        if run(name, args) != 0:
            print("流程在【%s】环节失败，已中止" % name)
            return 1
    print("\n全部完成：数据在 data/，截图在 screenshots/，日志在 logs/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
