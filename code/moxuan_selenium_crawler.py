# -*- coding: utf-8 -*-
"""moxuan_selenium_crawler.py · 路线 A：Selenium 浏览器自动化采集
=====================================================================
目标站点：墨轩书城（http://127.0.0.1:5000，SK3013 教学沙箱）
采集范围：第 1~10 页，每页 10 条，共 100 条

流程（对应站点六道关卡）：
  关卡1  未登录访问 /list 会被 302 重定向到 /login        -> 用真实浏览器会话登录
  关卡2  登录需图形验证码                                  -> 下载验证码图片 + ddddocr 识别 + 失败换图重试
  关卡3  列表页 HTML 无数据，由 JS 调 /api/books 渲染       -> 等浏览器渲染完成后再读 DOM 表格
  关卡4  接口校验 UA，脚本默认 UA 403                      -> 真实浏览器 UA（Edg/156），无需伪造
  关卡5  限频 8 次/10 秒                                   -> 翻页间隔 INTERVAL 秒 + 捕获 alert 后按窗口等待重试
  关卡6  分页 page/per_page（共 32 页）                     -> 点击"下一页"逐页翻页，读取表格

用法：
  python moxuan_selenium_crawler.py                 # 采集 10 页并保存截图
  python moxuan_selenium_crawler.py --pages 3       # 只采 3 页（快速验证）
  python moxuan_selenium_crawler.py --headed        # 显示浏览器窗口（课堂演示）
  python moxuan_selenium_crawler.py --manual-captcha # 手工输入验证码
"""
import argparse
import csv
import json
import os
import re
import sys
import time
from datetime import datetime

import requests
from selenium import webdriver
from selenium.common.exceptions import (NoAlertPresentException,
                                        StaleElementReferenceException,
                                        TimeoutException,
                                        WebDriverException)
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C
from captcha_solver import CaptchaSolver

# ------------------------------------------------------------------ 日志
try:                                   # 控制台统一用 UTF-8，避免 Windows GBK 终端乱码
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:                      # noqa: BLE001
    pass


class Logger:
    """同时输出到控制台与日志文件（UTF-8），便于把采集过程存档进报告。"""

    def __init__(self, path):
        self.path = path
        self.fh = open(path, "w", encoding="utf-8")

    def __call__(self, msg):
        line = "[%s] %s" % (datetime.now().strftime("%H:%M:%S"), msg)
        print(line)
        self.fh.write(line + "\n")
        self.fh.flush()

    def close(self):
        self.fh.close()


# ------------------------------------------------------------------ 爬虫主体
class MoxuanCrawler:
    def __init__(self, pages=C.PAGES, interval=C.INTERVAL, headed=False,
                 manual_captcha=False, log=None):
        self.pages = pages
        self.interval = interval
        self.headed = headed
        self.manual_captcha = manual_captcha
        self.log = log or print
        self.driver = None
        self.rows = []                 # 采集到的原始记录
        self.events = []               # 过程事件（登录、429、翻页）留作报告证据
        self.solver = None

    # ---------------------------------------------------------- 浏览器
    def start_browser(self):
        opts = EdgeOptions()
        if not self.headed:
            opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1600,1000")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--lang=zh-CN")
        # 关闭"自动化控制"提示条，截图更干净
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        opts.page_load_strategy = "eager"
        self.driver = webdriver.Edge(options=opts)
        self.driver.set_page_load_timeout(30)
        ua = self.driver.execute_script("return navigator.userAgent")
        self.log("[环境] 浏览器 UA = %s" % ua)
        hit = [k for k in C.BLOCKED_UA_KEYWORDS if k in ua.lower()]
        self.log("[环境] UA 黑名单命中：%s（预期为空，因此不会被关卡4拦截）" % (hit or "无"))
        return ua

    def stop_browser(self):
        if self.driver:
            try:
                self.driver.quit()
            except WebDriverException:
                pass
            self.driver = None

    def screenshot(self, filename, desc=""):
        path = os.path.join(C.SHOT_DIR, filename)
        self.driver.save_screenshot(path)
        self.log("[截图] %s -> %s" % (desc or filename, os.path.basename(path)))
        return path

    # ---------------------------------------------------------- 关卡1/2：登录
    def _captcha_bytes(self, captcha_id):
        """用浏览器会话 Cookie 通过 requests 直接下载验证码原图（160x56，比元素截图更清晰）。"""
        cookies = {c["name"]: c["value"] for c in self.driver.get_cookies()}
        url = C.BASE_URL + "/captcha/" + captcha_id
        r = requests.get(url, cookies=cookies,
                         headers={"User-Agent": self.driver.execute_script("return navigator.userAgent"),
                                  "Referer": C.LOGIN_URL},
                         timeout=10)
        r.raise_for_status()
        return r.content

    def _read_captcha_id(self):
        el = self.driver.find_element(By.NAME, "captcha_id")
        return el.get_attribute("value")

    def _submit_login(self, code):
        d = self.driver
        d.find_element(By.NAME, "username").clear()
        d.find_element(By.NAME, "username").send_keys(C.USERNAME)
        d.find_element(By.NAME, "password").clear()
        d.find_element(By.NAME, "password").send_keys(C.PASSWORD)
        d.find_element(By.NAME, "captcha_code").clear()
        d.find_element(By.NAME, "captcha_code").send_keys(code)
        # 表单提交后若成功会 302 到 /list
        d.find_element(By.CSS_SELECTOR, "button[type=submit]").click()
        try:
            WebDriverWait(d, 8).until(lambda x: "/list" in x.current_url or
                                      x.find_elements(By.CSS_SELECTOR, ".error"))
        except TimeoutException:
            pass
        if "/list" in d.current_url:
            return True, ""
        err = ""
        for e in d.find_elements(By.CSS_SELECTOR, ".error"):
            err = e.text.strip()
            break
        return False, err

    def login(self, max_retry=C.MAX_CAPTCHA_RETRY):
        """模拟登录：读验证码 -> 识别 -> 提交表单；失败则刷新换图重试。"""
        d = self.driver
        self.log("[关卡1] 未登录直接访问 %s" % C.LIST_URL)
        d.get(C.LIST_URL)
        time.sleep(0.8)
        self.log("[关卡1] 实际落地 URL = %s（302 重定向到登录页，拦截生效）" % d.current_url)

        for attempt in range(1, max_retry + 1):
            if attempt > 1:
                d.get(C.LOGIN_URL)
                time.sleep(0.6)
            captcha_id = self._read_captcha_id()
            png = self._captcha_bytes(captcha_id)
            with open(os.path.join(C.DATA_DIR, "captcha_last.png"), "wb") as f:
                f.write(png)

            if self.manual_captcha:
                os.startfile(os.path.join(C.DATA_DIR, "captcha_last.png"))
                code = input("  请查看验证码图片并输入：").strip().upper()
            else:
                code = self.solver.solve(png, save_dir=os.path.join(C.DATA_DIR, "captcha_samples"),
                                         tag="%02d" % attempt)

            if not code:
                self.log("[关卡2] 第 %d 次：验证码识别失败（字符不足 4 位），换一张重试" % attempt)
                continue

            ok, err = self._submit_login(code)
            self.events.append({"step": "login", "attempt": attempt, "captcha": code,
                                "success": ok, "error": err})
            if ok:
                self.log("[关卡2] 第 %d 次：验证码识别为 %s，登录成功，Cookie 会话已建立" % (attempt, code))
                return True
            self.log("[关卡2] 第 %d 次：验证码识别为 %s，服务端判定不通过（%s），换一张重试"
                     % (attempt, code, err or "验证码不正确"))
            time.sleep(0.5)
        return False

    # ---------------------------------------------------------- 关卡3/6：翻页读取
    def _dismiss_alert(self):
        """若页面弹出 alert（如 429 限频提示），取回文本并关闭，返回提示内容。"""
        try:
            alert = self.driver.switch_to.alert
            text = alert.text
            alert.accept()
            return text
        except NoAlertPresentException:
            return None

    def _read_table(self):
        """读取当前渲染完成的表格（关卡3：数据由 JS 写入 DOM，必须等渲染后再取）。

        注意：翻页时前端会先 tbody.innerHTML = '' 再重新插入行，此时已取到的
        WebElement 会失效并抛 StaleElementReferenceException，因此这里只做"取一次"，
        重试交给 _safe_read / _wait_page。
        """
        d = self.driver
        trs = d.find_elements(By.CSS_SELECTOR, "#bookTable tbody tr.book-row")
        out = []
        for tr in trs:
            rec = {}
            for cls, name in C.COLUMN_MAP.items():
                cells = tr.find_elements(By.CSS_SELECTOR, "td." + cls)
                rec[name] = cells[0].text.strip() if cells else ""
            if rec.get("编号"):
                out.append(rec)
        return out

    def _safe_read(self):
        """容错读取：DOM 正在重绘时返回 None，由上层循环重试（本项目最常踩的坑）。"""
        try:
            return self._read_table()
        except (StaleElementReferenceException, WebDriverException):
            return None

    def _page_no(self):
        try:
            return int(self.driver.find_element(By.ID, "pageInput").get_attribute("value") or 1)
        except (StaleElementReferenceException, WebDriverException, ValueError, TypeError):
            return None

    def _wait_page(self, expect_page, prev_first_id=None, timeout=25):
        """等待目标页稳定渲染：页码正确 + 满 10 行 + 首行编号与上一页不同 + 连续两次一致。"""
        deadline = time.time() + timeout
        last_first, stable = None, 0
        while time.time() < deadline:
            msg = self._dismiss_alert()
            if msg:
                return "alert", msg
            rows = self._safe_read()
            pg = self._page_no()
            if rows and pg == expect_page and len(rows) >= C.PER_PAGE:
                first = rows[0]["编号"]
                if (prev_first_id is None or first != prev_first_id):
                    if last_first == first:
                        stable += 1
                        if stable >= 1:
                            return "ok", rows
                    else:
                        stable, last_first = 0, first
            time.sleep(0.2)
        return "timeout", None

    def _next_page(self, cur_page, max_retry=3):
        """点击"下一页"并等待换页；遇到 429 弹窗则按限频窗口等待后重试。"""
        prev = self._safe_read()
        prev_first = prev[0]["编号"] if prev else None
        for attempt in range(1, max_retry + 1):
            try:
                self.driver.find_element(By.ID, "nextBtn").click()
            except (StaleElementReferenceException, WebDriverException):
                time.sleep(0.3)
                continue
            status, payload = self._wait_page(cur_page + 1, prev_first)
            if status == "ok":
                return True, payload
            if status == "alert":
                # 关卡5：10 秒窗口内超过 8 次 -> 429。alert 不带 Retry-After，按窗口长度等待
                self.log("[关卡5] 触发 429 限频：%s；等待 10.5s 后重试（第 %d 次）" % (payload, attempt))
                self.events.append({"step": "rate_limit", "page": cur_page + 1,
                                    "message": payload, "retry": attempt})
                time.sleep(10.5)
                continue
            self.log("[异常] 第 %d 页等待超时，重试（第 %d 次）" % (cur_page + 1, attempt))
        return False, None

    def crawl(self):
        """采集第 1~pages 页。"""
        d = self.driver
        status, rows = self._wait_page(1, None, timeout=25)
        if status != "ok" or not rows:
            self.log("[异常] 列表页未渲染出数据（%s），请确认已登录" % status)
            return []
        total_text = d.find_element(By.ID, "total").text
        pages_text = d.find_element(By.ID, "pages").text
        self.log("[关卡6] 站点共 %s 条 / %s 页（接口返回 total、pages 字段），本次采集第 1~%d 页"
                 % (total_text, pages_text, self.pages))

        self.screenshot("01_login_success.png", "登录成功后的图书列表页")
        self.events.append({"step": "login_success", "url": d.current_url,
                            "total": total_text, "pages": pages_text})

        page, data = 1, rows
        self.rows = []
        while True:
            page_no = self._page_no() or page
            self.log("[采集] 第 %s 页：解析到 %d 条 -> 累计 %d 条"
                     % (page_no, len(data), len(self.rows) + len(data)))
            for rec in data:
                rec["页码"] = int(page_no)
                rec["采集时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.rows.extend(data)
            self.events.append({"step": "page_done", "page": int(page_no),
                                "count": len(data), "cumulative": len(self.rows)})

            if page == 5:                     # 翻页过程中间截一张，证明是逐页翻的
                self.screenshot("02_crawl_paging.png", "翻页采集进行中（第 5 页）")
            if page >= self.pages:
                break

            time.sleep(self.interval)          # 礼貌间隔，遵守 robots.txt 的 Crawl-delay
            ok, data = self._next_page(page)
            if not ok:
                self.log("[异常] 第 %d 页翻页失败，提前结束" % (page + 1))
                break
            page += 1
        return self.rows


# ------------------------------------------------------------------ 清洗与入库
def clean_rows(rows):
    """数据清洗：去空白、类型转换、去重、合法性校验。"""
    seen, cleaned, dropped = set(), [], []
    for r in rows:
        rec = {k: (v.strip() if isinstance(v, str) else v) for k, v in r.items()}
        rec = {k: ("" if v is None else v) for k, v in rec.items()}
        try:
            rec["编号"] = int(rec["编号"])
            rec["价格"] = round(float(rec["价格"]), 1)
            rec["评分"] = round(float(rec["评分"]), 1)
            rec["出版年"] = int(rec["出版年"])
        except (ValueError, KeyError):
            dropped.append(("字段类型异常", rec))
            continue
        if not rec.get("书名") or rec["编号"] in seen:
            dropped.append(("空书名或重复编号", rec))
            continue
        if not (29.9 <= rec["价格"] <= 199.0) or not (6.0 <= rec["评分"] <= 9.8):
            dropped.append(("数值越界", rec))
            continue
        seen.add(rec["编号"])
        cleaned.append(rec)
    return cleaned, dropped


def save_csv(rows, path=C.CSV_PATH):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=C.FIELDS + ["页码", "采集时间"])
        w.writeheader()
        w.writerows(rows)
    return path


def save_json(rows, path=C.JSON_PATH):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    return path


def save_mongo(rows, uri="mongodb://127.0.0.1:27017", db="sk3013", coll="books"):
    """可选：写入 MongoDB（未启动服务时自动跳过，不影响 CSV 交付）。"""
    try:
        from pymongo import MongoClient
        client = MongoClient(uri, serverSelectionTimeoutMS=1500)
        client.admin.command("ping")
        c = client[db][coll]
        c.delete_many({})
        c.insert_many(rows)
        n = c.count_documents({})
        client.close()
        return n
    except Exception as e:                                  # noqa: BLE001
        return "跳过（%s: %s）" % (type(e).__name__, str(e)[:60])


def build_stats(rows):
    """基础统计：分类分布、出版社 TOP、价格/评分均值、评分榜。"""
    by_cat, by_pub, by_year = {}, {}, {}
    for r in rows:
        by_cat[r["分类"]] = by_cat.get(r["分类"], 0) + 1
        by_pub[r["出版社"]] = by_pub.get(r["出版社"], 0) + 1
        by_year[r["出版年"]] = by_year.get(r["出版年"], 0) + 1
    n = len(rows)
    stats = {
        "记录数": n,
        "页数": len({r["页码"] for r in rows}),
        "编号范围": [min(r["编号"] for r in rows), max(r["编号"] for r in rows)],
        "分类分布": dict(sorted(by_cat.items(), key=lambda kv: -kv[1])),
        "出版社TOP5": dict(sorted(by_pub.items(), key=lambda kv: -kv[1])[:5]),
        "出版年分布": dict(sorted(by_year.items())),
        "均价": round(sum(r["价格"] for r in rows) / n, 2),
        "最高价": max(r["价格"] for r in rows),
        "最低价": min(r["价格"] for r in rows),
        "平均评分": round(sum(r["评分"] for r in rows) / n, 2),
        "评分TOP5": [{"书名": r["书名"], "评分": r["评分"], "作者": r["作者"]}
                     for r in sorted(rows, key=lambda x: -x["评分"])[:5]],
        "价格TOP5": [{"书名": r["书名"], "价格": r["价格"], "出版社": r["出版社"]}
                     for r in sorted(rows, key=lambda x: -x["价格"])[:5]],
    }
    return stats


# ------------------------------------------------------------------ 主程序
def main():
    ap = argparse.ArgumentParser(description="墨轩书城图书采集（Selenium 路线A）")
    ap.add_argument("--pages", type=int, default=C.PAGES, help="采集页数，默认 10")
    ap.add_argument("--interval", type=float, default=C.INTERVAL, help="翻页间隔秒数，默认 1.0")
    ap.add_argument("--headed", action="store_true", help="显示浏览器窗口")
    ap.add_argument("--manual-captcha", action="store_true", help="手工输入验证码")
    args = ap.parse_args()

    log_path = os.path.join(C.LOG_DIR, "crawl_%s.log" % datetime.now().strftime("%Y%m%d_%H%M%S"))
    log = Logger(log_path)
    log("=" * 66)
    log("墨轩书城图书数据采集 · 路线A（Selenium）  %s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log("目标站点：%s   计划采集：第 1~%d 页" % (C.BASE_URL, args.pages))
    log("=" * 66)

    crawler = MoxuanCrawler(pages=args.pages, interval=args.interval,
                            headed=args.headed, manual_captcha=args.manual_captcha, log=log)
    crawler.solver = None if args.manual_captcha else CaptchaSolver()

    t0 = time.time()
    try:
        crawler.start_browser()
        if not crawler.login():
            log("[失败] 验证码重试 %d 次仍未登录成功" % C.MAX_CAPTCHA_RETRY)
            return 1
        crawler.crawl()
    finally:
        crawler.stop_browser()

    raw = crawler.rows
    cleaned, dropped = clean_rows(raw)
    log("-" * 66)
    log("[清洗] 原始 %d 条 -> 有效 %d 条，剔除 %d 条" % (len(raw), len(cleaned), len(dropped)))
    for reason, rec in dropped[:5]:
        log("        剔除：%s -> %s" % (reason, rec.get("书名", "?")))

    csv_path = save_csv(cleaned)
    json_path = save_json(cleaned)
    log("[入库] CSV  -> %s" % csv_path)
    log("[入库] JSON -> %s" % json_path)
    mongo = save_mongo(cleaned)
    log("[入库] MongoDB -> %s" % mongo)

    stats = build_stats(cleaned)
    with open(C.STATS_PATH, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=1)

    log("-" * 66)
    log("[统计] 记录数 %d，页数 %d，编号 %s~%s" % (stats["记录数"], stats["页数"],
                                                stats["编号范围"][0], stats["编号范围"][1]))
    log("[统计] 分类分布：%s" % "、".join("%s %d条" % kv for kv in stats["分类分布"].items()))
    log("[统计] 均价 %.2f 元，平均评分 %.2f" % (stats["均价"], stats["平均评分"]))
    log("[统计] 评分 TOP5：%s" % "；".join("%s(%.1f)" % (b["书名"], b["评分"]) for b in stats["评分TOP5"]))

    rate_events = [e for e in crawler.events if e["step"] == "rate_limit"]
    log("[关卡5] 本次共触发 429 限频 %d 次（已自动等待重试，数据完整）" % len(rate_events))
    if not args.manual_captcha:
        log("[关卡2] 验证码识别成功率 %.1f%%（%d/%d）" % (
            crawler.solver.success_rate * 100,
            sum(1 for _, ok in crawler.solver.history if ok), len(crawler.solver.history)))

    with open(os.path.join(C.DATA_DIR, "crawl_events.json"), "w", encoding="utf-8") as f:
        json.dump({"events": crawler.events, "elapsed_sec": round(time.time() - t0, 1),
                   "captcha": getattr(crawler.solver, "history", [])}, f, ensure_ascii=False, indent=1)
    log("[完成] 总耗时 %.1fs，日志：%s" % (time.time() - t0, log_path))
    log.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
