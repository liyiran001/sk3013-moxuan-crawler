# -*- coding: utf-8 -*-
"""toutiao_crawler.py · 进阶选做：黑马头条管理后台（授权练习平台）
=====================================================================
目标站点：http://mis-toutiao-python.itheima.net/（教材配套真实平台，仅限授权账号练习）

与墨轩书城（教学沙箱）不同，这里是**真实线上系统**，难点集中在两处：
  1) 登录：账号 demo + 密码，滑块验证码由**极验 GeeTest** 提供（不是自家 PIL 图形码）；
  2) 数据：登录后拿到 JWT Token，再以 Token 调用 /mis/v1_0/articles 分页接口采集文章。

极验滑块的破解思路（本项目采用，仅用于该授权练习平台）：
  * 极验在同一位置渲染了三块 canvas：
      .geetest_canvas_fullbg —— 完整背景（display:none，但像素已绘制）
      .geetest_canvas_bg     —— 挖掉缺口后的背景
      .geetest_canvas_slice  —— 被拖动的拼图块
  * 用 canvas.toDataURL() 直接取出像素（同源绘制，未污染画布）；
  * fullbg 与 bg 逐像素求差，差值区域就是**缺口**的位置；
  * 滑块初始贴在 x=0，因此拖动距离 = 缺口左边界 - 拼图块左边界；
  * 用 ActionChains 按"先快后慢 + 轻微回退"的拟人轨迹拖动，校验通过后登录按钮才可点击。

用法：
  python toutiao_crawler.py                 # 采集第 1~10 页（100 条）
  python toutiao_crawler.py --pages 3       # 只采 3 页
  python toutiao_crawler.py --headed        # 显示浏览器窗口
"""
import argparse
import base64
import csv
import datetime
import io
import json
import os
import random
import sys
import time

import numpy as np
import requests
from PIL import Image
from selenium import webdriver
from selenium.common.exceptions import (NoSuchElementException,
                                        TimeoutException,
                                        WebDriverException)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C

SITE = "https://mis-toutiao-python.itheima.net/"
API_ARTICLES = "https://api-toutiao-web.itheima.net/mis/v1_0/articles"
USERNAME = "demo"
# 文档给出的动态密码规则：固定串 + 当天日期；实测该账号当前也接受静态演示密码
PWD_PREFIX = "623itheima.CN032@."
STATIC_PWD = "ABCdefg123"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36 Edg/156.0.0.0")

OUT_CSV = os.path.join(C.DATA_DIR, "toutiao_articles.csv")
OUT_JSON = os.path.join(C.DATA_DIR, "toutiao_articles.json")


# ---------------------------------------------------------------- 日志
class Logger:
    def __init__(self, path):
        self.path = path
        self.fh = open(path, "w", encoding="utf-8")

    def __call__(self, msg):
        line = "[%s] %s" % (datetime.datetime.now().strftime("%H:%M:%S"), msg)
        print(line)
        self.fh.write(line + "\n")
        self.fh.flush()

    def close(self):
        self.fh.close()


# ---------------------------------------------------------------- 极验滑块
class GeetestSolver:
    """通过 fullbg 与 bg 的像素差定位缺口，并拟人拖动。"""

    def __init__(self, driver, log=print):
        self.d = driver
        self.log = log

    def _canvas_png(self, selector):
        data = self.d.execute_script(
            "return document.querySelector(arguments[0]).toDataURL('image/png')", selector)
        return base64.b64decode(data.split(",", 1)[1])

    def _distance(self):
        full = np.array(Image.open(io.BytesIO(self._canvas_png(".geetest_canvas_fullbg")))
                        .convert("RGB")).astype(int)
        bg = np.array(Image.open(io.BytesIO(self._canvas_png(".geetest_canvas_bg")))
                      .convert("RGB")).astype(int)
        sl = np.array(Image.open(io.BytesIO(self._canvas_png(".geetest_canvas_slice")))
                      .convert("RGBA"))
        diff = np.abs(full - bg).sum(axis=2)
        ys, xs = np.where(diff > 40)
        if len(xs) == 0:
            return 0, 0
        piece_x0 = np.where(sl[:, :, 3] > 60)[1].min() if (sl[:, :, 3] > 60).any() else 0
        return int(xs.min() - piece_x0), int(diff[ys, xs].mean())

    def _track(self, distance):
        """拟人轨迹：前半段加速（8~13px/步），后半段减速（2~5px/步），末段轻微回退再校正。"""
        track, cur = [], 0
        mid = distance * random.uniform(0.55, 0.72)
        while cur < distance:
            v = random.uniform(8, 13) if cur < mid else random.uniform(2, 5)
            v = max(1.0, min(v, distance - cur))
            cur += v
            track.append(int(round(v)))
        track.append(-random.randint(1, 3))
        track.append(random.randint(1, 2))
        return track

    def solve(self):
        WebDriverWait(self.d, 15).until(
            lambda x: x.find_elements(By.CSS_SELECTOR, ".geetest_canvas_bg"))
        time.sleep(0.8)
        dist, contrast = self._distance()
        if dist <= 0:
            self.log("[极验] 未识别到缺口")
            return False
        self.log("[极验] 缺口距离 %d px（差异强度 %d）" % (dist, contrast))

        slider = self.d.find_element(By.CSS_SELECTOR, ".geetest_slider_button")
        ac = ActionChains(self.d)
        ac.move_to_element(slider).click_and_hold()
        ac.pause(random.uniform(0.1, 0.2))
        for step in self._track(dist):
            ac.move_by_offset(step, random.choice([0, 0, 0, 1, -1]))
            ac.pause(random.uniform(0.01, 0.03))
        ac.pause(random.uniform(0.15, 0.3)).release().perform()
        time.sleep(2.0)

        # 验证成功的标志：出现"验证成功"字样，或登录按钮解除 disabled
        body = self.d.find_element(By.TAG_NAME, "body").text
        if "验证成功" in body:
            self.log("[极验] 滑块验证通过")
            return True
        btn = self.d.find_element(By.ID, "inp1")
        ok = btn.get_attribute("disabled") is None
        self.log("[极验] 滑块验证%s" % ("通过" if ok else "未通过"))
        return ok


# ---------------------------------------------------------------- 爬虫主体
class ToutiaoCrawler:
    def __init__(self, pages=10, headed=False, log=print):
        self.pages = pages
        self.headed = headed
        self.log = log
        self.driver = None
        self.token = None
        self.articles = []
        self.session = requests.Session()

    def start(self):
        opts = Options()
        if not self.headed:
            opts.add_argument("--headless=new")
        for arg in ("--window-size=1600,1000", "--disable-gpu", "--no-sandbox",
                    "--lang=zh-CN", "--user-agent=" + UA):
            opts.add_argument(arg)
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.driver = webdriver.Edge(options=opts)
        self.driver.set_page_load_timeout(40)

    def stop(self):
        if self.driver:
            try:
                self.driver.quit()
            except WebDriverException:
                pass

    def shot(self, name, desc=""):
        path = os.path.join(C.SHOT_DIR, name)
        self.driver.save_screenshot(path)
        self.log("[截图] %s -> %s" % (desc or name, name))

    # -------------------------------------------------- 登录
    def _try_login(self, password):
        d = self.driver
        d.get(SITE)
        time.sleep(4)
        u = d.find_element(By.CSS_SELECTOR, "input[name=username]")
        u.clear()
        u.send_keys(USERNAME)
        p = d.find_element(By.CSS_SELECTOR, "input[name=password]")
        p.clear()
        p.send_keys(password)
        d.find_element(By.CSS_SELECTOR, "input.yzm_btn").click()
        time.sleep(2)
        try:
            d.find_element(By.CSS_SELECTOR, ".geetest_radar_btn").click()
        except NoSuchElementException:
            self.log("[登录] 未出现极验按钮")
        time.sleep(3.5)
        self.shot("05_toutiao_login.png", "黑马头条登录页（极验滑块验证码已弹出）")

        if not GeetestSolver(d, self.log).solve():
            return False, "captcha"

        try:
            d.find_element(By.ID, "inp1").click()
        except WebDriverException:
            return False, "captcha"
        try:
            WebDriverWait(d, 12).until(
                lambda x: "/home" in x.current_url or "用户名或密码错误" in x.page_source)
        except TimeoutException:
            pass
        if "/home" in d.current_url:
            return True, "ok"
        return False, "password"

    def login(self):
        candidates = [PWD_PREFIX + datetime.date.today().strftime("%Y%m%d"),
                      PWD_PREFIX + (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y%m%d"),
                      STATIC_PWD]
        for pwd in candidates:
            ok, why = self._try_login(pwd)
            if ok:
                self.token = self.driver.execute_script("return localStorage.getItem('token')")
                self.log("[登录] 成功（密码 %s…），Token 长度 %d" % (pwd[:18], len(self.token)))
                self.log("[登录] 登录后 URL：%s" % self.driver.current_url)
                return True
            self.log("[登录] 使用密码 %s 未成功（%s），换下一个候选" % (pwd[:24], why))
            time.sleep(2)
        return False

    # -------------------------------------------------- 采集
    def open_article_page(self):
        """进入后台"内容审核"页，用于截图取证。"""
        self.driver.execute_script("location.hash = '#/home/contentAudit'")
        time.sleep(5)
        self.shot("06_toutiao_article_list.png", "黑马头条后台 · 内容审核（文章列表）")
        body = self.driver.find_element(By.TAG_NAME, "body").text
        self.log("[页面] 内容审核页可见文章标题 %d 行左右" % body.count("审核通过"))

    def fetch_articles(self):
        headers = {"Authorization": "Bearer " + self.token,
                   "User-Agent": UA,
                   "Referer": SITE}
        for page in range(1, self.pages + 1):
            url = API_ARTICLES + "?page=%d&per_page=10" % page
            for attempt in range(1, 4):
                r = self.session.get(url, headers=headers, timeout=20)
                if r.status_code == 200:
                    break
                if r.status_code in (401, 403):
                    self.log("[采集] 第 %d 页 Token 失效（%d）" % (page, r.status_code))
                    return False
                wait = 2 * attempt
                self.log("[采集] 第 %d 页返回 %d，%ds 后重试" % (page, r.status_code, wait))
                time.sleep(wait)
            body = r.json()
            items = (body.get("data") or {}).get("articles") or []
            for it in items:
                self.articles.append({
                    "文章ID": it.get("article_id"),
                    "标题": it.get("title"),
                    "作者": it.get("name"),
                    "频道": it.get("channel_name"),
                    "评论数": it.get("comment_count"),
                    "更新时间": it.get("update_time"),
                    "状态": self._status_text(it.get("status")),
                    "页码": page,
                })
            self.log("[采集] 第 %d 页：%d 条 -> 累计 %d 条" % (page, len(items), len(self.articles)))
            if not items:
                break
            time.sleep(0.6)                     # 对真实平台保持礼貌间隔
        return True

    @staticmethod
    def _status_text(code):
        return {2: "审核通过", 3: "审核失败", 4: "待审核", 5: "草稿"}.get(code, "未知(%s)" % code)


# ---------------------------------------------------------------- 主程序
def main():
    ap = argparse.ArgumentParser(description="黑马头条文章采集（进阶选做）")
    ap.add_argument("--pages", type=int, default=10)
    ap.add_argument("--headed", action="store_true")
    args = ap.parse_args()

    log_path = os.path.join(C.LOG_DIR, "toutiao_%s.log"
                            % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    log = Logger(log_path)
    log("=" * 68)
    log("进阶选做 · 黑马头条管理后台文章采集  %s"
        % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log("站点：%s  计划采集：第 1~%d 页" % (SITE, args.pages))
    log("=" * 68)

    cr = ToutiaoCrawler(pages=args.pages, headed=args.headed, log=log)
    cr.start()
    t0 = time.time()
    try:
        log("[登录] 打开登录页，准备识别极验滑块验证码")
        if not cr.login():
            log("[失败] 登录未成功")
            return 1
        cr.shot("06_toutiao_home.png", "登录成功后的后台首页（显示当前账号）")
        cr.open_article_page()
        if not cr.fetch_articles():
            log("[失败] 采集过程中 Token 失效")
            return 1
    finally:
        cr.stop()

    rows = cr.articles
    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)

    from collections import Counter
    ch = Counter(r["频道"] for r in rows)
    st = Counter(r["状态"] for r in rows)
    log("-" * 68)
    log("[结果] 共采集 %d 条，写入 %s" % (len(rows), os.path.basename(OUT_CSV)))
    log("[结果] 频道分布：%s" % "、".join("%s %d" % kv for kv in ch.most_common(8)))
    log("[结果] 状态分布：%s" % "、".join("%s %d" % kv for kv in st.most_common()))
    log("[完成] 总耗时 %.1fs" % (time.time() - t0))
    log.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
