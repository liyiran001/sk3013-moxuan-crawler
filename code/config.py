# -*- coding: utf-8 -*-
"""config.py · 全局配置
=================================================
SK3013 课程设计《墨轩书城图书数据采集》——路线 A（Selenium 浏览器自动化）

所有可调参数集中在此文件，方便复现与讲解：
  * BASE_URL     目标站点地址（默认本机沙箱 http://127.0.0.1:5000）
  * ACCOUNT      教学账号
  * PAGES        采集页数（要求 >=10 页、约 100 条）
  * INTERVAL     两次翻页之间的间隔秒数（遵守 robots.txt 的 Crawl-delay，避开 429 限频）
"""
import os

# ------------------------------------------------------------------ 站点与账号
BASE_URL = os.environ.get("MOXUAN_BASE", "http://127.0.0.1:5000")
LOGIN_URL = BASE_URL + "/login"
LIST_URL = BASE_URL + "/list"
API_URL = BASE_URL + "/api/books"

USERNAME = "student"
PASSWORD = "Sk3013#2026"

# ------------------------------------------------------------------ 采集参数
PAGES = 10                 # 采集第 1~10 页（连续 10 页）
PER_PAGE = 10              # 站点固定每页 10 条
INTERVAL = 1.0             # 翻页间隔（秒）：站点限频 8 次/10 秒，1.0s 间隔在第 9 页附近会触发一次 429
MAX_CAPTCHA_RETRY = 8      # 验证码识别失败时的最大重试次数（每次换一张）

# 站点 User-Agent 黑名单关键词（关卡 4），Selenium 使用真实浏览器 UA，天然不会命中
BLOCKED_UA_KEYWORDS = ["python-requests", "python-urllib", "scrapy", "httpx", "aiohttp",
                       "curl", "wget", "java/", "go-http-client", "okhttp"]

# ------------------------------------------------------------------ 目录结构
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
SHOT_DIR = os.path.join(ROOT, "screenshots")
LOG_DIR = os.path.join(ROOT, "logs")
REPORT_DIR = os.path.join(ROOT, "report")
for _d in (DATA_DIR, SHOT_DIR, LOG_DIR, REPORT_DIR):
    os.makedirs(_d, exist_ok=True)

CSV_PATH = os.path.join(DATA_DIR, "books_page1_10.csv")
JSON_PATH = os.path.join(DATA_DIR, "books_page1_10.json")
STATS_PATH = os.path.join(DATA_DIR, "stats.json")
CHART_PATH = os.path.join(SHOT_DIR, "04_category_price.png")
PREVIEW_PATH = os.path.join(REPORT_DIR, "storage_preview.html")

# ------------------------------------------------------------------ 字段映射
# 表格列（list.html 中 <td class="book-xxx">） -> 中文字段名
COLUMN_MAP = {
    "book-id": "编号",
    "book-title": "书名",
    "book-author": "作者",
    "book-publisher": "出版社",
    "book-category": "分类",
    "book-price": "价格",
    "book-rating": "评分",
    "book-year": "出版年",
    "book-intro": "简介",
}
FIELDS = list(COLUMN_MAP.values())
