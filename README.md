# 墨轩书城图书数据采集 · SK3013 课程设计

> 课程：《Python 网络爬虫技术》（SK3013）  
> 学生：黎展鹏　学号：2520148249　教学班：SK3013-2  
> 目标站点：墨轩书城（课程组自建教学沙箱，http://127.0.0.1:5000）  
> 技术路线：**路线 A —— Selenium 浏览器自动化登录并翻页采集**  
> 进阶选做：**已完成**——黑马头条管理后台（极验滑块登录 + Token 分页采集 100 篇）

---

## 一、项目简介

本课程设计在本地启动课程发放的教学靶站"墨轩书城"，用 Selenium 驱动真实 Edge 浏览器
完成 **模拟登录（图形验证码识别）→ 逐页翻页 → 解析表格 → 清洗 → CSV/JSON 入库 → 统计可视化**
的全流程，采集站点第 1~10 页共 **100 条**图书数据（站点共 320 条 / 32 页），并逐条应对站点设置的六道反爬关卡。

| 关卡 | 站点行为 | 本项目的应对 |
| --- | --- | --- |
| 1 | 未登录访问 `/list`、`/api/books` → 302 / 401 | 用真实浏览器会话登录，Cookie 由浏览器自动维护 |
| 2 | 登录需 4 位图形验证码（一次性、300 秒过期） | 解析隐藏域 `captcha_id` → 下载验证码原图 → ddddocr 识别 → 失败自动换图重试 |
| 3 | 列表页 HTML 不含数据，由前端 JS 调 `/api/books` 渲染 | 等浏览器把表格渲染完成后再读 DOM，不直接 `requests.get` 列表页 |
| 4 | 接口校验 UA，脚本默认 UA 返回 403 | 使用真实浏览器 UA（`Edg/156`），天然不命中黑名单 |
| 5 | 接口限频：10 秒内最多 8 次，超出 429 | 翻页间隔 1.0s + 捕获 429 弹窗后按 `Retry-After` 退避重试 |
| 6 | 分页 `page`/`per_page`，返回 `total`/`pages`/`items` | 点击"下一页"逐页翻，读取 `#pageInput` 与表格行校验页码 |

### 进阶选做：黑马头条（真实平台）

`code/toutiao_crawler.py` 走通了教材配套的授权练习平台
`http://mis-toutiao-python.itheima.net/`：

| 环节 | 做法 |
| --- | --- |
| 极验滑块 | 取三块 canvas 像素 → `fullbg` 与 `bg` 逐像素求差定位缺口 → 算拖动距离 → 拟人轨迹拖动 |
| 登录 | 账号 `demo`；文档的动态密码 + 平台实际演示密码 `ABCdefg123` 都作为候选尝试 |
| 取数 | 从 `localStorage.token` 取 JWT → `Authorization: Bearer` 调 `/mis/v1_0/articles` 分页 |
| 结果 | 10 页共 **100 篇**文章，落盘 `data/toutiao_articles.csv` / `.json` |

> 踩坑记录：无头浏览器默认 UA（含 `HeadlessChrome`）会被平台 WAF 直接拦成 405，
> 覆盖成正常 Edge UA 后才能打开登录页。详见报告 2.7 节。

## 二、目录结构

```
2520148249-黎展鹏/
├─ README.md                 本说明
├─ requirements.txt          依赖清单（版本与实测环境一致）
├─ code/                     源码
│  ├─ config.py              全局配置（站点地址、账号、页数、间隔、路径）
│  ├─ captcha_solver.py      验证码识别（ddddocr + PIL 预处理兜底）
│  ├─ moxuan_selenium_crawler.py   主程序：登录 + 翻页采集 + 清洗 + 入库 + 统计
│  ├─ toutiao_crawler.py     进阶选做：极验滑块登录 + Token 分页采集黑马头条
│  ├─ report_chart.py        统计可视化（四合一图表）
│  ├─ make_report_assets.py  生成"入库结果"预览页并整页截图
│  ├─ rate_limit_probe.py    关卡5 限频实测脚本（验证 429 与退避恢复）
│  ├─ bench_captcha.py       验证码识别基准测试（10 轮登录统计）
│  ├─ run_all.py             一键跑完整流程（墨轩书城）
│  ├─ inspect_docx.py / inspect_runs.py   模板结构探查辅助脚本
│  ├─ build_report.py        由模板生成最终课程设计报告 docx
│  └─ docx2png.py            报告 docx → PDF → 逐页 PNG（排版自检用）
├─ data/
│  ├─ books_page1_10.csv     清洗后的 100 条图书数据（utf-8-sig，Excel 可直接打开）
│  ├─ books_page1_10.json    同数据的 JSON 版本
│  ├─ stats.json             统计结果
│  ├─ crawl_events.json      采集过程事件（登录、翻页、限频）
│  ├─ toutiao_articles.csv   进阶选做：100 篇文章
│  ├─ toutiao_articles.json  同上（JSON）
│  └─ captcha_samples/       验证码样本
├─ screenshots/
│  ├─ 01_login_success.png   登录成功后的列表页
│  ├─ 02_crawl_paging.png    翻页采集中（第 5 页）
│  ├─ 03_data_stored.png     入库结果（整页，含统计图）
│  ├─ 03_data_stored_summary.png  入库结果（汇总部分，报告插图用）
│  ├─ 04_category_price.png  统计可视化
│  ├─ 05_toutiao_login.png   黑马头条登录页（极验滑块）
│  ├─ 06_toutiao_home.png    黑马头条后台首页（登录成功）
│  └─ 06_toutiao_article_list.png  内容审核页（文章列表）
├─ logs/                     运行日志与实测输出
└─ report/
   ├─ SK3013课程设计报告-黎展鹏-2520148249.docx   最终报告（提交用）
   ├─ SK3013课程设计报告-黎展鹏-2520148249.pdf    同一份报告的 PDF 版（便于查看）
   ├─ storage_preview.html   入库结果预览页
   └─ storage_summary.html   入库结果摘要页
```

## 三、运行环境

- Windows 10/11，Microsoft Edge 156（Selenium Manager 自动匹配驱动，无需手工下载 msedgedriver）
- Python 3.10+（本次实测 Python 3.12.14）
- 依赖：`pip install -r requirements.txt`

> 模板说明中提到 ddddocr 建议 Python 3.11；本次在 **Python 3.12 + onnxruntime 1.31.1**
> 下实测 ddddocr 1.6.1 可正常推理（见 `logs/captcha_benchmark.txt`，10/10 识别成功）。

## 四、运行步骤

1. **启动靶站**（另开一个终端）
   ```bash
   cd 靶站-墨轩书城
   pip install -r requirements.txt
   python seed_data.py        # 生成 320 条种子数据（只需一次）
   python app.py              # 启动后访问 http://127.0.0.1:5000
   ```
2. **运行采集**
   ```bash
   cd code
   python moxuan_selenium_crawler.py            # 默认采集 1~10 页
   # 可选参数：
   #   --pages 3            只采 3 页（快速验证）
   #   --interval 1.0       翻页间隔秒数
   #   --headed             显示浏览器窗口（课堂演示）
   #   --manual-captcha     手工输入验证码
   ```
3. **生成图表与入库结果截图**
   ```bash
   python report_chart.py
   python make_report_assets.py
   ```
4. **一键全流程**：`python run_all.py`
5. **生成报告 docx**：`python build_report.py`（需模板文件 `SK3013课程设计报告（学生填写模板）.docx` 在上一级目录）
6. **导出 PDF / 逐页图片自检**：`python docx2png.py 报告.docx out.pdf pages/`（依赖本机已安装 WPS）
7. **进阶选做（黑马头条）**
   ```bash
   python toutiao_crawler.py          # 自动过极验滑块 + 采集 1~10 页
   python toutiao_crawler.py --pages 3 --headed   # 少量页面 + 显示浏览器
   ```

## 五、实测结果（2026-10-29）

| 指标 | 结果 |
| --- | --- |
| 采集页数 / 记录数 | 10 页 / **100 条** |
| 编号范围 | 1000 ~ 1099（连续、无缺号） |
| 清洗剔除 | 0 条（原始 100 → 有效 100） |
| 采集耗时 | 41.9 s（含登录与 10 次翻页） |
| 验证码识别 | 10 轮独立登录首轮全部成功，单张合格率 10/10 |
| 接口限频实测 | 0.2s 间隔请求时第 **9** 次开始返回 429，按 `Retry-After=9` 退避后恢复 200 |
| 均价 / 平均评分 | 119.97 元 / 7.86 分 |
| 分类分布 | 计算机、经济管理、文学、历史哲学各 13 条，其余四类各 12 条 |

**进阶选做（黑马头条）实测**

| 指标 | 结果 |
| --- | --- |
| 极验滑块 | 像素差定位缺口，一次拖动通过（缺口距离 163 px） |
| 登录 | `demo` + 演示密码 `ABCdefg123`（任务书动态密码为备选） |
| 采集页数 / 条数 | 10 页 / **100 篇** |
| 频道分布 | html 36、设计 36、开发者资讯 11、python 4、css 4、人工智能 4、软件测试 3、android 1、ios 1 |
| 审核状态 | 审核通过 92 篇、待审核 8 篇 |
| 耗时 | 86.1 s（含登录与滑块识别） |
| 落盘 | `data/toutiao_articles.csv`、`data/toutiao_articles.json` |

## 六、合规声明

1. 主体采集仅对课程组自建的**教学沙箱**站点 `127.0.0.1:5000` 进行；进阶选做只在任务书授权的
   教材练习平台"黑马头条"（mis-toutiao-python.itheima.net）上、使用课程提供的账号 `demo` 练习，
   未对任何其他网站实施抓取；
2. 采集前阅读 `/robots.txt`：`Disallow: /login`、`Crawl-delay: 1`、限频 ≤8 次/10 秒，代码据此设置翻页间隔与退避策略；
3. 站内 320 条图书数据全部由 `seed_data.py` 程序生成，不含任何真实个人信息；
4. 未直接读取 `data.json`、未修改靶站代码取数，所有数据均通过登录后的正常请求获得；
5. 验证码识别与滑块破解仅用于本次课程练习的两个授权站点，不用于任何其他第三方平台；
6. 采集过程保持礼貌频率（沙箱 1.0s/页、黑马头条 0.6s/页），不并发压测、不批量注册、不绕开鉴权。

## 七、常见问题

- **`0xC0000142` 或浏览器起不来**：确认 Edge 已安装；首次运行 Selenium 需联网下载驱动。
- **登录偶尔失败**：属正常现象（验证码识别有概率失败），程序会自动换图重试，最多 8 次。
- **429 频繁出现**：把 `--interval` 调大（如 1.5），或减少单次采集页数。
- **MongoDB 未启动**：`save_mongo()` 会自动降级为"跳过"，不影响 CSV/JSON 交付。
- **中文乱码**：CSV 使用 `utf-8-sig`，Excel 直接双击打开即可；控制台已强制 UTF-8 输出。
- **黑马头条提示"用户名或密码错误"**：平台密码可能随日期/运维调整，程序已内置候选密码列表，
  必要时在 `toutiao_crawler.py` 顶部修改 `STATIC_PWD` 或 `PWD_PREFIX`。
- **黑马头条滑块连续失败**：极验有一定风控，重跑一次即可；`--headed` 模式成功率更高。
