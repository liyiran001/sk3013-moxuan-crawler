# -*- coding: utf-8 -*-
"""rate_limit_probe.py · 关卡5（接口限频）实测脚本
=====================================================================
站点配置：10 秒窗口内最多 8 次 /api/books 请求，超出返回 429 并带 Retry-After。

本脚本用于**验证并记录**该机制（而不是绕过它）：
  1) 先正常登录拿到会话；
  2) 以 0.2s 的"不礼貌"间隔连续请求 12 次，记录每次的状态码；
  3) 读到 429 后按 Retry-After 等待，验证请求恢复正常；
  4) 输出实测结论，作为报告"关卡5"的分析依据。

只请求分页接口本身，不改变数据，也不修改站点代码。
"""
import os
import re
import sys
import time

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C
from captcha_solver import CaptchaSolver

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36 Edg/156.0.0.0",
    "Referer": C.LIST_URL,
}


def login(sess, solver, max_try=6):
    for _ in range(max_try):
        r = sess.get(C.LOGIN_URL, headers=HEADERS, timeout=10)
        cid = re.search(r'name="captcha_id" value="([0-9a-z]+)"', r.text).group(1)
        png = sess.get(C.BASE_URL + "/captcha/" + cid, headers=HEADERS, timeout=10).content
        code = solver.solve(png)
        r = sess.post(C.LOGIN_URL, headers=HEADERS, timeout=10,
                      data={"username": C.USERNAME, "password": C.PASSWORD,
                            "captcha_id": cid, "captcha_code": code},
                      allow_redirects=False)
        if r.status_code == 302:
            return True
        time.sleep(0.4)
    return False


def main():
    sess = requests.Session()
    print("[1] 登录 %s" % C.BASE_URL)
    if not login(sess, CaptchaSolver()):
        sys.exit("登录失败")

    print("[2] 以 0.2s 间隔连续请求 /api/books 共 12 次（模拟不节流的脚本）")
    codes, first_429, retry_after = [], None, None
    for i in range(1, 13):
        r = sess.get(C.API_URL + "?page=1&per_page=10", headers=HEADERS, timeout=10)
        codes.append(r.status_code)
        print("    第 %2d 次 -> HTTP %d%s" % (i, r.status_code,
                                             "  Retry-After=%s" % r.headers.get("Retry-After")
                                             if r.status_code == 429 else ""))
        if r.status_code == 429 and first_429 is None:
            first_429 = i
            retry_after = float(r.headers.get("Retry-After", 10))
        time.sleep(0.2)

    print("-" * 62)
    print("[结论] 状态码序列：%s" % codes)
    print("[结论] 第 %s 次请求开始被限流（配置为 10 秒内最多 8 次）" % first_429)
    if first_429:
        wait = retry_after + 0.5
        print("[3] 按 Retry-After 等待 %.1fs 后重试，验证限流窗口已滑过" % wait)
        time.sleep(wait)
        r = sess.get(C.API_URL + "?page=1&per_page=10", headers=HEADERS, timeout=10)
        print("    恢复后请求 -> HTTP %d，返回 %d 条" %
              (r.status_code, len(r.json().get("data", {}).get("items", []))))
        print("[结论] 只要按 Retry-After 退避，即可在合规前提下继续采集")


if __name__ == "__main__":
    main()
