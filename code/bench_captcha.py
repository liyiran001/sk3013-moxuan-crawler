# -*- coding: utf-8 -*-
"""bench_captcha.py · 验证码识别模块基准测试（不采集任何业务数据）
=====================================================================
目的：量化"关卡2"的识别成本，为报告提供真实数据。
做法：独立进行 N 轮登录，每轮重新获取一张验证码，记录
      * 首轮识别即登录成功的次数
      * 平均每个会话需要的尝试次数
      * 单张验证码的"识别 + 校验"成功率

说明：本脚本只反复登录/登出，不抓取 /api/books，不涉及数据采集，
      因此不会与限频（关卡5）冲突。

用法：python bench_captcha.py [轮数，默认 10]
"""
import re
import sys

import requests

sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
import config as C
from captcha_solver import CaptchaSolver

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36 Edg/156.0.0.0",
    "Referer": C.LOGIN_URL,
}


def try_login(sess, solver, max_try=6):
    """返回 (是否成功, 尝试次数, 每次识别的结果)。"""
    codes = []
    for attempt in range(1, max_try + 1):
        r = sess.get(C.LOGIN_URL, headers=HEADERS, timeout=10)
        m = re.search(r'name="captcha_id" value="([0-9a-z]+)"', r.text)
        if not m:
            return False, attempt, codes
        cid = m.group(1)
        png = sess.get(C.BASE_URL + "/captcha/" + cid, headers=HEADERS, timeout=10).content
        code = solver.solve(png)
        codes.append(code)
        r = sess.post(C.LOGIN_URL, headers=HEADERS, timeout=10,
                      data={"username": C.USERNAME, "password": C.PASSWORD,
                            "captcha_id": cid, "captcha_code": code},
                      allow_redirects=False)
        if r.status_code == 302:
            return True, attempt, codes
    return False, max_try, codes


def main():
    rounds = int(sys.argv[1]) if len(sys.argv) > 1 else 10
    solver = CaptchaSolver()
    first_ok, total_try, fails = 0, 0, 0
    for i in range(1, rounds + 1):
        sess = requests.Session()
        ok, tries, codes = try_login(sess, solver)
        total_try += tries
        if ok:
            first_ok += 1 if tries == 1 else 0
        else:
            fails += 1
        print("第 %2d 轮：%s，尝试 %d 次，识别序列 %s" %
              (i, "成功" if ok else "失败", tries, codes))
        sess.get(C.BASE_URL + "/logout", headers=HEADERS, timeout=10)

    n = len(solver.history)
    valid = sum(1 for _, ok in solver.history if ok)
    print("-" * 60)
    print("登录轮数        : %d" % rounds)
    print("首轮即成功      : %d / %d（%.0f%%）" % (first_ok, rounds, first_ok * 100.0 / rounds))
    print("平均尝试次数    : %.2f 次/轮" % (total_try / rounds))
    print("最终登录成功率  : %.0f%%" % ((rounds - fails) * 100.0 / rounds))
    print("单张识别合格率  : %d / %d（%.1f%%）" % (valid, n, valid * 100.0 / n))


if __name__ == "__main__":
    main()
