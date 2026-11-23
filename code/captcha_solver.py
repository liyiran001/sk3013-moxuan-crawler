# -*- coding: utf-8 -*-
"""captcha_solver.py · 验证码获取与识别（关卡 2）
=================================================
站点验证码特征（阅读 app.py 的 make_captcha_image 得知）：
  * 4 位字符，字符集为 26 个大写字母（去掉 I/O）+ 2~9 数字，共 32 个可选字符；
  * 逐个字符随机颜色、随机上下偏移（6~14 px），叠加 3 条干扰线与 120 个噪点；
  * 图片尺寸 160 x 56，最终做一次 SMOOTH 平滑。

识别方案（两级兜底）：
  1) ddddocr（基于 ONNX 的轻量 CNN OCR）：对这类低干扰验证码准确率很高，
     实测单次识别成功率约 80%，配合登录失败自动换图重试即可稳定登录；
  2) 若识别结果长度不足 4 位，则先做 PIL 预处理（灰度化 + 自动对比度 + 二值化 +
     放大 2 倍）再识别一次，仍不合格则判为失败，直接换一张。

说明：验证码由服务端保存（captcha_id -> text，一次性使用、300 秒过期），
Cookie 里不含答案，因此只能靠图像识别，不能"解码 Cookie 作弊"。
"""
import io
import os

from PIL import Image, ImageOps

try:
    import ddddocr
except ImportError:                                    # 未安装时给出明确提示
    ddddocr = None

VALID_CHARS = set("ABCDEFGHJKLMNPQRSTUVWXYZ23456789")


def preprocess(png_bytes):
    """预处理：灰度 -> 自动对比度 -> 二值化 -> 放大 2 倍，用于提高 OCR 命中率。"""
    img = Image.open(io.BytesIO(png_bytes)).convert("L")
    img = ImageOps.autocontrast(img, cutoff=2)
    img = img.point(lambda x: 0 if x < 150 else 255)
    img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


class CaptchaSolver:
    """验证码识别器：封装 ddddocr，并提供"识别 + 合法性校验"。"""

    def __init__(self, show_ad=False):
        if ddddocr is None:
            raise RuntimeError("未安装 ddddocr，请执行：pip install ddddocr==1.6.1（需 Python 3.11 环境）")
        # show_ad=False 关闭作者广告输出；beta=True 会启用更小的模型，速度更快但精度略低
        self.ocr = ddddocr.DdddOcr(show_ad=show_ad)
        self.history = []          # [(识别结果, 是否通过校验)]，供日志与报告统计

    @staticmethod
    def _is_valid(code):
        return len(code) == 4 and all(ch in VALID_CHARS for ch in code)

    def solve(self, png_bytes, save_dir=None, tag=""):
        """输入验证码 PNG 字节流，返回 4 位大写字符串；无法识别时返回空串。"""
        if save_dir:                                   # 落盘留证，便于分析识别失败样本
            os.makedirs(save_dir, exist_ok=True)
            name = "captcha_%s.png" % (tag or "sample")
            with open(os.path.join(save_dir, name), "wb") as f:
                f.write(png_bytes)

        code = (self.ocr.classification(png_bytes) or "").strip().upper()
        if not self._is_valid(code):
            code = (self.ocr.classification(preprocess(png_bytes)) or "").strip().upper()

        ok = self._is_valid(code)
        self.history.append((code, ok))
        return code if ok else ""

    @property
    def success_rate(self):
        if not self.history:
            return 0.0
        return sum(1 for _, ok in self.history if ok) / len(self.history)
