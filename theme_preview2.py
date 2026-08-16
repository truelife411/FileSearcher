# -*- coding: utf-8 -*-
"""主题配色候选 v2 设计图生成器（讨论用，不入库）"""
from PIL import Image, ImageDraw, ImageFont

W, H = 1560, 1860
IMG = Image.new("RGB", (W, H), "#222222")
D = ImageDraw.Draw(IMG)


def font(size, bold=False):
    paths = [(r"C:\Windows\Fonts\msyhbd.ttc" if bold else r"C:\Windows\Fonts\msyh.ttc"),
             r"C:\Windows\Fonts\simhei.ttf"]
    for p in paths:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default()


F_TITLE = font(30, True)
F_NAME = font(19, True)
F_DESC = font(13)
F_BODY = font(14)
F_SMALL = font(11)
F_MICRO = font(9)


def lerp(a, b, r):
    ca = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
    cb = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join("%02x" % int(ca[k] + (cb[k] - ca[k]) * r) for k in range(3))


def rnd(box, r, fill=None, outline=None, w=1):
    D.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=w)


CANDIDATES = [
    ("方案 11「银盐」· 黑白照片的相纸——银盐纸白、墨黑文字、灰阶层次，像一张冲洗出来的黑白照片", {
        "bg": "#E9E9E6", "surface": "#F2F2EF", "surface_alt": "#DDDDD8",
        "surface_3": "#D0D0CA", "input": "#F6F6F3", "border": "#C6C6C0",
        "border_strong": "#A6A69E", "text": "#21211E", "muted": "#5F5F58",
        "muted_2": "#85857C", "accent": "#22221F", "accent_hover": "#3A3A36",
        "accent_pressed": "#111110", "accent_grad_a": "#3A3A36", "accent_grad_b": "#141412",
        "selected": "#D9D9D3", "selected_hover": "#CDCDC6", "row_alt": "#EEEEEA",
        "row_line": "#DADAD4", "title_bg": "#DFDFD9", "hover": "#ECECE8",
        "sel_text": "#111110", "on_accent": "#F2F2EF",
        "badges": ["#6E6E66", "#4A4A44", "#8A8A80", "#5A5A52"],
    }),
    ("方案 12「底片」· 黑白照片的底片——暗房黑、银灰文字、灰阶反相，像胶片底片的质感", {
        "bg": "#1C1C1A", "surface": "#232321", "surface_alt": "#2A2A27",
        "surface_3": "#343430", "input": "#171715", "border": "#3A3A36",
        "border_strong": "#52524C", "text": "#D4D4CC", "muted": "#8F8F86",
        "muted_2": "#6B6B64", "accent": "#D6D6CE", "accent_hover": "#E2E2DA",
        "accent_pressed": "#B8B8AE", "accent_grad_a": "#D6D6CE", "accent_grad_b": "#A8A89E",
        "selected": "#3A3A34", "selected_hover": "#43433C", "row_alt": "#191917",
        "row_line": "#272724", "title_bg": "#1C1C1A", "hover": "#20201E",
        "sel_text": "#E2E2DA", "on_accent": "#1C1C1A",
        "badges": ["#8F8F86", "#6B6B64", "#B0B0A6", "#7A7A72"],
    }),
]

SW, SH = 640, 470  # 模拟窗口


def draw_window(x, y, t):
    rnd((x, y, x + SW, y + SH), 12, fill=t["bg"], outline=t["border"], w=1)
    # ---- 标题栏 ----
    D.rectangle((x, y + 20, x + SW, y + 40), fill=t["title_bg"])
    rnd((x, y, x + SW, y + 40), 12, fill=t["title_bg"])
    for i in range(24):
        D.line((x + 16 + i, y + 8, x + 16 + i, y + 30), fill=lerp(t["accent_grad_a"], t["accent_grad_b"], i / 23))
    rnd((x + 16, y + 8, x + 40, y + 32), 7, outline=t["border_strong"], w=1)
    D.ellipse((x + 23, y + 13, x + 33, y + 23), outline=t["on_accent"], width=2)
    D.line((x + 31, y + 22, x + 36, y + 27), fill=t["on_accent"], width=2)
    D.text((x + 50, y + 10), "FileSearcher", font=F_BODY, fill=t["text"])
    for i, bx in enumerate((x + SW - 96, x + SW - 76, x + SW - 56)):
        rnd((bx, y + 13, bx + 14, y + 27), 4, fill=t["border_strong"])
    # ---- 搜索框 ----
    rnd((x + 18, y + 52, x + SW - 18, y + 92), 8, fill=t["input"], outline=t["accent"], w=2)
    D.text((x + 34, y + 61), "搜索全盘文件…", font=F_SMALL, fill=t["muted_2"])
    D.ellipse((x + SW - 46, y + 62, x + SW - 30, y + 78), fill=t["muted_2"])
    D.line((x + SW - 42, y + 67, x + SW - 34, y + 75), fill=t["input"], width=2)
    D.line((x + SW - 34, y + 67, x + SW - 42, y + 75), fill=t["input"], width=2)
    # ---- chips 行 ----
    chips = [("全部", t["accent"], t["on_accent"]), ("文档", t["surface_alt"], t["muted"]),
             ("图片", t["surface_alt"], t["muted"]), ("视频", t["surface_alt"], t["muted"])]
    cx = x + 18
    for label, bg, fg in chips:
        w = 16 + len(label) * 14
        rnd((cx, y + 104, cx + w, y + 124), 10, fill=bg, outline=t["border"], w=1)
        D.text((cx + 8, y + 107), label, font=F_SMALL, fill=fg)
        cx += w + 10
    # ---- 表头 ----
    hy = y + 136
    rnd((x + 8, hy, x + SW - 8, hy + 24), 5, fill=t["surface_alt"])
    D.text((x + 20, hy + 4), "文件名", font=F_SMALL, fill=t["text"])
    D.polygon([(x + 64, hy + 8), (x + 70, hy + 14), (x + 58, hy + 14)], fill=t["accent"])
    D.text((x + SW - 110, hy + 4), "大小", font=F_SMALL, fill=t["muted"], anchor="rs")
    D.text((x + SW - 20, hy + 4), "修改时间", font=F_SMALL, fill=t["muted"], anchor="rs")
    # ---- 数据行 ----
    files = [("设计规范.pdf", "2.4 MB", True, t["badges"][0]),
             ("会议纪要.docx", "86 KB", False, t["badges"][1]),
             ("项目报告.pptx", "1.1 MB", False, t["badges"][2]),
             ("工作照片集", "3.2 GB", False, t["badges"][3])]
    for i, (name, size, sel, badge) in enumerate(files):
        ry = hy + 28 + i * 38
        if sel:
            rnd((x + 8, ry, x + SW - 8, ry + 36), 6, fill=t["selected"])
            D.rectangle((x + 8, ry + 4, x + 12, ry + 32), fill=t["accent"])
            name_fill = t["sel_text"]
        else:
            rnd((x + 8, ry, x + SW - 8, ry + 36), 6, fill=t["surface"])
            D.line((x + 18, ry + 35, x + SW - 18, ry + 35), fill=t["row_line"])
            name_fill = t["text"]
        D.ellipse((x + 22, ry + 12, x + 30, ry + 20), fill=badge)
        D.text((x + 38, ry + 8), name, font=F_BODY, fill=name_fill)
        D.text((x + SW - 110, ry + 9), size, font=F_SMALL, fill=t["muted"], anchor="rs")
    # ---- 分页栏 ----
    py = hy + 28 + 4 * 38 + 10
    D.line((x + 8, py, x + SW - 8, py), fill=t["row_line"])
    D.text((x + 18, py + 8), "共 12,847 个结果", font=F_MICRO, fill=t["muted"])
    rnd((x + SW - 118, py + 4, x + SW - 76, py + 26), 6, fill=t["surface_alt"], outline=t["border"], w=1)
    D.text((x + SW - 97, py + 7), "‹", font=F_SMALL, fill=t["muted"], anchor="ms")
    rnd((x + SW - 70, py + 4, x + SW - 18, py + 26), 6, fill=t["accent"])
    D.text((x + SW - 44, py + 7), "1/12", font=F_MICRO, fill=t["on_accent"], anchor="ms")
    # ---- Kindle 阅读示例 ----
    ry0 = py + 36
    D.text((x + 18, ry0), "阅读示例", font=F_MICRO, fill=t["muted_2"])
    for i, ln in enumerate(("纸上得来终觉浅，绝知此事要躬行。", "读书破万卷，下笔如有神。")):
        D.text((x + 18, ry0 + 16 + i * 20), ln, font=F_SMALL, fill=t["muted"])


def draw_swatches(x, y, t):
    labels = ["bg", "surface", "surface_alt", "surface_3", "input", "border",
              "border_strong", "text", "muted", "muted_2", "accent",
              "selected", "row_line", "sel_text"]
    for i, name in enumerate(labels):
        col, row = i % 2, i // 2
        bx = x + col * 200
        by = y + row * 40
        rnd((bx, by, bx + 84, by + 24), 5, fill=t[name], outline="#4A4A4A", w=1)
        D.text((bx + 92, by + 5), name, font=F_MICRO, fill="#B8B8B8")
        D.text((bx + 150, by + 5), t[name], font=F_MICRO, fill="#8A8A8A")


D.text((36, 26), "FileSearcher 主题配色 · 黑白照片方案（银盐相纸浅色 + 底片深色，纯灰阶无彩色）", font=F_TITLE, fill="#F2F2F2")
D.text((36, 70), "预览为全界面模拟：标题栏 / 搜索框 / 筛选 chips / 表格(选中行+类型圆点) / 分页栏；右侧为关键色板", font=F_DESC, fill="#9E9E9E")

y = 108
for name, t in CANDIDATES:
    D.text((36, y), name, font=F_NAME, fill="#F2F2F2")
    draw_window(36, y + 28, t)
    draw_swatches(36 + SW + 40, y + 60, t)
    y += 28 + SH + 52

IMG.save("ui_themes_v2.png")
print("saved ui_themes_v2.png")
