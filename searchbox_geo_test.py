# -*- coding: utf-8 -*-
"""临时几何测试：验证搜索框输入区与清空按钮是否重叠（模拟用户 240DPI + text_pt=18 的框高）。"""
import sys
import tkinter as tk

sys.path.insert(0, r"C:\Users\hjf\Documents\代码\FileSearcher\file-searcher")
from file_searcher import RoundedSearchBox, THEMES  # noqa: E402

root = tk.Tk()
root.geometry("900x200")
root.tk.call("tk", "scaling", 4 / 3)
var = tk.StringVar(value="测试搜索词 abcdefghij")
for h in (58, 137):  # 96dpi 基准高 / 用户 240DPI+text_pt18 的实际高
    box = RoundedSearchBox(root, var, THEMES["mist"], lambda: var.set(""), height=h,
                           font_size=13, placeholder="搜索全盘文件…")
    box.pack(fill=tk.X, padx=24)
    root.update()
    w = box.winfo_width()
    bb = box.bbox(box._entry_window)
    entry_right = bb[2]
    btn_left = w - 44 - box._clear_r()
    overlap = entry_right - btn_left
    print(f"height={h}: entry_right={entry_right}, btn_left={btn_left}, "
          f"overlap={overlap}px ({'遮住!' if overlap > 0 else 'OK'})")
    box.destroy()
root.destroy()
