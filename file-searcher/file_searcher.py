#!/usr/bin/env python3
"""File Searcher — 全盘文件快速搜索工具，基于本地索引。"""

import os
import sys
import math
import subprocess
import sqlite3
import json
import shutil
import ctypes
import threading
import queue
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

# Windows DPI 感知
if sys.platform == "win32":
    try:
        ctypes.windll.shcore.SetProcessDPIAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

import tkinter as tk
from tkinter import ttk
from pathlib import Path

import pystray
from PIL import Image, ImageDraw, ImageTk


# ================================================================
#  全局常量配置
# ================================================================
INDEX_DIR = Path.home() / ".file_searcher_index"
INDEX_DB = INDEX_DIR / "index.db"
DEBUG_LOG = INDEX_DIR / "debug.log"
DEBUG_LOG_MAX = 1024 * 1024        # debug.log 超过 1MB 轮转为 debug.log.1
INDEX_SCAN_WORKERS = 4
INDEX_WRITE_BATCH_SIZE = 5000
INDEX_QUEUE_SIZE = 20000
PAGE_SIZE = 5000

FONT_FAMILY = "Microsoft YaHei UI" if sys.platform == "win32" else "sans-serif"
FONT_MONO = "Consolas" if sys.platform == "win32" else "monospace"
# 全局字号体系（基础 pt，最终渲染 pt = base × dpi_scale × font_scale，dpi_scale 为折中后的 DPI 系数）
# 240 DPI（250% 缩放）主字号 18 时正文 = 12 × 1.581 × 1.5 ≈ 28.5pt ≈ 38px 物理，接近系统应用观感
# MICRO≈19pt 表头·胶囊·徽章·状态栏 / SMALL≈21pt 辅助·路径·说明 / BODY≈25pt 正文·文件名·按钮
# INPUT≈27pt 搜索框（主角稍大）/ LG≈29pt 弹窗标题·空状态 / XL≈31pt 设置页大标题
FONT_MICRO = 9
FONT_SMALL = 10     # 辅助文字：路径列、说明、副文案
FONT_BODY = 12      # 正文：文件名、按钮、菜单项、状态文字
FONT_INPUT = 13     # 搜索框输入（视觉主角，比正文大一号）
FONT_LG = 14        # 弹窗标题、空状态主文案
FONT_XL = 15        # 设置页大标题
FONT_HEADER = FONT_MICRO   # 表头
FONT_TITLE = FONT_SMALL    # 标题栏标题
ROW_HEIGHT = 50     # 结果行高
TITLEBAR_H = 40     # 自绘标题栏高度
SPACING = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24}
THEMES = {
    # 深色「墨玉」：深空玻璃拟态（Raycast 风）。surface 系为"白上浮"层次，
    # border 系为"白描边"，选中态为 accent 半透明
    "dark": {
        "bg": "#0A0C10", "surface": "#14181F", "surface_alt": "#1A1F28",
        "surface_3": "#232935", "input": "#10141B", "border": "#232A35",
        "border_strong": "#333D4B", "text": "#EDF2F4", "muted": "#7A8792",
        "muted_2": "#5D6B77", "accent": "#6FD8C8", "accent_hover": "#8AE4D6",
        "accent_pressed": "#4FC4B2", "accent_grad_a": "#3EDAC6", "accent_grad_b": "#1FA190",
        "selected": "#16302D", "selected_hover": "#1B3B37", "row_alt": "#0E1116",
        "row_line": "#171C24", "title_bg": "#0A0C10", "hover": "#161B23",
        "sel_text": "#8AE4D6",
        "success": "#3ED598", "warning": "#F0B44C", "error": "#E4747E",
        "menu_bg": "#151920", "dialog_bg": "#151920",
    },
    # 浅色「凝脂」：宣纸暖白 + 黛青 accent（纸感国风，直角小圆角体系）
    "light": {
        "bg": "#F5F3EE", "surface": "#FFFFFF", "surface_alt": "#F7F5EF",
        "surface_3": "#EDEAE3", "input": "#FFFFFF", "border": "#E0DCD2",
        "border_strong": "#C9C3B4", "text": "#23282C", "muted": "#6B6558",
        "muted_2": "#8B8474", "accent": "#2E6E66", "accent_hover": "#3A837A",
        "accent_pressed": "#255A53", "accent_grad_a": "#3A837A", "accent_grad_b": "#255A53",
        "selected": "#EBF1EE", "selected_hover": "#DFEAE5", "row_alt": "#FAF8F3",
        "row_line": "#EFEDE6", "title_bg": "#EDEAE3", "hover": "#F7F5EF",
        "sel_text": "#255A53",
        "success": "#2E8B57", "warning": "#C08A2E", "error": "#B05353",
        "menu_bg": "#FFFFFF", "dialog_bg": "#FFFFFF",
    },
    # 浅色「雾沙」：柔和暖灰米白（低对比护眼，替代纯白扎眼）
    "mist": {
        "bg": "#ECE9E3", "surface": "#F5F2EC", "surface_alt": "#E5E1D9",
        "surface_3": "#DBD6CC", "input": "#F8F5F0", "border": "#D4CFC4",
        "border_strong": "#BDB7A9", "text": "#3E3C37", "muted": "#6E6A60",
        "muted_2": "#8F8A7E", "accent": "#5F7D72", "accent_hover": "#6E8C81",
        "accent_pressed": "#526C62", "accent_grad_a": "#6E8C81", "accent_grad_b": "#4C6B60",
        "selected": "#E2E6DF", "selected_hover": "#D8DED5", "row_alt": "#F1EEE7",
        "row_line": "#E4E0D7", "title_bg": "#E4E0D8", "hover": "#EFECE4",
        "sel_text": "#4C6B60",
        "success": "#5E8C6E", "warning": "#B08A45", "error": "#B0625C",
        "menu_bg": "#F5F2EC", "dialog_bg": "#F5F2EC",
    },
    # 深色「青墨」：柔和绿灰（护眼深色，替代纯黑压抑）
    "jade": {
        "bg": "#1C2320", "surface": "#242C28", "surface_alt": "#2A342F",
        "surface_3": "#333F39", "input": "#19201D", "border": "#36423B",
        "border_strong": "#4A5A51", "text": "#C8D1CB", "muted": "#87948C",
        "muted_2": "#66736B", "accent": "#7FBFA8", "accent_hover": "#93CBB6",
        "accent_pressed": "#6BAB95", "accent_grad_a": "#7FBFA8", "accent_grad_b": "#55967F",
        "selected": "#2C3D35", "selected_hover": "#32463D", "row_alt": "#1A211E",
        "row_line": "#28312C", "title_bg": "#1C2320", "hover": "#222A26",
        "sel_text": "#9CD4BF",
        "success": "#6FBF8F", "warning": "#D3A24E", "error": "#CE7878",
        "menu_bg": "#242C28", "dialog_bg": "#242C28",
    },
}

# 文件类型徽标色点 (fg, bg)，按主题区分；bg 供弹窗图标圆底等复用
BADGE_STYLES = {
    "dark": {
        "doc": ("#6FA8EF", "#1B2634"), "pdf": ("#E4747E", "#2C2027"),
        "xls": ("#7CC784", "#1D2C2A"), "ppt": ("#E8A569", "#2D2725"),
        "img": ("#BE8FE0", "#252536"), "code": ("#5AC8C8", "#182D32"),
        "zip": ("#D1A26B", "#2A2725"), "audio": ("#7FD4A8", "#1D2E2E"),
        "video": ("#8FA5E8", "#20243C"), "dir": ("#E5C56F", "#2C2C26"),
        "file": ("#7A8792", "#1F2733"),
    },
    "light": {
        "doc": ("#4A7FBC", "#EDF2F8"), "pdf": ("#B05353", "#F8EEEC"),
        "xls": ("#5B8A5B", "#EDF3EC"), "ppt": ("#C97A2B", "#F8F0E5"),
        "img": ("#8E5BB8", "#F3ECF8"), "code": ("#3A8A8A", "#E9F3F3"),
        "zip": ("#A07B46", "#F5F0E6"), "audio": ("#3E9468", "#EAF4EF"),
        "video": ("#5B6FC4", "#ECEDF7"), "dir": ("#A08A46", "#F5F1E6"),
        "file": ("#6B6558", "#EFEDE6"),
    },
    "mist": {
        "doc": ("#5A7FA8", "#E8EEF4"), "pdf": ("#A85A5A", "#F4EAE8"),
        "xls": ("#5E8A62", "#E9F1E8"), "ppt": ("#B0783C", "#F3ECE2"),
        "img": ("#8A62A8", "#EFEAF4"), "code": ("#4A8585", "#E6F0F0"),
        "zip": ("#95784E", "#F1ECE3"), "audio": ("#4A8A6E", "#E7F1EB"),
        "video": ("#5E6FB0", "#E9EBF4"), "dir": ("#937E44", "#F1EDE2"),
        "file": ("#6E6A60", "#EBE8E0"),
    },
    "jade": {
        "doc": ("#7FA8D8", "#1E2A38"), "pdf": ("#D88A8A", "#332527"),
        "xls": ("#8EC28E", "#1F2E23"), "ppt": ("#D8A56E", "#332A22"),
        "img": ("#A98ED0", "#272438"), "code": ("#6EC0C0", "#1B2E30"),
        "zip": ("#C0A070", "#2E2922"), "audio": ("#8EC8A8", "#1E2E28"),
        "video": ("#9AA8D8", "#232638"), "dir": ("#D0B46E", "#2E2C22"),
        "file": ("#87948C", "#26302B"),
    },
}

# 扩展名 → 徽章类别（构建自上方扩展名集合，见下方 _build_badge_kind_map）
DOCUMENT_EXTENSIONS = {".doc", ".docx", ".pdf", ".txt", ".rtf", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods", ".csv", ".md"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff", ".ico", ".svg"}
VIDEO_EXTENSIONS = {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm", ".m4v"}
AUDIO_EXTENSIONS = {".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a"}
ARCHIVE_EXTENSIONS = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".cab"}
CODE_EXTENSIONS = {".py", ".js", ".ts", ".tsx", ".jsx", ".java", ".c", ".cpp", ".h", ".hpp", ".cs", ".go", ".rs", ".php", ".rb", ".swift", ".kt", ".html", ".css", ".scss", ".json", ".xml", ".yaml", ".yml", ".sql", ".sh", ".ps1"}


def _build_badge_kind_map() -> dict:
    """扩展名 → 徽章类别映射表。"""
    m = {}
    for ext in DOCUMENT_EXTENSIONS:
        if ext in (".xls", ".xlsx", ".ods", ".csv"):
            m[ext] = "xls"
        elif ext in (".ppt", ".pptx"):
            m[ext] = "ppt"
        elif ext == ".pdf":
            m[ext] = "pdf"
        else:
            m[ext] = "doc"
    for ext in IMAGE_EXTENSIONS:
        m[ext] = "img"
    for ext in VIDEO_EXTENSIONS:
        m[ext] = "video"
    for ext in AUDIO_EXTENSIONS:
        m[ext] = "audio"
    for ext in ARCHIVE_EXTENSIONS:
        m[ext] = "zip"
    for ext in CODE_EXTENSIONS:
        m[ext] = "code"
    return m


BADGE_KIND_MAP = _build_badge_kind_map()


def badge_kind_for(name: str, is_dir: bool) -> str:
    """根据文件名与是否目录返回徽章类别。"""
    if is_dir:
        return "dir"
    ext = os.path.splitext(name)[1].lower()
    return BADGE_KIND_MAP.get(ext, "file")

IGNORE_DIRS = {
    "windows", "winnt", "system32", "syswow64", "winsxs",
    "$recycle.bin", "recycler", "recycled",
    "system volume information", "boot", "recovery",
    "config.msi", "msocache",
    "appdata", "application data", "local settings",
    "cookies", "history", "temporary internet files",
    "microsoft.net", "assembly", "installer",
    "node_modules", "__pycache__", ".git", ".svn", ".hg", ".file_searcher_index",
    "package cache", "driverstore", "servicing",
    # Linux 虚拟文件系统
    "proc", "sys", "dev", "run", "lost+found", "snap",
}

IGNORE_EXTENSIONS = {
    ".tmp", ".temp", ".log", ".etl", ".dmp", ".cache",
    ".pyc", ".pyo", ".pyd", ".class",
    ".ilk", ".pdb", ".obj", ".exp", ".lib",
}

IGNORE_PATH_CONTAINS = [
    "\\windows\\", "\\$recycle.bin\\", "\\system volume information\\",
    "\\appdata\\local\\temp\\", "\\windows\\temp\\",
    "\\temp\\", "\\cache\\", "\\microsoft\\windows\\",
    "\\driverstore\\", "\\servicing\\",
    "\\users\\all users\\", "\\users\\default user\\",
    "\\documents and settings\\",
    "\\programdata\\application data\\",
    "\\programdata\\desktop\\", "\\programdata\\documents\\",
    "\\programdata\\start menu\\", "\\programdata\\templates\\",
    "\\appdata\\local\\application data\\",
    "\\appdata\\local\\history\\",
    "\\appdata\\local\\temporary internet files\\",
    "\\application data\\", "\\local settings\\",
]


FILE_ATTR_SKIP = 0
try:
    import stat as _st
    FILE_ATTR_SKIP = _st.FILE_ATTRIBUTE_HIDDEN | _st.FILE_ATTRIBUTE_SYSTEM | _st.FILE_ATTRIBUTE_REPARSE_POINT
except Exception:
    import stat as _st  # 非 Windows：FILE_ATTRIBUTE_* 不存在，但仍需要 S_ISDIR/S_ISREG


# ================================================================
#  Windows Shell API 结构体 — 用于回收站操作
# ================================================================
class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", ctypes.c_void_p),
        ("wFunc", ctypes.c_uint),
        ("pFrom", ctypes.c_wchar_p),
        ("pTo", ctypes.c_wchar_p),
        ("fFlags", ctypes.c_ushort),
        ("fAnyOperationsAborted", ctypes.c_int),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", ctypes.c_wchar_p),
    ]

FO_DELETE = 3
FOF_ALLOWUNDO = 0x40
FOF_NOCONFIRMATION = 0x10
FOF_SILENT = 0x0004
FOF_WANTNUKEWARNING = 0x4000


class WindowsShellIconCache:
    """按扩展名缓存 Windows Shell 小图标，失败时返回统一占位图标。"""

    def __init__(self, root, background: str, size: int = 22):
        self.root = root
        self.background = background
        self.size = size
        self._cache = {}
        self._fallback = self._make_fallback()

    def _make_fallback(self):
        image = Image.new("RGBA", (self.size, self.size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((4, 2, self.size - 4, self.size - 2), radius=2,
                               outline="#87919D", fill="#5F6C79")
        draw.polygon((self.size - 8, 2, self.size - 4, 6, self.size - 8, 6), fill="#A5ADB6")
        return ImageTk.PhotoImage(image, master=self.root)

    def get(self, path: str, is_dir: bool = False):
        key = "__folder__" if is_dir else (Path(path).suffix.lower() or "__file__")
        if key not in self._cache:
            self._cache[key] = self._load_shell_icon(key, is_dir) or self._fallback
        return self._cache[key]

    def _load_shell_icon(self, key: str, is_dir: bool):
        if sys.platform != "win32":
            return None
        info = None
        shell32 = user32 = gdi32 = None
        screen_dc = mem_dc = bitmap = old_bitmap = brush = None
        try:
            from ctypes import wintypes

            class SHFILEINFOW(ctypes.Structure):
                _fields_ = [
                    ("hIcon", wintypes.HICON), ("iIcon", ctypes.c_int),
                    ("dwAttributes", wintypes.DWORD), ("szDisplayName", wintypes.WCHAR * 260),
                    ("szTypeName", wintypes.WCHAR * 80),
                ]

            class BITMAPINFOHEADER(ctypes.Structure):
                _fields_ = [
                    ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                    ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                    ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                    ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                    ("biClrImportant", wintypes.DWORD),
                ]

            shell32 = ctypes.windll.shell32
            user32 = ctypes.windll.user32
            gdi32 = ctypes.windll.gdi32
            shell32.SHGetFileInfoW.argtypes = [
                wintypes.LPCWSTR, wintypes.DWORD, ctypes.POINTER(SHFILEINFOW),
                wintypes.UINT, wintypes.UINT,
            ]
            shell32.SHGetFileInfoW.restype = ctypes.c_size_t
            user32.GetDC.argtypes = [wintypes.HWND]
            user32.GetDC.restype = wintypes.HDC
            user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
            user32.ReleaseDC.restype = ctypes.c_int
            user32.FillRect.argtypes = [wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.HBRUSH]
            user32.FillRect.restype = ctypes.c_int
            user32.DrawIconEx.argtypes = [
                wintypes.HDC, ctypes.c_int, ctypes.c_int, wintypes.HICON,
                ctypes.c_int, ctypes.c_int, wintypes.UINT, wintypes.HBRUSH, wintypes.UINT,
            ]
            user32.DrawIconEx.restype = wintypes.BOOL
            user32.DestroyIcon.argtypes = [wintypes.HICON]
            user32.DestroyIcon.restype = wintypes.BOOL
            gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
            gdi32.CreateCompatibleDC.restype = wintypes.HDC
            gdi32.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
            gdi32.CreateCompatibleBitmap.restype = wintypes.HBITMAP
            gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
            gdi32.SelectObject.restype = wintypes.HGDIOBJ
            gdi32.CreateSolidBrush.argtypes = [wintypes.COLORREF]
            gdi32.CreateSolidBrush.restype = wintypes.HBRUSH
            gdi32.GetDIBits.argtypes = [
                wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT,
                ctypes.c_void_p, ctypes.POINTER(BITMAPINFOHEADER), wintypes.UINT,
            ]
            gdi32.GetDIBits.restype = ctypes.c_int
            gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
            gdi32.DeleteObject.restype = wintypes.BOOL
            gdi32.DeleteDC.argtypes = [wintypes.HDC]
            gdi32.DeleteDC.restype = wintypes.BOOL

            info = SHFILEINFOW()
            sample = "folder" if is_dir else ("file" + key if key.startswith(".") else "file")
            attributes = 0x10 if is_dir else 0x80
            flags = 0x000000100 | 0x000000001 | 0x000000010
            if not shell32.SHGetFileInfoW(sample, attributes, ctypes.byref(info), ctypes.sizeof(info), flags):
                return None

            width = height = self.size
            screen_dc = user32.GetDC(None)
            if not screen_dc:
                return None
            mem_dc = gdi32.CreateCompatibleDC(screen_dc)
            if not mem_dc:
                return None
            bitmap = gdi32.CreateCompatibleBitmap(screen_dc, width, height)
            if not bitmap:
                return None
            old_bitmap = gdi32.SelectObject(mem_dc, bitmap)
            if not old_bitmap or old_bitmap == ctypes.c_void_p(-1).value:
                old_bitmap = None
                return None
            bg = self.background.lstrip("#")
            colorref = int(bg[4:6] + bg[2:4] + bg[0:2], 16)
            brush = gdi32.CreateSolidBrush(colorref)
            if not brush:
                return None
            rect = wintypes.RECT(0, 0, width, height)
            user32.FillRect(mem_dc, ctypes.byref(rect), brush)
            user32.DrawIconEx(mem_dc, 0, 0, info.hIcon, width, height, 0, None, 0x0003)

            header = BITMAPINFOHEADER()
            header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            header.biWidth = width
            header.biHeight = -height
            header.biPlanes = 1
            header.biBitCount = 32
            buffer = ctypes.create_string_buffer(width * height * 4)
            gdi32.SelectObject(mem_dc, old_bitmap)
            old_bitmap = None
            if not gdi32.GetDIBits(mem_dc, bitmap, 0, height, buffer, ctypes.byref(header), 0):
                return None
            image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1).copy()
            bg_rgb = tuple(int(bg[i:i + 2], 16) for i in (0, 2, 4))
            alpha = Image.new("L", image.size)
            alpha.putdata([
                0 if max(abs(r - bg_rgb[0]), abs(g - bg_rgb[1]), abs(b - bg_rgb[2])) <= 2 else 255
                for r, g, b, _a in image.getdata()
            ])
            image.putalpha(alpha)
            return ImageTk.PhotoImage(image, master=self.root)
        except Exception:
            return None
        finally:
            try:
                if gdi32 is not None and old_bitmap and mem_dc:
                    gdi32.SelectObject(mem_dc, old_bitmap)
                if gdi32 is not None and brush:
                    gdi32.DeleteObject(brush)
                if gdi32 is not None and bitmap:
                    gdi32.DeleteObject(bitmap)
                if gdi32 is not None and mem_dc:
                    gdi32.DeleteDC(mem_dc)
                if user32 is not None and screen_dc:
                    user32.ReleaseDC(None, screen_dc)
                if user32 is not None and info is not None and info.hIcon:
                    user32.DestroyIcon(info.hIcon)
            except Exception:
                pass


# ================================================================
#  工具函数
# ================================================================

def format_size(size: int) -> str:
    """将字节数格式化为人类可读的大小字符串"""
    if size < 1024:
        return f"{size} B"
    elif size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    elif size < 1024 * 1024 * 1024:
        return f"{size / (1024 * 1024):.1f} MB"
    elif size < 1024 * 1024 * 1024 * 1024:
        return f"{size / (1024 * 1024 * 1024):.1f} GB"
    else:
        return f"{size / (1024 * 1024 * 1024 * 1024):.1f} TB"


def open_with_default(path: str):
    """用系统默认软件打开文件（失败时抛出 OSError，由调用方提示）。

    必须在主线程调用 os.startfile：它底层走 ShellExecute，依赖调用线程的
    COM 初始化与消息循环。后台守护线程没有这些，对某些文件关联（DDE/shell
    extension）会出现「程序窗口起来了但文件没传过去」的空窗口。os.startfile
    本身是异步的（发消息即返回），不会阻塞 UI。文档中心在主线程调用故正常。
    """
    try:
        if sys.platform == "win32":
            os.startfile(os.path.normpath(path))
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError as e:
        try:
            _diag(f"[open-error] {path} -> {e}")
        except Exception:
            pass
        raise


def open_file_location(path: str):
    """在资源管理器中定位并选中文件"""
    if sys.platform == "win32":
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    else:
        subprocess.Popen(["xdg-open", os.path.dirname(path)])


def _find_trash_cmd() -> str | None:
    """查找可用的回收站命令。"""
    for cmd in ["gio", "trash-put", "kioclient5", "kioclient"]:
        if shutil.which(cmd):
            return cmd
    return None


def send_to_recycle_bin(paths: list[str]):
    """将文件列表移入回收站。"""
    if not paths:
        return
    if sys.platform == "win32":
        file_list = "\0".join(paths) + "\0\0"
        buf = ctypes.create_unicode_buffer(file_list)
        op = SHFILEOPSTRUCTW()
        op.wFunc = FO_DELETE
        op.pFrom = ctypes.cast(buf, ctypes.c_wchar_p)
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
        ret = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        if ret != 0:
            raise OSError(f"无法删除文件: {paths}")
    else:
        trash_cmd = _find_trash_cmd()
        if trash_cmd is None:
            raise OSError("未找到回收站命令，请安装 trash-cli (sudo apt install trash-cli)")
        for p in paths:
            try:
                if trash_cmd == "gio":
                    subprocess.run(["gio", "trash", p], check=True)
                else:
                    subprocess.run([trash_cmd, p], check=True)
            except subprocess.CalledProcessError as e:
                raise OSError(f"无法删除文件: {p}") from e


def permanent_delete(paths: list[str]):
    """彻底删除文件列表，不可恢复。"""
    if not paths:
        return
    if sys.platform == "win32":
        file_list = "\0".join(paths) + "\0\0"
        buf = ctypes.create_unicode_buffer(file_list)
        op = SHFILEOPSTRUCTW()
        op.wFunc = FO_DELETE
        op.pFrom = ctypes.cast(buf, ctypes.c_wchar_p)
        op.fFlags = FOF_NOCONFIRMATION | FOF_SILENT | FOF_WANTNUKEWARNING
        ret = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        if ret != 0:
            raise OSError(f"无法删除文件: {paths}")
    else:
        for p in paths:
            try:
                if os.path.isdir(p):
                    shutil.rmtree(p)
                else:
                    os.remove(p)
            except OSError as e:
                raise OSError(f"无法删除文件: {p}") from e


def rename_file(old_path: str, new_name: str) -> str:
    """重命名文件，如果目标已存在则抛出异常。返回新路径"""
    parent = os.path.dirname(old_path)
    new_path = os.path.join(parent, new_name)
    if os.path.exists(new_path):
        raise FileExistsError(f"目标已存在: {new_path}")
    os.rename(old_path, new_path)
    return new_path


# ================================================================
#  IndexEngine — 索引引擎
# ================================================================
class IndexEngine:
    """全盘文件索引引擎，使用 SQLite 存储文件元数据。"""

    _build_lock = threading.Lock()
    EXCLUDE_FILE = INDEX_DIR / "exclude.json"
    LAYOUT_FILE = INDEX_DIR / "layout.json"
    SETTINGS_FILE = INDEX_DIR / "settings.json"

    def __init__(self, progress_callback=None, cancel_check=None):
        self._progress = progress_callback or (lambda msg, n: None)
        self._cancel = cancel_check or (lambda: False)
        self._exclude_dirs, self._exclude_paths = self._load_exclude_list()

    # ---- 索引构建 ----

    def build_index(self):
        """并行扫描所有盘符，由当前线程统一批量写入 SQLite。"""
        if not self._build_lock.acquire(blocking=False):
            raise RuntimeError("索引任务已在运行")
        try:
            return self._build_index_locked()
        finally:
            self._build_lock.release()

    def _build_index_locked(self):
        """执行单个全量索引任务，并返回各阶段耗时。

        扫描采用「目录任务队列」：worker 一次领取一个目录扫描，子目录入共享队列，
        多 worker 在单盘/多盘下都能并行（不再按盘符分配导致单盘并行度 1）。
        modified 列存 epoch 秒（INTEGER），UI 层再格式化。
        """
        total_started = time.perf_counter()
        INDEX_DIR.mkdir(parents=True, exist_ok=True)

        conn = sqlite3.connect(str(INDEX_DB))
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("BEGIN IMMEDIATE")
        conn.execute("DROP TABLE IF EXISTS files")
        conn.execute("""
            CREATE TABLE files (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                name_lower TEXT NOT NULL,
                path TEXT NOT NULL UNIQUE,
                path_lower TEXT NOT NULL,
                size INTEGER NOT NULL,
                modified INTEGER NOT NULL,
                is_dir INTEGER NOT NULL DEFAULT 0
            )
        """)
        # FTS5 trigram 全文索引：子串匹配走倒排索引（搜索词 ≥3 字符时启用）
        fts_ok = False
        try:
            conn.execute("DROP TABLE IF EXISTS files_fts")
            conn.execute(
                "CREATE VIRTUAL TABLE files_fts USING fts5(name, path, tokenize='trigram')")
            fts_ok = True
        except Exception:
            fts_ok = False

        drives = [f"{d}:\\" for d in "CDEFGHIJKLMNOPQRSTUVWXYZAB" if os.path.exists(f"{d}:\\")]
        if not drives:
            conn.commit()
            conn.close()
            elapsed = time.perf_counter() - total_started
            return {
                "total_files": 0,
                "scan_write_seconds": elapsed,
                "optimize_seconds": 0.0,
                "total_seconds": elapsed,
            }

        result_queue = queue.Queue(maxsize=INDEX_QUEUE_SIZE)
        dir_queue = queue.Queue()
        stop_event = threading.Event()
        scan_cond = threading.Condition()
        pending_dirs = [len(drives)]   # 队列中待处理目录数（可变容器，worker/扫描共享）
        active_workers = 0             # 正在扫描目录的 worker 数
        scan_done = False              # 全部扫描结束
        worker_count = INDEX_SCAN_WORKERS
        total_files = 0
        insert_batch = []

        def put_queue_item(item):
            while not stop_event.is_set():
                try:
                    result_queue.put(item, timeout=0.2)
                    return True
                except queue.Full:
                    continue
            return False

        def scan_worker():
            nonlocal active_workers, scan_done
            while True:
                with scan_cond:
                    if self._cancel():
                        scan_done = True
                        scan_cond.notify_all()
                        return
                    while pending_dirs[0] == 0 and active_workers == 0 and not scan_done:
                        scan_done = True
                        scan_cond.notify_all()
                        return
                    if scan_done:
                        return
                    try:
                        dirpath = dir_queue.get_nowait()
                    except queue.Empty:
                        scan_cond.wait(timeout=0.3)
                        continue
                    pending_dirs[0] -= 1
                    active_workers += 1
                try:
                    self._scan_one_dir(dirpath, put_queue_item, dir_queue, scan_cond,
                                       pending_dirs)
                finally:
                    with scan_cond:
                        active_workers -= 1
                        scan_cond.notify_all()

        try:
            for drive in drives:
                dir_queue.put(drive)
            with ThreadPoolExecutor(max_workers=worker_count,
                                    thread_name_prefix="index-scan") as executor:
                futures = [executor.submit(scan_worker) for _ in range(worker_count)]
                while True:
                    try:
                        item = result_queue.get(timeout=0.2)
                    except queue.Empty:
                        if scan_done and result_queue.empty():
                            break
                        continue
                    insert_batch.append(item)
                    total_files += 1
                    if len(insert_batch) >= INDEX_WRITE_BATCH_SIZE:
                        conn.executemany(
                            "INSERT OR IGNORE INTO files(name, name_lower, path, path_lower, size, modified, is_dir) "
                            "VALUES(?,?,?,?,?,?,?)",
                            insert_batch,
                        )
                        self._progress(f"已收录 {total_files} 个文件", total_files)
                        insert_batch.clear()
                for future in futures:
                    future.result()

            if self._cancel():
                conn.rollback()
                return None

            if insert_batch:
                conn.executemany(
                    "INSERT OR IGNORE INTO files(name, name_lower, path, path_lower, size, modified, is_dir) "
                    "VALUES(?,?,?,?,?,?,?)",
                    insert_batch,
                )

            scan_write_seconds = time.perf_counter() - total_started
            self._progress("正在优化索引…", total_files)
            optimize_started = time.perf_counter()
            conn.execute("CREATE INDEX IF NOT EXISTS idx_name_lower ON files(name_lower)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_path_lower ON files(path_lower)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_size ON files(size)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_modified ON files(modified)")
            if fts_ok:
                try:
                    conn.execute(
                        "INSERT INTO files_fts(rowid, name, path) SELECT id, name, path FROM files")
                except Exception:
                    pass
            conn.commit()
            optimize_seconds = time.perf_counter() - optimize_started
            return {
                "total_files": total_files,
                "scan_write_seconds": scan_write_seconds,
                "optimize_seconds": optimize_seconds,
                "total_seconds": time.perf_counter() - total_started,
            }
        except Exception:
            stop_event.set()
            conn.rollback()
            raise
        finally:
            stop_event.set()
            conn.close()

    def _scan_one_dir(self, dirpath, put_queue_item, dir_queue, scan_cond, pending_dirs):
        """扫描单个目录：文件入结果队列，子目录入共享任务队列（含计数更新）。"""
        try:
            entries = os.scandir(dirpath)
        except (PermissionError, OSError):
            return
        with entries:
            for entry in entries:
                if self._cancel():
                    return
                full = entry.path
                try:
                    st = entry.stat(follow_symlinks=False)
                except OSError:
                    continue
                if _st.S_ISDIR(st.st_mode):
                    if not self._should_skip_dir(full, entry, st):
                        item = (
                            entry.name, entry.name.lower(), full, full.lower(), 0,
                            int(st.st_mtime), 1,
                        )
                        if not put_queue_item(item):
                            return
                        with scan_cond:
                            pending_dirs[0] += 1
                            dir_queue.put(full)
                            scan_cond.notify()
                    continue
                if not _st.S_ISREG(st.st_mode) or not self._should_include_file(full, entry):
                    continue
                item = (
                    entry.name,
                    entry.name.lower(),
                    full,
                    full.lower(),
                    st.st_size,
                    int(st.st_mtime),
                    0,
                )
                if not put_queue_item(item):
                    return

    def _should_skip_dir(self, dirpath: str, entry, st=None) -> bool:
        """判断目录是否应被跳过（系统目录、junction 点、排除列表等）。"""
        basename = entry.name.lower()
        try:
            if entry.is_junction():
                return True
        except Exception:
            pass
        if FILE_ATTR_SKIP:
            try:
                attrs = st.st_file_attributes
                if attrs & FILE_ATTR_SKIP:
                    return True
            except (AttributeError, OSError):
                pass
        if basename in IGNORE_DIRS:
            return True
        lower_path = dirpath.lower()
        for pat in IGNORE_PATH_CONTAINS:
            if pat in lower_path:
                return True
        if basename in self._exclude_dirs:
            return True
        for pat in self._exclude_paths:
            if pat in lower_path:
                return True
        return False

    def _should_include_file(self, filepath: str, entry) -> bool:
        """判断文件是否应被索引（排除指定扩展名和路径模式）。"""
        ext = os.path.splitext(entry.name)[1].lower()
        if ext in IGNORE_EXTENSIONS:
            return False
        lower_path = filepath.lower()
        for pat in self._exclude_paths:
            if pat in lower_path:
                return False
        return True

    @staticmethod
    def index_exists() -> bool:
        """检查索引数据库是否存在。"""
        return INDEX_DB.exists()

    @staticmethod
    def ensure_indexes():
        """幂等补齐排序列索引 + FTS5 表（老库升级用，重建时已含）。

        老库 FTS 表只建表不回填（回填交给后台 ensure_fts_backfill），
        回填完成前搜索走原 LIKE 路径，避免升级时卡启动。
        """
        if not INDEX_DB.exists():
            return
        try:
            conn = sqlite3.connect(str(INDEX_DB))
            conn.execute("CREATE INDEX IF NOT EXISTS idx_size ON files(size)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_modified ON files(modified)")
            try:
                conn.execute(
                    "CREATE VIRTUAL TABLE IF NOT EXISTS files_fts "
                    "USING fts5(name, path, tokenize='trigram')")
            except Exception:
                pass
            conn.commit()
            conn.close()
        except Exception:
            pass

    @staticmethod
    def ensure_fts_backfill():
        """后台回填 FTS 表缺失行（老库升级用）。表不存在或已完整时静默返回。

        索引任务（_build_lock）进行中时跳过本次回填——重建完成后 FTS 已随建随填。
        """
        if not INDEX_DB.exists() or not IndexEngine._build_lock.acquire(blocking=False):
            return
        try:
            conn = sqlite3.connect(str(INDEX_DB))
            try:
                conn.execute(
                    "CREATE VIRTUAL TABLE IF NOT EXISTS files_fts "
                    "USING fts5(name, path, tokenize='trigram')")
                have = conn.execute("SELECT count(*) FROM files_fts").fetchone()[0]
                total = conn.execute("SELECT count(*) FROM files").fetchone()[0]
                if have < total:
                    conn.execute(
                        "INSERT INTO files_fts(rowid, name, path) "
                        "SELECT id, name, path FROM files "
                        "WHERE id NOT IN (SELECT rowid FROM files_fts)")
                    conn.commit()
            except Exception:
                pass
            finally:
                conn.close()
        except Exception:
            pass
        finally:
            IndexEngine._build_lock.release()

    _schema_cache = (0.0, False, False)  # (db_mtime, has_is_dir, fts_ok)

    @classmethod
    def _db_schema(cls):
        """返回库结构信息（带 db_mtime 缓存）：(has_is_dir, fts_ok)。

        fts_ok = FTS 表存在且已有数据（空表视为未回填，走 LIKE 路径，
        回填完成后 mtime 变化自动切换）。
        """
        try:
            mtime = INDEX_DB.stat().st_mtime if INDEX_DB.exists() else 0.0
        except Exception:
            mtime = 0.0
        if cls._schema_cache[0] == mtime:
            return cls._schema_cache[1], cls._schema_cache[2]
        has_is_dir, fts_ok = False, False
        if INDEX_DB.exists():
            try:
                conn = sqlite3.connect(str(INDEX_DB))
                try:
                    cols = {row[1]: row[2]
                            for row in conn.execute("PRAGMA table_info(files)")}
                    has_is_dir = "is_dir" in cols
                    fts_ok = bool(conn.execute(
                        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='files_fts'"
                    ).fetchone())
                    if fts_ok:
                        n = conn.execute("SELECT count(*) FROM files_fts").fetchone()[0]
                        fts_ok = n > 0
                finally:
                    conn.close()
            except Exception:
                pass
        cls._schema_cache = (mtime, has_is_dir, fts_ok)
        return has_is_dir, fts_ok

    @staticmethod
    def _fts_query_ids(conn, q: str):
        """FTS5 trigram 子串匹配，返回匹配的 files.id 列表；失败返回 None（走 LIKE）。"""
        phrase = '"' + q.replace('"', " ") + '"'
        try:
            return [r[0] for r in conn.execute(
                "SELECT rowid FROM files_fts WHERE files_fts MATCH ?", (phrase,))]
        except Exception:
            return None

    @staticmethod
    def _fts_sync_row(cur, row_id: int, name: str | None, path: str | None):
        """同步 FTS 表单行（先删后插；name/path 传 None 时仅删除）。无 FTS 表时静默跳过。"""
        try:
            cur.execute("DELETE FROM files_fts WHERE rowid=?", (row_id,))
            if name is not None and path is not None:
                cur.execute("INSERT INTO files_fts(rowid, name, path) VALUES (?,?,?)",
                            (row_id, name, path))
        except sqlite3.OperationalError:
            pass

    _COL_DB = {"name": "name_lower", "path": "path_lower", "size": "size", "modified": "modified"}

    @staticmethod
    def _query_files(query: str = "", limit: int = PAGE_SIZE, offset: int = 0,
                     order_col: str = "name", order_desc: bool = False,
                     count_only: bool = False):
        """使用参数化条件查询索引，统一支持搜索、排序、计数和分页。

        搜索路径：关键词 ≥3 字符且 FTS 可用 → FTS5 trigram 子串匹配
        （MATCH 结果 rowid 落临时表后 JOIN files 排序分页）；
        否则 → LIKE 全表扫描（关键词做 %/_/\\ 转义）。
        """
        if not INDEX_DB.exists():
            return 0 if count_only else []
        conn = sqlite3.connect(str(INDEX_DB))
        conn.row_factory = sqlite3.Row
        has_is_dir, fts_ok = IndexEngine._db_schema()
        is_dir_sql = "is_dir" if has_is_dir else "0"
        clauses, params = [], []
        q = query.strip().lower()
        fts_ids = None
        if q:
            if fts_ok and len(q) >= 3:
                fts_ids = IndexEngine._fts_query_ids(conn, q)
            if fts_ids is None:
                eq = (q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_"))
                clauses.append("(name_lower LIKE ? ESCAPE '\\' OR path_lower LIKE ? ESCAPE '\\')")
                params.extend((f"%{eq}%", f"%{eq}%"))

        join = ""
        if fts_ids is not None:
            conn.execute("CREATE TEMP TABLE _fts_ids(id INTEGER PRIMARY KEY)")
            conn.executemany("INSERT OR IGNORE INTO _fts_ids(id) VALUES (?)",
                             ((r,) for r in fts_ids))
            join = " JOIN _fts_ids ON files.id = _fts_ids.id"
        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        if count_only:
            row = conn.execute("SELECT COUNT(*) FROM files" + join + where, params).fetchone()
            conn.close()
            return row[0] if row else 0

        col = IndexEngine._COL_DB.get(order_col, "name_lower")
        direction = "DESC" if order_desc else "ASC"
        if order_col == "type":
            type_parts = []
            if has_is_dir:
                type_parts.append("is_dir DESC")
            type_parts.extend((
                f"lower(substr(name, instr(name, '.'))) {direction}",
                "name_lower ASC",
                "path_lower ASC",
            ))
            order = ", ".join(type_parts)
        else:
            order = f"{col} {direction}, path_lower ASC"
        sql = (
            f"SELECT name, path, size, modified, {is_dir_sql} AS is_dir FROM files{join}"
            f"{where} ORDER BY {order} LIMIT ? OFFSET ?"
        )
        rows = conn.execute(sql, (*params, limit, offset)).fetchall()
        conn.close()
        return [dict(row) for row in rows]

    @staticmethod
    def search(query: str, limit: int = PAGE_SIZE, offset: int = 0, order_col: str = "name",
               order_desc: bool = False) -> list[dict]:
        return IndexEngine._query_files(query, limit, offset, order_col, order_desc)

    @staticmethod
    def load_all(limit: int = PAGE_SIZE, offset: int = 0, order_col: str = "size",
                 order_desc: bool = True) -> list[dict]:
        return IndexEngine._query_files("", limit, offset, order_col, order_desc)

    @staticmethod
    def result_count(query: str = "") -> int:
        return IndexEngine._query_files(query, count_only=True)

    @staticmethod
    def index_file_count() -> int:
        """返回索引中的文件总数。"""
        if not INDEX_DB.exists():
            return 0
        conn = sqlite3.connect(str(INDEX_DB))
        row = conn.execute("SELECT COUNT(*) FROM files").fetchone()
        conn.close()
        return row[0] if row else 0

    @staticmethod
    def remove_paths(paths: list[str]):
        """从索引中同步删除指定路径及其子路径（含 FTS 表）。"""
        if not paths or not INDEX_DB.exists():
            return
        conn = sqlite3.connect(str(INDEX_DB))
        try:
            for path in paths:
                normalized = os.path.normcase(os.path.normpath(path)).lower()
                child_prefix = normalized.rstrip("\\/") + os.sep
                escaped_child_prefix = (child_prefix.replace("\\", "\\\\")
                                        .replace("%", "\\%")
                                        .replace("_", "\\_"))
                try:
                    # 先删 FTS（files 行还在，可经子查询取 rowid）
                    conn.execute(
                        "DELETE FROM files_fts WHERE rowid IN "
                        "(SELECT id FROM files WHERE path_lower = ? OR path_lower LIKE ? ESCAPE '\\')",
                        (normalized, escaped_child_prefix + "%"),
                    )
                except sqlite3.OperationalError:
                    pass
                conn.execute(
                    "DELETE FROM files WHERE path_lower = ? OR path_lower LIKE ? ESCAPE '\\'",
                    (normalized, escaped_child_prefix + "%"),
                )
            conn.commit()
        finally:
            conn.close()

    @staticmethod
    def rename_path(old_path: str, new_path: str):
        """同步索引中的重命名：更新该行，若为文件夹则级联更新所有子路径。

        子路径在 Python 侧逐行计算新路径再 UPDATE，避开 LIKE 前缀匹配在
        Windows 反斜杠 + 转义下的坑，逻辑直观可靠。
        新路径若已在索引中（磁盘 rename 已成功覆盖），先删除该目标行再更新，
        避免撞 path UNIQUE 约束导致「磁盘已改名、索引未同步」的不一致。
        """
        if not INDEX_DB.exists():
            return
        old_n = os.path.normpath(old_path)
        new_n = os.path.normpath(new_path)
        old_lower = old_n.lower()
        old_prefix = old_lower.rstrip("\\/") + os.sep
        new_name = os.path.basename(new_n)
        conn = sqlite3.connect(str(INDEX_DB))
        try:
            cur = conn.cursor()
            # 先取自身行 id（重命名仅改大小写时 new_lower == old_lower，不能把自身当目标删掉）
            cur.execute("SELECT id FROM files WHERE path_lower=?", (old_lower,))
            self_row = cur.fetchone()
            # 0) 若目标行已存在（磁盘 rename 覆盖了旧文件），先删目标行及其 FTS 条目
            cur.execute("SELECT id FROM files WHERE path_lower=?", (new_n.lower(),))
            target_row = cur.fetchone()
            if target_row and (self_row is None or target_row[0] != self_row[0]):
                IndexEngine._fts_sync_row(cur, target_row[0], None, None)
                cur.execute("DELETE FROM files WHERE id=?", (target_row[0],))
            # 1) 更新自身行（文件或文件夹本身）
            if self_row:
                cur.execute(
                    "UPDATE files SET name=?, name_lower=?, path=?, path_lower=? WHERE id=?",
                    (new_name, new_name.lower(), new_n, new_n.lower(), self_row[0]),
                )
                IndexEngine._fts_sync_row(cur, self_row[0], new_name, new_n)
            # 2) 取出所有子行（path_lower 以旧前缀开头），逐个改路径
            cur.execute("SELECT id, path, name FROM files WHERE substr(path_lower, 1, ?) = ?",
                        (len(old_prefix), old_prefix))
            children = cur.fetchall()
            for row_id, child_path, child_name in children:
                # 用原 path 的后缀（保留原始大小写）拼到新前缀
                suffix = os.path.normpath(child_path)[len(old_n):].lstrip("\\/")
                child_new = os.path.join(new_n, suffix)
                cur.execute("UPDATE files SET path=?, path_lower=? WHERE id=?",
                            (child_new, child_new.lower(), row_id))
                IndexEngine._fts_sync_row(cur, row_id, child_name, child_new)
            conn.commit()
        finally:
            conn.close()

    # ---- 排除列表管理 ----

    @classmethod
    def _load_exclude_list(cls):
        """从 JSON 文件加载用户排除列表。"""
        if not cls.EXCLUDE_FILE.exists():
            return set(), []
        try:
            data = json.loads(cls.EXCLUDE_FILE.read_text(encoding="utf-8"))
            dirs = {d.lower() for d in data.get("dirs", [])}
            paths = [p.lower() for p in data.get("paths", [])]
            return dirs, paths
        except Exception:
            return set(), []

    @classmethod
    def save_exclude_list(cls, dirs: list[str], paths: list[str]):
        """保存排除列表到 JSON 文件（目录不存在时先创建，兼容未建索引的首次使用）。"""
        INDEX_DIR.mkdir(parents=True, exist_ok=True)
        cls.EXCLUDE_FILE.write_text(
            json.dumps({"dirs": dirs, "paths": paths}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    @classmethod
    def get_exclude_list(cls) -> dict:
        """获取排除列表原始数据（用于 GUI 展示）。"""
        if not cls.EXCLUDE_FILE.exists():
            return {"dirs": [], "paths": []}
        try:
            return json.loads(cls.EXCLUDE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {"dirs": [], "paths": []}

    # ---- 设置管理 ----

    DEFAULT_SETTINGS = {
        "auto_index_on_start": False,   # 启动时自动增量更新索引
        "tray_auto_index": False,        # 最小化到托盘后自动更新索引
        "tray_auto_index_minutes": 30,   # 托盘自动更新间隔（分钟）
        "theme": "mist",
        "text_pt": 16,                   # 主字号（表格正文基准 pt，10~40 可调）
    }

    @classmethod
    def load_settings(cls) -> dict:
        """从 JSON 文件读取设置，缺失项使用默认值。"""
        settings = dict(cls.DEFAULT_SETTINGS)
        if cls.SETTINGS_FILE.exists():
            try:
                data = json.loads(cls.SETTINGS_FILE.read_text(encoding="utf-8"))
                settings.update({k: v for k, v in data.items() if k in settings})
            except Exception:
                pass
        if settings.get("theme") not in THEMES:
            settings["theme"] = cls.DEFAULT_SETTINGS["theme"]
        return settings

    @classmethod
    def save_settings(cls, settings: dict):
        """将设置保存到 JSON 文件。"""
        INDEX_DIR.mkdir(parents=True, exist_ok=True)
        cls.SETTINGS_FILE.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )




def _diag(msg: str):
    """追加诊断行到 INDEX_DIR/debug.log（超过 1MB 自动轮转为 debug.log.1）。

    日志位置随索引目录走（~/.file_searcher_index/），不再依赖开发机绝对路径，
    打包 exe 移到任意位置都可正常记录。
    """
    try:
        INDEX_DIR.mkdir(parents=True, exist_ok=True)
        log_path = DEBUG_LOG
        if log_path.exists() and log_path.stat().st_size > DEBUG_LOG_MAX:
            try:
                log_path.replace(log_path.with_suffix(".log.1"))
            except Exception:
                pass
        with open(log_path, "a", encoding="utf-8") as fo:
            fo.write(msg + "\n")
    except Exception:
        pass


def _rounded_rect(canvas, x1, y1, x2, y2, radius, **kwargs):
    """在 Canvas 上绘制平滑圆角矩形。"""
    radius = max(1, min(radius, (x2 - x1) / 2, (y2 - y1) / 2))
    points = (
        x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
        x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
        x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
    )
    return canvas.create_polygon(points, smooth=True, splinesteps=24, **kwargs)


_GRADIENT_CACHE: dict = {}

def _make_gradient_pix(master, width: int, height: int, radius: int, c1: str, c2: str):
    """生成垂直渐变的圆角位图（带透明圆角遮罩），带缓存。"""
    key = (width, height, radius, c1, c2)
    if key in _GRADIENT_CACHE:
        return _GRADIENT_CACHE[key]
    w, h = max(2, width), max(2, height)
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    rgb1 = tuple(int(c1.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    rgb2 = tuple(int(c2.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    steps = max(2, h)
    for y in range(steps):
        t = y / (steps - 1)
        col = tuple(int(rgb1[i] + (rgb2[i] - rgb1[i]) * t) for i in range(3))
        d.line([(0, y), (w, y)], fill=col + (255,))
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, w - 1, h - 1), radius=radius, fill=255)
    img.putalpha(mask)
    pix = ImageTk.PhotoImage(img, master=master)
    _GRADIENT_CACHE[key] = pix
    return pix


class RoundedButton(tk.Canvas):
    """支持 hover、pressed、disabled、渐变主按钮且兼容 config(text/state) 的圆角按钮。

    kind: "normal"（灰底描边） | "accent"（霁青渐变主按钮） | "danger"（红色系）
          | "ghost"（透明底 accent 文字，hover 淡 accent 底）
    """

    def __init__(self, master, text="", command=None, width=120, height=48, radius=8,
                 colors=None, icon=None, kind="normal", font_size=None, **kwargs):
        self.colors = colors or THEMES["dark"]
        self._kind = kind
        self._font_size = font_size or FONT_BODY
        super().__init__(master, width=width, height=height, bd=0, highlightthickness=0,
                         bg=master.cget("bg"), cursor="hand2", **kwargs)
        self._text = text
        self._command = command
        self._radius = radius
        self._icon = icon
        self._state = tk.NORMAL
        self._visual_state = "normal"
        self.bind("<Configure>", lambda _e: self._draw())
        self.bind("<Enter>", lambda _e: self._set_visual("hover"))
        self.bind("<Leave>", lambda _e: self._set_visual("normal"))
        self.bind("<ButtonPress-1>", lambda _e: self._set_visual("pressed"))
        self.bind("<ButtonRelease-1>", self._release)
        self._draw()

    def _set_visual(self, state):
        if self._state != tk.DISABLED:
            self._visual_state = state
            self._draw()

    def _release(self, event):
        if self._state == tk.DISABLED:
            return
        inside = 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height()
        self._set_visual("hover" if inside else "normal")
        if inside and self._command:
            self._command()

    def _draw(self):
        self.delete("all")
        width, height = max(2, self.winfo_width()), max(2, self.winfo_height())
        c = self.colors
        if self._state == tk.DISABLED:
            if self._kind == "ghost":
                self.create_text(width / 2, height / 2, text=self._icon or self._text,
                                 fill=c["muted_2"], font=(FONT_FAMILY, self._font_size, "bold"))
                return
            fill, outline, fg = c["surface"], c["border"], c["muted_2"]
        elif self._kind == "accent":
            if self._visual_state == "pressed":
                fill, outline, fg = c["accent_pressed"], c["accent_pressed"], "#FFFFFF"
            elif self._visual_state == "hover":
                fill, outline, fg = c["accent_hover"], c["accent_hover"], "#FFFFFF"
            else:
                pix = _make_gradient_pix(self.master.winfo_toplevel(), width, height,
                                         self._radius, c["accent_grad_a"], c["accent_grad_b"])
                self.create_image(width / 2, height / 2, image=pix)
                self.create_text(width / 2, height / 2, text=self._icon or self._text,
                                 fill="#FFFFFF", font=(FONT_FAMILY, self._font_size, "bold"))
                return
            _rounded_rect(self, 1, 1, width - 1, height - 1, self._radius,
                          fill=fill, outline=outline, width=1)
            self.create_text(width / 2, height / 2, text=self._icon or self._text,
                             fill=fg, font=(FONT_FAMILY, self._font_size, "bold"))
            return
        elif self._kind == "danger":
            if self._visual_state == "pressed":
                pix = _make_gradient_pix(self.master.winfo_toplevel(), width, height,
                                         self._radius, "#C44854", "#A93A45")
                self.create_image(width / 2, height / 2, image=pix)
            elif self._visual_state == "hover":
                pix = _make_gradient_pix(self.master.winfo_toplevel(), width, height,
                                         self._radius, "#EE707B", "#D14F5B")
                self.create_image(width / 2, height / 2, image=pix)
            else:
                pix = _make_gradient_pix(self.master.winfo_toplevel(), width, height,
                                         self._radius, "#E2636E", "#C94854")
                self.create_image(width / 2, height / 2, image=pix)
            self.create_text(width / 2, height / 2, text=self._icon or self._text,
                             fill="#FFFFFF", font=(FONT_FAMILY, self._font_size, "bold"))
            return
        elif self._kind == "ghost":
            # 幽灵按钮：常态透明底 accent 文字；hover/pressed 淡 accent 底
            if self._visual_state == "pressed":
                _rounded_rect(self, 1, 1, width - 1, height - 1, self._radius,
                              fill=c["selected_hover"], outline="", width=0)
            elif self._visual_state == "hover":
                _rounded_rect(self, 1, 1, width - 1, height - 1, self._radius,
                              fill=c["selected"], outline="", width=0)
            self.create_text(width / 2, height / 2, text=self._icon or self._text,
                             fill=c["accent"], font=(FONT_FAMILY, self._font_size, "bold"))
            return
        elif self._visual_state == "pressed":
            fill, outline, fg = c["surface_3"], c["border_strong"], c["text"]
        elif self._visual_state == "hover":
            fill, outline, fg = c["surface_3"], c["border_strong"], c["text"]
        else:
            fill, outline, fg = c["surface_alt"], c["border"], c["text"]
        _rounded_rect(self, 1, 1, width - 1, height - 1, self._radius,
                      fill=fill, outline=outline, width=1)
        label = self._icon if self._icon and not self._text else self._text
        self.create_text(width / 2, height / 2, text=label, fill=fg,
                         font=(FONT_FAMILY, self._font_size, "bold" if self._text else "normal"))

    def configure(self, cnf=None, **kwargs):
        if cnf is None and not kwargs:
            return super().configure()
        if isinstance(cnf, str):
            if cnf == "text":
                return ("text", "text", "Text", "", self._text)
            if cnf == "state":
                return ("state", "state", "State", tk.NORMAL, self._state)
            if cnf == "command":
                return ("command", "command", "Command", "", self._command)
            return super().configure(cnf)
        if cnf:
            kwargs.update(cnf)
        if "text" in kwargs:
            self._text = kwargs.pop("text")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
            super().configure(cursor="" if self._state == tk.DISABLED else "hand2")
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        result = super().configure(**kwargs) if kwargs else None
        self._draw()
        return result

    def cget(self, key):
        if key == "text":
            return self._text
        if key == "state":
            return self._state
        if key == "command":
            return self._command
        return super().cget(key)

    config = configure


class RoundedSearchBox(tk.Canvas):
    """墨玉风大搜索框：玻璃底、放大镜、聚焦 accent 描边、placeholder、清空热区。"""

    def __init__(self, master, textvariable, colors, clear_command, height=64, font_size=None,
                 placeholder=""):
        super().__init__(master, height=height, bd=0, highlightthickness=0,
                         bg=master.cget("bg"))
        self.colors = colors
        self._focused = False
        self._hover = False
        self._radius = 6
        self._font_size = font_size or FONT_INPUT
        self._placeholder = placeholder
        self._placeholder_on = False
        self._var = textvariable
        self.entry = tk.Entry(self, textvariable=textvariable, relief="flat", bd=0,
                              bg=colors["input"], fg=colors["text"],
                              insertbackground=colors["accent"],
                              font=(FONT_FAMILY, self._font_size))
        self._entry_window = self.create_window(58, height / 2, window=self.entry, anchor=tk.W)
        self._clear_command = clear_command
        self.bind("<Configure>", self._layout)
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda _e: self._on_hover(True))
        self.bind("<Leave>", lambda _e: self._on_hover(False))
        self.entry.bind("<FocusIn>", self._focus_in, add="+")
        self.entry.bind("<FocusOut>", self._focus_out, add="+")
        self._var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _layout(self, _event=None):
        h = max(24, self.winfo_height() - 10)
        self.itemconfigure(self._entry_window, width=max(20, self.winfo_width() - 108), height=h)
        self._draw()

    def _on_hover(self, hover: bool):
        self._hover = hover
        self._draw()

    def _draw(self):
        self.delete("shell")
        width, height = max(2, self.winfo_width()), max(2, self.winfo_height())
        c = self.colors
        if self._focused:
            border = c["accent"]
        elif self._hover:
            border = c["border_strong"]
        else:
            border = c["border"]
        shell = _rounded_rect(self, 1.5, 1.5, width - 1.5, height - 1.5, self._radius,
                              fill=c["input"], outline=border,
                              width=1.8 if self._focused else 1, tags="shell")
        self.tag_lower(shell)
        # 放大镜（聚焦时 accent 色）
        icon_c = c["accent"] if self._focused else c["muted"]
        cx, cy = 32, height / 2
        self.create_oval(cx - 8.5, cy - 8.5, cx + 8.5, cy + 8.5,
                         outline=icon_c, width=2.4, tags="shell")
        self.create_line(cx + 7, cy + 7, cx + 14, cy + 14,
                         fill=icon_c, width=3, capstyle=tk.ROUND, tags="shell")
        has_text = bool(self._var.get())
        if has_text:
            # 清空按钮
            clear_x = width - 34
            self.create_oval(clear_x - 13, cy - 13, clear_x + 13, cy + 13,
                             fill=c["surface_3"], outline="", tags="shell")
            self.create_text(clear_x, cy, text="✕", fill=c["muted"],
                             font=(FONT_FAMILY, max(8, int(self._font_size * 0.62))), tags="shell")
        else:
            # placeholder：画在 entry 之上（Entry 无内容时背景透明区可见此文字）
            self.create_text(58, cy, text=self._placeholder, anchor=tk.W,
                             fill=c["muted_2"], font=(FONT_FAMILY, self._font_size),
                             tags=("shell", "placeholder"))
        self.tag_lower("shell", self._entry_window)
        if not has_text:
            # 占位文字须压在 entry 上层才可见（entry 透明，文字从其背景色区域透出）
            self.tag_raise("placeholder", self._entry_window)

    def _click(self, event):
        if event.x >= self.winfo_width() - 56 and self._var.get():
            self._clear_command()
        else:
            self.entry.focus_set()

    def _focus_in(self, _event=None):
        self._focused = True
        self._draw()

    def _focus_out(self, _event=None):
        self._focused = False
        self._draw()


# ================================================================
#  自绘弹窗组件：圆角对话框外壳 / 确认框 / 输入框 / 右键菜单 / 开关
# ================================================================

_DIALOG_MAGIC = "#010203"   # 透明魔法色（避开全部界面用色）


class ToggleSwitch(tk.Canvas):
    """自绘滑动开关：绑定 BooleanVar，点击切换并可回调。"""

    def __init__(self, master, colors, variable, command=None, width=46, height=26):
        super().__init__(master, width=width, height=height, bd=0, highlightthickness=0,
                         bg=master.cget("bg"), cursor="hand2")
        self.colors = colors
        self._var = variable
        self._command = command
        # 注意：不能用 self._w（tkinter 内部 widget 路径名属性），改用 _sw/_sh
        self._sw, self._sh = width, height
        self.bind("<Button-1>", self._toggle)
        self._var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _toggle(self, _event=None):
        self._var.set(not self._var.get())
        if self._command:
            self._command()

    def _draw(self):
        self.delete("all")
        c = self.colors
        w, h = self._sw, self._sh
        on = bool(self._var.get())
        if on:
            pix = _make_gradient_pix(self.winfo_toplevel(), w, h, h // 2,
                                     c["accent_grad_a"], c["accent_grad_b"])
            self.create_image(w / 2, h / 2, image=pix)
            cx = w - h / 2 - 1
        else:
            _rounded_rect(self, 1, 1, w - 1, h - 1, (h - 2) / 2,
                          fill=c["surface_3"], outline=c["border_strong"], width=1)
            cx = h / 2 + 1
        r = (h - 8) / 2
        self.create_oval(cx - r, h / 2 - r, cx + r, h / 2 + r, fill="#FFFFFF", outline="")


class RoundEntry(tk.Canvas):
    """弹窗用圆角输入框：聚焦 accent 描边 + 光晕。"""

    def __init__(self, master, colors, height=40, font=None, radius=10):
        super().__init__(master, height=height, bd=0, highlightthickness=0,
                         bg=master.cget("bg"))
        self.colors = colors
        self._radius = radius
        self._font = font or (FONT_FAMILY, FONT_BODY)
        self.var = tk.StringVar()
        self.entry = tk.Entry(self, textvariable=self.var, relief="flat", bd=0,
                              bg=colors["input"], fg=colors["text"],
                              insertbackground=colors["accent"], font=self._font)
        self._entry_window = self.create_window(14, height / 2, window=self.entry, anchor=tk.W)
        self._focused = False
        self.bind("<Configure>", self._layout)
        self.bind("<Button-1>", lambda _e: self.entry.focus_set())
        self.entry.bind("<FocusIn>", self._focus_in, add="+")
        self.entry.bind("<FocusOut>", self._focus_out, add="+")
        self._draw()

    def _layout(self, _event=None):
        self.itemconfigure(self._entry_window,
                           width=max(20, self.winfo_width() - 28),
                           height=max(20, self.winfo_height() - 12))
        self._draw()

    def _draw(self):
        self.delete("shell")
        c = self.colors
        w, h = max(2, self.winfo_width()), max(2, self.winfo_height())
        if self._focused:
            _rounded_rect(self, 1, 1, w - 1, h - 1, self._radius + 2,
                          fill="", outline=c["selected"], width=4, tags="shell")
        shell = _rounded_rect(self, 2, 2, w - 2, h - 2, self._radius,
                              fill=c["input"],
                              outline=c["accent"] if self._focused else c["border"],
                              width=1.5, tags="shell")
        self.tag_lower("shell", self._entry_window)

    def _focus_in(self, _event=None):
        self._focused = True
        self._draw()

    def _focus_out(self, _event=None):
        self._focused = False
        self._draw()

    def get(self):
        return self.var.get()

    def set(self, value):
        self.var.set(value)


class _DialogShell:
    """无边框圆角模态弹窗外壳：透明角、阴影、圆角卡片、标题区拖动、Esc 关闭。"""

    def __init__(self, app, width_px, height_px, radius=10, follow_mouse=False, at_xy=None):
        self.app = app
        self.root = app.root
        self.colors = app.colors
        self.result = None
        self._radius = radius
        top = tk.Toplevel(self.root)
        self.top = top
        top.withdraw()                          # 先隐藏，定位完成后再显示，避免 Windows 重定位
        top.overrideredirect(True)
        self._rounded_ok = True
        try:
            top.configure(bg=_DIALOG_MAGIC)
            top.attributes("-transparentcolor", _DIALOG_MAGIC)
        except tk.TclError:
            self._rounded_ok = False
            top.configure(bg=self.colors["dialog_bg"])
        # 注意：不设 transient(self.root)。transient 会让 Windows 把弹窗锚定到主窗口，
        # 导致我们 geometry 设置的跟随鼠标坐标被强制改回主窗附近。
        top.resizable(False, False)
        self.root.update_idletasks()
        if at_xy is not None:
            # 显式指定屏幕坐标（如右键点下位置），最可靠的跟随鼠标
            x, y = at_xy[0] + 12, at_xy[1] + 12
        elif follow_mouse:
            mx, my = self.root.winfo_pointerx(), self.root.winfo_pointery()
            x, y = mx + 12, my + 12
        else:
            pw, ph = self.root.winfo_width(), self.root.winfo_height()
            px, py = self.root.winfo_rootx(), self.root.winfo_rooty()
            x = px + (pw - width_px) // 2
            y = py + (ph - height_px) // 2
        # 防御：多显示器/主窗几何异常时，弹窗坐标钳制到屏幕内并强制置顶可见
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x = max(0, min(x, sw - width_px))
        y = max(0, min(y, sh - height_px))
        top.geometry(f"{width_px}x{height_px}+{x}+{y}")
        top.update_idletasks()                  # 让 geometry 真正应用
        top.geometry(f"{width_px}x{height_px}+{x}+{y}")  # 二次锁定，防 Windows 改动
        top.deiconify()                         # 定位完成后显示
        top.attributes("-topmost", True)          # 弹窗必须浮在无边框主窗之上
        top.lift()
        top.after(120, lambda: top.attributes("-topmost", False))  # 短暂置顶后归还层级

        bg = _DIALOG_MAGIC if self._rounded_ok else self.colors["dialog_bg"]
        self.canvas = tk.Canvas(top, width=width_px, height=height_px, bd=0,
                                highlightthickness=0, bg=bg)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        pad = 12 if self._rounded_ok else 1   # 阴影留白
        self._pad = pad
        self._draw_card(width_px, height_px)
        self.body = tk.Frame(self.canvas, bg=self.colors["dialog_bg"])
        self.canvas.create_window(pad + 1, pad + 1, window=self.body, anchor=tk.NW,
                                  width=width_px - pad * 2 - 2, height=height_px - pad * 2 - 2)
        self._drag = None
        self.canvas.bind("<Button-1>", self._drag_start)
        self.canvas.bind("<B1-Motion>", self._drag_move)
        top.bind("<Escape>", lambda _e: self.close())

    def _draw_card(self, w, h):
        c = self.colors
        p = self._pad
        # 纸感投影：浅色用极淡暖灰双层，深色用近黑单层
        if self._rounded_ok:
            if self.app._theme_name == "dark":
                _rounded_rect(self.canvas, p - 4, p + 3, w - p + 4, h - p + 7,
                              self._radius + 3, fill="#05080A", outline="", width=0)
            else:
                _rounded_rect(self.canvas, p - 3, p + 4, w - p + 3, h - p + 8,
                              self._radius + 3, fill="#DAD5C9", outline="", width=0)
                _rounded_rect(self.canvas, p - 1, p + 1, w - p + 1, h - p + 2,
                              self._radius + 1, fill="#E4DFD3", outline="", width=0)
        _rounded_rect(self.canvas, p, p, w - p, h - p, self._radius,
                      fill=c["dialog_bg"], outline=c["border"], width=1)

    def _drag_start(self, event):
        self._drag = (event.x_root - self.top.winfo_x(), event.y_root - self.top.winfo_y())

    def _drag_move(self, event):
        if self._drag:
            self.top.geometry(f"+{event.x_root - self._drag[0]}+{event.y_root - self._drag[1]}")

    def close(self, result=None):
        self.result = result
        try:
            self.top.grab_release()
        except Exception:
            pass
        self.top.destroy()

    def run(self, grab_focus=True):
        """模态运行，返回 result。grab_focus=False 时不抢占焦点（由调用方聚焦到具体输入框）。"""
        self.top.grab_set()
        if grab_focus:
            self.top.focus_set()
        self.root.wait_window(self.top)
        return self.result


def _dialog_confirm(app, title, desc, kind="warn", ok_text="确定", cancel_text="取消",
                    show_cancel=True):
    """确认/提示弹窗。kind: warn(黄) / danger(红) / info(青)。返回 True=确认。"""
    c = app.colors
    s = app._s
    w = s(400)
    # 估算高度：图标 46 + 标题 30 + 描述行数 + 按钮 52 + 边距。
    # 按实际字体测量换行（wraplength 生效），避免超长 desc 裁掉按钮区。
    try:
        import tkinter.font as tkfont
        f = tkfont.Font(root=app.root, font=app._f(FONT_SMALL))
        wrap = max(50, w - s(90))
        lines = 0
        for seg in desc.split("\n"):
            lines += max(1, math.ceil(f.measure(seg) / wrap))
    except Exception:
        approx_chars_per_line = 26
        lines = max(1, (len(desc) + approx_chars_per_line - 1) // approx_chars_per_line)
    h = s(150) + lines * s(20) + s(64)
    shell = _DialogShell(app, w, h)
    body = tk.Frame(shell.body, bg=c["dialog_bg"])
    body.pack(fill=tk.BOTH, expand=True, padx=s(22))
    icon_map = {"warn": ("!", c["warning"]), "danger": ("✕", c["error"]), "info": ("i", c["accent"])}
    icon_char, icon_c = icon_map.get(kind, icon_map["info"])
    badge_bg = {"warn": BADGE_STYLES[app._theme_name]["zip"][1],
                "danger": BADGE_STYLES[app._theme_name]["pdf"][1],
                "info": app.colors["selected"]}[kind]
    ic = tk.Canvas(body, width=s(44), height=s(44), bd=0, highlightthickness=0, bg=c["dialog_bg"])
    ic.pack(anchor=tk.W, pady=(s(18), s(10)))
    ic.create_oval(2, 2, s(44) - 2, s(44) - 2, fill=badge_bg, outline="")
    ic.create_text(s(22), s(22), text=icon_char, fill=icon_c, font=app._f(FONT_LG, "bold"))
    tk.Label(body, text=title, bg=c["dialog_bg"], fg=c["text"],
             font=app._f(FONT_LG, "bold")).pack(anchor=tk.W)
    tk.Label(body, text=desc, bg=c["dialog_bg"], fg=c["muted"],
             font=app._f(FONT_SMALL), justify=tk.LEFT, wraplength=w - s(90)).pack(
        anchor=tk.W, pady=(s(7), 0))
    btns = tk.Frame(body, bg=c["dialog_bg"])
    btns.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, s(16)))
    ok_kind = "danger" if kind == "danger" else "accent"
    RoundedButton(btns, text=ok_text, command=lambda: shell.close(True),
                  width=s(96), height=s(36), colors=c, kind=ok_kind,
                  font_size=app._f(FONT_BODY)[1]).pack(side=tk.RIGHT)
    if show_cancel:
        RoundedButton(btns, text=cancel_text, command=lambda: shell.close(False),
                      width=s(96), height=s(36), colors=c,
                      font_size=app._f(FONT_BODY)[1]).pack(side=tk.RIGHT, padx=(0, s(8)))
    shell.top.bind("<Return>", lambda _e: shell.close(True))
    return bool(shell.run())


def _dialog_input(app, title, desc, initial="", ok_text="确定", options=None,
                  selected_option=None, follow_mouse=False, cursor_start=False, at_xy=None):
    """输入弹窗。options 提供时显示分段胶囊选择器。

    at_xy：弹窗出现的屏幕坐标 (x,y)，优先于 follow_mouse，用于精确跟随右键点下位置。
    follow_mouse：弹窗跟随当前鼠标位置。
    cursor_start：光标落在初始文字最前（不选中、不再点一下），便于直接在最前输入。
    返回：未提供 options → 输入字符串或 None；提供 options → (option_value, 文本) 或 None。
    """
    c = app.colors
    s = app._s
    w = s(460)
    h = s(150) + (s(44) if options else 0) + s(90)
    shell = _DialogShell(app, w, h, follow_mouse=follow_mouse, at_xy=at_xy)
    body = tk.Frame(shell.body, bg=c["dialog_bg"])
    body.pack(fill=tk.BOTH, expand=True, padx=s(22))
    tk.Label(body, text=title, bg=c["dialog_bg"], fg=c["text"],
             font=app._f(FONT_LG, "bold")).pack(anchor=tk.W, pady=(s(18), 0))
    if desc:
        tk.Label(body, text=desc, bg=c["dialog_bg"], fg=c["muted"],
                 font=app._f(FONT_SMALL), anchor=tk.W, justify=tk.LEFT).pack(
            anchor=tk.W, pady=(s(6), 0))

    option_var = tk.StringVar(value=selected_option or (options[0][1] if options else ""))
    seg_canvases = []

    def _draw_seg():
        for cv, (label, value) in seg_canvases:
            cv.delete("all")
            cw = max(2, cv.winfo_width())
            ch = max(2, cv.winfo_height())
            active = option_var.get() == value
            if active:
                _rounded_rect(cv, 1, 1, cw - 1, ch - 1, (ch - 2) / 2,
                              fill=c["selected"], outline=c["accent"], width=1.2)
            cv.create_text(cw / 2, ch / 2, text=label,
                           fill=c["accent"] if active else c["muted"],
                           font=app._f(FONT_SMALL, "bold" if active else "normal"))

    if options:
        seg = tk.Frame(body, bg=c["dialog_bg"])
        seg.pack(anchor=tk.W, pady=(s(12), 0))
        for label, value in options:
            cv = tk.Canvas(seg, width=s(104), height=s(30), bd=0, highlightthickness=0,
                           bg=c["dialog_bg"], cursor="hand2")
            cv.pack(side=tk.LEFT, padx=(0, s(8)))
            cv.bind("<Button-1>", lambda _e, v=value: (option_var.set(v), _draw_seg()))
            cv.bind("<Configure>", lambda _e: _draw_seg())
            seg_canvases.append((cv, (label, value)))
        seg_canvases and shell.top.after_idle(_draw_seg)

    entry_box = RoundEntry(body, c, height=s(42), font=app._f(FONT_BODY))
    entry_box.pack(fill=tk.X, pady=(s(12), 0))
    entry_box.set(initial)

    def _place_cursor():
        # grab/topmost 时序后强制聚焦并定位光标（压过 shell.run 的 focus_set）
        entry_box.entry.focus_set()
        if cursor_start:
            entry_box.entry.selection_clear()
            entry_box.entry.icursor(0)
        else:
            entry_box.entry.selection_range(0, tk.END)
    # 立即设一次，再在事件循环空闲+短延迟后各设一次，确保生效
    _place_cursor()
    shell.top.after_idle(_place_cursor)
    shell.top.after(50, _place_cursor)

    def _ok():
        text = entry_box.get().strip()
        if not text:
            return
        shell.close((option_var.get(), text) if options else text)

    btns = tk.Frame(body, bg=c["dialog_bg"])
    btns.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, s(12)))
    RoundedButton(btns, text=ok_text, command=_ok,
                  width=s(96), height=s(36), colors=c, kind="accent",
                  font_size=app._f(FONT_BODY)[1]).pack(side=tk.RIGHT)
    RoundedButton(btns, text="取消", command=lambda: shell.close(None),
                  width=s(96), height=s(36), colors=c,
                  font_size=app._f(FONT_BODY)[1]).pack(side=tk.RIGHT, padx=(0, s(8)))
    entry_box.entry.bind("<Return>", lambda _e: _ok())
    # 不让 run() 抢占焦点，由 _place_cursor 聚焦到输入框
    return shell.run(grab_focus=False)


class CtxMenu:
    """自绘右键菜单：圆角、阴影、图标、悬停高亮、分隔线、危险项。"""

    ITEM_H = 32
    WIDTH = 236

    def __init__(self, app):
        self.app = app
        self.root = app.root
        self.colors = app.colors
        self.top = None
        self._items = []
        self._hover = None

    # ---- 对外接口 ----
    def show(self, x_root, y_root, items):
        """items: [dict(text, icon, cmd, kind='normal'|'danger', disabled=False)] 或 ('sep',)。"""
        self.close()
        self._items = items
        s = self.app._s
        c = self.colors
        item_h = s(self.ITEM_H)
        sep_h = s(11)
        pad_y = s(6)
        w = s(self.WIDTH)
        h = pad_y * 2 + sum(sep_h if it == ("sep",) else item_h for it in items)
        top = tk.Toplevel(self.root)
        self.top = top
        top.overrideredirect(True)
        self._rounded_ok = True
        try:
            top.configure(bg=_DIALOG_MAGIC)
            top.attributes("-transparentcolor", _DIALOG_MAGIC)
        except tk.TclError:
            self._rounded_ok = False
            top.configure(bg=c["menu_bg"])
        top.transient(self.root)
        # 边界翻转
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        if x_root + w > sw:
            x_root = max(0, x_root - w)
        if y_root + h > sh:
            y_root = max(0, sh - h - 4)
        top.geometry(f"{w}x{h}+{x_root}+{y_root}")
        top.attributes("-topmost", True)   # 无边框主窗下菜单必须置顶，否则被压在主窗后
        top.lift()
        bg = _DIALOG_MAGIC if self._rounded_ok else c["menu_bg"]
        self.canvas = tk.Canvas(top, width=w, height=h, bd=0, highlightthickness=0, bg=bg)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self._item_h, self._sep_h, self._pad_y = item_h, sep_h, pad_y
        self._draw()
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<Leave>", lambda _e: self._set_hover(None))
        self.canvas.bind("<Button-1>", self._on_click)
        top.bind("<Escape>", lambda _e: self.close())
        top.grab_set()
        # 点击菜单外：关闭（grab 下全局 Button-1 会路由到 grab 窗口）
        top.bind("<Button-1>", self._on_global_click, add="+")

    def close(self):
        if self.top is not None:
            try:
                self.top.grab_release()
                self.top.destroy()
            except Exception:
                pass
            self.top = None
            self._hover = None

    # ---- 布局命中 ----
    def _item_at(self, y):
        acc = self._pad_y
        for i, it in enumerate(self._items):
            hh = self._sep_h if it == ("sep",) else self._item_h
            if acc <= y < acc + hh:
                return i if it != ("sep",) else None
            acc += hh
        return None

    def _item_y(self, idx):
        acc = self._pad_y
        for i, it in enumerate(self._items):
            if i == idx:
                return acc
            acc += self._sep_h if it == ("sep",) else self._item_h
        return acc

    # ---- 事件 ----
    def _on_motion(self, event):
        self._set_hover(self._item_at(event.y))

    def _set_hover(self, idx):
        if idx != self._hover:
            self._hover = idx
            self._draw()

    def _on_click(self, event):
        idx = self._item_at(event.y)
        if idx is None:
            return
        item = self._items[idx]
        if item.get("disabled"):
            return
        cmd = item.get("cmd")
        self.close()
        if cmd:
            self.root.after_idle(cmd)

    def _on_global_click(self, event):
        # grab 下点击会路由到 grab 窗口：按屏幕坐标判断是否在菜单矩形外，在外则关闭
        if self.top is None:
            return
        try:
            x0, y0 = self.top.winfo_rootx(), self.top.winfo_rooty()
            x1 = x0 + self.top.winfo_width()
            y1 = y0 + self.top.winfo_height()
        except Exception:
            return
        if not (x0 <= event.x_root <= x1 and y0 <= event.y_root <= y1):
            self.close()

    # ---- 绘制 ----
    def _draw(self):
        cv = self.canvas
        cv.delete("all")
        c = self.colors
        s = self.app._s
        w = s(self.WIDTH)
        h = int(self.top.winfo_height() or 0) or (self._pad_y * 2 + sum(
            self._sep_h if it == ("sep",) else self._item_h for it in self._items))
        # 阴影 + 圆角底
        if self._rounded_ok:
            shadow = "#05080A" if self.app._theme_name == "dark" else "#C9D2CE"
            _rounded_rect(cv, 4, 8, w - 4, h - 2, 13, fill=shadow, outline="", width=0)
        _rounded_rect(cv, 3, 3, w - 3, h - 6, 12, fill=c["menu_bg"],
                      outline=c["border_strong"], width=1)
        inset = s(6)
        acc = self._pad_y
        # 右键菜单字号随主字号走（与设置一致：正文基准 pt 的同一缩放体系）
        font = self.app._f(FONT_BODY)
        icon_font = self.app._f(FONT_BODY)
        danger_hover_bg = BADGE_STYLES[self.app._theme_name]["pdf"][1]
        for i, it in enumerate(self._items):
            if it == ("sep",):
                y = acc + self._sep_h / 2
                cv.create_line(inset + s(6), y, w - inset - s(6), y, fill=c["row_line"])
                acc += self._sep_h
                continue
            y0, y1 = acc, acc + self._item_h
            cy = (y0 + y1) / 2
            disabled = it.get("disabled")
            danger = it.get("kind") == "danger"
            hovered = (i == self._hover) and not disabled
            if hovered:
                bg = danger_hover_bg if danger else c["selected"]
                _rounded_rect(cv, inset, y0 + 1, w - inset, y1 - 1, 8, fill=bg,
                              outline="", width=0)
            if disabled:
                fg = c["muted_2"]
            elif hovered:
                fg = c["error"] if danger else c["sel_text"]
            else:
                fg = c["error"] if danger else c["text"]
            cv.create_text(inset + s(16), cy, text=it.get("icon", ""), anchor=tk.W,
                           fill=c["muted_2"] if disabled else (c["error"] if danger else c["muted"]),
                           font=icon_font)
            cv.create_text(inset + s(40), cy, text=it.get("text", ""), anchor=tk.W,
                           fill=fg, font=font)
            acc += self._item_h


class FileTable(tk.Canvas):
    """自绘呼吸式表格：无网格线、行分隔线、类型徽章、列宽拖动、排序表头、行选中/悬停。

    行数据结构：rows = [{"result": dict, "icon": PhotoImage, "values": tuple}]
    values 顺序与列顺序对应（不含图标列）。
    """

    HEADER_H = 42

    def __init__(self, master, colors, icon_cache, font_body, font_header,
                 on_header_click=None, on_double=None, on_right=None,
                 on_col_resize=None, on_scroll_page=None, badge_styles=None):
        super().__init__(master, bd=0, highlightthickness=0, bg=colors["surface"],
                         cursor="")
        self.colors = colors
        self._icon_cache = icon_cache
        self._font_body = font_body
        self._font_header = font_header
        self._badge_styles = badge_styles or BADGE_STYLES["mist"]
        # 表格内所有文字统一正文大小（用户要求），仅保留风格区分：
        # 路径/数字同号、徽章同号加粗、数字列等宽字体、表头同号加粗
        body_pt = font_body[1] if isinstance(font_body, tuple) else FONT_BODY
        self._font_small = font_body
        self._font_badge = (FONT_FAMILY, body_pt, "bold")
        self._font_mono = (FONT_MONO, body_pt)
        self._font_header = (FONT_FAMILY, body_pt, "bold")
        # 表头高度随正文档位缩放（类常量 42 为兜底，构造时用实例值覆盖）
        self.HEADER_H = max(40, int(body_pt * 1.7))
        self._on_header_click = on_header_click
        self._on_double = on_double
        self._on_right = on_right
        self._on_col_resize = on_col_resize
        self._on_scroll_page = on_scroll_page   # 翻页回调 delta（±1 页），滚轮/PgUp/PgDn 使用
        self._wheel_accum = 0                    # 滚轮累积量（防抖：攒满 120 翻一页）
        self._cols = []            # [(key, width)]
        self._rows = []
        self._selected = []
        self._hover_row = None
        self._hover_col = None
        self._hover_handle = None   # 悬停/拖动中的列宽手柄（列 key）
        self._sort_col = None
        self._sort_asc = True
        self._drag_key = None
        self._drag_start_x = 0
        self._drag_orig_w = 0
        self._drag_pref = set()   # 用户主动拖过宽度的列（fit 时最后压缩）
        self._row_h = 50
        self._labels = {}
        self._scroll_y = 0        # 行区域纵向滚动偏移（页大小自适应下恒为 0，保留以兼容 _hit）
        try:
            import tkinter.font as tkfont
            self._font_body_obj = tkfont.Font(root=master.winfo_toplevel(), font=font_body)
            self._font_small_obj = tkfont.Font(root=master.winfo_toplevel(), font=self._font_small)
        except Exception:
            self._font_body_obj = None
            self._font_small_obj = None
        self.bind("<Configure>", lambda _e: self.redraw())
        self.bind("<Button-1>", self._on_click)
        self.bind("<Double-Button-1>", self._on_double_click)
        self.bind("<Button-3>", self._on_btn3)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<MouseWheel>", self._on_wheel)
        self.bind("<Motion>", self._on_motion)
        self.bind("<Leave>", lambda _e: self._set_hover(None, None))
        # 键盘导航：↑/↓ 移动选中，PgUp/PgDn 翻页（Linux 无 <MouseWheel>，用 Button-4/5）
        self.bind("<Up>", self._on_key_move)
        self.bind("<Down>", self._on_key_move)
        self.bind("<Prior>", lambda _e: self._scroll_page(-1))
        self.bind("<Next>", lambda _e: self._scroll_page(1))
        self.bind("<Button-4>", lambda _e: self._scroll_page(-1))
        self.bind("<Button-5>", lambda _e: self._scroll_page(1))

    # ============ 公共接口 ============

    def set_columns(self, cols):
        self._cols = list(cols)
        self.redraw()

    def column_width(self, key):
        for k, w in self._cols:
            if k == key:
                return w
        return 0

    def set_column_width(self, key, w):
        self._cols = [(k, max(40, w) if k == key else cw) for k, cw in self._cols]
        if self._on_col_resize:
            self._on_col_resize()
        self.redraw()

    def set_sort(self, col, asc):
        self._sort_col = col
        self._sort_asc = asc
        self.redraw()

    def set_rows(self, rows):
        self._rows = rows
        self._selected = []
        self._scroll_y = 0
        self.redraw()

    def set_row_h(self, h):
        self._row_h = max(28, h)
        self.redraw()

    def selected_results(self):
        return [self._rows[i]["result"] for i in self._selected if 0 <= i < len(self._rows)]

    def select_all(self):
        self._selected = list(range(len(self._rows)))
        self.redraw()

    # ============ 布局与命中 ============

    def _fit_cols(self, avail_w=None):
        """返回适配可视宽度的列宽列表（双向）：
        - 总宽超出可视区 → 按序压缩柔性列（path→name→type→modified→size，每列保底）；
        - 总宽不足可视区 → 按偏好比例拉宽柔性列，填满容器。
        刚性列 = 用户主动拖过的列（_drag_pref）+ 当前拖动的列，宽度不被调整；
        拖动任何列都是 1:1 响应且释放后不弹回。
        """
        if avail_w is None:
            avail_w = max(100, self.winfo_width())
        cols = list(self._cols)
        total = sum(w for _, w in cols)
        mins = {"path": 220, "name": 240, "type": 110, "modified": 180, "size": 130}
        rigid = set(getattr(self, "_drag_pref", None) or set())
        if self._drag_key:
            rigid.add(self._drag_key)
        flex_keys = [k for k in ("path", "name", "type", "modified", "size") if k not in rigid]
        flex_idx = [i for i, (k, _w) in enumerate(cols) if k in flex_keys]
        if total > avail_w:
            # 压缩：柔性列按序吸收溢出
            deficit = total - avail_w
            for key in flex_keys:
                if deficit <= 0:
                    break
                for i, (k2, w) in enumerate(cols):
                    if k2 == key:
                        lo = mins.get(key, 40)
                        take = min(max(0, w - lo), deficit)
                        if take > 0:
                            cols[i] = (key, w - take)
                            deficit -= take
                        break
        elif total < avail_w and flex_idx:
            # 拉宽：柔性列按偏好比例吸收剩余空间，填满容器
            extra = avail_w - total
            flex_total = sum(cols[i][1] for i in flex_idx)
            if flex_total > 0:
                for i in flex_idx:
                    cols[i] = (cols[i][0], cols[i][1] + int(extra * cols[i][1] / flex_total))
        return cols

    def _col_layout(self):
        x = 0
        layout = []
        for key, w in self._fit_cols():
            layout.append((key, x, x + w))
            x += w
        return layout

    def _hit(self, x, y):
        if y < self.HEADER_H:
            for key, x0, x1 in self._col_layout():
                if x0 <= x < x1:
                    return ("header", key)
            return None
        row = (y - self.HEADER_H + self._scroll_y) // self._row_h
        if 0 <= row < len(self._rows):
            return ("cell", row)
        return None

    def _resize_handle_at(self, x, y):
        """列宽拖动手柄：表头内列间竖线 ±4px（图标列除外）。"""
        if y >= self.HEADER_H:
            return None
        layout = self._col_layout()
        for i, (key, _x0, x1) in enumerate(layout):
            if i < len(self._cols) - 1 and key != "icon" and abs(x - x1) <= 4:
                return key
        return None

    def _truncate(self, text, max_w, font_obj=None):
        font_obj = font_obj or self._font_body_obj
        if font_obj is None or max_w < 20:
            return text
        if font_obj.measure(text) <= max_w:
            return text
        t = text
        while t and font_obj.measure(t + "…") > max_w:
            t = t[:-1]
        return (t + "…") if t else ""

    def _path_middle_ellipsis(self, path, max_w):
        """路径过长时中间省略，保留首尾目录信息。"""
        f = self._font_small_obj
        if f is None or max_w < 40:
            return path
        if f.measure(path) <= max_w:
            return path
        sep = "\\" if "\\" in path else "/"
        parts = path.split(sep)
        if len(parts) <= 2:
            return self._truncate(path, max_w, f)
        head, tail = parts[0], parts[-1]
        mid = parts[1:-1]
        while mid:
            cand = head + sep + "…" + sep + sep.join(mid[1:] + [tail]) if len(mid) > 1 else head + sep + "…" + sep + tail
            if f.measure(cand) <= max_w:
                return cand
            mid = mid[1:]
        cand = head + sep + "…" + sep + tail
        return cand if f.measure(cand) <= max_w else self._truncate(path, max_w, f)

    # ============ 事件 ============

    def _on_click(self, event):
        self.focus_set()
        handle = self._resize_handle_at(event.x, event.y)
        if handle:
            self._drag_key = handle
            self._drag_start_x = event.x
            # 基准用 fit 后的实际宽度：压缩态下拖动手柄位置是压缩后的，直接
            # 以实际宽度 + 位移写入偏好宽度，拖动才能"所见即所得"
            self._drag_orig_w = self._fitted_width(handle)
            return
        hit = self._hit(event.x, event.y)
        if hit is None:
            return
        kind, val = hit
        if kind == "header":
            if self._on_header_click:
                self._on_header_click(val)
            return
        row = val
        ctrl = bool(event.state & 0x0004)
        shift = bool(event.state & 0x0001)
        if ctrl:
            if row in self._selected:
                self._selected.remove(row)
            else:
                self._selected.append(row)
        elif shift and self._selected:
            anchor = self._selected[-1]
            lo, hi = sorted((anchor, row))
            self._selected = list(range(lo, hi + 1))
        else:
            self._selected = [row]
        self.redraw()

    def _on_double_click(self, event):
        hit = self._hit(event.x, event.y)
        if hit and hit[0] == "cell" and self._on_double:
            self._on_double(hit[1])

    def _on_btn3(self, event):
        hit = self._hit(event.x, event.y)
        if hit and hit[0] == "cell":
            row = hit[1]
            if row not in self._selected:
                self._selected = [row]
                self.redraw()
            if self._on_right:
                self._on_right(event, row)

    def _fitted_width(self, key):
        """取某列 fit 后的实际显示宽度（无压缩时等于偏好宽度）。"""
        for k, x0, x1 in self._col_layout():
            if k == key:
                return x1 - x0
        return self.column_width(key)

    def _on_drag(self, event):
        if self._drag_key:
            delta = event.x - self._drag_start_x
            self.set_column_width(self._drag_key, self._drag_orig_w + delta)

    def _on_release(self, event):
        """释放拖动：记录该列为手动布局列（fit 最后压缩），并刷新手柄高亮。"""
        if self._drag_key:
            self._drag_pref.add(self._drag_key)
            self._drag_key = None
            self.redraw()

    def _on_wheel(self, event):
        """滚轮翻页（防抖）：累积 delta，攒满一格（120）翻一页。

        一页正好铺满可视区，滚轮语义 = 翻页而非滚动行；
        快速连续滚动时累积量自然合并，不会一格格顿挫。
        """
        delta = event.delta
        if sys.platform == "darwin":
            delta = int(delta * 120)   # macOS 每格 delta=±1
        self._wheel_accum += delta
        if abs(self._wheel_accum) >= 120:
            pages = int(self._wheel_accum // 120)
            self._wheel_accum %= 120
            self._scroll_page(-pages)   # 向上滚（delta>0）→ 上一页

    def _scroll_page(self, delta: int):
        """翻页：delta=±1（负=上一页）。翻页后把焦点留在表格，继续可用键盘。"""
        if self._on_scroll_page:
            self._on_scroll_page(delta)
        self.focus_set()

    def _on_key_move(self, event):
        """↑/↓ 移动选中行（单选，clamp 到行范围）。"""
        if not self._rows:
            return "break"
        last = self._selected[-1] if self._selected else None
        if event.keysym == "Up":
            target = (last - 1) if last is not None else 0
        else:
            target = (last + 1) if last is not None else 0
        target = max(0, min(len(self._rows) - 1, target))
        if self._selected != [target]:
            self._selected = [target]
            self._hover_row = target
            self.redraw()
        return "break"

    def _on_motion(self, event):
        handle = self._resize_handle_at(event.x, event.y)
        if handle:
            self.configure(cursor="sb_h_double_arrow")
        else:
            self.configure(cursor="")
        if handle != self._hover_handle:
            self._hover_handle = handle
            self.redraw()
        hit = self._hit(event.x, event.y)
        if hit and hit[0] == "cell":
            self._set_hover(hit[1], None)
        else:
            self._set_hover(None, hit[1] if hit and hit[0] == "header" else None)

    def _set_hover(self, row, col):
        if row != self._hover_row or col != self._hover_col:
            self._hover_row = row
            self._hover_col = col
            self.redraw()

    # ============ 绘制 ============

    def _draw_badge(self, cx, cy, kind, text):
        """绘制类型徽章：圆角胶囊底 + 类型色文字。"""
        fg, bg = self._badge_styles.get(kind, self._badge_styles["file"])
        try:
            import tkinter.font as tkfont
            f = tkfont.Font(root=self.winfo_toplevel(), font=self._font_badge)
            tw = f.measure(text)
        except Exception:
            tw = len(text) * 12
        bw = tw + 18
        # 徽章高度随徽章字号缩放（胶囊需容纳同号加粗文字）
        bh = max(24, int(self._font_badge[1] * 1.9))
        _rounded_rect(self, cx - bw / 2, cy - bh / 2, cx + bw / 2, cy + bh / 2, bh / 2 - 1,
                      fill=bg, outline="", width=0)
        self.create_text(cx, cy, text=text, fill=fg, font=self._font_badge)

    def redraw(self):
        try:
            self.winfo_exists()
        except Exception:
            return
        self.delete("all")
        w = max(2, self.winfo_width())
        h = max(2, self.winfo_height())
        c = self.colors
        layout = self._col_layout()
        self.create_rectangle(0, 0, w, h, fill=c["surface"], outline="")

        pad_l = 18   # 单元格左内边距
        # ---- 数据行（含纵向滚动偏移，表头固定）----
        row_h = self._row_h
        sy = self._scroll_y
        for i, row in enumerate(self._rows):
            y0 = self.HEADER_H + i * row_h - sy
            if y0 + row_h <= self.HEADER_H:
                continue
            if y0 >= h:
                break
            y1 = y0 + row_h
            selected = i in self._selected
            if selected:
                bg = c["selected"]
            elif i == self._hover_row:
                bg = c["hover"]
            else:
                bg = c["surface"]
            self.create_rectangle(0, y0, w, y1, fill=bg, outline="")
            if selected:
                # 左侧 accent 圆头指示条
                inset = row_h * 0.18
                _rounded_rect(self, 0, y0 + inset, 6, y1 - inset, 3,
                              fill=c["accent"], outline="", width=0)
            # 行分隔线（最后一行也画，视觉收敛）
            self.create_line(0, y1, w, y1, fill=c["row_line"])
            vals = row.get("values", ())
            icon = row.get("icon")
            result = row.get("result", {})
            cy = (y0 + y1) / 2
            for j, (key, x0, x1) in enumerate(layout):
                if j == 0:
                    if icon:
                        self.create_image((x0 + x1) / 2, cy, image=icon)
                    continue
                idx = j - 1
                if key == "type":
                    kind = badge_kind_for(result.get("name", ""), bool(result.get("is_dir")))
                    text = str(vals[idx]) if idx < len(vals) else ""
                    if text:
                        self._draw_badge((x0 + x1) / 2, cy, kind, text)
                    continue
                if key == "name":
                    font, fg, fobj = self._font_body, (c["sel_text"] if selected else c["text"]), self._font_body_obj
                    anchor, tx = "w", x0 + pad_l
                elif key == "path":
                    font, fg, fobj = self._font_small, c["muted"], self._font_small_obj
                    anchor, tx = "w", x0 + pad_l
                elif key == "size":
                    font, fg, fobj = self._font_mono, c["muted"], self._font_small_obj
                    anchor, tx = "e", x1 - pad_l
                else:  # modified
                    font, fg, fobj = self._font_mono, c["muted"], self._font_small_obj
                    anchor, tx = "e", x1 - pad_l
                if idx < len(vals):
                    if key == "path":
                        text = self._path_middle_ellipsis(str(vals[idx]), max(10, x1 - x0 - pad_l * 2))
                    else:
                        text = self._truncate(str(vals[idx]), max(10, x1 - x0 - pad_l * 2), fobj)
                else:
                    text = ""
                if not text:
                    continue
                self.create_text(tx, cy, text=text, anchor=anchor, fill=fg, font=font)

        # ---- 表头 ----
        self.create_rectangle(0, 0, w, self.HEADER_H, fill=c["surface"], outline="")
        self.create_line(0, self.HEADER_H, w, self.HEADER_H, fill=c["border"])
        header_align = {"name": "w", "path": "w", "type": "center", "size": "e", "modified": "e"}
        for j, (key, x0, x1) in enumerate(layout):
            if j == 0:
                continue
            # 列间拖动手柄线：表头内中段竖线，悬停/拖动时 accent 加粗高亮
            if j < len(layout) - 1:
                active = (self._hover_handle == key) or (self._drag_key == key)
                self.create_line(x1, self.HEADER_H * 0.30, x1, self.HEADER_H * 0.80,
                                 fill=c["accent"] if active else c["border"],
                                 width=2 if active else 1)
            label = self._labels.get(key, key)
            anchor = header_align.get(key, "w")
            if anchor == "w":
                tx = x0 + pad_l
            elif anchor == "e":
                tx = x1 - pad_l
            else:
                tx = (x0 + x1) / 2
            if key == self._sort_col:
                fg = c["accent"]
            elif key == self._hover_col:
                fg = c["text"]
            else:
                fg = c["muted_2"]
            self.create_text(tx, self.HEADER_H / 2, text=label, anchor=anchor, fill=fg,
                             font=self._font_header)
            # 排序小三角（画在标签右侧）
            if key == self._sort_col:
                try:
                    import tkinter.font as tkfont
                    hf = tkfont.Font(root=self.winfo_toplevel(), font=self._font_header)
                    lw = hf.measure(label)
                except Exception:
                    lw = len(label) * 9
                tri_cx = tx + lw + 12 if anchor == "w" else (tx - lw - 12 if anchor == "e" else tx + lw / 2 + 10)
                tri_cy = self.HEADER_H / 2
                s = 5
                if self._sort_asc:
                    pts = (tri_cx - s, tri_cy + s * 0.6, tri_cx + s, tri_cy + s * 0.6, tri_cx, tri_cy - s * 0.8)
                else:
                    pts = (tri_cx - s, tri_cy - s * 0.6, tri_cx + s, tri_cy - s * 0.6, tri_cx, tri_cy + s * 0.8)
                self.create_polygon(pts, fill=c["accent"], outline="")
        self.configure(scrollregion=(0, 0, w, h))


# ================================================================
#  FileSearcherApp — GUI 主应用
# ================================================================
class FileSearcherApp:
    """全盘文件搜索 GUI 应用。使用 ttk.Treeview 展示文件列表。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("File Searcher — 全盘文件搜索")
        self.root.minsize(960, 620)

        self._engine_cancel = False
        self._index_running = False
        self._search_timer = None
        self._suppress_search_trace = False
        self._results: list[dict] = []
        self._sort_col = "size"
        self._sort_asc = False
        self._last_query = ""
        self._total_results = 0
        self._loading_more = False
        self._count_cache_key = None      # 计数缓存键 (query, 排序, filters, db_mtime)
        self._count_cache_val = 0
        self._query_seq = 0               # 查询序号（丢弃过期结果）
        self._query_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="db-query")
        self._tray_index_after_id = None
        self._settings = IndexEngine.load_settings()
        self._theme_name = self._resolve_theme(self._settings.get("theme", "mist"))
        self.colors = THEMES[self._theme_name]

        # 缩放体系：Tk 8.6 在 Windows 上按 96 DPI 布局，tk scaling 对字体渲染无效且会污染
        # winfo_fpixels，因此用原生 API GetDpiForSystem 取真实系统 DPI；
        # _dpi_scale（真实 DPI/96）× _font_scale（用户字号档位）放大所有 pt 字号与像素尺寸；
        # ui_scale 恒 1.0。主字号（text_pt = 表格正文基准 pt，10~40 可调）全局生效：
        # 所有页面（主界面/设置/弹窗/右键菜单）经 _f/_s 联动，其余文字按固定比例自适应
        try:
            _dpi = float(ctypes.windll.user32.GetDpiForSystem())
            if _dpi < 72:
                _dpi = 96.0
            # 部分 DPI 缩放：全量 dpi/96（如 250% → 2.5×）在 4K 高缩放屏上字号/行高
            # 会远超系统其他应用（实测观感"大得不像话"），取平方根折中
            # （250% → 1.581×、200% → 1.414×、150% → 1.225×）；
            # 96 DPI（100%）下恒为 1.0，低缩放屏行为不变。
            self._dpi_scale = max(1.0, (_dpi / 96.0) ** 0.5)
        except Exception:
            self._dpi_scale = 1.0
        # 正文目标 pt：BODY 渲染 = base × dpi_scale × font_scale = text_pt × dpi_scale
        # （Tk pt→px ≈ ×4/3，240 DPI 下 14/16/18pt 档分别约 30/34/38px 物理，
        # 96 DPI 下约 19/21/24px）。dpi_scale 取 sqrt(dpi/96) 折中（见上），
        # 高缩放屏不至于超过系统其他应用太多。
        # 旧档位（10/12/14，上一版校准）统一迁移到 14。
        try:
            self._text_pt = int(self._settings.get("text_pt", 16))
        except Exception:
            self._text_pt = 16
        if not (10 <= self._text_pt <= 40):
            self._text_pt = 16
            if "text_pt" in self._settings:
                self._settings["text_pt"] = 16
                IndexEngine.save_settings(self._settings)
        # font_scale = 用户档位相对基础正文的比例（不含 dpi_scale：dpi 因子由
        # _f/_s 里的 × dpi_scale 承担。曾误把 dpi_scale 放进分母导致其被完全
        # 抵消——程序是 per-monitor DPI aware，Windows 不会位图拉伸，高 DPI
        # 屏上字号始终只有系统正常水平的 1/dpi_scale，修复后 DPI 才真正生效）
        self._font_scale = self._text_pt / FONT_BODY
        try:
            self.root.tk.call("tk", "scaling", 4 / 3)  # 固定为标准 96dpi 行为
        except Exception:
            pass
        self._window_scale = 1.0
        self.ui_scale = 1.0

        # 原生标题栏（无边框方案已废弃）：铺满工作区、禁用最大化
        self._frameless = False
        self._normal_rect = None
        self._wnd_proc_holder = None   # 保活 WndProc 回调（防 GC）
        self._orig_wndproc = None

        # Tkinter 回调异常统一写日志（windowed 打包下 stderr 不可见，异常全丢）
        self.root.report_callback_exception = self._report_callback_exception

        IndexEngine.ensure_indexes()   # 老库幂等补齐 size/modified 排序索引 + FTS 空表
        # 老库升级：后台回填 FTS 表（回填完成前搜索自动走 LIKE，不卡启动）
        if IndexEngine.index_exists():
            threading.Thread(target=IndexEngine.ensure_fts_backfill,
                             daemon=True, name="fts-backfill").start()
        self._build_ui()
        self._load_layout()
        self._setup_frameless()
        self._setup_tray()

        # 分页状态：页大小按可视行数自适应
        self._page = 1
        self._page_size = self._compute_page_size()
        self._view_resize_timer = None
        self._layout_save_timer = None

        self.root.after(100, self._defer_initial_load)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        # 不再绑 <Unmap>：它在子控件 unmap 时也冒泡触发，会误把窗口收进托盘。
        # 「✕」进托盘走 _on_close；「—」留任务栏走 _minimize_window 的 iconify。
        self.root.bind("<Configure>", self._on_view_resize, add="+")
        self.root.after(1500, self._log_diag)
        if self._settings.get("auto_index_on_start") and IndexEngine.index_exists():
            self.root.after(500, self._do_index_silent)

    def _defer_initial_load(self):
        """延迟到结果表映射完成后加载首页。

        启动时表格尚未布局（winfo_height=1），_compute_page_size 会兜底成 5 行/页，
        直接查询会先拉一页 5 行、窗口显示后再重查一次。这里轮询等表格映射，
        映射完成后用一次性标志保证至少触发一次加载（_load_all）。

        注意：不能依赖窗口 resize 事件兜底——若本方法的映射检测先于 resize
        事件完成，会先把 _page_size 更新为实际值，随后 resize 发现页大小未变
        而跳过查询 → 首次查询永不执行、启动后列表空白。必须由本方法保证查询。
        """
        try:
            if self.root.state() in ("withdrawn", "iconic"):
                self.root.after(100, self._defer_initial_load)
                return
            if self.table.winfo_height() <= 1:
                self.root.after(100, self._defer_initial_load)
                return
        except Exception:
            self.root.after(100, self._defer_initial_load)
            return
        ps = self._compute_page_size()
        if ps != self._page_size:
            self._page_size = ps
        if not getattr(self, "_initial_query_done", False):
            self._initial_query_done = True
            self._load_all()
        else:
            self._update_result_status()

    def _report_callback_exception(self, exc, val, tb):
        """Tkinter 回调异常统一写 INDEX_DIR/debug.log（windowed 打包下 stderr 不可见）。"""
        try:
            import traceback
            _diag("[tk-error] " + "".join(traceback.format_exception(exc, val, tb)).rstrip())
        except Exception:
            pass

    def _log_diag(self):
        """启动 1.5s 后记录缩放关键参数到 INDEX_DIR/debug.log，用于诊断字号问题。"""
        try:
            import tkinter.font as tkfont
            try:
                real_dpi = float(ctypes.windll.user32.GetDpiForSystem())
            except Exception:
                real_dpi = 0.0
            f = tkfont.Font(root=self.root, family=FONT_FAMILY,
                            size=max(8, int(FONT_BODY * self.ui_scale * self._dpi_scale * self._font_scale)))
            line = (f"[diag] dpi_scale={self._dpi_scale:.3f} ui_scale={self.ui_scale:.3f} "
                    f"win_w={self.root.winfo_width()} sys_dpi={real_dpi:.0f} state={self.root.state()} "
                    f"font_body_pt={max(8, int(FONT_BODY * self.ui_scale * self._dpi_scale * self._font_scale))} "
                    f"real_px={f.metrics('linespace')}\n")
            _diag(line)
        except Exception:
            pass

    def _build_ui(self):
        """构建全部主界面控件（支持按新缩放重建）。"""
        self._configure_theme()
        self._icon_cache = WindowsShellIconCache(self.root, self.colors["surface"], size=self._s(22))
        # 用系统原生标题栏（含最小化/最大化/关闭按钮与任务栏图标），不再自绘标题栏
        self._build_toolbar()
        self._build_tree()
        self._build_statusbar()
        self._setup_drag_drop()
        self._build_context_menu()
        self._bind_shortcuts()
        self._update_index_button_text()

    # ---- 缩放工具 ----

    def _s(self, px: int) -> int:
        """像素尺寸 × ui_scale × DPI 比例 × 字号系数。"""
        return max(1, int(px * self.ui_scale * self._dpi_scale * self._font_scale))

    def _f(self, base_pt: int, weight: str = "normal"):
        """字号 pt × ui_scale × DPI 比例 × 字号系数的字体元组。"""
        return (FONT_FAMILY, max(8, int(base_pt * self.ui_scale * self._dpi_scale * self._font_scale)), weight)

    def _rebuild_ui(self):
        """按新的 ui_scale 重建全部主界面，保留搜索词、筛选与结果。"""
        if self._search_timer is not None:
            try:
                self.root.after_cancel(self._search_timer)
            except Exception:
                pass
            self._search_timer = None
        try:
            query = self._search_text()
        except Exception:
            query = ""
        saved = {"query": query}
        for w in self.root.winfo_children():
            w.destroy()
        self._build_ui()
        self._set_search_value(saved["query"] if saved["query"] else "")
        self._refresh_tree()
        self._update_result_status()

    def _apply_theme(self):
        """主题切换立即生效：换配色并重建主界面（保留搜索词与结果）。"""
        self._theme_name = self._resolve_theme(self._settings.get("theme", "mist"))
        self.colors = THEMES[self._theme_name]
        self._configure_theme()
        self._rebuild_ui()

    def _apply_text_scale(self):
        """文字大小档位切换立即生效：重算全局缩放系数并重建全部界面。"""
        try:
            self._text_pt = int(self._settings.get("text_pt", 16))
        except Exception:
            self._text_pt = 16
        if not (10 <= self._text_pt <= 40):
            self._text_pt = 16
        self._font_scale = self._text_pt / FONT_BODY
        self._configure_theme()
        self._rebuild_ui()
        # 重建后表格尚未布局（winfo_height=1），页大小必须等布局完成后重算，
        # 否则 _page_size 停留旧行高算出的值 → 切档位后可见行数不随字号变化
        self.root.after(80, self._apply_view_resize)

    def _resolve_theme(self, theme: str) -> str:
        # 无「跟随系统」：未知/旧值（含 system）一律落到默认雾沙
        return theme if theme in THEMES else "mist"

    def _configure_theme(self):
        c = self.colors
        self.root.configure(bg=c["bg"])
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", background=c["bg"], foreground=c["text"], font=self._f(FONT_BODY))
        style.configure("TFrame", background=c["bg"])
        style.configure("Surface.TFrame", background=c["surface"])
        style.configure("TLabel", background=c["bg"], foreground=c["text"], font=self._f(FONT_BODY))
        style.configure("Muted.TLabel", background=c["bg"], foreground=c["muted"], font=self._f(FONT_SMALL))
        # 排除列表（设置页）
        style.configure("Ex.Treeview", background=c["surface"], fieldbackground=c["surface"],
                        foreground=c["text"], borderwidth=0, relief="flat",
                        rowheight=self._s(38), font=self._f(FONT_SMALL))
        style.map("Ex.Treeview", background=[("selected", c["selected"])],
                  foreground=[("selected", c["sel_text"])])
        style.configure("Ex.Treeview.Heading", background=c["surface_alt"], foreground=c["muted"],
                        font=self._f(FONT_MICRO, "bold"), relief="flat", padding=(8, 7))
        style.map("Ex.Treeview.Heading", background=[("active", c["surface_3"])])
        style.configure("Vertical.TScrollbar", background=c["surface_3"], troughcolor=c["surface"],
                        bordercolor=c["surface"], arrowcolor=c["muted_2"], width=self._s(12))
        style.configure("TProgressbar", troughcolor=c["surface_alt"], background=c["accent"],
                        bordercolor=c["surface_alt"])
        style.configure("TSeparator", background=c["border"])
        # 主字号滑动条（设置页）：轨道用 surface_alt，滑块用 accent
        style.configure("TextPt.Horizontal.TScale", background=c["surface"],
                        troughcolor=c["surface_alt"], bordercolor=c["surface"],
                        lightcolor=c["accent"], darkcolor=c["accent"])

    # ================================================================
    #  自绘标题栏 + 无边框窗口（Windows）
    # ================================================================

    def _setup_frameless(self):
        """使用系统原生标题栏：任务栏图标/最小化/关闭走系统；启动即铺满工作区。

        启动铺满工作区（保留任务栏）= _maximize_to_workarea()。
        最大化「不可用」：窗口保持可缩放（铺满生效的前提，resizable(False) 会钳制尺寸），
        通过 ① 去掉 WS_MAXIMIZEBOX 样式位（最大化钮置灰）+ ② 拦截双击标题栏最大化
        + ③ _maximize_to_workarea 里程序内禁缩放拖拽，三路共同保证不能真正放大窗口。
        托盘行为：✕ 收进托盘，— 最小化到任务栏。
        """
        self._frameless = False
        if sys.platform != "win32":
            try:
                self.root.state("zoomed")
            except Exception:
                pass
            return
        try:
            self.root.deiconify()
            self._maximize_to_workarea()   # 启动即铺满工作区（不依赖 zoomed）
            # 去掉最大化样式位（最大化钮置灰；去掉后即使用户点也无最大化能力）
            try:
                hwnd = self.root.winfo_id()
                user32 = ctypes.windll.user32
                GWL_STYLE = -16
                WS_MAXIMIZEBOX = 0x00010000
                SWP_FRAMECHANGED, SWP_NOMOVE, SWP_NOSIZE = 0x0020, 0x0002, 0x0001
                SWP_NOZORDER, SWP_NOOWNERZORDER = 0x0004, 0x0200
                cur = user32.GetWindowLongW(hwnd, GWL_STYLE)
                user32.SetWindowLongW(hwnd, GWL_STYLE, cur & ~WS_MAXIMIZEBOX)
                user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                    SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE
                                    | SWP_NOZORDER | SWP_NOOWNERZORDER)
                after = user32.GetWindowLongW(hwnd, GWL_STYLE)
                _diag(f"[frameless] style before={cur:#010x} after={after:#010x} maximbox_removed={not bool(after & WS_MAXIMIZEBOX)}")
            except Exception:
                pass
            # 子类化窗口过程：系统层拦截最大化（最大化钮/双击标题栏/系统菜单都无效）
            self._install_no_maximize_wndproc()
            _diag(f"[frameless] native titlebar, maximized-to-workarea, no-maximize-box, state={self.root.state()}")
        except Exception as e:
            _diag(f"[frameless] FAIL {e!r}")

    def _install_no_maximize_wndproc(self):
        """子类化窗口过程，系统层拦截最大化：最大化钮/双击标题栏/系统菜单/Aero 吸附全部失效。

        Tk 的原生标题栏对按钮控制很弱（去掉 WS_MAXIMIZEBOX 也不置灰按钮）。这条路直接在
        WndProc 里吞掉 WM_SYSCOMMAND(SC_MAXIMIZE) 和 WM_NCLBUTTONDBLCLK(HTCAPTION)，
        最大化这个操作根本到不了窗口 → 按钮点了没反应、双击标题栏不放大。窗口仍可缩放
        （铺满工作区生效的前提），拖边角由 _maximize_to_workarea 的 maxsize 钳在工作区。
        """
        if sys.platform != "win32":
            return
        try:
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            if ctypes.sizeof(ctypes.c_void_p) == 8:
                GetW, SetW = user32.GetWindowLongPtrW, user32.SetWindowLongPtrW
            else:
                GetW, SetW = user32.GetWindowLongW, user32.SetWindowLongW
            SetW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
            SetW.restype = ctypes.c_ssize_t
            GetW.argtypes = [wintypes.HWND, ctypes.c_int]
            GetW.restype = ctypes.c_ssize_t
            user32.CallWindowProcW.argtypes = [ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT,
                                               ctypes.c_ssize_t, ctypes.c_ssize_t]
            user32.CallWindowProcW.restype = ctypes.c_ssize_t

            GWLP_WNDPROC = -4
            WM_SYSCOMMAND, WM_NCLBUTTONDBLCLK, WM_GETMINMAXINFO = 0x0112, 0x00A3, 0x0024
            SC_MAXIMIZE, SC_RESTORE = 0xF030, 0xF120
            HTCAPTION = 2
            app = self

            def wnd_proc(hwnd, msg, wparam, lparam):
                # 吞掉最大化/还原命令：窗口锁定在最大化，点「还原」钮/系统菜单/任务栏右键都无效
                if msg == WM_SYSCOMMAND:
                    cmd = wparam & 0xFFF0
                    if cmd in (SC_MAXIMIZE, SC_RESTORE):
                        return 0
                # 吞掉双击标题栏（默认会最大化/还原）
                if msg == WM_NCLBUTTONDBLCLK and wparam == HTCAPTION:
                    return 0
                return user32.CallWindowProcW(app._orig_wndproc, hwnd, msg, wparam, lparam)

            WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT,
                                         ctypes.c_ssize_t, ctypes.c_ssize_t)
            proc = WNDPROC(wnd_proc)
            hwnd = self.root.winfo_id()
            self._orig_wndproc = GetW(hwnd, GWLP_WNDPROC)
            if not self._orig_wndproc:
                return
            self._wnd_proc_holder = proc   # 保活回调，防 GC
            SetW(hwnd, GWLP_WNDPROC, ctypes.cast(proc, ctypes.c_void_p).value or proc)
            _diag("[frameless] no-maximize wndproc installed")
        except Exception as e:
            _diag(f"[frameless] wndproc FAIL {e!r}")

    def _maximize_to_workarea(self):
        """启动即「真正最大化」：Tk 原生 zoomed。

        之前手动 SetWindowPos/ShowWindow(SW_MAXIMIZE) 都被 Tk 布局用「内容请求尺寸」覆盖
        （实测 SetWindowPos 2560x1368 → 缩回 640x426；ShowWindow → 缩回 960x620，state 仍
        normal）。Tk 原生 `state('zoomed')` 会让 Tk 进入 zoomed 状态，Tk 自己维护窗口大小
        （=工作区），不再被内容尺寸覆盖。原 overrideredirect 时代 zoomed 失效，原生标题栏
        下无此问题。窗口进入 zoomed 后右上角「还原」钮被 WndProc 拦 SC_RESTORE → 锁定最大化。
        """
        if sys.platform != "win32":
            return
        try:
            from ctypes import wintypes
            self.root.state("zoomed")
            self.root.update_idletasks()
            self._normal_rect = None
            _diag(f"[maximize] state('zoomed') -> state={self.root.state()}")
            # idle 后回读实际窗口矩形确认铺满
            def _verify():
                try:
                    hwnd = self.root.winfo_id()
                    r = wintypes.RECT()
                    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r))
                    _diag(f"[maximize] actual rect=({r.left},{r.top}) "
                          f"{r.right - r.left}x{r.bottom - r.top} state={self.root.state()}")
                except Exception:
                    pass
            self.root.after_idle(_verify)
        except Exception as e:
            _diag(f"[maximize] zoomed FAIL {e!r}")


    # ================================================================
    #  UI 构建
    # ================================================================

    def _build_toolbar(self):
        """凝脂风搜索区：全宽大搜索框（视觉主角，无其他干扰元素）。"""
        c = self.colors
        zone = tk.Frame(self.root, bg=c["bg"])
        zone.pack(fill=tk.X, padx=self._s(24), pady=(self._s(18), 0))
        self._header = zone

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._on_search_changed())
        row_h = self._s(58)
        self.search_box = RoundedSearchBox(zone, self.search_var, c, self._clear_search,
                                           height=row_h, font_size=self._f(FONT_INPUT)[1],
                                           placeholder="搜索全盘文件…")
        self.search_box.pack(fill=tk.X, padx=self._s(24))
        self.search_entry = self.search_box.entry
        self.search_entry.bind("<Return>", self._do_search)

    def _set_search_value(self, value: str):
        """更新搜索变量但不触发查询，用于占位符和程序化清空。"""
        self._suppress_search_trace = True
        try:
            self.search_var.set(value)
        finally:
            self._suppress_search_trace = False

    def _search_text(self):
        return self.search_var.get().strip()

    def _build_tree(self):
        """构建凝脂风结果区：纸面卡片容器 + 呼吸式表格 + 底部分页栏（左计数/右页码）。"""
        c = self.colors
        outer = tk.Frame(self.root, bg=c["bg"])
        outer.pack(fill=tk.BOTH, expand=True, padx=self._s(28), pady=(self._s(10), self._s(10)))
        self.result_canvas = tk.Canvas(outer, bd=0, highlightthickness=0, bg=c["bg"])
        self.result_canvas.pack(fill=tk.BOTH, expand=True)
        self.result_surface = tk.Frame(self.result_canvas, bg=c["surface"])
        self._result_window = self.result_canvas.create_window(1, 1, window=self.result_surface, anchor=tk.NW)
        self.result_canvas.bind("<Configure>", self._layout_result_container)

        # 呼吸式表格（保留排序/列宽拖动/徽章）
        self.table = FileTable(self.result_surface, c, self._icon_cache,
                               self._f(FONT_BODY), self._f(FONT_HEADER),
                               on_header_click=self._sort_by,
                               on_double=self._on_double_click,
                               on_right=self._on_right_click,
                               on_col_resize=self._on_col_resize,
                               on_scroll_page=self._goto_page_relative,
                               badge_styles=BADGE_STYLES[self._theme_name])
        self.table.pack(fill=tk.BOTH, expand=True, padx=1, pady=(1, 0))
        self.table._labels = {"name": "文件名", "path": "路径", "type": "类型",
                              "size": "大小", "modified": "修改时间"}
        self.table._row_h = self._s(ROW_HEIGHT)
        self._default_cols = [("name", self._s(340)), ("path", self._s(620)),
                              ("type", self._s(130)), ("size", self._s(140)),
                              ("modified", self._s(210))]
        self.table._cols = [("icon", self._s(46))] + list(self._default_cols)

        # 底部分页栏：左「共 N 个结果」、大号数字页码居中、翻页钮钉在两端的内侧固定区
        pager = tk.Frame(self.result_surface, bg=c["surface"], height=self._s(52),
                         highlightthickness=1, highlightbackground=c["row_line"])
        pager.pack(fill=tk.X, padx=1, pady=(0, 1))
        pager.pack_propagate(False)
        self.pager_total_var = tk.StringVar(value="共 0 个结果")
        total_lbl = tk.Label(pager, textvariable=self.pager_total_var, bg=c["surface"],
                             fg=c["muted"], font=self._f(FONT_SMALL))
        total_lbl.place(relx=0.0, rely=0.5, anchor=tk.W, x=self._s(16))
        # 数字页码容器（整体居中，动态重建）
        self._pager_btns_frame = tk.Frame(pager, bg=c["surface"])
        self._pager_btns_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        # 「上一页/下一页」固定容器：锚定中轴两侧，内容向内对齐 → 位置永不随页码数量变化
        self._pager_prev_frame = tk.Frame(pager, bg=c["surface"])
        self._pager_prev_frame.place(relx=0.5, rely=0.5, anchor=tk.E, x=-self._s(200))
        self._pager_next_frame = tk.Frame(pager, bg=c["surface"])
        self._pager_next_frame.place(relx=0.5, rely=0.5, anchor=tk.W, x=self._s(200))

        # 空状态：同心圆 + 放大镜图标 + 主文案 + 副文案
        empty_frame = tk.Frame(self.result_surface, bg=c["surface"])
        icon_size = self._s(64)
        self._empty_icon = tk.Canvas(empty_frame, width=icon_size, height=icon_size, bd=0,
                                     highlightthickness=0, bg=c["surface"])
        self._empty_icon.pack()
        u = icon_size / 64.0
        self._empty_icon.create_oval(2 * u, 2 * u, 62 * u, 62 * u, outline=c["surface_3"],
                                     width=max(1.5, 2 * u))
        self._empty_icon.create_oval(10 * u, 10 * u, 54 * u, 54 * u, outline=c["border_strong"],
                                     width=max(1, 1.5 * u), dash=(3, 5))
        self._empty_icon.create_oval(19 * u, 19 * u, 37 * u, 37 * u, outline=c["muted_2"],
                                     width=max(2, 2.5 * u))
        self._empty_icon.create_line(35 * u, 35 * u, 44 * u, 44 * u, fill=c["muted_2"],
                                     width=max(2.5, 3 * u), capstyle=tk.ROUND)
        self.empty_title_var = tk.StringVar(value="没有匹配的结果")
        self.empty_sub_var = tk.StringVar(value="换个关键词试试，或检查拼写")
        tk.Label(empty_frame, textvariable=self.empty_title_var, bg=c["surface"], fg=c["muted"],
                 font=self._f(FONT_LG)).pack(pady=(self._s(14), self._s(4)))
        tk.Label(empty_frame, textvariable=self.empty_sub_var, bg=c["surface"], fg=c["muted_2"],
                 font=self._f(FONT_SMALL)).pack()
        self.empty_state = empty_frame

    def _show_empty_state(self, show: bool):
        if show:
            self.empty_state.place(in_=self.table, relx=0.5, rely=0.45, anchor=tk.CENTER)
        else:
            self.empty_state.place_forget()

    # ---- 分页 ----

    def _compute_page_size(self) -> int:
        """页大小 = 表头下方可视行数，一页正好铺满不滚动。"""
        try:
            h = self.table.winfo_height() - self.table.HEADER_H
            return max(5, h // self.table._row_h)
        except Exception:
            return 20

    def _total_pages(self) -> int:
        if self._total_results <= 0:
            return 1
        return max(1, (self._total_results + self._page_size - 1) // self._page_size)

    def _goto_page(self, page: int):
        total_pages = self._total_pages()
        page = max(1, min(total_pages, page))
        if page != self._page:
            self._page = page
            self.table._scroll_y = 0   # 翻页回到顶部
            self._run_query(self._search_text())
        else:
            self._update_result_status()

    def _goto_page_relative(self, delta: int):
        self._goto_page(self._page + delta)

    def _on_jump_to_page(self, _event=None):
        """内联跳页器回车：解析输入并跳转，非法/超范围提示并还原为当前页。"""
        total_pages = self._total_pages()
        raw = self._pager_jump_var.get().strip()
        try:
            page = int(raw)
        except (TypeError, ValueError):
            self.status_var.set("请输入有效的页码数字")
            self._pager_jump_var.set(str(self._page))
            return
        if page < 1 or page > total_pages:
            self.status_var.set(f"页码超出范围（共 {total_pages} 页）")
            self._pager_jump_var.set(str(self._page))
            return
        # 跳完把焦点还给表格，方便继续用快捷键翻页
        self._goto_page(page)
        self.table.focus_set()

    def _on_col_resize(self):
        """列宽拖动后防抖保存布局。"""
        if getattr(self, "_layout_save_timer", None) is not None:
            try:
                self.root.after_cancel(self._layout_save_timer)
            except Exception:
                pass
        self._layout_save_timer = self.root.after(500, self._save_layout)

    def _on_view_resize(self, event=None):
        """窗口尺寸变化：重算页大小（自适应可视行数），刷新当前页。

        <Configure> 在窗口移动时也触发（width/height 不变），
        尺寸无变化直接跳过，避免移动窗口白跑一轮页大小计算。
        """
        if event is not None and getattr(event, "widget", None) is not self.root:
            return
        if event is not None:
            size = (event.width, event.height)
            if getattr(self, "_last_view_size", None) == size:
                return
            self._last_view_size = size
        if getattr(self, "_view_resize_timer", None) is not None:
            try:
                self.root.after_cancel(self._view_resize_timer)
            except Exception:
                pass
        self._view_resize_timer = self.root.after(300, self._apply_view_resize)

    def _apply_view_resize(self):
        self._view_resize_timer = None
        try:
            if self.root.state() in ("withdrawn", "iconic"):
                return
        except Exception:
            pass
        new_ps = self._compute_page_size()
        if new_ps != getattr(self, "_page_size", 0):
            self._page_size = new_ps
            if IndexEngine.index_exists():
                self._run_query(self._search_text())
            else:
                self._update_result_status()

    def _layout_result_container(self, event):
        if event.width < 4 or event.height < 4:
            return
        self.result_canvas.itemconfigure(self._result_window, width=event.width - 2, height=event.height - 2)
        self.result_canvas.delete("container")
        shape = _rounded_rect(self.result_canvas, 1, 1, event.width - 1, event.height - 1, 6,
                              fill=self.colors["surface"], outline=self.colors["border"],
                              width=1, tags="container")
        self.result_canvas.tag_lower(shape)

    def _build_context_menu(self):
        """构建自绘右键菜单（圆角、阴影、悬停高亮）。"""
        self._ctx_menu = CtxMenu(self)

    def _on_right_click(self, event, row_idx: int):
        """右键：选中行已由表格处理，根据选中数量置灰部分菜单项。"""
        if not (0 <= row_idx < len(self._results)):
            return
        if not self.table.selected_results():
            return
        # 记录右键点下的屏幕坐标，供后续弹窗（如重命名）跟随鼠标定位。
        # 不能用 winfo_pointerxy：菜单项经 after_idle 触发时指针位置已不可靠。
        self._last_right_click_xy = (event.x_root, event.y_root)
        multi = len(self.table.selected_results()) > 1
        items = [
            {"text": "打开", "icon": "↗", "cmd": self._open_selected, "disabled": multi},
            {"text": "打开所在文件夹", "icon": "⌂", "cmd": self._open_file_location_selected, "disabled": multi},
            ("sep",),
            {"text": "复制", "icon": "⧉", "cmd": self._copy_path},
            {"text": "剪切", "icon": "✂", "cmd": self._cut_path},
            {"text": "复制完整路径", "icon": "≡", "cmd": self._copy_full_path_text},
            ("sep",),
            {"text": "重命名", "icon": "✎", "cmd": self._rename_file_dialog, "disabled": multi},
            ("sep",),
            {"text": "删除到回收站", "icon": "⌫", "cmd": self._delete_file_recycle},
            {"text": "彻底删除", "icon": "✕", "cmd": self._delete_file_permanent, "kind": "danger"},
        ]
        self._ctx_menu.show(event.x_root, event.y_root, items)

    def _build_statusbar(self):
        """底部快捷键栏（Raycast 风）：左侧快捷键提示，右侧索引状态 + 结果计数 + 页码。"""
        c = self.colors
        bar = tk.Frame(self.root, bg=c["title_bg"], height=self._s(44),
                       highlightthickness=1, highlightbackground=c["border"])
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)

        txt_font = self._f(FONT_SMALL)

        # 左侧：索引状态点 + 文字（动作反馈也复用此处）
        left = tk.Frame(bar, bg=c["title_bg"])
        left.pack(side=tk.LEFT, padx=(self._s(18), 0))

        self._statusbar_dot = tk.Label(left, text="●", bg=c["title_bg"], fg=c["muted_2"],
                                       font=self._f(FONT_MICRO))
        self._statusbar_dot.pack(side=tk.LEFT, padx=(0, self._s(6)))
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = tk.Label(left, textvariable=self.status_var, bg=c["title_bg"],
                                     fg=c["muted"], font=txt_font, anchor=tk.W)
        self.status_label.pack(side=tk.LEFT)

        # 不确定进度条（索引时显示）
        self.progress_slot = tk.Frame(left, bg=c["title_bg"], width=self._s(120),
                                      height=self._s(30))
        self.progress_slot.pack(side=tk.LEFT, padx=(self._s(14), 0))
        self.progress_slot.pack_propagate(False)
        self.progress = ttk.Progressbar(self.progress_slot, mode="indeterminate",
                                        length=self._s(100))

        # 右侧：设置钮 · 重建索引钮
        right = tk.Frame(bar, bg=c["title_bg"])
        right.pack(side=tk.RIGHT, padx=(0, self._s(10)))

        self.index_btn = RoundedButton(right, text="⟳ 重建索引", command=self._toggle_index,
                                       width=self._s(104), height=self._s(30), colors=c,
                                       kind="ghost", font_size=self._f(FONT_SMALL)[1])
        self.index_btn.pack(side=tk.LEFT, padx=(0, self._s(4)))
        self.settings_btn = RoundedButton(right, icon="⚙", command=self._open_settings,
                                          width=self._s(34), height=self._s(30), colors=c,
                                          font_size=self._f(FONT_BODY)[1])
        self.settings_btn.pack(side=tk.LEFT)

    def _set_status(self, text: str, kind: str = "normal"):
        self.status_var.set(text)
        color = {"success": self.colors["success"], "error": self.colors["error"],
                 "warning": self.colors["warning"]}.get(kind, self.colors["muted"])
        self.status_label.configure(foreground=color)
        if hasattr(self, "_statusbar_dot"):
            self._statusbar_dot.configure(foreground=color)

    def _set_loading(self, loading: bool, text: str = "正在加载结果…"):
        if loading:
            self._set_status(text)
            self.root.configure(cursor="watch")
            self.root.update_idletasks()
        else:
            self.root.configure(cursor="")

    # ================================================================
    #  索引管理
    # ================================================================

    def _update_index_button_text(self):
        """刷新索引按钮文字与底部状态文字。"""
        if IndexEngine.index_exists():
            count = IndexEngine.index_file_count()
            updated = datetime.fromtimestamp(INDEX_DB.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            self.index_btn.config(text="⟳ 重建索引")
            self._index_state_text = f"索引就绪 · {count:,} 个文件 · 更新于 {updated}"
            self._set_status(self._index_state_text)
        else:
            self.index_btn.config(text="⟳ 创建索引")
            self._index_state_text = "尚未创建索引 · 点击右下按钮创建"
            self._set_status(self._index_state_text, "warning")

    def _toggle_index(self):
        """点击索引按钮：运行中 → 询问是否停止；空闲 → 确认后创建/重建。"""
        if self._index_running:
            keep = "，现有索引保持不变" if IndexEngine.index_exists() else ""
            if _dialog_confirm(self, "停止索引",
                               f"索引正在后台运行，确定停止吗？\n本次扫描的全部进度将被丢弃{keep}。",
                               kind="warn", ok_text="停止"):
                self._stop_index()
                self._set_status("正在停止索引…", "warning")
            return
        if self._engine_cancel:
            return
        if IndexEngine.index_exists():
            if not _dialog_confirm(self, "重建索引",
                                   "重建索引将扫描所有磁盘，可能需要几分钟。\n"
                                   "中途停止将丢弃本次扫描的全部进度（现有索引保持不变）。继续？",
                                   kind="warn", ok_text="重建"):
                return
        else:
            if not _dialog_confirm(self, "创建索引",
                                   "将扫描所有磁盘建立索引，可能需要几分钟。\n"
                                   "中途停止将丢弃本次扫描的全部进度。继续？",
                                   kind="warn", ok_text="创建"):
                return
        self._do_index()

    def _do_index(self):
        """在后台线程中执行索引构建。"""
        if self._index_running:
            return
        self._engine_cancel = False
        self._index_running = True
        self.index_btn.config(text="停止索引")   # 不禁用：运行中可点击以停止
        self.progress.pack(expand=True)
        self.progress.start(12)
        self._set_status("正在创建索引，扫描全盘文件…")

        engine = IndexEngine(
            progress_callback=lambda msg, n: self.root.after(0, self._on_index_progress, msg, n),
            cancel_check=lambda: self._engine_cancel,
        )

        def run():
            try:
                stats = engine.build_index()
                self.root.after(0, lambda result=stats: self._on_index_done(result))
            except Exception as e:
                self.root.after(0, lambda error=str(e): self._on_index_error(error))

        threading.Thread(target=run, daemon=True).start()

    def _stop_index(self):
        """停止正在进行的索引构建。"""
        if not self._engine_cancel:
            self._engine_cancel = True
            self.status_var.set("正在停止索引…")

    def _on_index_progress(self, msg: str, count: int):
        """索引进度回调。"""
        self._set_status(f"正在建立索引 · 已收录 {count:,} 个文件", "warning")

    def _on_index_done(self, stats):
        """索引完成回调，显示文件数和各阶段耗时。"""
        self._engine_cancel = False
        self._index_running = False
        self.progress.stop()
        self.progress.pack_forget()
        self._update_index_button_text()
        if stats is None:
            self._set_status("索引已停止，原索引保持不变", "warning")
            return
        total = stats["total_files"]
        scan_write = stats["scan_write_seconds"]
        optimize = stats["optimize_seconds"]
        elapsed = stats["total_seconds"]
        self._run_query(self._search_text())
        self._set_status(
            f"索引完成 — {total:,} 项｜扫描写入 {scan_write:.1f} 秒｜"
            f"优化 {optimize:.1f} 秒｜总计 {elapsed:.1f} 秒", "success"
        )

    def _on_index_error(self, err: str):
        """索引出错回调。"""
        self._engine_cancel = False
        self._index_running = False
        self.progress.stop()
        self.progress.pack_forget()
        self._update_index_button_text()
        self._set_status(f"索引出错: {err}", "error")

    def _do_index_silent(self):
        """静默后台重建索引，始终保证同一时间只有一个索引任务。"""
        if self._index_running or self._engine_cancel:
            return
        self._engine_cancel = False
        self._index_running = True
        self.index_btn.config(text="停止索引")   # 不禁用：运行中可点击以停止
        self.progress.pack(expand=True)
        self.progress.start(12)
        engine = IndexEngine(
            progress_callback=lambda msg, n: self.root.after(0, self._on_index_progress, msg, n),
            cancel_check=lambda: self._engine_cancel,
        )

        def run():
            try:
                stats = engine.build_index()
                self.root.after(0, lambda result=stats: self._on_index_done(result))
            except Exception as e:
                self.root.after(0, lambda error=str(e): self._on_index_error(error))

        threading.Thread(target=run, daemon=True).start()

    # ================================================================
    #  设置对话框
    # ================================================================

    def _open_settings(self):
        """设置窗口：左侧圆角导航 + 右侧分组卡片（Toggle 开关、主题预览卡），修改即保存。"""
        c = self.colors
        s = self._s
        dlg = tk.Toplevel(self.root)
        dlg.title("设置")
        dlg.resizable(True, True)
        dlg.grab_set()
        dlg.transient(self.root)

        dw, dh = s(900), s(620)
        dlg.minsize(s(780), s(540))
        dlg.configure(bg=c["bg"])
        dlg.update_idletasks()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        px, py = self.root.winfo_rootx(), self.root.winfo_rooty()
        dlg.geometry(f"{dw}x{dh}+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

        body = tk.Frame(dlg, bg=c["bg"])
        body.pack(fill=tk.BOTH, expand=True, padx=s(14), pady=s(14))

        # ====== 左侧导航 ======
        NAV_W = s(200)
        nav_panel = tk.Frame(body, bg=c["title_bg"], highlightthickness=1,
                             highlightbackground=c["border"], width=NAV_W)
        nav_panel.pack(side=tk.LEFT, fill=tk.Y)
        nav_panel.pack_propagate(False)

        tk.Label(nav_panel, text="设置", bg=c["title_bg"], fg=c["text"],
                 font=self._f(FONT_XL, "bold")).pack(anchor=tk.W, padx=s(16),
                                                     pady=(s(18), s(12)))
        nav_items = []

        def _draw_nav(cv, active):
            cv.delete("all")
            w = max(2, cv.winfo_width())
            h = max(2, cv.winfo_height())
            if active:
                _rounded_rect(cv, 1, 1, w - 1, h - 1, 10, fill=c["selected"],
                              outline="", width=0)
            cv.create_text(s(13), h / 2, anchor=tk.W, text=cv._nav_icon,
                           fill=c["accent"] if active else c["muted"],
                           font=self._f(FONT_BODY))
            cv.create_text(s(36), h / 2, anchor=tk.W, text=cv._nav_text,
                           fill=c["sel_text"] if active else c["muted"],
                           font=self._f(FONT_BODY, "bold" if active else "normal"))

        def _make_nav_item(icon, text):
            f = tk.Frame(nav_panel, bg=c["title_bg"], cursor="hand2")
            f.pack(fill=tk.X, padx=s(8), pady=2)
            cv = tk.Canvas(f, height=s(38), bd=0, highlightthickness=0, bg=c["title_bg"])
            cv.pack(fill=tk.X)
            cv._nav_icon = icon
            cv._nav_text = text
            nav_items.append(cv)
            cv.bind("<Configure>", lambda _e, cvs=cv: _draw_nav(cvs, cvs is nav_state["active"]))
            return f, cv

        nav_idx, idx_cv = _make_nav_item("◉", "常规")
        nav_ex, ex_cv = _make_nav_item("⊘", "排除列表")
        nav_about, about_cv = _make_nav_item("ℹ", "关于")
        nav_state = {"active": idx_cv}

        # ====== 右侧内容区 ======
        content = tk.Frame(body, bg=c["bg"])
        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(s(14), 0))

        page_index = tk.Frame(content, bg=c["bg"])
        page_exclude = tk.Frame(content, bg=c["bg"])
        page_about = tk.Frame(content, bg=c["bg"])

        # ---- 设置项变量 ----
        auto_start_var = tk.BooleanVar(value=self._settings.get("auto_index_on_start", False))
        tray_auto_var = tk.BooleanVar(value=self._settings.get("tray_auto_index", False))
        minutes_var = tk.IntVar(value=self._settings.get("tray_auto_index_minutes", 30))
        theme_var = tk.StringVar(value=self._settings.get("theme", "mist"))

        def _persist_general():
            try:
                mins = int(minutes_var.get())
                mins = max(5, min(120, mins))
            except Exception:
                mins = 30
            self._settings["auto_index_on_start"] = auto_start_var.get()
            self._settings["tray_auto_index"] = tray_auto_var.get()
            self._settings["tray_auto_index_minutes"] = mins
            IndexEngine.save_settings(self._settings)

        auto_start_var.trace_add("write", lambda *_: _persist_general())
        tray_auto_var.trace_add("write", lambda *_: _persist_general())
        minutes_var.trace_add("write", lambda *_: _persist_general())

        def _on_theme_change(*_):
            new_theme = theme_var.get()
            if new_theme == self._settings.get("theme"):
                return
            self._settings["theme"] = new_theme
            IndexEngine.save_settings(self._settings)
            self._apply_theme()
            # 主题切换后重建设置窗口，使窗口内全部控件跟随新配色
            try:
                dlg.destroy()
            except Exception:
                pass
            self.root.after(60, self._open_settings)

        theme_var.trace_add("write", lambda *_: _on_theme_change())

        # ---- 通用小构件 ----
        def _card(parent, title):
            card = tk.Frame(parent, bg=c["surface"], highlightthickness=1,
                            highlightbackground=c["border"])
            card.pack(fill=tk.X, pady=(0, s(14)))
            tk.Label(card, text=title, bg=c["surface"], fg=c["text"],
                     font=self._f(FONT_BODY, "bold")).pack(anchor=tk.W, padx=s(18),
                                                           pady=(s(14), s(4)))
            return card

        def _option_row(parent, name, desc):
            row = tk.Frame(parent, bg=c["surface"])
            row.pack(fill=tk.X, padx=s(18), pady=s(6))
            text_f = tk.Frame(row, bg=c["surface"])
            text_f.pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Label(text_f, text=name, bg=c["surface"], fg=c["text"],
                     font=self._f(FONT_BODY), anchor=tk.W).pack(anchor=tk.W)
            if desc:
                tk.Label(text_f, text=desc, bg=c["surface"], fg=c["muted_2"],
                         font=self._f(FONT_SMALL), anchor=tk.W).pack(anchor=tk.W,
                                                                     pady=(2, 0))
            return row

        def _page_head(parent, title, desc):
            tk.Label(parent, text=title, bg=c["bg"], fg=c["text"],
                     font=self._f(FONT_XL, "bold")).pack(anchor=tk.W, pady=(s(6), 0))
            tk.Label(parent, text=desc, bg=c["bg"], fg=c["muted_2"],
                     font=self._f(FONT_SMALL)).pack(anchor=tk.W, pady=(2, s(14)))

        # ================= 常规页 =================
        _page_head(page_index, "常规", "所有修改即时保存，无需重启。")

        card_idx = _card(page_index, "索引")
        row = _option_row(card_idx, "启动时自动重建索引", "每次启动软件后在后台静默重建全盘索引")
        ToggleSwitch(row, c, auto_start_var, width=s(46), height=s(26)).pack(
            side=tk.RIGHT, padx=(s(10), 0))
        row = _option_row(card_idx, "托盘自动重建", "最小化到系统托盘后按固定间隔重建索引")
        ToggleSwitch(row, c, tray_auto_var, width=s(46), height=s(26)).pack(
            side=tk.RIGHT, padx=(s(10), 0))

        # 间隔步进器
        row = _option_row(card_idx, "重建间隔", "托盘自动重建的时间间隔（5 ~ 120 分钟）")
        stepper = tk.Frame(row, bg=c["surface"])
        stepper.pack(side=tk.RIGHT)
        minutes_label = tk.Label(stepper, text=f"{minutes_var.get()} 分钟", bg=c["input"],
                                 fg=c["text"], font=(FONT_MONO, self._f(FONT_SMALL)[1]),
                                 width=9, pady=4, highlightthickness=1,
                                 highlightbackground=c["border"])
        RoundedButton(stepper, text="−", width=s(28), height=s(28), colors=c,
                      command=lambda: minutes_var.set(max(5, int(minutes_var.get() or 30) - 5)),
                      font_size=self._f(FONT_BODY)[1]).pack(side=tk.LEFT, padx=(0, s(6)))
        minutes_label.pack(side=tk.LEFT)
        RoundedButton(stepper, text="＋", width=s(28), height=s(28), colors=c,
                      command=lambda: minutes_var.set(min(120, int(minutes_var.get() or 30) + 5)),
                      font_size=self._f(FONT_BODY)[1]).pack(side=tk.LEFT, padx=(s(6), 0))
        minutes_var.trace_add("write", lambda *_: minutes_label.configure(
            text=f"{minutes_var.get()} 分钟"))
        tk.Frame(card_idx, bg=c["surface"], height=s(8)).pack()

        # ---- 外观卡：主题预览 ----
        card_theme = _card(page_index, "外观")
        theme_row = tk.Frame(card_theme, bg=c["surface"])
        theme_row.pack(fill=tk.X, padx=s(18), pady=(0, s(16)))
        theme_canvases = []

        def _draw_theme_card(cv, value):
            cv.delete("all")
            w = max(2, cv.winfo_width())
            h = max(2, cv.winfo_height())
            active = theme_var.get() == value
            name_h = s(26)
            pv_h = h - name_h - 6
            # 迷你界面预览（按主题实际配色绘制）
            t = THEMES.get(value, THEMES["mist"])
            pv_bg = t["bg"]
            pv_bar = t["surface_alt"]
            pv_box = t["input"]
            pv_row = t["row_line"]
            _rounded_rect(cv, 3, 3, w - 3, pv_h, 9, fill=pv_bg,
                          outline=c["border_strong"], width=1)
            _rounded_rect(cv, s(10), s(9), w * 0.55, s(15), 3, fill=pv_bar, outline="")
            _rounded_rect(cv, s(10), s(20), w - s(10), s(34), 6, fill=pv_box,
                          outline=c["accent"] if active else c["border"], width=1)
            _rounded_rect(cv, s(10), s(40), w * 0.8, s(48), 4, fill=c["selected"], outline="")
            _rounded_rect(cv, s(10), s(54), w * 0.65, s(62), 4, fill=pv_row, outline="")
            # 名称区
            cv.create_text(w / 2, pv_h + name_h / 2 + 3, text=cv._theme_name,
                           fill=c["sel_text"] if active else c["muted"],
                           font=self._f(FONT_SMALL, "bold" if active else "normal"))
            # 选中态：accent 描边 + 角标
            if active:
                _rounded_rect(cv, 1, 1, w - 1, h - 1, 10, fill="",
                              outline=c["accent"], width=2)
                cv.create_oval(w - s(22), 6, w - 6, s(22) - 0, fill=c["accent"], outline="")
                cv.create_text(w - s(14), (6 + s(22)) / 2, text="✓", fill="#FFFFFF",
                               font=self._f(FONT_MICRO, "bold"))

        for name, value in (("雾沙（柔和浅色）", "mist"), ("青墨（柔和深色）", "jade"),
                            ("凝脂（浅色）", "light"), ("墨玉（深色）", "dark")):
            cv = tk.Canvas(theme_row, width=s(138), height=s(96), bd=0,
                           highlightthickness=0, bg=c["surface"], cursor="hand2")
            cv.pack(side=tk.LEFT, padx=(0, s(12)))
            cv._theme_name = name
            cv._theme_value = value
            cv.bind("<Button-1>", lambda _e, v=value: theme_var.set(v))
            cv.bind("<Configure>", lambda _e, cvs=cv: _draw_theme_card(cvs, cvs._theme_value))
            theme_canvases.append(cv)
        theme_var.trace_add("write", lambda *_: [
            _draw_theme_card(cvs, cvs._theme_value) for cvs in theme_canvases])
        dlg.after_idle(lambda: [_draw_theme_card(cvs, cvs._theme_value)
                                for cvs in theme_canvases])

        # ---- 主字号（全局生效：主界面/弹窗/右键菜单/设置全部联动）----
        # 主字号 = 表格正文基准 pt（96 DPI 语义，实际渲染再乘 dpi_scale）；
        # 表头/路径/按钮/标题等其余文字、行高与间距按固定比例随主字号自适应
        text_row = tk.Frame(card_theme, bg=c["surface"])
        text_row.pack(fill=tk.X, padx=s(18), pady=(0, s(16)))
        tk.Label(text_row, text="主字号", bg=c["surface"], fg=c["text"],
                 font=self._f(FONT_BODY)).pack(side=tk.LEFT, padx=(0, s(16)))
        text_pt = int(self._settings.get("text_pt", 16))
        TEXT_PT_MIN, TEXT_PT_MAX = 10, 40

        val_lbl = tk.Label(text_row, text=f"{text_pt} pt", bg=c["surface"], fg=c["accent"],
                           font=self._f(FONT_BODY, "bold"), width=6, anchor=tk.E)
        val_lbl.pack(side=tk.RIGHT, padx=(s(8), 0))

        scale = ttk.Scale(text_row, from_=TEXT_PT_MIN, to=TEXT_PT_MAX,
                          orient=tk.HORIZONTAL, style="TextPt.Horizontal.TScale",
                          command=lambda v: val_lbl.config(text=f"{int(float(v))} pt"))
        scale.set(text_pt)
        scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, s(12)))

        def _commit_text_pt(_e=None):
            """松手/键盘调整后才保存生效：拖动过程只更新数值预览。"""
            pt = int(float(scale.get()))
            if pt == self._settings.get("text_pt"):
                return
            self._settings["text_pt"] = pt
            IndexEngine.save_settings(self._settings)
            self._apply_text_scale()
            # 字号变化后整窗尺寸全变，重建设置窗口以刷新配色与布局
            try:
                dlg.destroy()
            except Exception:
                pass
            self.root.after(60, self._open_settings)

        scale.bind("<ButtonRelease-1>", _commit_text_pt)
        scale.bind("<KeyRelease>", _commit_text_pt)
        tk.Label(card_theme, text="主字号即表格正文的字号；表头、路径、按钮、标题等其余文字及行高、间距会随主字号自动缩放。",
                 bg=c["surface"], fg=c["muted_2"], font=self._f(FONT_MICRO)).pack(
            anchor=tk.W, padx=s(18), pady=(0, s(6)))

        # ================= 排除列表页 =================
        _page_head(page_exclude, "排除列表", "索引时跳过匹配的目录（修改即时保存，需重建索引生效）。")

        tb = tk.Frame(page_exclude, bg=c["bg"])
        tb.pack(fill=tk.X, pady=(0, s(10)))
        RoundedButton(tb, text="＋ 添加", command=lambda: self._exclude_add(ex_list, dlg),
                      width=s(96), height=s(34), colors=c,
                      font_size=self._f(FONT_BODY)[1]).pack(side=tk.LEFT, padx=(0, s(8)))
        RoundedButton(tb, text="✎ 编辑", command=lambda: self._exclude_edit(ex_list, dlg),
                      width=s(96), height=s(34), colors=c,
                      font_size=self._f(FONT_BODY)[1]).pack(side=tk.LEFT, padx=(0, s(8)))
        RoundedButton(tb, text="✕ 删除", command=lambda: self._exclude_delete(ex_list),
                      width=s(96), height=s(34), colors=c, kind="danger",
                      font_size=self._f(FONT_BODY)[1]).pack(side=tk.LEFT)

        ex_frame = tk.Frame(page_exclude, bg=c["surface"], highlightthickness=1,
                            highlightbackground=c["border"])
        ex_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("type", "value")
        ex_list = ttk.Treeview(ex_frame, columns=columns, show="headings", selectmode="browse",
                               style="Ex.Treeview", height=14)
        ex_list.heading("type", text="类型")
        ex_list.heading("value", text="排除内容")
        ex_list.column("type", width=s(110), minwidth=80, anchor=tk.CENTER)
        ex_list.column("value", width=s(430), minwidth=220)

        scroll_y = ttk.Scrollbar(ex_frame, orient=tk.VERTICAL, command=ex_list.yview)
        ex_list.configure(yscrollcommand=scroll_y.set)
        ex_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        ex_list.bind("<Double-1>", lambda e: self._exclude_edit(ex_list, dlg))

        data = IndexEngine.get_exclude_list()
        for d in data.get("dirs", []):
            ex_list.insert("", tk.END, values=("目录名", d))
        for p in data.get("paths", []):
            ex_list.insert("", tk.END, values=("路径包含", p))

        # ================= 关于页 =================
        _page_head(page_about, "关于", "")
        card_about = _card(page_about, "File Searcher")
        about_top = tk.Frame(card_about, bg=c["surface"])
        about_top.pack(fill=tk.X, padx=s(18), pady=(s(4), s(14)))
        logo_size = s(48)
        logo_cv = tk.Canvas(about_top, width=logo_size, height=logo_size, bd=0,
                            highlightthickness=0, bg=c["surface"])
        logo_cv.pack(side=tk.LEFT, padx=(0, s(14)))
        grad = _make_gradient_pix(self.root, logo_size, logo_size, s(11),
                                  c["accent_grad_a"], c["accent_grad_b"])
        logo_cv.create_image(0, 0, image=grad, anchor=tk.NW)
        u = logo_size / 20.0
        logo_cv.create_oval(5 * u, 5 * u, 11.5 * u, 11.5 * u, outline="#FFFFFF",
                            width=max(1.5, 1.8 * u))
        logo_cv.create_line(10.8 * u, 10.8 * u, 15 * u, 15 * u, fill="#FFFFFF",
                            width=max(2, 2.2 * u), capstyle=tk.ROUND)
        about_text = tk.Frame(about_top, bg=c["surface"])
        about_text.pack(side=tk.LEFT)
        tk.Label(about_text, text="File Searcher · 凝脂版", bg=c["surface"], fg=c["text"],
                 font=self._f(FONT_LG, "bold")).pack(anchor=tk.W)
        tk.Label(about_text, text="全盘文件快速搜索工具，基于本地索引",
                 bg=c["surface"], fg=c["muted"], font=self._f(FONT_SMALL)).pack(
            anchor=tk.W, pady=(4, 0))
        tk.Label(about_text, text="github.com/truelife411/FileSearcher",
                 bg=c["surface"], fg=c["muted_2"],
                 font=(FONT_MONO, self._f(FONT_SMALL)[1])).pack(anchor=tk.W, pady=(4, 0))

        # ====== 导航切换 ======
        def _show_page(page, active_cv):
            for p in (page_index, page_exclude, page_about):
                p.pack_forget()
            page.pack(fill=tk.BOTH, expand=True)
            nav_state["active"] = active_cv
            for cv in nav_items:
                _draw_nav(cv, cv is active_cv)

        _on_idx = lambda _e=None: _show_page(page_index, idx_cv)
        _on_ex = lambda _e=None: _show_page(page_exclude, ex_cv)
        _on_about = lambda _e=None: _show_page(page_about, about_cv)
        for nav_frame, handler in ((nav_idx, _on_idx), (nav_ex, _on_ex), (nav_about, _on_about)):
            nav_frame.bind("<Button-1>", handler)
            for w in nav_frame.winfo_children():
                w.bind("<Button-1>", handler)

        _show_page(page_index, idx_cv)

    # ================================================================
    #  搜索逻辑
    # ================================================================

    def _load_all(self):
        """按当前筛选、排序加载无关键词的第一页。"""
        self._run_query("")

    def _clear_search(self, event=None):
        """清空搜索并恢复当前筛选下的全部结果。"""
        if self._search_timer is not None:
            self.root.after_cancel(self._search_timer)
            self._search_timer = None
        self._set_search_value("")
        self.root.focus_set()
        self._run_query("")
        return "break"

    def _on_search_changed(self):
        """真实搜索文字变化时触发（300ms 防抖延迟）。"""
        if self._suppress_search_trace:
            return
        if self._search_timer is not None:
            self.root.after_cancel(self._search_timer)
        self._search_timer = self.root.after(300, self._do_search)

    def _do_search(self, event=None):
        """按当前关键词、筛选和排序执行第一页查询。"""
        if self._search_timer is not None:
            self.root.after_cancel(self._search_timer)
            self._search_timer = None
        self._run_query(self._search_text())
        return "break" if event is not None else None

    def _count_cached(self, query: str) -> int:
        """结果计数：同一 query+排序 且 DB 未变时走缓存，翻页不重算。"""
        try:
            mtime = INDEX_DB.stat().st_mtime if INDEX_DB.exists() else 0
        except Exception:
            mtime = 0
        key = (query, self._sort_col, self._sort_asc, mtime)
        if key == self._count_cache_key:
            return self._count_cache_val
        val = IndexEngine.result_count(query)
        self._count_cache_key = key
        self._count_cache_val = val
        return val

    def _run_query(self, query: str):
        """统一执行搜索查询：后台线程跑 DB，主线程只渲染；计数走缓存。"""
        self._last_query = query
        if not IndexEngine.index_exists():
            self._results = []
            self._total_results = 0
            self._page = 1
            self._refresh_tree()
            self._update_result_status()
            self._set_status("请先创建索引再搜索", "warning")
            return

        self._set_loading(True)
        self._query_seq += 1
        seq = self._query_seq
        page, page_size = self._page, self._page_size
        sort_col, sort_desc = self._sort_col, not self._sort_asc

        def work():
            total = self._count_cached(query)
            total_pages = max(1, (total + page_size - 1) // page_size) if total > 0 else 1
            pg = max(1, min(page, total_pages))
            offset = (pg - 1) * page_size
            if query:
                rows = IndexEngine.search(query, limit=page_size, offset=offset,
                                          order_col=sort_col, order_desc=sort_desc)
            else:
                rows = IndexEngine.load_all(limit=page_size, offset=offset,
                                            order_col=sort_col, order_desc=sort_desc)
            return total, pg, rows

        def done(fut):
            if seq != self._query_seq:   # 期间又有新查询，丢弃过期结果
                return
            try:
                total, pg, rows = fut.result()
            except Exception:
                self._set_loading(False)
                return
            self._total_results = total
            self._page = pg
            self._results = rows
            self._refresh_tree()
            self._update_sort_heading()
            self._update_result_status()
            self._set_loading(False)

        fut = self._query_executor.submit(work)
        fut.add_done_callback(lambda f: self.root.after(0, done, f))

    def _update_result_status(self):
        """分页栏承载计数与数字页码按钮；状态文字仅在索引状态之外显示搜索概况。"""
        total = self._total_results
        total_pages = self._total_pages()
        self.pager_total_var.set(f"共 {total:,} 个结果")
        self._render_pager_buttons(total_pages)
        if self._index_running:
            return
        if self._last_query:
            self._set_status(f"搜索「{self._last_query}」— 共 {total:,} 个结果")
        elif IndexEngine.index_exists():
            self._set_status(getattr(self, "_index_state_text", "就绪"))

    # ---- 大号数字页码 ----

    def _page_sequence(self, total_pages: int) -> list:
        """生成页码序列（含省略号）：1 … c-1 c c+1 … last。页码恒 >= 1。"""
        cur = max(1, self._page)
        if total_pages <= 7:
            return list(range(1, total_pages + 1))
        pages = {1, total_pages, cur - 1, cur, cur + 1}
        if cur <= 3:
            pages.update((2, 3, 4))
        if cur >= total_pages - 2:
            pages.update((total_pages - 1, total_pages - 2, total_pages - 3))
        # 下限保护：过滤非法页码（如 cur=1 时 cur-1=0）
        pages = {p for p in pages if 1 <= p <= total_pages}
        seq, prev = [], 0
        for p in sorted(pages):
            if prev and p - prev > 1:
                seq.append("…")
            seq.append(p)
            prev = p
        return seq

    def _render_pager_buttons(self, total_pages: int):
        """重建分页按钮：翻页钮钉在中轴两侧固定位（不随页码数量动），数字页码居中。"""
        center = getattr(self, "_pager_btns_frame", None)
        prev_f = getattr(self, "_pager_prev_frame", None)
        next_f = getattr(self, "_pager_next_frame", None)
        if center is None or prev_f is None or next_f is None:
            return
        for f in (center, prev_f, next_f):
            for w in f.winfo_children():
                w.destroy()
        c = self.colors
        s = self._s
        if total_pages <= 1:
            return
        btn_h = s(34)
        num_w = s(40)
        nav_w = s(72)        # 「上一页/下一页」按钮宽
        gap = s(12)          # 数字页码间距
        fsize = self._f(FONT_BODY)[1]
        nav_fsize = self._f(FONT_SMALL)[1]

        # 翻页钮：prev 容器内右对齐（anchor=E），next 容器内左对齐（anchor=W）→ 始终贴着中轴两侧固定位
        pb = RoundedButton(prev_f, text="上一页", command=lambda: self._goto_page_relative(-1),
                           width=nav_w, height=btn_h, colors=c, font_size=nav_fsize)
        if self._page <= 1:
            pb.config(state=tk.DISABLED)
        pb.pack(side=tk.RIGHT)
        # 内联跳页器放进 next_f、位于「下一页」左侧（同容器相邻，永不与页码重叠）
        self._pager_jump_var = tk.StringVar(value=str(self._page))
        entry_w = max(4, len(str(total_pages)) + 1)
        je = tk.Entry(
            next_f, textvariable=self._pager_jump_var, width=entry_w,
            justify=tk.CENTER, bg=c["input"], fg=c["text"],
            insertbackground=c["accent"], relief="flat", bd=0,
            font=self._f(FONT_BODY),
            highlightthickness=1, highlightbackground=c["border"],
            highlightcolor=c["accent"])
        je.pack(side=tk.LEFT, ipady=s(3), padx=(0, s(3)))
        je.bind("<Return>", self._on_jump_to_page)
        # 聚焦时全选，方便直接覆盖输入目标页码
        je.bind("<FocusIn>", lambda _e: je.after_idle(
            lambda: je.selection_range(0, tk.END)))
        tk.Label(next_f, text=f"/ {total_pages}", bg=c["surface"],
                 fg=c["muted"], font=self._f(FONT_SMALL)).pack(side=tk.LEFT,
                 padx=(0, s(10)))
        nb = RoundedButton(next_f, text="下一页", command=lambda: self._goto_page_relative(1),
                           width=nav_w, height=btn_h, colors=c, font_size=nav_fsize)
        if self._page >= total_pages:
            nb.config(state=tk.DISABLED)
        nb.pack(side=tk.LEFT)

        # 数字页码：居中排布
        for item in self._page_sequence(total_pages):
            if item == "…":
                tk.Label(center, text="…", bg=c["surface"], fg=c["muted_2"],
                         font=self._f(FONT_BODY)).pack(side=tk.LEFT, padx=(0, gap))
                continue
            p = item
            active = (p == self._page)
            b = RoundedButton(center, text=str(p), command=lambda pg=p: self._goto_page(pg),
                              width=num_w, height=btn_h, colors=c,
                              kind="accent" if active else "normal", font_size=fsize)
            b.pack(side=tk.LEFT, padx=(0, gap))

    # ================================================================
    #  排除列表管理
    # ================================================================

    def _exclude_add(self, ex_list, parent):
        """添加排除项（自绘输入弹窗 + 类型分段选择）。"""
        result = _dialog_input(
            self, "添加排除项", "索引时将跳过匹配该内容的目录。",
            options=[("目录名", "目录名"), ("路径包含", "路径包含")],
            selected_option="路径包含", ok_text="添加")
        if result:
            type_value, text = result
            ex_list.insert("", tk.END, values=(type_value, text))
            self._exclude_save(ex_list)

    def _exclude_edit(self, ex_list, parent):
        """编辑选中排除项。"""
        sel = ex_list.selection()
        if not sel:
            return
        vals = ex_list.item(sel[0], "values")
        result = _dialog_input(
            self, "编辑排除项", "索引时将跳过匹配该内容的目录。",
            initial=vals[1],
            options=[("目录名", "目录名"), ("路径包含", "路径包含")],
            selected_option=vals[0], ok_text="保存")
        if result:
            type_value, text = result
            ex_list.item(sel[0], values=(type_value, text))
            self._exclude_save(ex_list)

    def _exclude_save(self, ex_list):
        """将排除列表保存到 JSON 文件。"""
        dirs, paths = [], []
        for iid in ex_list.get_children():
            v = ex_list.item(iid, "values")
            if v[0] == "目录名":
                dirs.append(v[1])
            else:
                paths.append(v[1])
        IndexEngine.save_exclude_list(dirs, paths)

    def _exclude_delete(self, ex_list):
        """删除选中的排除项并即时保存。"""
        sel = ex_list.selection()
        if sel:
            ex_list.delete(sel[0])
            self._exclude_save(ex_list)

    # ================================================================
    #  文件列表显示与交互
    # ================================================================

    def _refresh_tree(self):
        """将当前页结果填充到自绘表格，空结果显示空状态。"""
        rows = []
        for result in self._results:
            icon = self._icon_cache.get(result["path"], bool(result.get("is_dir")))
            rows.append({"result": result, "icon": icon, "values": self._tree_item_values(result)})
        self.table.set_rows(rows)
        # 空状态文案区分「无索引」与「真无结果」
        if hasattr(self, "empty_title_var"):
            if not IndexEngine.index_exists():
                self.empty_title_var.set("尚未创建索引")
                self.empty_sub_var.set("点击右下角「创建索引」开始全盘搜索")
            else:
                self.empty_title_var.set("没有匹配的结果")
                self.empty_sub_var.set("换个关键词试试，或检查拼写")
        self._show_empty_state(not self._results)

    def _tree_item_values(self, result: dict):
        is_dir = bool(result.get("is_dir"))
        ext = os.path.splitext(result["name"])[1]
        type_text = "文件夹" if is_dir else (ext.upper().lstrip(".") if ext else "文件")
        size_text = "" if is_dir else format_size(result["size"])
        modified = result["modified"]
        if isinstance(modified, (int, float)):
            modified = datetime.fromtimestamp(int(modified)).strftime("%Y-%m-%d %H:%M:%S")
        return (result["name"], result["path"], type_text, size_text, modified)

    def _get_selected_path(self) -> str | None:
        """获取当前选中文件的完整路径（兼容单选）。"""
        paths = self._get_selected_paths()
        return paths[0] if paths else None

    def _get_selected_paths(self) -> list[str]:
        """获取所有选中文件的完整路径列表。"""
        return [r["path"] for r in self.table.selected_results()]

    def _on_double_click(self, row_idx: int):
        """双击文件名：检查文件是否存在后打开。"""
        if 0 <= row_idx < len(self._results):
            path = self._results[row_idx]["path"]
            if not os.path.exists(path):
                _dialog_confirm(self, "文件不存在", f"文件可能已被移动或删除：\n{path}",
                                kind="warn", ok_text="知道了", show_cancel=False)
                return
            try:
                open_with_default(path)
            except OSError as e:
                _dialog_confirm(self, "打开失败", str(e), kind="warn",
                                ok_text="知道了", show_cancel=False)

    def _open_selected(self):
        """右键菜单 → 打开文件。"""
        path = self._get_selected_path()
        if path:
            if not os.path.exists(path):
                _dialog_confirm(self, "文件不存在", f"文件可能已被移动或删除：\n{path}",
                                kind="warn", ok_text="知道了", show_cancel=False)
                return
            try:
                open_with_default(path)
            except OSError as e:
                _dialog_confirm(self, "打开失败", str(e), kind="warn",
                                ok_text="知道了", show_cancel=False)

    def _open_file_location_selected(self):
        """右键菜单 → 在资源管理器中定位文件。"""
        path = self._get_selected_path()
        if path:
            if not os.path.exists(path):
                parent = os.path.dirname(path)
                if not os.path.exists(parent):
                    _dialog_confirm(self, "路径不存在", f"路径可能已被移动或删除：\n{path}",
                                    kind="warn", ok_text="知道了", show_cancel=False)
                    return
            open_file_location(path)

    # ================================================================
    #  重命名
    # ================================================================

    def _rename_file_dialog(self):
        """弹出重命名对话框（自绘圆角输入弹窗），执行文件重命名并刷新列表。"""
        path = self._get_selected_path()
        if not path:
            return
        old_name = os.path.basename(path)
        at_xy = getattr(self, "_last_right_click_xy", None)
        _diag(f"[rename] dialog at_xy={at_xy}")
        new_name = _dialog_input(self, "重命名", "输入新的文件名：", initial=old_name,
                                 cursor_start=True, at_xy=at_xy)
        if not new_name or new_name == old_name:
            return
        try:
            new_path = rename_file(path, new_name)
            IndexEngine.rename_path(path, new_path)   # 同步索引 DB（含子路径级联）
            selected = self.table.selected_results()
            if selected:
                r = selected[0]
                r["name"] = new_name
                r["path"] = new_path
            self._refresh_tree()
            self.status_var.set(f"已重命名: {old_name} → {new_name}")
        except Exception as e:
            _dialog_confirm(self, "重命名失败", str(e), kind="warn",
                            ok_text="知道了", show_cancel=False)

    def _open_new_window(self):
        """右键菜单 → 新开一个程序窗口。

        PyInstaller onefile 下 __file__ 指向 _MEIxxxx 临时解包目录，把脚本路径
        传给 exe 无意义（还会在每次新窗口时重新解包），打包环境直接启动 exe；
        源码环境带脚本路径启动解释器。
        """
        try:
            python = sys.executable
            flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            if getattr(sys, "frozen", False):
                subprocess.Popen([python], creationflags=flags)
            else:
                subprocess.Popen([python, os.path.abspath(__file__)], creationflags=flags)
            self.status_var.set("已打开新窗口")
        except Exception as e:
            _dialog_confirm(self, "打开失败", str(e), kind="warn",
                            ok_text="知道了", show_cancel=False)

    # ================================================================
    #  删除
    # ================================================================

    def _delete_file_recycle(self):
        """删除到回收站（可恢复），支持多选。"""
        paths = self._get_selected_paths()
        if not paths:
            return
        count = len(paths)
        if count == 1:
            msg = f"确定将「{os.path.basename(paths[0])}」移动到回收站吗？可以在系统回收站中恢复。"
        else:
            msg = f"确定将选中的 {count} 个文件移动到回收站吗？可以在系统回收站中恢复。"
        if not _dialog_confirm(self, "确认删除", msg, kind="danger", ok_text="删除"):
            return
        try:
            send_to_recycle_bin(paths)
            IndexEngine.remove_paths(paths)
            self._run_query(self._last_query)
            if count == 1:
                self.status_var.set(f"已删除到回收站: {os.path.basename(paths[0])}")
            else:
                self.status_var.set(f"已删除 {count} 个文件到回收站")
        except Exception as e:
            _dialog_confirm(self, "删除失败", str(e), kind="warn",
                            ok_text="知道了", show_cancel=False)

    def _delete_file_permanent(self):
        """彻底删除文件（不可恢复，有二次确认），支持多选。"""
        paths = self._get_selected_paths()
        if not paths:
            return
        count = len(paths)
        if count == 1:
            msg = f"确定彻底删除「{os.path.basename(paths[0])}」吗？\n此操作不可恢复！"
        else:
            msg = f"确定彻底删除选中的 {count} 个文件吗？\n此操作不可恢复！"
        if not _dialog_confirm(self, "确认彻底删除", msg, kind="danger", ok_text="彻底删除"):
            return
        try:
            permanent_delete(paths)
            IndexEngine.remove_paths(paths)
            self._run_query(self._last_query)
            if count == 1:
                self.status_var.set(f"已彻底删除: {os.path.basename(paths[0])}")
            else:
                self.status_var.set(f"已彻底删除 {count} 个文件")
        except Exception as e:
            _dialog_confirm(self, "删除失败", str(e), kind="warn",
                            ok_text="知道了", show_cancel=False)

    # ================================================================
    #  拖拽到外部程序
    # ================================================================

    def _setup_drag_drop(self):
        """设置文件拖拽功能（依赖 tkdnd 扩展，经 tkinterdnd2 包提供二进制）。

        tkinterdnd2 装好后按平台把 tkdnd 的 Tcl 目录加入 auto_path 再 package require；
        源码与 PyInstaller 打包环境均适用（spec 已收集 tkinterdnd2 数据）。
        """
        try:
            try:
                import tkinterdnd2
                _tkdnd_root = os.path.join(os.path.dirname(tkinterdnd2.__file__), "tkdnd")
                if sys.platform == "win32":
                    _platform = "win-x64"
                elif sys.platform == "darwin":
                    _platform = "osx-x64"
                else:
                    _platform = "linux-x64"
                try:
                    _tcl9 = int(self.root.tk.call("info", "tclversion").split(".")[0]) >= 9
                except Exception:
                    _tcl9 = False
                _p = os.path.join(_tkdnd_root, _platform + ("-tcl9" if _tcl9 else ""))
                if not os.path.isdir(_p):
                    _p = os.path.join(_tkdnd_root, _platform)
                self.root.tk.call("lappend", "auto_path", _p)
            except Exception:
                pass
            self.root.tk.call("package", "require", "tkdnd")
            self.root.tk.eval(f"tkdnd::drag_source register {self.table._w} DND_Files")
            self._dnd_cb = self.root.register(self._on_dnd_data)
            self.root.tk.eval(f"tkdnd::drag_source handler {self.table._w} drag {self._dnd_cb}")
            self._dnd_ok = True
        except tk.TclError:
            self._dnd_ok = False
            self.table.bind("<B1-Motion>", self._on_drag_fallback)
            self.status_var.set("拖拽不可用（未找到 tkdnd 扩展）")

    def _on_dnd_data(self, *args):
        """拖拽时返回文件路径（格式化为 tkdnd 需要的格式）。"""
        path = self._get_selected_path()
        if path:
            return "{" + path.replace(chr(92), "/") + "}"
        return ""

    def _on_drag_fallback(self, event):
        """tkdnd 不可用时的拖拽回退事件处理。"""
        self._drag_started = getattr(self, '_drag_started', False)
        if not self._drag_started:
            self._drag_started = True
            self.root.after(200, self._reset_drag_flag)

    def _reset_drag_flag(self):
        """重置拖拽标志。"""
        self._drag_started = False

    # ================================================================
    #  排序
    # ================================================================

    def _update_sort_heading(self):
        """刷新表格表头的当前排序箭头。"""
        self.table.set_sort(self._sort_col, self._sort_asc)

    def _sort_by(self, col: str):
        """点击列头排序，并按当前关键词重新查询第一页。"""
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._page = 1
        self._run_query(self._search_text())

    def _bind_shortcuts(self):
        """绑定全局搜索快捷键和仅在结果列表生效的文件操作快捷键。"""
        self.root.bind("<Control-l>", self._focus_search, add="+")
        self.root.bind("<Control-L>", self._focus_search, add="+")
        self.root.bind("<Control-f>", self._focus_search, add="+")
        self.root.bind("<Control-F>", self._focus_search, add="+")
        self.root.bind("<Escape>", self._clear_search, add="+")
        self.table.bind("<Return>", lambda _e: (self._open_selected(), "break")[1])
        self.table.bind("<Control-a>", self._select_all_results)
        self.table.bind("<Control-A>", self._select_all_results)
        self.table.bind("<Control-c>", lambda _e: (self._copy_path(), "break")[1])
        self.table.bind("<Control-C>", lambda _e: (self._copy_path(), "break")[1])
        self.table.bind("<Control-x>", lambda _e: (self._cut_path(), "break")[1])
        self.table.bind("<Control-X>", lambda _e: (self._cut_path(), "break")[1])
        self.table.bind("<Delete>", lambda _e: (self._delete_file_recycle(), "break")[1])
        self.table.bind("<Shift-Delete>", lambda _e: (self._delete_file_permanent(), "break")[1])
        self.table.bind("<Alt-Return>", lambda _e: (self._open_file_location_selected(), "break")[1])

    def _focus_search(self, _event=None):
        """聚焦搜索框并全选搜索文字。"""
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, tk.END)
        self.search_entry.icursor(tk.END)
        return "break"

    def _select_all_results(self, _event=None):
        """仅在结果列表中全选有效结果行。"""
        self.table.select_all()
        return "break"

    # ================================================================
    #  列宽布局持久化
    # ================================================================

    def _load_layout(self):
        """从 JSON 恢复上次列宽（跨显示器按缩放系数换算；无 scale 的旧文件直接忽略）。"""
        if not IndexEngine.LAYOUT_FILE.exists() or not hasattr(self, "table"):
            return
        try:
            data = json.loads(IndexEngine.LAYOUT_FILE.read_text(encoding="utf-8"))
            saved_scale = data.pop("scale", None)
            current = self._dpi_scale * self._font_scale * self.ui_scale
            cols = [("icon", self._s(46))]
            for key, default_w in self._default_cols:
                w = int(data.get(key, default_w))
                if saved_scale and saved_scale > 0 and abs(saved_scale - current) > 0.05:
                    w = max(40, int(w * current / saved_scale))
                if w < 40 or w > default_w * 1.5:
                    w = default_w
                cols.append((key, w))
            self.table._cols = cols
            self.table.redraw()
        except Exception:
            pass

    def _save_layout(self):
        """保存当前列宽（含缩放系数，便于跨显示器还原）。"""
        data = {"scale": round(self._dpi_scale * self._font_scale * self.ui_scale, 4)}
        for key, w in self.table._cols:
            if key != "icon":
                data[key] = w
        try:
            IndexEngine.LAYOUT_FILE.write_text(
                json.dumps(data, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

    def _on_close(self):
        """点「✕」→ 收进系统托盘（任务栏不留图标，托盘可还原/退出）。

        托盘不可用（初始化失败）时直接退出，避免窗口被隐藏后无法找回。
        """
        if self._tray is None:
            self._do_exit()
            return
        self.root.withdraw()
        self._start_tray_auto_index_timer()

    def _minimize_window(self):
        """点「—」→ 标准最小化到任务栏（保留原生窗口框架后 iconify 有效）。"""
        try:
            self.root.iconify()
        except Exception as e:
            _diag(f"[min] iconify fail {e!r}")

    # ================================================================
    #  系统托盘
    # ================================================================

    def _set_win_icon(self, img=None):
        """设置窗口图标（标题栏 + 任务栏 + Alt+Tab）。

        任务栏图标机制：Windows 任务栏按钮请求 WM_GETICON 优先取 ICON_BIG(32)，取不到才
        回退 ICON_SMALL(16)，两者都没有则回退进程可执行文件图标（pythonw 的 python 图标）。
        所以必须同时设 ICON_SMALL + ICON_BIG 并保活 HICON；再叠 Tk iconbitmap 双保险。
        独立 AppUserModelID 让任务栏不归组到 pythonw。
        """
        import tempfile as _tf
        from ctypes import wintypes
        # 生成 / 复用 .ico（延迟重设时 img=None 直接复用已保存文件）
        if getattr(self, "_ico_path", None) is None:
            if img is None:
                return
            try:
                tmp = _tf.NamedTemporaryFile(suffix=".ico", delete=False)
                tmp.close()
                img.save(tmp.name, format="ICO",
                         sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
                self._ico_path = tmp.name   # 保活防 GC
            except Exception:
                return
        ico_path = self._ico_path
        if sys.platform != "win32":
            try:
                self.root.iconbitmap(ico_path)
            except Exception:
                pass
            return
        try:
            user32 = ctypes.windll.user32
            hwnd = self.root.winfo_id()
            IMAGE_ICON, LR_LOADFROMFILE = 1, 0x0010
            WM_SETICON = 0x0080
            ICON_SMALL, ICON_BIG = 0, 1
            user32.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR,
                                          wintypes.UINT, ctypes.c_int, ctypes.c_int, wintypes.UINT]
            user32.LoadImageW.restype = wintypes.HANDLE
            user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                            wintypes.WPARAM, wintypes.LPARAM]
            user32.SendMessageW.restype = wintypes.LPARAM
            h_small = user32.LoadImageW(0, ico_path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
            h_big = user32.LoadImageW(0, ico_path, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
            # ① 先设窗口图标（大+小）——必须在 AUMID 之前！
            #    Windows 首次设置 AUMID 时会用「当时的窗口图标」给任务栏分组做快照；
            #    若 AUMID 在前，快照到的是通用默认图标（第一次启动显示通用图标的根因）。
            if h_small:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_small)
                self._hicon_small = h_small   # 保活
            if h_big:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_big)
                self._hicon_big = h_big       # 保活
            # ② Tk 官方双保险
            try:
                self.root.iconbitmap(ico_path)
            except Exception:
                pass
            # ③ 设 WS_EX_APPWINDOW：让窗口拥有独立任务栏按钮，图标取窗口 WM_SETICON
            #    （进程级 AUMID 的图标 Windows 从可执行文件 pythonw.exe 取，拿不到窗口图标，
            #    导致通用图标；故弃用 AUMID，改窗口级方案）
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            SWP_FRAMECHANGED, SWP_NOMOVE, SWP_NOSIZE = 0x0020, 0x0002, 0x0001
            SWP_NOZORDER, SWP_NOOWNERZORDER = 0x0004, 0x0200
            cur_ex = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(hwnd, GWL_EXSTYLE, cur_ex | WS_EX_APPWINDOW)
            user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE
                                | SWP_NOZORDER | SWP_NOOWNERZORDER)
            _diag(f"[icon] small={bool(h_small)} big={bool(h_big)} appwindow=1")
        except Exception as e:
            _diag(f"[icon] FAIL {e!r}")

    def _create_tray_icon(self, size: int = 64) -> Image.Image:
        """生成程序图标（渐变圆角方块 + 白色放大镜）。size=64 供任务栏高清；托盘用 32 缩放。"""
        s = max(16, int(size))
        img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        c = self.colors
        rgb1 = tuple(int(c["accent_grad_a"].lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
        rgb2 = tuple(int(c["accent_grad_b"].lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
        # 垂直渐变圆角方块底
        grad = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        gd = ImageDraw.Draw(grad)
        for y in range(s):
            t = y / (s - 1)
            col = tuple(int(rgb1[i] + (rgb2[i] - rgb1[i]) * t) for i in range(3))
            gd.line([(0, y), (s - 1, y)], fill=col + (255,))
        mask = Image.new("L", (s, s), 0)
        ImageDraw.Draw(mask).rounded_rectangle((0, 0, s - 1, s - 1), radius=max(2, s // 4), fill=255)
        grad.putalpha(mask)
        # 白色放大镜（按 s 缩放）
        draw = ImageDraw.Draw(grad)
        u = s / 32.0
        draw.ellipse([9 * u, 8 * u, 21 * u, 20 * u], outline=(255, 255, 255, 255),
                     width=max(2, int(3 * u)))
        draw.line([18.5 * u, 18.5 * u, 25 * u, 25 * u], fill=(255, 255, 255, 255),
                  width=max(2, int(4 * u)))
        return grad

    def _setup_tray(self):
        """创建系统托盘图标和菜单。"""
        try:
            icon = self._create_tray_icon(32)
            try:
                # 用 WM_SETICON(small+big) + AppUserModelID + iconbitmap 设置窗口/任务栏图标
                self._set_win_icon(self._create_tray_icon(64))
                # 窗口显示后再重设一次（应对任务栏图标快照时机 / 缩放后丢失）
                self.root.after(400, lambda: self._set_win_icon(None))
                self.root.after(1200, lambda: self._set_win_icon(None))
            except Exception:
                pass
            menu = pystray.Menu(
                pystray.MenuItem("显示窗口", self._tray_restore, default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("新窗口", self._tray_new_window),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("退出", self._tray_exit),
            )
            self._tray = pystray.Icon("FileSearcher", icon, "File Searcher", menu)
            self._tray._default_action = self._tray_restore
            threading.Thread(target=self._tray.run, daemon=True).start()
        except Exception:
            self._tray = None

    def _tray_restore(self, icon=None, item=None):
        """双击托盘图标：恢复窗口。"""
        if self._tray is None:
            return
        self.root.after(0, self._do_restore)

    def _do_restore(self):
        # 取消待执行的托盘定时索引
        self._cancel_tray_auto_index_timer()
        self.root.deiconify()
        self.root.lift()
        self._maximize_to_workarea()   # 恢复也铺满工作区（不依赖 zoomed）
        self.root.focus_force()
        # 搜索框聚焦；有内容则全选，无内容光标归零。
        self.search_entry.focus_set()
        if self._search_text():
            self.search_entry.select_range(0, tk.END)
        else:
            self.search_entry.icursor(0)

    def _start_tray_auto_index_timer(self):
        """如果设置了托盘自动更新，启动定时器，N 分钟后触发一次静默索引。"""
        self._cancel_tray_auto_index_timer()
        if self._settings.get("tray_auto_index") and IndexEngine.index_exists():
            minutes = self._settings.get("tray_auto_index_minutes", 30)
            ms = max(5, minutes) * 60 * 1000
            self._tray_index_after_id = self.root.after(ms, self._on_tray_auto_index)

    def _cancel_tray_auto_index_timer(self):
        """取消待执行的托盘定时索引。"""
        if self._tray_index_after_id is not None:
            try:
                self.root.after_cancel(self._tray_index_after_id)
            except Exception:
                pass
            self._tray_index_after_id = None

    def _on_tray_auto_index(self):
        """托盘定时触发：静默更新索引。"""
        self._tray_index_after_id = None
        if not self._index_running:
            self._do_index_silent()

    def _tray_exit(self, icon=None, item=None):
        """右键菜单「退出」：停止托盘并销毁窗口。"""
        if self._tray is not None:
            self._tray.stop()
        self.root.after(0, self._do_exit)

    def _tray_new_window(self, icon=None, item=None):
        """托盘右键菜单「新窗口」：新开一个程序实例。"""
        self._open_new_window()

    def _do_exit(self):
        """托盘「退出」：彻底结束进程，不留后台 pythonw 残留。

        只 root.destroy() 不够：ThreadPoolExecutor 的工作线程、pystray 图标线程、
        以及挂着的 after 定时器都会让 pythonw 进程继续存活 → 后台残留。
        这里 shutdown 线程池、停托盘、销毁窗口、清理临时图标，最后 os._exit 兜底确保进程结束。
        """
        self._cancel_tray_auto_index_timer()
        # 清理临时窗口图标文件（每次启动在 %TEMP% 生成，退出时删除）
        ico = getattr(self, "_ico_path", None)
        if ico:
            try:
                os.remove(ico)
            except Exception:
                pass
            self._ico_path = None
        # 停托盘图标线程
        tray = getattr(self, "_tray", None)
        if tray is not None:
            try:
                tray.stop()
            except Exception:
                pass
            self._tray = None
        # 关掉 DB 查询线程池（不等新查询）
        executor = getattr(self, "_query_executor", None)
        if executor is not None:
            try:
                executor.shutdown(wait=False)
            except Exception:
                pass
            self._query_executor = None
        # 销毁窗口并退出主循环
        try:
            self.root.destroy()
        except Exception:
            pass
        # 兜底：确保进程真正结束（托盘/线程若仍挂住则强退）
        try:
            os._exit(0)
        except Exception:
            sys.exit(0)

    # ================================================================
    #  复制 / 剪切路径
    # ================================================================

    def _copy_full_path_text(self, event=None):
        """将选中项的完整路径作为纯文本写入 Tk 剪贴板。"""
        paths = self._get_selected_paths()
        if not paths:
            return "break" if event is not None else None
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(paths))
        self.root.update_idletasks()
        self._set_status(f"已复制 {len(paths):,} 个完整路径", "success")
        return "break" if event is not None else None

    def _copy_path(self, event=None):
        """Ctrl+C：将选中文件复制到剪贴板（资源管理器可粘贴）。"""
        paths = self._get_selected_paths()
        if paths and self._set_clipboard_hdrop(paths, move=False):
            count = len(paths)
            if count == 1:
                self.status_var.set(f"已复制: {os.path.basename(paths[0])}")
            else:
                self.status_var.set(f"已复制 {count} 个文件")

    def _cut_path(self, event=None):
        """Ctrl+X：将选中文件剪切到剪贴板（资源管理器可粘贴）。"""
        paths = self._get_selected_paths()
        if paths and self._set_clipboard_hdrop(paths, move=True):
            count = len(paths)
            if count == 1:
                self.status_var.set(f"已剪切: {os.path.basename(paths[0])}")
            else:
                self.status_var.set(f"已剪切 {count} 个文件")

    def _set_clipboard_hdrop(self, paths: list[str], move: bool = False) -> bool:
        """将文件列表写入剪贴板。Windows 用 CF_HDROP 支持资源管理器粘贴；Linux 降级为纯文本路径。"""
        try:
            if sys.platform == "win32":
                return self._set_clipboard_win32(paths, move)
            else:
                return self._set_clipboard_linux(paths, move)
        except Exception as e:
            print(f"Clipboard error: {e}")
            return False

    def _set_clipboard_win32(self, paths: list[str], move: bool) -> bool:
        """Windows：CF_HDROP 格式，资源管理器可粘贴文件。"""
        CF_HDROP = 15
        GMEM_MOVEABLE = 0x0002
        DROPEFFECT_MOVE = 2

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint32, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.restype = ctypes.c_int
        kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
        kernel32.GlobalFree.restype = ctypes.c_void_p

        user32.OpenClipboard.argtypes = [ctypes.c_void_p]
        user32.OpenClipboard.restype = ctypes.c_int
        user32.EmptyClipboard.restype = ctypes.c_int
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p
        user32.CloseClipboard.restype = ctypes.c_int
        user32.RegisterClipboardFormatW.argtypes = [ctypes.c_wchar_p]
        user32.RegisterClipboardFormatW.restype = ctypes.c_uint

        class DROPFILES(ctypes.Structure):
            _fields_ = [
                ("pFiles", ctypes.c_uint32),
                ("x", ctypes.c_long),
                ("y", ctypes.c_long),
                ("fNC", ctypes.c_int32),
                ("fWide", ctypes.c_int32),
            ]

        file_list = "\0".join(paths) + "\0\0"
        file_data = file_list.encode("utf-16-le")

        df = DROPFILES()
        df.pFiles = ctypes.sizeof(DROPFILES)
        df.x = 0
        df.y = 0
        df.fNC = 0
        df.fWide = 1

        total_size = ctypes.sizeof(DROPFILES) + len(file_data)

        h_global = kernel32.GlobalAlloc(GMEM_MOVEABLE, total_size)
        if not h_global:
            return False

        ptr = kernel32.GlobalLock(h_global)
        if not ptr:
            kernel32.GlobalFree(h_global)
            return False

        ctypes.memmove(ptr, ctypes.addressof(df), ctypes.sizeof(DROPFILES))
        ctypes.memmove(ptr + ctypes.sizeof(DROPFILES), file_data, len(file_data))
        kernel32.GlobalUnlock(h_global)

        if not user32.OpenClipboard(0):
            kernel32.GlobalFree(h_global)
            return False

        user32.EmptyClipboard()
        user32.SetClipboardData(CF_HDROP, h_global)

        if move:
            fmt = user32.RegisterClipboardFormatW("Preferred DropEffect")
            if fmt:
                h_eff = kernel32.GlobalAlloc(GMEM_MOVEABLE, 4)
                if h_eff:
                    p_eff = kernel32.GlobalLock(h_eff)
                    if p_eff:
                        ctypes.c_int32.from_address(p_eff).value = DROPEFFECT_MOVE
                        kernel32.GlobalUnlock(h_eff)
                        user32.SetClipboardData(fmt, h_eff)

        user32.CloseClipboard()
        return True

    def _set_clipboard_linux(self, paths: list[str], move: bool) -> bool:
        """Linux：通过 xclip/wl-copy 写入纯文本路径到剪贴板。"""
        text = "\n".join(paths)
        # 优先 Wayland (wl-copy)，其次 X11 (xclip)
        for cmd in ["wl-copy", "xclip", "xsel"]:
            if shutil.which(cmd):
                try:
                    if cmd == "xclip":
                        subprocess.run([cmd, "-selection", "clipboard"],
                                       input=text.encode("utf-8"), check=True)
                    elif cmd == "xsel":
                        subprocess.run([cmd, "--clipboard", "--input"],
                                       input=text.encode("utf-8"), check=True)
                    else:
                        subprocess.run([cmd], input=text.encode("utf-8"), check=True)
                    return True
                except Exception:
                    continue
        # 最后兜底：Tkinter 自带的剪贴板
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            return True
        except Exception:
            return False


# ================================================================
#  程序入口
# ================================================================

def main():
    """创建 Tkinter 根窗口并启动应用。"""
    root = tk.Tk()
    # 注意：不在 main 里设置 tk scaling —— 它会污染 winfo_fpixels 的返回值，
    # DPI 处理统一在 FileSearcherApp 里用原生 API 完成
    FileSearcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
