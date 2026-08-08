#!/usr/bin/env python3
"""File Searcher — 全盘文件快速搜索工具，基于本地索引。"""

import os
import sys
import subprocess
import sqlite3
import json
import shutil
import ctypes
import threading
import queue
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

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
INDEX_SCAN_WORKERS = 4
INDEX_WRITE_BATCH_SIZE = 5000
INDEX_QUEUE_SIZE = 20000
PAGE_SIZE = 5000

FONT_FAMILY = "Microsoft YaHei UI" if sys.platform == "win32" else "sans-serif"
FONT_MONO = "Consolas" if sys.platform == "win32" else "monospace"
# 全局字号体系（基础 pt，最终渲染 = base × dpi_scale × font_scale ≈ base × 2.08）
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
    # 深色「墨青」：近黑石墨底 + 青玉 accent
    "dark": {
        "bg": "#0C1014", "surface": "#131A21", "surface_alt": "#19222B",
        "surface_3": "#212D38", "input": "#0F151B", "border": "#222D38",
        "border_strong": "#31404E", "text": "#E9F0F3", "muted": "#8E9CAA",
        "muted_2": "#5E6C79", "accent": "#3FC1B0", "accent_hover": "#59D2C2",
        "accent_pressed": "#2CA295", "accent_grad_a": "#46C9B6", "accent_grad_b": "#2B9C8E",
        "selected": "#193034", "selected_hover": "#1E3A3E", "row_alt": "#151D24",
        "row_line": "#1B232B", "title_bg": "#10161C", "hover": "#1B242D",
        "sel_text": "#59D2C2",
        "success": "#3ED598", "warning": "#F0B44C", "error": "#E4747E",
        "menu_bg": "#151D25", "dialog_bg": "#151D25",
    },
    # 浅色「晴白」：暖白底 + 深青 accent（保证对比度）
    "light": {
        "bg": "#F2F5F4", "surface": "#FFFFFF", "surface_alt": "#EFF3F2",
        "surface_3": "#E2E9E7", "input": "#FFFFFF", "border": "#DFE6E3",
        "border_strong": "#C3CFCB", "text": "#182228", "muted": "#61727B",
        "muted_2": "#93A1A8", "accent": "#0F9C8B", "accent_hover": "#23B3A1",
        "accent_pressed": "#0B8577", "accent_grad_a": "#23B3A1", "accent_grad_b": "#0B8577",
        "selected": "#E4F4F1", "selected_hover": "#D8EEE9", "row_alt": "#F7FAF9",
        "row_line": "#EDF1F0", "title_bg": "#E9EEEC", "hover": "#F0F5F3",
        "sel_text": "#0B8577",
        "success": "#18A058", "warning": "#D9962C", "error": "#D64545",
        "menu_bg": "#FFFFFF", "dialog_bg": "#FFFFFF",
    },
}

# 文件类型徽章：(文字色, 底色)，按主题区分；底色为类型色低透明度混合预算值
BADGE_STYLES = {
    "dark": {
        "doc": ("#6FA8EF", "#1F2C3C"), "pdf": ("#E4747E", "#2E262D"),
        "xls": ("#7CC784", "#21312E"), "ppt": ("#E8A569", "#2F2C2A"),
        "img": ("#BE8FE0", "#29293A"), "code": ("#5AC8C8", "#1C3137"),
        "zip": ("#D1A26B", "#2C2C2B"), "audio": ("#7FD4A8", "#213232"),
        "video": ("#8FA5E8", "#232840"), "dir": ("#E5C56F", "#2E302B"),
        "file": ("#8E9CAA", "#212D38"),
    },
    "light": {
        "doc": ("#2F6FD0", "#E9F1FB"), "pdf": ("#C94F5A", "#FAEDEE"),
        "xls": ("#3E9B4F", "#EBF5ED"), "ppt": ("#C97A2B", "#FAF1E7"),
        "img": ("#9A5BC4", "#F4EDFA"), "code": ("#1D9B9B", "#E7F5F5"),
        "zip": ("#B07E3F", "#F7F1E8"), "audio": ("#3FA873", "#EBF6F0"),
        "video": ("#5B6FC4", "#EDEFF9"), "dir": ("#B08F2E", "#F7F3E7"),
        "file": ("#61727B", "#EFF3F2"),
    },
}

# 扩展名 → 徽章类别（构建自上方扩展名集合，见下方 _build_badge_kind_map）
DOCUMENT_EXTENSIONS = {".doc", ".docx", ".pdf", ".txt", ".rtf", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods", ".csv", ".md"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff", ".ico", ".svg"}
VIDEO_EXTENSIONS = {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm", ".m4v"}
AUDIO_EXTENSIONS = {".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a"}
ARCHIVE_EXTENSIONS = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".cab"}
CODE_EXTENSIONS = {".py", ".js", ".ts", ".tsx", ".jsx", ".java", ".c", ".cpp", ".h", ".hpp", ".cs", ".go", ".rs", ".php", ".rb", ".swift", ".kt", ".html", ".css", ".scss", ".json", ".xml", ".yaml", ".yml", ".sql", ".sh", ".ps1"}
TYPE_FILTERS = {
    "文档": DOCUMENT_EXTENSIONS, "图片": IMAGE_EXTENSIONS, "视频": VIDEO_EXTENSIONS,
    "音频": AUDIO_EXTENSIONS, "压缩包": ARCHIVE_EXTENSIONS, "代码": CODE_EXTENSIONS,
}


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
    pass


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
    """用系统默认软件打开文件（失败时抛出 OSError，由调用方提示）。"""
    if sys.platform == "win32":
        os.startfile(os.path.normpath(path))
    else:
        subprocess.Popen(["xdg-open", path])


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
        """执行单个全量索引任务，并返回各阶段耗时。"""
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
                modified TEXT NOT NULL,
                is_dir INTEGER NOT NULL DEFAULT 0
            )
        """)

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
        producer_done = object()
        stop_event = threading.Event()
        worker_count = min(len(drives), INDEX_SCAN_WORKERS)
        completed_producers = 0
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

        def scan_drive(drive):
            try:
                self._scan_drive(drive, put_queue_item)
            finally:
                put_queue_item(producer_done)

        try:
            with ThreadPoolExecutor(max_workers=worker_count, thread_name_prefix="index-scan") as executor:
                futures = [executor.submit(scan_drive, drive) for drive in drives]

                while completed_producers < len(drives):
                    item = result_queue.get()
                    if item is producer_done:
                        completed_producers += 1
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

    def _scan_drive(self, drive, put_queue_item):
        """扫描单个盘符，并把文件元数据放入有界队列。"""
        dir_stack = [drive]
        while dir_stack and not self._cancel():
            dirpath = dir_stack.pop()
            try:
                entries = os.scandir(dirpath)
            except (PermissionError, OSError):
                continue

            with entries:
                for entry in entries:
                    if self._cancel():
                        return
                    full = entry.path
                    try:
                        is_dir = entry.is_dir(follow_symlinks=False)
                    except OSError:
                        continue
                    if is_dir:
                        if not self._should_skip_dir(full, entry):
                            try:
                                st = entry.stat(follow_symlinks=False)
                                item = (
                                    entry.name, entry.name.lower(), full, full.lower(), 0,
                                    datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"), 1,
                                )
                                if not put_queue_item(item):
                                    return
                            except OSError:
                                pass
                            dir_stack.append(full)
                        continue

                    try:
                        is_file = entry.is_file(follow_symlinks=False)
                    except OSError:
                        continue
                    if not is_file or not self._should_include_file(full, entry):
                        continue
                    try:
                        st = entry.stat()
                    except OSError:
                        continue

                    item = (
                        entry.name,
                        entry.name.lower(),
                        full,
                        full.lower(),
                        st.st_size,
                        datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                        0,
                    )
                    if not put_queue_item(item):
                        return

    def _should_skip_dir(self, dirpath: str, entry) -> bool:
        """判断目录是否应被跳过（系统目录、junction 点、排除列表等）。"""
        basename = entry.name.lower()
        try:
            if entry.is_junction():
                return True
        except Exception:
            pass
        if FILE_ATTR_SKIP:
            try:
                attrs = entry.stat(follow_symlinks=False).st_file_attributes
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
    def index_file_count() -> int:
        """返回索引中的文件总数。"""
        if not INDEX_DB.exists():
            return 0
        conn = sqlite3.connect(str(INDEX_DB))
        row = conn.execute("SELECT COUNT(*) FROM files").fetchone()
        conn.close()
        return row[0] if row else 0

    _COL_DB = {"name": "name_lower", "path": "path_lower", "type": "name_lower", "size": "size", "modified": "modified"}

    @staticmethod
    def _query_files(query: str = "", limit: int = PAGE_SIZE, offset: int = 0,
                     order_col: str = "name", order_desc: bool = False, filters: dict | None = None,
                     count_only: bool = False):
        """使用参数化条件查询索引，统一支持搜索、筛选、排序、计数和分页。"""
        if not INDEX_DB.exists():
            return 0 if count_only else []
        filters = filters or {}
        conn = sqlite3.connect(str(INDEX_DB))
        conn.row_factory = sqlite3.Row
        has_is_dir = any(row[1] == "is_dir" for row in conn.execute("PRAGMA table_info(files)"))
        is_dir_sql = "is_dir" if has_is_dir else "0"
        clauses, params = [], []
        q = query.strip().lower()
        if q:
            clauses.append("(name_lower LIKE ? OR path_lower LIKE ?)")
            params.extend((f"%{q}%", f"%{q}%"))

        path_prefix = filters.get("path_prefix")
        if path_prefix:
            prefix = os.path.normcase(os.path.normpath(path_prefix)).lower()
            child_prefix = prefix.rstrip("\\/") + os.sep
            escaped_child_prefix = (child_prefix.replace("\\", "\\\\")
                                    .replace("%", "\\%")
                                    .replace("_", "\\_"))
            clauses.append("(path_lower = ? OR path_lower LIKE ? ESCAPE '\\')")
            params.extend((prefix, escaped_child_prefix + "%"))

        type_name = filters.get("type", "全部")
        if type_name == "文件夹":
            clauses.append(f"{is_dir_sql} = 1")
        elif type_name != "全部":
            clauses.append(f"{is_dir_sql} = 0")
            extensions = TYPE_FILTERS.get(type_name, set())
            if extensions:
                ext_clauses = []
                for ext in sorted(extensions):
                    ext_clauses.append("name_lower LIKE ?")
                    params.append("%" + ext)
                clauses.append("(" + " OR ".join(ext_clauses) + ")")

        time_name = filters.get("time", "全部")
        days = {"今天": 0, "近7天": 7, "近30天": 30}.get(time_name)
        if days is not None:
            start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            if days:
                start -= timedelta(days=days - 1)
            clauses.append("modified >= ?")
            params.append(start.strftime("%Y-%m-%d %H:%M:%S"))

        size_name = filters.get("size", "全部")
        if size_name == "<1MB":
            clauses.append(f"({is_dir_sql} = 0 AND size < ?)")
            params.append(1024 * 1024)
        elif size_name == "1-100MB":
            clauses.append(f"({is_dir_sql} = 0 AND size >= ? AND size <= ?)")
            params.extend((1024 * 1024, 100 * 1024 * 1024))
        elif size_name == ">100MB":
            clauses.append(f"({is_dir_sql} = 0 AND size > ?)")
            params.append(100 * 1024 * 1024)

        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        if count_only:
            row = conn.execute("SELECT COUNT(*) FROM files" + where, params).fetchone()
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
            f"SELECT name, path, size, modified, {is_dir_sql} AS is_dir FROM files"
            f"{where} ORDER BY {order} LIMIT ? OFFSET ?"
        )
        rows = conn.execute(sql, (*params, limit, offset)).fetchall()
        conn.close()
        return [dict(row) for row in rows]

    @staticmethod
    def search(query: str, limit: int = PAGE_SIZE, offset: int = 0, order_col: str = "name",
               order_desc: bool = False, filters: dict | None = None) -> list[dict]:
        return IndexEngine._query_files(query, limit, offset, order_col, order_desc, filters)

    @staticmethod
    def load_all(limit: int = PAGE_SIZE, offset: int = 0, order_col: str = "size",
                 order_desc: bool = True, filters: dict | None = None) -> list[dict]:
        return IndexEngine._query_files("", limit, offset, order_col, order_desc, filters)

    @staticmethod
    def result_count(query: str = "", filters: dict | None = None) -> int:
        return IndexEngine._query_files(query, filters=filters, count_only=True)

    @staticmethod
    def remove_paths(paths: list[str]):
        """从索引中同步删除指定路径及其子路径。"""
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
                conn.execute(
                    "DELETE FROM files WHERE path_lower = ? OR path_lower LIKE ? ESCAPE '\\'",
                    (normalized, escaped_child_prefix + "%"),
                )
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
        """保存排除列表到 JSON 文件。"""
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
        "theme": "dark",
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
        if settings.get("theme") not in {"dark", "light", "system"}:
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

    def __init__(self, master, text="", command=None, width=120, height=48, radius=12,
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
    """大圆角搜索外壳：内嵌 Entry、放大镜、聚焦光晕、Ctrl+F 提示胶囊与清空热区。"""

    def __init__(self, master, textvariable, colors, clear_command, height=56, font_size=None):
        super().__init__(master, height=height, bd=0, highlightthickness=0,
                         bg=master.cget("bg"))
        self.colors = colors
        self._focused = False
        self._hover = False
        self._radius = 15
        self._font_size = font_size or FONT_INPUT
        self.entry = tk.Entry(self, textvariable=textvariable, relief="flat", bd=0,
                              bg=colors["input"], fg=colors["text"],
                              insertbackground=colors["accent"],
                              font=(FONT_FAMILY, self._font_size))
        self._entry_window = self.create_window(52, height / 2, window=self.entry, anchor=tk.W)
        self._clear_command = clear_command
        self._kbd_font = (FONT_MONO, max(8, self._font_size - 5))
        self.bind("<Configure>", self._layout)
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda _e: self._on_hover(True))
        self.bind("<Leave>", lambda _e: self._on_hover(False))
        self.entry.bind("<FocusIn>", self._focus_in, add="+")
        self.entry.bind("<FocusOut>", self._focus_out, add="+")
        self._draw()

    def _layout(self, _event=None):
        # entry 高度必须给足（= 壳高 - 上下留白），否则文字被裁剪
        h = max(24, self.winfo_height() - 10)
        # 右侧预留：清空钮区 + kbd 胶囊区
        self.itemconfigure(self._entry_window, width=max(20, self.winfo_width() - 176), height=h)
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
        # 聚焦光晕：外层淡 accent 圆角描边
        if self._focused:
            _rounded_rect(self, 1, 1, width - 1, height - 1, self._radius + 2,
                          fill="", outline=c["selected"], width=5, tags="shell")
        shell = _rounded_rect(self, 2.5, 2.5, width - 2.5, height - 2.5, self._radius,
                              fill=c["input"], outline=border, width=1.6, tags="shell")
        self.tag_lower(shell)
        # 搜索图标（放大镜，聚焦时 accent 色）
        icon_c = c["accent"] if self._focused else c["muted"]
        cx, cy = 28, height / 2
        self.create_oval(cx - 7.5, cy - 7.5, cx + 7.5, cy + 7.5,
                         outline=icon_c, width=2.2, tags="shell")
        self.create_line(cx + 6, cy + 6, cx + 12, cy + 12,
                         fill=icon_c, width=2.8, capstyle=tk.ROUND, tags="shell")
        # 右侧：Ctrl+F 快捷键胶囊（两块小胶囊）
        kx = width - 96
        for i, t in enumerate(("Ctrl", "F")):
            w = 34 if i == 0 else 22
            x0 = kx if i == 0 else kx + 38
            _rounded_rect(self, x0, cy - 11, x0 + w, cy + 11, 6,
                          fill=c["surface_3"], outline="", width=0, tags="shell")
            self.create_text(x0 + w / 2, cy, text=t, fill=c["muted"],
                             font=self._kbd_font, tags="shell")
        # 清空按钮（hover/focus 时显示）
        clear_x = width - 30
        if self._hover or self._focused:
            self.create_oval(clear_x - 12, cy - 12, clear_x + 12, cy + 12,
                             fill=c["surface_3"], outline="", tags="shell")
            self.create_text(clear_x, cy, text="✕", fill=c["muted"],
                             font=(FONT_FAMILY, FONT_SMALL), tags="shell")
        self.tag_lower("shell", self._entry_window)

    def _click(self, event):
        if event.x >= self.winfo_width() - 50:
            self._clear_command()
        else:
            self.entry.focus_set()

    def _focus_in(self, _event=None):
        self._focused = True
        self._draw()

    def _focus_out(self, _event=None):
        self._focused = False
        self._draw()


class StatusPill(tk.Canvas):
    """索引状态胶囊：圆角全胶囊 + 状态圆点 + 文字。kind: ok 绿 / warn 黄 / off 灰。"""

    def __init__(self, master, colors, height=30, font=None):
        super().__init__(master, height=height, bd=0, highlightthickness=0,
                         bg=master.cget("bg"))
        self.colors = colors
        self._font = font or (FONT_FAMILY, FONT_MICRO)
        self._text = ""
        self._kind = "off"
        self.bind("<Configure>", lambda _e: self._draw())

    def set_status(self, kind: str, text: str):
        if kind != self._kind or text != self._text:
            self._kind = kind
            self._text = text
            self._draw()

    def _draw(self):
        self.delete("all")
        c = self.colors
        h = max(2, self.winfo_height())
        try:
            import tkinter.font as tkfont
            f = tkfont.Font(root=self.winfo_toplevel(), font=self._font)
            tw = f.measure(self._text)
        except Exception:
            tw = len(self._text) * 9
        w = tw + 58
        self.configure(width=w)
        _rounded_rect(self, 1, 1, w - 1, h - 1, (h - 2) / 2,
                      fill=c["surface"], outline=c["border"], width=1)
        dot_c = {"ok": c["success"], "warn": c["warning"]}.get(self._kind, c["muted_2"])
        cy = h / 2
        self.create_oval(16, cy - 4, 24, cy + 4, fill=dot_c, outline="")
        self.create_text(34, cy, text=self._text, anchor=tk.W, fill=c["muted"], font=self._font)


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
        self._w, self._h = width, height
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
        w, h = self._w, self._h
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

    def __init__(self, app, width_px, height_px, radius=16):
        self.app = app
        self.root = app.root
        self.colors = app.colors
        self.result = None
        self._radius = radius
        top = tk.Toplevel(self.root)
        self.top = top
        top.overrideredirect(True)
        self._rounded_ok = True
        try:
            top.configure(bg=_DIALOG_MAGIC)
            top.attributes("-transparentcolor", _DIALOG_MAGIC)
        except tk.TclError:
            self._rounded_ok = False
            top.configure(bg=self.colors["dialog_bg"])
        top.transient(self.root)
        top.resizable(False, False)
        self.root.update_idletasks()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        px, py = self.root.winfo_rootx(), self.root.winfo_rooty()
        top.geometry(f"{width_px}x{height_px}+{px + (pw - width_px) // 2}+{py + (ph - height_px) // 2}")

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
        # 柔和投影（两层偏移深色圆角矩形）
        if self._rounded_ok:
            shadow = "#05080A" if self.app._theme_name == "dark" else "#C9D2CE"
            _rounded_rect(self.canvas, p - 4, p + 2, w - p + 4, h - p + 6, self._radius + 3,
                          fill=shadow, outline="", width=0)
        _rounded_rect(self.canvas, p, p, w - p, h - p, self._radius,
                      fill=c["dialog_bg"], outline=c["border_strong"], width=1)

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

    def run(self):
        """模态运行，返回 result。"""
        self.top.grab_set()
        self.top.focus_set()
        self.root.wait_window(self.top)
        return self.result


def _dialog_confirm(app, title, desc, kind="warn", ok_text="确定", cancel_text="取消",
                    show_cancel=True):
    """确认/提示弹窗。kind: warn(黄) / danger(红) / info(青)。返回 True=确认。"""
    c = app.colors
    s = app._s
    w = s(400)
    # 估算高度：图标 46 + 标题 30 + 描述行数 + 按钮 52 + 边距
    approx_chars_per_line = 26
    lines = max(1, (len(desc) + approx_chars_per_line - 1) // approx_chars_per_line)
    h = s(150) + lines * s(20) + s(64)
    shell = _DialogShell(app, w, h)
    body = shell.body
    icon_map = {"warn": ("!", c["warning"]), "danger": ("✕", c["error"]), "info": ("i", c["accent"])}
    icon_char, icon_c = icon_map.get(kind, icon_map["info"])
    badge_bg = {"warn": BADGE_STYLES[app._theme_name]["zip"][1],
                "danger": BADGE_STYLES[app._theme_name]["pdf"][1],
                "info": app.colors["selected"]}[kind]
    ic = tk.Canvas(body, width=s(44), height=s(44), bd=0, highlightthickness=0, bg=c["dialog_bg"])
    ic.pack(anchor=tk.W, pady=(s(14), s(12)))
    ic.create_oval(2, 2, s(44) - 2, s(44) - 2, fill=badge_bg, outline="")
    ic.create_text(s(22), s(22), text=icon_char, fill=icon_c, font=app._f(FONT_LG, "bold"))
    tk.Label(body, text=title, bg=c["dialog_bg"], fg=c["text"],
             font=app._f(FONT_LG, "bold")).pack(anchor=tk.W)
    tk.Label(body, text=desc, bg=c["dialog_bg"], fg=c["muted"],
             font=app._f(FONT_SMALL), justify=tk.LEFT, wraplength=w - s(60)).pack(
        anchor=tk.W, pady=(s(6), 0))
    btns = tk.Frame(body, bg=c["dialog_bg"])
    btns.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, s(12)))
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
                  selected_option=None):
    """输入弹窗。options 提供时显示分段胶囊选择器。

    返回：未提供 options → 输入字符串或 None；提供 options → (option_value, 文本) 或 None。
    """
    c = app.colors
    s = app._s
    w = s(460)
    h = s(150) + (s(44) if options else 0) + s(90)
    shell = _DialogShell(app, w, h)
    body = shell.body
    tk.Label(body, text=title, bg=c["dialog_bg"], fg=c["text"],
             font=app._f(FONT_LG, "bold")).pack(anchor=tk.W, pady=(s(14), 0))
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
    entry_box.entry.focus_set()
    entry_box.entry.selection_range(0, tk.END)

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
    return shell.run()


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
        font = self.app._f(FONT_SMALL)
        icon_font = self.app._f(FONT_SMALL)
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
                 on_scroll_page=None, on_col_resize=None, badge_styles=None):
        super().__init__(master, bd=0, highlightthickness=0, bg=colors["surface"],
                         cursor="")
        self.colors = colors
        self._icon_cache = icon_cache
        self._font_body = font_body
        self._font_header = font_header
        self._badge_styles = badge_styles or BADGE_STYLES["dark"]
        # 派生字体：路径列小一号、徽章 MICRO 加粗、数字列等宽（按 BODY 比例换算）
        body_pt = font_body[1] if isinstance(font_body, tuple) else FONT_BODY
        self._font_small = (FONT_FAMILY, max(8, round(body_pt * FONT_SMALL / FONT_BODY)))
        self._font_badge = (FONT_FAMILY, max(8, round(body_pt * FONT_MICRO / FONT_BODY)), "bold")
        self._font_mono = (FONT_MONO, max(8, round(body_pt * FONT_SMALL / FONT_BODY)))
        self._on_header_click = on_header_click
        self._on_double = on_double
        self._on_right = on_right
        self._on_scroll_page = on_scroll_page
        self._on_col_resize = on_col_resize
        self._cols = []            # [(key, width)]
        self._rows = []
        self._selected = []
        self._hover_row = None
        self._hover_col = None
        self._sort_col = None
        self._sort_asc = True
        self._drag_key = None
        self._drag_start_x = 0
        self._drag_orig_w = 0
        self._row_h = 50
        self._labels = {}
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
        self.bind("<ButtonRelease-1>", lambda _e: setattr(self, "_drag_key", None))
        self.bind("<MouseWheel>", self._on_wheel)
        self.bind("<Motion>", self._on_motion)
        self.bind("<Leave>", lambda _e: self._set_hover(None, None))

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
        """返回适配可视宽度的列宽列表：总宽超出可视区时等比压缩非图标列（保底 40%）。

        不修改 self._cols（偏好宽度）——拖动列宽与布局保存仍基于偏好值，
        仅绘制/命中/拖动手柄使用压缩后的实际宽度。窗口 resize 或列宽变化后
        由 <Configure> 触发 redraw 自动重新适配。
        """
        if avail_w is None:
            avail_w = max(100, self.winfo_width())
        total = sum(w for _, w in self._cols)
        if total <= avail_w:
            return list(self._cols)
        icon_w = self._cols[0][1] if self._cols else 0
        flex = total - icon_w
        if flex <= 0:
            return list(self._cols)
        k = max(0.4, (avail_w - icon_w) / flex)
        return [(key, w if key == "icon" else max(8, int(w * k)))
                for key, w in self._cols]

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
        row = (y - self.HEADER_H) // self._row_h
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

    # ============ 事件 ============

    def _on_click(self, event):
        self.focus_set()
        handle = self._resize_handle_at(event.x, event.y)
        if handle:
            self._drag_key = handle
            self._drag_start_x = event.x
            self._drag_orig_w = self.column_width(handle)
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

    def _on_drag(self, event):
        if self._drag_key:
            delta = event.x - self._drag_start_x
            self.set_column_width(self._drag_key, self._drag_orig_w + delta)

    def _on_wheel(self, event):
        if self._on_scroll_page:
            self._on_scroll_page(1 if event.delta < 0 else -1)

    def _on_motion(self, event):
        if self._resize_handle_at(event.x, event.y):
            self.configure(cursor="sb_h_double_arrow")
        else:
            self.configure(cursor="")
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
        bh = min(24, self._row_h * 0.46)
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
        # ---- 数据行 ----
        row_h = self._row_h
        for i, row in enumerate(self._rows):
            y0 = self.HEADER_H + i * row_h
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
                text = self._truncate(str(vals[idx]), max(10, x1 - x0 - pad_l * 2), fobj) if idx < len(vals) else ""
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
        self._last_filters = {}
        self._total_results = 0
        self._loading_more = False
        self._tray_index_after_id = None
        self._settings = IndexEngine.load_settings()
        self._theme_name = self._resolve_theme(self._settings.get("theme", "dark"))
        self.colors = THEMES[self._theme_name]

        # 缩放体系：Tk 8.6 在 Windows 上按 96 DPI 布局，tk scaling 对字体渲染无效且会污染
        # winfo_fpixels，因此用原生 API GetDpiForSystem 取真实系统 DPI；
        # _dpi_scale（真实 DPI/96）× _font_scale（用户字号偏好）放大所有 pt 字号与像素尺寸；
        # ui_scale 仅由窗口宽度驱动（1150 基准 0.8~1.8）
        try:
            _dpi = float(ctypes.windll.user32.GetDpiForSystem())
            if _dpi < 72:
                _dpi = 96.0
            self._dpi_scale = max(1.0, _dpi / 96.0)
        except Exception:
            self._dpi_scale = 1.0
        # 固定字号：正文固定 25pt（当前最大化 27pt 再小 2 号），不再随窗口缩放
        self._font_scale = 25 / (FONT_BODY * self._dpi_scale)
        try:
            self.root.tk.call("tk", "scaling", 4 / 3)  # 固定为标准 96dpi 行为
        except Exception:
            pass
        self._window_scale = 1.0
        self.ui_scale = 1.0

        # 自绘标题栏 + 无边框窗口（Windows 专属，失败自动回退原生标题栏）
        self._frameless = False
        self._normal_rect = None
        self._tb_buttons = []
        self._tb_hit_rects = []
        self._dbl_click_flag = False
        self._orig_wndproc = None

        self._path_options = self._build_path_options()
        self._build_ui()
        self._load_layout()
        self._setup_frameless()
        self._setup_tray()

        # 分页状态：页大小按可视行数自适应
        self._page = 1
        self._page_size = self._compute_page_size()
        self._view_resize_timer = None
        self._layout_save_timer = None

        self.root.after(100, self._load_all)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Unmap>", self._on_minimize)
        self.root.bind("<Configure>", self._on_view_resize, add="+")
        self.root.after(150, self._ensure_maximized)
        self.root.after(1500, self._log_diag)
        if self._settings.get("auto_index_on_start") and IndexEngine.index_exists():
            self.root.after(500, self._do_index_silent)

    def _log_diag(self):
        """启动 1.5s 后记录缩放关键参数到 debug.log，用于诊断字号问题。"""
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
            with open(r"C:\Users\hjf\Documents\代码\FileSearcher\debug.log", "a", encoding="utf-8") as fo:
                fo.write(line)
        except Exception:
            pass

    def _build_ui(self):
        """构建全部主界面控件（支持按新缩放重建）。"""
        self._configure_theme()
        self._icon_cache = WindowsShellIconCache(self.root, self.colors["surface"], size=self._s(22))
        self._build_titlebar()
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
        self._placeholder_visible = False
        if saved["query"]:
            self._set_search_value(saved["query"])
            self.search_entry.configure(fg=self.colors["text"])
        else:
            self._show_placeholder()
        self._refresh_tree()
        self._update_sort_heading()
        self._update_result_status()

    def _ensure_maximized(self):
        """启动后强制铺满工作区（无边框模式）。"""
        if self._frameless and self._normal_rect is None:
            self._maximize_to_workarea()

    def _apply_theme(self):
        """主题切换立即生效：换配色并重建主界面（保留搜索词与结果）。"""
        self._theme_name = self._resolve_theme(self._settings.get("theme", "dark"))
        self.colors = THEMES[self._theme_name]
        self._configure_theme()
        self._rebuild_ui()

    def _resolve_theme(self, theme: str) -> str:
        if theme != "system":
            return theme if theme in THEMES else "dark"
        if sys.platform == "win32":
            try:
                import winreg
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                return "light" if winreg.QueryValueEx(key, "AppsUseLightTheme")[0] else "dark"
            except Exception:
                pass
        return "dark"

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

    # ================================================================
    #  自绘标题栏 + 无边框窗口（Windows）
    # ================================================================

    def _build_titlebar(self):
        """构建自绘标题栏：渐变圆角 logo、标题+副标题、最小化/关闭按钮。"""
        c = self.colors
        self._titlebar_h = self._s(TITLEBAR_H)
        self._tb_buttons = []
        self._tb_hit_rects = []
        bar = tk.Frame(self.root, bg=c["title_bg"], height=self._titlebar_h)
        bar.pack(fill=tk.X, side=tk.TOP)
        bar.pack_propagate(False)
        self._titlebar = bar

        # logo：渐变圆角方块 + 白色放大镜
        logo_size = self._s(20)
        logo = tk.Canvas(bar, width=logo_size, height=logo_size, bd=0, highlightthickness=0,
                         bg=c["title_bg"])
        logo.pack(side=tk.LEFT, padx=(self._s(14), self._s(10)),
                  pady=(self._titlebar_h - logo_size) // 2)
        grad = _make_gradient_pix(self.root, logo_size, logo_size, self._s(5),
                                  c["accent_grad_a"], c["accent_grad_b"])
        logo.create_image(0, 0, image=grad, anchor=tk.NW)
        u = logo_size / 20.0  # 以 20px 为基准的单位缩放
        logo.create_oval(5 * u, 5 * u, 11.5 * u, 11.5 * u, outline="#FFFFFF",
                         width=max(1.5, 1.8 * u))
        logo.create_line(10.8 * u, 10.8 * u, 15 * u, 15 * u, fill="#FFFFFF",
                         width=max(2, 2.2 * u), capstyle=tk.ROUND)

        tk.Label(bar, text="File Searcher", bg=c["title_bg"], fg=c["text"],
                 font=self._f(FONT_TITLE, "bold")).pack(side=tk.LEFT)
        tk.Label(bar, text="全盘文件搜索", bg=c["title_bg"], fg=c["muted_2"],
                 font=self._f(FONT_MICRO)).pack(side=tk.LEFT, padx=(self._s(9), 0))

        def _make_tb_btn(text, hover_bg=None, command=None):
            btn = tk.Label(bar, text=text, bg=c["title_bg"], fg=c["muted_2"],
                           font=self._f(FONT_TITLE, "normal"), width=4, cursor="hand2")
            btn.pack(side=tk.RIGHT, fill=tk.Y)
            btn.bind("<Enter>", lambda _e: btn.configure(bg=hover_bg or c["surface_3"], fg=c["text"]))
            btn.bind("<Leave>", lambda _e: btn.configure(bg=c["title_bg"], fg=c["muted_2"]))
            if command:
                btn.bind("<Button-1>", lambda _e: command())
            self._tb_buttons.append(btn)
            return btn

        # 右侧按钮从右往左：关闭 → 最小化（程序启动即最大化，无需最大化按钮）
        _make_tb_btn("✕", hover_bg="#D64545", command=self._on_close)
        _make_tb_btn("—", command=self._on_close)
        # 双击标题栏空白处：铺满/还原
        bar.bind("<Double-Button-1>", self._toggle_maximize)
        for child in bar.winfo_children():
            if child not in self._tb_buttons:
                child.bind("<Double-Button-1>", self._toggle_maximize)
        bar.bind("<Configure>", lambda _e: self.root.after_idle(self._update_tb_hit_rects))

    def _update_tb_hit_rects(self):
        """缓存标题栏按钮在窗口内的矩形，供 WM_NCHITTEST 区分按钮点击。"""
        rects = []
        for w in self._tb_buttons:
            try:
                rects.append((w.winfo_x(), w.winfo_y(), w.winfo_width(), w.winfo_height()))
            except Exception:
                pass
        self._tb_hit_rects = rects

    def _setup_frameless(self):
        """尝试启用无边框自绘标题栏；任何失败都回退到系统原生标题栏。"""
        if sys.platform != "win32":
            # 非 Windows：隐藏自绘标题栏，用原生标题栏 + 最大化
            try:
                self._titlebar.pack_forget()
                self.root.state("zoomed")
            except Exception:
                pass
            return
        try:
            self.root.overrideredirect(True)
            self._apply_frameless_wndproc()
            self._maximize_to_workarea()
            self._frameless = True
            self.root.after(200, self._poll_dbl_click)
        except Exception:
            self._frameless = False
            try:
                self.root.overrideredirect(False)
                self._titlebar.pack_forget()
                self.root.state("zoomed")
            except Exception:
                pass

    def _apply_frameless_wndproc(self):
        """挂载 WM_NCHITTEST 窗口过程：标题栏系统级拖动、四边缩放、双击最大化。"""
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        if ctypes.sizeof(ctypes.c_void_p) == 8:
            SetWindowLong = ctypes.windll.user32.SetWindowLongPtrW
            GetWindowLong = ctypes.windll.user32.GetWindowLongPtrW
        else:
            SetWindowLong = ctypes.windll.user32.SetWindowLongW
            GetWindowLong = ctypes.windll.user32.GetWindowLongW
        SetWindowLong.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
        SetWindowLong.restype = ctypes.c_ssize_t
        GetWindowLong.argtypes = [wintypes.HWND, ctypes.c_int]
        GetWindowLong.restype = ctypes.c_ssize_t
        user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
        # 64 位 Windows 上 WPARAM/LPARAM 是 64 位，wintypes 里是 32 位，必须用 c_ssize_t
        user32.CallWindowProcW.argtypes = [ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT,
                                           ctypes.c_ssize_t, ctypes.c_ssize_t]
        user32.CallWindowProcW.restype = ctypes.c_ssize_t

        WM_NCHITTEST = 0x0084
        WM_NCLBUTTONDBLCLK = 0x00A3
        HTCLIENT = 1
        HTCAPTION = 2
        HTLEFT, HTRIGHT, HTTOP, HTBOTTOM = 10, 11, 12, 15
        HTTOPLEFT, HTTOPRIGHT = 13, 14
        HTBOTTOMLEFT, HTBOTTOMRIGHT = 16, 17
        EDGE = 6

        hwnd = self.root.winfo_id()
        self._frameless_hwnd = hwnd
        self._tb_hit_rects = []
        self.root.update_idletasks()
        self._update_tb_hit_rects()
        app = self

        def wnd_proc(hwnd, msg, wparam, lparam):
            if msg == WM_NCHITTEST:
                rect = wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                x = lparam & 0xFFFF
                y = (lparam >> 16) & 0xFFFF
                if x >= 0x8000:
                    x -= 0x10000
                if y >= 0x8000:
                    y -= 0x10000
                lx, ly = x - rect.left, y - rect.top
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                # 标题栏区域：按钮 → HTCLIENT（可点击），其余 → HTCAPTION（可拖动）
                if 0 <= ly < app._titlebar_h:
                    for (bx, by, bw, bh) in app._tb_hit_rects:
                        if bx <= lx < bx + bw and by <= ly < by + bh:
                            return HTCLIENT
                    return HTCAPTION
                # 边框缩放
                if lx <= EDGE and ly <= EDGE:
                    return HTTOPLEFT
                if lx >= w - EDGE and ly <= EDGE:
                    return HTTOPRIGHT
                if lx <= EDGE and ly >= h - EDGE:
                    return HTBOTTOMLEFT
                if lx >= w - EDGE and ly >= h - EDGE:
                    return HTBOTTOMRIGHT
                if ly <= EDGE:
                    return HTTOP
                if ly >= h - EDGE:
                    return HTBOTTOM
                if lx <= EDGE:
                    return HTLEFT
                if lx >= w - EDGE:
                    return HTRIGHT
                return HTCLIENT
            if msg == WM_NCLBUTTONDBLCLK and wparam == HTCAPTION:
                app._dbl_click_flag = True
                return 0
            return user32.CallWindowProcW(app._orig_wndproc, hwnd, msg, wparam, lparam)

        self._wnd_proc_fn = wnd_proc
        WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT,
                                     ctypes.c_ssize_t, ctypes.c_ssize_t)
        self._wnd_proc_type = WNDPROC
        proc_ptr = WNDPROC(wnd_proc)
        self._wnd_proc_holder = proc_ptr
        self._orig_wndproc = GetWindowLong(hwnd, -4)  # GWL_WNDPROC
        if not self._orig_wndproc:
            raise OSError("GetWindowLong failed")
        SetWindowLong(hwnd, -4, ctypes.cast(proc_ptr, ctypes.c_void_p).value or proc_ptr)
        # 让无边框窗口保留任务栏按钮
        try:
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            cur = GetWindowLong(hwnd, GWL_EXSTYLE)
            SetWindowLong(hwnd, GWL_EXSTYLE, cur | WS_EX_APPWINDOW)
        except Exception:
            pass

    def _maximize_to_workarea(self):
        """将窗口铺满工作区（排除任务栏）。用 Tk geometry 设置，避免被内部布局重置。"""
        if sys.platform != "win32":
            return
        try:
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            rect = wintypes.RECT()
            user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)  # SPI_GETWORKAREA
            w = rect.right - rect.left
            h = rect.bottom - rect.top
            self.root.geometry(f"{w}x{h}+{rect.left}+{rect.top}")
            self.root.update_idletasks()
            self._normal_rect = None
        except Exception:
            pass

    def _restore_normal_size(self):
        """从铺满状态还原为居中的常规尺寸。"""
        try:
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            rect = wintypes.RECT()
            user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
            wa_w = rect.right - rect.left
            wa_h = rect.bottom - rect.top
            w = int(min(1200, wa_w * 0.86))
            h = int(min(800, wa_h * 0.88))
            x = rect.left + (wa_w - w) // 2
            y = rect.top + (wa_h - h) // 2
            self.root.geometry(f"{w}x{h}+{x}+{y}")
            self.root.update_idletasks()
            self._normal_rect = (x, y, w, h)
        except Exception:
            pass

    def _toggle_maximize(self, _event=None):
        """双击标题栏：铺满工作区 / 还原常规尺寸 之间切换。"""
        if self._frameless:
            if self._normal_rect is None:
                self._restore_normal_size()
            else:
                self._maximize_to_workarea()
        else:
            try:
                if self.root.state() == "zoomed":
                    self.root.state("normal")
                else:
                    self.root.state("zoomed")
            except Exception:
                pass

    def _poll_dbl_click(self):
        """轮询 WndProc 里的双击标志（WndProc 中不能直接调用 Tk）。"""
        if getattr(self, "_dbl_click_flag", False):
            self._dbl_click_flag = False
            self._toggle_maximize()
        self.root.after(150, self._poll_dbl_click)

    def _build_path_options(self):
        home = Path.home()
        candidates = {
            "全部": None,
            "桌面": home / "Desktop",
            "文档": home / "Documents",
            "下载": home / "Downloads",
        }
        return {name: str(path) if path else None for name, path in candidates.items()}


    # ================================================================
    #  UI 构建
    # ================================================================

    def _build_toolbar(self):
        """构建搜索区：全宽大搜索框 + 工具行（状态胶囊 + 幽灵索引按钮 + 设置图标钮）。"""
        c = self.colors
        header = tk.Frame(self.root, bg=c["bg"])
        header.pack(fill=tk.X, padx=self._s(26), pady=(self._s(16), self._s(10)))
        header.columnconfigure(0, weight=1)
        self._header = header

        # ---- 全宽大搜索框 ----
        row_h = self._s(48)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._on_search_changed())
        self.search_box = RoundedSearchBox(header, self.search_var, c, self._clear_search,
                                           height=row_h, font_size=self._f(FONT_INPUT)[1])
        self.search_box.grid(row=0, column=0, sticky="ew")
        self.search_entry = self.search_box.entry
        self.search_entry.bind("<Return>", self._do_search)
        self.search_entry.bind("<FocusIn>", self._hide_placeholder, add="+")
        self.search_entry.bind("<FocusOut>", self._show_placeholder, add="+")
        self._placeholder = "搜索文件名或完整路径…"
        self._placeholder_visible = False
        self._show_placeholder()

        # ---- 工具行：左 状态胶囊 / 右 重建索引 + 设置 ----
        tool_row = tk.Frame(header, bg=c["bg"], height=self._s(34))
        tool_row.grid(row=1, column=0, sticky="ew", pady=(self._s(12), 0))
        tool_row.grid_propagate(False)

        self.index_status_var = tk.StringVar(value="尚未创建索引")
        self.index_count_var = tk.StringVar(value="0 项")
        self.index_updated_var = tk.StringVar(value="未更新")
        self.index_info_var = tk.StringVar(value="尚未创建索引")
        self.status_pill = StatusPill(tool_row, c, height=self._s(30),
                                      font=self._f(FONT_MICRO))
        self.status_pill.pack(side=tk.LEFT)

        self.settings_btn = RoundedButton(tool_row, icon="⚙", command=self._open_settings,
                                          width=self._s(34), height=self._s(34), colors=c,
                                          font_size=self._f(FONT_BODY)[1])
        self.settings_btn.pack(side=tk.RIGHT)
        self.index_btn = RoundedButton(tool_row, text="⟳  重建索引", command=self._toggle_index,
                                       width=self._s(118), height=self._s(34), colors=c,
                                       kind="ghost", font_size=self._f(FONT_BODY)[1])
        self.index_btn.pack(side=tk.RIGHT, padx=(0, self._s(2)))

    def _update_search_width(self):
        """兼容旧调用：搜索框已全宽，无需联动调整。"""
        return

    def _set_search_value(self, value: str):
        """更新搜索变量但不触发查询，用于占位符和程序化清空。"""
        self._suppress_search_trace = True
        try:
            self.search_var.set(value)
        finally:
            self._suppress_search_trace = False

    def _hide_placeholder(self, _event=None):
        if self._placeholder_visible:
            self._set_search_value("")
            self.search_entry.configure(fg=self.colors["text"])
            self._placeholder_visible = False

    def _show_placeholder(self, _event=None):
        if not self.search_var.get() and self.root.focus_get() != self.search_entry:
            self._placeholder_visible = True
            self.search_entry.configure(fg=self.colors["muted"])
            self._set_search_value(self._placeholder)

    def _search_text(self):
        return "" if self._placeholder_visible else self.search_var.get().strip()

    def _current_filters(self):
        """筛选功能已从界面移除，查询不带过滤条件。"""
        return {}

    def _build_tree(self):
        """构建圆角结果容器、自绘呼吸式表格（FileTable）和底部固定分页栏。"""
        c = self.colors
        outer = tk.Frame(self.root, bg=c["bg"])
        outer.pack(fill=tk.BOTH, expand=True, padx=self._s(26), pady=(0, self._s(10)))
        self.result_canvas = tk.Canvas(outer, bd=0, highlightthickness=0, bg=c["bg"])
        self.result_canvas.pack(fill=tk.BOTH, expand=True)
        self.result_surface = tk.Frame(self.result_canvas, bg=c["surface"])
        self._result_window = self.result_canvas.create_window(1, 1, window=self.result_surface, anchor=tk.NW)
        self.result_canvas.bind("<Configure>", self._layout_result_container)

        # 自绘呼吸式表格
        self.table = FileTable(self.result_surface, c, self._icon_cache,
                               self._f(FONT_BODY), self._f(FONT_HEADER),
                               on_header_click=self._sort_by,
                               on_double=self._on_double_click,
                               on_right=self._on_right_click,
                               on_scroll_page=self._goto_page_relative,
                               on_col_resize=self._on_col_resize,
                               badge_styles=BADGE_STYLES[self._theme_name])
        self.table.pack(fill=tk.BOTH, expand=True, padx=1, pady=(1, 0))
        self.table._labels = {"name": "文件名", "path": "路径", "type": "类型",
                              "size": "大小", "modified": "修改时间"}
        self.table._row_h = self._s(ROW_HEIGHT)
        self._default_cols = [("name", self._s(340)), ("path", self._s(620)),
                              ("type", self._s(130)), ("size", self._s(140)),
                              ("modified", self._s(210))]
        self.table._cols = [("icon", self._s(46))] + list(self._default_cols)

        # 底部固定分页栏：左 总数 / 右 ‹ 1/12 ›
        pager = tk.Frame(self.result_surface, bg=c["surface"], height=self._s(40),
                         highlightthickness=1, highlightbackground=c["row_line"])
        pager.pack(fill=tk.X, padx=1, pady=(0, 1))
        pager.pack_propagate(False)
        self.pager_total_var = tk.StringVar(value="共 0 个结果")
        tk.Label(pager, textvariable=self.pager_total_var, bg=c["surface"], fg=c["muted"],
                 font=self._f(FONT_MICRO)).pack(side=tk.LEFT, padx=self._s(14))
        self.pager_next = RoundedButton(pager, text="›",
                                        command=lambda: self._goto_page_relative(1),
                                        width=self._s(30), height=self._s(28), colors=c,
                                        font_size=self._f(FONT_BODY)[1])
        self.pager_next.pack(side=tk.RIGHT, padx=(self._s(4), self._s(10)),
                             pady=(self._s(6), 0))
        self.page_info_var = tk.StringVar(value="1 / 1")
        tk.Label(pager, textvariable=self.page_info_var, bg=c["surface"], fg=c["muted"],
                 font=(FONT_MONO, self._f(FONT_MICRO)[1])).pack(side=tk.RIGHT, padx=self._s(6),
                                                                pady=(self._s(6), 0))
        self.pager_prev = RoundedButton(pager, text="‹",
                                        command=lambda: self._goto_page_relative(-1),
                                        width=self._s(30), height=self._s(28), colors=c,
                                        font_size=self._f(FONT_BODY)[1])
        self.pager_prev.pack(side=tk.RIGHT, pady=(self._s(6), 0))

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
        tk.Label(empty_frame, text="没有匹配的结果", bg=c["surface"], fg=c["muted"],
                 font=self._f(FONT_LG)).pack(pady=(self._s(14), self._s(4)))
        tk.Label(empty_frame, text="换个关键词试试，或检查拼写", bg=c["surface"], fg=c["muted_2"],
                 font=self._f(FONT_SMALL)).pack()
        self.empty_state = empty_frame

    def _show_empty_state(self, show: bool):
        if show:
            self.empty_state.place(in_=self.table, relx=0.5, rely=0.45, anchor=tk.CENTER)
        else:
            self.empty_state.place_forget()

    # ---- 分页 ----

    def _compute_page_size(self) -> int:
        """页大小 = 可视行数（表头下方可容纳的行数），保证无页内滚动。"""
        try:
            h = self.table.winfo_height() - FileTable.HEADER_H
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
            self._run_query(self._search_text())
        else:
            self._update_result_status()

    def _goto_page_relative(self, delta: int):
        self._goto_page(self._page + delta)

    def _on_col_resize(self):
        """列宽拖动后防抖保存布局。"""
        if getattr(self, "_layout_save_timer", None) is not None:
            try:
                self.root.after_cancel(self._layout_save_timer)
            except Exception:
                pass
        self._layout_save_timer = self.root.after(500, self._save_layout)

    def _on_view_resize(self, event=None):
        """窗口尺寸变化：页大小自适应（不动字号，只重查当前页）。"""
        if event is not None and getattr(event, "widget", None) is not self.root:
            return
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
        shape = _rounded_rect(self.result_canvas, 1, 1, event.width - 1, event.height - 1, 14,
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
        """构建极简状态栏：title_bg 底、顶部细线、左状态点+文字、右页码。"""
        c = self.colors
        bar = tk.Frame(self.root, bg=c["title_bg"], height=self._s(34),
                       highlightthickness=1, highlightbackground=c["border"])
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)
        bar.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="就绪 — 请先创建索引")
        self._statusbar_dot = tk.Label(bar, text="●", bg=c["title_bg"], fg=c["muted_2"],
                                       font=self._f(FONT_MICRO))
        self._statusbar_dot.grid(row=0, column=0, sticky="w", padx=(self._s(22), self._s(8)))
        self.status_label = tk.Label(bar, textvariable=self.status_var, bg=c["title_bg"], fg=c["muted"],
                                     font=self._f(FONT_SMALL), anchor=tk.W)
        self.status_label.grid(row=0, column=1, sticky="nsew")
        self.progress_slot = tk.Frame(bar, bg=c["title_bg"], width=self._s(150), height=self._s(32))
        self.progress_slot.grid(row=0, column=2, sticky="ns")
        self.progress_slot.grid_propagate(False)
        self.progress = ttk.Progressbar(self.progress_slot, mode="indeterminate", length=self._s(130))
        self.status_right_var = tk.StringVar(value="")
        tk.Label(bar, textvariable=self.status_right_var, bg=c["title_bg"], fg=c["muted_2"],
                 font=self._f(FONT_SMALL), anchor=tk.E, padx=self._s(22)).grid(row=0, column=3, sticky="nsew")
        bar.rowconfigure(0, weight=1)

    def _set_status(self, text: str, kind: str = "normal"):
        self.status_var.set(text)
        color = {"success": self.colors["success"], "error": self.colors["error"],
                 "warning": self.colors["warning"]}.get(kind, self.colors["muted"])
        self.status_label.configure(foreground=color)
        if hasattr(self, "_statusbar_dot"):
            self._statusbar_dot.configure(foreground=color)

    def _set_index_dot(self, kind: str):
        """更新状态胶囊圆点颜色（文字保持不变）。kind: ok 绿 / warn 黄 / off 灰。"""
        if hasattr(self, "status_pill"):
            self.status_pill.set_status(kind, self.status_pill._text)

    def _pill(self, kind: str, text: str):
        """一次性设置状态胶囊的颜色与文字。"""
        if hasattr(self, "status_pill"):
            self.status_pill.set_status(kind, text)

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
        """更新索引按钮、状态胶囊、数量和更新时间。"""
        if IndexEngine.index_exists():
            count = IndexEngine.index_file_count()
            updated = datetime.fromtimestamp(INDEX_DB.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            self.index_btn.config(text="⟳  重建索引")
            self.index_status_var.set("索引就绪")
            self.index_count_var.set(f"{count:,} 项")
            self.index_updated_var.set(f"更新于 {updated}")
            self.index_info_var.set(f"索引就绪 · {count:,} 个文件 · 更新于 {updated}")
            self._pill("ok", f"索引就绪 · {count:,} 个文件 · 更新于 {updated}")
        else:
            self.index_btn.config(text="⟳  创建索引")
            self.index_status_var.set("尚未创建索引")
            self.index_count_var.set("0 项")
            self.index_updated_var.set("未更新")
            self.index_info_var.set("尚未创建索引 · 0 个文件 · 未更新")
            self._pill("off", "尚未创建索引 · 点击右侧按钮创建")

    def _toggle_index(self):
        """点击索引按钮：创建或重建索引。"""
        if self._index_running or self._engine_cancel:
            return
        if IndexEngine.index_exists():
            if not _dialog_confirm(self, "重建索引", "重建索引将扫描所有磁盘，可能需要几分钟。继续？",
                                   kind="warn", ok_text="重建"):
                return
        self._do_index()

    def _do_index(self):
        """在后台线程中执行索引构建。"""
        if self._index_running:
            return
        self._engine_cancel = False
        self._index_running = True
        self.index_btn.config(state=tk.DISABLED)
        self.progress.pack(expand=True)
        self.progress.start(12)
        self._set_status("正在创建索引，扫描全盘文件…")
        self._pill("warn", "正在建立索引 · 扫描全盘文件中…")

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
        self._set_status(msg)
        self.index_count_var.set(f"已收录 {count:,} 项")
        self.index_info_var.set(f"正在建立索引 · 已收录 {count:,} 个文件")
        self._pill("warn", f"正在建立索引 · 已收录 {count:,} 个文件")

    def _on_index_done(self, stats):
        """索引完成回调，显示文件数和各阶段耗时。"""
        self._engine_cancel = False
        self._index_running = False
        self.progress.stop()
        self.progress.pack_forget()
        self.index_btn.config(state=tk.NORMAL)
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
        self.index_btn.config(state=tk.NORMAL)
        self._set_status(f"索引出错: {err}", "error")
        self._pill("off", "索引出错 · 点击重试")

    def _do_index_silent(self):
        """静默后台重建索引，始终保证同一时间只有一个索引任务。"""
        if self._index_running or self._engine_cancel:
            return
        self._engine_cancel = False
        self._index_running = True
        self.index_btn.config(state=tk.DISABLED)
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
        theme_var = tk.StringVar(value=self._settings.get("theme", "dark"))

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
        row = _option_row(card_idx, "启动时自动更新索引", "每次启动软件后在后台静默重建全盘索引")
        ToggleSwitch(row, c, auto_start_var, width=s(46), height=s(26)).pack(
            side=tk.RIGHT, padx=(s(10), 0))
        row = _option_row(card_idx, "托盘自动更新", "最小化到系统托盘后按固定间隔更新索引")
        ToggleSwitch(row, c, tray_auto_var, width=s(46), height=s(26)).pack(
            side=tk.RIGHT, padx=(s(10), 0))

        # 间隔步进器
        row = _option_row(card_idx, "更新间隔", "托盘自动更新的时间间隔（5 ~ 120 分钟）")
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
            # 迷你界面预览
            if value == "system":
                half = w / 2
                cv.create_rectangle(3, 3, half, pv_h, fill="#0C1014", outline="")
                cv.create_rectangle(half, 3, w - 3, pv_h, fill="#F2F5F4", outline="")
                _rounded_rect(cv, s(10), s(10), w - s(10), s(26), 5,
                              fill="", outline=c["accent"], width=1.2)
            else:
                pv_bg = "#0C1014" if value == "dark" else "#F2F5F4"
                pv_bar = "#1C2630" if value == "dark" else "#DFE6E3"
                pv_box = "#131A21" if value == "dark" else "#FFFFFF"
                pv_row = "#1A222B" if value == "dark" else "#EFF3F2"
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

        for name, value in (("深色", "dark"), ("浅色", "light"), ("跟随系统", "system")):
            cv = tk.Canvas(theme_row, width=s(150), height=s(96), bd=0,
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
        tk.Label(about_text, text="File Searcher · 霁青版", bg=c["surface"], fg=c["text"],
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
        """清空搜索并恢复占位符和当前筛选下的全部结果。"""
        if self._search_timer is not None:
            self.root.after_cancel(self._search_timer)
            self._search_timer = None
        self._placeholder_visible = False
        self._set_search_value("")
        self.root.focus_set()
        self._show_placeholder()
        self._run_query("")
        return "break"

    def _on_search_changed(self):
        """真实搜索文字变化时触发（300ms 防抖延迟）。"""
        if self._suppress_search_trace or self._placeholder_visible:
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

    def _run_query(self, query: str, filters: dict | None = None):
        """统一执行搜索查询：分页拉取当前页并刷新表格与分页栏。"""
        filters = (filters if filters is not None else self._current_filters()).copy()
        self._last_query = query
        self._last_filters = filters
        if not IndexEngine.index_exists():
            self._results = []
            self._total_results = 0
            self._page = 1
            self._refresh_tree()
            self._update_result_status()
            self._set_status("请先创建索引再搜索", "warning")
            return

        self._set_loading(True)
        try:
            self._total_results = IndexEngine.result_count(query, filters=filters)
            total_pages = self._total_pages()
            self._page = max(1, min(self._page, total_pages))
            offset = (self._page - 1) * self._page_size
            if query:
                self._results = IndexEngine.search(
                    query, limit=self._page_size, offset=offset,
                    order_col=self._sort_col, order_desc=not self._sort_asc, filters=filters,
                )
            else:
                self._results = IndexEngine.load_all(
                    limit=self._page_size, offset=offset,
                    order_col=self._sort_col, order_desc=not self._sort_asc, filters=filters,
                )
            self._refresh_tree()
            self._update_sort_heading()
            self._update_result_status()
        finally:
            self._set_loading(False)

    def _update_result_status(self):
        """统一更新状态栏计数、分页信息与状态文字。"""
        total = self._total_results
        total_pages = self._total_pages()
        self.pager_total_var.set(f"共 {total:,} 个结果")
        self.page_info_var.set(f"{self._page} / {total_pages}")
        self.status_right_var.set(f"第 {self._page} / {total_pages} 页")
        if self._last_query:
            self._set_status(f"搜索「{self._last_query}」— 第 {self._page} / {total_pages} 页，共 {total:,} 个结果")
        else:
            self._set_status(f"第 {self._page} / {total_pages} 页，共 {total:,} 个结果")

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
        self._show_empty_state(not self._results)

    def _tree_item_values(self, result: dict):
        is_dir = bool(result.get("is_dir"))
        ext = os.path.splitext(result["name"])[1]
        type_text = "文件夹" if is_dir else (ext.upper().lstrip(".") if ext else "文件")
        size_text = "" if is_dir else format_size(result["size"])
        return (result["name"], result["path"], type_text, size_text, result["modified"])

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
        new_name = _dialog_input(self, "重命名", "输入新的文件名：", initial=old_name)
        if not new_name or new_name == old_name:
            return
        try:
            new_path = rename_file(path, new_name)
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
        """右键菜单 → 新开一个程序窗口。"""
        try:
            script = os.path.abspath(__file__)
            python = sys.executable
            flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            subprocess.Popen([python, script], creationflags=flags)
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

    def _remove_from_results(self, path: str):
        """从当前结果列表中移除指定路径并同步当前计数。"""
        before = len(self._results)
        self._results = [f for f in self._results if f["path"] != path]
        if len(self._results) < before:
            self._total_results = max(0, self._total_results - 1)

    # ================================================================
    #  拖拽到外部程序
    # ================================================================

    def _setup_drag_drop(self):
        """设置文件拖拽功能（依赖 tkdnd 库）。"""
        try:
            self.root.tk.call("package", "require", "tkdnd")
            self.root.tk.eval(f"tkdnd::drag_source register {self.table._w} DND_Files")
            self._dnd_cb = self.root.register(self._on_dnd_data)
            self.root.tk.eval(f"tkdnd::drag_source handler {self.table._w} drag {self._dnd_cb}")
            self._dnd_ok = True
        except tk.TclError:
            self._dnd_ok = False
            self.table.bind("<B1-Motion>", self._on_drag_fallback)

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
        """聚焦搜索框并全选真实搜索文字。"""
        self.search_entry.focus_set()
        self._hide_placeholder()
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
        """从 JSON 文件恢复上次的列宽（跨显示器按缩放系数换算）。

        旧版文件没有 "scale" 字段（旧显示器绝对值），无法可靠换算，
        直接忽略并采用当前显示器下的默认列宽。
        """
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
                cols.append((key, w))
            self.table._cols = cols
            self.table.redraw()
        except Exception:
            pass

    def _save_layout(self):
        """将当前列宽保存到 JSON 文件（含缩放系数，便于跨显示器还原）。"""
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
        """关闭窗口 → 最小化到系统托盘。"""
        self.root.withdraw()
        self._start_tray_auto_index_timer()

    # ================================================================
    #  系统托盘
    # ================================================================

    def _create_tray_icon(self) -> Image.Image:
        """用 Pillow 生成 32x32 的托盘图标（渐变底 + 白色放大镜）。"""
        img = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
        c = self.colors
        rgb1 = tuple(int(c["accent_grad_a"].lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
        rgb2 = tuple(int(c["accent_grad_b"].lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
        # 垂直渐变圆角方块底
        grad = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
        gd = ImageDraw.Draw(grad)
        for y in range(32):
            t = y / 31
            col = tuple(int(rgb1[i] + (rgb2[i] - rgb1[i]) * t) for i in range(3))
            gd.line([(0, y), (31, y)], fill=col + (255,))
        mask = Image.new("L", (32, 32), 0)
        ImageDraw.Draw(mask).rounded_rectangle((0, 0, 31, 31), radius=8, fill=255)
        grad.putalpha(mask)
        # 白色放大镜
        draw = ImageDraw.Draw(grad)
        draw.ellipse([9, 8, 21, 20], outline=(255, 255, 255, 255), width=3)
        draw.line([18.5, 18.5, 25, 25], fill=(255, 255, 255, 255), width=4)
        return grad

    def _setup_tray(self):
        """创建系统托盘图标和菜单。"""
        try:
            icon = self._create_tray_icon()
            try:
                self.root.iconphoto(True, ImageTk.PhotoImage(icon, master=self.root))
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

    def _on_minimize(self, event=None):
        """窗口最小化时隐藏到托盘，并按设置启动定时索引。"""
        if self.root.state() == "iconic":
            self.root.after(100, self.root.withdraw)
            self._start_tray_auto_index_timer()

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
        if self._frameless:
            self._maximize_to_workarea()
        else:
            try:
                if sys.platform == "win32":
                    self.root.state("zoomed")
                else:
                    self.root.attributes("-zoomed", True)
            except Exception:
                pass
        self.root.focus_force()
        # 搜索框全选真实文字；占位符状态只聚焦。
        self.search_entry.focus_set()
        if self._search_text():
            self.search_entry.select_range(0, tk.END)
        else:
            self._hide_placeholder()
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
        self._cancel_tray_auto_index_timer()
        self._save_layout()
        self.root.destroy()

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
