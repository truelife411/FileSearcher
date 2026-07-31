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
from tkinter import ttk, messagebox
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
# 全局字号体系（Tk points，100% 缩放下 12pt ≈ 16px）
FONT_XS = 10        # 最小辅助文字（保留兜底）
FONT_SMALL = 11     # 筛选标签 / 类型标签
FONT_BODY = 12      # 正文：表格、信息行、状态栏
FONT_LG = 13        # 按钮
FONT_INPUT = 14     # 搜索框输入
FONT_HEADER = 12    # 表头
FONT_TITLE = 12     # 标题栏标题
ROW_HEIGHT = 56     # 结果行高
TITLEBAR_H = 42     # 自绘标题栏高度
SPACING = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24}
THEMES = {
    "dark": {
        "bg": "#0E1116", "surface": "#171B22", "surface_alt": "#1F252E",
        "surface_3": "#262D38", "input": "#14181F", "border": "#2A3140",
        "border_strong": "#374052", "text": "#ECF0F7", "muted": "#8B94A7",
        "muted_2": "#5F6A7E", "accent": "#5B7CFA", "accent_hover": "#7A93FC",
        "accent_pressed": "#3D57C8", "accent_grad_a": "#6A8BFF", "accent_grad_b": "#4A66E8",
        "selected": "#2E4570", "selected_hover": "#365080", "row_alt": "#1A1F28",
        "success": "#4ADE80", "warning": "#F5B84C", "error": "#F26D6D",
        "title_bg": "#141820", "hover": "#262D38",
    },
    "light": {
        "bg": "#F2F4F8", "surface": "#FFFFFF", "surface_alt": "#E9EDF3",
        "surface_3": "#DFE5EE", "input": "#FFFFFF", "border": "#D5DBE5",
        "border_strong": "#B9C3D2", "text": "#141B26", "muted": "#5F6B7E",
        "muted_2": "#8A94A6", "accent": "#3B6CF0", "accent_hover": "#5B84F5",
        "accent_pressed": "#2C57D4", "accent_grad_a": "#5B84F5", "accent_grad_b": "#3B6CF0",
        "selected": "#DCE7FB", "selected_hover": "#C9DAF8", "row_alt": "#F5F7FB",
        "success": "#16A34A", "warning": "#B45309", "error": "#DC2626",
        "title_bg": "#E7EBF2", "hover": "#E3E9F2",
    },
}

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
    """用系统默认软件打开文件"""
    try:
        if sys.platform == "win32":
            os.startfile(os.path.normpath(path))
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError as e:
        messagebox.showerror("打开失败", str(e))


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

    kind: "normal"（灰底） | "accent"（靛蓝渐变主按钮） | "danger"（红色系）
    """

    def __init__(self, master, text="", command=None, width=120, height=48, radius=12,
                 colors=None, icon=None, kind="normal", font_size=None, **kwargs):
        self.colors = colors or THEMES["dark"]
        self._kind = kind
        self._font_size = font_size or FONT_LG
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
                fill, outline, fg = "#A62B2B", "#A62B2B", "#FFFFFF"
            elif self._visual_state == "hover":
                fill, outline, fg = "#C42B2B", "#C42B2B", "#FFFFFF"
            else:
                fill, outline, fg = c["error"], c["error"], "#FFFFFF"
        elif self._visual_state == "pressed":
            fill, outline, fg = c["surface_3"], c["border_strong"], c["text"]
        elif self._visual_state == "hover":
            fill, outline, fg = c["surface_3"], c["accent"], c["text"]
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
    """大号圆角搜索外壳：内嵌 Entry、搜索图标、聚焦高亮和清空热区。"""

    def __init__(self, master, textvariable, colors, clear_command, height=58, font_size=None):
        super().__init__(master, height=height, bd=0, highlightthickness=0,
                         bg=master.cget("bg"))
        self.colors = colors
        self._focused = False
        self._hover = False
        self._font_size = font_size or FONT_INPUT
        self.entry = tk.Entry(self, textvariable=textvariable, relief="flat", bd=0,
                              bg=colors["input"], fg=colors["text"],
                              insertbackground=colors["text"], font=(FONT_FAMILY, self._font_size))
        self._entry_window = self.create_window(50, height / 2, window=self.entry, anchor=tk.W)
        self._clear_command = clear_command
        self.bind("<Configure>", self._layout)
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda _e: self._on_hover(True))
        self.bind("<Leave>", lambda _e: self._on_hover(False))
        self.entry.bind("<FocusIn>", self._focus_in, add="+")
        self.entry.bind("<FocusOut>", self._focus_out, add="+")
        self._draw()

    def _layout(self, _event=None):
        self.itemconfigure(self._entry_window, width=max(20, self.winfo_width() - 96), height=30)
        self._draw()

    def _on_hover(self, hover: bool):
        self._hover = hover
        self._draw()

    def _draw(self):
        self.delete("shell")
        width, height = max(2, self.winfo_width()), max(2, self.winfo_height())
        c = self.colors
        if self._focused:
            border, glow = c["accent"], 1
        elif self._hover:
            border, glow = c["border_strong"], 0
        else:
            border, glow = c["border"], 0
        shell = _rounded_rect(self, 1.5, 1.5, width - 1.5, height - 1.5, 13,
                              fill=c["input"], outline=border, width=1.5, tags="shell")
        self.tag_lower(shell)
        if glow:
            _rounded_rect(self, 3, 3, width - 3, height - 3, 12,
                          outline=c["accent"], width=1, tags="shell")
        # 搜索图标（放大镜）
        cx, cy = 23, height / 2
        self.create_oval(cx - 6.5, cy - 6.5, cx + 6.5, cy + 6.5,
                         outline=c["muted"], width=1.8, tags="shell")
        self.create_line(cx + 5.5, cy + 5.5, cx + 11, cy + 11,
                         fill=c["muted"], width=2.2, tags="shell")
        # 清空按钮（hover 时显示）
        clear_x = width - 26
        if self._hover or self._focused:
            self.create_oval(clear_x - 11, cy - 11, clear_x + 11, cy + 11,
                             fill=c["surface_3"], outline="", tags="shell")
            self.create_text(clear_x, cy, text="✕", fill=c["muted"],
                             font=(FONT_FAMILY, FONT_SMALL), tags="shell")
        self.tag_lower("shell", self._entry_window)

    def _click(self, event):
        if event.x >= self.winfo_width() - 46:
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
        self._has_more = False
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

        self.root.after(100, self._load_all)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Unmap>", self._on_minimize)
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
        self.path_filter_var = tk.StringVar(value="全部")
        self.type_filter_var = tk.StringVar(value="全部")
        self.time_filter_var = tk.StringVar(value="全部")
        self.size_filter_var = tk.StringVar(value="全部")
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
        saved = {
            "query": query,
            "path": self.path_filter_var.get() if hasattr(self, "path_filter_var") else "全部",
            "type": self.type_filter_var.get() if hasattr(self, "type_filter_var") else "全部",
            "time": self.time_filter_var.get() if hasattr(self, "time_filter_var") else "全部",
            "size": self.size_filter_var.get() if hasattr(self, "size_filter_var") else "全部",
        }
        for w in self.root.winfo_children():
            w.destroy()
        self._build_ui()
        if hasattr(self, "path_filter_var"):
            self.path_filter_var.set(saved["path"])
            self.type_filter_var.set(saved["type"])
            self.time_filter_var.set(saved["time"])
            self.size_filter_var.set(saved["size"])
        self._placeholder_visible = False
        if saved["query"]:
            self._set_search_value(saved["query"])
            self.search_entry.configure(fg=self.colors["text"])
        else:
            self._show_placeholder()
        self._refresh_tree()
        self._update_sort_heading()
        self._update_result_status()
        self._update_filter_button()

    def _ensure_maximized(self):
        """启动后强制铺满工作区（无边框模式）。"""
        if self._frameless and self._normal_rect is None:
            self._maximize_to_workarea()

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
        style.configure("Status.TLabel", background=c["surface"], foreground=c["muted"], padding=(10, 6))
        style.configure("TButton", background=c["surface_alt"], foreground=c["text"], bordercolor=c["border"], padding=(self._s(14), self._s(9)))
        style.map("TButton", background=[("active", c["border_strong"]), ("disabled", c["surface"])], foreground=[("disabled", c["muted_2"])])
        style.configure("Accent.TButton", background=c["accent"], foreground="#FFFFFF", bordercolor=c["accent"], font=self._f(FONT_BODY, "bold"))
        style.map("Accent.TButton", background=[("active", c["accent_hover"]), ("disabled", c["border"])])
        style.configure("TEntry", fieldbackground=c["input"], foreground=c["text"], insertcolor=c["text"], bordercolor=c["border"], padding=10)
        style.configure("Filter.TCombobox", fieldbackground=c["surface_2"] if "surface_2" in c else c["surface_alt"],
                        background=c["surface_alt"], foreground=c["text"], arrowcolor=c["muted"],
                        bordercolor=c["border"], lightcolor=c["border"], darkcolor=c["border"],
                        padding=(self._s(10), self._s(8)), font=self._f(FONT_BODY))
        style.map("Filter.TCombobox", fieldbackground=[("readonly", c["surface_2"] if "surface_2" in c else c["surface_alt"])],
                  foreground=[("readonly", c["text"])], bordercolor=[("focus", c["accent"])])
        style.configure("Results.Treeview", background=c["surface"], fieldbackground=c["surface"],
                        foreground=c["text"], borderwidth=0, relief="flat", rowheight=self._s(ROW_HEIGHT),
                        font=self._f(FONT_BODY))
        style.map("Results.Treeview", background=[("selected", c["selected"])], foreground=[("selected", c["text"])])
        style.layout("Results.Treeview.Heading", [])
        style.configure("Ex.Treeview", background=c["surface"], fieldbackground=c["surface"],
                        foreground=c["text"], borderwidth=0, relief="flat", rowheight=self._s(36),
                        font=self._f(FONT_BODY))
        style.map("Ex.Treeview", background=[("selected", c["selected"])], foreground=[("selected", c["text"])])
        style.configure("Ex.Treeview.Heading", background=c["surface_alt"], foreground=c["text"],
                        font=self._f(FONT_SMALL, "bold"), relief="flat", padding=(8, 6))
        style.map("Ex.Treeview.Heading", background=[("active", c["surface_3"])])
        style.configure("Vertical.TScrollbar", background=c["surface_3"], troughcolor=c["surface"],
                        bordercolor=c["surface"], arrowcolor=c["muted_2"], width=self._s(12))
        style.configure("TProgressbar", troughcolor=c["surface_alt"], background=c["accent"], bordercolor=c["surface_alt"])
        style.configure("TCheckbutton", background=c["bg"], foreground=c["text"], font=self._f(FONT_BODY))
        style.configure("TSeparator", background=c["border"])

    # ================================================================
    #  自绘标题栏 + 无边框窗口（Windows）
    # ================================================================

    def _build_titlebar(self):
        """构建自绘标题栏：渐变放大镜 logo、标题、最小化/最大化/关闭按钮。"""
        c = self.colors
        self._titlebar_h = self._s(TITLEBAR_H)
        self._tb_buttons = []
        self._tb_hit_rects = []
        bar = tk.Frame(self.root, bg=c["title_bg"], height=self._titlebar_h)
        bar.pack(fill=tk.X, side=tk.TOP)
        bar.pack_propagate(False)
        self._titlebar = bar

        # logo：用 Canvas 画放大镜
        logo = tk.Canvas(bar, width=self._s(26), height=self._s(26), bd=0, highlightthickness=0,
                         bg=c["title_bg"])
        logo.pack(side=tk.LEFT, padx=(self._s(16), self._s(10)), pady=(self._titlebar_h - self._s(26)) // 2)
        grad = _make_gradient_pix(self.root, 2, self._s(26), 0, c["accent_grad_a"], c["accent_grad_b"])
        logo.create_image(1, self._s(13), image=grad)
        s = self.ui_scale
        logo.create_oval(4 * s, 4 * s, 16 * s, 16 * s, outline=c["title_bg"], width=max(1.5, 2.2 * s))
        logo.create_line(14.5 * s, 14.5 * s, 20 * s, 20 * s, fill=c["title_bg"], width=max(2, 2.6 * s), capstyle=tk.ROUND)

        tk.Label(bar, text="File Searcher", bg=c["title_bg"], fg=c["text"],
                 font=self._f(FONT_TITLE, "bold")).pack(side=tk.LEFT)

        def _make_tb_btn(text, hover_bg=None, command=None):
            btn = tk.Label(bar, text=text, bg=c["title_bg"], fg=c["muted"],
                           font=self._f(FONT_TITLE, "normal"), width=4, cursor="hand2")
            btn.pack(side=tk.RIGHT)
            btn.bind("<Enter>", lambda _e: btn.configure(bg=hover_bg or c["surface_3"], fg=c["text"]))
            btn.bind("<Leave>", lambda _e: btn.configure(bg=c["title_bg"], fg=c["muted"]))
            if command:
                btn.bind("<Button-1>", lambda _e: command())
            self._tb_buttons.append(btn)
            return btn

        # 右侧按钮从右往左：关闭 → 最小化（程序启动即最大化，无需最大化按钮）
        _make_tb_btn("✕", hover_bg="#C42B2B", command=self._on_close)
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
        """构建大号搜索行、带状态徽章的信息行和卡片式筛选面板。"""
        c = self.colors
        header = tk.Frame(self.root, bg=c["bg"])
        header.pack(fill=tk.X, padx=self._s(24), pady=(self._s(20), self._s(12)))
        header.columnconfigure(0, weight=1)

        search_row = tk.Frame(header, bg=c["bg"], height=self._s(58))
        search_row.grid(row=0, column=0, sticky="ew")
        search_row.grid_propagate(False)
        search_row.columnconfigure(0, weight=1)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._on_search_changed())
        self.search_box = RoundedSearchBox(search_row, self.search_var, c, self._clear_search,
                                           height=self._s(58), font_size=self._f(FONT_INPUT)[1])
        self.search_box.grid(row=0, column=0, sticky="nsew", padx=(0, self._s(12)))
        self.search_entry = self.search_box.entry
        self.search_entry.bind("<Return>", self._do_search)
        self.search_entry.bind("<FocusIn>", self._hide_placeholder, add="+")
        self.search_entry.bind("<FocusOut>", self._show_placeholder, add="+")
        self._placeholder = "搜索文件名或完整路径…"
        self._placeholder_visible = False
        self._show_placeholder()

        self.filter_btn = RoundedButton(search_row, text="筛选", command=self._toggle_filters,
                                        width=self._s(104), height=self._s(58), colors=c,
                                        font_size=self._f(FONT_LG)[1])
        self.filter_btn.grid(row=0, column=1, padx=(0, self._s(12)))
        self.index_btn = RoundedButton(search_row, text="创建索引", command=self._toggle_index,
                                       width=self._s(138), height=self._s(58), colors=c, kind="accent",
                                       font_size=self._f(FONT_LG)[1])
        self.index_btn.grid(row=0, column=2, padx=(0, self._s(12)))
        self.settings_btn = RoundedButton(search_row, icon="⚙", command=self._open_settings,
                                          width=self._s(58), height=self._s(58), colors=c,
                                          font_size=self._f(FONT_LG)[1])
        self.settings_btn.grid(row=0, column=3)

        # 信息行：左侧索引状态徽章（彩色圆点），右侧结果计数
        info_row = tk.Frame(header, bg=c["bg"], height=self._s(30))
        info_row.grid(row=1, column=0, sticky="ew", pady=(self._s(12), 0))
        info_row.grid_propagate(False)
        info_row.columnconfigure(0, weight=1)
        self.index_status_var = tk.StringVar(value="尚未创建索引")
        self.index_count_var = tk.StringVar(value="0 项")
        self.index_updated_var = tk.StringVar(value="未更新")
        self.index_info_var = tk.StringVar(value="尚未创建索引 · 0 项 · 未更新")
        self.result_count_var = tk.StringVar(value="0 个结果")
        self._status_dot = tk.Label(info_row, text="●", bg=c["bg"], fg=c["muted_2"],
                                    font=self._f(FONT_SMALL))
        self._status_dot.grid(row=0, column=0, sticky="w")
        tk.Label(info_row, textvariable=self.index_info_var, bg=c["bg"], fg=c["muted"],
                 font=self._f(FONT_BODY), anchor=tk.W).grid(row=0, column=1, sticky="w", padx=(self._s(8), 0))
        tk.Label(info_row, textvariable=self.result_count_var, bg=c["bg"], fg=c["muted"],
                 font=self._f(FONT_BODY), anchor=tk.E).grid(row=0, column=2, sticky="e")

        # 筛选面板：卡片式四列
        self.filter_panel = tk.Frame(header, bg=c["surface"], highlightthickness=1,
                                     highlightbackground=c["border"], padx=self._s(16), pady=self._s(14))
        for column in range(4):
            self.filter_panel.columnconfigure(column, weight=1, uniform="filter")
        specs = (
            ("位置", self.path_filter_var, list(self._path_options)),
            ("类型", self.type_filter_var, ["全部", "文件夹", "文档", "图片", "视频", "音频", "压缩包", "代码"]),
            ("时间", self.time_filter_var, ["全部", "今天", "近7天", "近30天"]),
            ("大小", self.size_filter_var, ["全部", "<1MB", "1-100MB", ">100MB"]),
        )
        for column, (label, variable, values) in enumerate(specs):
            cell = tk.Frame(self.filter_panel, bg=c["surface"])
            cell.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else self._s(8), self._s(8) if column < 3 else 0))
            tk.Label(cell, text=label, bg=c["surface"], fg=c["muted_2"],
                     font=self._f(FONT_SMALL)).pack(anchor=tk.W, pady=(0, self._s(6)))
            box = ttk.Combobox(cell, textvariable=variable, values=values, state="readonly",
                               style="Filter.TCombobox", height=8)
            box.pack(fill=tk.X)
            box.bind("<<ComboboxSelected>>", self._on_filter_changed)
        self._filters_visible = False
        self._update_filter_button()

    def _toggle_filters(self):
        self._filters_visible = not self._filters_visible
        if self._filters_visible:
            self.filter_panel.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        else:
            self.filter_panel.grid_remove()

    def _on_filter_changed(self, _event=None):
        self._update_filter_button()
        self._do_search()

    def _update_filter_button(self):
        count = sum(var.get() != "全部" for var in (
            self.path_filter_var, self.type_filter_var, self.time_filter_var, self.size_filter_var
        ))
        self.filter_btn.config(text=f"筛选 {count}" if count else "筛选")

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
        return {
            "path_prefix": self._path_options.get(self.path_filter_var.get()),
            "type": self.type_filter_var.get(),
            "time": self.time_filter_var.get(),
            "size": self.size_filter_var.get(),
        }

    def _build_tree(self):
        """构建圆角结果容器、自绘表头、hover 高亮与保留原生交互的 Treeview。"""
        c = self.colors
        outer = tk.Frame(self.root, bg=c["bg"])
        outer.pack(fill=tk.BOTH, expand=True, padx=self._s(24), pady=(0, self._s(14)))
        self.result_canvas = tk.Canvas(outer, bd=0, highlightthickness=0, bg=c["bg"])
        self.result_canvas.pack(fill=tk.BOTH, expand=True)
        self.result_surface = tk.Frame(self.result_canvas, bg=c["surface"])
        self._result_window = self.result_canvas.create_window(1, 1, window=self.result_surface, anchor=tk.NW)
        self.result_surface.bind("<Configure>", lambda _e: self._draw_tree_header())
        self.result_canvas.bind("<Configure>", self._layout_result_container)

        self.header_canvas = tk.Canvas(self.result_surface, height=self._s(44), bd=0, highlightthickness=0,
                                       bg=c["surface_alt"], cursor="hand2")
        self.header_canvas.pack(fill=tk.X, padx=1, pady=(1, 0))
        self.header_canvas.bind("<Configure>", lambda _e: self._draw_tree_header(True))
        self.header_canvas.bind("<Button-1>", self._on_header_click)
        self.header_canvas.bind("<Motion>", self._on_header_motion)
        self.header_canvas.bind("<Leave>", lambda _e: self._set_header_hover(None))
        self._header_hover_col = None

        body = tk.Frame(self.result_surface, bg=c["surface"])
        body.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0, 1))
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        columns = ("name", "path", "type", "size", "modified")
        self.tree = ttk.Treeview(body, columns=columns, show="tree headings", selectmode="extended",
                                 style="Results.Treeview")
        self.tree.heading("#0", text="")
        for col in columns:
            self.tree.heading(col, text="")
        self.tree.column("#0", width=self._s(44), minwidth=self._s(44), stretch=False, anchor=tk.CENTER)
        self.tree.column("name", width=self._s(320), minwidth=self._s(200))
        self.tree.column("path", width=self._s(600), minwidth=self._s(280))
        self.tree.column("type", width=self._s(130), minwidth=self._s(100), anchor=tk.CENTER)
        self.tree.column("size", width=self._s(150), minwidth=self._s(110), anchor=tk.E)
        self.tree.column("modified", width=self._s(200), minwidth=self._s(160), anchor=tk.CENTER)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y = ttk.Scrollbar(body, orient=tk.VERTICAL, style="Vertical.TScrollbar",
                                    command=self._on_tree_scroll_wrapper)
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        self.tree.tag_configure("odd", background=c["surface"])
        self.tree.tag_configure("even", background=c["row_alt"])
        self.tree.tag_configure("hover", background=c["hover"])
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<MouseWheel>", lambda _e: self.root.after_idle(self._on_tree_scroll))
        self.tree.bind("<Configure>", lambda _e: self._draw_tree_header(True))
        self._hover_iid = None
        self._hover_tag = "odd"
        self.tree.bind("<Motion>", self._on_tree_motion)
        self.tree.bind("<Leave>", self._on_tree_leave)

        # 空状态：放大镜图标 + 主文案 + 副文案
        empty_frame = tk.Frame(body, bg=c["surface"])
        icon_size = self._s(56)
        self._empty_icon = tk.Canvas(empty_frame, width=icon_size, height=icon_size, bd=0,
                                     highlightthickness=0, bg=c["surface"])
        self._empty_icon.pack()
        s = self.ui_scale
        self._empty_icon.create_oval(10 * s, 10 * s, 40 * s, 40 * s, outline=c["muted_2"],
                                     width=max(2, 3 * s))
        self._empty_icon.create_line(37 * s, 37 * s, 50 * s, 50 * s, fill=c["muted_2"],
                                     width=max(3, 4 * s), capstyle=tk.ROUND)
        tk.Label(empty_frame, text="没有匹配的结果", bg=c["surface"], fg=c["muted"],
                 font=self._f(FONT_LG)).pack(pady=(self._s(14), self._s(4)))
        tk.Label(empty_frame, text="换个关键词或调整筛选条件试试", bg=c["surface"], fg=c["muted_2"],
                 font=self._f(FONT_SMALL)).pack()
        self.empty_state = empty_frame

    def _on_tree_motion(self, event):
        """鼠标悬停行高亮（基于 tag 的 hover 效果）。"""
        iid = self.tree.identify_row(event.y)
        if iid == self._hover_iid:
            return
        self._restore_hover()
        if iid:
            try:
                tags = self.tree.item(iid, "tags")
                self._hover_tag = tags[0] if tags else "odd"
                self.tree.item(iid, tags=("hover",))
            except Exception:
                pass
        self._hover_iid = iid

    def _on_tree_leave(self, _event=None):
        self._restore_hover()

    def _restore_hover(self):
        if self._hover_iid:
            try:
                self.tree.item(self._hover_iid, tags=(self._hover_tag,))
            except Exception:
                pass
        self._hover_iid = None

    def _on_header_motion(self, event):
        """表头悬停列高亮。"""
        x = event.x + self.header_canvas.canvasx(0)
        current = 0
        col = None
        for c, width in self._tree_column_layout():
            if current <= x < current + width:
                col = c
                break
            current += width
        if col != self._header_hover_col:
            self._header_hover_col = col
            self._draw_tree_header()

    def _set_header_hover(self, col):
        """表头鼠标移出时清除悬停高亮。"""
        self._header_hover_col = col
        self._draw_tree_header()

    def _layout_result_container(self, event):
        if event.width < 4 or event.height < 4:
            return
        self.result_canvas.itemconfigure(self._result_window, width=event.width - 2, height=event.height - 2)
        self.result_canvas.delete("container")
        shape = _rounded_rect(self.result_canvas, 1, 1, event.width - 1, event.height - 1, 12,
                              fill=self.colors["surface"], outline=self.colors["border"],
                              width=1, tags="container")
        self.result_canvas.tag_lower(shape)

    def _tree_column_layout(self, force: bool = False):
        """计算列布局；hover 重绘等非布局变化时使用缓存，避免列宽抖动。"""
        if force or getattr(self, "_col_layout", None) is None:
            available = max(1, self.header_canvas.winfo_width() - 11)
            fixed = {
                "type": self.tree.column("type", "width"),
                "size": self.tree.column("size", "width"),
                "modified": self.tree.column("modified", "width"),
            }
            flexible = max(420, available - sum(fixed.values()))
            name_width = max(200, int(flexible * 0.34))
            path_width = max(240, flexible - name_width)
            self.tree.column("name", width=name_width - self.tree.column("#0", "width"))
            self.tree.column("path", width=path_width)
            self._col_layout = [("name", name_width), ("path", path_width), *fixed.items()]
        return self._col_layout

    def _draw_tree_header(self, force: bool = False):
        if not hasattr(self, "header_canvas"):
            return
        c = self.colors
        self.header_canvas.delete("all")
        labels = {"name": "文件名", "path": "路径", "type": "类型", "size": "大小", "modified": "修改时间"}
        x = 0
        for col, width in self._tree_column_layout(force):
            anchor = tk.E if col == "size" else (tk.CENTER if col in ("type", "modified") else tk.W)
            if col == "name":
                text_x = x + self.tree.column("#0", "width") + 16
            elif anchor == tk.E:
                text_x = x + width - 16
            elif anchor == tk.CENTER:
                text_x = x + width / 2
            else:
                text_x = x + 16
            if col == self._sort_col:
                fg = c["accent_hover"]
                label = labels[col] + ("  ▲" if self._sort_asc else "  ▼")
            else:
                fg = c["muted"]
                label = labels[col]
            if col == self._header_hover_col:
                fg = c["text"]
            cy = self._s(22)
            self.header_canvas.create_text(text_x, cy, text=label, anchor=anchor, fill=fg,
                                           font=self._f(FONT_HEADER, "bold"))
            if x:
                self.header_canvas.create_line(x, self._s(10), x, self._s(34), fill=c["border"])
            x += width
        self.header_canvas.configure(scrollregion=(0, 0, x, self._s(44)))

    def _on_header_click(self, event):
        x = event.x + self.header_canvas.canvasx(0)
        current = 0
        for col, width in self._tree_column_layout():
            if current <= x < current + width:
                self._sort_by(col)
                return
            current += width

    def _build_context_menu(self):
        """构建与主题一致的右键菜单。"""
        c = self.colors
        self._ctx_menu = tk.Menu(self.root, tearoff=0, bg=c["surface_alt"], fg=c["text"],
                                 activebackground=c["selected"], activeforeground=c["text"],
                                 bd=0, relief="flat", font=self._f(FONT_BODY))
        self._ctx_menu.add_command(label="打开", command=self._open_selected)
        self._ctx_menu.add_command(label="打开所在文件夹", command=self._open_file_location_selected)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="复制", command=self._copy_path)
        self._ctx_menu.add_command(label="剪切", command=self._cut_path)
        self._ctx_menu.add_command(label="复制完整路径", command=self._copy_full_path_text)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="重命名", command=self._rename_file_dialog)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="删除到回收站", command=self._delete_file_recycle)
        self._ctx_menu.add_command(label="彻底删除", command=self._delete_file_permanent)

    def _build_statusbar(self):
        """构建固定高度、顶部细分隔线的左右状态栏。"""
        c = self.colors
        bar = tk.Frame(self.root, bg=c["surface"], height=self._s(46),
                       highlightthickness=1, highlightbackground=c["border"])
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)
        bar.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="就绪 — 请先创建索引")
        self._statusbar_dot = tk.Label(bar, text="●", bg=c["surface"], fg=c["muted_2"],
                                       font=self._f(FONT_SMALL))
        self._statusbar_dot.grid(row=0, column=0, sticky="w", padx=(self._s(24), self._s(8)))
        self.status_label = tk.Label(bar, textvariable=self.status_var, bg=c["surface"], fg=c["muted"],
                                     font=self._f(FONT_BODY), anchor=tk.W)
        self.status_label.grid(row=0, column=1, sticky="nsew")
        self.progress_slot = tk.Frame(bar, bg=c["surface"], width=self._s(150), height=self._s(44))
        self.progress_slot.grid(row=0, column=2, sticky="ns")
        self.progress_slot.grid_propagate(False)
        self.progress = ttk.Progressbar(self.progress_slot, mode="indeterminate", length=self._s(130))
        self.status_right_var = tk.StringVar(value="0 个结果")
        tk.Label(bar, textvariable=self.status_right_var, bg=c["surface"], fg=c["muted"],
                 font=self._f(FONT_BODY), anchor=tk.E, padx=self._s(24)).grid(row=0, column=3, sticky="nsew")
        bar.rowconfigure(0, weight=1)

    def _set_status(self, text: str, kind: str = "normal"):
        self.status_var.set(text)
        color = {"success": self.colors["success"], "error": self.colors["error"],
                 "warning": self.colors["warning"]}.get(kind, self.colors["muted"])
        self.status_label.configure(foreground=color)
        if hasattr(self, "_statusbar_dot"):
            self._statusbar_dot.configure(foreground=color)

    def _set_index_dot(self, kind: str):
        """顶部信息行状态徽章：ok 绿 / warn 黄 / off 灰。"""
        if not hasattr(self, "_status_dot"):
            return
        color = {"ok": self.colors["success"], "warn": self.colors["warning"],
                 "off": self.colors["muted_2"]}.get(kind, self.colors["muted_2"])
        self._status_dot.configure(foreground=color)

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
        """更新索引按钮、就绪状态、数量和更新时间。"""
        if IndexEngine.index_exists():
            count = IndexEngine.index_file_count()
            updated = datetime.fromtimestamp(INDEX_DB.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            self.index_btn.config(text="重建索引")
            self.index_status_var.set("索引就绪")
            self.index_count_var.set(f"{count:,} 项")
            self.index_updated_var.set(f"更新于 {updated}")
            self.index_info_var.set(f"索引就绪 · {count:,} 个文件 · 更新于 {updated}")
            self._set_index_dot("ok")
        else:
            self.index_btn.config(text="创建索引")
            self.index_status_var.set("尚未创建索引")
            self.index_count_var.set("0 项")
            self.index_updated_var.set("未更新")
            self.index_info_var.set("尚未创建索引 · 0 个文件 · 未更新")
            self._set_index_dot("off")

    def _toggle_index(self):
        """点击索引按钮：创建或重建索引。"""
        if self._index_running or self._engine_cancel:
            return
        if IndexEngine.index_exists():
            if not messagebox.askyesno("确认", "重建索引将扫描所有磁盘，可能需要几分钟。继续？"):
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
        self._set_index_dot("warn")

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
        self._set_index_dot("warn")
        self.index_count_var.set(f"已收录 {count:,} 项")
        self.index_info_var.set(f"正在建立索引 · 已收录 {count:,} 个文件")

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
        self._set_index_dot("off")

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
        """弹出设置对话框：左侧圆角导航、卡片式内容，修改即保存。"""
        dlg = tk.Toplevel(self.root)
        dlg.title("设置")
        dlg.resizable(True, True)
        dlg.grab_set()
        dlg.transient(self.root)

        try:
            scale = float(self.root.tk.call('tk', 'scaling'))
        except Exception:
            scale = 1.0
        s = max(1.0, scale)
        dw = int(780 * s)
        dh = int(560 * s)
        dlg.minsize(int(620 * s), int(460 * s))
        dlg.configure(bg=self.colors["bg"])
        dlg.update_idletasks()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        px, py = self.root.winfo_rootx(), self.root.winfo_rooty()
        dlg.geometry(f"{dw}x{dh}+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

        c = self.colors
        body = tk.Frame(dlg, bg=c["bg"])
        body.pack(fill=tk.BOTH, expand=True, padx=int(16 * s), pady=int(16 * s))

        # ====== 左侧导航卡片 ======
        NAV_W = int(172 * s)
        nav_panel = tk.Frame(body, bg=c["surface"], highlightthickness=1,
                             highlightbackground=c["border"], width=NAV_W)
        nav_panel.pack(side=tk.LEFT, fill=tk.Y)
        nav_panel.pack_propagate(False)

        tk.Label(nav_panel, text="设置", bg=c["surface"], fg=c["text"],
                 font=(FONT_FAMILY, FONT_LG + 2, "bold")).pack(anchor=tk.W, padx=20, pady=(22, 20))

        nav_items = []

        def _make_nav_item(text):
            f = tk.Frame(nav_panel, bg=c["surface"], cursor="hand2")
            f.pack(fill=tk.X, padx=10, pady=3)
            cv = tk.Canvas(f, height=42, bd=0, highlightthickness=0, bg=c["surface"])
            cv.pack(fill=tk.X)
            cv._nav_text = text
            nav_items.append(cv)
            return f, cv

        nav_idx, idx_cv = _make_nav_item("索引设置")
        nav_ex, ex_cv = _make_nav_item("排除列表")

        def _draw_nav(cv, active):
            cv.delete("all")
            w = max(2, cv.winfo_width())
            h = 42
            if active:
                cv.create_rectangle(0, 1, w, h - 1, fill=c["surface_3"], outline="")
                cv.create_rectangle(0, 8, 3.5, h - 8, fill=c["accent"], outline="")
            cv.create_text(17, h / 2, anchor=tk.W, text=cv._nav_text,
                           fill=c["text"] if active else c["muted"],
                           font=(FONT_FAMILY, FONT_BODY, "bold" if active else "normal"))

        # ====== 右侧内容卡片 ======
        content = tk.Frame(body, bg=c["surface"], highlightthickness=1,
                           highlightbackground=c["border"])
        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(int(14 * s), 0))

        page_index = tk.Frame(content, bg=c["surface"])
        page_exclude = tk.Frame(content, bg=c["surface"])

        # ---- 索引设置页 ----
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
            self._settings["theme"] = theme_var.get()
            IndexEngine.save_settings(self._settings)

        auto_start_var.trace_add("write", lambda *_: _persist_general())
        tray_auto_var.trace_add("write", lambda *_: _persist_general())
        minutes_var.trace_add("write", lambda *_: _persist_general())
        theme_var.trace_add("write", lambda *_: _persist_general())

        pad = int(30 * s)

        def _section_title(text):
            tk.Label(page_index, text=text, bg=c["surface"], fg=c["text"],
                     font=(FONT_FAMILY, FONT_BODY + 1, "bold")).pack(anchor=tk.W, pady=(0, 10))

        tk.Label(page_index, text="常规", bg=c["surface"], fg=c["text"],
                 font=(FONT_FAMILY, FONT_LG + 2, "bold")).pack(anchor=tk.W, padx=pad, pady=(pad, 22))

        _section_title("启动时自动更新索引")
        cb1 = ttk.Checkbutton(page_index, text="启动后自动重建全盘文件索引",
                              variable=auto_start_var)
        cb1.pack(anchor=tk.W, padx=(pad + 18, 0), pady=(0, 20))

        _section_title("托盘自动更新")
        cb2 = ttk.Checkbutton(page_index, text="最小化到托盘后，自动更新索引",
                              variable=tray_auto_var)
        cb2.pack(anchor=tk.W, padx=(pad + 18, 0), pady=(0, 8))
        minutes_frame = tk.Frame(page_index, bg=c["surface"])
        minutes_frame.pack(anchor=tk.W, padx=(pad + 18, 0))
        tk.Label(minutes_frame, text="间隔", bg=c["surface"], fg=c["muted"],
                 font=(FONT_FAMILY, FONT_BODY)).pack(side=tk.LEFT, padx=(0, 8))
        spin = ttk.Spinbox(minutes_frame, from_=5, to=120, width=6, textvariable=minutes_var)
        spin.pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(minutes_frame, text="分钟（5~120）", bg=c["surface"], fg=c["muted"],
                 font=(FONT_FAMILY, FONT_BODY)).pack(side=tk.LEFT)

        tk.Frame(page_index, bg=c["border"], height=1).pack(fill=tk.X, padx=pad, pady=20)
        _section_title("界面主题")
        theme_frame = tk.Frame(page_index, bg=c["surface"])
        theme_frame.pack(anchor=tk.W, padx=(pad + 18, 0))
        for text, value in (("深色", "dark"), ("浅色", "light"), ("跟随系统", "system")):
            ttk.Radiobutton(theme_frame, text=text, value=value, variable=theme_var).pack(
                side=tk.LEFT, padx=(0, 20))
        tk.Label(page_index, text="主题修改已保存，重启程序后完全生效。", bg=c["surface"],
                 fg=c["muted_2"], font=(FONT_FAMILY, FONT_SMALL)).pack(
            anchor=tk.W, padx=(pad + 18, 0), pady=(10, 0))

        # ---- 排除列表页 ----
        tk.Label(page_exclude, text="索引时跳过匹配的目录（修改即时保存，需重建索引生效）",
                 bg=c["surface"], fg=c["muted_2"], font=(FONT_FAMILY, FONT_BODY)).pack(
            anchor=tk.W, padx=pad, pady=(pad, 12))

        tb = tk.Frame(page_exclude, bg=c["surface"])
        tb.pack(fill=tk.X, padx=pad, pady=(0, 10))
        RoundedButton(tb, text="＋ 添加", command=lambda: self._exclude_add(ex_list, dlg),
                      width=96, height=36, colors=c).pack(side=tk.LEFT, padx=(0, 8))
        RoundedButton(tb, text="✎ 编辑", command=lambda: self._exclude_edit(ex_list, dlg),
                      width=96, height=36, colors=c).pack(side=tk.LEFT, padx=(0, 8))
        RoundedButton(tb, text="✕ 删除", command=lambda: self._exclude_delete(ex_list),
                      width=96, height=36, colors=c, kind="danger").pack(side=tk.LEFT)

        ex_frame = tk.Frame(page_exclude, bg=c["surface"])
        ex_frame.pack(fill=tk.BOTH, expand=True, padx=pad, pady=(0, pad))

        columns = ("type", "value")
        ex_list = ttk.Treeview(ex_frame, columns=columns, show="headings", selectmode="browse",
                               style="Ex.Treeview", height=14)
        ex_list.heading("type", text="类型")
        ex_list.heading("value", text="排除内容")
        ex_list.column("type", width=int(110 * s), minwidth=80, anchor=tk.CENTER)
        ex_list.column("value", width=int(430 * s), minwidth=220)

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

        # ====== 导航切换 ======
        def _highlight(active_cv):
            for cv in nav_items:
                _draw_nav(cv, cv is active_cv)

        def _on_idx_click(_e=None):
            page_exclude.pack_forget()
            page_index.pack(fill=tk.BOTH, expand=True)
            _highlight(idx_cv)

        def _on_ex_click(_e=None):
            page_index.pack_forget()
            page_exclude.pack(fill=tk.BOTH, expand=True)
            _highlight(ex_cv)

        nav_idx.bind("<Button-1>", _on_idx_click)
        nav_ex.bind("<Button-1>", _on_ex_click)
        for w in nav_idx.winfo_children():
            w.bind("<Button-1>", _on_idx_click)
        for w in nav_ex.winfo_children():
            w.bind("<Button-1>", _on_ex_click)
        body.bind("<Configure>", lambda _e: (_draw_nav(idx_cv, True), _draw_nav(ex_cv, False)))

        _on_idx_click()

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
        """统一执行搜索/筛选查询，并同步计数、分页和状态。"""
        filters = (filters if filters is not None else self._current_filters()).copy()
        self._last_query = query
        self._last_filters = filters
        if not IndexEngine.index_exists():
            self._results = []
            self._total_results = 0
            self._has_more = False
            self._refresh_tree()
            self._update_result_status()
            self._set_status("请先创建索引再搜索", "warning")
            return

        self._set_loading(True)
        try:
            self._total_results = IndexEngine.result_count(query, filters=filters)
            if query:
                self._results = IndexEngine.search(
                    query, limit=PAGE_SIZE, offset=0, order_col=self._sort_col,
                    order_desc=not self._sort_asc, filters=filters,
                )
            else:
                self._results = IndexEngine.load_all(
                    limit=PAGE_SIZE, offset=0, order_col=self._sort_col,
                    order_desc=not self._sort_asc, filters=filters,
                )
            self._has_more = len(self._results) < self._total_results
            self._refresh_tree()
            self._update_sort_heading()
            self._update_result_status()
        finally:
            self._set_loading(False)

    def _update_result_status(self):
        """统一更新顶部计数、右侧计数和底部状态文字。"""
        shown = len(self._results)
        total = self._total_results
        self.result_count_var.set(f"{total:,} 个结果")
        self.status_right_var.set(f"已显示 {shown:,} / {total:,}")
        if self._last_query:
            self._set_status(f"搜索「{self._last_query}」— 已显示 {shown:,} / {total:,} 个结果")
        else:
            self._set_status(f"已显示 {shown:,} / {total:,} 个结果")

    # ================================================================
    #  排除列表管理
    # ================================================================

    def _manage_exclude(self):
        """打开排除列表管理对话框。支持添加、编辑、删除排除项，操作即时保存。"""
        dialog = tk.Toplevel(self.root)
        dialog.title("排除列表管理")
        dialog.geometry("1000x700")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="索引时跳过匹配的目录（修改后需重建索引生效）:",
                  padding=(8, 8)).pack(anchor=tk.W)

        tb = ttk.Frame(dialog, padding=(8, 0, 8, 4))
        tb.pack(fill=tk.X)
        ttk.Button(tb, text="＋ 添加", command=lambda: self._exclude_add(ex_list, dialog)).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(tb, text="✎ 编辑", command=lambda: self._exclude_edit(ex_list, dialog)).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(tb, text="✕ 删除", command=lambda: self._exclude_delete(ex_list)).pack(side=tk.LEFT)

        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        columns = ("type", "value")
        ex_list = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        ex_list.heading("type", text="类型")
        ex_list.heading("value", text="排除内容")
        ex_list.column("type", width=80, minwidth=60)
        ex_list.column("value", width=860, minwidth=200)

        scroll_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=ex_list.yview)
        ex_list.configure(yscrollcommand=scroll_y.set)
        ex_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        ex_list.bind("<Double-1>", lambda e: self._exclude_edit(ex_list, dialog))

        data = IndexEngine.get_exclude_list()
        for d in data.get("dirs", []):
            ex_list.insert("", tk.END, values=("目录名", d))
        for p in data.get("paths", []):
            ex_list.insert("", tk.END, values=("路径包含", p))

        btn_frame = ttk.Frame(dialog, padding=(8, 0, 8, 8))
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="关闭", command=dialog.destroy).pack(side=tk.RIGHT)

    def _exclude_add(self, ex_list, parent):
        """添加排除项子对话框。"""
        sub = tk.Toplevel(parent)
        sub.title("添加排除项")
        sub.geometry("500x200")
        sub.transient(parent)
        sub.grab_set()
        sub.resizable(True, True)
        sub.minsize(400, 180)

        f = ttk.Frame(sub, padding=(12, 12))
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(f, text="类型:").grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        type_var = tk.StringVar(value="路径包含")
        cb = ttk.Combobox(f, textvariable=type_var, values=["目录名", "路径包含"], state="readonly", width=12)
        cb.grid(row=0, column=1, sticky=tk.W, pady=(0, 8))

        ttk.Label(f, text="内容:").grid(row=1, column=0, sticky=tk.W)
        val_entry = ttk.Entry(f, width=40)
        val_entry.grid(row=1, column=1, sticky=tk.EW, padx=(4, 0))
        val_entry.focus_set()

        f.columnconfigure(1, weight=1)

        def ok():
            v = val_entry.get().strip()
            if v:
                ex_list.insert("", tk.END, values=(type_var.get(), v))
                self._exclude_save(ex_list)
            sub.destroy()

        btn_f = ttk.Frame(sub, padding=(12, 0, 12, 8))
        btn_f.pack(fill=tk.X)
        ttk.Button(btn_f, text="取消", command=sub.destroy).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btn_f, text="确定", command=ok).pack(side=tk.RIGHT)
        val_entry.bind("<Return>", lambda e: ok())

    def _exclude_edit(self, ex_list, parent):
        """编辑选中排除项子对话框。"""
        sel = ex_list.selection()
        if not sel:
            return
        vals = ex_list.item(sel[0], "values")
        sub = tk.Toplevel(parent)
        sub.title("编辑排除项")
        sub.geometry("500x200")
        sub.transient(parent)
        sub.grab_set()
        sub.resizable(True, True)
        sub.minsize(400, 180)

        f = ttk.Frame(sub, padding=(12, 12))
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(f, text="类型:").grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        type_var = tk.StringVar(value=vals[0])
        cb = ttk.Combobox(f, textvariable=type_var, values=["目录名", "路径包含"], state="readonly", width=12)
        cb.grid(row=0, column=1, sticky=tk.W, pady=(0, 8))

        ttk.Label(f, text="内容:").grid(row=1, column=0, sticky=tk.W)
        val_entry = ttk.Entry(f, width=40)
        val_entry.insert(0, vals[1])
        val_entry.grid(row=1, column=1, sticky=tk.EW, padx=(4, 0))
        val_entry.focus_set()
        val_entry.selection_range(0, tk.END)

        f.columnconfigure(1, weight=1)

        def ok():
            v = val_entry.get().strip()
            if v:
                ex_list.item(sel[0], values=(type_var.get(), v))
                self._exclude_save(ex_list)
            sub.destroy()

        btn_f = ttk.Frame(sub, padding=(12, 0, 12, 8))
        btn_f.pack(fill=tk.X)
        ttk.Button(btn_f, text="取消", command=sub.destroy).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btn_f, text="确定", command=ok).pack(side=tk.RIGHT)
        val_entry.bind("<Return>", lambda e: ok())

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
        """清空并重新填充列表，空结果使用容器中央的真实空状态。"""
        self.tree.delete(*self.tree.get_children())
        self._item_to_result = {}
        if not self._results:
            self.empty_state.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            return
        self.empty_state.place_forget()
        self._append_to_tree(self._results)

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
        sel = self.tree.selection()
        paths = []
        for iid in sel:
            result = self._item_to_result.get(iid)
            if result:
                paths.append(result["path"])
        return paths

    def _on_double_click(self, event):
        """双击文件名：检查文件是否存在后打开。"""
        path = self._get_selected_path()
        if path:
            if not os.path.exists(path):
                messagebox.showwarning("文件不存在", "文件可能已被移动或删除：\n" + path)
                return
            open_with_default(path)

    def _on_right_click(self, event):
        """右键：不丢失多选，并根据选中数量置灰部分菜单项。"""
        row = self.tree.identify_row(event.y)
        current_sel = self.tree.selection()
        if row not in self._item_to_result:
            return
        # 如果右键点击的行不在当前选中列表中，则只选这一行
        if row and row not in current_sel:
            self.tree.selection_set(row)
            current_sel = (row,)

        # 根据选中数量设置菜单项状态
        multi = len(current_sel) > 1
        self._ctx_menu.entryconfig("打开", state="disabled" if multi else "normal")
        self._ctx_menu.entryconfig("打开所在文件夹", state="disabled" if multi else "normal")
        self._ctx_menu.entryconfig("重命名", state="disabled" if multi else "normal")
        # 删除类始终可用
        self._ctx_menu.entryconfig("删除到回收站", state="normal")
        self._ctx_menu.entryconfig("彻底删除", state="normal")

        if current_sel:
            self._ctx_menu.post(event.x_root, event.y_root)

    def _open_selected(self):
        """右键菜单 → 打开文件。"""
        path = self._get_selected_path()
        if path:
            if not os.path.exists(path):
                messagebox.showwarning("文件不存在", "文件可能已被移动或删除：\n" + path)
                return
            open_with_default(path)

    def _open_file_location_selected(self):
        """右键菜单 → 在资源管理器中定位文件。"""
        path = self._get_selected_path()
        if path:
            if not os.path.exists(path):
                parent = os.path.dirname(path)
                if not os.path.exists(parent):
                    messagebox.showwarning("路径不存在", "路径可能已被移动或删除：\n" + path)
                    return
            open_file_location(path)

    # ================================================================
    #  重命名
    # ================================================================

    def _rename_file_dialog(self):
        """弹出重命名对话框，执行文件重命名并刷新列表。"""
        path = self._get_selected_path()
        if not path:
            return
        old_name = os.path.basename(path)
        dlg = tk.Toplevel(self.root)
        dlg.title("重命名")
        dlg.geometry("600x240")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(True, True)
        dlg.minsize(500, 220)
        frm = ttk.Frame(dlg, padding=(24, 20))
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="新文件名:").grid(row=0, column=0, sticky=tk.W, pady=(0, 12))
        name_var = tk.StringVar(value=old_name)
        entry = ttk.Entry(frm, textvariable=name_var, width=50)
        entry.grid(row=1, column=0, sticky=tk.EW, pady=(0, 12))
        entry.selection_range(0, tk.END)
        entry.focus_set()
        frm.columnconfigure(0, weight=1)
        result = {"name": None}
        def ok():
            result["name"] = name_var.get().strip()
            dlg.destroy()
        bf = ttk.Frame(frm)
        bf.grid(row=2, column=0, sticky=tk.E)
        ttk.Button(bf, text="取消", command=dlg.destroy).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(bf, text="确定", command=ok).pack(side=tk.RIGHT)
        entry.bind("<Return>", lambda e: ok())
        self.root.wait_window(dlg)
        new_name = result["name"]
        if not new_name or new_name == old_name:
            return
        try:
            new_path = rename_file(path, new_name)
            sel = self.tree.selection()
            if sel:
                r = self._item_to_result.get(sel[0])
                if r:
                    r["name"] = new_name
                    r["path"] = new_path
            self._refresh_tree()
            self.status_var.set(f"已重命名: {old_name} → {new_name}")
        except Exception as e:
            messagebox.showerror("重命名失败", str(e))

    def _open_new_window(self):
        """右键菜单 → 新开一个程序窗口。"""
        try:
            script = os.path.abspath(__file__)
            python = sys.executable
            flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            subprocess.Popen([python, script], creationflags=flags)
            self.status_var.set("已打开新窗口")
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

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
            msg = f"确定将「{os.path.basename(paths[0])}」移动到回收站？"
        else:
            msg = f"确定将选中的 {count} 个文件移动到回收站？"
        if not messagebox.askyesno("确认删除", msg):
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
            messagebox.showerror("删除失败", str(e))

    def _delete_file_permanent(self):
        """彻底删除文件（不可恢复，有二次确认），支持多选。"""
        paths = self._get_selected_paths()
        if not paths:
            return
        count = len(paths)
        if count == 1:
            msg = f"确定彻底删除「{os.path.basename(paths[0])}」？\n\n此操作不可恢复！"
        else:
            msg = f"确定彻底删除选中的 {count} 个文件？\n\n此操作不可恢复！"
        if not messagebox.askyesno("确认彻底删除", msg):
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
            messagebox.showerror("删除失败", str(e))

    def _remove_from_results(self, path: str):
        """从当前结果列表中移除指定路径并同步当前计数。"""
        before = len(self._results)
        self._results = [f for f in self._results if f["path"] != path]
        if len(self._results) < before:
            self._total_results = max(0, self._total_results - 1)
            self._has_more = len(self._results) < self._total_results

    # ================================================================
    #  拖拽到外部程序
    # ================================================================

    def _setup_drag_drop(self):
        """设置文件拖拽功能（依赖 tkdnd 库）。"""
        try:
            self.root.tk.call("package", "require", "tkdnd")
            self.root.tk.eval(f"tkdnd::drag_source register {self.tree._w} DND_Files")
            self._dnd_cb = self.root.register(self._on_dnd_data)
            self.root.tk.eval(f"tkdnd::drag_source handler {self.tree._w} drag {self._dnd_cb}")
            self._dnd_ok = True
        except tk.TclError:
            self._dnd_ok = False
            self.tree.bind("<B1-Motion>", self._on_drag_fallback)

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
    #  无限滚动加载
    # ================================================================

    def _on_tree_scroll_wrapper(self, *args):
        """拦截滚动条事件，同时调用原始滚动和加载更多检测。"""
        self.tree.yview(*args)
        self._on_tree_scroll()

    def _load_more(self):
        """按当前关键词、筛选和排序加载下一页。"""
        if not self._has_more or self._loading_more:
            return
        self._loading_more = True
        try:
            offset = len(self._results)
            query = self._last_query
            filters = self._last_filters
            if query:
                more = IndexEngine.search(
                    query, limit=PAGE_SIZE, offset=offset,
                    order_col=self._sort_col, order_desc=not self._sort_asc, filters=filters,
                )
            else:
                more = IndexEngine.load_all(
                    limit=PAGE_SIZE, offset=offset,
                    order_col=self._sort_col, order_desc=not self._sort_asc, filters=filters,
                )
            if more:
                self._results.extend(more)
                self._append_to_tree(more)
            self._has_more = len(self._results) < self._total_results
            self._update_result_status()
        finally:
            self._loading_more = False

    def _append_to_tree(self, items):
        """追加结果，使用 Shell 图标和基于全局行号的交替行色。"""
        self.empty_state.place_forget()
        start = len(self._item_to_result)
        for index, result in enumerate(items, start=start):
            icon = self._icon_cache.get(result["path"], bool(result.get("is_dir")))
            tag = "even" if index % 2 else "odd"
            iid = self.tree.insert(
                "", tk.END, text="", image=icon,
                values=self._tree_item_values(result), tags=(tag,),
            )
            self._item_to_result[iid] = result

    def _on_tree_scroll(self, *args):
        """检测滚动条是否到达底部（>= 95%），触发加载更多。"""
        if not self._has_more or self._loading_more:
            return
        try:
            _, bottom = self.tree.yview()
            if bottom >= 0.95:
                self.root.after(100, self._load_more)
        except Exception:
            pass

    # ================================================================
    #  排序
    # ================================================================

    def _update_sort_heading(self):
        """刷新自绘表头的当前排序箭头。"""
        self._draw_tree_header()

    def _sort_by(self, col: str):
        """点击列头排序，并按当前关键词和筛选重新查询第一页。"""
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._run_query(self._search_text())

    def _bind_shortcuts(self):
        """绑定全局搜索快捷键和仅在结果列表生效的文件操作快捷键。"""
        self.root.bind("<Control-l>", self._focus_search, add="+")
        self.root.bind("<Control-L>", self._focus_search, add="+")
        self.root.bind("<Control-f>", self._focus_search, add="+")
        self.root.bind("<Control-F>", self._focus_search, add="+")
        self.root.bind("<Escape>", self._clear_search, add="+")
        self.tree.bind("<Return>", lambda _e: (self._open_selected(), "break")[1])
        self.tree.bind("<Control-a>", self._select_all_results)
        self.tree.bind("<Control-A>", self._select_all_results)
        self.tree.bind("<Control-c>", lambda _e: (self._copy_path(), "break")[1])
        self.tree.bind("<Control-C>", lambda _e: (self._copy_path(), "break")[1])
        self.tree.bind("<Control-x>", lambda _e: (self._cut_path(), "break")[1])
        self.tree.bind("<Control-X>", lambda _e: (self._cut_path(), "break")[1])
        self.tree.bind("<Delete>", lambda _e: (self._delete_file_recycle(), "break")[1])
        self.tree.bind("<Shift-Delete>", lambda _e: (self._delete_file_permanent(), "break")[1])
        self.tree.bind("<Alt-Return>", lambda _e: (self._open_file_location_selected(), "break")[1])

    def _focus_search(self, _event=None):
        """聚焦搜索框并全选真实搜索文字。"""
        self.search_entry.focus_set()
        self._hide_placeholder()
        self.search_entry.selection_range(0, tk.END)
        self.search_entry.icursor(tk.END)
        return "break"

    def _select_all_results(self, _event=None):
        """仅在结果列表中全选有效结果行。"""
        items = tuple(self._item_to_result)
        if items:
            self.tree.selection_set(items)
        return "break"

    # ================================================================
    #  列宽布局持久化
    # ================================================================

    def _load_layout(self):
        """从 JSON 文件恢复上次的列宽。"""
        if not IndexEngine.LAYOUT_FILE.exists():
            return
        try:
            data = json.loads(IndexEngine.LAYOUT_FILE.read_text(encoding="utf-8"))
            for col in ("name", "path", "type", "size", "modified"):
                if col in data:
                    self.tree.column(col, width=data[col])
        except Exception:
            pass

    def _save_layout(self):
        """将当前列宽保存到 JSON 文件。"""
        data = {}
        for col in ("name", "path", "type", "size", "modified"):
            data[col] = self.tree.column(col, "width")
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
