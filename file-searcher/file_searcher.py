try:
    import ctypes
    ctypes.windll.shcore.SetProcessDPIAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

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

import hashlib

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime

try:
    from pypinyin import lazy_pinyin
    _PYIN_OK = True
except ImportError:
    _PYIN_OK = False


# ================================================================
#  全局常量配置
# ================================================================
INDEX_DIR = Path.home() / ".file_searcher_index"
INDEX_DB = INDEX_DIR / "index.db"

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
        os.startfile(os.path.normpath(path))
    except OSError as e:
        messagebox.showerror("打开失败", str(e))


def open_file_location(path: str):
    """在资源管理器中定位并选中文件"""
    subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])


def send_to_recycle_bin(path: str):
    """将文件移入回收站（可通过 Windows Shell API 恢复）"""
    buf = ctypes.create_unicode_buffer(path + "\0\0")
    op = SHFILEOPSTRUCTW()
    op.wFunc = FO_DELETE
    op.pFrom = ctypes.cast(buf, ctypes.c_wchar_p)
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
    ret = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    if ret != 0:
        raise OSError(f"无法删除文件: {path}")


def permanent_delete(path: str):
    """彻底删除文件，不可恢复"""
    buf = ctypes.create_unicode_buffer(path + "\0\0")
    op = SHFILEOPSTRUCTW()
    op.wFunc = FO_DELETE
    op.pFrom = ctypes.cast(buf, ctypes.c_wchar_p)
    op.fFlags = FOF_NOCONFIRMATION | FOF_SILENT | FOF_WANTNUKEWARNING
    ret = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    if ret != 0:
        raise OSError(f"无法删除文件: {path}")


def rename_file(old_path: str, new_name: str) -> str:
    """重命名文件，如果目标已存在则抛出异常。返回新路径"""
    parent = os.path.dirname(old_path)
    new_path = os.path.join(parent, new_name)
    if os.path.exists(new_path):
        raise FileExistsError(f"目标已存在: {new_path}")
    os.rename(old_path, new_path)
    return new_path


def to_pinyin(text: str) -> str:
    """将中文文本转为拼音串（空格分隔），用于拼音搜索。"""
    if not _PYIN_OK:
        return ""
    try:
        return " ".join(lazy_pinyin(text))
    except Exception:
        return ""


# ================================================================
#  IndexEngine — 索引引擎
# ================================================================
class IndexEngine:
    """全盘文件索引引擎，使用 SQLite 存储文件元数据。"""

    EXCLUDE_FILE = INDEX_DIR / "exclude.json"
    LAYOUT_FILE = INDEX_DIR / "layout.json"

    def __init__(self, progress_callback=None, cancel_check=None):
        self._progress = progress_callback or (lambda msg, n: None)
        self._cancel = cancel_check or (lambda: False)
        self._exclude_dirs, self._exclude_paths = self._load_exclude_list()

    # ---- 索引构建 ----

    def build_index(self):
        """构建全盘文件索引。遍历所有盘符，递归扫描目录，批量写入 SQLite。"""
        INDEX_DIR.mkdir(parents=True, exist_ok=True)

        conn = sqlite3.connect(str(INDEX_DB))
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=OFF")
        conn.execute("DROP TABLE IF EXISTS files")
        conn.execute("DROP TABLE IF EXISTS files_fts")
        conn.execute("""
            CREATE TABLE files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                name_lower TEXT NOT NULL,
                path TEXT NOT NULL UNIQUE,
                path_lower TEXT NOT NULL,
                size INTEGER NOT NULL,
                modified TEXT NOT NULL,
                hash TEXT
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_name_lower ON files(name_lower)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_path_lower ON files(path_lower)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_size ON files(size)")
        conn.execute("""
            CREATE VIRTUAL TABLE files_fts USING fts5(
                name, path, name_pinyin,
                content='', tokenize='unicode61'
            )
        """)
        conn.commit()

        drives = [f"{d}:" for d in "CDEFGHIJKLMNOPQRSTUVWXYZAB" if os.path.exists(f"{d}:")]
        total_files = 0
        insert_batch = []
        dir_stack = [d + "\\" for d in drives]

        self._progress("开始扫描磁盘…", 0)

        while dir_stack:
            if self._cancel():
                break
            d = dir_stack.pop()
            total_files = self._scan_directory(d, conn, insert_batch, total_files, dir_stack)

        if insert_batch:
            conn.executemany(
                "INSERT OR IGNORE INTO files(name, name_lower, path, path_lower, size, modified) "
                "VALUES(?,?,?,?,?,?)",
                insert_batch,
            )
            conn.commit()
            self._progress(f"已收录 {total_files:,} 个文件", total_files)

        conn.close()
        return total_files

    def _scan_directory(self, dirpath, conn, insert_batch, total_files, dir_stack):
        """扫描单个目录：将文件信息加入批处理队列，子目录加入栈。"""
        try:
            entries = os.scandir(dirpath)
        except (PermissionError, OSError):
            return total_files

        with entries:
            for entry in entries:
                if self._cancel():
                    break
                full = entry.path
                try:
                    is_dir = entry.is_dir(follow_symlinks=False)
                except OSError:
                    continue
                if is_dir:
                    if self._should_skip_dir(full, entry):
                        continue
                    dir_stack.append(full)
                else:
                    try:
                        is_file = entry.is_file(follow_symlinks=False)
                    except OSError:
                        continue
                    if not is_file:
                        continue
                    if not self._should_include_file(full, entry):
                        continue
                    try:
                        st = entry.stat()
                    except OSError:
                        continue
                    insert_batch.append((
                        entry.name,
                        entry.name.lower(),
                        full,
                        full.lower(),
                        st.st_size,
                        datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    ))
                    total_files += 1
                    if len(insert_batch) >= 500:
                        conn.executemany(
                            "INSERT OR IGNORE INTO files(name, name_lower, path, path_lower, size, modified) "
                            "VALUES(?,?,?,?,?,?)",
                            insert_batch,
                        )
                        conn.commit()
                        insert_batch.clear()
                    # 每处理 500 个文件发送一次进度更新
                    if total_files % 500 == 0:
                        self._progress(f"已收录 {total_files:,} 个文件", total_files)
        return total_files

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
    def search(query: str, limit: int = 5000, offset: int = 0, order_col: str = "name", order_desc: bool = False) -> list[dict]:
        """在索引中搜索文件。类型列排序在 Python 侧完成以与显示一致。"""
        if not INDEX_DB.exists():
            return []
        q = query.strip().lower()
        conn = sqlite3.connect(str(INDEX_DB))
        conn.row_factory = sqlite3.Row
        if order_col == "type":
            sql = "SELECT name, path, size, modified FROM files WHERE name_lower LIKE ? OR path_lower LIKE ?"
            rows = conn.execute(sql, ("%" + q + "%", "%" + q + "%")).fetchall()
            conn.close()
            results = [dict(r) for r in rows]
            results.sort(key=lambda r: os.path.splitext(r["name"])[1].lower(), reverse=order_desc)
            return results[offset:offset + limit]
        col = IndexEngine._COL_DB.get(order_col, "name_lower")
        direction = "DESC" if order_desc else "ASC"
        sql = f"SELECT name, path, size, modified FROM files WHERE name_lower LIKE ? OR path_lower LIKE ? ORDER BY {col} {direction} LIMIT ? OFFSET ?"
        rows = conn.execute(sql, ("%" + q + "%", "%" + q + "%", limit, offset)).fetchall()
        conn.close()
        return [dict(r) for r in rows]

    @staticmethod
    def load_all(limit: int = 5000, offset: int = 0, order_col: str = "size", order_desc: bool = True) -> list[dict]:
        """加载全部文件，默认按大小降序。类型列排序在 Python 侧完成。"""
        if not INDEX_DB.exists():
            return []
        conn = sqlite3.connect(str(INDEX_DB))
        conn.row_factory = sqlite3.Row
        if order_col == "type":
            sql = "SELECT name, path, size, modified FROM files"
            rows = conn.execute(sql).fetchall()
            conn.close()
            results = [dict(r) for r in rows]
            results.sort(key=lambda r: os.path.splitext(r["name"])[1].lower(), reverse=order_desc)
            return results[offset:offset + limit]
        col = IndexEngine._COL_DB.get(order_col, "size")
        direction = "DESC" if order_desc else "ASC"
        sql = f"SELECT name, path, size, modified FROM files ORDER BY {col} {direction} LIMIT ? OFFSET ?"
        rows = conn.execute(sql, (limit, offset)).fetchall()
        conn.close()
        return [dict(r) for r in rows]

    # ---- FTS5 模糊搜索 ----

    @staticmethod
    def search_fts(query: str, limit: int = 5000, offset: int = 0) -> list[dict]:
        """使用 FTS5 进行模糊搜索，支持中文分词和拼音搜索。"""
        if not INDEX_DB.exists():
            return []
        q = query.strip()
        if not q:
            return []

        conn = sqlite3.connect(str(INDEX_DB))
        conn.row_factory = sqlite3.Row

        # 检查 FTS 表是否存在
        fts_exists = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='files_fts'"
        ).fetchone()
        if not fts_exists:
            conn.close()
            return IndexEngine.search(query, limit, offset)

        # FTS 表为空时回退到 LIKE
        fts_count = conn.execute("SELECT COUNT(*) FROM files_fts").fetchone()[0]
        if fts_count == 0:
            conn.close()
            return IndexEngine.search(query, limit, offset)

        # 构建 FTS 查询：在 name、path、name_pinyin 三列中搜索
        # 将用户输入转为拼音也加入搜索词，支持用拼音搜中文文件名
        pinyin_q = to_pinyin(q)
        search_terms = q
        if pinyin_q:
            search_terms = f"{q} {pinyin_q}"

        # FTS5 查询语法：各列 OR 匹配
        sql = """
            SELECT f.name, f.path, f.size, f.modified,
                   rank
            FROM files_fts
            JOIN files f ON f.path = files_fts.path
            WHERE files_fts MATCH ?
            ORDER BY rank
            LIMIT ? OFFSET ?
        """
        try:
            rows = conn.execute(sql, (search_terms, limit, offset)).fetchall()
        except sqlite3.OperationalError:
            conn.close()
            return IndexEngine.search(query, limit, offset)

        conn.close()
        return [dict(r) for r in rows]

    # ---- FTS 后台填充 ----

    @staticmethod
    def populate_fts(batch_size: int = 2000, progress_cb=None) -> int:
        """后台填充 FTS5 表（含拼音），从已索引的 files 表读取。返回处理条数。"""
        if not INDEX_DB.exists():
            return 0
        conn = sqlite3.connect(str(INDEX_DB))
        conn.execute("PRAGMA journal_mode=WAL")

        # 确保 FTS 表存在
        fts_exists = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='files_fts'"
        ).fetchone()
        if not fts_exists:
            conn.execute("""
                CREATE VIRTUAL TABLE files_fts USING fts5(
                    name, path, name_pinyin,
                    content='', tokenize='unicode61'
                )
            """)
            conn.commit()

        # 统计总数
        total_row = conn.execute("SELECT COUNT(*) FROM files").fetchone()
        total = total_row[0] if total_row else 0
        if total == 0:
            conn.close()
            return 0

        offset = 0
        processed = 0
        while offset < total:
            rows = conn.execute(
                "SELECT name, path FROM files LIMIT ? OFFSET ?",
                (batch_size, offset),
            ).fetchall()

            fts_data = [(name, path, to_pinyin(name)) for name, path in rows]
            conn.executemany(
                "INSERT INTO files_fts(name, path, name_pinyin) VALUES(?,?,?)",
                fts_data,
            )
            conn.commit()

            processed += len(rows)
            offset += len(rows)
            if progress_cb:
                progress_cb(processed, total)

            if not rows:
                break

        conn.close()
        return processed

    # ---- 重复文件检测 ----

    @staticmethod
    def _hash_column_exists() -> bool:
        """检查 files 表是否有 hash 列（兼容旧索引）。"""
        if not INDEX_DB.exists():
            return False
        conn = sqlite3.connect(str(INDEX_DB))
        cols = [r[1] for r in conn.execute("PRAGMA table_info(files)").fetchall()]
        conn.close()
        return "hash" in cols

    @staticmethod
    def find_duplicate_candidates(min_size: int = 1024) -> list[tuple[int, int]]:
        """找出大小相同的文件组（重复候选），返回 [(size, count), ...] 按 count 降序。"""
        if not INDEX_DB.exists():
            return []
        if not IndexEngine._hash_column_exists():
            return []  # 旧索引，需要重建
        conn = sqlite3.connect(str(INDEX_DB))
        rows = conn.execute(
            "SELECT size, COUNT(*) as cnt FROM files "
            "WHERE size >= ? "
            "GROUP BY size HAVING cnt > 1 ORDER BY cnt DESC",
            (min_size,),
        ).fetchall()
        conn.close()
        return [(r[0], r[1]) for r in rows]

    @staticmethod
    def compute_hashes_for_size(size: int, progress_cb=None) -> list[dict]:
        """对指定大小的所有文件计算 MD5 哈希，返回含 hash 的文件列表。"""
        if not INDEX_DB.exists():
            return []
        conn = sqlite3.connect(str(INDEX_DB))
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT id, name, path, size, modified FROM files WHERE size = ? AND hash IS NULL",
            (size,),
        ).fetchall()

        results = []
        total = len(rows)
        for i, r in enumerate(rows):
            if progress_cb and i % 10 == 0:
                progress_cb(i, total)
            try:
                md5 = IndexEngine._file_md5(r["path"])
            except (OSError, PermissionError):
                continue
            conn.execute("UPDATE files SET hash = ? WHERE id = ?", (md5, r["id"]))
            results.append({**dict(r), "hash": md5})

        conn.commit()
        conn.close()
        if progress_cb:
            progress_cb(total, total)
        return results

    @staticmethod
    def get_duplicates() -> list[dict]:
        """获取所有重复文件组，返回 [{group_id, hash, files: [{name, path, size, modified}]}]。"""
        if not INDEX_DB.exists():
            return []
        conn = sqlite3.connect(str(INDEX_DB))
        conn.row_factory = sqlite3.Row
        dup_hashes = conn.execute(
            "SELECT hash, COUNT(*) as cnt FROM files "
            "WHERE hash IS NOT NULL "
            "GROUP BY hash HAVING cnt > 1 ORDER BY cnt DESC"
        ).fetchall()

        groups = []
        for gid, (h, cnt) in enumerate(dup_hashes, 1):
            files = conn.execute(
                "SELECT name, path, size, modified FROM files WHERE hash = ? ORDER BY path",
                (h,),
            ).fetchall()
            groups.append({
                "group_id": gid,
                "hash": h,
                "count": cnt,
                "files": [dict(f) for f in files],
            })

        conn.close()
        return groups

    @staticmethod
    def _file_md5(filepath: str, chunk_size: int = 65536) -> str:
        """计算文件的 MD5 哈希值。"""
        md5 = hashlib.md5()
        with open(filepath, "rb") as f:
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                md5.update(chunk)
        return md5.hexdigest()

    @staticmethod
    def clear_hashes():
        """清除所有文件的哈希值（用于重新扫描重复文件）。"""
        if not INDEX_DB.exists():
            return
        conn = sqlite3.connect(str(INDEX_DB))
        conn.execute("UPDATE files SET hash = NULL")
        conn.commit()
        conn.close()

    # ---- 排除列表管理 ----
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


# ================================================================
#  FileSearcherApp — GUI 主应用
# ================================================================
class FileSearcherApp:
    """全盘文件搜索 GUI 应用。使用 ttk.Treeview 展示文件列表。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("File Searcher — 全盘文件搜索")
        self.root.state("zoomed")
        self.root.minsize(800, 400)

        self._engine_cancel = False
        self._search_timer = None
        self._results: list[dict] = []
        self._sort_col = "size"
        self._sort_asc = False
        self._has_more = False
        self._last_query = ""
        self._loading_more = False
        self._search_mode = "exact"   # "exact" | "fuzzy"
        self._view_mode = "normal"    # "normal" | "duplicates"
        self._dup_groups: list[dict] = []  # 重复文件组缓存

        self._build_toolbar()
        self._build_tree()
        self._load_layout()
        self._setup_drag_drop()
        self._build_context_menu()
        self._build_statusbar()
        self._update_index_button_text()

        self.root.after(100, self._load_all)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)


    # ================================================================
    #  UI 构建
    # ================================================================

    def _build_toolbar(self):
        """构建顶部工具栏：搜索框 + 模式切换 + 索引按钮 + 重复检测。"""
        toolbar = ttk.Frame(self.root, padding=(8, 6))
        toolbar.pack(fill=tk.X)

        ttk.Label(toolbar, text="\U0001f50d 搜索:").pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._on_search_changed())
        search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=40)
        search_entry.pack(side=tk.LEFT, padx=(0, 4))
        search_entry.bind("<Return>", lambda _e: self._do_search())

        # 搜索模式切换
        self.search_mode_var = tk.StringVar(value="精确")
        mode_cb = ttk.Combobox(toolbar, textvariable=self.search_mode_var,
                               values=["精确", "模糊"], state="readonly", width=5)
        mode_cb.pack(side=tk.LEFT, padx=(0, 12))
        mode_cb.bind("<<ComboboxSelected>>", lambda e: self._on_search_mode_changed())

        self.root.bind("<Escape>", lambda _e: self._clear_search())
        self._index_icon = tk.PhotoImage(width=1, height=1)

        self.index_btn = ttk.Button(toolbar, text="创建索引", command=self._toggle_index)
        self.index_btn.pack(side=tk.LEFT, padx=(4, 0))
        self.index_count_var = tk.StringVar(value="")
        ttk.Label(toolbar, textvariable=self.index_count_var, foreground="gray").pack(side=tk.LEFT, padx=(8, 0))

        ttk.Button(toolbar, text="排除列表", command=self._manage_exclude).pack(side=tk.LEFT, padx=(4, 0))
        self.dup_btn = ttk.Button(toolbar, text="查找重复文件", command=self._start_duplicate_scan)
        self.dup_btn.pack(side=tk.LEFT, padx=(8, 0))
        self.dup_clear_btn = ttk.Button(toolbar, text="退出重复视图", command=self._exit_duplicate_view, state=tk.DISABLED)
        self.dup_clear_btn.pack(side=tk.LEFT, padx=(4, 0))

    def _build_tree(self):
        """构建中央文件列表 Treeview。列顺序：文件名 | 路径 | 类型 | 大小 | 修改时间。"""
        frame = ttk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

        columns = ("name", "path", "type", "size", "modified")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("name", text="文件名", command=lambda: self._sort_by("name"))
        self.tree.heading("path", text="路径", command=lambda: self._sort_by("path"))
        self.tree.heading("type", text="类型", command=lambda: self._sort_by("type"))
        self.tree.heading("size", text="大小", command=lambda: self._sort_by("size"))
        self.tree.heading("modified", text="修改时间", command=lambda: self._sort_by("modified"))

        self.tree.column("name", width=220, minwidth=120)
        self.tree.column("path", width=460, minwidth=160)
        self.tree.column("type", width=80, minwidth=60, anchor=tk.CENTER)
        self.tree.column("size", width=80, minwidth=60, anchor=tk.E)
        self.tree.column("modified", width=75, minwidth=65, anchor=tk.CENTER)

        scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.configure(command=self._on_tree_scroll_wrapper)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        style = ttk.Style()
        style.map("Treeview",
            background=[("selected", "#0078D4")],
            foreground=[("selected", "white")],
        )
        import tkinter.font as tkfont
        FONT = ("Microsoft YaHei", 9)
        _font_obj = tkfont.Font(family="Microsoft YaHei", size=9)
        _row_h = _font_obj.metrics("linespace") + 6
        FONT_HEAD = ("Microsoft YaHei", 9, "bold")
        style.configure("Treeview", font=FONT, rowheight=_row_h)
        style.configure("Treeview.Heading", background="#D0D0D0", relief="flat", font=FONT_HEAD)
        style.map("Treeview.Heading",
            background=[("active", "#A0A0A0")])

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Delete>", lambda e: self._delete_file_recycle())
        self.tree.bind("<Control-c>", self._copy_path)
        self.tree.bind("<Control-x>", self._cut_path)

    def _build_context_menu(self):
        """构建右键弹出菜单。"""
        self._ctx_menu = tk.Menu(self.root, tearoff=0)
        self._ctx_menu.add_command(label="打开", command=self._open_selected)
        self._ctx_menu.add_command(label="打开所在文件夹", command=self._open_file_location_selected)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="重命名", command=self._rename_file_dialog)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="删除到回收站", command=self._delete_file_recycle)
        self._ctx_menu.add_command(label="彻底删除", command=self._delete_file_permanent)

    def _build_statusbar(self):
        """构建底部状态栏。"""
        self.status_var = tk.StringVar(value="就绪 — 请先创建索引再搜索文件")
        statusbar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(8, 2))
        statusbar.pack(fill=tk.X, side=tk.BOTTOM)

    # ================================================================
    #  索引管理
    # ================================================================

    def _update_index_button_text(self):
        """根据索引状态更新按钮文字。"""
        if IndexEngine.index_exists():
            count = IndexEngine.index_file_count()
            self.index_btn.config(text="重建索引")
            self.index_count_var.set(f"已索引 {count:,} 个文件")
        else:
            self.index_btn.config(text="创建索引")
            self.index_count_var.set("尚未创建索引")

    def _toggle_index(self):
        """点击索引按钮：创建或重建索引。"""
        if self._engine_cancel:
            return
        if IndexEngine.index_exists():
            if not messagebox.askyesno("确认", "重建索引将扫描所有磁盘，可能需要几分钟。继续？"):
                return
        self._do_index()

    def _do_index(self):
        """在后台线程中执行索引构建。"""
        self._engine_cancel = False
        self.index_btn.config(state=tk.DISABLED)
        self.status_var.set("正在创建索引，扫描全盘文件…")

        engine = IndexEngine(
            progress_callback=lambda msg, n: self.root.after(0, self._on_index_progress, msg, n),
            cancel_check=lambda: self._engine_cancel,
        )

        def run():
            try:
                total = engine.build_index()
                self.root.after(0, lambda: self._on_index_done(total))
            except Exception as e:
                self.root.after(0, lambda: self._on_index_error(str(e)))

        threading.Thread(target=run, daemon=True).start()

    def _stop_index(self):
        """停止正在进行的索引构建。"""
        if not self._engine_cancel:
            self._engine_cancel = True
            self.status_var.set("正在停止索引…")

    def _on_index_progress(self, msg: str, count: int):
        """索引进度回调。"""
        self.status_var.set(msg)
        self.index_count_var.set(f"已收录 {count:,} 个文件")

    def _on_index_done(self, total: int):
        """索引完成回调。启动 FTS 模糊搜索索引后台填充。"""
        self._engine_cancel = False
        self.index_btn.config(state=tk.NORMAL)
        self._update_index_button_text()
        self.status_var.set(f"索引完成 — 共收录 {total:,} 个文件，正在准备模糊搜索索引…")
        self._load_all()

        # 后台填充 FTS（拼音索引），不阻塞正常搜索
        def run_fts():
            IndexEngine.populate_fts(
                batch_size=2000,
                progress_cb=lambda done, tot: self.root.after(
                    0, lambda: self.status_var.set(
                        f"模糊搜索索引准备中… {done:,} / {tot:,}"
                    )
                ),
            )
            self.root.after(0, lambda: self._update_index_button_text())

        threading.Thread(target=run_fts, daemon=True).start()

    def _on_index_error(self, err: str):
        """索引出错回调。"""
        self._engine_cancel = False
        self.index_btn.config(state=tk.NORMAL)
        self.status_var.set(f"索引出错: {err}")

    # ================================================================
    #  搜索逻辑
    # ================================================================

    def _load_all(self):
        """加载首页（默认按大小降序的前 5000 个文件）。"""
        if not IndexEngine.index_exists():
            return
        self._view_mode = "normal"
        self._dup_groups = []
        self.dup_clear_btn.config(state=tk.DISABLED)
        self._last_query = ""
        self._sort_col = "size"
        self._sort_asc = False
        self._results = IndexEngine.load_all(limit=5000, offset=0, order_col=self._sort_col, order_desc=not self._sort_asc)
        self._has_more = len(self._results) == 5000
        self._refresh_tree()
        self._update_sort_heading()
        total = IndexEngine.index_file_count()
        self.status_var.set(f"就绪 — 已索引 {total:,} 个文件，已显示 {len(self._results):,} 个（按大小降序）")

    def _clear_search(self):
        """清空搜索框，恢复显示全部文件。"""
        self.search_var.set("")
        self._load_all()

    def _on_search_changed(self):
        """搜索框内容变化时触发（300ms 防抖延迟）。"""
        if self._search_timer is not None:
            self.root.after_cancel(self._search_timer)
        self._search_timer = self.root.after(300, self._do_search)

    def _on_search_mode_changed(self):
        """搜索模式切换时立即重新搜索。"""
        mode = self.search_mode_var.get()
        self._search_mode = "fuzzy" if mode == "模糊" else "exact"
        self._do_search()

    def _do_search(self):
        """执行搜索：精确模式用 LIKE，模糊模式用 FTS5。"""
        query = self.search_var.get().strip()
        if not query:
            self._load_all()
            return
        if not IndexEngine.index_exists():
            self.status_var.set("请先创建索引再搜索")
            return
        self._last_query = query
        self._view_mode = "normal"

        if self._search_mode == "fuzzy":
            self._results = IndexEngine.search_fts(query, limit=5000, offset=0)
            self._has_more = len(self._results) == 5000
            mode_label = "模糊"
        else:
            self._results = IndexEngine.search(query, limit=5000, offset=0,
                                               order_col=self._sort_col, order_desc=not self._sort_asc)
            self._has_more = len(self._results) == 5000
            mode_label = "精确"

        self._refresh_tree()
        self.status_var.set(f"搜索「{query}」({mode_label}) — 已显示 {len(self._results):,} 个文件")

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
    #  重复文件检测 GUI
    # ================================================================

    def _start_duplicate_scan(self):
        """启动重复文件扫描（后台线程）。"""
        if not IndexEngine.index_exists():
            messagebox.showwarning("提示", "请先创建索引再查找重复文件")
            return

        if not IndexEngine._hash_column_exists():
            if messagebox.askyesno("需要重建索引",
                                   "当前索引缺少哈希列（旧版本索引），需要重建索引才能查找重复文件。\n\n是否立即重建？"):
                self._do_index()
            return

        # 确认是否重新扫描
        existing = IndexEngine.get_duplicates()
        if existing:
            total_dup = sum(g["count"] for g in existing)
            if messagebox.askyesno("已有扫描结果",
                                   f"已有重复文件扫描结果：{len(existing)} 组，共 {total_dup} 个文件。\n\n"
                                   "是否重新扫描？\n（选择「否」将直接查看已有结果）"):
                IndexEngine.clear_hashes()
            else:
                self._show_duplicates(existing)
                return

        # 获取候选：大小相同的文件组
        candidates = IndexEngine.find_duplicate_candidates(min_size=1024)
        if not candidates:
            messagebox.showinfo("结果", "没有找到大小相同的文件，无需哈希比对。")
            return

        total_sizes = len(candidates)
        total_files_to_hash = sum(c[1] for c in candidates)
        msg = (f"发现 {total_sizes} 组大小相同的文件\n"
               f"共需比对 {total_files_to_hash:,} 个文件\n"
               f"（首次扫描需要计算哈希值，可能较慢，请耐心等待）\n\n继续扫描？")
        if not messagebox.askyesno("确认扫描", msg):
            return

        self.dup_btn.config(state=tk.DISABLED)
        self.status_var.set("正在扫描重复文件，计算哈希值…")

        def run():
            try:
                done_groups = 0
                for size, cnt in candidates:
                    if self._engine_cancel:
                        break
                    done_groups += 1
                    self.root.after(0, lambda s=size, d=done_groups, t=total_sizes:
                                    self.status_var.set(
                                        f"扫描重复文件 [{d}/{t}] — 正在比对 {format_size(s)} 文件组…"))
                    IndexEngine.compute_hashes_for_size(size)

                if not self._engine_cancel:
                    groups = IndexEngine.get_duplicates()
                    self.root.after(0, lambda: self._on_dup_scan_done(groups))
                else:
                    self.root.after(0, lambda: self.status_var.set("重复文件扫描已取消"))
            except Exception as e:
                self.root.after(0, lambda: self._on_dup_scan_error(str(e)))
            finally:
                self.root.after(0, lambda: self.dup_btn.config(state=tk.NORMAL))

        self._engine_cancel = False
        threading.Thread(target=run, daemon=True).start()

    def _on_dup_scan_done(self, groups: list[dict]):
        """重复扫描完成，展示结果。"""
        self._engine_cancel = False
        if not groups:
            messagebox.showinfo("结果", "没有发现重复文件。")
            self.status_var.set("就绪 — 未发现重复文件")
            return
        total_files = sum(g["count"] for g in groups)
        self.status_var.set(f"发现 {len(groups)} 组重复文件，共 {total_files} 个文件（可释放空间：{self._calc_dup_space(groups)}）")
        self._show_duplicates(groups)

    def _on_dup_scan_error(self, err: str):
        self._engine_cancel = False
        self.status_var.set(f"扫描出错: {err}")

    def _calc_dup_space(self, groups: list[dict]) -> str:
        """计算重复文件可释放空间（每组保留一个，其余可删）。"""
        wasted = 0
        for g in groups:
            size = g["files"][0]["size"] if g["files"] else 0
            wasted += size * (g["count"] - 1)
        return format_size(wasted)

    def _show_duplicates(self, groups: list[dict]):
        """在 treeview 中展示重复文件组，每组用颜色区分。"""
        self._view_mode = "duplicates"
        self._dup_groups = groups
        self._results = []

        self.tree.delete(*self.tree.get_children())
        self._item_to_result = {}

        # 用交替颜色区分不同组
        tag_colors = ["#FFF8E1", "#E8F5E9"]  # 浅黄、浅绿交替

        for g in groups:
            gid = g["group_id"]
            tag = f"dup_{gid}"
            color = tag_colors[gid % len(tag_colors)]
            self.tree.tag_configure(tag, background=color)

            # 先插入组标题行
            header_vals = (
                f"▸ 重复组 #{gid}（{g['count']} 个文件）",
                f"MD5: {g['hash']}",
                "",
                format_size(g["files"][0]["size"]),
                "",
            )
            self.tree.insert("", tk.END, values=header_vals, tags=(tag,))

            for f in g["files"]:
                ext = os.path.splitext(f["name"])[1]
                ext_text = ext.upper().lstrip(".") if ext else ""
                vals = (
                    f["name"],
                    f["path"],
                    ext_text,
                    format_size(f["size"]),
                    f["modified"],
                )
                iid = self.tree.insert("", tk.END, values=vals, tags=(tag,))
                self._item_to_result[iid] = f
                self._results.append(f)

        self.dup_clear_btn.config(state=tk.NORMAL)

    def _exit_duplicate_view(self):
        """退出重复文件视图，返回正常模式。"""
        self._view_mode = "normal"
        self._dup_groups = []
        self.dup_clear_btn.config(state=tk.DISABLED)
        query = self.search_var.get().strip()
        if query:
            self._do_search()
        else:
            self._load_all()

    # ================================================================
    #  文件列表显示与交互
    # ================================================================

    def _refresh_tree(self):
        """清空并重新填充 Treeview（用于排序和首次加载）。"""
        if self._view_mode == "duplicates":
            # 重复视图下，直接重建列表（无组标题，删除后保持一致）
            self.tree.delete(*self.tree.get_children())
            self._item_to_result = {}
            for f in self._results:
                ext = os.path.splitext(f["name"])[1]
                ext_text = ext.upper().lstrip(".") if ext else ""
                vals = (
                    f["name"],
                    f["path"],
                    ext_text,
                    format_size(f["size"]),
                    f["modified"],
                )
                iid = self.tree.insert("", tk.END, values=vals)
                self._item_to_result[iid] = f
            return
        self.tree.delete(*self.tree.get_children())
        self._item_to_result = {}
        for f in self._results:
            ext = os.path.splitext(f["name"])[1]
            ext_text = ext.upper().lstrip(".") if ext else ""
            vals = (
                f["name"],
                f["path"],
                ext_text,
                format_size(f["size"]),
                f["modified"],
            )
            iid = self.tree.insert("", tk.END, values=vals)
            self._item_to_result[iid] = f

    def _get_selected_path(self) -> str | None:
        """获取当前选中文件的完整路径。"""
        sel = self.tree.selection()
        if not sel:
            return None
        result = self._item_to_result.get(sel[0])
        return result["path"] if result else None

    def _on_double_click(self, event):
        """双击文件名：检查文件是否存在后打开。"""
        path = self._get_selected_path()
        if path:
            if not os.path.exists(path):
                messagebox.showwarning("文件不存在", "文件可能已被移动或删除：\n" + path)
                return
            open_with_default(path)

    def _on_right_click(self, event):
        """右键：选中行并弹出上下文菜单。"""
        self.tree.selection_set(self.tree.identify_row(event.y))
        sel = self.tree.selection()
        if sel:
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

    # ================================================================
    #  删除
    # ================================================================

    def _delete_file_recycle(self):
        """删除到回收站（可恢复）。"""
        path = self._get_selected_path()
        if not path:
            return
        name = os.path.basename(path)
        if not messagebox.askyesno("确认删除", f"确定将「{name}」移动到回收站？"):
            return
        try:
            send_to_recycle_bin(path)
            self._remove_from_results(path)
            self._refresh_tree()
            self.status_var.set(f"已删除到回收站: {name}")
        except Exception as e:
            messagebox.showerror("删除失败", str(e))

    def _delete_file_permanent(self):
        """彻底删除文件（不可恢复，有二次确认）。"""
        path = self._get_selected_path()
        if not path:
            return
        name = os.path.basename(path)
        if not messagebox.askyesno("确认彻底删除", f"确定彻底删除「{name}」？\n\n此操作不可恢复！"):
            return
        try:
            permanent_delete(path)
            self._remove_from_results(path)
            self._refresh_tree()
            self.status_var.set(f"已彻底删除: {name}")
        except Exception as e:
            messagebox.showerror("删除失败", str(e))

    def _remove_from_results(self, path: str):
        """从当前结果列表中移除指定路径的文件。"""
        self._results = [f for f in self._results if f["path"] != path]

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
        """加载下一页数据（5000 条），追加到列表末尾。"""
        if not self._has_more or self._loading_more:
            return
        if self._view_mode == "duplicates":
            return
        self._loading_more = True
        offset = len(self._results)
        if self._last_query:
            if self._search_mode == "fuzzy":
                more = IndexEngine.search_fts(self._last_query, limit=5000, offset=offset)
            else:
                more = IndexEngine.search(
                    self._last_query, limit=5000, offset=offset,
                    order_col=self._sort_col, order_desc=not self._sort_asc,
                )
        else:
            more = IndexEngine.load_all(
                limit=5000, offset=offset,
                order_col=self._sort_col, order_desc=not self._sort_asc,
            )
        if more:
            self._results.extend(more)
            self._append_to_tree(more)
            self._has_more = len(more) == 5000
            if self._last_query:
                self.status_var.set(f"搜索「{self._last_query}」— 已显示 {len(self._results):,} 个文件")
            else:
                total = IndexEngine.index_file_count()
                self.status_var.set(f"已显示 {len(self._results):,} / {total:,} 个文件")
        else:
            self._has_more = False
        self._loading_more = False

    def _append_to_tree(self, items):
        """将新加载的文件追加到 Treeview（不重建整个树，避免闪烁）。"""
        for f in items:
            ext = os.path.splitext(f["name"])[1]
            ext_text = ext.upper().lstrip(".") if ext else ""
            vals = (
                f["name"],
                f["path"],
                ext_text,
                format_size(f["size"]),
                f["modified"],
            )
            iid = self.tree.insert("", tk.END, values=vals)
            self._item_to_result[iid] = f

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
        """更新列头排序箭头指示（▲ 升序 / ▼ 降序）。"""
        arrow = " \u25b2" if self._sort_asc else " \u25bc"
        for c in ("name", "path", "type", "size", "modified"):
            base = {"name": "文件名", "path": "路径", "type": "类型", "size": "大小", "modified": "修改时间"}
            self.tree.heading(c, text=base[c] + (arrow if c == self._sort_col else ""))

    def _sort_by(self, col: str):
        """点击列头排序。同列再次点击切换升降序，切换列默认升序。"""
        if self._view_mode == "duplicates":
            self._exit_duplicate_view()
            return

        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True

        query = self.search_var.get().strip()
        if query:
            self._results = IndexEngine.search(
                query, limit=5000, offset=0, order_col=col, order_desc=not self._sort_asc,
            )
            self._has_more = len(self._results) == 5000
        else:
            self._results = IndexEngine.load_all(
                limit=5000, offset=0, order_col=col, order_desc=not self._sort_asc,
            )
            self._has_more = len(self._results) == 5000

        self._refresh_tree()
        self._update_sort_heading()

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
        """窗口关闭时保存列宽并销毁。"""
        self._save_layout()
        self.root.destroy()

    # ================================================================
    #  复制 / 剪切路径
    # ================================================================

    def _copy_path(self, event=None):
        """Ctrl+C：将选中文件的路径复制到剪贴板。"""
        path = self._get_selected_path()
        if path:
            self.root.clipboard_clear()
            self.root.clipboard_append(path)
            self.status_var.set(f"已复制: {os.path.basename(path)}")

    def _cut_path(self, event=None):
        """Ctrl+X：将选中文件的路径剪切到剪贴板。"""
        path = self._get_selected_path()
        if path:
            self.root.clipboard_clear()
            self.root.clipboard_append(path)
            self.status_var.set(f"已剪切: {os.path.basename(path)}")


# ================================================================
#  程序入口
# ================================================================

def main():
    """创建 Tkinter 根窗口并启动应用。"""
    root = tk.Tk()
    # 告诉 Tkinter 当前系统 DPI，确保高 DPI 屏幕文字清晰不模糊
    try:
        dpi = root.winfo_fpixels('1i')
        root.tk.call('tk', 'scaling', dpi / 72.0)
    except Exception:
        pass
    FileSearcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
