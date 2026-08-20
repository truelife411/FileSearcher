# -*- coding: utf-8 -*-
"""USN Journal 增量同步（Windows NTFS 变更日志）。

机制：NTFS 卷自带变更日志（USN Journal），记录卷上所有文件/目录变更。
本模块负责：
  * 枚举 NTFS 固定盘、查询/激活/读取各卷的 USN Journal（ctypes 直调 Win32）
  * 记录每卷的 last_usn（增量游标），断档/换卷时通知调用方全量重建
  * 把 USN 记录翻译成文件系统动作（新增/删除/重命名/修改），
    通过回调调用索引引擎完成增量更新

权限：读 USN Journal 无需管理员权限（已实测）；创建 journal 需要管理员。
"""
import ctypes
import ctypes.wintypes as wt
import json
import os
import threading

_USN_STATE_FILE = "usn_state.json"

# ---- Win32 常量 ----
GENERIC_READ = 0x80000000
FILE_SHARE_READ = 0x1
FILE_SHARE_WRITE = 0x2
OPEN_EXISTING = 3
FILE_READ_ATTRIBUTES = 0x80
FILE_FLAG_BACKUP_SEMANTICS = 0x02000000
DRIVE_FIXED = 3
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

FSCTL_QUERY_USN_JOURNAL = 0x000900F4
FSCTL_READ_USN_JOURNAL = 0x000900BB
FSCTL_CREATE_USN_JOURNAL = 0x000900E7

# USN Reason 掩码
USN_REASON_DATA_OVERWRITE = 0x1
USN_REASON_DATA_EXTEND = 0x2
USN_REASON_DATA_TRUNCATION = 0x4
USN_REASON_FILE_CREATE = 0x100
USN_REASON_FILE_DELETE = 0x200
USN_REASON_RENAME_OLD_NAME = 0x4000
USN_REASON_RENAME_NEW_NAME = 0x8000
USN_REASON_CLOSE = 0x80000000

# 读取时关注的原因：创建/删除/重命名 + 关闭（ReturnOnlyOnClose=1 时
# 修改类原因会合并进 CLOSE 记录，故 CLOSE 必须请求）
USN_REASON_MASK = (USN_REASON_FILE_CREATE | USN_REASON_FILE_DELETE |
                   USN_REASON_RENAME_OLD_NAME | USN_REASON_RENAME_NEW_NAME |
                   USN_REASON_CLOSE)

_k32 = None


def _kernel32():
    """按需初始化 kernel32 句柄函数（惰性，避免导入即初始化）。"""
    global _k32
    if _k32 is None:
        k = ctypes.WinDLL("kernel32", use_last_error=True)
        k.CreateFileW.restype = wt.HANDLE
        k.CreateFileW.argtypes = [wt.LPCWSTR, wt.DWORD, wt.DWORD, ctypes.c_void_p,
                                  wt.DWORD, wt.DWORD, wt.HANDLE]
        k.CloseHandle.argtypes = [wt.HANDLE]
        k.DeviceIoControl.restype = wt.BOOL
        k.DeviceIoControl.argtypes = [wt.HANDLE, wt.DWORD, ctypes.c_void_p, wt.DWORD,
                                      ctypes.c_void_p, wt.DWORD,
                                      ctypes.POINTER(wt.DWORD), ctypes.c_void_p]
        k.GetLogicalDrives.restype = wt.DWORD
        k.GetDriveTypeW.restype = wt.UINT
        k.GetDriveTypeW.argtypes = [wt.LPCWSTR]
        k.GetVolumeInformationW.restype = wt.BOOL
        k.GetVolumeInformationW.argtypes = [
            wt.LPCWSTR, ctypes.c_wchar_p, wt.DWORD, ctypes.POINTER(wt.DWORD),
            ctypes.POINTER(wt.DWORD), ctypes.POINTER(wt.DWORD),
            ctypes.c_wchar_p, wt.DWORD]
        _k32 = k
    return _k32


class READ_USN_JOURNAL_DATA(ctypes.Structure):
    _fields_ = [
        ("StartUsn", ctypes.c_int64),
        ("ReasonMask", wt.DWORD),
        ("ReturnOnlyOnClose", wt.DWORD),
        ("Timeout", ctypes.c_int64),
        ("BytesToWaitFor", ctypes.c_int64),
        ("UsnJournalID", ctypes.c_int64),
    ]


class USN_RECORD_V2(ctypes.Structure):
    _fields_ = [
        ("RecordLength", wt.DWORD),
        ("MajorVersion", wt.WORD),
        ("MinorVersion", wt.WORD),
        ("FileReferenceNumber", ctypes.c_int64),
        ("ParentFileReferenceNumber", ctypes.c_int64),
        ("Usn", ctypes.c_int64),
        ("TimeStamp", ctypes.c_int64),
        ("Reason", wt.DWORD),
        ("SourceInfo", wt.DWORD),
        ("SecurityId", wt.DWORD),
        ("FileAttributes", wt.DWORD),
        ("FileNameLength", wt.WORD),
        ("FileNameOffset", wt.WORD),
        ("FileName", ctypes.c_wchar * 1),
    ]


class BY_HANDLE_FILE_INFORMATION(ctypes.Structure):
    """注意：FILETIME 必须用 2×DWORD（对齐 4），用 c_int64 会在 x64 上
    产生 8 字节对齐 padding 导致后续字段错位（nFileIndex 读错）。"""
    _fields_ = [
        ("dwFileAttributes", wt.DWORD),
        ("ftCreationTime", wt.DWORD * 2),
        ("ftLastAccessTime", wt.DWORD * 2),
        ("ftLastWriteTime", wt.DWORD * 2),
        ("dwVolumeSerialNumber", wt.DWORD),
        ("nFileSizeHigh", wt.DWORD),
        ("nFileSizeLow", wt.DWORD),
        ("nNumberOfLinks", wt.DWORD),
        ("nFileIndexHigh", wt.DWORD),
        ("nFileIndexLow", wt.DWORD),
    ]


def is_admin() -> bool:
    """当前进程是否以管理员权限运行。"""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def list_ntfs_volumes() -> list[str]:
    """枚举 NTFS 固定盘，返回盘符根路径列表（如 ['C:\\', 'D:\\']）。"""
    k = _kernel32()
    drives = []
    mask = k.GetLogicalDrives()
    for i in range(26):
        if not (mask & (1 << i)):
            continue
        root = f"{chr(ord('A') + i)}:\\"
        if k.GetDriveTypeW(root) != DRIVE_FIXED:
            continue
        fs = ctypes.create_unicode_buffer(32)
        ok = k.GetVolumeInformationW(root, None, 0, None, None, None, fs, 32)
        if ok and fs.value.upper() == "NTFS":
            drives.append(root)
    return drives


def volume_serial(root: str) -> int | None:
    """返回卷序列号（用于检测换卷/格式化）。"""
    k = _kernel32()
    serial = wt.DWORD()
    ok = k.GetVolumeInformationW(root, None, 0, ctypes.byref(serial), None, None,
                                 None, 0)
    return int(serial.value) if ok else None


def open_volume(root: str) -> int:
    """打开卷句柄（设备名形式 \\\\.\\C:）。失败返回 INVALID_HANDLE_VALUE。"""
    k = _kernel32()
    path = "\\\\.\\" + root.rstrip("\\/:") + ":"
    return k.CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE,
                         None, OPEN_EXISTING, 0, None)


def query_journal(h: int) -> dict | None:
    """查询卷的 journal 信息；未激活返回 None。"""
    k = _kernel32()
    buf = (ctypes.c_byte * 64)()
    n = wt.DWORD()
    if not k.DeviceIoControl(h, FSCTL_QUERY_USN_JOURNAL, None, 0,
                             buf, len(buf), ctypes.byref(n), None):
        return None
    vals = ctypes.cast(buf, ctypes.POINTER(ctypes.c_int64))
    return {
        "journal_id": vals[0],
        "first_usn": vals[1],
        "next_usn": vals[2],
        "lowest_valid": vals[3],
    }


def create_journal(h: int) -> bool:
    """激活卷的 USN Journal（需要管理员权限）。"""
    k = _kernel32()
    buf = (ctypes.c_byte * 8)()   # CREATE_USN_JOURNAL_DATA: 最大/分配大小(4+4)
    ctypes.memset(buf, 0, 8)
    n = wt.DWORD()
    return bool(k.DeviceIoControl(h, FSCTL_CREATE_USN_JOURNAL, buf, 8,
                                  buf, 8, ctypes.byref(n), None))


def read_records(h: int, start_usn: int, journal_id: int,
                 reason_mask: int = USN_REASON_MASK,
                 buffer_size: int = 1 << 20) -> tuple[list[dict], int]:
    """从 start_usn 读取增量记录（ReturnOnlyOnClose=1）。

    返回 (records, next_usn)：records 为记录字典列表，next_usn 为下一条位置；
    journal 断档（返回的 first 大于 start_usn）时 next_usn 为 -1。
    """
    k = _kernel32()
    records: list[dict] = []
    cursor = start_usn
    for _round in range(64):   # 单轮同步最多 64 次读取（防异常死循环）
        in_data = READ_USN_JOURNAL_DATA()
        in_data.StartUsn = cursor
        in_data.ReasonMask = reason_mask
        in_data.ReturnOnlyOnClose = 1
        in_data.Timeout = 0
        in_data.BytesToWaitFor = 0
        in_data.UsnJournalID = journal_id
        out = (ctypes.c_byte * buffer_size)()
        n = wt.DWORD()
        if not k.DeviceIoControl(h, FSCTL_READ_USN_JOURNAL,
                                 ctypes.byref(in_data), ctypes.sizeof(in_data),
                                 out, len(out), ctypes.byref(n), None):
            return records, -1
        nbytes = n.value
        if nbytes < 8:
            return records, cursor
        next_usn = ctypes.cast(out, ctypes.POINTER(ctypes.c_int64)).contents.value
        if next_usn <= cursor:   # 无新记录（或已到末尾）
            return records, next_usn
        pos = 8
        base = ctypes.addressof(out)
        while pos + 4 <= nbytes:
            rl = ctypes.cast(base + pos, ctypes.POINTER(wt.DWORD)).contents.value
            if rl < 32 or pos + rl > nbytes:
                break
            rec = ctypes.cast(base + pos, ctypes.POINTER(USN_RECORD_V2)).contents
            fname = ""
            try:
                fname = ctypes.wstring_at(base + pos + rec.FileNameOffset,
                                          rec.FileNameLength // 2)
            except Exception:
                pass
            records.append({
                "usn": rec.Usn,
                "file_ref": rec.FileReferenceNumber,
                "parent_ref": rec.ParentFileReferenceNumber,
                "reason": rec.Reason,
                "name": fname,
                "is_dir": bool(rec.FileAttributes & 0x10),  # FILE_ATTRIBUTE_DIRECTORY
            })
            pos += rl
        cursor = next_usn
        if nbytes < buffer_size:   # 本次已读完
            return records, next_usn
    return records, cursor


def dir_file_ref(path: str) -> int | None:
    """打开目录句柄获取其 NTFS 文件引用号；失败返回 None。"""
    k = _kernel32()
    k.GetFileInformationByHandle.restype = wt.BOOL
    k.GetFileInformationByHandle.argtypes = [
        wt.HANDLE, ctypes.POINTER(BY_HANDLE_FILE_INFORMATION)]
    h = k.CreateFileW(path, FILE_READ_ATTRIBUTES,
                      FILE_SHARE_READ | FILE_SHARE_WRITE, None, OPEN_EXISTING,
                      FILE_FLAG_BACKUP_SEMANTICS, None)
    if h == INVALID_HANDLE_VALUE:
        return None
    try:
        info = BY_HANDLE_FILE_INFORMATION()
        if not k.GetFileInformationByHandle(h, ctypes.byref(info)):
            return None
        return (int(info.nFileIndexHigh) << 32) | int(info.nFileIndexLow)
    finally:
        k.CloseHandle(h)


class UsnState:
    """按卷序列号持久化 last_usn 游标（usn_state.json）。"""

    def __init__(self, index_dir):
        self._file = os.path.join(index_dir, _USN_STATE_FILE)
        self._lock = threading.Lock()
        self._data = self._load()

    def _load(self) -> dict:
        try:
            with open(self._file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def save(self):
        with self._lock:
            try:
                os.makedirs(os.path.dirname(self._file), exist_ok=True)
                with open(self._file, "w", encoding="utf-8") as f:
                    json.dump(self._data, f, ensure_ascii=False, indent=1)
            except Exception:
                pass

    def get(self, root: str) -> dict | None:
        return self._data.get(root.rstrip("\\/"))

    def set(self, root: str, serial: int, journal_id: int, last_usn: int):
        self._data[root.rstrip("\\/")] = {
            "serial": serial, "journal_id": journal_id, "last_usn": last_usn,
        }


class UsnSyncEngine:
    """一轮 USN 增量同步：遍历 NTFS 卷读增量，翻译成动作并应用。

    事件翻译依赖索引库的 dir_refs 映射（父引用号 → 目录路径）；
    翻译失败的记录计入 untranslated（映射缺失时靠目录重扫自愈）。
    """

    def __init__(self, engine, index_dir, exclude_dirs=(), exclude_paths=()):
        self.engine = engine            # IndexEngine（注入避免循环导入）
        self.state = UsnState(index_dir)
        self.exclude_dirs = set(exclude_dirs)
        self.exclude_paths = list(exclude_paths)

    def _excluded(self, path: str) -> bool:
        lower = path.lower()
        if os.path.basename(os.path.normpath(path)).lower() in self.exclude_dirs:
            return True
        return any(pat in lower for pat in self.exclude_paths)

    def sync_once(self) -> dict:
        """执行一轮同步。返回统计：
        {"changed", "rescan_dirs", "untranslated", "reset", "error"}"""
        stats = {"changed": 0, "rescan_dirs": 0, "untranslated": 0,
                 "reset": False, "error": None}
        dirty: set[str] = set()
        deletes: list[dict] = []
        renames: list[dict] = []
        try:
            vols = list_ntfs_volumes()
        except Exception as e:
            stats["error"] = str(e)
            return stats

        for root in vols:
            h = open_volume(root)
            if h == INVALID_HANDLE_VALUE:
                continue
            try:
                q = query_journal(h)
                if q is None:
                    continue   # journal 未激活 → 跳过该卷（App 层可提示）
                serial = volume_serial(root)
                st = self.state.get(root)
                if st is None:
                    # 首次接触该卷：仅初始化游标（索引来自全量扫描，天然一致）
                    self.state.set(root, serial, q["journal_id"], q["next_usn"])
                    self.state.save()
                    continue
                ok = True
                if st.get("serial") != serial or \
                        st.get("journal_id") != q["journal_id"]:
                    ok = False   # 换卷/格式化/journal 重建
                elif st.get("last_usn", 0) < q["lowest_valid"] or \
                        st.get("last_usn", 0) < q["first_usn"]:
                    ok = False   # 记录已被截断
                if not ok:
                    stats["reset"] = True
                    self.state.set(root, serial, q["journal_id"], q["next_usn"])
                    self.state.save()
                    continue
                records, next_usn = read_records(h, st["last_usn"], q["journal_id"])
                if next_usn < 0:
                    stats["reset"] = True
                    self.state.set(root, serial, q["journal_id"], q["next_usn"])
                    self.state.save()
                    continue
                for rec in records:
                    reason = rec["reason"]
                    if reason & USN_REASON_FILE_DELETE:
                        deletes.append(rec)
                    elif reason & (USN_REASON_RENAME_OLD_NAME |
                                   USN_REASON_RENAME_NEW_NAME):
                        renames.append(rec)
                    elif reason & USN_REASON_FILE_CREATE:
                        p = self._translate_parent(rec)
                        if p:
                            dirty.add(p)
                    elif reason & USN_REASON_CLOSE:
                        p = self._translate_parent(rec)
                        if p:
                            dirty.add(p)
                if next_usn > st.get("last_usn", 0):
                    self.state.set(root, serial, q["journal_id"], next_usn)
                    self.state.save()
            finally:
                try:
                    _kernel32().CloseHandle(h)
                except Exception:
                    pass

        # ---- 应用：删除 ----
        for rec in deletes:
            path = self._translate(rec)
            if path and not self._excluded(path):
                try:
                    self.engine.remove_paths([path])
                    stats["changed"] += 1
                    continue
                except Exception:
                    pass
            stats["untranslated"] += 1

        # ---- 应用：重命名（同引用号配对 旧名+新名）----
        by_ref: dict[int, list[dict]] = {}
        for rec in renames:
            by_ref.setdefault(rec["file_ref"], []).append(rec)
        for _ref, recs in by_ref.items():
            old = new = None
            for r in recs:
                if r["reason"] & USN_REASON_RENAME_OLD_NAME:
                    old = r
                if r["reason"] & USN_REASON_RENAME_NEW_NAME:
                    new = r
            if old and new:
                old_path = self._translate(old)
                parent_new = self._translate_parent(new)
                new_path = os.path.join(parent_new, new["name"]) if parent_new else None
                if old_path and new_path and not self._excluded(old_path):
                    try:
                        self.engine.rename_path(old_path, new_path)
                        stats["changed"] += 1
                        continue
                    except Exception:
                        pass
            for r in recs:
                p = self._translate_parent(r)
                if p:
                    dirty.add(p)

        # ---- 应用：dirty 目录重扫 ----
        dirty.discard(None)
        if dirty:
            try:
                n = self.engine.rescan_dirs(sorted(dirty),
                                            self.exclude_dirs, self.exclude_paths)
                stats["changed"] += n
                stats["rescan_dirs"] = len(dirty)
            except Exception as e:
                stats["error"] = str(e)
        return stats

    def _translate(self, rec: dict) -> str | None:
        """记录 → 完整路径（父引用号查 dir_refs）。查不到返回 None。"""
        parent = self.engine.dir_ref_lookup(rec["parent_ref"])
        if not parent:
            return None
        return os.path.join(parent, rec["name"])

    def _translate_parent(self, rec: dict) -> str | None:
        """记录 → 父目录完整路径。"""
        return self.engine.dir_ref_lookup(rec["parent_ref"])
