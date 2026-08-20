# -*- coding: utf-8 -*-
"""IndexEngine 增量同步单元测试：rescan_dirs / dir_ref_lookup / 增删改。
使用临时索引库，不碰真实索引。"""
import os
import shutil
import sqlite3
import sys
import tempfile

sys.path.insert(0, r"C:\Users\hjf\Documents\代码\FileSearcher\file-searcher")
import file_searcher as fs
import usn_sync as u

# ---- 准备临时索引库 ----
tmp_idx = tempfile.mkdtemp(prefix="fs_idx_test_")
fs.INDEX_DB = fs.Path(tmp_idx) / "files.db"
fs.INDEX_DIR = fs.Path(tmp_idx)
# 模拟 build_index 建表（ensure_indexes 仅用于老库升级）
conn = sqlite3.connect(str(fs.INDEX_DB))
conn.execute("CREATE TABLE files (id INTEGER PRIMARY KEY, name TEXT NOT NULL, "
             "name_lower TEXT NOT NULL, path TEXT NOT NULL UNIQUE, path_lower TEXT NOT NULL, "
             "size INTEGER NOT NULL, modified INTEGER NOT NULL, is_dir INTEGER NOT NULL DEFAULT 0)")
conn.execute("CREATE TABLE dir_refs (dir_path TEXT PRIMARY KEY, file_ref INTEGER NOT NULL)")
conn.execute("CREATE INDEX IF NOT EXISTS idx_dir_refs_ref ON dir_refs(file_ref)")
conn.execute("CREATE VIRTUAL TABLE files_fts USING fts5(name, path, tokenize='trigram')")
conn.commit()
conn.close()
conn = sqlite3.connect(fs.INDEX_DB)
tables = {r[0] for r in conn.execute(
    "SELECT name FROM sqlite_master WHERE type='table'")}
conn.close()
assert {"files", "dir_refs", "files_fts"} <= tables, tables
print("[ok] 临时库建表:", sorted(tables))

# ---- 准备测试目录树（避开 %TEMP%，它被默认排除规则跳过）----
root = os.path.join(r"C:\Users\hjf\Documents\代码\FileSearcher", "_usn_test_data")
shutil.rmtree(root, ignore_errors=True)
os.makedirs(os.path.join(root, "sub"))
for f in ("a.txt", "sub/b.txt"):
    with open(os.path.join(root, f), "w", encoding="utf-8") as fh:
        fh.write("x")

# ---- 1. rescan 新增目录树 ----
n = fs.IndexEngine.rescan_dirs([root])
conn = sqlite3.connect(fs.INDEX_DB)
cnt = conn.execute("SELECT count(*) FROM files").fetchone()[0]
dir_refs = conn.execute("SELECT count(*) FROM dir_refs").fetchone()[0]
fts = conn.execute("SELECT count(*) FROM files_fts").fetchone()[0]
conn.close()
print(f"[1] rescan 新增: changed={n} files={cnt} dir_refs={dir_refs} fts={fts}")
# root 自身条目由父目录重扫处理；此处应有 a.txt + sub + b.txt 三项
assert n >= 3 and cnt == 3 and dir_refs >= 2 and fts == cnt, "rescan 新增失败"

# ---- 2. dir_ref_lookup 翻译 ----
ref = u.dir_file_ref(root)
path = fs.IndexEngine.dir_ref_lookup(ref)
print(f"[2] 引用号翻译: {ref:#x} -> {path}")
assert path and os.path.normcase(path) == os.path.normcase(root), "翻译失败"
ref_sub = u.dir_file_ref(os.path.join(root, "sub"))
assert fs.IndexEngine.dir_ref_lookup(ref_sub) == os.path.join(root, "sub")
print("[2] 子目录翻译 ok")

# ---- 3. 删除文件后重扫 ----
os.remove(os.path.join(root, "a.txt"))
n = fs.IndexEngine.rescan_dirs([root])
conn = sqlite3.connect(fs.INDEX_DB)
cnt = conn.execute("SELECT count(*) FROM files").fetchone()[0]
fts = conn.execute("SELECT count(*) FROM files_fts").fetchone()[0]
conn.close()
print(f"[3] 删除 a.txt 后重扫: changed={n} files={cnt} fts={fts}")
assert cnt == 2 and fts == 2, "删除同步失败"

# ---- 4. 新增文件后重扫 ----
with open(os.path.join(root, "c.txt"), "w", encoding="utf-8") as fh:
    fh.write("y")
n = fs.IndexEngine.rescan_dirs([root])
conn = sqlite3.connect(fs.INDEX_DB)
cnt = conn.execute("SELECT count(*) FROM files").fetchone()[0]
row = conn.execute("SELECT name, size FROM files WHERE name='c.txt'").fetchone()
conn.close()
print(f"[4] 新增 c.txt 后重扫: changed={n} files={cnt} row={row}")
assert cnt == 3 and row and row[1] == 1, "新增同步失败"

# ---- 5. 删除整个目录（rescan 时目录不存在 → remove_tree）----
shutil.rmtree(os.path.join(root, "sub"))
n = fs.IndexEngine.rescan_dirs([os.path.join(root, "sub")])
conn = sqlite3.connect(fs.INDEX_DB)
cnt = conn.execute("SELECT count(*) FROM files").fetchone()[0]
dr = conn.execute("SELECT count(*) FROM dir_refs WHERE dir_path LIKE ?",
                  (os.path.join(root, "sub") + "%",)).fetchone()[0]
conn.close()
print(f"[5] 删除子目录后重扫: changed={n} files={cnt} sub_dir_refs={dr}")
assert cnt == 1 and dr == 0, "目录删除同步失败"

# ---- 6. rename_path 级联 + dir_refs 维护 ----
old = os.path.join(root, "c.txt")
new = os.path.join(root, "c2.txt")
os.rename(old, new)
fs.IndexEngine.rename_path(old, new)
conn = sqlite3.connect(fs.INDEX_DB)
row = conn.execute("SELECT path FROM files WHERE name='c2.txt'").fetchone()
conn.close()
print(f"[6] rename: {row}")
assert row and row[0] == new, "rename 同步失败"

shutil.rmtree(root, ignore_errors=True)
shutil.rmtree(tmp_idx, ignore_errors=True)
print("\nALL PASS")
