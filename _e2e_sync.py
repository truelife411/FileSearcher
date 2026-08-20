# -*- coding: utf-8 -*-
"""全链路端到端：真实文件变更 → USN 事件 → UsnSyncEngine → 临时索引库。"""
import os
import shutil
import sqlite3
import sys
import tempfile
import time

sys.path.insert(0, r"C:\Users\hjf\Documents\代码\FileSearcher\file-searcher")
import file_searcher as fs
import usn_sync as u

# ---- 临时索引库（不碰真实索引）----
tmp_idx = tempfile.mkdtemp(prefix="fs_e2e_idx_")
fs.INDEX_DB = fs.Path(tmp_idx) / "files.db"
fs.INDEX_DIR = fs.Path(tmp_idx)
conn = sqlite3.connect(str(fs.INDEX_DB))
conn.execute("CREATE TABLE files (id INTEGER PRIMARY KEY, name TEXT NOT NULL, "
             "name_lower TEXT NOT NULL, path TEXT NOT NULL UNIQUE, path_lower TEXT NOT NULL, "
             "size INTEGER NOT NULL, modified INTEGER NOT NULL, is_dir INTEGER NOT NULL DEFAULT 0)")
conn.execute("CREATE TABLE dir_refs (dir_path TEXT PRIMARY KEY, file_ref INTEGER NOT NULL)")
conn.execute("CREATE INDEX IF NOT EXISTS idx_dir_refs_ref ON dir_refs(file_ref)")
conn.execute("CREATE VIRTUAL TABLE files_fts USING fts5(name, path, tokenize='trigram')")
conn.commit()
conn.close()

# ---- 测试目录（避开 %TEMP% 排除规则）----
root = os.path.join(r"C:\Users\hjf\Documents\代码\FileSearcher", "_usn_test_data")
shutil.rmtree(root, ignore_errors=True)
os.makedirs(root)
with open(os.path.join(root, "a.txt"), "w", encoding="utf-8") as fh:
    fh.write("a")

# 1) 先建立索引 + dir_refs（模拟初始状态）
fs.IndexEngine.rescan_dirs([root])
print("[1] 初始索引建立")

# 2) UsnSyncEngine，last_usn 指向当前 journal 末尾
eng = u.UsnSyncEngine(fs.IndexEngine, tmp_idx)
h = u.open_volume("C:\\")
q = u.query_journal(h)
serial = u.volume_serial("C:\\")
eng.state.set("C:", serial, q["journal_id"], q["next_usn"])
eng.state.save()
u._kernel32().CloseHandle(h)
print(f"[2] last_usn={q['next_usn']:#x}")

# 3) 触发真实变更：新建文件 + 重命名 + 删除
with open(os.path.join(root, "new_file.txt"), "w", encoding="utf-8") as fh:
    fh.write("new")
os.rename(os.path.join(root, "a.txt"), os.path.join(root, "renamed.txt"))
os.remove(os.path.join(root, "new_file.txt"))
print("[3] 变更已触发（新建+重命名+删除），等待 20s 让 USN 记录刷盘...")
time.sleep(20.0)

# 4) 执行同步
stats = eng.sync_once()
print(f"[4] sync_once: {stats}")

# 5) 验证索引结果
conn = sqlite3.connect(str(fs.INDEX_DB))
names = sorted(r[0] for r in conn.execute("SELECT name FROM files"))
conn.close()
print(f"[5] 索引内容: {names}")
ok = "renamed.txt" in names and "a.txt" not in names and "new_file.txt" not in names
print("RESULT:", "PASS" if ok else "FAIL")

shutil.rmtree(root, ignore_errors=True)
shutil.rmtree(tmp_idx, ignore_errors=True)
