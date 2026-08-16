# File Searcher 项目笔记

## 用户偏好
- 每次改完代码自动做本地 `git commit`
- `git push` 到 GitHub 需要用户明确要求
- **【铁律·2026-08-15 用户定】改完代码一律「PyInstaller 打包 exe 到桌面」，绝不再从工具链拉起 FS 进程**。流程：杀旧实例（让用户自己退/关，或确认无残留）→ `python -m PyInstaller --noconfirm --clean FileSearcher.spec` → 覆盖 `C:\Users\hjf\Desktop\FileSearcher.exe` → **让用户自己双击桌面 exe 启动**（用户在自己的 Session 1 桌面会话里启动，一切正常）。
- **【血泪教训·会话隔离】绝不能用 `Start-Process`/subprocess 从 WorkBuddy 工具链启动 FS 验证**！工具拉起的进程跑在隔离沙箱会话（非用户 Session 1），会导致：① `os.startfile`/`ShellExecute`/`explorer.exe` 拉起另一个 GUI 程序（视频播放器）时**静默失败或卡死**（双击视频无反应/空窗）② 任务栏图标、窗口行为异常 ③ 跨会话杀不掉进程残留。这些「bug」全是假象，根因是进程不在用户桌面会话。**双击打开文件异常时，第一反应必须是「FS 是不是被工具拉起的」，别再怀疑 os.startfile 写法/标志位/线程**（这些都试过、都不是根因，纯弯路）。
- × 关闭按钮 = 最小化到托盘，最小化按钮也缩到托盘，彻底退出靠托盘右键菜单
- 托盘恢复窗口时：最大化 + 搜索框全选/聚焦

## 打开文件的最终实现（2026-08-15 定稿）
- `open_with_default`：**主线程直接 `os.startfile(os.path.normpath(path))`**，不包后台线程、不加 ShellExecute 标志位。与文档中心/资源管理器/迅雷完全同款。os.startfile 本就异步不阻塞 UI。
- 失败的弯路（勿再走）：后台守护线程调 startfile、ShellExecuteExW 加 `SEE_MASK_NOASYNC`/`DDEWAIT`、怀疑播放器单实例 emit 丢失——全是被「会话隔离」假象误导。

## 技术栈
- **运行/打包用系统 Python 3.12**（`C:\Users\hjf\AppData\Local\Programs\Python\Python312\python.exe`，PyInstaller 6.20 已装）— managed 3.13 没装 pystray/PIL！开发语法检查用 `ast.parse`（不要 `py_compile`，会写 __pycache__ 被沙箱拦）
- Tkinter + SQLite (WAL模式)；系统托盘：pystray + Pillow；**打包：`python -m PyInstaller --noconfirm --clean FileSearcher.spec`**（spec 已配好 tcl/tk collect_all + **`icon='app.ico'`**（项目根，程序图标：青绿渐变圆角方块+白色放大镜，16~256px 六档）），--onefile --windowed 含在 spec 里），产物 `dist\FileSearcher.exe`（约23MB）→ 复制到桌面
- **exe 图标验证**：Shell32 ExtractAssociatedIcon 提取 exe 图标与 app.ico 逐像素对比（0% 差异）；桌面图标不刷新时清 `%LocalAppData%\Microsoft\Windows\Explorer\iconcache_*.db` + 重启 explorer（ie4uinit -show 刷新）
- **拖拽依赖 tkinterdnd2（0.6.2，已装）**：tkdnd 是独立 Tcl 扩展，spec 里 `collect_all('tkinterdnd2')`（未装则 try/except 跳过，运行时降级+状态栏提示）；代码里 `_setup_drag_drop` 按平台/tcl9 把 `tkinterdnd2/tkdnd/<平台>` 加 auto_path 再 `package require tkdnd`
- GitHub: https://github.com/truelife411/FileSearcher

## 索引引擎（2026-08-16 升级：FTS5 + 目录队列扫描）
- **搜索 = FTS5 trigram 全文索引**：`files_fts(name, path, tokenize='trigram')` 虚拟表，rowid 对应 files.id；搜索词 ≥3 字符走 `MATCH "词"`（子串语义，trigram 下 1-2 字符 MATCH 返回空！），<3 字符回退 LIKE（`%`/`_`/`\` 已转义 + ESCAPE）；FTS 命中 rowid 落 TEMP 表 JOIN files 排序分页
- **FTS 同步**：重建时随建随填；老库 `ensure_indexes` 只建空表 + 启动后台线程 `ensure_fts_backfill` 回填（持有 _build_lock，避开并发重建）；`remove_paths`/`rename_path` 用 `_fts_sync_row`（先删后插，name/path=None 仅删）同步；`_db_schema` 检测 fts 表有数据才启用（空表视为未回填走 LIKE）
- **modified 列存 epoch 秒（INTEGER）**：扫描期不再格式化字符串；UI 层 `_tree_item_values` 格式化，兼容老库 TEXT（`_db_schema` 的 modified_is_text 分支，时间过滤分别用 strftime 字符串或 int(timestamp)）
- **扫描 = 目录任务队列**：`dir_queue` + Condition 计数（pending_dirs[0]/active_workers），worker 一次领一个目录扫、子目录入队 → 单盘也 4 worker 并行；取消/异常路径 scan_done 收尾
- **每条目 1 次 stat**：`S_ISDIR/S_ISREG(st.st_mode)` 判断类型（Windows 上 junction 的 stat(follow_symlinks=False) 为 S_IFLNK → 与旧 is_dir(False) 行为一致被跳过）
- `_db_schema()` 带 db_mtime 缓存的 PRAGMA（has_is_dir/modified_is_text/fts_ok），老库 TEXT 列不迁移
- 实测全盘 101,256 文件：扫描 7.9s + 优化(FTS回填+索引) 1.8s
- **日志**：`_diag` 写 `~/.file_searcher_index/debug.log`（不再硬编码开发机路径），>1MB 轮转 debug.log.1；`report_callback_exception` 已接线（Tkinter 回调异常写 [tk-error] 行）
- **启动竞态坑位（2026-08-16 实测）**：_defer_initial_load 轮询映射时若先于 resize 事件更新 _page_size，resize 会因"页大小未变"跳过查询 → 首次查询永不执行、列表空白。修法：映射后一次性 _initial_query_done 标志保证至少触发一次 _load_all（默认全部文件、size DESC）
- **键盘/滚轮导航**：FileTable 内实现——滚轮翻页（累积 delta≥120 翻一页，防抖）、↑/↓ 移选中（clamp）、PgUp/PgDn 翻页（on_scroll_page 回调），Linux Button-4/5、macOS delta 缩放
- **搜索重置页码（2026-08-16）**：_do_search(防抖搜索)与 _clear_search(清空) 必须 `_page = 1`；_sort_by 本就重置；_goto_page/resize/索引刷新保持页码
- **搜索框交互（2026-08-16 用户定）**：「一点即改 + 大清空键」。未聚焦时点击框内任意处（放大镜区/输入区/留白）→ 聚焦并**全选文字**（浏览器地址栏习惯，直接打字即替换，解决"修改搜索内容"）；已聚焦时点击 → 正常光标定位（再点一下可精确编辑）。✕ 清空按钮：半径随框高 0.32 倍放大（_clear_r），**点击热区与圆外沿对齐**（_clear_hit_start = 宽-44-r），悬停 accent 填充+白字+手型光标（_on_motion 仅在状态翻转时重绘）。**【坑位】输入区宽度预留必须随按钮半径同步**（_entry_width = w-106-r）：固定 126px 在高 DPI（240）+大字号（text_pt18）下框高达 137px、按钮半径 43px，曾致按钮左缘被输入区窗口遮住 19px（窗口永远盖在画布图元之上）。无文字时恢复 108px。Esc 清空 / Ctrl+L 聚焦全选保留。几何验证脚本 searchbox_geo_test.py（96dpi 基准高 58 vs 用户实际高 137 双档检查重叠）
- **快捷移动（2026-08-16 用户定；菜单名 2026-08-16 改「快捷移动上层目录到 xxx」）**：右键菜单「快捷移动上层目录到 xxx」（重命名上方，多选置灰，未配置显示"（未配置）"）。作用：选中文件→移动其父目录，选中文件夹→移动本身。设置项 quick_move_dir（设置-常规页，filedialog.askdirectory 选择，点击路径标签也可重新选择）。**不弹确认**（用户要求快捷）；安全边界：磁盘根不可移/目标在源内部不可移/目标已有同名提示取消/保护目录拦截（_block_if_protected）。**Windows 走 SHFileOperation FO_MOVE（shell_move，系统原生进度框：进度条/剩余时间/取消/冲突询问，与资源管理器同机制；取消=DE_OPERATION_CANCELLED 0x75）**，非 Windows 后台线程 shutil.move 兜底。成功后 IndexEngine.rename_path 级联同步索引（无需重扫）→ 刷新结果。索引同步失败仅忽略（重建时纠正）
- **删除所在目录（2026-08-16）**：右键删除组下方新增「删除所在目录到回收站 / 彻底删除所在目录」（多选置灰）。语义同快捷移动：文件→父目录，文件夹→本身；**有确认框**（目录删除比文件危险，文案明示"及其全部内容"，**确认框带醒目全路径条**——_dialog_confirm 的 path_highlight 参数：等宽字体+描边+红色，超长换行计入弹窗高度）。复用 send_to_recycle_bin/permanent_delete（SHFileOperation 原生支持目录）+ IndexEngine.remove_paths 递归删索引。磁盘根目录拦截（_selected_dir_path/_is_drive_root 可复用）
- **按星级重命名（2026-08-16 用户定）**：右键重命名下方「按星级重命名」（多选置灰）。_dialog_stars 星选弹窗：7 星点选（点第 N 颗=1~N 星点亮，金色 warning/灰 muted_2）+ 等宽预览条（★×N + 去星标名）。**替换语义（幂等）**：`re.sub(r"^(★+\s*)+", "", name)` 剥旧星再添新星；弹窗自动识别当前星数预选。rename_file + rename_path 同步索引。★ U+2605 合法文件名字符
- **保护目录（2026-08-16 用户定，防高危误操作）**：设置第四页「保护目录」（⛔ 图标导航），预置 C:\Windows / Program Files / Program Files (x86) / ProgramData。**拦截三个目录级操作**：快捷移动、删除到回收站、彻底删除所在目录（_block_if_protected 在确认框之前拦，弹 danger 警告 + path_highlight 显示命中项）。**匹配规则（用户拍板：拦本身+上层，不拦子目录）**：`_is_protected_dir` = T==P 或 P.startswith(T+sep)（normcase+normpath+去尾分隔符，大小写不敏感、分隔符边界不误伤「重要」vs「重要2」）。存储于 settings.json `protected_dirs`（DEFAULT_SETTINGS 预置，load_settings 自动补齐老配置）。设置页增删改：filedialog.askdirectory 目录选择器（不用手输），_protect_add/_protect_edit/_protect_delete/_protect_save 即时保存。9 组边界用例已验证（本身/上层拦、子目录放行、边界不误伤、大小写）
- build/ dist/ 已加入 .gitignore（构建产物不入库）；app.ico 是项目资源（入库）

## UI 设计基线（2026-08-14「凝脂纸感」，方向 C，当前线上版）
- **凝脂主题**（B 墨玉实机被否决后改选 C）：宣纸暖白底 #F5F3EE + 黛青 accent #2E6E66（渐变 #3A837A→#255A53）+ 墨色文字 #23282C。深色变体「墨玉」#0A0C10/#6FD8C8 保留在 THEMES["dark"]。默认主题 = light。sel_text = 选中行/激活态文字色
- **柔和配色（2026-08-16 定稿，共 4 主题）**：纯白扎眼/纯黑压抑 → 新增「雾沙」mist（**米色**浅色 #F2ECDB/#F8F4E7/#403B2F/#5F7D72，黄白调亮而干净）与「天青」jade（**青瓷浅色** #E6EDE9/#F1F5F1/#33413C/#5C8A7F，用户从 10+ 个候选方案中拍板选定，替换原青墨深色）。**「跟随系统」已按用户要求移除**（_resolve_theme 对未知/旧 system 值落回 mist）；默认主题 = mist。BADGE_STYLES 同步四套（雾沙徽章底色米白调、天青为浅色青瓷系）。主题卡 4 张一行（s(138)），迷你预览色按主题查 THEMES 表绘制。设计图脚本 theme_preview2.py 可复用（候选色板 → 模拟窗口+色板 PNG）
- **纸感语言 = 小圆角 + 细线**：结果卡片 6 / 按钮 8 / 搜索框 6 / 弹窗 10；分隔线 row_line、描边 border 1px；无网格线
- **布局（C 版定型）**：标题栏 40px（渐变 logo+副标题）；搜索区 = 居中搜索框（max 980、高 62、placeholder「搜索全盘文件…」、聚焦 accent 描边、有内容时右侧 ✕ 清空）；下方 chips 类型筛选排（全部/文件夹/文档/图片/视频/音频/压缩包/代码，接入 filters.type）；结果纸面卡片（_layout_result_container 圆角 6）内嵌 FileTable + 底部分页栏（左"共 N 个结果"/右 ‹ 1/12 ›）；**底部快捷键栏 44px**（左：状态点+文字 + kbd 提示 ↵打开/⌥↵定位/^C复制/⌫删除/⇵翻页；右：⟳重建索引 ghost 钮 + ⚙ 钮）
- **呼吸式表格 FileTable**：行分隔线、选中行 selected 底+左侧 6px accent 指示条、类型徽章（BADGE_STYLES[theme][kind]=(fg,bg)）、表头加粗+排序 accent+自绘小三角、列宽拖动（_fit_cols 按序压缩/拉宽、_drag_pref 刚性列、_fitted_width 1:1 手感、layout.json 跨显示器换算）。**路径列用 _path_middle_ellipsis 中间省略**（保留首尾目录）
- **字号层级不变**：MICRO=9/SMALL=10/BODY=12/INPUT=13/LG=14/XL=15 base pt；**主字号 = text_pt（表格正文基准，10~40 可调）**，其余文字按固定比例自适应（表头/路径=0.83×、按钮=1×、搜索框=1.08×、标题≈1.17×）；_f/_s 全局缩放；数字列 Consolas。设置页用 ttk.Scale 滑动条（TextPt.Horizontal.TScale 样式），**拖动只预览数值、松手才保存生效**（ButtonRelease-1/KeyRelease 提交 → _apply_text_scale → after(80) 重算页大小）
- **全自绘弹窗体系**：_DialogShell（透明角+阴影+置顶钳制）、_dialog_confirm、_dialog_input（分段胶囊）、CtxMenu、ToggleSwitch、RoundEntry
- **设置窗口**：左导航三页（常规/排除列表/关于）；主题卡 = 墨玉(深)/凝脂(浅)/跟随系统，切换 destroy+重开；排除列表 ttk.Treeview（Ex.Treeview）
- **分页**：页大小 = 表头下方可视行数（_compute_page_size），滚轮翻页，resize 重算
- Windows 无边框：overrideredirect + WM_NCHITTEST；启动铺满工作区
- **缩放坑位（2026-08-16 两轮实测定稿）**：Tk 8.6 按 96 DPI 布局，tk scaling 污染 winfo_fpixels；程序是 per-monitor DPI aware（SetProcessDPIAwareness(2)），Windows 不做位图拉伸，**一切缩放必须自己按 DPI 算**。① 曾把 `_font_scale = text_pt/(FONT_BODY×dpi_scale)` 把 dpi_scale 放分母，与 `_f/_s` 里的 ×dpi_scale 完全互抵 → 字号不随 DPI 变，250% 屏上只有系统应用 40% 大。② 改为全量 dpi/96（250%→2.5×）后用户反馈"紧凑都大得不像话"（46.7px 字、145px 行高远超系统 30px）。**定稿：`_dpi_scale = sqrt(dpi/96)` 折中**（250%→1.581×、200%→1.414×、150%→1.225×、100%→1.0 不变）+ `_font_scale = text_pt/FONT_BODY`。240 DPI 下 14/16/18 档渲染约 30/34/38px 物理。③ **切档位后必须 `root.after(80, _apply_view_resize)` 重算页大小**：_rebuild_ui 重建时表格 winfo_height=1，页大小若停留旧行高算出的值，可见行数不随字号变化
- 诊断：_log_diag 写 debug.log [diag] 行；进程管理：PowerShell CIM 杀进程（Bash 的 wmic/tasklist 沙箱拦截）+ debug.log 行数验证
- **已移除的死代码**：StatusPill、_path_options、_remove_from_results、排序残留之外的 _update_search_width
- 设计稿：ui_preview_v2.html（A 霁青精修 / B 墨玉 / C 凝脂 三方向对比，含色板与取舍说明）
