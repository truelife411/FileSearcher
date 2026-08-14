# File Searcher 项目笔记

## 用户偏好
- 每次改完代码自动做本地 `git commit`
- `git push` 到 GitHub 需要用户明确要求
- **改完代码必须杀旧进程再启动新程序**（2026-08-08 用户明确强调）：杀掉所有 file_searcher 实例（`wmic process where "name='pythonw.exe'" get ProcessId,CommandLine` 过滤 → `wmic process where ProcessId=X call terminate`），再用 Python subprocess detached（DETACHED|CREATE_NO_WINDOW）启动 pythonw 新实例，确认 debug.log 新增 [diag] 行
- × 关闭按钮 = 最小化到托盘，最小化按钮也缩到托盘，彻底退出靠托盘右键菜单
- 托盘恢复窗口时：最大化 + 搜索框全选/聚焦

## 技术栈
- **运行用系统 Python 3.12**（`C:\Users\hjf\AppData\Local\Programs\Python\Python312\pythonw.exe`）— managed 3.13 没装 pystray/PIL！开发检查可用 managed 3.13 py_compile
- Tkinter + SQLite (WAL模式)；系统托盘：pystray + Pillow；打包：PyInstaller (--onefile --windowed)
- GitHub: https://github.com/truelife411/FileSearcher

## UI 设计基线（2026-08-14「凝脂纸感」，方向 C，当前线上版）
- **凝脂主题**（B 墨玉实机被否决后改选 C）：宣纸暖白底 #F5F3EE + 黛青 accent #2E6E66（渐变 #3A837A→#255A53）+ 墨色文字 #23282C。深色变体「墨玉」#0A0C10/#6FD8C8 保留在 THEMES["dark"]。默认主题 = light。sel_text = 选中行/激活态文字色
- **纸感语言 = 小圆角 + 细线**：结果卡片 6 / 按钮 8 / 搜索框 6 / 弹窗 10；分隔线 row_line、描边 border 1px；无网格线
- **布局（C 版定型）**：标题栏 40px（渐变 logo+副标题）；搜索区 = 居中搜索框（max 980、高 62、placeholder「搜索全盘文件…」、聚焦 accent 描边、有内容时右侧 ✕ 清空）；下方 chips 类型筛选排（全部/文件夹/文档/图片/视频/音频/压缩包/代码，接入 filters.type）；结果纸面卡片（_layout_result_container 圆角 6）内嵌 FileTable + 底部分页栏（左"共 N 个结果"/右 ‹ 1/12 ›）；**底部快捷键栏 44px**（左：状态点+文字 + kbd 提示 ↵打开/⌥↵定位/^C复制/⌫删除/⇵翻页；右：⟳重建索引 ghost 钮 + ⚙ 钮）
- **呼吸式表格 FileTable**：行分隔线、选中行 selected 底+左侧 6px accent 指示条、类型徽章（BADGE_STYLES[theme][kind]=(fg,bg)）、表头加粗+排序 accent+自绘小三角、列宽拖动（_fit_cols 按序压缩/拉宽、_drag_pref 刚性列、_fitted_width 1:1 手感、layout.json 跨显示器换算）。**路径列用 _path_middle_ellipsis 中间省略**（保留首尾目录）
- **字号层级不变**：MICRO=9/SMALL=10/BODY=12/INPUT=13/LG=14/XL=15 base pt；text_pt 档位 14/16/18；_f/_s 全局缩放；数字列 Consolas
- **全自绘弹窗体系**：_DialogShell（透明角+阴影+置顶钳制）、_dialog_confirm、_dialog_input（分段胶囊）、CtxMenu、ToggleSwitch、RoundEntry
- **设置窗口**：左导航三页（常规/排除列表/关于）；主题卡 = 墨玉(深)/凝脂(浅)/跟随系统，切换 destroy+重开；排除列表 ttk.Treeview（Ex.Treeview）
- **分页**：页大小 = 表头下方可视行数（_compute_page_size），滚轮翻页，resize 重算
- Windows 无边框：overrideredirect + WM_NCHITTEST；启动铺满工作区
- **缩放坑位**：Tk 8.6 按 96 DPI 布局，tk scaling 污染 winfo_fpixels；_dpi_scale=GetDpiForSystem()/96 × _font_scale；_s() 输出恒 ≈ px×2.083（系数互抵），像素尺寸不随 DPI 变
- 诊断：_log_diag 写 debug.log [diag] 行；进程管理：PowerShell CIM 杀进程（Bash 的 wmic/tasklist 沙箱拦截）+ debug.log 行数验证
- **已移除的死代码**：StatusPill、_path_options、_remove_from_results、排序残留之外的 _update_search_width
- 设计稿：ui_preview_v2.html（A 霁青精修 / B 墨玉 / C 凝脂 三方向对比，含色板与取舍说明）
