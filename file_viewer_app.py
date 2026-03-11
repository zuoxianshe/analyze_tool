# -*- coding: utf-8 -*-
from app_common import *
from text_compare_window import TextCompareWindow
from alarm_monitor_window import AlarmMonitorWindow
from tkinter import simpledialog

class FileViewerApp(tkdnd.Tk):
    def __init__(self):
        super().__init__()
        self.title("文件查看器｜文件夹优先+字母序")
        self.geometry("1200x750")
        self.configure(bg="#f4f4f4")

        # 核心变量
        self.current_path = None
        self.max_file_size = 10 * 1024 * 1024  # 10MB
        self.node_data = {}          # 节点数据
        self.node_type = {}          # 节点类型
        self.parent_node_files = {}  # 去重映射
        self.node_full_path = {}     # 节点完整路径
        self.current_text_content = ""  # 当前文本内容
        self.pdf_lazy_source = None
        self.pdf_lazy_name = ""
        self.pdf_page_count = 0
        self.pdf_current_page = 1
        self.pdf_page_cache = {}
        self.pdf_rendering = False
        self.txt_embedded_images = []
        self.editor_font_size = 10
        self.editor_font = tkfont.Font(family="Consolas", size=self.editor_font_size)
        self.editor_bold_font = tkfont.Font(family="Consolas", size=self.editor_font_size, weight="bold")
        # 行号字号与正文一致，避免滚动后行号与正文错位
        self.line_num_font = tkfont.Font(family="Consolas", size=self.editor_font_size)
        
        # 搜索相关变量（分开存储避免冲突）
        self.file_search_hits = []   # 文件名搜索结果 [(节点ID, 显示名称)]
        self.content_search_hits = []# 内容搜索结果 [(行号, 列号, 内容)]
        self.multi_content_search_hits = []  # 多文件内容搜索结果 [(节点ID, 行号, 列号, 行内容)]
        self.preview_line_to_hit_index = []  # 兼容旧逻辑（保留）
        self.preview_jump_entries = []       # 预览框行号到跳转目标映射（统一）
        self.current_search_type = ""  # "file" 或 "content"
        self.search_text_cache = {}    # 多文件搜索文本缓存：node_id -> {"key": ..., "lines": [...]}
        self._pending_multi_jump = None
        self.search_history = []       # 搜索历史（最近20次）
        self.content_search_result_history = []      # 内容搜索结果历史（最近10次）
        self.multi_content_result_history = []       # 多文件搜索结果历史（最近10次）
        self.virtual_text_boxes = []                 # 虚拟多文本框（页签）
        self.active_text_box_id = None
        self._text_box_seq = 1
        self.node_text_box_map = {}                  # 文件节点 -> 文本框页签id

        # ========== 顶部综合工具区（文件操作 + 文字编辑） ==========
        self.frame_tools = tk.Frame(self, bg="#e8eef8", bd=2, relief=tk.RIDGE)
        self.frame_tools.pack(fill="x", padx=5, pady=1)
        tools_inner = tk.Frame(self.frame_tools, bg="#e8eef8")
        tools_inner.pack(fill="x", padx=4, pady=3)
        tools_inner.grid_columnconfigure(0, weight=1)
        tools_inner.grid_columnconfigure(1, weight=0)
        tools_inner.grid_columnconfigure(2, weight=1)

        file_group = tk.Frame(tools_inner, bg="#e8eef8")
        file_group.grid(row=0, column=0, sticky="ew", padx=(2, 8))
        tk.Label(file_group, text="📁 文件操作", bg="#e8eef8", font=("微软雅黑",8,"bold")).pack(anchor="w", pady=(0, 2))
        file_btn_row = tk.Frame(file_group, bg="#e8eef8")
        file_btn_row.pack(anchor="w")
        
        self.folder_archive_btn = tk.Button(file_btn_row, text="📂 文件夹/压缩包", width=14, command=self.open_folder_or_archive,
                                   font=("微软雅黑",9), bg="#2196F3", fg="white")
        self.folder_archive_btn.pack(side="left", padx=3)
        self.file_btn = tk.Button(file_btn_row, text="📄 选择文件", width=14, command=self.open_file,
                                 font=("微软雅黑",9), bg="#4CAF50", fg="white")
        self.file_btn.pack(side="left", padx=3)
        self.compare_btn = tk.Button(file_btn_row, text="🆚 文本比较", width=14, command=self.open_compare_window,
                                    font=("微软雅黑",9), bg="#9C27B0", fg="white")
        self.compare_btn.pack(side="left", padx=3)
        self.monitor_btn = tk.Button(file_btn_row, text="🚨 监控告警", width=12, command=self.open_alarm_monitor_window,
                                     font=("微软雅黑",9), bg="#FF7043", fg="white")
        self.monitor_btn.pack(side="left", padx=3)

        sep = tk.Frame(tools_inner, width=1, bg="#b9c7de")
        sep.grid(row=0, column=1, sticky="ns", padx=6)

        edit_group = tk.Frame(tools_inner, bg="#e8eef8")
        edit_group.grid(row=0, column=2, sticky="ew", padx=(8, 2))
        tk.Label(edit_group, text="✏️ 文字编辑", bg="#e8eef8", font=("微软雅黑",8,"bold")).pack(anchor="w", pady=(0, 2))
        edit_btn_row = tk.Frame(edit_group, bg="#e8eef8")
        edit_btn_row.pack(anchor="w")

        self.bold_btn = tk.Button(edit_btn_row, text="𝐁 加粗", width=7, command=self.set_bold, font=("微软雅黑",8))
        self.bold_btn.pack(side="left", padx=5)
        self.color_btn = tk.Menubutton(edit_btn_row, text="🎨 颜色", width=7, font=("微软雅黑",8), relief=tk.RAISED, bg="#f0f0f0")
        self.color_btn.pack(side="left", padx=5)
        self.color_menu = tk.Menu(self.color_btn, tearoff=0, bg="white", bd=1)
        self.color_btn.config(menu=self.color_menu)
        self.build_word_style_color_menu()
        self.reset_btn = tk.Button(edit_btn_row, text="🔄 恢复默认", width=9, command=self.reset_style, font=("微软雅黑",8))
        self.reset_btn.pack(side="left", padx=5)
        self.md_render_btn = tk.Button(edit_btn_row, text="📝 Markdown渲染", width=12, command=self.render_current_text_as_markdown, font=("微软雅黑",8))
        self.md_render_btn.pack(side="left", padx=5)
        self.save_btn = tk.Button(edit_btn_row, text="保存文本", width=9, command=self.save_current_text, font=("微软雅黑",8))
        self.save_btn.pack(side="left", padx=5)

        # ========== 搜索栏 ==========
        search_frame = tk.Frame(self, bg="#f4f4f4")
        search_frame.pack(fill="x", pady=2, padx=10)
        
        tk.Label(search_frame, text="智能搜索：", bg="#f4f4f4", font=("微软雅黑",9)).pack(side="left", padx=2)
        self.search_entry = ttk.Combobox(search_frame, width=50, font=("Consolas",10), state="normal")
        self.search_entry.pack(side="left", padx=8, fill="x", expand=True)
        self.search_entry.bind("<Return>", lambda e: self.smart_search())
        
        self.search_btn = tk.Button(search_frame, text="🔍 搜索", width=8, command=self.smart_search, bg="#2196F3", fg="white")
        self.search_btn.pack(side="left", padx=4)
        self.search_multi_btn = tk.Button(search_frame, text="📚 多文件", width=8, command=self.search_content_multi, bg="#009688", fg="white")
        self.search_multi_btn.pack(side="left", padx=4)
        self.clear_btn = tk.Button(search_frame, text="🧹 清空", width=8, command=self.clear_all_highlights, bg="#f4f4f4", fg="black")
        self.clear_btn.pack(side="left", padx=4)
        
        self.search_status_label = tk.Label(search_frame, text="当前模式：未搜索", bg="#f4f4f4", fg="gray", font=("微软雅黑",9))
        self.search_status_label.pack(side="left", padx=8)

        # ========== 主布局 ==========
        self.main_paned = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief=tk.RIDGE, sashwidth=6, bg="#f4f4f4")
        self.main_paned.pack(fill="both", expand=True, padx=10, pady=5)

        # 左侧：文件结构树
        self.tree_frame = tk.Frame(self.main_paned, bg="#f4f4f4")
        self.tree_label = tk.Label(self.tree_frame, text="📂 文件结构（文件夹优先+字母序）", font=("微软雅黑",10,"bold"), bg="#e0e0e0")
        self.tree_label.pack(fill="x")
        self.tree = ttk.Treeview(self.tree_frame, show="tree")
        self.tree.pack(fill="both", expand=True, pady=4)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.tag_configure("search_hit", background="yellow")
        self.main_paned.add(self.tree_frame, width=280)

        # 右侧面板（垂直分割）
        self.right_paned = tk.PanedWindow(self.main_paned, orient=tk.VERTICAL, sashrelief=tk.RIDGE, sashwidth=6, bg="#f4f4f4")
        self.main_paned.add(self.right_paned)

        # 右上：文本编辑区
        self.text_panel = tk.Frame(self.right_paned, bg="#f4f4f4")
        self.text_box_selector = tk.Frame(self.text_panel, bg="#eef2f8")
        self.text_box_selector.pack(fill="x", padx=4, pady=(2, 0))
        self.current_file_label = tk.Label(
            self.text_panel, text="当前文件：-", bg="#f4f4f4", fg="#37474f", anchor="w", font=("微软雅黑", 9, "bold")
        )
        self.current_file_label.pack(fill="x", padx=6, pady=(2, 0))
        self.txt_container = tk.Frame(self.text_panel, bg="#f4f4f4")
        self.txt_container.pack(fill="both", expand=True)
        self.pdf_nav_frame = tk.Frame(self.text_panel, bg="#f4f4f4", bd=0, highlightthickness=0, height=34)
        self.pdf_nav_frame.pack_propagate(False)
        self.pdf_nav_label = tk.Label(
            self.pdf_nav_frame, text="PDF分页：0/0", bg="#f4f4f4", font=("微软雅黑", 9, "bold"), bd=0, highlightthickness=0
        )
        self.pdf_nav_label.pack(side="left", padx=6, pady=2)
        self.pdf_page_slider = tk.Scale(
            self.pdf_nav_frame, from_=1, to=1, orient=tk.HORIZONTAL, showvalue=0,
            length=320, command=self.on_pdf_slider_change,
            bg="#f4f4f4", bd=0, highlightthickness=0, troughcolor="#d9d9d9", relief=tk.FLAT
        )
        self.pdf_page_slider.pack(side="left", padx=4, fill="x", expand=True)
        self.pdf_page_slider.config(state=tk.DISABLED)
        self.pdf_nav_frame.place_forget()
        self.line_num = tk.Text(self.txt_container, width=6, state="disabled", bg="#f0f0f0", font=self.line_num_font, wrap=tk.NONE)
        self.line_num.pack(side="left", fill="y")
        self.line_num.configure(cursor="arrow")
        self.line_num.tag_config("ln", justify="right")
        self.txt_ybar = tk.Scrollbar(self.txt_container, orient="vertical")
        self.txt_ybar.pack(side="right", fill="y")
        self.txt = tk.Text(self.txt_container, font=self.editor_font, wrap=tk.NONE, bg="white", undo=True)
        self.txt.pack(side="right", fill="both", expand=True)
        self.txt.configure(yscrollcommand=self.on_text_vertical_scroll)
        self.txt_ybar.config(command=self.txt.yview)
        self.txt_xbar_bar = tk.Frame(self.text_panel, bg="#f4f4f4", height=16)
        self.txt_xbar_bar.pack(fill="x", padx=2, pady=(0, 2))
        self.txt_xbar_bar.pack_propagate(False)
        self.txt_xbar = tk.Scrollbar(self.txt_xbar_bar, orient="horizontal", command=self.txt.xview)
        self.txt_xbar.pack(fill="both", expand=True)
        self.txt.configure(xscrollcommand=self.txt_xbar.set)
        self.pdf_corner_label = tk.Label(
            self.txt_container, text="", bg="#eef3ff", fg="#1f3a93", font=("微软雅黑", 9, "bold"),
            bd=1, relief=tk.SOLID, padx=6, pady=2
        )
        self.pdf_corner_label.place_forget()
        self.txt.bind("<KeyRelease>", lambda e: self.on_text_edited())
        self.txt.bind("<Control-s>", self.save_current_text)
        self.txt.bind("<Control-S>", self.save_current_text)
        self.txt.bind("<MouseWheel>", self.on_pdf_mousewheel)
        self.txt.bind("<Button-4>", self.on_pdf_mousewheel)
        self.txt.bind("<Button-5>", self.on_pdf_mousewheel)
        self.txt.bind("<Shift-MouseWheel>", self.on_text_horizontal_scroll)
        self.txt.bind("<Shift-Button-4>", self.on_text_horizontal_scroll)
        self.txt.bind("<Shift-Button-5>", self.on_text_horizontal_scroll)
        self.txt.bind("<Control-MouseWheel>", self.on_editor_zoom)
        self.txt.bind("<Control-Button-4>", self.on_editor_zoom)
        self.txt.bind("<Control-Button-5>", self.on_editor_zoom)
        self.line_num.bind("<MouseWheel>", self.on_line_num_mousewheel)
        self.line_num.bind("<Button-4>", self.on_line_num_mousewheel)
        self.line_num.bind("<Button-5>", self.on_line_num_mousewheel)
        self.text_panel.bind("<Configure>", lambda e: self.on_text_panel_resize())
        # 文本标签配置
        self.txt.tag_config("hl", background="yellow")
        self.txt.tag_config("jump_hl", background="#FFE4B5")
        self.txt.tag_config("bold_tag", font=self.editor_bold_font)
        self.txt.tag_config("color_tag", foreground="black")
        self.txt.tag_config("md_h1", font=("微软雅黑", 18, "bold"), foreground="#1b4f9c", spacing1=10, spacing3=6)
        self.txt.tag_config("md_h2", font=("微软雅黑", 16, "bold"), foreground="#1f5fae", spacing1=8, spacing3=5)
        self.txt.tag_config("md_h3", font=("微软雅黑", 14, "bold"), foreground="#2d6bbf", spacing1=6, spacing3=4)
        self.txt.tag_config("md_quote", foreground="#5f6b7a", lmargin1=18, lmargin2=18)
        self.txt.tag_config("md_codeblock", font=("Consolas", 10), background="#f5f7fa", lmargin1=14, lmargin2=14)
        self.txt.tag_config("md_inline_code", font=("Consolas", 10), background="#eef2f7")
        self.txt.tag_config("md_bold", font=("Consolas", self.editor_font_size, "bold"))
        self.txt.tag_config("md_italic", font=("Consolas", self.editor_font_size, "italic"))
        self.txt.tag_config("md_list", lmargin1=20, lmargin2=30)
        self.txt.tag_config("md_link", foreground="#1a73e8", underline=True)
        self.right_paned.add(self.text_panel, height=430, minsize=230)

        # 右下：搜索结果预览
        self.result_frame = tk.Frame(self.right_paned, bg="#f0f8f8")
        self.result_title = tk.Label(self.result_frame, text="🔍 搜索结果预览 | 匹配数：0", bg="#f0f8f8",
                                    font=("微软雅黑",9,"bold"), fg="#2196F3")
        self.result_title.pack(anchor="w", padx=5, pady=2)
        self.result_txt_container = tk.Frame(self.result_frame, bg="#f0f8f8")
        self.result_txt_container.pack(fill="both", expand=True, padx=5, pady=(0, 0))
        self.result_txt_ybar = tk.Scrollbar(self.result_txt_container, orient="vertical")
        self.result_txt_ybar.pack(side="right", fill="y")
        self.result_txt = tk.Text(self.result_txt_container, font=("Consolas",10), bg="#fffff8", height=1, wrap=tk.NONE)
        self.result_txt.pack(side="left", fill="both", expand=True)
        self.result_txt.configure(yscrollcommand=self.result_txt_ybar.set)
        self.result_txt_ybar.config(command=self.result_txt.yview)
        self.result_txt_xbar_bar = tk.Frame(self.result_frame, bg="#f0f8f8", height=12)
        self.result_txt_xbar_bar.pack(fill="x", padx=5, pady=(0, 0))
        self.result_txt_xbar_bar.pack_propagate(False)
        self.result_txt_xbar = tk.Scrollbar(self.result_txt_xbar_bar, orient="horizontal", command=self.result_txt.xview)
        self.result_txt_xbar.pack(fill="both", expand=True)
        self.result_txt.configure(xscrollcommand=self.result_txt_xbar.set)
        self.result_txt.config(state=tk.DISABLED)
        # 双击跳转绑定
        self.result_txt.bind("<Double-1>", self.on_double_click_jump)
        self.right_paned.add(self.result_frame, height=58, minsize=24)
        self.result_collapsed_height = 24
        self.result_expanded_height = 120

        # 事件绑定
        self.txt.bind("<FocusIn>", lambda e: self.set_search_mode("content"))
        self.tree.bind("<FocusIn>", lambda e: self.set_search_mode("file"))
        self.search_entry.bind("<FocusIn>", lambda e: self.set_search_mode("file"))
        
        # 拖拽支持：仅允许拖到“文件结构区”或“文本框”
        self.tree_frame.drop_target_register(tkdnd.DND_FILES)
        self.tree_frame.dnd_bind('<<Drop>>', self.on_drop)
        self.txt.drop_target_register(tkdnd.DND_FILES)
        self.txt.dnd_bind('<<Drop>>', self.on_drop)
        self.bind_all("<Control-f>", self.focus_search_entry)
        self.bind_all("<Control-F>", self.focus_search_entry)
        self._init_virtual_text_boxes()

    # ==============================
    # 基础工具函数
    # ==============================
    def is_dark_color(self, hex_color):
        try:
            hex_color = hex_color.lstrip('#')
            if len(hex_color) == 3:
                hex_color = ''.join([c*2 for c in hex_color])
            if len(hex_color) != 6:
                return False
            r, g, b = int(hex_color[0:2],16), int(hex_color[2:4],16), int(hex_color[4:6],16)
            return (0.299*r + 0.587*g + 0.114*b)/255 < 0.5
        except:
            return False

    def build_word_style_color_menu(self):
        colors = [
            ("自动/黑色", "#000000"), ("白色", "#FFFFFF"), ("红色", "#FF0000"),
            ("橙色", "#FF6600"), ("黄色", "#FFFF00"), ("绿色", "#00B050"),
            ("蓝色", "#0070C0"), ("紫色", "#7030A0"), ("深红", "#C00000"),
            ("深绿", "#007030"), ("深蓝", "#002060"), ("浅红", "#FFC0CB"),
            ("浅绿", "#92D050"), ("浅蓝", "#B7DEE8")
        ]
        for name, hex_color in colors:
            fg = "white" if self.is_dark_color(hex_color) else "black"
            self.color_menu.add_command(
                label=name, background=hex_color, foreground=fg,
                command=lambda c=hex_color: self.apply_selected_color(c)
            )

    def apply_selected_color(self, color_hex):
        try:
            start = self.txt.index("sel.first")
            end = self.txt.index("sel.last")
            self.txt.tag_config("color_tag", foreground=color_hex)
            self.txt.tag_add("color_tag", start, end)
        except tk.TclError:
            messagebox.showinfo("提示", "请先选中要修改颜色的文字")

    def set_bold(self):
        try:
            start = self.txt.index("sel.first")
            end = self.txt.index("sel.last")
            if "bold_tag" in self.txt.tag_names(start):
                self.txt.tag_remove("bold_tag", start, end)
            else:
                self.txt.tag_add("bold_tag", start, end)
        except tk.TclError:
            messagebox.showinfo("提示", "请先选中要加粗的文字")

    def reset_style(self):
        self.txt.tag_remove("bold_tag", "1.0", tk.END)
        self.txt.tag_remove("color_tag", "1.0", tk.END)
        self.txt.tag_remove("jump_hl", "1.0", tk.END)
        self.txt.tag_config("color_tag", foreground="black")
        messagebox.showinfo("提示", "已恢复所有文字默认样式")

    def set_search_mode(self, mode):
        """设置搜索模式：file（文件名）/ content（内容）"""
        self.current_search_type = mode
        mode_text = "文件名搜索" if mode == "file" else "内容搜索"
        self.search_status_label.config(text=f"当前模式：{mode_text}")

    def focus_search_entry(self, event=None):
        """Ctrl+F: 聚焦智能搜索框并选中现有内容"""
        self.search_entry.focus_set()
        text = self.search_entry.get()
        if text:
            self.search_entry.selection_range(0, tk.END)
            self.search_entry.icursor(tk.END)
        return "break"

    def set_current_file_label(self, file_name):
        name = str(file_name).strip() if file_name else "-"
        self.current_file_label.config(text=f"当前文件：{name}")
        if self.active_text_box_id:
            box = self._get_active_text_box()
            if box is not None and box.get("auto_title", True):
                box["title"] = name
                self._refresh_text_box_selector()

    def _init_virtual_text_boxes(self):
        self.virtual_text_boxes = [{"id": "tb_1", "title": "主文本框", "content": "", "auto_title": True}]
        self.active_text_box_id = "tb_1"
        self._text_box_seq = 2
        self._refresh_text_box_selector()

    def _create_text_box(self, title, content="", auto_title=True, switch_to=True):
        new_id = f"tb_{self._text_box_seq}"
        self._text_box_seq += 1
        self.virtual_text_boxes.append({
            "id": new_id,
            "title": str(title or f"文本框{self._text_box_seq}"),
            "content": str(content or ""),
            "auto_title": bool(auto_title),
        })
        if switch_to:
            self.switch_text_box(new_id)
        else:
            self._refresh_text_box_selector()
        return new_id

    def _ensure_text_box_for_node(self, node_id, node_title):
        box_id = self.node_text_box_map.get(node_id)
        if box_id:
            exists = any(b.get("id") == box_id for b in self.virtual_text_boxes)
            if exists:
                self.switch_text_box(box_id)
                return box_id
        # 首次解析该文件节点：自动新建文本框页签
        new_id = self._create_text_box(node_title, content="", auto_title=True, switch_to=True)
        self.node_text_box_map[node_id] = new_id
        return new_id

    def _get_active_text_box(self):
        for b in self.virtual_text_boxes:
            if b.get("id") == self.active_text_box_id:
                return b
        return None

    def _sync_active_text_box_content(self):
        box = self._get_active_text_box()
        if box is None:
            return
        box["content"] = self.txt.get("1.0", "end-1c")

    def _refresh_text_box_selector(self):
        for w in self.text_box_selector.winfo_children():
            w.destroy()
        for idx, b in enumerate(self.virtual_text_boxes, start=1):
            bid = b.get("id")
            title = b.get("title") or f"文本框{idx}"
            label = f"{idx}:{title}"
            is_active = bid == self.active_text_box_id
            tab = tk.Frame(
                self.text_box_selector,
                bg="#dfe8f6" if is_active else "#f4f7fb",
                bd=1,
                relief=tk.SUNKEN if is_active else tk.RAISED
            )
            tab.pack(side="left", padx=2, pady=2)

            title_btn = tk.Button(
                tab,
                text=label,
                relief=tk.FLAT,
                bd=0,
                font=("微软雅黑", 8),
                bg="#dfe8f6" if is_active else "#f4f7fb",
                activebackground="#dfe8f6",
                command=lambda x=bid: self.switch_text_box(x)
            )
            title_btn.pack(side="left", padx=(4, 1), pady=1)
            title_btn.bind("<Double-1>", lambda e, x=bid: self.rename_text_box(x))
            title_btn.bind("<Button-3>", lambda e, x=bid: self.rename_text_box(x))

            close_btn = tk.Button(
                tab,
                text="x",
                width=2,
                relief=tk.FLAT,
                bd=0,
                font=("Consolas", 8, "bold"),
                fg="#b71c1c",
                bg="#dfe8f6" if is_active else "#f4f7fb",
                activebackground="#ffcdd2",
                command=lambda x=bid: self.close_text_box(x)
            )
            close_btn.pack(side="left", padx=(1, 3), pady=1)

    def rename_text_box(self, box_id):
        box = None
        for b in self.virtual_text_boxes:
            if b.get("id") == box_id:
                box = b
                break
        if box is None:
            return
        current_title = box.get("title") or "文本框"
        new_title = simpledialog.askstring("重命名页签", "请输入新页签名称：", initialvalue=current_title, parent=self)
        if new_title is None:
            return
        new_title = str(new_title).strip()
        if not new_title:
            return
        box["title"] = new_title
        box["auto_title"] = False
        if box_id == self.active_text_box_id:
            self.current_file_label.config(text=f"当前文件：{new_title}")
        self._refresh_text_box_selector()

    def close_text_box(self, box_id):
        if len(self.virtual_text_boxes) <= 1:
            messagebox.showinfo("提示", "至少保留一个文本框页签")
            return
        self._sync_active_text_box_content()
        idx = -1
        for i, b in enumerate(self.virtual_text_boxes):
            if b.get("id") == box_id:
                idx = i
                break
        if idx < 0:
            return
        was_active = (box_id == self.active_text_box_id)
        del self.virtual_text_boxes[idx]
        self.node_text_box_map = {k: v for k, v in self.node_text_box_map.items() if v != box_id}
        if was_active:
            new_idx = max(0, idx - 1)
            self.active_text_box_id = self.virtual_text_boxes[new_idx].get("id")
            target = self.virtual_text_boxes[new_idx]
            content = target.get("content", "")
            self.txt.delete("1.0", tk.END)
            self.txt.insert("1.0", content)
            self.current_text_content = content
            self.refresh_line_numbers()
            self.current_file_label.config(text=f"当前文件：{target.get('title', '-')}")
        self._refresh_text_box_selector()

    def switch_text_box(self, box_id):
        if box_id == self.active_text_box_id:
            return
        self._sync_active_text_box_content()
        target = None
        for b in self.virtual_text_boxes:
            if b.get("id") == box_id:
                target = b
                break
        if target is None:
            return
        self.active_text_box_id = box_id
        content = target.get("content", "")
        self.txt.delete("1.0", tk.END)
        self.txt.insert("1.0", content)
        self.current_text_content = content
        self.refresh_line_numbers()
        self.current_file_label.config(text=f"当前文件：{target.get('title', '-')}")
        self._refresh_text_box_selector()

    def record_search_history(self, keyword):
        kw = str(keyword or "").strip()
        if not kw:
            return
        # 去重：已有关键词移除后再追加，保证最新在末尾
        self.search_history = [x for x in self.search_history if x != kw]
        self.search_history.append(kw)
        if len(self.search_history) > 20:
            self.search_history = self.search_history[-20:]
        self.search_entry["values"] = list(reversed(self.search_history))

    def set_search_keyword(self, keyword):
        self.search_entry.delete(0, tk.END)
        self.search_entry.insert(0, str(keyword))
        self.search_entry.focus_set()

    def _upsert_result_history(self, mode, keyword, hits):
        kw = str(keyword or "").strip()
        if not kw:
            return
        history = self.content_search_result_history if mode == "content" else self.multi_content_result_history
        history[:] = [h for h in history if h.get("keyword") != kw]
        history.append({"keyword": kw, "hits": list(hits), "collapsed": False})
        if len(history) > 10:
            del history[:-10]

    def _toggle_result_history(self, mode, keyword):
        history = self.content_search_result_history if mode == "content" else self.multi_content_result_history
        for item in history:
            if item.get("keyword") == keyword:
                item["collapsed"] = not item.get("collapsed", False)
                break

    def _render_result_history(self, mode):
        history = self.content_search_result_history if mode == "content" else self.multi_content_result_history
        self.preview_line_to_hit_index.clear()
        self.preview_jump_entries.clear()
        self.result_txt.config(state=tk.NORMAL)
        self.result_txt.delete("1.0", tk.END)

        if not history:
            title = "内容搜索结果历史" if mode == "content" else "多文件搜索结果历史"
            self.result_txt.insert("1.0", "暂无历史搜索结果")
            self.result_title.config(text=f"🔍 {title} | 记录数：0")
            self.result_txt.config(state=tk.DISABLED)
            self.set_result_panel_collapsed(True)
            return

        lines = []
        for record in reversed(history):
            kw = record.get("keyword", "")
            hits = record.get("hits", [])
            collapsed = bool(record.get("collapsed", False))
            marker = "▶" if collapsed else "▼"
            header = f"{marker} 关键词：{kw} | 命中数：{len(hits)}"
            lines.append(header)
            self.preview_jump_entries.append(("history_toggle", mode, kw))
            if collapsed:
                continue

            max_show = 500
            show_hits = hits[:max_show]
            if mode == "content":
                for row, col, content in show_hits:
                    lines.append(f"  第{row}行: {content}")
                    self.preview_jump_entries.append(("content", int(row), int(col)))
            else:
                for nid, row, col, line in show_hits:
                    name = self.tree.item(nid, "text")
                    lines.append(f"  [{name}] 第{row}行: {line}")
                    self.preview_jump_entries.append(("content_multi", nid, int(row), int(col)))
            if len(hits) > max_show:
                lines.append(f"  ... 共 {len(hits)} 条，仅显示前 {max_show} 条")
                self.preview_jump_entries.append(None)

        self.result_txt.insert("1.0", "\n".join(lines))
        title = "内容搜索结果历史" if mode == "content" else "多文件搜索结果历史"
        self.result_title.config(text=f"🔍 {title} | 记录数：{len(history)}（最近10次）")
        self.result_txt.config(state=tk.DISABLED)
        self.set_result_panel_collapsed(False)

    def set_result_panel_collapsed(self, collapsed):
        h = self.result_collapsed_height if collapsed else self.result_expanded_height
        try:
            self.right_paned.paneconfigure(self.result_frame, height=h)
        except Exception:
            pass

    def refresh_line_numbers(self):
        """刷新行号"""
        self.line_num.config(state=tk.NORMAL)
        self.line_num.delete("1.0", tk.END)
        try:
            line_count = int(self.txt.index("end-1c").split(".")[0])
            self.line_num.insert("end", "\n".join(str(i) for i in range(1, line_count+1)), "ln")
        except:
            pass
        self.line_num.config(state=tk.DISABLED)
        try:
            first, _ = self.txt.yview()
            self.line_num.yview_moveto(first)
        except Exception:
            pass

    def on_text_vertical_scroll(self, first, last):
        self.txt_ybar.set(first, last)
        try:
            self.line_num.yview_moveto(float(first))
        except Exception:
            pass

    def on_line_num_mousewheel(self, event):
        handled = self.on_pdf_mousewheel(event)
        if handled == "break":
            return "break"
        if hasattr(event, "num") and event.num in (4, 5):
            step = -3 if event.num == 4 else 3
        else:
            step = -3 if getattr(event, "delta", 0) > 0 else 3
        self.txt.yview_scroll(step, "units")
        return "break"

    def on_editor_zoom(self, event):
        if hasattr(event, "num") and event.num in (4, 5):
            delta = 1 if event.num == 4 else -1
        else:
            delta = 1 if event.delta > 0 else -1
        new_size = max(8, min(40, self.editor_font_size + delta))
        if new_size == self.editor_font_size:
            return "break"
        self.editor_font_size = new_size
        self.editor_font.configure(size=new_size)
        self.editor_bold_font.configure(size=new_size)
        self.line_num_font.configure(size=new_size)
        self.txt.tag_config("md_bold", font=("Consolas", new_size, "bold"))
        self.txt.tag_config("md_italic", font=("Consolas", new_size, "italic"))
        self.refresh_line_numbers()
        return "break"

    def render_current_text_as_markdown(self):
        """Render current editor content as markdown regardless of file extension."""
        content = self.txt.get("1.0", "end-1c")
        self.display_text_in_editor(content, file_ext=".md")

    def set_markdown_render_button_state(self, rendered):
        if not hasattr(self, "md_render_btn"):
            return
        if rendered:
            self.md_render_btn.config(text="✅ 已渲染", relief=tk.SUNKEN)
        else:
            self.md_render_btn.config(text="📝 Markdown渲染", relief=tk.RAISED)

    def display_text_in_editor(self, content, file_ext=""):
        self.txt_embedded_images = []
        self.txt.delete("1.0", tk.END)
        if str(file_ext).lower() == ".md":
            self.render_markdown(content)
            self.set_markdown_render_button_state(True)
        else:
            self.txt.insert("1.0", content)
            self.set_markdown_render_button_state(False)
        self.current_text_content = self.txt.get("1.0", "end-1c")
        self._sync_active_text_box_content()
        self.refresh_line_numbers()

    def _resolve_jump_target(self, expect_row, expect_col, keyword):
        """Ensure jump lands on a real hit line; fallback to nearest matched line."""
        kw = str(keyword or "").strip().lower()
        row = max(1, int(expect_row))
        col = max(0, int(expect_col))
        max_row = int(self.txt.index("end-1c").split(".")[0])
        row = min(row, max_row)
        line_text = self.txt.get(f"{row}.0", f"{row}.end")
        if kw and kw in line_text.lower():
            found_col = line_text.lower().find(kw)
            if found_col >= 0:
                return row, found_col
            return row, col

        if not kw:
            return row, col

        candidate_rows = []
        for r in range(1, max_row + 1):
            t = self.txt.get(f"{r}.0", f"{r}.end")
            if kw in t.lower():
                candidate_rows.append(r)
        if not candidate_rows:
            return row, col
        best_row = min(candidate_rows, key=lambda r: abs(r - row))
        t = self.txt.get(f"{best_row}.0", f"{best_row}.end").lower()
        best_col = t.find(kw)
        if best_col < 0:
            best_col = 0
        return best_row, best_col

    def _apply_jump_to_editor(self, row, col, keyword):
        row, col = self._resolve_jump_target(row, col, keyword)
        target_pos = f"{row}.{max(0, col)}"
        self.txt.tag_remove("jump_hl", "1.0", tk.END)
        self.txt.tag_remove("hl", "1.0", tk.END)
        self.txt.tag_add("jump_hl", f"{row}.0", f"{row}.end")
        if col >= 0:
            end_col = col + max(1, len(str(keyword or "").strip()))
            self.txt.tag_add("hl", f"{row}.{col}", f"{row}.{end_col}")
        self.txt.mark_set(tk.INSERT, target_pos)
        self.txt.see(target_pos)
        self.txt.focus_set()

    def _insert_markdown_inline(self, text):
        i = 0
        n = len(text)
        while i < n:
            if text.startswith("**", i):
                j = text.find("**", i + 2)
                if j != -1:
                    self.txt.insert(tk.END, text[i + 2:j], ("md_bold",))
                    i = j + 2
                    continue
            if text.startswith("*", i):
                j = text.find("*", i + 1)
                if j != -1:
                    self.txt.insert(tk.END, text[i + 1:j], ("md_italic",))
                    i = j + 1
                    continue
            if text.startswith("`", i):
                j = text.find("`", i + 1)
                if j != -1:
                    self.txt.insert(tk.END, text[i + 1:j], ("md_inline_code",))
                    i = j + 1
                    continue
            if text.startswith("[", i):
                m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", text[i:])
                if m:
                    self.txt.insert(tk.END, m.group(1), ("md_link",))
                    i += len(m.group(0))
                    continue
            self.txt.insert(tk.END, text[i])
            i += 1

    def render_markdown(self, content):
        lines = str(content).splitlines()
        in_code = False
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("```"):
                in_code = not in_code
                self.txt.insert(tk.END, "\n")
                continue
            if in_code:
                self.txt.insert(tk.END, line + "\n", ("md_codeblock",))
                continue

            if line.startswith("# "):
                self.txt.insert(tk.END, line[2:].strip() + "\n", ("md_h1",))
                continue
            if line.startswith("## "):
                self.txt.insert(tk.END, line[3:].strip() + "\n", ("md_h2",))
                continue
            if line.startswith("### "):
                self.txt.insert(tk.END, line[4:].strip() + "\n", ("md_h3",))
                continue
            if line.startswith(">"):
                self.txt.insert(tk.END, line[1:].lstrip() + "\n", ("md_quote",))
                continue
            if re.match(r"^\s*[-*]\s+", line) or re.match(r"^\s*\d+\.\s+", line):
                item = re.sub(r"^\s*([-*]|\d+\.)\s+", "• ", line)
                start = self.txt.index(tk.END)
                self._insert_markdown_inline(item)
                end = self.txt.index(tk.END)
                self.txt.tag_add("md_list", start, end)
                self.txt.insert(tk.END, "\n")
                continue

            self._insert_markdown_inline(line)
            self.txt.insert(tk.END, "\n")

    def display_image_in_editor(self, image_obj, title=""):
        self.txt_embedded_images = []
        self.txt.delete("1.0", tk.END)
        if title:
            self.txt.insert("1.0", f"{title}\n\n")
        if not PIL_AVAILABLE:
            self.txt.insert(tk.END, "(未安装Pillow，无法显示图片。请安装：pip install pillow)")
            self.current_text_content = self.txt.get("1.0", "end-1c")
            self.refresh_line_numbers()
            return
        try:
            max_w = max(500, self.txt.winfo_width() - 40)
            img = image_obj.copy()
            if img.width > max_w:
                ratio = max_w / float(img.width)
                img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.LANCZOS)
            tk_img = ImageTk.PhotoImage(img)
            self.txt.image_create(tk.END, image=tk_img)
            self.txt_embedded_images.append(tk_img)
        except Exception as e:
            self.txt.insert(tk.END, f"(图片显示失败: {str(e)})")
        self.current_text_content = self.txt.get("1.0", "end-1c")
        self.refresh_line_numbers()

    def clear_pdf_lazy_state(self):
        self.pdf_lazy_source = None
        self.pdf_lazy_name = ""
        self.pdf_page_count = 0
        self.pdf_current_page = 1
        self.pdf_page_cache.clear()
        self.pdf_rendering = False
        self.pdf_nav_label.config(text="")
        self.pdf_page_slider.config(state=tk.DISABLED, from_=1, to=1)
        self.pdf_nav_frame.place_forget()
        self.pdf_corner_label.place_forget()

    def update_pdf_corner_label(self, page_no=None):
        if not self.pdf_lazy_source or self.pdf_page_count <= 0:
            self.pdf_corner_label.place_forget()
            return
        page = page_no if page_no is not None else self.pdf_current_page
        self.pdf_corner_label.config(text=f"第 {page} / {self.pdf_page_count} 页")
        self.pdf_corner_label.place(relx=1.0, rely=1.0, x=-10, y=-10, anchor="se")

    def setup_pdf_lazy(self, source, display_name):
        self.pdf_lazy_source = source
        self.pdf_lazy_name = display_name
        self.set_current_file_label(display_name)
        self.pdf_page_cache.clear()
        self.pdf_page_count = get_pdf_page_count(
            file_path=source.get("path"),
            file_bytes=source.get("bytes"),
        )
        self.pdf_current_page = 1
        if self.pdf_page_count <= 0:
            self.txt.delete("1.0", tk.END)
            self.txt.insert("1.0", "(PDF页数读取失败)")
            self.refresh_line_numbers()
            self.pdf_nav_label.config(text="")
            self.pdf_page_slider.config(state=tk.DISABLED, from_=1, to=1)
            self.pdf_nav_frame.place_forget()
            return
        self.pdf_page_slider.config(state=tk.NORMAL, from_=1, to=self.pdf_page_count)
        self.pdf_page_slider.set(1)
        self.pdf_nav_label.config(text=f"PDF分页：1/{self.pdf_page_count}")
        self.pdf_nav_frame.place(relx=0.5, rely=1.0, y=-6, anchor="s")
        self.update_pdf_corner_label(1)
        self.render_pdf_page(1)

    def render_pdf_page(self, page_no):
        if not self.pdf_lazy_source or self.pdf_rendering:
            return
        page_no = max(1, min(int(page_no), self.pdf_page_count))
        self.pdf_current_page = page_no
        self.pdf_nav_label.config(text=f"PDF分页：{page_no}/{self.pdf_page_count}（解析中）")
        if page_no in self.pdf_page_cache:
            page_payload = self.pdf_page_cache[page_no]
            self._finish_render_pdf(page_no, page_payload.get("text", ""), page_payload.get("image_bytes"))
            self.pdf_nav_label.config(text=f"PDF分页：{page_no}/{self.pdf_page_count}")
            self.update_pdf_corner_label(page_no)
            return

        self.pdf_rendering = True

        def worker():
            content, _ = extract_pdf_single_page(
                file_path=self.pdf_lazy_source.get("path"),
                file_bytes=self.pdf_lazy_source.get("bytes"),
                page_no=page_no,
                display_name=self.pdf_lazy_name
            )
            image_bytes = extract_pdf_single_page_image_bytes(
                file_path=self.pdf_lazy_source.get("path"),
                file_bytes=self.pdf_lazy_source.get("bytes"),
                page_no=page_no
            )
            payload = {"text": content, "image_bytes": image_bytes}
            self.pdf_page_cache[page_no] = payload
            self.after(0, lambda: self._finish_render_pdf(page_no, content, image_bytes))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_render_pdf(self, page_no, content, image_bytes=None):
        self.pdf_rendering = False
        if not self.pdf_lazy_source:
            return
        self.txt_embedded_images = []
        self.txt.delete("1.0", tk.END)
        self.txt.insert("1.0", content)
        if image_bytes and PIL_AVAILABLE:
            try:
                self.txt.insert(tk.END, "\n\n")
                pil_img = Image.open(io.BytesIO(image_bytes))
                max_w = max(500, self.txt.winfo_width() - 40)
                if pil_img.width > max_w:
                    ratio = max_w / float(pil_img.width)
                    pil_img = pil_img.resize((int(pil_img.width * ratio), int(pil_img.height * ratio)), Image.LANCZOS)
                tk_img = ImageTk.PhotoImage(pil_img)
                self.txt.image_create(tk.END, image=tk_img)
                self.txt_embedded_images.append(tk_img)
            except Exception:
                pass
        self.current_text_content = self.txt.get("1.0", "end-1c")
        self.refresh_line_numbers()
        self.pdf_nav_label.config(text=f"PDF分页：{page_no}/{self.pdf_page_count}")
        self.update_pdf_corner_label(page_no)

    def on_pdf_slider_change(self, value):
        if not self.pdf_lazy_source:
            return
        try:
            page = int(float(value))
        except Exception:
            page = self.pdf_current_page
        if page != self.pdf_current_page:
            self.render_pdf_page(page)

    def on_pdf_mousewheel(self, event):
        if not self.pdf_lazy_source or self.pdf_page_count <= 1:
            return
        if self.pdf_rendering:
            return "break"
        if hasattr(event, "num") and event.num in (4, 5):
            delta = -1 if event.num == 4 else 1
        else:
            delta = -1 if event.delta > 0 else 1
        target = max(1, min(self.pdf_current_page + delta, self.pdf_page_count))
        if target != self.pdf_current_page:
            self.pdf_page_slider.set(target)
            self.render_pdf_page(target)
            return "break"

    def on_text_horizontal_scroll(self, event):
        if hasattr(event, "num") and event.num in (4, 5):
            step = -4 if event.num == 4 else 4
        else:
            step = -4 if event.delta > 0 else 4
        self.txt.xview_scroll(step, "units")
        return "break"

    def on_text_panel_resize(self):
        if self.pdf_lazy_source:
            self.pdf_nav_frame.place(relx=0.5, rely=1.0, y=-6, anchor="s")

    def on_text_edited(self):
        self.current_text_content = self.txt.get("1.0", "end-1c")
        self._sync_active_text_box_content()
        self.refresh_line_numbers()

    def save_current_text(self, event=None):
        content = self.txt.get("1.0", "end-1c")
        selection = self.tree.selection()
        current_node = selection[0] if selection else None
        current_path = self.node_full_path.get(current_node, "") if current_node else ""

        initial_name = os.path.basename(current_path) if current_path else "smart_helper_text.txt"
        defaultextension = os.path.splitext(initial_name)[1] or ".txt"
        save_path = filedialog.asksaveasfilename(
            title="保存文本内容",
            initialfile=initial_name,
            defaultextension=defaultextension,
            filetypes=[("文本文件", "*.txt"), ("JSON", "*.json"), ("CSV", "*.csv"), ("Python", "*.py"), ("所有文件", "*.*")]
        )
        if not save_path:
            return "break" if event else None

        try:
            with open(save_path, "w", encoding="utf-8", errors="replace") as f:
                f.write(content)
            self.current_text_content = content
            messagebox.showinfo("提示", f"文本已保存：\n{save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")
        return "break" if event else None

    # ==============================
    # 节点管理函数
    # ==============================
    def _check_node_exists(self, parent_node, name):
        """检查节点是否已存在"""
        if parent_node not in self.parent_node_files:
            self.parent_node_files[parent_node] = {}
        return name in self.parent_node_files[parent_node]

    def _add_node(self, parent_node, name, node_id, node_type, node_data, node_full_path):
        """添加节点（去重）"""
        pure_name = os.path.basename(name) if name else "未知名称"
        if parent_node not in self.parent_node_files:
            self.parent_node_files[parent_node] = {}
        
        # 仅在不存在时创建节点
        if pure_name not in self.parent_node_files[parent_node]:
            self.tree.insert(parent_node, "end", text=pure_name, iid=node_id)
            self.parent_node_files[parent_node][pure_name] = node_id
        
        # 更新节点数据
        self.node_type[node_id] = node_type
        self.node_data[node_id] = node_data
        self.node_full_path[node_id] = node_full_path
        
        return self.parent_node_files[parent_node][pure_name]

    def clear_all(self):
        """清空所有数据"""
        self.clear_pdf_lazy_state()
        # 清空树
        self.tree.delete(*self.tree.get_children())
        # 清空映射表
        self.node_data.clear()
        self.node_type.clear()
        self.parent_node_files.clear()
        self.node_full_path.clear()
        # 清空文本
        self.txt.delete("1.0", tk.END)
        # 清空搜索结果
        self.clear_all_highlights()
        # 重置变量
        self.current_text_content = ""
        self.current_path = None
        self.file_search_hits.clear()
        self.content_search_hits.clear()
        self.current_search_type = ""
        self.node_text_box_map.clear()
        self._init_virtual_text_boxes()
        self._pending_multi_jump = None
        self.set_current_file_label("")

    # ==============================
    # 文件/文件夹加载函数
    # ==============================
    def open_folder_or_archive(self):
        """打开文件夹或压缩包"""
        folder_path = filedialog.askdirectory(title="选择文件夹（若取消则继续选择压缩包）")
        if folder_path:
            folder_path = normalize_input_path(folder_path)
            self.parent_node_files.clear()
            self.load_folder(folder_path)
            return

        archive_path = filedialog.askopenfilename(
            title="选择压缩包",
            filetypes=[
                ("压缩包", "*.zip *.tar *.gz *.bz2 *.tar.gz *.tar.bz2 *.tgz *.tbz2"),
                ("所有文件", "*.*")
            ]
        )
        if archive_path:
            archive_path = normalize_input_path(archive_path)
            self.parent_node_files.clear()
            self.load_file_or_archive(archive_path)

    def open_file(self):
        """打开普通文件"""
        file_path = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[
                ("支持的文件", "*.txt *.py *.md *.json *.xml *.csv *.xlsx *.xls *.html *.log *.pdf *.doc *.docx"),
                ("文本文件", "*.txt *.py *.md *.json *.xml *.html *.log *.pdf *.doc *.docx"),
                ("表格文件", "*.csv *.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            file_path = normalize_input_path(file_path)
            self.parent_node_files.clear()
            self.load_file_or_archive(file_path)

    def load_folder(self, folder_path):
        """加载文件夹结构"""
        self.clear_all()
        self.current_path = folder_path
        
        # 创建根节点
        root_name = os.path.basename(folder_path)
        root_node = f"root_{uuid.uuid4()}"
        self._add_node("", root_name, root_node, "local_dir", folder_path, folder_path)
        self.tree.item(root_node, open=True)
        
        # 递归遍历
        self._traverse_folder(folder_path, root_node, 1)
        
        self.title(f"文件查看器 - {root_name}")
        self.set_current_file_label(root_name)
        self.txt.insert("1.0", f"✅ 已加载文件夹：{folder_path}\n排序规则：文件夹优先 + 字母序（不区分大小写）")
        self.refresh_line_numbers()

    def _traverse_folder(self, folder_path, parent_node, level):
        """递归遍历文件夹"""
        try:
            entries = os.listdir(folder_path)
            
            # 分离文件夹和文件，分别排序
            dirs = []
            files = []
            for entry in entries:
                entry_path = os.path.join(folder_path, entry)
                if os.path.isdir(entry_path):
                    dirs.append(entry)
                else:
                    files.append(entry)
            
            # 按字母序排序（不区分大小写）
            dirs_sorted = sorted(dirs, key=lambda x: x.lower())
            files_sorted = sorted(files, key=lambda x: x.lower())
            
            # 先处理文件夹
            for dir_name in dirs_sorted:
                dir_path = os.path.join(folder_path, dir_name)
                dir_node_id = f"dir_{uuid.uuid4()}"
                actual_node_id = self._add_node(
                    parent_node, dir_name, dir_node_id, 
                    "local_dir", dir_path, dir_path
                )
                self.tree.item(actual_node_id, open=True)
                self._traverse_folder(dir_path, actual_node_id, level + 1)
            
            # 后处理文件
            for file_name in files_sorted:
                file_path = os.path.join(folder_path, file_name)
                file_node_id = f"file_{uuid.uuid4()}"
                self._add_node(
                    parent_node, file_name, file_node_id, 
                    "local_file", file_path, file_path
                )
                
        except PermissionError:
            # 权限不足
            err_node_id = f"err_{uuid.uuid4()}"
            self._add_node(
                parent_node, "🔒 无访问权限", err_node_id, 
                "error", "权限不足，无法访问该目录", ""
            )
        except Exception as e:
            # 其他错误
            err_node_id = f"err_{uuid.uuid4()}"
            self._add_node(
                parent_node, "❌ 访问错误", err_node_id, 
                "error", f"访问失败：{str(e)}", ""
            )

    def guess_archive_type(self, name, data=None):
        """判断压缩包类型"""
        name_lower = name.lower()
        if name_lower.endswith('.zip'):
            return 'zip'
        elif name_lower.endswith('.tar'):
            return 'tar'
        elif name_lower.endswith(('.tar.gz', '.tgz')):
            return 'tar.gz'
        elif name_lower.endswith(('.tar.bz2', '.tbz2')):
            return 'tar.bz2'
        elif name_lower.endswith('.gz'):
            return 'gz'
        elif name_lower.endswith('.bz2'):
            return 'bz2'
        
        # 通过文件头判断
        if data:
            if data.startswith(b'PK\x03\x04'):
                return 'zip'
            elif data.startswith(b'\x1f\x8b'):
                return 'gz'
            elif data.startswith(b'BZh'):
                return 'bz2'
        return None

    def load_file_or_archive(self, file_path):
        """加载文件或压缩包"""
        file_path = normalize_input_path(file_path)
        self.clear_all()
        self.current_path = file_path
        
        # 判断是否是压缩包
        if self.guess_archive_type(file_path):
            self.scan_archive(file_path, "", 0)
        else:
            # 普通文件（支持多格式）
            try:
                file_ext = os.path.splitext(file_path)[1].lower()
                content = ""
                file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0

                if file_ext == '.pdf':
                    self.clear_pdf_lazy_state()
                    file_node_id = f"file_{uuid.uuid4()}"
                    self._add_node("", os.path.basename(file_path), file_node_id, "file", {"kind": "pdf_path", "path": file_path}, file_path)
                    self.setup_pdf_lazy({"path": file_path}, os.path.basename(file_path))
                    self.title(f"文件查看器 - {os.path.basename(file_path)}")
                    return
                if file_ext in IMAGE_EXTENSIONS:
                    self.clear_pdf_lazy_state()
                    file_node_id = f"file_{uuid.uuid4()}"
                    self._add_node("", os.path.basename(file_path), file_node_id, "file", {"kind": "image_path", "path": file_path}, file_path)
                    if PIL_AVAILABLE:
                        with Image.open(file_path) as img:
                            self.display_image_in_editor(img, title=f"=== 图片文件：{os.path.basename(file_path)} ===")
                    else:
                        self.display_text_in_editor("(未安装Pillow，无法显示图片。请安装：pip install pillow)")
                    self.set_current_file_label(os.path.basename(file_path))
                    self.title(f"文件查看器 - {os.path.basename(file_path)}")
                    return

                if file_size > LARGE_FILE_THRESHOLD:
                    if file_ext == '.pdf':
                        parsed = extract_pdf_text(file_path=file_path, max_pages=PDF_LARGE_MAX_PAGES)
                        content = (
                            f"=== 大PDF预览：{os.path.basename(file_path)} ({file_size/1024/1024:.2f}MB) ===\n"
                            f"[性能模式] 仅提取前{PDF_LARGE_MAX_PAGES}页文本。\n\n{parsed}"
                        )
                    else:
                        preview = read_text_preview_from_path(file_path)
                        content = f"=== 大文件预览：{os.path.basename(file_path)} ({file_size/1024/1024:.2f}MB) ===\n\n{preview}"
                else:
                
                    # Excel/CSV文件处理
                    if file_ext in ['.xlsx', '.xls', '.csv']:
                        if PANDAS_AVAILABLE:
                            if file_ext in ['.xlsx', '.xls']:
                                excel_data = pd.ExcelFile(file_path)
                                content = []
                                content.append(f"=== Excel文件：{os.path.basename(file_path)} ===")
                                for sheet_name in excel_data.sheet_names:
                                    df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=TABULAR_PREVIEW_ROWS)
                                    content.append(f"\n--- Sheet: {sheet_name} ---")
                                    content.append(df.to_string(index=False))
                                content = "\n".join(content)
                            elif file_ext == '.csv':
                                df = pd.read_csv(file_path, nrows=TABULAR_PREVIEW_ROWS)
                                content = df.to_string(index=False)
                        else:
                            # 降级为普通文本读取
                            content = read_text_file_auto(file_path)
                    # JSON文件格式化
                    elif file_ext == '.json':
                        json_content = read_text_file_auto(file_path)
                        try:
                            json_data = json.loads(json_content)
                            content = json.dumps(json_data, indent=4, ensure_ascii=False)
                        except:
                            content = json_content
                    elif file_ext in ['.doc', '.docx']:
                        content = extract_doc_text(file_path)
                    else:
                        # 普通文本文件
                        content = read_text_file_auto(file_path)
                
                self.clear_pdf_lazy_state()
                self.current_text_content = content
                file_node_id = f"file_{uuid.uuid4()}"
                self._add_node("", os.path.basename(file_path), file_node_id, "file", content, file_path)
                self.display_text_in_editor(content, file_ext=file_ext)
                self.set_current_file_label(os.path.basename(file_path))
                
            except Exception as e:
                messagebox.showerror("错误", f"加载文件失败：{str(e)}")
        
        self.title(f"文件查看器 - {os.path.basename(file_path)}")

    def scan_archive(self, src, parent_node, level):
        """解析压缩包"""
        raw_data = None
        archive_name = ""
        
        if isinstance(src, str):
            # 本地压缩包文件
            if not os.path.isfile(src):
                return
            with open(src, 'rb') as f:
                raw_data = f.read()
            archive_name = os.path.basename(src)
            
            # 创建压缩包根节点
            arch_node_id = f"arch_{uuid.uuid4()}"
            arch_node = self._add_node(
                parent_node, archive_name, arch_node_id, 
                "archive", raw_data, src
            )
            self.tree.item(arch_node, open=True)
            parent_node = arch_node
        else:
            # 压缩包内的嵌套压缩包
            raw_data = src["data"]
            archive_name = os.path.basename(src["name"])
        
        buf = io.BytesIO(raw_data)
        files = []
        dirs = set()
        arch_type = self.guess_archive_type(archive_name, raw_data)
        
        try:
            if arch_type == 'zip':
                with zipfile.ZipFile(buf) as zf:
                    for info in zf.infolist():
                        if info.is_dir():
                            dirs.add(info.filename)
                        elif info.file_size > 0:
                            with zf.open(info) as f:
                                read_limit = info.file_size if info.file_size <= ARCHIVE_MEMBER_FULL_READ_BYTES else (ARCHIVE_MEMBER_PREVIEW_BYTES + 1)
                                member_data = f.read(read_limit)
                            files.append((info.filename, member_data, info.file_size))
                            # 添加文件所在目录
                            file_dir = os.path.dirname(info.filename)
                            if file_dir:
                                dirs.add(file_dir)
                            
            elif arch_type in ('tar', 'tar.gz', 'tar.bz2'):
                mode = {'tar':'r', 'tar.gz':'r:gz', 'tar.bz2':'r:bz2'}[arch_type]
                with tarfile.open(fileobj=buf, mode=mode) as tf:
                    for member in tf.getmembers():
                        if member.isdir():
                            dirs.add(member.name)
                        elif member.isfile() and member.size > 0:
                            extracted = tf.extractfile(member)
                            if extracted:
                                read_limit = member.size if member.size <= ARCHIVE_MEMBER_FULL_READ_BYTES else (ARCHIVE_MEMBER_PREVIEW_BYTES + 1)
                                file_data = extracted.read(read_limit)
                                files.append((member.name, file_data, member.size))
                            file_dir = os.path.dirname(member.name)
                            if file_dir:
                                dirs.add(file_dir)
                            
            elif arch_type == 'gz':
                with gzip.GzipFile(fileobj=buf) as gf:
                    new_name = archive_name[:-3] if archive_name.lower().endswith('.gz') else archive_name + '.unzip'
                    data = gf.read(ARCHIVE_MEMBER_FULL_READ_BYTES + 1)
                    original_size = len(data)
                    if len(data) > ARCHIVE_MEMBER_FULL_READ_BYTES:
                        data = data[:ARCHIVE_MEMBER_PREVIEW_BYTES + 1]
                    files.append((new_name, data, original_size))
                    
            elif arch_type == 'bz2':
                with bz2.BZ2File(fileobj=buf) as bf:
                    new_name = archive_name[:-4] if archive_name.lower().endswith('.bz2') else archive_name + '.unzip'
                    data = bf.read(ARCHIVE_MEMBER_FULL_READ_BYTES + 1)
                    original_size = len(data)
                    if len(data) > ARCHIVE_MEMBER_FULL_READ_BYTES:
                        data = data[:ARCHIVE_MEMBER_PREVIEW_BYTES + 1]
                    files.append((new_name, data, original_size))
                    
        except Exception as e:
            messagebox.showerror("解析错误", f"压缩包解析失败：{str(e)}")
            return
        
        # 排序：先文件夹（字母序），后文件（字母序）
        dirs_sorted = sorted([d for d in dirs if d], key=lambda x: x.lower())
        for dir_path in dirs_sorted:
            self._create_dir_node(dir_path.split('/'), parent_node, level)
        
        files_sorted = sorted(files, key=lambda x: os.path.basename(x[0]).lower())
        for file_path, file_data, original_size in files_sorted:
            file_name = os.path.basename(file_path)
            if not file_name:
                continue
                
            file_node_id = f"f_{uuid.uuid4()}"
            # 获取父目录节点
            dir_path = os.path.dirname(file_path)
            if dir_path:
                dir_node = self._create_dir_node(dir_path.split('/'), parent_node, level)
            else:
                dir_node = parent_node
            
            # 判断是否是嵌套压缩包
            if self.guess_archive_type(file_name, file_data):
                if original_size > len(file_data):
                    tip_content = (
                        f"=== 嵌套压缩包预览：{file_name} ===\n"
                        f"[性能模式] 仅加载了前{ARCHIVE_MEMBER_PREVIEW_BYTES // 1024 // 1024}MB，"
                        f"已跳过递归解析以避免卡顿。"
                    )
                    self._add_node(
                        dir_node, file_name, file_node_id,
                        "file", tip_content, file_path
                    )
                    continue
                # 嵌套压缩包
                nest_node = self._add_node(
                    dir_node, file_name, file_node_id, 
                    "archive", file_data, file_path
                )
                self.tree.item(nest_node, open=True)
                self.scan_archive({"name": file_name, "data": file_data}, nest_node, level+1)
            else:
                # 普通文件（支持多格式解析）
                try:
                    file_ext = os.path.splitext(file_name)[1].lower()
                    is_truncated_in_archive = original_size > len(file_data)
                    effective_data = file_data[:ARCHIVE_MEMBER_PREVIEW_BYTES] if is_truncated_in_archive else file_data
                    if is_truncated_in_archive:
                        if file_ext == '.pdf':
                            file_content = {
                                "kind": "pdf_bytes",
                                "name": file_name,
                                "bytes": effective_data,
                                "truncated": True,
                                "original_size": original_size
                            }
                        else:
                            preview = read_text_preview_from_bytes(effective_data)
                            file_content = (
                                f"=== 大文件预览：{file_name} ({original_size/1024/1024:.2f}MB) ===\n"
                                f"[性能模式] 压缩包内仅加载前{ARCHIVE_MEMBER_PREVIEW_BYTES // 1024 // 1024}MB数据。\n\n{preview}"
                            )
                    else:
                        if file_ext in ['.xlsx', '.xls', '.csv'] and PANDAS_AVAILABLE:
                            # 处理表格文件
                            if file_ext in ['.xlsx', '.xls']:
                                excel_buf = io.BytesIO(effective_data)
                                excel_data = pd.ExcelFile(excel_buf)
                                content = []
                                content.append(f"=== Excel文件：{file_name} ===")
                                for sheet_name in excel_data.sheet_names:
                                    excel_buf.seek(0)
                                    df = pd.read_excel(excel_buf, sheet_name=sheet_name)
                                    content.append(f"\n--- Sheet: {sheet_name} ---")
                                    content.append(df.to_string(index=False))
                                file_content = "\n".join(content)
                            elif file_ext == '.csv':
                                csv_buf = io.StringIO(decode_bytes_auto(effective_data))
                                df = pd.read_csv(csv_buf)
                                file_content = df.to_string(index=False)
                        elif file_ext == '.json':
                            # JSON格式化
                            json_content = decode_bytes_auto(effective_data)
                            try:
                                json_data = json.loads(json_content)
                                file_content = json.dumps(json_data, indent=4, ensure_ascii=False)
                            except:
                                file_content = json_content
                        elif file_ext == '.docx':
                            file_content = extract_docx_text(file_bytes=effective_data)
                        elif file_ext == '.doc':
                            file_content = "(压缩包内 .doc 暂不支持直接解析，请先解压到本地后打开)"
                        elif file_ext == '.pdf':
                            file_content = {
                                "kind": "pdf_bytes",
                                "name": file_name,
                                "bytes": effective_data,
                                "truncated": False,
                                "original_size": original_size
                            }
                        elif file_ext in IMAGE_EXTENSIONS:
                            file_content = {
                                "kind": "image_bytes",
                                "name": file_name,
                                "bytes": effective_data,
                                "original_size": original_size
                            }
                        else:
                            # 普通文本解码
                            file_content = decode_bytes_auto(effective_data)
                except:
                    file_content = "(二进制文件/解码失败)"
                
                self._add_node(
                    dir_node, file_name, file_node_id, 
                    "file", file_content, file_path
                )

    def _create_dir_node(self, path_parts, parent_node, level):
        """创建目录节点"""
        if not path_parts:
            return parent_node
        
        current_dir = path_parts[0]
        if not current_dir:
            return self._create_dir_node(path_parts[1:], parent_node, level)
            
        dir_node_id = f"dir_{uuid.uuid4()}"
        actual_node_id = self._add_node(
            parent_node, current_dir, dir_node_id, 
            "archive_dir", "", os.path.join(self.node_full_path.get(parent_node, ""), current_dir)
        )
        
        return self._create_dir_node(path_parts[1:], actual_node_id, level+1)

    # ==============================
    # 事件处理函数
    # ==============================
    def on_tree_select(self, event):
        """选择树节点"""
        selection = self.tree.selection()
        if not selection:
            return
        node_id = selection[0]
        node_text = self.tree.item(node_id, "text")
        node_type = self.node_type.get(node_id)
        if node_type in ("file", "local_file"):
            self._ensure_text_box_for_node(node_id, node_text)
        self.set_current_file_label(node_text)
        content = ""
        
        if node_type == "file":
            # 压缩包内的文件
            node_data = self.node_data.get(node_id, "")
            if isinstance(node_data, dict):
                if node_data.get("kind") == "pdf_bytes":
                    self.setup_pdf_lazy({"bytes": node_data.get("bytes", b"")}, node_data.get("name", self.tree.item(node_id, "text")))
                    return
                if node_data.get("kind") == "pdf_path":
                    path = node_data.get("path")
                    if path and os.path.exists(path):
                        self.setup_pdf_lazy({"path": path}, os.path.basename(path))
                        return
                if node_data.get("kind") == "image_bytes":
                    self.clear_pdf_lazy_state()
                    if PIL_AVAILABLE:
                        try:
                            with Image.open(io.BytesIO(node_data.get("bytes", b""))) as img:
                                self.display_image_in_editor(img, title=f"=== 图片文件：{node_data.get('name', '')} ===")
                        except Exception as e:
                            self.display_text_in_editor(f"(图片显示失败: {str(e)})")
                    else:
                        self.display_text_in_editor("(未安装Pillow，无法显示图片。请安装：pip install pillow)")
                    return
                if node_data.get("kind") == "image_path":
                    self.clear_pdf_lazy_state()
                    path = node_data.get("path")
                    if path and os.path.exists(path) and PIL_AVAILABLE:
                        try:
                            with Image.open(path) as img:
                                self.display_image_in_editor(img, title=f"=== 图片文件：{os.path.basename(path)} ===")
                        except Exception as e:
                            self.display_text_in_editor(f"(图片显示失败: {str(e)})")
                    else:
                        self.display_text_in_editor("(未安装Pillow或图片路径无效)")
                    return
            self.clear_pdf_lazy_state()
            content = node_data if isinstance(node_data, str) else str(node_data)
            self.current_text_content = content
        elif node_type == "local_file":
            # 本地文件
            file_path = normalize_input_path(self.node_data.get(node_id))
            if file_path and os.path.exists(file_path):
                try:
                    pending = self._pending_multi_jump
                    force_full_parse = bool(pending and pending.get("node_id") == node_id)
                    # 检查文件大小
                    file_size = os.path.getsize(file_path)
                    if file_size > self.max_file_size:
                        if not messagebox.askyesno("提示", 
                            f"文件大小 {file_size/1024/1024:.1f} MB，超过10MB限制，是否继续？"):
                            return
                    
                    # 支持多格式读取
                    file_ext = os.path.splitext(file_path)[1].lower()
                    if file_ext == '.pdf':
                        self.setup_pdf_lazy({"path": file_path}, os.path.basename(file_path))
                        return
                    if file_ext in IMAGE_EXTENSIONS:
                        self.clear_pdf_lazy_state()
                        if PIL_AVAILABLE:
                            with Image.open(file_path) as img:
                                self.display_image_in_editor(img, title=f"=== 图片文件：{os.path.basename(file_path)} ===")
                        else:
                            self.display_text_in_editor("(未安装Pillow，无法显示图片。请安装：pip install pillow)")
                        return
                    self.clear_pdf_lazy_state()
                    if file_size > LARGE_FILE_THRESHOLD and not force_full_parse:
                        preview = read_text_preview_from_path(file_path)
                        content = f"=== 大文件预览：{os.path.basename(file_path)} ({file_size/1024/1024:.2f}MB) ===\n\n{preview}"
                    else:
                        if file_ext in ['.xlsx', '.xls', '.csv'] and PANDAS_AVAILABLE:
                            if file_ext in ['.xlsx', '.xls']:
                                excel_data = pd.ExcelFile(file_path)
                                content = []
                                content.append(f"=== Excel文件：{os.path.basename(file_path)} ===")
                                for sheet_name in excel_data.sheet_names:
                                    if force_full_parse:
                                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                                    else:
                                        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=TABULAR_PREVIEW_ROWS)
                                    content.append(f"\n--- Sheet: {sheet_name} ---")
                                    content.append(df.to_string(index=False))
                                content = "\n".join(content)
                            elif file_ext == '.csv':
                                if force_full_parse:
                                    df = pd.read_csv(file_path)
                                else:
                                    df = pd.read_csv(file_path, nrows=TABULAR_PREVIEW_ROWS)
                                content = df.to_string(index=False)
                        elif file_ext == '.json':
                            json_content = read_text_file_auto(file_path)
                            try:
                                json_data = json.loads(json_content)
                                content = json.dumps(json_data, indent=4, ensure_ascii=False)
                            except:
                                content = json_content
                        elif file_ext in ['.doc', '.docx']:
                            content = extract_doc_text(file_path)
                        else:
                            content = read_text_file_auto(file_path)
                    
                    self.current_text_content = content
                except Exception as e:
                    content = f"读取失败：{str(e)}"
        elif node_type in ("archive", "archive_dir", "local_dir"):
            # 目录/压缩包
            self.clear_pdf_lazy_state()
            node_text = self.tree.item(node_id, "text")
            content = f"▶ {node_text}\n\n按「文件夹优先+字母序」显示子项（已默认展开）"
            self.current_text_content = ""
        elif node_type == "error":
            # 错误节点
            self.clear_pdf_lazy_state()
            content = self.node_data.get(node_id, "无法访问该目录/文件")
            self.current_text_content = ""

        # 更新文本框
        node_path = self.node_full_path.get(node_id, "")
        ext = os.path.splitext(str(node_path))[1].lower()
        self.display_text_in_editor(content, file_ext=ext)
        pending = self._pending_multi_jump
        if pending and pending.get("node_id") == node_id:
            self._pending_multi_jump = None
            row = pending.get("row", 1)
            col = pending.get("col", 0)
            keyword = pending.get("keyword", "")
            self.after_idle(lambda r=row, c=col, kw=keyword: self._apply_jump_to_editor(r, c, kw))

    def on_drop(self, event):
        """拖拽文件/文件夹"""
        try:
            # 解析拖拽路径
            file_path = event.data
            if "} {" in file_path:
                file_path = file_path.split("} {")[0]
            file_path = normalize_input_path(file_path)
            
            # 清空映射表
            self.parent_node_files.clear()
            
            # 加载
            if os.path.isdir(file_path):
                self.load_folder(file_path)
            else:
                self.load_file_or_archive(file_path)
        except Exception as e:
            messagebox.showerror("拖拽错误", f"解析路径失败：{str(e)}\n\n建议：请检查文件路径是否正确，或尝试使用按钮选择文件")

    def open_compare_window(self):
        """打开文本比较窗口"""
        TextCompareWindow(self)

    def open_alarm_monitor_window(self):
        """打开监控告警窗口"""
        AlarmMonitorWindow(self)

    def open_stacked_text_box(self):
        """在文本框上方新增可选择的文本框页签。"""
        self._sync_active_text_box_content()
        title = self.current_file_label.cget("text").replace("当前文件：", "").strip() or f"文本框{self._text_box_seq}"
        new_id = f"tb_{self._text_box_seq}"
        self._text_box_seq += 1
        self.virtual_text_boxes.append({
            "id": new_id,
            "title": title,
            "content": self.txt.get("1.0", "end-1c"),
            "auto_title": True,
        })
        self.active_text_box_id = new_id
        self._refresh_text_box_selector()

    # ==============================
    # 搜索功能
    # ==============================
    def smart_search(self):
        """智能搜索"""
        keyword = self.search_entry.get().strip()
        if not keyword:
            messagebox.showinfo("提示", "请输入搜索关键词")
            return
        self.record_search_history(keyword)
        
        # 清空之前的搜索结果
        self.clear_all_highlights()
        
        if self.current_search_type == "file":
            # 文件名搜索
            self.search_filename(keyword)
        else:
            # 内容搜索
            self.search_content(keyword)

    def search_filename(self, keyword):
        """搜索文件名"""
        self.file_search_hits.clear()
        self.preview_line_to_hit_index.clear()
        self.preview_jump_entries.clear()
        
        # 递归遍历树节点
        def traverse(node):
            for child in self.tree.get_children(node):
                node_text = self.tree.item(child, "text").lower()
                if keyword.lower() in node_text:
                    # 标记为搜索命中
                    self.tree.item(child, tags=("search_hit",))
                    # 保存命中结果
                    self.file_search_hits.append((child, self.tree.item(child, "text")))
                # 递归子节点
                traverse(child)
        
        # 开始遍历
        traverse("")
        
        # 更新预览框
        self.result_txt.config(state=tk.NORMAL)
        self.result_txt.delete("1.0", tk.END)
        
        if self.file_search_hits:
            # 显示命中结果
            hit_texts = []
            for idx, (nid, name) in enumerate(self.file_search_hits, start=1):
                hit_texts.append(f"{idx}. {name}")
                self.preview_line_to_hit_index.append(idx - 1)
                self.preview_jump_entries.append(("file", nid))
            self.result_txt.insert("1.0", "\n".join(hit_texts))
            self.result_title.config(text=f"🔍 文件名搜索结果 | 匹配数：{len(self.file_search_hits)}")
            self.set_result_panel_collapsed(False)
            # 定位到第一个结果
            first_node = self.file_search_hits[0][0]
            self.tree.selection_set(first_node)
            self.tree.see(first_node)
        else:
            self.result_txt.insert("1.0", "未找到匹配的文件名")
            self.result_title.config(text="🔍 文件名搜索结果 | 匹配数：0")
            self.set_result_panel_collapsed(True)
        
        self.result_txt.config(state=tk.DISABLED)

    def search_content(self, keyword):
        """搜索内容"""
        self.current_text_content = self.txt.get("1.0", "end-1c")
        if not self.current_text_content:
            messagebox.showwarning("提示", "请先选择文件并确保有可搜索的文本内容")
            return
        
        self.content_search_hits.clear()
        self.preview_line_to_hit_index.clear()
        self.preview_jump_entries.clear()
        self.txt.tag_remove("hl", "1.0", tk.END)
        
        # 重新插入文本（避免干扰）
        self.txt.delete("1.0", tk.END)
        self.txt.insert("1.0", self.current_text_content)
        
        start_pos = "1.0"
        hit_count = 0
        
        # 查找所有匹配项
        while True:
            # 不区分大小写搜索
            pos = self.txt.search(keyword, start_pos, stopindex=tk.END, nocase=True)
            if not pos:
                break
            
            # 计算结束位置
            row, col = pos.split(".")
            end_pos = f"{row}.{int(col) + len(keyword)}"
            
            # 标记高亮
            self.txt.tag_add("hl", pos, end_pos)
            
            # 获取行内容
            line_start = f"{row}.0"
            line_end = f"{row}.end"
            line_content = self.txt.get(line_start, line_end).strip()
            
            # 保存命中结果
            self.content_search_hits.append((int(row), int(col), line_content))
            
            # 更新起始位置
            start_pos = end_pos
            hit_count += 1
        
        self._upsert_result_history("content", keyword, self.content_search_hits)
        self._render_result_history("content")
        self.refresh_line_numbers()

    def _extract_text_lines_for_search(self, node_id):
        node_type = self.node_type.get(node_id)

        # 缓存命中判断（避免重复解析大文件/PDF/DOC）
        cache_key = None
        if node_type == "file":
            data = self.node_data.get(node_id, "")
            if isinstance(data, str):
                sample_head = data[:512]
                sample_tail = data[-512:] if len(data) > 512 else data
                cache_key = ("mem", len(data), sample_head, sample_tail)
        elif node_type == "local_file":
            file_path = normalize_input_path(self.node_data.get(node_id))
            if file_path and os.path.exists(file_path):
                try:
                    st = os.stat(file_path)
                    cache_key = ("path", file_path, int(st.st_mtime), st.st_size)
                except Exception:
                    cache_key = ("path", file_path, 0, 0)

        if cache_key is not None:
            cached = self.search_text_cache.get(node_id)
            if cached and cached.get("key") == cache_key:
                return cached.get("lines", [])

        if node_type == "file":
            data = self.node_data.get(node_id, "")
            if isinstance(data, str):
                lines = data.splitlines()
                if cache_key is not None:
                    self.search_text_cache[node_id] = {"key": cache_key, "lines": lines}
                return lines
            return []
        if node_type == "local_file":
            file_path = normalize_input_path(self.node_data.get(node_id))
            if not file_path or not os.path.exists(file_path):
                return []
            file_ext = os.path.splitext(file_path)[1].lower()
            try:
                if file_ext == ".json":
                    content = read_text_file_auto(file_path)
                    try:
                        lines = json.dumps(json.loads(content), indent=4, ensure_ascii=False).splitlines()
                        if cache_key is not None:
                            self.search_text_cache[node_id] = {"key": cache_key, "lines": lines}
                        return lines
                    except Exception:
                        lines = content.splitlines()
                        if cache_key is not None:
                            self.search_text_cache[node_id] = {"key": cache_key, "lines": lines}
                        return lines
                if file_ext in [".doc", ".docx"]:
                    lines = extract_doc_text(file_path).splitlines()
                    if cache_key is not None:
                        self.search_text_cache[node_id] = {"key": cache_key, "lines": lines}
                    return lines
                if file_ext == ".pdf":
                    lines = extract_pdf_text(file_path=file_path, max_pages=PDF_LARGE_MAX_PAGES).splitlines()
                    if cache_key is not None:
                        self.search_text_cache[node_id] = {"key": cache_key, "lines": lines}
                    return lines
                if file_ext in IMAGE_EXTENSIONS:
                    return []
                lines = read_text_file_auto(file_path).splitlines()
                if cache_key is not None:
                    self.search_text_cache[node_id] = {"key": cache_key, "lines": lines}
                return lines
            except Exception:
                return []
        return []

    def search_content_multi(self):
        """多文件内容搜索"""
        keyword = self.search_entry.get().strip()
        if not keyword:
            messagebox.showinfo("提示", "请输入搜索关键词")
            return
        self.record_search_history(keyword)

        self.clear_all_highlights()
        self.multi_content_search_hits.clear()
        self.preview_line_to_hit_index.clear()
        self.preview_jump_entries.clear()
        self.current_search_type = "content_multi"
        kw_low = keyword.lower()
        self.config(cursor="watch")
        self.update_idletasks()

        def traverse(node):
            for child in self.tree.get_children(node):
                node_type = self.node_type.get(child)
                if node_type in ("file", "local_file"):
                    lines = self._extract_text_lines_for_search(child)
                    for row_idx, line in enumerate(lines, start=1):
                        if not line:
                            continue
                        line_low = line.lower()
                        col = line_low.find(kw_low)
                        if col >= 0:
                            self.multi_content_search_hits.append((child, row_idx, col, line.strip()))
                traverse(child)

        try:
            traverse("")
        finally:
            self.config(cursor="")

        self._upsert_result_history("content_multi", keyword, self.multi_content_search_hits)
        self._render_result_history("content_multi")

    def on_double_click_jump(self, event):
        """双击搜索结果跳转"""
        try:
            click_index = self.result_txt.index(f"@{event.x},{event.y}")
            line_num = int(click_index.split(".")[0]) - 1
            if line_num < 0:
                return
            # 点击空白区域或末尾时，回退到最后一条有效映射
            if line_num >= len(self.preview_jump_entries):
                line_num = len(self.preview_jump_entries) - 1
            if line_num < 0:
                return
            entry = self.preview_jump_entries[line_num]
            if not entry:
                return
            jump_type = entry[0]
            if jump_type == "history_toggle":
                mode, kw = entry[1], entry[2]
                self._toggle_result_history(mode, kw)
                self._render_result_history(mode)
                return

            if jump_type == "file":
                node_id = entry[1]
                self.tree.selection_set(node_id)
                self.tree.see(node_id)
                self.tree.focus_set()
            elif jump_type == "content":
                row, col = entry[1], entry[2]
                self._apply_jump_to_editor(row, col, self.search_entry.get().strip())
            elif jump_type == "content_multi":
                node_id, row, col = entry[1], entry[2], entry[3]
                kw = self.search_entry.get().strip()
                current_selection = self.tree.selection()
                if current_selection and current_selection[0] == node_id:
                    self._apply_jump_to_editor(row, col, kw)
                else:
                    self._pending_multi_jump = {"node_id": node_id, "row": int(row), "col": int(col), "keyword": kw}
                    self.tree.selection_set(node_id)
                    self.tree.see(node_id)
                    self.tree.focus_set()
        except Exception as e:
            messagebox.showinfo("提示", f"跳转失败：{str(e)}\n请选择有效的搜索结果行")

    def clear_all_highlights(self):
        """清空所有高亮"""
        # 清空树节点高亮
        def clear_tree(node):
            for child in self.tree.get_children(node):
                self.tree.item(child, tags=())
                clear_tree(child)
        clear_tree("")
        
        # 清空文本高亮
        self.txt.tag_remove("hl", "1.0", tk.END)
        self.txt.tag_remove("jump_hl", "1.0", tk.END)
        
        # 清空预览框
        self.result_txt.config(state=tk.NORMAL)
        self.result_txt.delete("1.0", tk.END)
        self.result_txt.config(state=tk.DISABLED)
        
        # 重置标题
        self.result_title.config(text="🔍 搜索结果预览 | 匹配数：0")
        self.set_result_panel_collapsed(True)
        
        # 清空搜索结果列表
        self.file_search_hits.clear()
        self.content_search_hits.clear()
        self.multi_content_search_hits.clear()
        self.preview_line_to_hit_index.clear()
        self.preview_jump_entries.clear()
        self.search_text_cache.clear()

