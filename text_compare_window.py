# -*- coding: utf-8 -*-
from app_common import *

class TextCompareWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("文本比较工具 | 智能差异高亮 | 支持TXT/CSV/XLSX/JSON等")
        self.geometry("1200x700")
        self.configure(bg="#f4f4f4")
        self.resizable(True, True)

        self.text1_content = ""
        self.text2_content = ""
        self.text1_raw = ""  # 原始内容
        self.text2_raw = ""  # 原始内容
        self.file1_format = ""  # 文件格式
        self.file2_format = ""  # 文件格式
        self.file1_path = ""
        self.file2_path = ""
        self.compare_font_size = 10
        self.compare_font = tkfont.Font(family="Consolas", size=self.compare_font_size)
        self.compare_bold_font = tkfont.Font(family="Consolas", size=self.compare_font_size, weight="bold")

        # 加载区域
        self.frame_load = tk.Frame(self, bg="#e0f0e0", bd=2, relief=tk.RIDGE)
        self.frame_load.pack(fill="x", padx=5, pady=3)
        tk.Label(self.frame_load, text="📁 加载对比文件", bg="#e0f0e0", font=("微软雅黑", 10, "bold")).pack(pady=2)
        load_inner = tk.Frame(self.frame_load, bg="#e0f0e0")
        load_inner.pack(pady=4)

        self.btn_load1 = tk.Button(load_inner, text="加载文件1", width=15, command=lambda: self.load_file(1),
                                  font=("微软雅黑", 9), bg="#4CAF50", fg="white")
        self.btn_load1.pack(side="left", padx=5)
        self.btn_load2 = tk.Button(load_inner, text="加载文件2", width=15, command=lambda: self.load_file(2),
                                  font=("微软雅黑", 9), bg="#2196F3", fg="white")
        self.btn_load2.pack(side="left", padx=5)
        self.btn_compare = tk.Button(load_inner, text="开始对比", width=15, command=self.compare_texts,
                                    font=("微软雅黑", 9), bg="#FF9800", fg="white")
        self.btn_compare.pack(side="left", padx=5)
        
        # 对比模式选择
        self.compare_mode = tk.StringVar(value="smart")
        mode_frame = tk.Frame(self.frame_load, bg="#e0f0e0")
        mode_frame.pack(pady=2)
        tk.Label(mode_frame, text="对比模式：", bg="#e0f0e0", font=("微软雅黑", 9)).pack(side="left", padx=2)
        tk.Radiobutton(mode_frame, text="智能模式（按格式）", variable=self.compare_mode, value="smart", 
                      bg="#e0f0e0", font=("微软雅黑", 8)).pack(side="left", padx=5)
        tk.Radiobutton(mode_frame, text="行级对比（通用）", variable=self.compare_mode, value="line", 
                      bg="#e0f0e0", font=("微软雅黑", 8)).pack(side="left", padx=5)
        tk.Radiobutton(mode_frame, text="单词级对比", variable=self.compare_mode, value="word", 
                      bg="#e0f0e0", font=("微软雅黑", 8)).pack(side="left", padx=5)

        tip_text = "💡 支持格式：TXT/CSV/XLSX/JSON/XML/PY/Md | 拖拽文件到文本框即可"
        if not PANDAS_AVAILABLE:
            tip_text += " (未安装pandas，XLSX/CSV功能受限)"
        tk.Label(self.frame_load, text=tip_text, bg="#e0f0e0", font=("微软雅黑", 9, "italic")).pack(pady=2)

        # 编辑工具栏
        self.frame_edit = tk.Frame(self, bg="#f0f0f0", bd=2, relief=tk.RIDGE)
        self.frame_edit.pack(fill="x", padx=5, pady=2)
        tk.Label(self.frame_edit, text="✏️ 文字编辑工具", bg="#f0f0f0", font=("微软雅黑", 9, "bold")).pack(pady=1)
        edit_inner = tk.Frame(self.frame_edit, bg="#f0f0f0")
        edit_inner.pack(pady=1)

        self.bold_btn = tk.Button(edit_inner, text="𝐁 加粗", width=7, command=self.set_bold, font=("微软雅黑", 8))
        self.bold_btn.pack(side="left", padx=5)
        self.color_btn = tk.Menubutton(edit_inner, text="🎨 颜色", width=7, font=("微软雅黑", 8), relief=tk.RAISED, bg="#f0f0f0")
        self.color_btn.pack(side="left", padx=5)
        self.color_menu = tk.Menu(self.color_btn, tearoff=0, bg="white", bd=1)
        self.color_btn.config(menu=self.color_menu)
        self.build_color_menu()
        self.reset_btn = tk.Button(edit_inner, text="🔄 恢复默认", width=9, command=self.reset_style, font=("微软雅黑", 8))
        self.reset_btn.pack(side="left", padx=5)
        self.save_btn = tk.Button(edit_inner, text="保存文本", width=9, command=self.save_active_text, font=("微软雅黑", 8))
        self.save_btn.pack(side="left", padx=5)

        # 对比区域
        self.paned_main = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief=tk.RIDGE, bg="#f4f4f4")
        self.paned_main.pack(fill="both", expand=True, padx=5, pady=3)

        # 文本框1
        self.frame_text1 = tk.Frame(self.paned_main, bg="#f8f8f8")
        tk.Label(self.frame_text1, text="文件1内容", bg="#f8f8f8", font=("微软雅黑", 10, "bold")).pack(fill="x")
        self.text1_container = tk.Frame(self.frame_text1, bg="#f8f8f8")
        self.text1_container.pack(fill="both", expand=True, padx=2, pady=(2, 0))
        self.text1_ybar = tk.Scrollbar(self.text1_container, orient="vertical")
        self.text1_ybar.pack(side="right", fill="y")
        self.text1 = tk.Text(self.text1_container, font=self.compare_font, wrap=tk.NONE, bg="white", undo=True)
        self.text1.pack(side="left", fill="both", expand=True)
        self.text1.configure(yscrollcommand=self.text1_ybar.set)
        self.text1_ybar.config(command=self.text1.yview)
        self.text1_xbar_bar = tk.Frame(self.frame_text1, bg="#f8f8f8", height=16)
        self.text1_xbar_bar.pack(fill="x", padx=2, pady=(0, 2))
        self.text1_xbar_bar.pack_propagate(False)
        self.text1_xbar = tk.Scrollbar(self.text1_xbar_bar, orient="horizontal", command=self.text1.xview)
        self.text1_xbar.pack(fill="both", expand=True)
        self.text1.configure(xscrollcommand=self.text1_xbar.set)
        self.text1.drop_target_register(tkdnd.DND_FILES)
        self.text1.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, 1))
        self.text1.bind("<FocusIn>", lambda e: self.set_active_compare_text(self.text1))
        self.text1.bind("<Control-s>", self.save_active_text)
        self.text1.bind("<Control-S>", self.save_active_text)
        self.text1.bind("<Control-MouseWheel>", self.on_compare_zoom)
        self.text1.bind("<Control-Button-4>", self.on_compare_zoom)
        self.text1.bind("<Control-Button-5>", self.on_compare_zoom)

        # 文本框2
        self.frame_text2 = tk.Frame(self.paned_main, bg="#f8f8f8")
        tk.Label(self.frame_text2, text="文件2内容", bg="#f8f8f8", font=("微软雅黑", 10, "bold")).pack(fill="x")
        self.text2_container = tk.Frame(self.frame_text2, bg="#f8f8f8")
        self.text2_container.pack(fill="both", expand=True, padx=2, pady=(2, 0))
        self.text2_ybar = tk.Scrollbar(self.text2_container, orient="vertical")
        self.text2_ybar.pack(side="right", fill="y")
        self.text2 = tk.Text(self.text2_container, font=self.compare_font, wrap=tk.NONE, bg="white", undo=True)
        self.text2.pack(side="left", fill="both", expand=True)
        self.text2.configure(yscrollcommand=self.text2_ybar.set)
        self.text2_ybar.config(command=self.text2.yview)
        self.text2_xbar_bar = tk.Frame(self.frame_text2, bg="#f8f8f8", height=16)
        self.text2_xbar_bar.pack(fill="x", padx=2, pady=(0, 2))
        self.text2_xbar_bar.pack_propagate(False)
        self.text2_xbar = tk.Scrollbar(self.text2_xbar_bar, orient="horizontal", command=self.text2.xview)
        self.text2_xbar.pack(fill="both", expand=True)
        self.text2.configure(xscrollcommand=self.text2_xbar.set)
        self.text2.drop_target_register(tkdnd.DND_FILES)
        self.text2.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, 2))
        self.text2.bind("<FocusIn>", lambda e: self.set_active_compare_text(self.text2))
        self.text2.bind("<Control-s>", self.save_active_text)
        self.text2.bind("<Control-S>", self.save_active_text)
        self.text2.bind("<Control-MouseWheel>", self.on_compare_zoom)
        self.text2.bind("<Control-Button-4>", self.on_compare_zoom)
        self.text2.bind("<Control-Button-5>", self.on_compare_zoom)

        self.paned_main.add(self.frame_text1, width=580)
        self.paned_main.add(self.frame_text2, width=580)
        self.active_compare_text = self.text1
        self.init_tags()

    def init_tags(self):
        # 不同类型的差异高亮
        self.text1.tag_config("diff_add", background="#E6FFE6")    # 新增内容 - 浅绿
        self.text1.tag_config("diff_remove", background="#FFE6E6") # 删除内容 - 浅红
        self.text1.tag_config("diff_change", background="#FFFFE6") # 修改内容 - 浅黄
        self.text2.tag_config("diff_add", background="#E6FFE6")
        self.text2.tag_config("diff_remove", background="#FFE6E6")
        self.text2.tag_config("diff_change", background="#FFFFE6")
        
        # 兼容原有标签
        self.text1.tag_config("diff", background="#FFE4B5")
        self.text2.tag_config("diff", background="#FFE4B5")
        
        self.text1.tag_config("bold", font=self.compare_bold_font)
        self.text2.tag_config("bold", font=self.compare_bold_font)
        self.text1.tag_config("color", foreground="black")
        self.text2.tag_config("color", foreground="black")

    def get_widget_content(self, widget):
        return widget.get("1.0", "end-1c")

    def sync_text_contents(self):
        self.text1_content = self.get_widget_content(self.text1)
        self.text2_content = self.get_widget_content(self.text2)

    def set_active_compare_text(self, widget):
        self.active_compare_text = widget

    def on_compare_zoom(self, event):
        if hasattr(event, "num") and event.num in (4, 5):
            delta = 1 if event.num == 4 else -1
        else:
            delta = 1 if event.delta > 0 else -1
        new_size = max(8, min(36, self.compare_font_size + delta))
        if new_size == self.compare_font_size:
            return "break"
        self.compare_font_size = new_size
        self.compare_font.configure(size=new_size)
        self.compare_bold_font.configure(size=new_size)
        return "break"

    def save_active_text(self, event=None):
        active_text = self.focus_get()
        if active_text not in [self.text1, self.text2]:
            active_text = self.active_compare_text
        if active_text not in [self.text1, self.text2]:
            active_text = self.text1

        file_path = self.file1_path if active_text == self.text1 else self.file2_path
        file_format = self.file1_format if active_text == self.text1 else self.file2_format
        initial_name = os.path.basename(file_path) if file_path else ("compare_text_1.txt" if active_text == self.text1 else "compare_text_2.txt")
        defaultextension = file_format if file_format else ".txt"
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
                f.write(self.get_widget_content(active_text))
            if active_text == self.text1:
                self.file1_path = save_path
                self.file1_format = os.path.splitext(save_path)[1].lower()
            else:
                self.file2_path = save_path
                self.file2_format = os.path.splitext(save_path)[1].lower()
            self.sync_text_contents()
            messagebox.showinfo("提示", f"文本已保存：\n{save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")
        return "break" if event else None

    def build_color_menu(self):
        colors = [
            ("黑色", "#000000"), ("红色", "#FF0000"), ("橙色", "#FF6600"),
            ("黄色", "#FFFF00"), ("绿色", "#00B050"), ("蓝色", "#0070C0"), ("紫色", "#7030A0")
        ]
        for name, hex_color in colors:
            fg = "white" if self.is_dark_color(hex_color) else "black"
            self.color_menu.add_command(
                label=name, background=hex_color, foreground=fg,
                command=lambda c=hex_color: self.set_color(c)
            )

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

    def parse_json(self, content):
        """解析JSON并格式化"""
        try:
            data = json.loads(content)
            return json.dumps(data, indent=4, ensure_ascii=False)
        except:
            return content

    def read_file_content(self, file_path):
        """增强版文件读取：支持多格式转换为文本"""
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            self.current_file_format = file_ext
            file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0

            if file_size > LARGE_FILE_THRESHOLD:
                if file_ext == '.pdf':
                    content = extract_pdf_text(file_path=file_path, max_pages=PDF_LARGE_MAX_PAGES)
                    return (
                        f"=== 大PDF预览：{os.path.basename(file_path)} ({file_size/1024/1024:.2f}MB) ===\n"
                        f"[性能模式] 仅提取前{PDF_LARGE_MAX_PAGES}页文本。\n\n{content}",
                        file_ext
                    )
                preview = read_text_preview_from_path(file_path)
                return f"=== 大文件预览：{os.path.basename(file_path)} ({file_size/1024/1024:.2f}MB) ===\n\n{preview}", file_ext
            
            # Excel文件 (.xlsx/.xls)
            if file_ext in ['.xlsx', '.xls']:
                if not PANDAS_AVAILABLE:
                    messagebox.showwarning("提示", "未安装pandas库，请先安装：pip install pandas openpyxl")
                    with open(file_path, 'rb') as f:
                        return "(Excel文件需要pandas支持)", file_ext
                try:
                    # 读取Excel所有sheet
                    excel_data = pd.ExcelFile(file_path)
                    content = []
                    content.append(f"=== Excel文件：{os.path.basename(file_path)} ===")
                    for sheet_name in excel_data.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=TABULAR_PREVIEW_ROWS)
                        content.append(f"\n--- Sheet: {sheet_name} ---")
                        content.append(df.to_string(index=False))
                    return "\n".join(content), file_ext
                except Exception as e:
                    messagebox.showerror("错误", f"读取Excel失败：{str(e)}")
                    return None, file_ext
            
            # CSV文件
            elif file_ext == '.csv':
                if not PANDAS_AVAILABLE:
                    # 降级为普通文本读取
                    return read_text_file_auto(file_path), file_ext
                try:
                    df = pd.read_csv(file_path, nrows=TABULAR_PREVIEW_ROWS)
                    return df.to_string(index=False), file_ext
                except Exception as e:
                    # 降级读取
                    return read_text_file_auto(file_path), file_ext
            
            # JSON文件
            elif file_ext == '.json':
                content = read_text_file_auto(file_path)
                # 格式化JSON以便对比
                return self.parse_json(content), file_ext
            elif file_ext in ['.doc', '.docx']:
                return extract_doc_text(file_path), file_ext
            elif file_ext == '.pdf':
                return extract_pdf_text(file_path=file_path), file_ext
            
            # 普通文本文件 (txt/xml/md/py等)
            else:
                return read_text_file_auto(file_path), file_ext
                    
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败：{str(e)}")
            return None, ""

    def load_file(self, file_num):
        """加载文件（支持多格式）"""
        file_path = filedialog.askopenfilename(
            title=f"选择文件{file_num}",
            filetypes=[
                ("所有支持的文件", "*.txt *.csv *.xlsx *.xls *.json *.xml *.md *.py *.html *.log *.pdf *.doc *.docx"),
                ("文本文件", "*.txt *.log *.md *.py *.json *.xml *.html *.pdf *.doc *.docx"),
                ("表格文件", "*.csv *.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if not file_path:
            return
        
        content, file_format = self.read_file_content(file_path)
        if content is None:
            return
        
        if file_num == 1:
            self.file1_path = file_path
            self.text1_content = content
            self.text1_raw = content
            self.file1_format = file_format
            self.text1.delete("1.0", tk.END)
            self.text1.insert("1.0", content)
        else:
            self.file2_path = file_path
            self.text2_content = content
            self.text2_raw = content
            self.file2_format = file_format
            self.text2.delete("1.0", tk.END)
            self.text2.insert("1.0", content)

    def on_drop(self, event, file_num):
        """拖拽加载文件（支持多格式）"""
        try:
            file_path = event.data
            if "} {" in file_path:
                file_path = file_path.split("} {")[0]
            file_path = normalize_input_path(file_path)
            
            content, file_format = self.read_file_content(file_path)
            if content is None:
                return
            
            if file_num == 1:
                self.file1_path = file_path
                self.text1_content = content
                self.text1_raw = content
                self.file1_format = file_format
                self.text1.delete("1.0", tk.END)
                self.text1.insert("1.0", content)
            else:
                self.file2_path = file_path
                self.text2_content = content
                self.text2_raw = content
                self.file2_format = file_format
                self.text2.delete("1.0", tk.END)
                self.text2.insert("1.0", content)
        except Exception as e:
            messagebox.showerror("错误", f"拖拽文件失败：{str(e)}")

    def highlight_word_differences(self, text_widget, original_text, compare_text, is_text1=True):
        """单词级差异高亮"""
        # 清除原有高亮
        text_widget.tag_remove("diff_add", "1.0", tk.END)
        text_widget.tag_remove("diff_remove", "1.0", tk.END)
        text_widget.tag_remove("diff_change", "1.0", tk.END)
        
        # 使用SequenceMatcher找差异
        s = SequenceMatcher(None, original_text, compare_text)
        
        # 重新插入原始文本
        text_widget.delete("1.0", tk.END)
        text_widget.insert("1.0", original_text)
        
        # 标记差异
        for tag, i1, i2, j1, j2 in s.get_opcodes():
            if tag == 'replace':
                # 内容修改
                start = text_widget.index(f"1.0 + {i1} chars")
                end = text_widget.index(f"1.0 + {i2} chars")
                text_widget.tag_add("diff_change", start, end)
            elif tag == 'delete':
                # 内容删除
                start = text_widget.index(f"1.0 + {i1} chars")
                end = text_widget.index(f"1.0 + {i2} chars")
                text_widget.tag_add("diff_remove", start, end)
            elif tag == 'insert':
                # 只在第二个文本框标记新增
                if not is_text1:
                    start = text_widget.index(f"1.0 + {j1} chars")
                    end = text_widget.index(f"1.0 + {j2} chars")
                    text_widget.tag_add("diff_add", start, end)

    def compare_json_structures(self):
        """JSON结构化对比"""
        try:
            # 尝试解析JSON
            json1 = json.loads(self.text1_raw)
            json2 = json.loads(self.text2_raw)
            
            # 递归比较JSON
            def compare_json(json_a, json_b, path=""):
                diffs = []
                
                if type(json_a) != type(json_b):
                    diffs.append(f"{path}: 类型不同 - {type(json_a).__name__} vs {type(json_b).__name__}")
                    return diffs
                
                if isinstance(json_a, dict):
                    # 比较字典
                    all_keys = set(json_a.keys()) | set(json_b.keys())
                    for key in all_keys:
                        new_path = f"{path}.{key}" if path else key
                        if key not in json_a:
                            diffs.append(f"{new_path}: 新增键 - 值: {json_b[key]}")
                        elif key not in json_b:
                            diffs.append(f"{new_path}: 删除键 - 值: {json_a[key]}")
                        else:
                            diffs.extend(compare_json(json_a[key], json_b[key], new_path))
                            
                elif isinstance(json_a, list):
                    # 比较列表
                    max_len = max(len(json_a), len(json_b))
                    for i in range(max_len):
                        new_path = f"{path}[{i}]"
                        if i >= len(json_a):
                            diffs.append(f"{new_path}: 新增元素 - 值: {json_b[i]}")
                        elif i >= len(json_b):
                            diffs.append(f"{new_path}: 删除元素 - 值: {json_a[i]}")
                        else:
                            diffs.extend(compare_json(json_a[i], json_b[i], new_path))
                            
                else:
                    # 基本类型比较
                    if json_a != json_b:
                        diffs.append(f"{path}: 值不同 - {json_a} vs {json_b}")
                
                return diffs
            
            # 获取差异
            differences = compare_json(json1, json2)
            
            # 显示差异摘要
            diff_summary = "\n".join(differences) if differences else "两个JSON文件内容完全相同"
            
            # 格式化JSON以便行级对比
            formatted1 = json.dumps(json1, indent=4, ensure_ascii=False)
            formatted2 = json.dumps(json2, indent=4, ensure_ascii=False)
            
            # 更新文本框
            self.text1.delete("1.0", tk.END)
            self.text1.insert("1.0", formatted1)
            self.text2.delete("1.0", tk.END)
            self.text2.insert("1.0", formatted2)
            
            # 行级高亮差异
            self.highlight_line_differences(formatted1, formatted2)
            
            # 显示差异统计
            messagebox.showinfo("JSON对比结果", 
                               f"找到 {len(differences)} 处差异：\n\n{diff_summary}")
            
        except json.JSONDecodeError:
            # 不是有效的JSON，退回到普通对比
            self.compare_texts_fallback()
        except Exception as e:
            messagebox.showerror("错误", f"JSON对比失败：{str(e)}")
            self.compare_texts_fallback()

    def highlight_line_differences(self, text1, text2):
        """行级差异高亮"""
        # 清除原有高亮
        self.text1.tag_remove("diff", "1.0", tk.END)
        self.text2.tag_remove("diff", "1.0", tk.END)
        self.text1.tag_remove("diff_add", "1.0", tk.END)
        self.text1.tag_remove("diff_remove", "1.0", tk.END)
        self.text1.tag_remove("diff_change", "1.0", tk.END)
        self.text2.tag_remove("diff_add", "1.0", tk.END)
        self.text2.tag_remove("diff_remove", "1.0", tk.END)
        self.text2.tag_remove("diff_change", "1.0", tk.END)
        
        # 按行对比
        lines1 = text1.splitlines()
        lines2 = text2.splitlines()
        
        differ = Differ()
        diff_result = list(differ.compare(lines1, lines2))
        
        line1_idx = 0
        line2_idx = 0
        
        for line in diff_result:
            if line.startswith('- '):
                # 文件1独有的行（删除）
                if line1_idx < len(lines1):
                    self.text1.tag_add("diff_remove", f"{line1_idx+1}.0", f"{line1_idx+1}.end")
                    line1_idx += 1
            elif line.startswith('+ '):
                # 文件2独有的行（新增）
                if line2_idx < len(lines2):
                    self.text2.tag_add("diff_add", f"{line2_idx+1}.0", f"{line2_idx+1}.end")
                    line2_idx += 1
            elif line.startswith('? '):
                # 行内差异标记，跳过
                continue
            elif line.startswith('  '):
                # 相同的行
                line1_idx += 1
                line2_idx += 1
            else:
                # 修改的行
                if line1_idx < len(lines1):
                    self.text1.tag_add("diff_change", f"{line1_idx+1}.0", f"{line1_idx+1}.end")
                    line1_idx += 1
                if line2_idx < len(lines2):
                    self.text2.tag_add("diff_change", f"{line2_idx+1}.0", f"{line2_idx+1}.end")
                    line2_idx += 1

    def compare_texts_fallback(self):
        """降级对比方案"""
        self.sync_text_contents()
        if not self.text1_content or not self.text2_content:
            messagebox.showwarning("提示", "请先加载两个文件！")
            return
        
        # 行级对比
        self.highlight_line_differences(self.text1_content, self.text2_content)
        messagebox.showinfo("完成", "文本对比完成，差异部分已高亮显示！")

    def compare_texts(self):
        """智能文本对比"""
        self.sync_text_contents()
        self.text1_raw = self.text1_content
        self.text2_raw = self.text2_content
        if not self.text1_content or not self.text2_content:
            messagebox.showwarning("提示", "请先加载两个文件！")
            return
        
        compare_mode = self.compare_mode.get()
        
        if compare_mode == "smart":
            # 智能模式：根据文件格式选择对比方式
            if self.file1_format == ".json" and self.file2_format == ".json":
                # JSON专用对比
                self.compare_json_structures()
            elif self.file1_format in [".csv", ".xlsx", ".xls"] or self.file2_format in [".csv", ".xlsx", ".xls"]:
                # 表格文件行级对比
                self.highlight_line_differences(self.text1_content, self.text2_content)
                messagebox.showinfo("完成", "表格文件对比完成，差异行已高亮显示！")
            else:
                # 默认行级对比
                self.highlight_line_differences(self.text1_content, self.text2_content)
                messagebox.showinfo("完成", "文本对比完成，差异部分已高亮显示！")
                
        elif compare_mode == "line":
            # 行级对比
            self.highlight_line_differences(self.text1_content, self.text2_content)
            messagebox.showinfo("完成", "行级对比完成，差异行已高亮显示！")
            
        elif compare_mode == "word":
            # 单词级对比
            self.highlight_word_differences(self.text1, self.text1_content, self.text2_content, is_text1=True)
            self.highlight_word_differences(self.text2, self.text2_content, self.text1_content, is_text1=False)
            messagebox.showinfo("完成", "单词级对比完成，差异内容已高亮显示！")

    def set_bold(self):
        active_text = self.focus_get()
        if active_text not in [self.text1, self.text2]:
            messagebox.showinfo("提示", "请先选中要编辑的文本！")
            return
        try:
            start = active_text.index("sel.first")
            end = active_text.index("sel.last")
            if "bold" in active_text.tag_names(start):
                active_text.tag_remove("bold", start, end)
            else:
                active_text.tag_add("bold", start, end)
        except tk.TclError:
            messagebox.showinfo("提示", "请先选中要加粗的文字！")

    def set_color(self, color_hex):
        active_text = self.focus_get()
        if active_text not in [self.text1, self.text2]:
            messagebox.showinfo("提示", "请先选中要编辑的文本！")
            return
        try:
            start = active_text.index("sel.first")
            end = active_text.index("sel.last")
            active_text.tag_config("color", foreground=color_hex)
            active_text.tag_add("color", start, end)
        except tk.TclError:
            messagebox.showinfo("提示", "请先选中要改色的文字！")

    def reset_style(self):
        for w in [self.text1, self.text2]:
            w.tag_remove("bold", "1.0", tk.END)
            w.tag_remove("color", "1.0", tk.END)
            w.tag_remove("diff_add", "1.0", tk.END)
            w.tag_remove("diff_remove", "1.0", tk.END)
            w.tag_remove("diff_change", "1.0", tk.END)
            w.tag_remove("diff", "1.0", tk.END)
        messagebox.showinfo("提示", "文字样式和高亮已恢复默认！")


