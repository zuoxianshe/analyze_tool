# -*- coding: utf-8 -*-
from app_common import *

class AlarmMonitorWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("监控告警筛选")
        self.geometry("1080x680")
        self.configure(bg="#f4f4f4")
        self.sheets = {}
        self.result_df = pd.DataFrame() if PANDAS_AVAILABLE else None
        self.loaded_excel_path = ""
        self.grid_col_widths = {}
        self._last_click_row = ""
        self._last_click_col = ""
        self._cell_editor = None
        self._cell_editor_target = ("", "")

        top = tk.Frame(self, bg="#e8eef8", bd=2, relief=tk.RIDGE)
        top.pack(fill="x", padx=6, pady=4)
        tk.Label(top, text="监控告警筛选（Excel）", bg="#e8eef8", font=("微软雅黑", 10, "bold")).pack(anchor="w", padx=8, pady=4)

        ctrl = tk.Frame(top, bg="#e8eef8")
        ctrl.pack(fill="x", padx=8, pady=2)
        self.load_btn = tk.Button(ctrl, text="加载Excel", width=12, command=self.load_excel, bg="#2196F3", fg="white", font=("微软雅黑", 9))
        self.load_btn.pack(side="left", padx=3)
        self.filter_btn = tk.Button(ctrl, text="筛选告警", width=10, command=self.filter_alarms, bg="#4CAF50", fg="white", font=("微软雅黑", 9))
        self.filter_btn.pack(side="left", padx=3)
        self.full_btn = tk.Button(ctrl, text="输出全量告警", width=12, command=self.output_all_alarms, bg="#795548", fg="white", font=("微软雅黑", 9))
        self.full_btn.pack(side="left", padx=3)
        self.save_btn = tk.Button(ctrl, text="保存结果", width=10, command=self.save_result, bg="#607D8B", fg="white", font=("微软雅黑", 9))
        self.save_btn.pack(side="left", padx=3)

        self.file_label = tk.Label(top, text="未加载Excel（支持拖拽 .xlsx/.xls/.csv）", bg="#e8eef8", fg="#37474f", font=("微软雅黑", 9))
        self.file_label.pack(anchor="w", padx=8, pady=(2, 4))
        self.result_count_label = tk.Label(top, text="当前告警数：0", bg="#e8eef8", fg="#37474f", font=("微软雅黑", 9))
        self.result_count_label.pack(anchor="w", padx=8, pady=(0, 4))

        filters = tk.Frame(self, bg="#f4f4f4")
        filters.pack(fill="x", padx=8, pady=2)
        tk.Label(filters, text="FRU对象：", bg="#f4f4f4", font=("微软雅黑", 9)).grid(row=0, column=0, sticky="w", padx=4, pady=3)
        self.fru_entry = tk.Entry(filters, width=26, font=("Consolas", 10))
        self.fru_entry.grid(row=0, column=1, sticky="w", padx=4, pady=3)
        tk.Label(filters, text="机型包含：", bg="#f4f4f4", font=("微软雅黑", 9)).grid(row=0, column=2, sticky="w", padx=10, pady=3)
        self.model_entry = tk.Entry(filters, width=26, font=("Consolas", 10))
        self.model_entry.grid(row=0, column=3, sticky="w", padx=4, pady=3)
        tk.Label(filters, text="仅限机型：", bg="#f4f4f4", font=("微软雅黑", 9)).grid(row=1, column=0, sticky="w", padx=4, pady=3)
        self.model_only_entry = tk.Entry(filters, width=26, font=("Consolas", 10))
        self.model_only_entry.grid(row=1, column=1, sticky="w", padx=4, pady=3)
        tk.Label(filters, text="（仅限机型：该行机型若含其他机型将被排除）", bg="#f4f4f4", fg="gray", font=("微软雅黑", 8)).grid(
            row=1, column=2, columnspan=2, sticky="w", padx=10, pady=3
        )

        result_frame = tk.Frame(self, bg="#f4f4f4")
        result_frame.pack(fill="both", expand=True, padx=8, pady=6)
        self.result_tree = ttk.Treeview(result_frame, show="headings")
        self.result_tree.pack(side="left", fill="both", expand=True)
        self.result_ybar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        self.result_ybar.pack(side="right", fill="y")
        self.result_tree.configure(yscrollcommand=self.result_ybar.set)
        self.result_xbar = ttk.Scrollbar(self, orient="horizontal", command=self.result_tree.xview)
        self.result_xbar.pack(fill="x", padx=8, pady=(0, 6))
        self.result_tree.configure(xscrollcommand=self.result_xbar.set)
        self.result_tree.bind("<ButtonRelease-1>", self.on_tree_click_release)
        self.result_tree.bind("<Double-1>", self.on_tree_double_click_edit)
        self.result_tree.bind("<Control-c>", self.copy_selected_cell_or_row)

        self.bind("<Control-s>", self.save_result)
        self.result_tree.bind("<Control-s>", self.save_result)

        self.drop_target_register(tkdnd.DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop_excel)

    def load_excel(self):
        if not PANDAS_AVAILABLE:
            messagebox.showerror("错误", "未安装pandas，无法读取Excel。请安装：pip install pandas openpyxl")
            return
        path = filedialog.askopenfilename(
            title="选择Excel/CSV文件",
            filetypes=[("表格文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        )
        if path:
            self._load_excel_path(path)

    def on_drop_excel(self, event):
        path = event.data
        if "} {" in path:
            path = path.split("} {")[0]
        path = normalize_input_path(path)
        self._load_excel_path(path)

    def _load_excel_path(self, path):
        path = normalize_input_path(path)
        ext = os.path.splitext(path)[1].lower()
        if ext not in [".xlsx", ".xls", ".csv"]:
            messagebox.showwarning("提示", "仅支持.xlsx/.xls/.csv")
            return
        if not PANDAS_AVAILABLE:
            messagebox.showerror("错误", "未安装pandas，无法读取表格文件")
            return
        try:
            if ext == ".csv":
                self.sheets = {"CSV": pd.read_csv(path, dtype=str).fillna("")}
            else:
                all_sheets = pd.read_excel(path, sheet_name=None, dtype=str)
                self.sheets = {name: df.fillna("") for name, df in all_sheets.items()}
            self.loaded_excel_path = path
            total_rows = sum(len(df) for df in self.sheets.values())
            self.file_label.config(text=f"已加载：{path} | Sheet数：{len(self.sheets)} | 总行数：{total_rows}")
            self.result_df = pd.DataFrame()
            self._render_df_to_grid(self.result_df)
        except Exception as e:
            messagebox.showerror("错误", f"加载失败：{str(e)}")

    def _pick_col(self, df, candidates):
        normalized = {str(c).replace(" ", ""): c for c in df.columns}
        for name in candidates:
            key = name.replace(" ", "")
            if key in normalized:
                return normalized[key]
        return None

    def _row_text(self, row):
        vals = []
        for v in row.values:
            s = str(v).strip()
            if s:
                vals.append(s)
        return " | ".join(vals)

    def _is_only_model(self, model_text, target):
        target_raw = str(target).strip()
        if not target_raw:
            return True

        # 统一归一化，避免“包含可命中、仅限不命中”的空格/大小写差异
        def _norm(s):
            return re.sub(r"\s+", "", str(s)).lower()

        text = _norm(model_text)
        target_norm = _norm(target_raw)
        if not target_norm:
            return True
        if target_norm not in text:
            return False

        # 支持常见分隔符；若出现多个机型，只允许全部都等于目标机型
        if re.search(r"[/、,，;；|]+", text):
            parts = [p for p in re.split(r"[/、,，;；|]+", text) if p]
            if not parts:
                return False
            uniq = set(parts)
            return len(uniq) == 1 and target_norm in uniq

        # 无分隔符时保留“单机型字段可包含附加说明”的兼容行为
        return target_norm in text

    def _extract_model_text_for_filter(self, row, model_col=None):
        target_fields = {"支持产品列表", "芯片规划部署形态"}

        # 1) 明确的列头优先（若列名符合目标字段）
        if model_col and str(model_col).strip().replace(" ", "") in target_fields:
            return str(row.get(model_col, "")).strip()

        # 2) 从单元格首行解析“字段: 值”或“字段：值”
        values = [str(v).strip() for v in row.values]
        for i, cell in enumerate(values):
            if not cell:
                continue
            lines = [ln.strip() for ln in cell.splitlines() if ln.strip()]
            if not lines:
                continue
            first_line = lines[0]
            key = first_line
            val = ""
            if "：" in first_line:
                key, val = first_line.split("：", 1)
            elif ":" in first_line:
                key, val = first_line.split(":", 1)
            key = str(key).strip().replace(" ", "")
            val = str(val).strip()
            if key not in target_fields:
                continue
            if val:
                return val
            if i + 1 < len(values) and values[i + 1]:
                return values[i + 1]
            if len(lines) > 1:
                return "\n".join(lines[1:]).strip()
            return ""

        return ""

    def _collect_alarm_df(self, df, fru_kw="", model_kw="", model_only_kw=""):
        fru_col = self._pick_col(df, ["FRU对象", "FRU"])
        model_col = self._pick_col(df, ["支持产品列表", "芯片规划部署形态", "支持产品", "产品列表"])

        if df is None or len(df) == 0:
            return df.iloc[0:0].copy()

        out_rows = []
        for _, row in df.iterrows():
            fru_val = str(row.get(fru_col, "")).strip() if fru_col else ""
            model_val = self._extract_model_text_for_filter(row, model_col=model_col)

            if fru_kw:
                if not fru_col:
                    continue
                if fru_kw.lower() not in fru_val.lower():
                    continue

            if model_kw:
                if not model_val:
                    continue
                if model_kw.lower() not in model_val.lower():
                    continue

            if model_only_kw:
                if not model_val:
                    continue
                if not self._is_only_model(model_val, model_only_kw):
                    continue

            out_rows.append(row)
        if not out_rows:
            return df.iloc[0:0].copy()
        return pd.DataFrame(out_rows, columns=df.columns).fillna("")

    def _merge_all_sheets(self):
        if not self.sheets:
            return pd.DataFrame()
        dfs = [df.copy() for df in self.sheets.values()]
        merged = pd.concat(dfs, ignore_index=True, sort=False).fillna("")
        return merged

    def _sort_treeview_by_col(self, col_name, descending=False):
        rows = [(self.result_tree.set(k, col_name), k) for k in self.result_tree.get_children("")]
        def key_fn(item):
            v = item[0]
            try:
                return (0, float(str(v).replace(",", "")))
            except Exception:
                return (1, str(v).lower())
        rows.sort(key=key_fn, reverse=descending)
        for idx, (_, k) in enumerate(rows):
            self.result_tree.move(k, "", idx)
        self.result_tree.heading(col_name, command=lambda: self._sort_treeview_by_col(col_name, not descending))

    def on_tree_click_release(self, event):
        row_id = self.result_tree.identify_row(event.y)
        col_id = self.result_tree.identify_column(event.x)
        self._last_click_row = row_id
        self._last_click_col = col_id
        for col in self.result_tree["columns"]:
            self.grid_col_widths[col] = self.result_tree.column(col, "width")

    def copy_selected_cell_or_row(self, event=None):
        row_id = self._last_click_row or self.result_tree.focus()
        if not row_id:
            return "break"
        values = list(self.result_tree.item(row_id, "values"))
        cols = list(self.result_tree["columns"])
        text = ""
        if self._last_click_col and self._last_click_col.startswith("#"):
            try:
                idx = int(self._last_click_col[1:]) - 1
                if 0 <= idx < len(values):
                    text = str(values[idx])
            except Exception:
                text = ""
        if not text:
            text = "\t".join([str(v) for v in values])
        self.clipboard_clear()
        self.clipboard_append(text)
        return "break"

    def on_tree_double_click_edit(self, event):
        row_id = self.result_tree.identify_row(event.y)
        col_id = self.result_tree.identify_column(event.x)
        if not row_id or not col_id or not col_id.startswith("#"):
            return
        try:
            col_index = int(col_id[1:]) - 1
        except Exception:
            return
        cols = list(self.result_tree["columns"])
        if not (0 <= col_index < len(cols)):
            return
        bbox = self.result_tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        value = self.result_tree.set(row_id, cols[col_index])
        if self._cell_editor is not None:
            self._cell_editor.destroy()
            self._cell_editor = None
        self._cell_editor_target = (row_id, cols[col_index])
        self._cell_editor = tk.Entry(self.result_tree, font=("Consolas", 10))
        self._cell_editor.place(x=x, y=y, width=w, height=h)
        self._cell_editor.insert(0, value)
        self._cell_editor.focus_set()
        self._cell_editor.bind("<Return>", self.commit_cell_edit)
        self._cell_editor.bind("<Escape>", self.cancel_cell_edit)
        self._cell_editor.bind("<FocusOut>", self.commit_cell_edit)

    def commit_cell_edit(self, event=None):
        if self._cell_editor is None:
            return
        row_id, col_name = self._cell_editor_target
        new_val = self._cell_editor.get()
        if row_id and col_name:
            self.result_tree.set(row_id, col_name, new_val)
        self._cell_editor.destroy()
        self._cell_editor = None
        self._cell_editor_target = ("", "")

    def cancel_cell_edit(self, event=None):
        if self._cell_editor is not None:
            self._cell_editor.destroy()
            self._cell_editor = None
        self._cell_editor_target = ("", "")

    def _get_treeview_df(self):
        cols = list(self.result_tree["columns"])
        rows = []
        for item in self.result_tree.get_children(""):
            values = list(self.result_tree.item(item, "values"))
            if len(values) < len(cols):
                values += [""] * (len(cols) - len(values))
            rows.append(values[:len(cols)])
        if not cols:
            return pd.DataFrame()
        return pd.DataFrame(rows, columns=cols)

    def _render_df_to_grid(self, df):
        self.cancel_cell_edit()
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        cols = list(df.columns) if df is not None else []
        self.result_tree["columns"] = cols
        count = len(df) if df is not None else 0
        self.result_count_label.config(text=f"当前告警数：{count}")
        if not cols:
            self.result_tree["displaycolumns"] = ()
            return
        self.result_tree["displaycolumns"] = cols
        for col in cols:
            self.result_tree.heading(col, text=str(col), command=lambda c=col: self._sort_treeview_by_col(c, False))
            width = self.grid_col_widths.get(col, 140)
            self.result_tree.column(col, width=width, anchor="w", stretch=True)
        for _, row in df.iterrows():
            values = [str(row.get(c, "")) for c in cols]
            self.result_tree.insert("", "end", values=values)

    def filter_alarms(self):
        if not self.sheets:
            messagebox.showwarning("提示", "请先加载Excel文件")
            return

        fru_kw = self.fru_entry.get().strip()
        model_kw = self.model_entry.get().strip()
        model_only_kw = self.model_only_entry.get().strip()
        if not (fru_kw or model_kw or model_only_kw):
            messagebox.showinfo("提示", "请至少输入一个筛选条件")
            return
        merged = self._merge_all_sheets()
        self.result_df = self._collect_alarm_df(merged, fru_kw=fru_kw, model_kw=model_kw, model_only_kw=model_only_kw)
        self._render_df_to_grid(self.result_df)

    def output_all_alarms(self):
        if not self.sheets:
            messagebox.showwarning("提示", "请先加载Excel文件")
            return
        self.result_df = self._merge_all_sheets()
        self._render_df_to_grid(self.result_df)

    def save_result(self, event=None):
        if not PANDAS_AVAILABLE:
            messagebox.showerror("错误", "未安装pandas，无法保存为Excel")
            return "break" if event else None
        if self.result_df is None or len(self.result_df) == 0:
            messagebox.showwarning("提示", "当前没有可保存的筛选结果")
            return "break" if event else None
        path = filedialog.asksaveasfilename(
            title="保存筛选结果",
            defaultextension=".xlsx",
            initialfile="alarm_result.xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not path:
            return "break" if event else None
        try:
            current_df = self._get_treeview_df()
            self.result_df = current_df.copy()
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                current_df.to_excel(writer, sheet_name="告警结果", index=False)
            messagebox.showinfo("提示", f"已保存：\n{path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")
        return "break" if event else None

# ==============================
# 主文件查看器（原有功能完全保留）
# ==============================
