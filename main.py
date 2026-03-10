import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk, colorchooser
import tkinter.font as tkfont
import tkinterdnd2 as tkdnd
import zipfile
import tarfile
import gzip
import bz2
import os
import io
import locale
import re
import xml.etree.ElementTree as ET
import subprocess
import tempfile
import shutil
from urllib.parse import unquote
from difflib import Differ, SequenceMatcher
import uuid
import json
import threading

# 新增依赖：处理Excel和CSV
try:
    import pandas as pd

    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

try:
    import pdfplumber

    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    from pypdf import PdfReader

    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

try:
    from PIL import Image, ImageTk

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import win32com.client as win32

    WIN32COM_AVAILABLE = True
except Exception:
    WIN32COM_AVAILABLE = False

# 修复中文排序
try:
    locale.setlocale(locale.LC_ALL, '')
except locale.Error:
    locale.setlocale(locale.LC_ALL, 'C')

LARGE_FILE_THRESHOLD = 8 * 1024 * 1024
PREVIEW_READ_BYTES = 256 * 1024
PDF_LARGE_MAX_PAGES = 10
TABULAR_PREVIEW_ROWS = 3000
ARCHIVE_MEMBER_PREVIEW_BYTES = 2 * 1024 * 1024
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}


def read_text_preview_from_path(file_path, max_bytes=PREVIEW_READ_BYTES):
    try:
        with open(file_path, "rb") as f:
            data = f.read(max_bytes + 1)
        truncated = len(data) > max_bytes
        text = data[:max_bytes].decode("utf-8", errors="replace")
        if truncated:
            text += f"\n\n[预览模式] 文件较大，仅显示前 {max_bytes // 1024}KB。"
        return text
    except Exception as e:
        return f"(预览读取失败: {str(e)})"


def read_text_preview_from_bytes(file_data, max_bytes=PREVIEW_READ_BYTES):
    data = file_data[: max_bytes + 1]
    truncated = len(data) > max_bytes
    text = data[:max_bytes].decode("utf-8", errors="replace")
    if truncated:
        text += f"\n\n[预览模式] 文件较大，仅显示前 {max_bytes // 1024}KB。"
    return text


def extract_pdf_text(file_path=None, file_bytes=None, display_name="", max_pages=None):
    """Extract text from PDF using pdfplumber first, fallback to pypdf."""
    if not PDFPLUMBER_AVAILABLE and not PYPDF_AVAILABLE:
        return "未安装 PDF 解析依赖。请安装：pip install pdfplumber pypdf"

    if file_bytes is None and not file_path:
        return "(PDF路径无效)"

    def _format_pdf_text(text):
        title = display_name or (os.path.basename(file_path) if file_path else "PDF")
        content = (text or "").strip()
        if not content:
            return f"=== PDF文件：{title} ===\n(未提取到可读文本，可能为扫描版图片PDF)"
        return f"=== PDF文件：{title} ===\n\n{content}"

    try:
        if PDFPLUMBER_AVAILABLE:
            pages_text = []
            if file_bytes is not None:
                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                    pages = pdf.pages[:max_pages] if max_pages else pdf.pages
                    for page in pages:
                        pages_text.append(page.extract_text() or "")
            else:
                with pdfplumber.open(file_path) as pdf:
                    pages = pdf.pages[:max_pages] if max_pages else pdf.pages
                    for page in pages:
                        pages_text.append(page.extract_text() or "")
            return _format_pdf_text("\n\n".join(pages_text))
    except Exception:
        pass

    try:
        if PYPDF_AVAILABLE:
            if file_bytes is not None:
                reader = PdfReader(io.BytesIO(file_bytes))
            else:
                reader = PdfReader(file_path)
            pages_text = []
            pages = reader.pages[:max_pages] if max_pages else reader.pages
            for page in pages:
                pages_text.append(page.extract_text() or "")
            return _format_pdf_text("\n\n".join(pages_text))
    except Exception as e:
        return f"(PDF解析失败: {str(e)})"

    return "(PDF解析失败)"


def get_pdf_page_count(file_path=None, file_bytes=None):
    try:
        if PDFPLUMBER_AVAILABLE:
            if file_bytes is not None:
                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                    return len(pdf.pages)
            with pdfplumber.open(file_path) as pdf:
                return len(pdf.pages)
        if PYPDF_AVAILABLE:
            reader = PdfReader(io.BytesIO(file_bytes)) if file_bytes is not None else PdfReader(file_path)
            return len(reader.pages)
    except Exception:
        return 0
    return 0


def extract_pdf_single_page(file_path=None, file_bytes=None, page_no=1, display_name=""):
    """Extract one PDF page with text/table/image metadata."""
    page_no = max(1, int(page_no))
    title = display_name or (os.path.basename(file_path) if file_path else "PDF")
    try:
        if PDFPLUMBER_AVAILABLE:
            if file_bytes is not None:
                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                    total = len(pdf.pages)
                    if total <= 0:
                        return "(PDF无页面)", 0
                    idx = min(page_no, total) - 1
                    page = pdf.pages[idx]
                    text = (page.extract_text() or "").strip()
                    tables = page.extract_tables() or []
                    images = page.images or []
            else:
                with pdfplumber.open(file_path) as pdf:
                    total = len(pdf.pages)
                    if total <= 0:
                        return "(PDF无页面)", 0
                    idx = min(page_no, total) - 1
                    page = pdf.pages[idx]
                    text = (page.extract_text() or "").strip()
                    tables = page.extract_tables() or []
                    images = page.images or []

            parts = [f"=== PDF文件：{title} | 第 {idx + 1}/{total} 页 ==="]
            if text:
                parts.append("")
                parts.append(text)
            if tables:
                for t_idx, table in enumerate(tables, start=1):
                    parts.append("")
                    parts.append(f"-- 表格 {t_idx} --")
                    for row in table:
                        safe_row = [(c if c is not None else "") for c in row]
                        parts.append("\t".join(safe_row))

            return "\n".join(parts), total

        if PYPDF_AVAILABLE:
            reader = PdfReader(io.BytesIO(file_bytes)) if file_bytes is not None else PdfReader(file_path)
            total = len(reader.pages)
            if total <= 0:
                return "(PDF无页面)", 0
            idx = min(page_no, total) - 1
            text = (reader.pages[idx].extract_text() or "").strip()
            content = f"=== PDF文件：{title} | 第 {idx + 1}/{total} 页 ==="
            if text:
                content += f"\n\n{text}"
            return content, total
    except Exception as e:
        return f"(PDF单页解析失败: {str(e)})", 0

    return "(未安装PDF解析依赖)", 0


def extract_pdf_single_page_image_bytes(file_path=None, file_bytes=None, page_no=1):
    """Render one PDF page to PNG bytes for visual display."""
    if not PDFPLUMBER_AVAILABLE:
        return None
    try:
        if file_bytes is not None:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                total = len(pdf.pages)
                if total <= 0:
                    return None
                idx = min(max(1, int(page_no)), total) - 1
                page_img = pdf.pages[idx].to_image(resolution=110).original
        else:
            with pdfplumber.open(file_path) as pdf:
                total = len(pdf.pages)
                if total <= 0:
                    return None
                idx = min(max(1, int(page_no)), total) - 1
                page_img = pdf.pages[idx].to_image(resolution=110).original

        buf = io.BytesIO()
        page_img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None


def extract_docx_text(file_path):
    try:
        with zipfile.ZipFile(file_path) as zf:
            xml_bytes = zf.read("word/document.xml")
        root = ET.fromstring(xml_bytes)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        lines = []
        for para in root.findall(".//w:p", ns):
            texts = []
            for t in para.findall(".//w:t", ns):
                texts.append(t.text or "")
            line = "".join(texts).strip()
            if line:
                lines.append(line)
        text = "\n".join(lines).strip()
        return text if text else "(DOCX中未提取到文本)"
    except Exception as e:
        return f"(DOCX解析失败: {str(e)})"


def extract_doc_text(file_path):
    file_path = normalize_input_path(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        return extract_docx_text(file_path)
    if ext == ".doc":
        # 1) Try Word COM first
        if WIN32COM_AVAILABLE:
            word = None
            doc = None
            try:
                word = win32.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(file_path, ReadOnly=True)
                text = doc.Content.Text or ""
                return text.strip() if text.strip() else "(DOC中未提取到文本)"
            except Exception as e:
                return f"(DOC解析失败: {str(e)})"
            finally:
                try:
                    if doc is not None:
                        doc.Close(False)
                except Exception:
                    pass
                try:
                    if word is not None:
                        word.Quit()
                except Exception:
                    pass
        # 2) Fallback: try Word COM import at runtime (pywin32 may be installed after startup)
        try:
            import win32com.client as _win32
            word = None
            doc = None
            try:
                word = _win32.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(file_path, ReadOnly=True)
                text = doc.Content.Text or ""
                return text.strip() if text.strip() else "(DOC中未提取到文本)"
            finally:
                try:
                    if doc is not None:
                        doc.Close(False)
                except Exception:
                    pass
                try:
                    if word is not None:
                        word.Quit()
                except Exception:
                    pass
        except Exception:
            pass

        # 3) Fallback: LibreOffice soffice headless convert .doc -> .txt
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            tmpdir = tempfile.mkdtemp(prefix="doc_parse_")
            try:
                subprocess.run(
                    [soffice, "--headless", "--convert-to", "txt:Text", "--outdir", tmpdir, file_path],
                    check=False,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=60
                )
                base = os.path.splitext(os.path.basename(file_path))[0]
                txt_path = os.path.join(tmpdir, base + ".txt")
                if os.path.exists(txt_path):
                    with open(txt_path, "r", encoding="utf-8", errors="replace") as f:
                        text = f.read()
                    return text.strip() if text.strip() else "(DOC中未提取到文本)"
            except Exception:
                pass
            finally:
                try:
                    shutil.rmtree(tmpdir, ignore_errors=True)
                except Exception:
                    pass

        return "(DOC解析失败：请安装 Microsoft Word+pywin32 或 LibreOffice（soffice）)"
    return "(不支持的文档格式)"


def normalize_input_path(path):
    if path is None:
        return ""
    p = str(path).strip().strip('"').strip("'")
    if p.startswith("{") and p.endswith("}"):
        p = p[1:-1]
    p = unquote(p)
    p = p.replace("/", os.sep).replace("\\", os.sep)
    p = os.path.normpath(p)
    return p


# ==============================
# 文本比较窗口（增强版：智能差异高亮）
# ==============================
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
        self.color_btn = tk.Menubutton(edit_inner, text="🎨 颜色", width=7, font=("微软雅黑", 8), relief=tk.RAISED,
                                       bg="#f0f0f0")
        self.color_btn.pack(side="left", padx=5)
        self.color_menu = tk.Menu(self.color_btn, tearoff=0, bg="white", bd=1)
        self.color_btn.config(menu=self.color_menu)
        self.build_color_menu()
        self.reset_btn = tk.Button(edit_inner, text="🔄 恢复默认", width=9, command=self.reset_style,
                                   font=("微软雅黑", 8))
        self.reset_btn.pack(side="left", padx=5)
        self.save_btn = tk.Button(edit_inner, text="保存文本", width=9, command=self.save_active_text,
                                  font=("微软雅黑", 8))
        self.save_btn.pack(side="left", padx=5)

        # 对比区域
        self.paned_main = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief=tk.RIDGE, bg="#f4f4f4")
        self.paned_main.pack(fill="both", expand=True, padx=5, pady=3)

        # 文本框1
        self.frame_text1 = tk.Frame(self.paned_main, bg="#f8f8f8")
        tk.Label(self.frame_text1, text="文件1内容", bg="#f8f8f8", font=("微软雅黑", 10, "bold")).pack(fill="x")
        self.text1 = scrolledtext.ScrolledText(self.frame_text1, font=self.compare_font, wrap="word", bg="white")
        self.text1.pack(fill="both", expand=True, padx=2, pady=2)
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
        self.text2 = scrolledtext.ScrolledText(self.frame_text2, font=self.compare_font, wrap="word", bg="white")
        self.text2.pack(fill="both", expand=True, padx=2, pady=2)
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
        self.text1.tag_config("diff_add", background="#E6FFE6")  # 新增内容 - 浅绿
        self.text1.tag_config("diff_remove", background="#FFE6E6")  # 删除内容 - 浅红
        self.text1.tag_config("diff_change", background="#FFFFE6")  # 修改内容 - 浅黄
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
        initial_name = os.path.basename(file_path) if file_path else (
            "compare_text_1.txt" if active_text == self.text1 else "compare_text_2.txt")
        defaultextension = file_format if file_format else ".txt"
        save_path = filedialog.asksaveasfilename(
            title="保存文本内容",
            initialfile=initial_name,
            defaultextension=defaultextension,
            filetypes=[("文本文件", "*.txt"), ("JSON", "*.json"), ("CSV", "*.csv"), ("Python", "*.py"),
                       ("所有文件", "*.*")]
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
                hex_color = ''.join([c * 2 for c in hex_color])
            if len(hex_color) != 6:
                return False
            r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
            return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5
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
                        f"=== 大PDF预览：{os.path.basename(file_path)} ({file_size / 1024 / 1024:.2f}MB) ===\n"
                        f"[性能模式] 仅提取前{PDF_LARGE_MAX_PAGES}页文本。\n\n{content}",
                        file_ext
                    )
                preview = read_text_preview_from_path(file_path)
                return f"=== 大文件预览：{os.path.basename(file_path)} ({file_size / 1024 / 1024:.2f}MB) ===\n\n{preview}", file_ext

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
                    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                        return f.read(), file_ext
                try:
                    df = pd.read_csv(file_path, nrows=TABULAR_PREVIEW_ROWS)
                    return df.to_string(index=False), file_ext
                except Exception as e:
                    # 降级读取
                    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                        return f.read(), file_ext

            # JSON文件
            elif file_ext == '.json':
                with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                    content = f.read()
                # 格式化JSON以便对比
                return self.parse_json(content), file_ext
            elif file_ext in ['.doc', '.docx']:
                return extract_doc_text(file_path), file_ext
            elif file_ext == '.pdf':
                return extract_pdf_text(file_path=file_path), file_ext

            # 普通文本文件 (txt/xml/md/py等)
            else:
                with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                    return f.read(), file_ext

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
                    self.text1.tag_add("diff_remove", f"{line1_idx + 1}.0", f"{line1_idx + 1}.end")
                    line1_idx += 1
            elif line.startswith('+ '):
                # 文件2独有的行（新增）
                if line2_idx < len(lines2):
                    self.text2.tag_add("diff_add", f"{line2_idx + 1}.0", f"{line2_idx + 1}.end")
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
                    self.text1.tag_add("diff_change", f"{line1_idx + 1}.0", f"{line1_idx + 1}.end")
                    line1_idx += 1
                if line2_idx < len(lines2):
                    self.text2.tag_add("diff_change", f"{line2_idx + 1}.0", f"{line2_idx + 1}.end")
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
        tk.Label(top, text="监控告警筛选（Excel）", bg="#e8eef8", font=("微软雅黑", 10, "bold")).pack(anchor="w", padx=8,
                                                                                                    pady=4)

        ctrl = tk.Frame(top, bg="#e8eef8")
        ctrl.pack(fill="x", padx=8, pady=2)
        self.load_btn = tk.Button(ctrl, text="加载Excel", width=12, command=self.load_excel, bg="#2196F3", fg="white",
                                  font=("微软雅黑", 9))
        self.load_btn.pack(side="left", padx=3)
        self.filter_btn = tk.Button(ctrl, text="筛选告警", width=10, command=self.filter_alarms, bg="#4CAF50",
                                    fg="white", font=("微软雅黑", 9))
        self.filter_btn.pack(side="left", padx=3)
        self.full_btn = tk.Button(ctrl, text="输出全量告警", width=12, command=self.output_all_alarms, bg="#795548",
                                  fg="white", font=("微软雅黑", 9))
        self.full_btn.pack(side="left", padx=3)
        self.save_btn = tk.Button(ctrl, text="保存结果", width=10, command=self.save_result, bg="#607D8B", fg="white",
                                  font=("微软雅黑", 9))
        self.save_btn.pack(side="left", padx=3)

        self.file_label = tk.Label(top, text="未加载Excel（支持拖拽 .xlsx/.xls/.csv）", bg="#e8eef8", fg="#37474f",
                                   font=("微软雅黑", 9))
        self.file_label.pack(anchor="w", padx=8, pady=(2, 4))
        self.result_count_label = tk.Label(top, text="当前告警数：0", bg="#e8eef8", fg="#37474f", font=("微软雅黑", 9))
        self.result_count_label.pack(anchor="w", padx=8, pady=(0, 4))

        filters = tk.Frame(self, bg="#f4f4f4")
        filters.pack(fill="x", padx=8, pady=2)
        tk.Label(filters, text="FRU对象：", bg="#f4f4f4", font=("微软雅黑", 9)).grid(row=0, column=0, sticky="w", padx=4,
                                                                                    pady=3)
        self.fru_entry = tk.Entry(filters, width=26, font=("Consolas", 10))
        self.fru_entry.grid(row=0, column=1, sticky="w", padx=4, pady=3)
        tk.Label(filters, text="机型包含：", bg="#f4f4f4", font=("微软雅黑", 9)).grid(row=0, column=2, sticky="w",
                                                                                     padx=10, pady=3)
        self.model_entry = tk.Entry(filters, width=26, font=("Consolas", 10))
        self.model_entry.grid(row=0, column=3, sticky="w", padx=4, pady=3)
        tk.Label(filters, text="仅限机型：", bg="#f4f4f4", font=("微软雅黑", 9)).grid(row=1, column=0, sticky="w",
                                                                                     padx=4, pady=3)
        self.model_only_entry = tk.Entry(filters, width=26, font=("Consolas", 10))
        self.model_only_entry.grid(row=1, column=1, sticky="w", padx=4, pady=3)
        tk.Label(filters, text="（仅限机型：该行机型若含其他机型将被排除）", bg="#f4f4f4", fg="gray",
                 font=("微软雅黑", 8)).grid(
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
class FileViewerApp(tkdnd.Tk):
    def __init__(self):
        super().__init__()
        self.title("文件查看器｜文件夹优先+字母序")
        self.geometry("1200x750")
        self.configure(bg="#f4f4f4")

        # 核心变量
        self.current_path = None
        self.max_file_size = 10 * 1024 * 1024  # 10MB
        self.node_data = {}  # 节点数据
        self.node_type = {}  # 节点类型
        self.parent_node_files = {}  # 去重映射
        self.node_full_path = {}  # 节点完整路径
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
        self.line_num_font = tkfont.Font(family="Consolas", size=max(8, self.editor_font_size - 1))

        # 搜索相关变量（分开存储避免冲突）
        self.file_search_hits = []  # 文件名搜索结果 [(节点ID, 显示名称)]
        self.content_search_hits = []  # 内容搜索结果 [(行号, 列号, 内容)]
        self.multi_content_search_hits = []  # 多文件内容搜索结果 [(节点ID, 行号, 列号, 行内容)]
        self.preview_line_to_hit_index = []  # 兼容旧逻辑（保留）
        self.preview_jump_entries = []  # 预览框行号到跳转目标映射（统一）
        self.current_search_type = ""  # "file" 或 "content"
        self.search_text_cache = {}  # 多文件搜索文本缓存：node_id -> {"key": ..., "lines": [...]}

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
        tk.Label(file_group, text="📁 文件操作", bg="#e8eef8", font=("微软雅黑", 8, "bold")).pack(anchor="w",
                                                                                                 pady=(0, 2))
        file_btn_row = tk.Frame(file_group, bg="#e8eef8")
        file_btn_row.pack(anchor="w")

        self.folder_archive_btn = tk.Button(file_btn_row, text="📂 文件夹/压缩包", width=14,
                                            command=self.open_folder_or_archive,
                                            font=("微软雅黑", 9), bg="#2196F3", fg="white")
        self.folder_archive_btn.pack(side="left", padx=3)
        self.file_btn = tk.Button(file_btn_row, text="📄 选择文件", width=14, command=self.open_file,
                                  font=("微软雅黑", 9), bg="#4CAF50", fg="white")
        self.file_btn.pack(side="left", padx=3)
        self.compare_btn = tk.Button(file_btn_row, text="🆚 文本比较", width=14, command=self.open_compare_window,
                                     font=("微软雅黑", 9), bg="#9C27B0", fg="white")
        self.compare_btn.pack(side="left", padx=3)
        self.monitor_btn = tk.Button(file_btn_row, text="🚨 监控告警", width=12, command=self.open_alarm_monitor_window,
                                     font=("微软雅黑", 9), bg="#FF7043", fg="white")
        self.monitor_btn.pack(side="left", padx=3)

        sep = tk.Frame(tools_inner, width=1, bg="#b9c7de")
        sep.grid(row=0, column=1, sticky="ns", padx=6)

        edit_group = tk.Frame(tools_inner, bg="#e8eef8")
        edit_group.grid(row=0, column=2, sticky="ew", padx=(8, 2))
        tk.Label(edit_group, text="✏️ 文字编辑", bg="#e8eef8", font=("微软雅黑", 8, "bold")).pack(anchor="w",
                                                                                                  pady=(0, 2))
        edit_btn_row = tk.Frame(edit_group, bg="#e8eef8")
        edit_btn_row.pack(anchor="w")

        self.bold_btn = tk.Button(edit_btn_row, text="𝐁 加粗", width=7, command=self.set_bold, font=("微软雅黑", 8))
        self.bold_btn.pack(side="left", padx=5)
        self.color_btn = tk.Menubutton(edit_btn_row, text="🎨 颜色", width=7, font=("微软雅黑", 8), relief=tk.RAISED,
                                       bg="#f0f0f0")
        self.color_btn.pack(side="left", padx=5)
        self.color_menu = tk.Menu(self.color_btn, tearoff=0, bg="white", bd=1)
        self.color_btn.config(menu=self.color_menu)
        self.build_word_style_color_menu()
        self.reset_btn = tk.Button(edit_btn_row, text="🔄 恢复默认", width=9, command=self.reset_style,
                                   font=("微软雅黑", 8))
        self.reset_btn.pack(side="left", padx=5)
        self.md_render_btn = tk.Button(edit_btn_row, text="📝 Markdown渲染", width=12,
                                       command=self.render_current_text_as_markdown, font=("微软雅黑", 8))
        self.md_render_btn.pack(side="left", padx=5)
        self.save_btn = tk.Button(edit_btn_row, text="保存文本", width=9, command=self.save_current_text,
                                  font=("微软雅黑", 8))
        self.save_btn.pack(side="left", padx=5)

        # ========== 搜索栏 ==========
        search_frame = tk.Frame(self, bg="#f4f4f4")
        search_frame.pack(fill="x", pady=2, padx=10)

        tk.Label(search_frame, text="智能搜索：", bg="#f4f4f4", font=("微软雅黑", 9)).pack(side="left", padx=2)
        self.search_entry = tk.Entry(search_frame, width=50, font=("Consolas", 10))
        self.search_entry.pack(side="left", padx=8, fill="x", expand=True)
        self.search_entry.bind("<Return>", lambda e: self.smart_search())

        self.search_btn = tk.Button(search_frame, text="🔍 搜索", width=8, command=self.smart_search, bg="#2196F3",
                                    fg="white")
        self.search_btn.pack(side="left", padx=4)
        self.search_multi_btn = tk.Button(search_frame, text="📚 多文件", width=8, command=self.search_content_multi,
                                          bg="#009688", fg="white")
        self.search_multi_btn.pack(side="left", padx=4)
        self.clear_btn = tk.Button(search_frame, text="🧹 清空", width=8, command=self.clear_all_highlights,
                                   bg="#f4f4f4", fg="black")
        self.clear_btn.pack(side="left", padx=4)

        self.search_status_label = tk.Label(search_frame, text="当前模式：未搜索", bg="#f4f4f4", fg="gray",
                                            font=("微软雅黑", 9))
        self.search_status_label.pack(side="left", padx=8)

        # ========== 主布局 ==========
        self.main_paned = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief=tk.RIDGE, sashwidth=6, bg="#f4f4f4")
        self.main_paned.pack(fill="both", expand=True, padx=10, pady=5)

        # 左侧：文件结构树
        self.tree_frame = tk.Frame(self.main_paned, bg="#f4f4f4")
        self.tree_label = tk.Label(self.tree_frame, text="📂 文件结构（文件夹优先+字母序）", font=("微软雅黑", 10, "bold"),
                                   bg="#e0e0e0")
        self.tree_label.pack(fill="x")
        self.tree = ttk.Treeview(self.tree_frame, show="tree")
        self.tree.pack(fill="both", expand=True, pady=4)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.tag_configure("search_hit", background="yellow")
        self.main_paned.add(self.tree_frame, width=280)

        # 右侧面板（垂直分割）
        self.right_paned = tk.PanedWindow(self.main_paned, orient=tk.VERTICAL, sashrelief=tk.RIDGE, sashwidth=6,
                                          bg="#f4f4f4")
        self.main_paned.add(self.right_paned)

        # 右上：文本编辑区
        self.text_panel = tk.Frame(self.right_paned, bg="#f4f4f4")
        self.txt_container = tk.Frame(self.text_panel, bg="#f4f4f4")
        self.txt_container.pack(fill="both", expand=True)
        self.pdf_nav_frame = tk.Frame(self.text_panel, bg="#f4f4f4", bd=0, highlightthickness=0, height=34)
        self.pdf_nav_frame.pack_propagate(False)
        self.pdf_nav_label = tk.Label(
            self.pdf_nav_frame, text="PDF分页：0/0", bg="#f4f4f4", font=("微软雅黑", 9, "bold"), bd=0,
            highlightthickness=0
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
        self.line_num = tk.Text(self.txt_container, width=5, state="disabled", bg="#f0f0f0", font=self.line_num_font)
        self.line_num.pack(side="left", fill="y")
        self.txt = scrolledtext.ScrolledText(self.txt_container, font=self.editor_font, wrap="word", bg="white")
        self.txt.pack(side="right", fill="both", expand=True)
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
        self.txt.bind("<Control-MouseWheel>", self.on_editor_zoom)
        self.txt.bind("<Control-Button-4>", self.on_editor_zoom)
        self.txt.bind("<Control-Button-5>", self.on_editor_zoom)
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
                                     font=("微软雅黑", 9, "bold"), fg="#2196F3")
        self.result_title.pack(anchor="w", padx=5, pady=2)
        self.result_txt = scrolledtext.ScrolledText(self.result_frame, font=("Consolas", 10), bg="#fffff8", height=1)
        self.result_txt.pack(fill="both", expand=True, padx=5, pady=2)
        self.result_txt.config(state=tk.DISABLED)
        # 双击跳转绑定
        self.result_txt.bind("<Double-1>", self.on_double_click_jump)
        self.right_paned.add(self.result_frame, height=38, minsize=28)

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

    # ==============================
    # 基础工具函数
    # ==============================
    def is_dark_color(self, hex_color):
        try:
            hex_color = hex_color.lstrip('#')
            if len(hex_color) == 3:
                hex_color = ''.join([c * 2 for c in hex_color])
            if len(hex_color) != 6:
                return False
            r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
            return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5
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

    def refresh_line_numbers(self):
        """刷新行号"""
        self.line_num.config(state=tk.NORMAL)
        self.line_num.delete("1.0", tk.END)
        try:
            line_count = int(self.txt.index("end-1c").split(".")[0])
            self.line_num.insert("end", "\n".join(str(i) for i in range(1, line_count + 1)))
        except:
            pass
        self.line_num.config(state=tk.DISABLED)

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
        self.line_num_font.configure(size=max(8, new_size - 1))
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
        self.refresh_line_numbers()

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

    def on_text_panel_resize(self):
        if self.pdf_lazy_source:
            self.pdf_nav_frame.place(relx=0.5, rely=1.0, y=-6, anchor="s")

    def on_text_edited(self):
        self.current_text_content = self.txt.get("1.0", "end-1c")
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
            filetypes=[("文本文件", "*.txt"), ("JSON", "*.json"), ("CSV", "*.csv"), ("Python", "*.py"),
                       ("所有文件", "*.*")]
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
                    self._add_node("", os.path.basename(file_path), file_node_id, "file",
                                   {"kind": "pdf_path", "path": file_path}, file_path)
                    self.setup_pdf_lazy({"path": file_path}, os.path.basename(file_path))
                    self.title(f"文件查看器 - {os.path.basename(file_path)}")
                    return
                if file_ext in IMAGE_EXTENSIONS:
                    self.clear_pdf_lazy_state()
                    file_node_id = f"file_{uuid.uuid4()}"
                    self._add_node("", os.path.basename(file_path), file_node_id, "file",
                                   {"kind": "image_path", "path": file_path}, file_path)
                    if PIL_AVAILABLE:
                        with Image.open(file_path) as img:
                            self.display_image_in_editor(img, title=f"=== 图片文件：{os.path.basename(file_path)} ===")
                    else:
                        self.display_text_in_editor("(未安装Pillow，无法显示图片。请安装：pip install pillow)")
                    self.title(f"文件查看器 - {os.path.basename(file_path)}")
                    return

                if file_size > LARGE_FILE_THRESHOLD:
                    if file_ext == '.pdf':
                        parsed = extract_pdf_text(file_path=file_path, max_pages=PDF_LARGE_MAX_PAGES)
                        content = (
                            f"=== 大PDF预览：{os.path.basename(file_path)} ({file_size / 1024 / 1024:.2f}MB) ===\n"
                            f"[性能模式] 仅提取前{PDF_LARGE_MAX_PAGES}页文本。\n\n{parsed}"
                        )
                    else:
                        preview = read_text_preview_from_path(file_path)
                        content = f"=== 大文件预览：{os.path.basename(file_path)} ({file_size / 1024 / 1024:.2f}MB) ===\n\n{preview}"
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
                            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                                content = f.read()
                    # JSON文件格式化
                    elif file_ext == '.json':
                        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                            json_content = f.read()
                        try:
                            json_data = json.loads(json_content)
                            content = json.dumps(json_data, indent=4, ensure_ascii=False)
                        except:
                            content = json_content
                    elif file_ext in ['.doc', '.docx']:
                        content = extract_doc_text(file_path)
                    else:
                        # 普通文本文件
                        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                            content = f.read()

                self.clear_pdf_lazy_state()
                self.current_text_content = content
                file_node_id = f"file_{uuid.uuid4()}"
                self._add_node("", os.path.basename(file_path), file_node_id, "file", content, file_path)
                self.display_text_in_editor(content, file_ext=file_ext)

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
                                member_data = f.read(min(info.file_size, ARCHIVE_MEMBER_PREVIEW_BYTES) + 1)
                            files.append((info.filename, member_data, info.file_size))
                            # 添加文件所在目录
                            file_dir = os.path.dirname(info.filename)
                            if file_dir:
                                dirs.add(file_dir)

            elif arch_type in ('tar', 'tar.gz', 'tar.bz2'):
                mode = {'tar': 'r', 'tar.gz': 'r:gz', 'tar.bz2': 'r:bz2'}[arch_type]
                with tarfile.open(fileobj=buf, mode=mode) as tf:
                    for member in tf.getmembers():
                        if member.isdir():
                            dirs.add(member.name)
                        elif member.isfile() and member.size > 0:
                            extracted = tf.extractfile(member)
                            if extracted:
                                file_data = extracted.read(min(member.size, ARCHIVE_MEMBER_PREVIEW_BYTES) + 1)
                                files.append((member.name, file_data, member.size))
                            file_dir = os.path.dirname(member.name)
                            if file_dir:
                                dirs.add(file_dir)

            elif arch_type == 'gz':
                with gzip.GzipFile(fileobj=buf) as gf:
                    new_name = archive_name[:-3] if archive_name.lower().endswith('.gz') else archive_name + '.unzip'
                    data = gf.read(ARCHIVE_MEMBER_PREVIEW_BYTES + 1)
                    files.append((new_name, data, len(data)))

            elif arch_type == 'bz2':
                with bz2.BZ2File(fileobj=buf) as bf:
                    new_name = archive_name[:-4] if archive_name.lower().endswith('.bz2') else archive_name + '.unzip'
                    data = bf.read(ARCHIVE_MEMBER_PREVIEW_BYTES + 1)
                    files.append((new_name, data, len(data)))

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
                if len(file_data) > ARCHIVE_MEMBER_PREVIEW_BYTES:
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
                self.scan_archive({"name": file_name, "data": file_data}, nest_node, level + 1)
            else:
                # 普通文件（支持多格式解析）
                try:
                    file_ext = os.path.splitext(file_name)[1].lower()
                    is_truncated_in_archive = len(file_data) > ARCHIVE_MEMBER_PREVIEW_BYTES
                    effective_data = file_data[:ARCHIVE_MEMBER_PREVIEW_BYTES] if is_truncated_in_archive else file_data
                    if original_size > LARGE_FILE_THRESHOLD or is_truncated_in_archive:
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
                                f"=== 大文件预览：{file_name} ({original_size / 1024 / 1024:.2f}MB) ===\n"
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
                                    df = pd.read_excel(excel_buf, sheet_name=sheet_name, nrows=TABULAR_PREVIEW_ROWS)
                                    content.append(f"\n--- Sheet: {sheet_name} ---")
                                    content.append(df.to_string(index=False))
                                file_content = "\n".join(content)
                            elif file_ext == '.csv':
                                csv_buf = io.StringIO(effective_data.decode('utf-8', errors='replace'))
                                df = pd.read_csv(csv_buf, nrows=TABULAR_PREVIEW_ROWS)
                                file_content = df.to_string(index=False)
                        elif file_ext == '.json':
                            # JSON格式化
                            json_content = effective_data.decode('utf-8', errors='replace')
                            try:
                                json_data = json.loads(json_content)
                                file_content = json.dumps(json_data, indent=4, ensure_ascii=False)
                            except:
                                file_content = json_content
                        elif file_ext == '.docx':
                            try:
                                tmp_docx = io.BytesIO(effective_data)
                                with zipfile.ZipFile(tmp_docx) as zf:
                                    xml_bytes = zf.read("word/document.xml")
                                root = ET.fromstring(xml_bytes)
                                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                                lines = []
                                for para in root.findall(".//w:p", ns):
                                    texts = [t.text or "" for t in para.findall(".//w:t", ns)]
                                    line = "".join(texts).strip()
                                    if line:
                                        lines.append(line)
                                file_content = "\n".join(lines) if lines else "(DOCX中未提取到文本)"
                            except Exception as e:
                                file_content = f"(DOCX解析失败: {str(e)})"
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
                            file_content = effective_data.decode("utf-8", errors="replace")
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

        return self._create_dir_node(path_parts[1:], actual_node_id, level + 1)

    # ==============================
    # 事件处理函数
    # ==============================
    def on_tree_select(self, event):
        """选择树节点"""
        selection = self.tree.selection()
        if not selection:
            return
        node_id = selection[0]
        node_type = self.node_type.get(node_id)
        content = ""

        if node_type == "file":
            # 压缩包内的文件
            node_data = self.node_data.get(node_id, "")
            if isinstance(node_data, dict):
                if node_data.get("kind") == "pdf_bytes":
                    self.setup_pdf_lazy({"bytes": node_data.get("bytes", b"")},
                                        node_data.get("name", self.tree.item(node_id, "text")))
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
                    # 检查文件大小
                    file_size = os.path.getsize(file_path)
                    if file_size > self.max_file_size:
                        if not messagebox.askyesno("提示",
                                                   f"文件大小 {file_size / 1024 / 1024:.1f} MB，超过10MB限制，是否继续？"):
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
                                self.display_image_in_editor(img,
                                                             title=f"=== 图片文件：{os.path.basename(file_path)} ===")
                        else:
                            self.display_text_in_editor("(未安装Pillow，无法显示图片。请安装：pip install pillow)")
                        return
                    self.clear_pdf_lazy_state()
                    if file_size > LARGE_FILE_THRESHOLD:
                        preview = read_text_preview_from_path(file_path)
                        content = f"=== 大文件预览：{os.path.basename(file_path)} ({file_size / 1024 / 1024:.2f}MB) ===\n\n{preview}"
                    else:
                        if file_ext in ['.xlsx', '.xls', '.csv'] and PANDAS_AVAILABLE:
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
                        elif file_ext == '.json':
                            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                                json_content = f.read()
                            try:
                                json_data = json.loads(json_content)
                                content = json.dumps(json_data, indent=4, ensure_ascii=False)
                            except:
                                content = json_content
                        elif file_ext in ['.doc', '.docx']:
                            content = extract_doc_text(file_path)
                        else:
                            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                                content = f.read()

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
            messagebox.showerror("拖拽错误",
                                 f"解析路径失败：{str(e)}\n\n建议：请检查文件路径是否正确，或尝试使用按钮选择文件")

    def open_compare_window(self):
        """打开文本比较窗口"""
        TextCompareWindow(self)

    def open_alarm_monitor_window(self):
        """打开监控告警窗口"""
        AlarmMonitorWindow(self)

    # ==============================
    # 搜索功能
    # ==============================
    def smart_search(self):
        """智能搜索"""
        keyword = self.search_entry.get().strip()
        if not keyword:
            messagebox.showinfo("提示", "请输入搜索关键词")
            return

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
            # 定位到第一个结果
            first_node = self.file_search_hits[0][0]
            self.tree.selection_set(first_node)
            self.tree.see(first_node)
        else:
            self.result_txt.insert("1.0", "未找到匹配的文件名")
            self.result_title.config(text="🔍 文件名搜索结果 | 匹配数：0")

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

        # 更新预览框
        self.result_txt.config(state=tk.NORMAL)
        self.result_txt.delete("1.0", tk.END)

        if self.content_search_hits:
            # 显示命中结果
            hit_texts = []
            for idx, (row, col, content) in enumerate(self.content_search_hits):
                hit_texts.append(f"第{row}行: {content}")
                self.preview_line_to_hit_index.append(idx)
                self.preview_jump_entries.append(("content", int(row), int(col)))
            self.result_txt.insert("1.0", "\n".join(hit_texts))
            self.result_title.config(text=f"🔍 内容搜索结果 | 匹配数：{hit_count}")
        else:
            self.result_txt.insert("1.0", "未找到匹配的内容")
            self.result_title.config(text="🔍 内容搜索结果 | 匹配数：0")

        self.result_txt.config(state=tk.DISABLED)
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
                    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                        content = f.read()
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
                with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                    lines = f.read().splitlines()
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

        self.result_txt.config(state=tk.NORMAL)
        self.result_txt.delete("1.0", tk.END)
        if self.multi_content_search_hits:
            max_show = 500
            show_hits = self.multi_content_search_hits[:max_show]
            texts = []
            for idx, (nid, row, col, line) in enumerate(show_hits, start=1):
                name = self.tree.item(nid, "text")
                texts.append(f"{idx}. [{name}] 第{row}行: {line}")
                self.preview_line_to_hit_index.append(idx - 1)
                self.preview_jump_entries.append(("content_multi", nid, int(row), int(col)))
            if len(self.multi_content_search_hits) > max_show:
                texts.append(f"... 共 {len(self.multi_content_search_hits)} 条，仅显示前 {max_show} 条")
                self.preview_jump_entries.append(None)
            self.result_txt.insert("1.0", "\n".join(texts))
            self.result_title.config(text=f"🔍 多文件内容搜索 | 匹配数：{len(self.multi_content_search_hits)}")
        else:
            self.result_txt.insert("1.0", "未找到匹配的内容")
            self.result_title.config(text="🔍 多文件内容搜索 | 匹配数：0")
        self.result_txt.config(state=tk.DISABLED)

    def on_double_click_jump(self, event):
        """双击搜索结果跳转"""
        try:
            click_index = self.result_txt.index(f"@{event.x},{event.y}")
            line_num = int(click_index.split(".")[0]) - 1
            if line_num < 0 or line_num >= len(self.preview_jump_entries):
                return
            entry = self.preview_jump_entries[line_num]
            if not entry:
                return
            jump_type = entry[0]

            if jump_type == "file":
                node_id = entry[1]
                self.tree.selection_set(node_id)
                self.tree.see(node_id)
                self.tree.focus_set()
            elif jump_type == "content":
                row, col = entry[1], entry[2]
                target_pos = f"{row}.{col}"
                self.txt.tag_remove("jump_hl", "1.0", tk.END)
                self.txt.tag_add("jump_hl", f"{row}.0", f"{row}.end")
                self.txt.mark_set(tk.INSERT, target_pos)
                self.txt.see(target_pos)
                self.txt.focus_set()
            elif jump_type == "content_multi":
                node_id, row, col = entry[1], entry[2], entry[3]
                self.tree.selection_set(node_id)
                self.tree.see(node_id)
                self.on_tree_select(None)
                target_pos = f"{row}.{max(0, col)}"
                self.txt.tag_remove("jump_hl", "1.0", tk.END)
                self.txt.tag_add("jump_hl", f"{row}.0", f"{row}.end")
                self.txt.mark_set(tk.INSERT, target_pos)
                self.txt.see(target_pos)
                self.txt.focus_set()
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

        # 清空搜索结果列表
        self.file_search_hits.clear()
        self.content_search_hits.clear()
        self.multi_content_search_hits.clear()
        self.preview_line_to_hit_index.clear()
        self.preview_jump_entries.clear()
        self.search_text_cache.clear()


# ==============================
# 程序入口
# ==============================
if __name__ == "__main__":
    try:
        app = FileViewerApp()
        app.mainloop()
    except Exception as e:
        error_msg = f"错误信息：{str(e)}\n\n解决方法：\n1. 确保安装基础依赖：pip install tkinterdnd2\n2. 确保Python版本≥3.6"
        if not PANDAS_AVAILABLE:
            error_msg += "\n3. 如需Excel/CSV支持，请安装：pip install pandas openpyxl"
        messagebox.showerror("启动失败", error_msg)
