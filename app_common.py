# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
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
import html
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
ARCHIVE_MEMBER_FULL_READ_BYTES = 32 * 1024 * 1024
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}


def decode_bytes_auto(data):
    """Best-effort bytes->text decode for mixed Chinese/Unicode files."""
    if data is None:
        return ""
    if isinstance(data, str):
        return data
    if not isinstance(data, (bytes, bytearray)):
        return str(data)

    b = bytes(data)
    # UTF BOM first
    for enc in ("utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "utf-32"):
        try:
            return b.decode(enc)
        except Exception:
            pass
    # Common Windows/Chinese encodings
    for enc in ("utf-8", "gb18030", "gbk", "big5", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            pass
    return b.decode("utf-8", errors="replace")


def read_text_file_auto(file_path):
    """Read text file with encoding fallback."""
    with open(file_path, "rb") as f:
        return decode_bytes_auto(f.read())


def sanitize_xml_text(text):
    """Remove illegal XML 1.0 control chars to avoid parser abort on malformed runs."""
    if not text:
        return ""
    # Keep: TAB(0x09), LF(0x0A), CR(0x0D)
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)


def extract_docx_part_text_relaxed(xml_text):
    """
    Fallback extractor for malformed DOCX XML parts.
    Uses regex-based extraction so one bad run/format does not cut off subsequent text.
    """
    text = sanitize_xml_text(xml_text or "")
    if not text:
        return []

    para_chunks = re.split(r"</w:p\s*>", text, flags=re.IGNORECASE)
    out_lines = []
    for chunk in para_chunks:
        if not chunk:
            continue
        chunk = re.sub(r"<w:tab\b[^>]*/>", "\t", chunk, flags=re.IGNORECASE)
        chunk = re.sub(r"<w:(?:br|cr)\b[^>]*/>", "\n", chunk, flags=re.IGNORECASE)
        pieces = re.findall(
            r"<w:(?:t|instrText|delText)\b[^>]*>(.*?)</w:(?:t|instrText|delText)\s*>",
            chunk,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if not pieces:
            continue
        line = "".join(html.unescape(p) for p in pieces).strip()
        if line:
            out_lines.append(line)
    return out_lines


def read_text_preview_from_path(file_path, max_bytes=PREVIEW_READ_BYTES):
    try:
        with open(file_path, "rb") as f:
            data = f.read(max_bytes + 1)
        truncated = len(data) > max_bytes
        text = decode_bytes_auto(data[:max_bytes])
        if truncated:
            text += f"\n\n[预览模式] 文件较大，仅显示前 {max_bytes // 1024}KB。"
        return text
    except Exception as e:
        return f"(预览读取失败: {str(e)})"


def read_text_preview_from_bytes(file_data, max_bytes=PREVIEW_READ_BYTES):
    data = file_data[: max_bytes + 1]
    truncated = len(data) > max_bytes
    text = decode_bytes_auto(data[:max_bytes])
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


def extract_docx_text(file_path=None, file_bytes=None):
    try:
        zsrc = io.BytesIO(file_bytes) if file_bytes is not None else file_path
        with zipfile.ZipFile(zsrc) as zf:
            names = set(zf.namelist())
            doc_parts = []
            if "word/document.xml" in names:
                doc_parts.append("word/document.xml")
            doc_parts.extend(sorted([n for n in names if n.startswith("word/header") and n.endswith(".xml")]))
            doc_parts.extend(sorted([n for n in names if n.startswith("word/footer") and n.endswith(".xml")]))
            for p in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"):
                if p in names:
                    doc_parts.append(p)
            if not doc_parts:
                return "(DOCX中未找到可解析的XML文本内容)"

            xml_docs = []
            for part in doc_parts:
                try:
                    xml_docs.append((part, zf.read(part)))
                except Exception:
                    continue

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        lines = []
        parsed_any = False
        for part_name, xml_bytes in xml_docs:
            xml_text = decode_bytes_auto(xml_bytes)
            xml_text = sanitize_xml_text(xml_text)
            try:
                root = ET.fromstring(xml_text)
                parsed_any = True
                for para in root.findall(".//w:p", ns):
                    parts = []
                    for node in para.iter():
                        tag = node.tag.split("}")[-1] if "}" in node.tag else node.tag
                        if tag in ("t", "instrText", "delText"):
                            parts.append(node.text or "")
                        elif tag == "tab":
                            parts.append("\t")
                        elif tag in ("br", "cr"):
                            parts.append("\n")
                    line = "".join(parts).strip()
                    if line:
                        lines.append(line)
            except Exception:
                # Fallback: tolerate malformed XML and keep extracting later content.
                fallback_lines = extract_docx_part_text_relaxed(xml_text)
                if fallback_lines:
                    parsed_any = True
                    lines.extend(fallback_lines)
        if not parsed_any:
            return "(DOCX解析失败: XML内容不可解析)"
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
