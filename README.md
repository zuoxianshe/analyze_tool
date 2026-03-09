# Smart Helper 使用说明

`smart.py` 是一个基于 `tkinter` 的桌面工具，集成了文件加载、文本解析/编辑、全文搜索、多文件搜索、PDF分页预览、文本对比、监控告警筛选等功能。

## 1. 功能概览

- 主界面文件解析
  - 支持加载：文件夹、压缩包、单文件
  - 支持拖拽加载：文件夹/压缩包/文件（仅拖到文件结构区或文本区才触发解析）
  - 支持格式：`txt/json/xml/md/csv/xlsx/xls/pdf/doc/docx` 及常见图片格式
- 文本编辑
  - 加粗、颜色、恢复默认、保存文本
  - `Ctrl + S` 保存当前文本
  - `Ctrl + 滚轮` 缩放文本显示
  - Markdown 渲染（支持按钮状态显示“已渲染”）
- 搜索能力
  - 文件名搜索、单文件内容搜索、多文件内容搜索
  - 搜索结果双击跳转到对应节点/文本位置（已修复偶发不跳转问题）
  - `Ctrl + F` 快速聚焦到智能搜索框；若已有内容会自动全选
- PDF 大文件能力
  - 单页懒加载解析，滚轮翻页
  - 下方居中分页进度显示
  - 右下角显示当前页数
- 文本对比窗口
  - 支持双文本加载、编辑、差异高亮
  - 支持拖拽加载和 `Ctrl + S` 保存
- 监控告警窗口（独立 UI）
  - Excel/CSV 加载（支持拖拽）
  - 支持按 `FRU对象`、`机型包含`、`仅限机型` 筛选
  - 多 Sheet 自动合并后筛选与输出
  - 结果以近似 Excel 网格展示（排序、复制、单元格编辑）
  - 结果保存为 `xlsx`

## 2. 运行环境

- 操作系统：Windows（`.doc` 解析对 Windows 更友好）
- Python：建议 3.9+

## 3. 依赖安装

建议先创建虚拟环境后安装：

```powershell
pip install tkinterdnd2 pandas openpyxl pdfplumber pypdf pillow
```

可选依赖（用于 `.doc` 解析）：

- 方案A：安装 Microsoft Word + `pywin32`
```powershell
pip install pywin32
```
- 方案B：安装 LibreOffice（提供 `soffice` 命令行转换）

## 4. 启动方式

在项目目录执行：

```powershell
python smart.py
```

如果系统没有 `python` 命令，可尝试：

```powershell
py -3 smart.py
```

## 5. 使用说明

### 5.1 文件加载

- `📂 文件夹/压缩包`：先选文件夹，取消后可继续选择压缩包
- `📄 选择文件`：选择单个文件
- 也可直接拖拽到文件结构区或文本框进行解析

### 5.2 智能搜索

- 在“智能搜索”输入关键字后执行搜索
- 搜索模式：
  - 光标在文件树侧重“文件名搜索”
  - 光标在文本框侧重“内容搜索”
  - 点击“多文件”执行跨文件内容搜索
- 双击结果可跳转到对应位置

### 5.3 监控告警筛选

- 点击 `🚨 监控告警` 进入独立窗口
- 加载 `xlsx/xls/csv`
- 输入筛选条件后点击“筛选告警”
- “输出全量告警”可直接显示全部告警
- 可编辑网格并保存为 Excel

## 6. 常见问题

- `ModuleNotFoundError: No module named 'tkinterdnd2'`
  - 执行：`pip install tkinterdnd2`
- `加载失败：import openpyxl failed`
  - 执行：`pip install openpyxl`
- `import xlrd failed`
  - 执行：`pip install xlrd`
- `.doc` 解析失败
  - 优先检查是否安装 Microsoft Word + `pywin32`
  - 或安装 LibreOffice，确保 `soffice` 可在命令行调用
- PDF 乱码/慢
  - 已内置多解析器回退与分页懒加载；超大文件建议使用分页浏览与关键词搜索

## 7. 项目文件

- 主程序：`smart.py`
- 规则文档：`BMC_BIOS_Redfish_Registry_规则整理.md`

