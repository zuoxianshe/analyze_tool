# -*- coding: utf-8 -*-
from app_common import PANDAS_AVAILABLE, messagebox
from file_viewer_app import FileViewerApp


if __name__ == "__main__":
    try:
        app = FileViewerApp()
        app.mainloop()
    except Exception as e:
        error_msg = f"错误信息：{str(e)}\n\n解决方法：\n1. 确保安装基础依赖：pip install tkinterdnd2\n2. 确保Python版本≥3.6"
        if not PANDAS_AVAILABLE:
            error_msg += "\n3. 如需Excel/CSV支持，请安装：pip install pandas openpyxl"
        messagebox.showerror("启动失败", error_msg)
