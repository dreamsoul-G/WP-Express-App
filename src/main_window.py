"""
WP快通（文档处理）软件
Copyright [2026] [郭宇轩]

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

import logging
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox

from PIL import Image, ImageTk

from .core.converter import Converter
from .core.pdf_handler import PDFHandler
# 导入核心模块
from .core.word_handler import WordHandler
from .dialogs import (
    BatchCreateDialog, MergeDialog,
    SplitDialog, MoveDialog, ConvertDialog
)


class MainWindow:
    """主窗口类"""

    def __init__(self):
        self.root = tk.Tk()
        self.setup_window()
        self.setup_handlers()
        self.load_images()
        self.create_widgets()
        self.center_window()

    def setup_window(self):
        """窗口设置"""
        self.root.title("WP快通 v1.1.0.2026217")
        self.root.geometry("1200x700")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")

        # 设置图标
        icon_path = Path(__file__).parent.parent / "assets" / "icons" / "app.ico"
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except:
                pass

    def setup_handlers(self):
        """初始化处理器"""
        self.word_handler = WordHandler()
        self.pdf_handler = PDFHandler()
        self.converter = Converter()
        self.logger = logging.getLogger("WPQuickPass")

    def load_images(self):
        """加载图片"""
        # 图片资源 assets/images/*.png
        assets_dir = Path(__file__).parent.parent / "assets" / "images"

        self.word_img = None
        self.pdf_img = None
        self.convert_img = None

        try:
            # Word图片
            word_path = assets_dir / "word_bg.png"
            if word_path.exists():
                img = Image.open(word_path)
                img = img.resize((280, 160), Image.Resampling.LANCZOS)
                self.word_img = ImageTk.PhotoImage(img)

            # PDF图片
            pdf_path = assets_dir / "pdf_bg.png"
            if pdf_path.exists():
                img = Image.open(pdf_path)
                img = img.resize((280, 160), Image.Resampling.LANCZOS)
                self.pdf_img = ImageTk.PhotoImage(img)

            # 转换图片
            convert_path = assets_dir / "convert_bg.png"
            if convert_path.exists():
                img = Image.open(convert_path)
                img = img.resize((280, 160), Image.Resampling.LANCZOS)
                self.convert_img = ImageTk.PhotoImage(img)

        except Exception as e:
            print(f"图片加载失败: {e}")

    def create_widgets(self):
        """创建界面组件"""
        # 主容器
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 标题
        title = ttk.Label(
            main_container,
            text="WP快通 文档处理 软件",
            font=("微软雅黑", 16, "bold"),
            foreground="#2c3e50"
        )
        title.pack(pady=(0, 20))

        # 三栏主框架
        main_frame = ttk.Frame(main_container)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧 - Word功能
        self.create_word_section(main_frame)

        # 中间 - PDF功能
        self.create_pdf_section(main_frame)

        # 右侧 - 转换功能
        self.create_convert_section(main_frame)

        # 状态栏
        self.create_status_bar(main_container)

    def create_word_section(self, parent):
        """创建Word功能区"""
        frame = ttk.LabelFrame(parent, text="Word文档处理", padding=15)
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # 图片
        if self.word_img:
            ttk.Label(frame, image=self.word_img).pack(pady=(0, 15))
        else:
            ttk.Label(frame, text="📄 Word处理", font=("Arial", 12)).pack(pady=(0, 15))

        # 按钮
        buttons = [
            ("📄 批量生成", self.on_word_batch),
            ("📂 移动/重命名", self.on_word_move),
            ("🔗 合并文档", self.on_word_merge),
            ("✂️ 拆分文档", self.on_word_split)
        ]

        for text, command in buttons:
            btn = ttk.Button(frame, text=text, command=command, width=18)
            btn.pack(pady=6)

    def create_pdf_section(self, parent):
        """创建PDF功能区"""
        frame = ttk.LabelFrame(parent, text="PDF文档处理", padding=15)
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # 图片
        if self.pdf_img:
            ttk.Label(frame, image=self.pdf_img).pack(pady=(0, 15))
        else:
            ttk.Label(frame, text="📋 PDF处理", font=("Arial", 12)).pack(pady=(0, 15))

        # 按钮
        buttons = [
            ("📄 批量生成", self.on_pdf_batch),
            ("📂 移动/重命名", self.on_pdf_move),
            ("🔗 合并文档", self.on_pdf_merge),
            ("✂️ 拆分文档", self.on_pdf_split)
        ]

        for text, command in buttons:
            btn = ttk.Button(frame, text=text, command=command, width=18)
            btn.pack(pady=6)

    def create_convert_section(self, parent):
        """创建转换功能区"""
        frame = ttk.LabelFrame(parent, text="格式转换", padding=15)
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # 图片
        if self.convert_img:
            ttk.Label(frame, image=self.convert_img).pack(pady=(0, 15))
        else:
            ttk.Label(frame, text="🔄 格式转换", font=("Arial", 12)).pack(pady=(0, 15))

        # 按钮
        buttons = [
            ("📄 → 📋 Word转PDF", self.on_word_to_pdf),
            ("📋 → 📄 PDF转Word", self.on_pdf_to_word)
        ]

        for text, command in buttons:
            btn = ttk.Button(frame, text=text, command=command, width=18)
            btn.pack(pady=10)

    def create_status_bar(self, parent):
        """创建状态栏"""
        frame = ttk.Frame(parent, relief=tk.SUNKEN)
        frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))

        self.status_label = ttk.Label(frame, text="WP", anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, padx=10)

    def center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def update_status(self, text):
        """更新状态栏"""
        self.status_label.config(text=text)
        self.root.update()

    # Word功能回调
    def on_word_batch(self):
        """Word批量生成"""

        def callback(count, prefix, save_path):
            if self.word_handler.create_multiple(count, prefix, save_path):
                messagebox.showinfo("成功", f"创建了 {count} 个Word文档")

        BatchCreateDialog(self.root, "Word", callback)

    def on_word_move(self):
        """Word移动/重命名"""

        def callback(source, target):
            if self.word_handler.move_file(source, target):
                messagebox.showinfo("成功", "文件移动成功")

        MoveDialog(self.root, "Word", callback)

    def on_word_merge(self):
        """Word合并"""

        def callback(files, output):
            if self.word_handler.merge_files(files, output):
                messagebox.showinfo("成功", "文档合并成功")

        MergeDialog(self.root, "Word", callback)

    def on_word_split(self):
        """Word拆分"""

        def callback(source, position, output_dir):
            results = self.word_handler.split_file(source, position, output_dir)
            if results:
                messagebox.showinfo("成功", f"拆分为 {len(results)} 个文件")

        SplitDialog(self.root, "Word", callback)

    # PDF功能回调
    def on_pdf_batch(self):
        """PDF批量生成"""

        def callback(count, prefix, save_path):
            if self.pdf_handler.create_multiple(count, prefix, save_path):
                messagebox.showinfo("成功", f"创建了 {count} 个PDF文档")

        BatchCreateDialog(self.root, "PDF", callback)

    def on_pdf_move(self):
        """PDF移动/重命名"""

        def callback(source, target):
            if self.pdf_handler.move_file(source, target):
                messagebox.showinfo("成功", "文件移动成功")

        MoveDialog(self.root, "PDF", callback)

    def on_pdf_merge(self):
        """PDF合并"""

        def callback(files, output):
            if self.pdf_handler.merge_files(files, output):
                messagebox.showinfo("成功", "文档合并成功")

        MergeDialog(self.root, "PDF", callback)

    def on_pdf_split(self):
        """PDF拆分"""

        def callback(source, position, output_dir):
            results = self.pdf_handler.split_file(source, position, output_dir)
            if results:
                messagebox.showinfo("成功", f"拆分为 {len(results)} 个文件")

        SplitDialog(self.root, "PDF", callback)

    # 转换功能回调
    def on_word_to_pdf(self):
        """Word转PDF"""

        def callback(source, target, batch=False):
            if batch:
                # 批量转换
                messagebox.showwarning("提示", "批量转换功能正在开发中")
            else:
                if self.converter.word_to_pdf(source, target):
                    messagebox.showinfo("成功", "转换成功")

        ConvertDialog(self.root, "Word转PDF", callback)

    def on_pdf_to_word(self):
        """PDF转Word"""

        def callback(source, target, batch=False):
            if batch:
                # 批量转换
                messagebox.showwarning("提示", "批量转换功能正在开发中")
            else:
                if self.converter.pdf_to_word(source, target):
                    messagebox.showinfo("成功", "转换成功")

        ConvertDialog(self.root, "PDF转Word", callback)

    def run(self):
        """运行主循环"""
        self.root.mainloop()
