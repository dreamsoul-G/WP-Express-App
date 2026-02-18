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

import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox


class BaseDialog:
    """对话框基类"""

    def __init__(self, parent, title, width=400, height=300):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 居中显示
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2

        self.dialog.geometry(f"{width}x{height}+{x}+{y}")

        # 主框架
        self.main_frame = ttk.Frame(self.dialog, padding=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

    def add_label(self, text, **kwargs):
        """添加标签"""
        label = ttk.Label(self.main_frame, text=text, **kwargs)
        label.pack(anchor=tk.W, pady=2)
        return label

    def add_entry(self, default="", **kwargs):
        """添加输入框"""
        entry = ttk.Entry(self.main_frame, **kwargs)
        entry.insert(0, default)
        entry.pack(fill=tk.X, pady=5)
        return entry

    def add_button(self, text, command, side=tk.RIGHT):
        """添加按钮"""
        btn = ttk.Button(self.main_frame, text=text, command=command)
        btn.pack(side=side, padx=5, pady=10)
        return btn

    def destroy(self):
        """销毁对话框"""
        self.dialog.destroy()


class BatchCreateDialog(BaseDialog):
    """批量创建对话框"""

    def __init__(self, parent, file_type, callback):
        super().__init__(parent, f"批量创建{file_type}文件", 400, 350)
        self.file_type = file_type
        self.callback = callback

        # 文件数量
        self.add_label("创建数量:")
        self.count_entry = self.add_entry("1")

        # 文件前缀
        self.add_label("文件前缀:")
        self.prefix_entry = self.add_entry("文档")

        # 保存路径
        self.add_label("保存路径:")

        path_frame = ttk.Frame(self.main_frame)
        path_frame.pack(fill=tk.X, pady=5)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(path_frame, text="浏览",
                   command=self.browse_path).pack(side=tk.LEFT)

        # 使用桌面
        self.use_desktop = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self.main_frame,
            text="保存到桌面",
            variable=self.use_desktop
        ).pack(anchor=tk.W, pady=10)

        # 按钮
        self.add_button("取消", self.destroy, tk.LEFT)
        self.add_button("创建", self.create_files, tk.RIGHT)

    def browse_path(self):
        """浏览路径"""
        path = filedialog.askdirectory(title="选择保存路径")
        if path:
            self.path_var.set(path)
            self.use_desktop.set(False)

    def create_files(self):
        """创建文件"""
        try:
            count = int(self.count_entry.get())
            prefix = self.prefix_entry.get()

            if count <= 0:
                messagebox.showerror("错误", "数量必须大于0")
                return

            if self.use_desktop.get():
                save_path = None
            else:
                save_path = Path(self.path_var.get())
                if not save_path.exists():
                    messagebox.showerror("错误", "保存路径不存在")
                    return

            self.callback(count, prefix, save_path)
            self.destroy()

        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")


class MergeDialog(BaseDialog):
    """合并文件对话框"""

    def __init__(self, parent, file_type, callback):
        super().__init__(parent, f"合并{file_type}文件", 500, 400)
        self.file_type = file_type
        self.callback = callback
        self.files = []

        # 文件列表
        self.add_label(f"选择{file_type}文件:")

        # 列表框
        frame = ttk.Frame(self.main_frame)
        frame.pack(fill=tk.BOTH, expand=True, pady=5)

        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, height=6)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.listbox.yview)

        # 按钮框架
        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame, text="添加文件",
                   command=self.add_files).pack(side=tk.LEFT, padx=2)

        ttk.Button(btn_frame, text="移除选中",
                   command=self.remove_selected).pack(side=tk.LEFT, padx=2)

        ttk.Button(btn_frame, text="清空列表",
                   command=self.clear_list).pack(side=tk.LEFT, padx=2)

        # 输出文件名
        self.add_label("输出文件名:")
        self.output_entry = self.add_entry("合并文档")

        # 保存路径
        self.add_label("保存路径:")

        path_frame = ttk.Frame(self.main_frame)
        path_frame.pack(fill=tk.X, pady=5)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(path_frame, text="浏览",
                   command=self.browse_save_path).pack(side=tk.LEFT)

        # 按钮
        self.add_button("取消", self.destroy, tk.LEFT)
        self.add_button("合并", self.merge_files, tk.RIGHT)

    def add_files(self):
        """添加文件"""
        if self.file_type == "Word":
            filetypes = [("Word文档", "*.docx"), ("所有文件", "*.*")]
        else:
            filetypes = [("PDF文档", "*.pdf"), ("所有文件", "*.*")]

        files = filedialog.askopenfilenames(
            title=f"选择{self.file_type}文件",
            filetypes=filetypes
        )

        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.listbox.insert(tk.END, Path(file).name)

    def remove_selected(self):
        """移除选中"""
        selections = self.listbox.curselection()
        for idx in reversed(selections):
            self.listbox.delete(idx)
            self.files.pop(idx)

    def clear_list(self):
        """清空列表"""
        self.listbox.delete(0, tk.END)
        self.files.clear()

    def browse_save_path(self):
        """浏览保存路径"""
        if self.file_type == "Word":
            filetypes = [("Word文档", "*.docx")]
        else:
            filetypes = [("PDF文档", "*.pdf")]

        file = filedialog.asksaveasfilename(
            title="保存合并文件",
            defaultextension=filetypes[0][1].replace("*", ""),
            filetypes=filetypes
        )

        if file:
            self.path_var.set(file)

    def merge_files(self):
        """合并文件"""
        if not self.files:
            messagebox.showerror("错误", "请先添加文件")
            return

        output_name = self.output_entry.get()
        if not output_name:
            messagebox.showerror("错误", "请输入输出文件名")
            return

        save_path = self.path_var.get()
        if not save_path:
            messagebox.showerror("错误", "请选择保存路径")
            return

        self.callback([Path(f) for f in self.files], Path(save_path))
        self.destroy()


class SplitDialog(BaseDialog):
    """拆分文件对话框"""

    def __init__(self, parent, file_type, callback):
        super().__init__(parent, f"拆分{file_type}文件", 400, 300)
        self.file_type = file_type
        self.callback = callback

        # 选择文件
        self.add_label(f"选择{file_type}文件:")

        file_frame = ttk.Frame(self.main_frame)
        file_frame.pack(fill=tk.X, pady=5)

        self.file_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_var)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(file_frame, text="浏览",
                   command=self.browse_file).pack(side=tk.LEFT)

        # 拆分位置
        self.add_label("拆分位置（段落/页数）:")
        self.pos_entry = self.add_entry("1")

        # 保存路径
        self.add_label("保存路径:")

        path_frame = ttk.Frame(self.main_frame)
        path_frame.pack(fill=tk.X, pady=5)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(path_frame, text="浏览",
                   command=self.browse_path).pack(side=tk.LEFT)

        # 按钮
        self.add_button("取消", self.destroy, tk.LEFT)
        self.add_button("拆分", self.split_file, tk.RIGHT)

    def browse_file(self):
        """浏览文件"""
        if self.file_type == "Word":
            filetypes = [("Word文档", "*.docx"), ("所有文件", "*.*")]
        else:
            filetypes = [("PDF文档", "*.pdf"), ("所有文件", "*.*")]

        file = filedialog.askopenfilename(
            title=f"选择{self.file_type}文件",
            filetypes=filetypes
        )

        if file:
            self.file_var.set(file)

    def browse_path(self):
        """浏览路径"""
        path = filedialog.askdirectory(title="选择保存路径")
        if path:
            self.path_var.set(path)

    def split_file(self):
        """拆分文件"""
        file_path = self.file_var.get()
        if not file_path:
            messagebox.showerror("错误", "请选择文件")
            return

        try:
            position = int(self.pos_entry.get())
            if position <= 0:
                messagebox.showerror("错误", "位置必须大于0")
                return
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
            return

        save_path = Path(self.path_var.get()) if self.path_var.get() else None

        self.callback(Path(file_path), position, save_path)
        self.destroy()


class MoveDialog(BaseDialog):
    """移动/重命名对话框"""

    def __init__(self, parent, file_type, callback):
        super().__init__(parent, f"移动/重命名{file_type}文件", 400, 250)
        self.file_type = file_type
        self.callback = callback

        # 源文件
        self.add_label("选择源文件:")

        source_frame = ttk.Frame(self.main_frame)
        source_frame.pack(fill=tk.X, pady=5)

        self.source_var = tk.StringVar()
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_var)
        self.source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(source_frame, text="浏览",
                   command=self.browse_source).pack(side=tk.LEFT)

        # 目标文件
        self.add_label("目标位置:")

        target_frame = ttk.Frame(self.main_frame)
        target_frame.pack(fill=tk.X, pady=5)

        self.target_var = tk.StringVar()
        self.target_entry = ttk.Entry(target_frame, textvariable=self.target_var)
        self.target_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(target_frame, text="浏览",
                   command=self.browse_target).pack(side=tk.LEFT)

        # 按钮
        self.add_button("取消", self.destroy, tk.LEFT)
        self.add_button("移动", self.move_file, tk.RIGHT)

    def browse_source(self):
        """浏览源文件"""
        if self.file_type == "Word":
            filetypes = [("Word文档", "*.docx"), ("所有文件", "*.*")]
        else:
            filetypes = [("PDF文档", "*.pdf"), ("所有文件", "*.*")]

        file = filedialog.askopenfilename(
            title=f"选择{self.file_type}文件",
            filetypes=filetypes
        )

        if file:
            self.source_var.set(file)

    def browse_target(self):
        """浏览目标位置"""
        # 如果是文件，保存对话框；如果是目录，选择目录
        if self.source_var.get():
            # 尝试作为文件保存
            if self.file_type == "Word":
                filetypes = [("Word文档", "*.docx")]
            else:
                filetypes = [("PDF文档", "*.pdf")]

            file = filedialog.asksaveasfilename(
                title="选择目标文件",
                defaultextension=filetypes[0][1].replace("*", ""),
                filetypes=filetypes
            )

            if file:
                self.target_var.set(file)
        else:
            # 选择目录
            path = filedialog.askdirectory(title="选择目标目录")
            if path:
                self.target_var.set(path)

    def move_file(self):
        """移动文件"""
        source = self.source_var.get()
        target = self.target_var.get()

        if not source:
            messagebox.showerror("错误", "请选择源文件")
            return

        if not target:
            messagebox.showerror("错误", "请选择目标位置")
            return

        self.callback(Path(source), Path(target))
        self.destroy()


class ConvertDialog(BaseDialog):
    """转换对话框"""

    def __init__(self, parent, convert_type, callback):
        title = "Word转PDF" if convert_type == "Word转PDF" else "PDF转Word"
        super().__init__(parent, title, 400, 250)
        self.convert_type = convert_type
        self.callback = callback

        # 源文件
        if convert_type == "Word转PDF":
            self.add_label("选择Word文档:")
        else:
            self.add_label("选择PDF文档:")

        source_frame = ttk.Frame(self.main_frame)
        source_frame.pack(fill=tk.X, pady=5)

        self.source_var = tk.StringVar()
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_var)
        self.source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(source_frame, text="浏览",
                   command=self.browse_source).pack(side=tk.LEFT)

        # 目标文件
        if convert_type == "Word转PDF":
            self.add_label("保存为PDF:")
        else:
            self.add_label("保存为Word文档:")

        target_frame = ttk.Frame(self.main_frame)
        target_frame.pack(fill=tk.X, pady=5)

        self.target_var = tk.StringVar()
        self.target_entry = ttk.Entry(target_frame, textvariable=self.target_var)
        self.target_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(target_frame, text="浏览",
                   command=self.browse_target).pack(side=tk.LEFT)

        # 按钮
        self.add_button("取消", self.destroy, tk.LEFT)
        self.add_button("转换", self.convert_file, tk.RIGHT)

    def browse_source(self):
        """浏览源文件"""
        if self.convert_type == "Word转PDF":
            filetypes = [("Word文档", "*.docx;*.doc"), ("所有文件", "*.*")]
        else:
            filetypes = [("PDF文档", "*.pdf"), ("所有文件", "*.*")]

        file = filedialog.askopenfilename(
            title="选择源文件",
            filetypes=filetypes
        )

        if file:
            self.source_var.set(file)

    def browse_target(self):
        """浏览目标文件"""
        if self.convert_type == "Word转PDF":
            filetypes = [("PDF文档", "*.pdf"), ("所有文件", "*.*")]
            defaultext = ".pdf"
        else:
            filetypes = [("Word文档", "*.docx"), ("所有文件", "*.*")]
            defaultext = ".docx"

        file = filedialog.asksaveasfilename(
            title="保存转换文件",
            defaultextension=defaultext,
            filetypes=filetypes
        )

        if file:
            self.target_var.set(file)

    def convert_file(self):
        """转换文件"""
        source = self.source_var.get()
        target = self.target_var.get()

        if not source:
            messagebox.showerror("错误", "请选择源文件")
            return

        # 如果没选目标，使用默认
        if not target:
            target = None

        self.callback(Path(source), Path(target) if target else None, False)
        self.destroy()
