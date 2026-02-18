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

from pathlib import Path
from typing import List, Optional

from .base import BaseHandler

try:
    from pypdf import PdfWriter, PdfReader

    HAVE_PYPDF = True
except ImportError:
    HAVE_PYPDF = False
    print("警告: pypdf 未安装，PDF功能不可用")

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4

    HAVE_REPORTLAB = True
except ImportError:
    HAVE_REPORTLAB = False
    print("警告: reportlab 未安装，PDF生成功能不可用")


class PDFHandler(BaseHandler):
    """PDF文档处理器"""

    def __init__(self):
        super().__init__()
        self.file_ext = ".pdf"

    def create_single_file(self, file_name: str, save_path: Optional[Path] = None) -> bool:
        """创建单个PDF文档"""
        if not HAVE_REPORTLAB:
            self.logger.error("reportlab 未安装")
            return False

        try:
            # 确保文件名有扩展名
            if not file_name.endswith(self.file_ext):
                file_name += self.file_ext

            # 确定保存路径
            if save_path:
                output_path = save_path / file_name
            else:
                output_path = self.default_save_path / file_name

            # 确保目录存在
            output_path.parent.mkdir(parents=True, exist_ok=True)

            # 创建PDF
            c = canvas.Canvas(str(output_path), pagesize=A4)
            c.drawString(100, 750, "这是一个新PDF文档")
            c.showPage()
            c.save()

            self.logger.info(f"创建PDF文档: {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"创建PDF文档失败: {e}")
            return False

    def create_multiple(self, count: int, prefix: str, save_path: Optional[Path] = None) -> bool:
        """批量创建PDF文档"""
        if not HAVE_REPORTLAB:
            self.logger.error("reportlab 未安装")
            return False

        try:
            # 确定保存路径
            if save_path:
                folder_path = save_path / prefix
            else:
                folder_path = self.default_save_path / prefix

            # 创建文件夹
            folder_path.mkdir(parents=True, exist_ok=True)

            # 批量创建
            for i in range(1, count + 1):
                file_name = f"{prefix}_{i}{self.file_ext}"
                file_path = folder_path / file_name

                c = canvas.Canvas(str(file_path), pagesize=A4)
                c.drawString(100, 750, f"这是第 {i} 个PDF文档")
                c.showPage()
                c.save()

            self.logger.info(f"批量创建 {count} 个PDF文档到: {folder_path}")
            return True

        except Exception as e:
            self.logger.error(f"批量创建PDF文档失败: {e}")
            return False

    def merge_files(self, source_files: List[Path], output_path: Path) -> bool:
        """合并PDF文档"""
        if not HAVE_PYPDF:
            self.logger.error("pypdf 未安装")
            return False

        if not source_files:
            self.logger.error("没有源文件")
            return False

        try:
            # 确保输出目录存在
            output_path.parent.mkdir(parents=True, exist_ok=True)

            # 创建PDF写入器
            pdf_writer = PdfWriter()

            # 合并所有PDF
            for file_path in source_files:
                with open(file_path, 'rb') as f:
                    pdf_reader = PdfReader(f)
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)

            # 写入输出文件
            with open(output_path, 'wb') as f:
                pdf_writer.write(f)

            self.logger.info(f"合并PDF文档到: {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"合并PDF文档失败: {e}")
            return False

    def split_file(self, source_path: Path, split_pos: int, output_dir: Optional[Path] = None) -> List[Path]:
        """拆分PDF文档"""
        if not HAVE_PYPDF:
            self.logger.error("pypdf 未安装")
            return []

        try:
            # 确定输出目录
            if output_dir:
                save_dir = output_dir
            else:
                save_dir = source_path.parent

            # 打开源PDF
            with open(source_path, 'rb') as f:
                pdf_reader = PdfReader(f)
                total_pages = len(pdf_reader.pages)

                # 验证拆分位置
                if split_pos <= 0 or split_pos >= total_pages:
                    self.logger.error(f"拆分位置无效: {split_pos}，总页数: {total_pages}")
                    return []

                # 创建两个写入器
                writer1 = PdfWriter()
                writer2 = PdfWriter()

                # 拆分页面
                for i, page in enumerate(pdf_reader.pages):
                    if i < split_pos:
                        writer1.add_page(page)
                    else:
                        writer2.add_page(page)

            # 生成输出路径
            source_name = source_path.stem
            output1 = save_dir / f"{source_name}_拆分1{self.file_ext}"
            output2 = save_dir / f"{source_name}_拆分2{self.file_ext}"

            # 保存拆分后的PDF
            with open(output1, 'wb') as f:
                writer1.write(f)

            with open(output2, 'wb') as f:
                writer2.write(f)

            self.logger.info(f"拆分PDF文档: {output1}, {output2}")
            return [output1, output2]

        except Exception as e:
            self.logger.error(f"拆分PDF文档失败: {e}")
            return []
