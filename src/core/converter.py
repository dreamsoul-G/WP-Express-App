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

import sys
from pathlib import Path
from typing import Optional


class Converter:
    """文档格式转换器"""

    def __init__(self):
        self.logger = None
        self._setup_logger()

    def _setup_logger(self):
        """设置日志"""
        import logging
        self.logger = logging.getLogger("WP_Express")

    def word_to_pdf(self, source_path: Path, output_path: Optional[Path] = None) -> bool:
        """Word转PDF"""
        # 检查是否支持转换
        if not sys.platform.startswith('win'):
            self.logger.error("Word转PDF仅支持Windows系统")
            return False

        try:
            import win32com.client
            import pythoncom
        except ImportError:
            self.logger.error("请安装pywin32: pip install pywin32")
            return False

        try:
            # 确定输出路径
            if output_path is None:
                output_path = source_path.with_suffix('.pdf')

            # 启动Word应用程序
            pythoncom.CoInitialize()
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0

            # 打开Word文档
            doc = word_app.Documents.Open(str(source_path))

            # 保存为PDF
            doc.SaveAs(str(output_path), FileFormat=17)  # 17 = wdFormatPDF

            # 清理
            doc.Close()
            word_app.Quit()
            pythoncom.CoUninitialize()

            self.logger.info(f"Word转PDF成功: {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"Word转PDF失败: {e}")
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            return False

    def pdf_to_word(self, source_path: Path, output_path: Optional[Path] = None) -> bool:
        """PDF转Word"""
        # 检查是否支持转换
        if not sys.platform.startswith('win'):
            self.logger.error("PDF转Word仅支持Windows系统")
            return False

        try:
            import win32com.client
            import pythoncom
        except ImportError:
            self.logger.error("请安装pywin32: pip install pywin32")
            return False

        try:
            # 确定输出路径
            if output_path is None:
                output_path = source_path.with_suffix('.docx')

            # 启动Word应用程序
            pythoncom.CoInitialize()
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0

            # 打开PDF文档
            doc = word_app.Documents.Open(str(source_path))

            # 保存为Word
            doc.SaveAs(str(output_path), FileFormat=16)  # 16 = wdFormatDocumentDefault

            # 清理
            doc.Close()
            word_app.Quit()
            pythoncom.CoUninitialize()

            self.logger.info(f"PDF转Word成功: {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"PDF转Word失败: {e}")
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            return False
