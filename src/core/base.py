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
from pathlib import Path
from typing import List, Optional


class BaseHandler:
    """基础文件处理器"""

    def __init__(self):
        self.default_save_path = Path.home() / "Desktop"
        self.logger = logging.getLogger("WP_Express")

    def validate_path(self, path: Path) -> bool:
        """验证路径"""
        try:
            if path.exists():
                return True
            # 如果路径不存在，检查父目录是否存在
            if path.parent.exists():
                return True
            return False
        except Exception as e:
            self.logger.error(f"路径验证失败: {e}")
            return False

    def create_single_file(self, file_name: str, save_path: Optional[Path] = None) -> bool:
        """创建单个文件 - 子类实现"""
        raise NotImplementedError

    def create_multiple(self, count: int, prefix: str, save_path: Optional[Path] = None) -> bool:
        """批量创建文件 - 子类实现"""
        raise NotImplementedError

    def merge_files(self, source_files: List[Path], output_path: Path) -> bool:
        """合并文件 - 子类实现"""
        raise NotImplementedError

    def split_file(self, source_path: Path, split_pos: int, output_dir: Optional[Path] = None) -> List[Path]:
        """拆分文件 - 子类实现"""
        raise NotImplementedError

    def move_file(self, source_path: Path, target_path: Path) -> bool:
        """移动文件"""
        try:
            import shutil
            shutil.move(str(source_path), str(target_path))
            self.logger.info(f"文件移动: {source_path} -> {target_path}")
            return True
        except Exception as e:
            self.logger.error(f"文件移动失败: {e}")
            return False
