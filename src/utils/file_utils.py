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

import shutil
from pathlib import Path
from typing import List, Optional


def ensure_directory(path: Path) -> bool:
    """确保目录存在"""
    try:
        path.mkdir(parents=True, exist_ok=True)
        return True
    except Exception:
        return False


def get_unique_filename(path: Path) -> Path:
    """获取唯一的文件名（避免覆盖）"""
    if not path.exists():
        return path

    counter = 1
    while True:
        new_path = path.parent / f"{path.stem}_{counter}{path.suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


def copy_file(source: Path, target: Path) -> bool:
    """复制文件"""
    try:
        shutil.copy2(str(source), str(target))
        return True
    except Exception:
        return False


def move_file(source: Path, target: Path) -> bool:
    """移动文件"""
    try:
        shutil.move(str(source), str(target))
        return True
    except Exception:
        return False


def delete_file(path: Path) -> bool:
    """删除文件"""
    try:
        if path.exists():
            path.unlink()
        return True
    except Exception:
        return False


def get_file_size(path: Path) -> int:
    """获取文件大小（字节）"""
    try:
        return path.stat().st_size
    except Exception:
        return 0


def list_files(directory: Path, extensions: Optional[List[str]] = None) -> List[Path]:
    """列出目录中的文件"""
    if not directory.exists() or not directory.is_dir():
        return []

    files = []
    for item in directory.iterdir():
        if item.is_file():
            if extensions is None or item.suffix.lower() in extensions:
                files.append(item)

    return files
