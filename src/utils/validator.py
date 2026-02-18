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

import re
from pathlib import Path


def validate_file_path(path: Path, must_exist: bool = True) -> tuple:
    """验证文件路径"""
    if not path or str(path).strip() == "":
        return False, "路径不能为空"

    try:
        path_obj = Path(path)

        if must_exist and not path_obj.exists():
            return False, "文件不存在"

        if not must_exist:
            # 检查父目录是否存在
            if not path_obj.parent.exists():
                return False, "父目录不存在"

        return True, "验证通过"

    except Exception as e:
        return False, f"路径无效: {e}"


def validate_directory_path(path: Path, must_exist: bool = True) -> tuple:
    """验证目录路径"""
    if not path or str(path).strip() == "":
        return False, "路径不能为空"

    try:
        path_obj = Path(path)

        if must_exist and not path_obj.exists():
            return False, "目录不存在"

        if must_exist and not path_obj.is_dir():
            return False, "路径不是目录"

        return True, "验证通过"

    except Exception as e:
        return False, f"路径无效: {e}"


def validate_file_name(name: str) -> tuple:
    """验证文件名"""
    if not name or name.strip() == "":
        return False, "文件名不能为空"

    # 检查非法字符
    illegal_chars = r'[<>:"/\\|?*]'
    if re.search(illegal_chars, name):
        return False, "文件名包含非法字符"

    # 检查长度
    if len(name) > 255:
        return False, "文件名过长"

    # 检查保留名称
    reserved_names = [
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    ]

    if name.upper() in reserved_names:
        return False, "文件名是系统保留名称"

    return True, "验证通过"


def sanitize_file_name(name: str) -> str:
    """清理文件名"""
    if not name:
        return "未命名文件"

    # 移除非法字符
    illegal_chars = r'[<>:"/\\|?*]'
    name = re.sub(illegal_chars, '_', name)

    # 移除首尾空格
    name = name.strip()

    # 如果为空，使用默认名称
    if not name:
        name = "未命名文件"

    return name


def validate_number(value: str, min_val: int = 1, max_val: int = 1000) -> tuple:
    """验证数字"""
    try:
        num = int(value)

        if num < min_val:
            return False, f"数值不能小于 {min_val}"

        if num > max_val:
            return False, f"数值不能大于 {max_val}"

        return True, "验证通过"

    except ValueError:
        return False, "请输入有效的数字"
