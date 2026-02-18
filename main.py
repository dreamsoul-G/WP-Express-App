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
import traceback
from pathlib import Path


def print_banner():
    banner = "WP快通（文档处理）软件 v1.1.0.2026217"
    print(banner)
    print("功能：Word/PDF文档的创建、合并、拆分、转换")
    print("=" * 60)


def setup_environment():
    """设置运行环境"""
    # 获取项目根目录
    root_dir = Path(__file__).parent.absolute()

    # 添加src目录到Python路径
    src_path = root_dir / "src"
    sys.path.insert(0, str(src_path))
    sys.path.insert(0, str(root_dir))

    return root_dir


def check_dependencies():
    """检查依赖包"""
    required = {
        'python-docx': 'docx',
        'pypdf': 'pypdf',
        'reportlab': 'reportlab',
        'Pillow': 'PIL',
    }

    if sys.platform.startswith('win'):
        required['pywin32'] = 'win32com'

    missing = []
    for pkg_name, module_name in required.items():
        try:
            __import__(module_name)
            print(f"✓ {pkg_name}")
        except ImportError:
            missing.append(pkg_name)
            print(f"✗ {pkg_name}")

    if missing:
        print(f"\n缺少依赖包，请运行:")
        print(f"pip install {' '.join(missing)}")
        if sys.platform.startswith('win') and 'pywin32' in missing:
            print("\n注意：Windows用户可能需要以管理员身份运行命令提示符")
        return False
    return True


def create_directories(root_dir):
    """创建必要的目录"""
    dirs = [
        root_dir / "logs",
        root_dir / "assets" / "icons",
        root_dir / "assets" / "images"
    ]

    for dir_path in dirs:
        dir_path.mkdir(parents=True, exist_ok=True)
        print(f"✓ 目录就绪: {dir_path.relative_to(root_dir)}")


def setup_logging(root_dir=None):
    """设置日志系统"""
    import logging

    # 确定日志文件路径：打包环境用当前目录，开发环境用项目目录
    if root_dir is None or getattr(sys, 'frozen', False):
        # 用户模式：日志放在 .exe 同级目录下
        log_file = Path(sys.executable).parent / "WP_Express.log"
    else:
        # 开发模式：日志放在项目 logs 目录下
        log_file = root_dir / "logs" / "WP_Express.log"

    # 确保日志文件的目录存在
    log_file.parent.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler() if not getattr(sys, 'frozen', False) else logging.NullHandler()
        ]
    )

    return logging.getLogger("WP_Express")


def launch_gui():
    """启动GUI界面"""
    # 如果是打包环境， 进行非常明确的路径设置
    if getattr(sys, 'frozen', False):
        print("【用户模式】正在设置模块搜索路径...")
        base_path = Path(sys.executable).parent  # .exe所在目录

        possible_src_paths = [
            base_path / 'src',
            base_path / '..' / 'src',
            base_path,
            ]

        for p in possible_src_paths:
            p = p.resolve()
            if p.exists():
                str_path = str(p)
                if str_path not in sys.path:
                    sys.path.insert(0, str_path)
                    print(f"  已添加路径: {str_path}")

        # 打印当前所有路径供调试
        print("【当前模块搜索路径】:")
        for idx, path in enumerate(sys.path[:5]):  # 显示前几个关键路径
            print(f"  {idx}: {path}")
        print("  ...")

    # 现在尝试导入，使用更灵活的方式
    main_window = None
    last_error = None

    # 尝试从可能的不同模块名导入
    for module_name in ['src.main_window', 'main_window']:
        try:
            print(f"尝试导入模块: '{module_name}'")
            module = __import__(module_name, fromlist=['MainWindow'])
            main_window = getattr(module, 'MainWindow')
            print(f"✓ 成功从 '{module_name}' 导入 MainWindow")
            break
        except ImportError as e:
            last_error = e
            print(f"  从 '{module_name}' 导入失败: {e}")
            continue

    if main_window is None:
        error_msg = f"无法加载程序界面。所有导入尝试均失败。\n最后错误: {last_error}"
        print(f"✗ {error_msg}")
        # 尝试记录日志
        try:
            import logging
            logger = logging.getLogger("WPQuickPass")
            logger.error(error_msg, exc_info=True)
        except:
            pass
        # 弹窗提示
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("启动错误",
                                 "核心界面模块加载失败。\n这可能是因为打包不完整。\n请检查dist文件夹内是否有'src'目录。")
            root.destroy()
        except:
            pass
        return False

    # 运行主程序
    try:
        app = main_window()
        print("✓ GUI应用程序启动成功，进入主循环。")
        app.run()
        return True
    except Exception as e:
        error_msg = f"GUI运行时错误: {e}"
        print(f"✗ {error_msg}")
        try:
            import logging
            logger = logging.getLogger("WPQuickPass")
            logger.error(error_msg, exc_info=True)
        except:
            pass
        return False


def main():
    """主函数"""
    try:
        # 1. 显示横幅
        print_banner()

        # 2. 关键判断：如果是打包后的环境，则跳过控制台依赖检查
        if getattr(sys, 'frozen', False):
            # 用户模式 - 简化启动流程
            print("检测到打包环境，跳过依赖检查，直接启动...")

            # 设置日志（使用默认路径，与.exe文件同目录）
            logger = setup_logging()
            logger.info("应用程序启动（用户模式）")

            # 直接进入GUI启动
            gui_started = launch_gui()
            return 0 if gui_started else 1

        else:
            # 开发模式 - 完整启动流程
            # 2.1 设置环境
            root_dir = setup_environment()
            print(f"项目目录: {root_dir}")

            # 2.2 检查依赖
            print("\n检查依赖包:")
            if not check_dependencies():
                input("\n按Enter键退出...")
                return 1

            # 2.3 创建目录
            print("\n检查目录结构:")
            create_directories(root_dir)

            # 2.4 设置日志
            logger = setup_logging(root_dir)
            logger.info("应用程序启动（开发模式）")

            # 2.5 启动GUI
            print("\n启动应用程序...")
            return launch_gui()

    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
        return 0
    except Exception as e:
        print(f"\n✗ 程序错误: {e}")
        traceback.print_exc()

        # 记录错误到日志（如果日志系统已初始化）
        try:
            import logging
            logger = logging.getLogger("WPQuickPass")
            logger.error(f"程序启动失败: {e}", exc_info=True)
        except:
            pass

        return 1


if __name__ == "__main__":
    # 确保UTF-8编码
    if sys.platform.startswith('win'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

    # 运行主程序
    exit_code = main()

    # 程序执行完毕
    print("\n" + "=" * 60)

    # 用户模式下直接退出，避免无控制台时input()卡住
    if getattr(sys, 'frozen', False):
        print("\n程序执行完毕，窗口即将关闭。")
        sys.exit(exit_code)  # 用户模式直接退出，无延迟
    else:
        print("\n程序执行完毕，按Enter键退出...")
        try:
            input()
        except (KeyboardInterrupt, EOFError):
            pass

    sys.exit(exit_code)
