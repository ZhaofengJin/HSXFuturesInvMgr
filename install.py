"""
环境检查与自动安装脚本

功能：
1. 检查 Python 版本（>= 3.12）
2. 检查 openpyxl 是否安装
3. 自动安装缺失依赖
4. 验证目录结构

使用方式：
    python install.py
"""

import sys
import subprocess
import os


def check_python_version():
    """检查 Python 版本"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 12):
        print(f"[失败] Python 版本过低: {version.major}.{version.minor}.{version.micro}")
        print("[提示] 需要 Python 3.12 或更高版本")
        return False
    print(f"[通过] Python 版本: {version.major}.{version.minor}.{version.micro}")
    return True


def check_module(module_name, install_name=None):
    """检查模块是否安装，未安装则自动安装"""
    if install_name is None:
        install_name = module_name

    try:
        __import__(module_name)
        print(f"[通过] {module_name} 已安装")
        return True
    except ImportError:
        print(f"[缺失] {module_name} 未安装，正在自动安装...")
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", install_name],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print(f"[通过] {module_name} 安装成功")
            return True
        except subprocess.CalledProcessError:
            print(f"[失败] {module_name} 安装失败，请手动运行: pip install {install_name}")
            return False


def check_directory_structure():
    """检查项目目录结构"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    required_files = [
        "main.py", "cli.py", "config.py", "utils.py",
        "models.py", "excel_handler.py", "processor.py",
    ]

    all_ok = True
    for fname in required_files:
        fpath = os.path.join(script_dir, fname)
        if os.path.exists(fpath):
            print(f"[通过] {fname}")
        else:
            print(f"[缺失] {fname}")
            all_ok = False

    # 检查 results 目录
    results_dir = os.path.join(script_dir, "results")
    if os.path.isdir(results_dir):
        print(f"[通过] results/ 目录存在")
    else:
        print(f"[提示] results/ 目录不存在，运行时将自动创建")

    return all_ok


def main():
    print("=" * 60)
    print("HSXFuturesInvMgr 环境检查与安装")
    print("=" * 60)
    print()

    ok = True
    ok = check_python_version() and ok
    print()
    ok = check_module("openpyxl") and ok
    ok = check_module("pytest", "pytest") and ok
    print()
    ok = check_directory_structure() and ok
    print()

    print("=" * 60)
    if ok:
        print("环境检查全部通过！可以开始使用。")
        print()
        print("运行方式:")
        print("  1. 双击 run.bat（推荐，有进度窗口）")
        print("  2. 双击 run_silent.vbs（静默运行，完成后弹窗提示）")
        print("  3. 命令行: python cli.py")
        print("  4. AI 调用: python cli.py --json")
    else:
        print("环境检查未通过，请根据上方提示修复后重试。")
    print("=" * 60)

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
