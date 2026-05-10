"""
启动器 - 无窗口静默运行期货库存工具 v5

功能：
- 自动查找项目目录并启动 main.py
- 支持 Windows 和 macOS/Linux 环境
- 静默运行，无控制台窗口弹出（Windows）
"""

import os
import sys
import subprocess


def find_python_executable():
    """查找可用的 Python 解释器"""
    # 优先查找 .workbuddy 中的 Python（Windows 特定路径）
    workbuddy_python = r"C:\Users\77188\.workbuddy\binaries\python\versions\3.13.12\python.exe"
    if os.path.exists(workbuddy_python):
        return workbuddy_python

    # 其次查找系统 PATH 中的 python
    for cmd in ["python3", "python"]:
        python_path = subprocess.run(
            [cmd, "--version"],
            capture_output=True,
            text=True,
        )
        if python_path.returncode == 0:
            return cmd

    return None


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(script_dir, "main.py")

    if not os.path.exists(main_script):
        print(f"Error: {main_script} not found")
        sys.exit(1)

    python_exe = find_python_executable()
    if not python_exe:
        print("Error: Python interpreter not found")
        sys.exit(1)

    # Windows: 无窗口运行；其他平台：正常运行
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        subprocess.Popen(
            [python_exe, main_script],
            cwd=script_dir,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            startupinfo=startupinfo,
        )
    else:
        subprocess.run(
            [python_exe, main_script],
            cwd=script_dir,
        )


if __name__ == "__main__":
    main()
