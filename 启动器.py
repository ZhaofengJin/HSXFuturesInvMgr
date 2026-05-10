"""
启动器 - 无窗口静默运行期货库存工具v4
"""
import subprocess
import os

# 获取脚本目录
script_dir = os.path.dirname(os.path.abspath(__file__))
main_script = os.path.join(script_dir, "期货库存工具_v4.py")
python_exe = r"C:\Users\77188\.workbuddy\binaries\python\versions\3.13.12\python.exe"

# 直接运行Python脚本（无窗口）
subprocess.Popen(
    [python_exe, main_script],
    cwd=script_dir,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL
)
