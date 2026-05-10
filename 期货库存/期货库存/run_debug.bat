@echo off
chcp 65001 >nul
cd /d "C:\Users\77188\Desktop\期货库存"
"C:\Users\77188\AppData\Local\Programs\Python\Python312\python.exe" debug_data.py > debug_output.log 2>&1
type debug_output.log
