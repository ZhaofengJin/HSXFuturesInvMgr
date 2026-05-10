@echo off
cd /d "C:\Users\77188\Desktop\期货库存"
"C:\Users\77188\AppData\Local\Programs\Python\Python312\python.exe" test3.py > output.log 2>&1
type output.log
pause
