' -*- coding: utf-8 -*-
' VBS启动器 - 无窗口静默运行期货库存工具v4

Set WshShell = CreateObject("WScript.Shell")
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)
pythonScript = scriptDir & "\期货库存工具_v4.py"
pythonExe = "C:\Users\77188\AppData\Local\Programs\Python\Python312\python.exe"
Set oExec = WshShell.Exec("cmd /c """ & pythonExe & """ """ & pythonScript & """")

Do While oExec.Status = 0
    WScript.Sleep 100
Loop
