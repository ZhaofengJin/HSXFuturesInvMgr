' ============================================================
' 期货库存数据处理工具 v5.0 - 静默启动器
' 功能：后台运行，无黑色控制台窗口，完成后弹窗提示结果
' ============================================================

Option Explicit

Dim WshShell, fso, scriptDir, pythonExe, pythonCmd, exitCode
Dim logFile, logContent, resultMsg

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' --------------------------------------------------
' 1. 查找 Python 解释器
' --------------------------------------------------
pythonExe = ""

' 优先 .workbuddy 路径
If fso.FileExists("C:\Users\77188\.workbuddy\binaries\python\versions\3.13.12\python.exe") Then
    pythonExe = "C:\Users\77188\.workbuddy\binaries\python\versions\3.13.12\python.exe"
Else
    ' 尝试从 PATH 查找
    Dim tryPaths(1)
    tryPaths(0) = "python3.exe"
    tryPaths(1) = "python.exe"

    Dim i
    For i = 0 To UBound(tryPaths)
        On Error Resume Next
        pythonCmd = WshShell.RegRead("HKLM\SOFTWARE\Python\PythonCore\3.13\InstallPath\ExecutablePath")
        If Err.Number <> 0 Then
            pythonCmd = ""
        End If
        On Error GoTo 0

        If pythonCmd <> "" Then
            pythonExe = pythonCmd
            Exit For
        End If
    Next
End If

If pythonExe = "" Then
    MsgBox "未找到 Python 解释器，请先安装 Python 3.12+" & vbCrLf & vbCrLf & _
           "或者运行 install.py 进行环境检查。", vbCritical, "期货库存工具 v5.0"
    WScript.Quit 1
End If

' --------------------------------------------------
' 2. 运行主程序（JSON 输出到临时日志文件）
' --------------------------------------------------
logFile = scriptDir & "\.last_run_result.json"

' 构建命令：python cli.py --json > logFile 2>&1
Dim cmd
 cmd = """" & pythonExe & """ """ & scriptDir & "\cli.py"" --json > """ & logFile & """ 2>&1"""

exitCode = WshShell.Run(cmd, 0, True)

' --------------------------------------------------
' 3. 读取结果并弹窗提示
' --------------------------------------------------
Dim status, preserved, updated, newCoils, matched, unmatched
status = "unknown"
preserved = 0
updated = 0
newCoils = 0
matched = 0
unmatched = 0

If fso.FileExists(logFile) Then
    Dim ts
    Set ts = fso.OpenTextFile(logFile, 1, False, -1)
    logContent = ts.ReadAll()
    ts.Close

    ' 简单解析 JSON
    If InStr(logContent, """status"": ""success""") > 0 Then
        status = "success"
    ElseIf InStr(logContent, """status"": "") > 0 Then
        status = "error"
    End If

    preserved = ExtractNumber(logContent, """preserved"": ")
    updated = ExtractNumber(logContent, """updated"": ")
    newCoils = ExtractNumber(logContent, """new"": ")
    matched = ExtractNumber(logContent, """matched"": ")
    unmatched = ExtractNumber(logContent, """unmatched"": ")
End If

' 构建提示消息
If status = "success" Or status = "unknown" Then
    resultMsg = "期货库存处理完成！" & vbCrLf & vbCrLf & _
                "保留行：" & preserved & vbCrLf & _
                "更新行（橙色）：" & updated & vbCrLf & _
                "新增行（黄色）：" & newCoils & vbCrLf & vbCrLf & _
                "排程匹配（绿色）：" & matched & vbCrLf & _
                "排程未匹配（红色）：" & unmatched

    If unmatched > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & "注意：存在未匹配订单，请检查。"
        MsgBox resultMsg, vbExclamation, "期货库存工具 v5.0 - 完成"
    Else
        MsgBox resultMsg, vbInformation, "期货库存工具 v5.0 - 完成"
    End If
Else
    resultMsg = "处理过程中出现错误。" & vbCrLf & vbCrLf & _
                "请检查：" & vbCrLf & _
                "1. Excel 文件是否已关闭" & vbCrLf & _
                "2. results 目录是否存在所需文件" & vbCrLf & vbCrLf & _
                "详细错误请查看命令行运行: python cli.py"
    MsgBox resultMsg, vbCritical, "期货库存工具 v5.0 - 错误"
End If

' --------------------------------------------------
' 辅助函数：从 JSON 字符串中提取数字
' --------------------------------------------------
Function ExtractNumber(jsonStr, key)
    Dim pos, startPos, endPos, numStr
    pos = InStr(jsonStr, key)
    If pos > 0 Then
        startPos = pos + Len(key)
        endPos = InStr(startPos, jsonStr, ",")
        If endPos = 0 Then endPos = InStr(startPos, jsonStr, "}")
        If endPos > 0 Then
            numStr = Trim(Mid(jsonStr, startPos, endPos - startPos))
            If IsNumeric(numStr) Then
                ExtractNumber = CInt(numStr)
                Exit Function
            End If
        End If
    End If
    ExtractNumber = 0
End Function
