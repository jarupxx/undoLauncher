Option Explicit

Const vbHide = 0
Const vbNormalFocus = 1
Const vbMinimizedFocus = 2
Const vbMaximizedFocus = 3
Const vbNormalNoFocus = 4
Const vbMinimizedNoFocus = 6
Dim appPath, filePath

'-----------------------------------------------------------
'ここをアプリのパスに書き換えてください
appPath = """C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"""
'-----------------------------------------------------------

TerminateProcess("explorer.exe")

Sub TerminateProcess(ProcessName)
    ' WshShellオブジェクトの作成
    Dim wShell
    Set wShell = CreateObject("WScript.Shell")

    ' 名前付き引数の定義
    Const vbHide = 0
    Const vbNormalFocus = 1
    Const vbMinimizedFocus = 2
    Const vbMaximizedFocus = 3
    Const vbNormalNoFocus = 4
    Const vbMinimizedNoFocus = 6

    ' アプリケーションの起動 Chr(34)はダブルクォーテーション
    filePath = Chr(34) & WScript.Arguments.Item(0) & Chr(34)
    If Wscript.Arguments.Count = 0 Then
       wShell.Run appPath, vbNormalFocus, False
    ElseIf Wscript.Arguments.Count = 1 Then
       wShell.Run appPath & filePath, vbNormalFocus, False
    Else
       MsgBox "引数は一つにして下さい。", 48, "undoLauncher"
       Wscript.Quit
    End If

    ' Explorerの再起動
    Dim Service,QfeSet,Qfe
    Set Service = CreateObject("WbemScripting.SWbemLocator").ConnectServer
    Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='" & ProcessName & "'")
    For Each Qfe In QfeSet
        Qfe.Terminate
    Next
End Sub
