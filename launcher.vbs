Option Explicit

Const vbHide = 0
Const vbNormalFocus = 1
Const vbMinimizedFocus = 2
Const vbMaximizedFocus = 3
Const vbNormalNoFocus = 4
Const vbMinimizedNoFocus = 6
Dim appPath, filePath

'-----------------------------------------------------------
'�������A�v���̃p�X�ɏ��������Ă�������
appPath = """C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"""
'-----------------------------------------------------------

TerminateProcess("explorer.exe")

Sub TerminateProcess(ProcessName)
    ' WshShell�I�u�W�F�N�g�̍쐬
    Dim wShell
    Set wShell = CreateObject("WScript.Shell")

    ' ���O�t�������̒�`
    Const vbHide = 0
    Const vbNormalFocus = 1
    Const vbMinimizedFocus = 2
    Const vbMaximizedFocus = 3
    Const vbNormalNoFocus = 4
    Const vbMinimizedNoFocus = 6

    ' �A�v���P�[�V�����̋N�� Chr(34)�̓_�u���N�H�[�e�[�V����
    filePath = Chr(34) & WScript.Arguments.Item(0) & Chr(34)
    If Wscript.Arguments.Count = 0 Then
       wShell.Run appPath, vbNormalFocus, False
    ElseIf Wscript.Arguments.Count = 1 Then
       wShell.Run appPath & filePath, vbNormalFocus, False
    Else
       MsgBox "�����͈�ɂ��ĉ������B", 48, "undoLauncher"
       Wscript.Quit
    End If

    ' Explorer�̍ċN��
    Dim Service,QfeSet,Qfe
    Set Service = CreateObject("WbemScripting.SWbemLocator").ConnectServer
    Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='" & ProcessName & "'")
    For Each Qfe In QfeSet
        Qfe.Terminate
    Next
End Sub
