' ====================================================
' تشغيل تطبيق ألمدون — الثانوية الإعدادية ألمدون
' انقر مزدوج مباشرة لتشغيل التطبيق بدون نافذة CMD
' ====================================================
Dim fso, shell, appDir, scriptPath, pythonwPath

Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appDir     = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = appDir & "\main.py"

' Try pythonw first (no console window), fall back to python
Dim cmd
On Error Resume Next

' Check if pythonw is available
shell.Run "pythonw --version", 0, True
If Err.Number = 0 Then
    cmd = "pythonw """ & scriptPath & """"
Else
    Err.Clear
    cmd = "python """ & scriptPath & """"
End If

On Error GoTo 0

' Launch the app silently (0 = hidden window, False = don't wait)
shell.Run cmd, 0, False
