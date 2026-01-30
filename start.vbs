' Shift Optimizer Application Launcher
' This script launches the application without showing a command prompt window

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = strScriptPath

' Check if venv exists and dependencies are installed
strVenvPath = strScriptPath & "\venv"
strDepsFlag = strScriptPath & "\.deps_installed"
strPythonW = strVenvPath & "\Scripts\pythonw.exe"
strPython = strVenvPath & "\Scripts\python.exe"

' If venv doesn't exist or dependencies not installed, run admin_setup.bat first (with console)
If Not objFSO.FolderExists(strVenvPath) Or Not objFSO.FileExists(strDepsFlag) Then
    ' Need to run setup with console visible
    objShell.Run "cmd /c """ & strScriptPath & "\admin_setup.bat""", 1, True
Else
    ' Run the application without console window
    If objFSO.FileExists(strPythonW) Then
        objShell.Run """" & strPythonW & """ """ & strScriptPath & "\app.py""", 0, False
    ElseIf objFSO.FileExists(strPython) Then
        ' Fallback to python.exe if pythonw.exe not found
        objShell.Run """" & strPython & """ """ & strScriptPath & "\app.py""", 0, False
    Else
        MsgBox "Python not found in virtual environment." & vbCrLf & vbCrLf & "Please run admin_setup.bat first to set up the environment.", vbExclamation, "Shift Optimizer"
    End If
End If
