' Fan Control Launcher — auto-elevates to Administrator, no console window
' Double-click this file to launch the Fan Control GUI

Set objShell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the folder this script is in
strFolder = fso.GetParentFolderName(WScript.ScriptFullName)
strScript = strFolder & "\fan_control_gui.py"

' Find pythonw.exe (no console window)
Set wshShell = CreateObject("WScript.Shell")

' Try to find pythonw.exe via py launcher first, then PATH
Dim pythonw
pythonw = "pythonw.exe"

' Launch with admin elevation ("runas"), hidden window (0)
objShell.ShellExecute pythonw, """" & strScript & """", strFolder, "runas", 0
