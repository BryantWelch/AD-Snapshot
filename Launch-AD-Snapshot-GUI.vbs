Option Explicit

' AD Snapshot GUI Launcher
' This script launches the AD Snapshot GUI application without showing the PowerShell console window

' Get the script directory
Dim objFSO, objShell, strScriptDir, strPSCommand, intWindowStyle

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Get the directory where this VBS file is located
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Set the PowerShell command to run
strPSCommand = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strScriptDir & "\AD-Snapshot-GUI-Main.ps1"""

' Run the command without showing the window (0 = hidden)
intWindowStyle = 0
objShell.Run strPSCommand, intWindowStyle, False

' Clean up
Set objShell = Nothing
Set objFSO = Nothing
