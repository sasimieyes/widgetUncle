' start_widget.vbs - run BtBatteryWidget.ps1 without a console window
Option Explicit
Dim fso, sh, base, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & base & "\BtBatteryWidget.ps1"""
sh.Run cmd, 0, False
