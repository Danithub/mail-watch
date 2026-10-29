' Launch watch.ps1 without a visible console window (Task Scheduler entry point).
Set sh = CreateObject("WScript.Shell")
dir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & dir & "\watch.ps1""", 0, True
