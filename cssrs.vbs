Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "C:\Perflog\cssrs.vbs" & Chr(34), 0
Set WshShell = Nothing