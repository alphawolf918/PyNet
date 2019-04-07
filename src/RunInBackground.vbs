Dim wsh
Set wsh = CreateObject("WScript.Shell")
wsh.Run "C:\Tasks\PyNet.Py", 0, False
Set wsh = Nothing