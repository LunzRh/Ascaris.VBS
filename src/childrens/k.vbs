Dim Shell
Set Shell = CreateObject("WScript.Shell")

While True
	Shell.SendKeys "{CAPSLOCK}"
	WScript.Sleep 200
Wend
