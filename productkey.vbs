Set WshShell = CreateObject("WScript.Shell")
cdkey = ConvertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))
MsgBox  "Copied to clipboard:"& chr(13) & cdkey
WshShell.Run "cmd.exe /c echo " & cdkey & " | clip"

Function ConvertToKey(Key)
	Const KeyOffset = 52
	i = 28
	Chars = "BCDFGHJKMPQRTVWXY2346789"
	Do
		Cur = 0
		x = 14
		Do
			Cur = Cur * 256
			Cur = Key(x + KeyOffset) + Cur
			Key(x + KeyOffset) = (Cur \ 24) And 255
			Cur = Cur Mod 24
			x = x -1
		Loop While x >= 0
		i = i - 1
		KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
		If (((29 - i) Mod 6) = 0) And (i <> -1) Then
			i = i -1
			KeyOutput = "-" & KeyOutput
		End If
	Loop While i >= 0
	ConvertToKey = KeyOutput
End Function
