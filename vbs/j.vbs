Set fs = CreateObject("Scripting.FileSystemObject")

For Each f In fs.GetFolder("..").Files
	If InStr(1, f.Name, "mode") = 1 Then
		Select Case Replace(f.Name, "mode=", "")
			Case "n" Editor.Down
			Case "i" Editor.InsText("j")
		End Select
		Exit For
	End If
Next
