Set fs = CreateObject("Scripting.FileSystemObject")

For Each f In fs.GetFolder("..").Files
	If InStr(1, f.Name, "mode") = 1 Then
		Select Case Replace(f.Name, "mode=", "")
			Case "n"
				Editor.Right
				f.Name = "mode=i"
			Case "i" Editor.InsText("a")
		End Select
		Exit For
	End If
Next
