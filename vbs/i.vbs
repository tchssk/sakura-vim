Set fs = CreateObject("Scripting.FileSystemObject")

For Each f In fs.GetFolder("..").Files
	If InStr(1, f.Name, "mode") = 1 Then
		Select Case Replace(f.Name, "mode=", "")
			Case "n" f.Name ="mode=i"
			Case "i" Editor.InsText("i")
		End Select
		Exit For
	End If
Next
