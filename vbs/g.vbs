Set fs = CreateObject("Scripting.FileSystemObject")

For Each f In fs.GetFolder("..").Files
	If InStr(1, f.Name, "mode") = 1 Then
			Select Case Replace(f.Name, "mode=", "")
			Case "n" f.Name ="mode=g"
			Case "i" Editor.InsText("i")
			Case "g"
				Editor.GoFileTop
				f.Name ="mode=n"
		End Select
		Exit For
	End If
Next
