' Draft file
Sub Main()
    Dim hello
    Set hello = "hello, "
    For i = 0 To 12 Step 5
        msgBox "i"
        msgBox i
        Set hello = hello + ", " & "helo"
        For j = 5 To 0 Step 0 - 2
            msgBox j
        Next
        msgBox "--------"
    Next i
    msgBox ReturnArg(hello)

    For i = 0 To 12
        msgBox Oddity(i)
    Next i
End Sub

Function ReturnArg(arg)
    Set ReturnArg = arg
End Function

Function Oddity(n)
    If n Mod 2 = 0 Then
        Set Oddity = "Not that odd"
    Else
        Set Oddity= "High level of oddity"
    End If
End Function