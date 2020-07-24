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
End Sub

Function ReturnArg(arg)
    Set ReturnArg = arg
End Function
