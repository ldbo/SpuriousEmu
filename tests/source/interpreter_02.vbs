' Draft file
Sub Main()
    VBAEnv.VBA.Interaction.msgBox "Hey"
    Dim hello
    Dim plop As New Integer
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

    Exexexexe
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

Sub Exexexexe()
    Dim command
    Set command = "plop plop plop /"
    Dim shell As Object
    Set shell = WScript.CreateObject("WScript.Shell")
    shell.Run(command)
End Sub