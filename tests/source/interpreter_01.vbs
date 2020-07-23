Sub Main
    Dim var1
    Set var1 = 12
    msgBox var1
    Set var1 = Return1()
    msgBox Return1()
    msgBox ReturnArg(12)
End Sub

Function Return1()
    Set Return1 = 1 * 1
    msgBox Return1
End Function

Function ReturnArg(arg)
    Set ReturnArg = 23
End Function