Call jean
Call jean(a, b, c)
msgBox
msgBox jacques(12)
msgBox abcd, efgh
Dim Shared variable
Dim a As Integer, b, c%
Static abcd As Long, def As String
Const a = 1, b As Long = 12, c! = 30
ReDim MyArray()
ReDim Array1(), Array2(), Array3()
ReDim Preserve Array1(), Array2(), Array3()
ReDim arr(1, 2 To 12, 3)
ReDim arr(12) As Integer
ReDim array(12 To 15) As New MyClass
Erase a
Erase a, b, c
Mid(str, 12, 3) = "abc"
MidB$(name, 0) = str + "12"
LSet b = "Hey"
RSet a = "abcd"
Let a = b.c() + 12
b = 5 + 12
Set a.b(12, 14) = 12
On Error Goto abcd : On Error Resume Next
On Error Goto 123 :
Resume Next : Resume 12 : Resume abcd : Error abcd : Error 123
Open hey As 12
Open "hi my name is" For Binary Access Read Lock Read As (12 + "abcd") _
Len = 1
Open s Access Write Lock Read Write As f
Open s Access Read Write Lock Read As f
Open s Lock Write As #f
Open s Shared As # f
Open s As f Len=12
Open s As f Len = 15
Reset Close hey
Reset
Close abcd + 5, abcd(), efg
Close
Seek ab, 123 + 5
Lock c
Lock d, To b
Lock f, a To 12 + b
Unlock g, b
Line Input #123, ab
Width #145, abcd
Print #1,
Print #1, "This is a test"
Print #1, "Zone 1",Tab; "Zone 2"
Print #1, "Hello" ;" " , "World"
Print #1, Spc(5); "5 leading spaces "
Print #1, Tab(10) ; "Hello"
Write #1, "Hello World", 234
Write #1,
Write #1, MyBool ; " is a Boolean value"
Write #1, MyDate ; " is a date"
Write #1, MyNull ; " is a null value"
Write #1, MyError ; " is an error value"
Input #1, MyString, MyNumber
Put #1, RecordNumber, MyRecord
Get #4,,FileBuffer
Get #1, Position, MyRecord

While a = 12
    msgBox Hey
    a = a / 2
Wend

While False

Wend

For i = 0 To 12

Next i

For j = 12 To 16 Step 5
    msgBox j
Next

For nasty = 0 to 1
    For amnesty = 12 To 10
        msgBox amnesty
    Next amnesty
    Exit For
Next nasty

For Each carrot in bag
    eat carrot
Next carrot

Do
    msgBox "Hey, when will I stop ?"
    Exit Do
Loop

Do While a > 12
    a = a - 1
Loop

Do
    b = b + 1
Loop Until b > 10

If a Then
    printf "Youhou \\%s", hey
EndIf

If fifi Then
    Let fofo = false
ElseIf fofo Then
    Let fifi = True
Elseif trou Then
    Let coupcoup = False
Elseif coupcoup Then
EndIf

If a = 10 Then
    msgBox "Game over"
ElseIf b > 10 Then
Else
    msgBox "You win"
EndIf

If True Then msgBox 12
If a < 12 Then Let a = a + 12 Else msgBox "Nothing to do here"
If True Then If False Then msgBox "Trou" Else msgBox "False"

Select Case number
Case Is <= 12
    msgBox "rikiki"
Case Is = 42
    msgBox "duh"
Case Else
    msgBox "Duh ?"
End Select

Select Case a
Case Else
    Let a = 12
    Dim b As String
End Select

Select Case cas
Case Is <= 15, 17, 26, Is >= 100
    msgBox "Ho"
Case 20 To 22
    msgBox "D"
End Select

Stop
GoTo 12
Go To abcd
On ex GoTo 12, my_line, 1
On expr GoTo label
On Number GoSub Sub1, Sub2
On Number GoSub Sub12
GoSub subway
Return
Exit Sub
Exit Function
Exit Property
RaiseEvent LogonCompleted ("AntoineJan")
