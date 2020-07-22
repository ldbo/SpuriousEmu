' Test functions symbol extraction
Function CalculateSquareRoot(NumberArg)
 If NumberArg Then ' Evaluate argument.
  Exit Function ' Exit to calling procedure.
 Else
  Set CalculateSquareRoot = Sqr(NumberArg) ' Return square root.
 End If
End Function