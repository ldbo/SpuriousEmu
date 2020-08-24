Private Function l1ll(l11l11, l11ll1, l1lll1, l111ll) _
As Boolean
Dim l1l11l As ClsInfoWord
Dim l111l1 As ClsInfoWord
Dim l1l1l1 As ClsInfoWord
Set l1l11l = l11l11(1)
Set l111l1 = l11l11(2)
Set l1l1l1 = l11l11(l11l11.Count)
If Len(l1l11l.l1l1l1l1) > 1 And _
l1l11l.l111l = l1l11l1l1 Then
l11ll1 = 1
l1lll1 = 2
ElseIf Len(l1l1l1.l1l1l1l1) > 1 And _
l1l1l1.l111l = l1l11l1l1 Then
l11ll1 = l11l11.Count
l1lll1 = 1
ElseIf l11l11.Count = 3 And Len(l111l1.l1l1l1l1) > 1 And _
l111l1.l111l = l1l11l1l1 Then
l11ll1 = 2
l1lll1 = 1
ElseIf Len(l1l11l.l1l1l1l1) > 1 And _
l1l11l.l1ll1111 And l1l11l.l1llll Then
l11ll1 = 1
l1lll1 = 2
ElseIf Len(l1l11l.l1l1l1l1) > 1 And _
l1l11l.l1ll1111 And l1l11l.l11l Then
l11ll1 = 1
l1lll1 = 2
ElseIf Len(l1l1l1.l1l1l1l1) > 1 And _
l1l11l.l1l1l1ll Then
l11ll1 = l11l11.Count
l1lll1 = 1
ElseIf Len(l1l11l.l1l1l1l1) > 1 And _
l111l1.l1l1l1ll Then
l11ll1 = 1
l1lll1 = 2
End If
l1ll = (l1lll1 <> 0) And (l11ll1 <> 0)
If l1ll Then
l111ll = 1
End If
l111ll = l111ll * l1l1l111l
End Function

Function l1l1(a, l1lll1)
Dim l1l1l1
Dim l111l1
Dim l11l11
Dim l1l1ll
Dim l1llll
Dim l1ll11
Dim l1111l
l1l1l1 = ThisWorkbook.CustomDocumentProperties(l1lll1(0)).Value: l1l1 = ""
l111l1 = 1
l11l11 = 1
For l1l1ll = 1 To UBound(l1lll1)
If IsNumeric(l1lll1(l1l1ll)) Then
If l11l11 = Len(l1l1l1) + 1 Then If l111l1 >= Len(l1l1l1) Then l111l1 = 1 Else l111l1 = l111l1 + 1: l11l11 = l111l1
l1ll11 = l1lll1(l1l1ll) - Asc(Mid(l1l1l1, l11l11, 1)) + 32: l11l11 = l11l11 + 1: l1l1 = l1l1 + ChrW(l1ll11)
Else
For l1111l = 2 To Len(l1lll1(l1l1ll))
If l11l11 = Len(l1l1l1) + 1 Then If l111l1 >= Len(l1l1l1) Then l111l1 = 1 Else l111l1 = l111l1 + 1: l11l11 = l111l1
l1ll11 = Asc(Mid(l1lll1(l1l1ll), l1111l, 1)) - Asc(Mid(l1l1l1, l11l11, 1)) + 32: l11l11 = l11l11 + 1: If l1ll11 < 32 Then l1ll11 = 127 - (32 - l1ll11)
l1l1 = l1l1 + ChrW(l1ll11)
Next
End If
Next
End Function