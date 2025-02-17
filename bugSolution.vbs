Function f(x)
  Dim result
  result = 1
  If x > 0 Then
    For i = 1 To x
      result = result * i
    Next
  End If
  f = result
End Function
MsgBox f(5) 