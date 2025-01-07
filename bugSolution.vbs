Function f(a, b)
  If VarType(a) = vbEmpty Then
    a = 0
  End If
  If VarType(b) = vbEmpty Then
    b = 0
  End If
  f = a + b
End Function

MsgBox f(1, Empty) 'This will now correctly output 1.