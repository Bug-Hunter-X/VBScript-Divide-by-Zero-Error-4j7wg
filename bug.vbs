Function f(a, b)
  If a = 0 Then Exit Function
  If b = 0 Then Exit Function
  f = a / b
End Function

' Calling the function with b = 0 will cause a divide by zero error
MsgBox f(10, 0)