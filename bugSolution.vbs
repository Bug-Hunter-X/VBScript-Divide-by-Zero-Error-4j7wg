Function f(a, b)
  If a = 0 Or b = 0 Then
    f = 0 ' Or handle the error in a more appropriate way
    Exit Function
  End If
  f = a / b
End Function

' Calling the function with b = 0 will now return 0 instead of causing an error.
MsgBox f(10, 0)