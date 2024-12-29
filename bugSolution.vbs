Function f(x)
  If x = 0 Then
    f = 1
  ElseIf x < 0 Then
    f = 0 ' Handle negative input (optional)
  Else
    f = x * f(x - 1)
  End If
End Function
MsgBox f(5) ' Test with a valid input
MsgBox f(0) 'Test base case
MsgBox f(-1) 'Test negative input (optional)