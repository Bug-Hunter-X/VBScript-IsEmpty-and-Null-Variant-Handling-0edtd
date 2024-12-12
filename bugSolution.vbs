Function f(a, b)
  If IsNull(a) Or IsNull(b) Then
    Err.Raise 13, , "Null value encountered"
  ElseIf IsEmpty(a) Or IsEmpty(b) Then
    Err.Raise 13, , "Empty value encountered"
  End If
  ' ... rest of function
End Function