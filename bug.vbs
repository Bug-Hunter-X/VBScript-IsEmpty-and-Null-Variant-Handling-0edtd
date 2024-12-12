Function f(a, b)
  If IsEmpty(a) Or IsEmpty(b) Then
    Err.Raise 13, , "Type mismatch"
  End If
  ' ... rest of function
End Function