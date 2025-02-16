Function MyFunction(param1, param2)
  On Error Resume Next 'Enable error handling
  If IsNumeric(param1) And IsNumeric(param2) Then
    If param2 = 0 Then
      result = "Division by zero error"
    Else
      result = param1 / param2
    End If
  Else
    result = "Type mismatch error"
  End If
  MyFunction = result
End Function