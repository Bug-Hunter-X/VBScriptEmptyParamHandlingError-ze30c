Function MyFunction(param1)
  If IsEmpty(param1) Then
    ' Handle empty parameter
    Exit Function 'This line was missing, causing the script to continue and throw an error.
  End If
  ' ... rest of the function
End Function