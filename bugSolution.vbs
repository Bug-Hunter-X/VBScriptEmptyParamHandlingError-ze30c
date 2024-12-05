Function MyFunction(param1)
  If IsEmpty(param1) Then
    ' Handle empty parameter
    Exit Function
  End If
  ' ... rest of the function
  result = param1 * 2
  MyFunction = result
End Function

'Example usage
d = MyFunction(10)
MsgBox d

e = MyFunction()
MsgBox e 'Will display nothing because of the correct handling of empty parameter.