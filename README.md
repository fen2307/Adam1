Sub EBITDAFIND()


Dim CompanyRange As Integer
Dim EbitdaDataArray(1 To 10000) As Double
Dim x As Double
Dim y As Double
Dim z As Integer
z = 1
For i = 17 To 10000
If Range("G" & i).Value <> "" Then
CompanyRange = i
Else
Exit For
End If

Next i

For i = 17 To CompanyRange

EbitdaDataArray(z) = Range("g" & i).Value
z = z + 1
Next i

For i = CompanyRange + 5 To CompanyRange + CompanyRange + 5

z = 1
While z < CompanyRange - 16
x = Range("f" & i).Value
y = Range("e" & i).Value
If x / y = EbitdaDataArray(z) Then
Range("g" & i).Value = EbitdaDataArray(z)
End If
z = z + 1
Wend

Next i


End Sub
