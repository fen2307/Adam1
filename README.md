# Adam1
Sub EBITDAFIND()


Dim CompanyRange As Integer
Dim EbitdaDataArray() As Double
Dim x As Double
Dim y As Double
Dim z As Integer
z = 1
For i = 17 To 10000
If Range("G" & i).Value = "" Then
Exit For
End If
ComapnyRange = i
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









NOWE NOWENOWE


 Worksheets("Ex_1").Activate
 Dim Srednia As Double
 Dim LiczbaFirm As Integer
 Srednia = 100000000
 Dim InputSrednia As Double
 InputSrednia = Range("C14").Value
 Dim Dane(2 To 505) As Double
  Dim DaneNazwy(2 To 505) As String
 Worksheets("Data_Ex_1").Activate
 For i = 2 To 505
 Dane(i) = Range("M" & i).Value
 Next i
 For i = 2 To 505
 DaneNazwy(i) = Range("C" & i).Value
 Next i

 
 While Srednia <> InputSrednia
 Randomize
 LiczbaFirm = Int((505 - 2 + 1) * Rnd + lowerbound)
 
 Dim Tablica(1 To LiczbaFirm) As Double
 
     For i = 1 To LiczbaFirm
     
     Randomize
     Tablica(i) = Int((505 - 2 + 1) * Rnd + lowerbound)
     
    
            For Z = 1 To LiczbaFirm
            If Tablica(i) = Tablica(Z) And i <> Z Then
            
            Randomize
            Tablica(i) = Int((505 - 2 + 1) * Rnd + lowerbound)
            
            Z = 1
            End If
            Next Z
     Next i
     
        Dim Suma As Double
        Suma = 0
        For x = 1 To LiczbaFirm
        Suma = Suma + Dane(Tablica(x))
        Next x
        
 Srednia = Suma / LiczbaFirm
 
        If Srednia = InputSrednia Then
        Worksheets("Ex_1").Activate
               For h = 1 To LiczbaFirm
               Dim w As Integer
               w = 16 + i
               Range("B" & w).Value = DaneNazwy(Tablica(h))
               Next h
        End If
 Wend
