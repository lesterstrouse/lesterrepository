Attribute VB_Name = "Module1"
Sub stockmacro()
'find last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'set variables initially
Dim i, j As Long
Dim greatpi, greatpd, openp, greatv, totvol As Double
tick = Cells(2, 1).Value
openp = Cells(2, 3).Value
greatpi = 0
greatpd = 0
greatv = 0
totvol = 0
Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly change"
Cells(1, 11).Value = "percent change"
Cells(1, 12).Value = "total stock volume"
j = 2
'loop through spreadsheet
For i = 2 To lastrow
  If Cells(i, 1).Value <> tick Then
     Cells(j, 9).Value = tick
     Cells(j, 10).Value = closep - openp
     If openp <> 0 Then
        Cells(j, 11).Value = Round(((Cells(j, 10) / openp) * 100), 2)
     Else
         'if aboveformula used closep instead of openp - the percentage would be 100 - so if openp is 0 ...
         Cells(j, 11) = 100
     End If
     Cells(j, 12).Value = totvol
     If totvol > greatv Then
         'greatest volume chk
         greatv = totvol
         tickv = tick
     End If
     If Cells(j, 11).Value > 0 Then
        'increase stock price
        Cells(j, 10).Interior.ColorIndex = 4
        If Cells(j, 11).Value > greatpi Then
            greatpi = Cells(j, 11).Value
            tickpi = tick
        End If
     Else
        'decrease stock price
        Cells(j, 10).Interior.ColorIndex = 3
        If Cells(j, 11) < greatpd Then
            greatpd = Cells(j, 11).Value
            tickpd = tick
        End If
     End If
'final processing
     totvol = 0
     tick = Cells(i, 1).Value
     openp = Cells(i, 3).Value
     j = j + 1
  End If
  totvol = totvol + Cells(i, 7).Value
  closep = Cells(i, 6).Value
Next i

Cells(j, 9) = tick
Cells(j, 10).Value = closep - openp
Cells(j, 11).Value = Round(((Cells(j, 10).Value / openp) * 100), 2)
Cells(j, 12) = totvol
If totvol > greatv Then
    greatv = totvol
End If
If Cells(j, 11).Value > 0 Then
    Cells(j, 10).Interior.ColorIndex = 4
    If Cells(j, 11).Value > greatpi Then
        greatpi = Cells(j, 11).Value
    End If
    Else
        Cells(j, 10).Interior.ColorIndex = 3
        If Cells(j, 11) < greatpd Then
            greatpd = Cells(j, 11).Value
        End If
End If

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 16) = tickpi
Cells(2, 17) = greatpi
Cells(3, 16) = tickpd
Cells(3, 17) = greatpd
Cells(4, 16) = tickv
Cells(4, 17) = greatv

End Sub

