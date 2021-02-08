Sub easyOption():

' declare variable
Dim total As Double

' get number of rows
RowCount = Cells(Rows.Count, "a").End(xlUp).Row

' create new column
Range("J1").Value = "TickerSymbol"
Range("k1").Value = "Total Stock Volume"

For i = 2 To RowCount

' Print result when Ticker Symbol changes
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Print symbol
Range("J" & 2 + k).Value = Cells(i, 1).Value

' print total
Range("k" & 2 + k).Value = total

total = 0

'do next row
k = k + 1

Else
    total = total + Cells(i, 7).Value
End If

Next i

End Sub

