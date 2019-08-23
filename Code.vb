Sub StockMarket()

    Dim lastRow As Long
    Dim counter As Long
    Dim volume As Double
    counter = 2
    lastRow = Cells(rows.Count, 1).End(xlUp).Row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    For i = 2 To lastRow
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            Cells(counter, 9).Value = Cells(i, 1).Value
            Cells(counter, 10).Value = volume
            counter = counter + 1
            volume = 0
        Else
            volume = volume + Cells(i + 1, 7)
        End If
    Next i

End Sub
