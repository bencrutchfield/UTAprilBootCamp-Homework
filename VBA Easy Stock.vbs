Sub EasyStockExchange()

Dim total_volume As Double


For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total STOCK Volume"
        Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        j = 0
    For I = 2 To Lastrow
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            total_volume = total_volume + ws.Cells(I, 7).Value
            ws.Range("I" & 2 + j).Value = ws.Cells(I, 1).Value
            ws.Range("J" & 2 + j).Value = total_volume
             j = j + 1
            total_volume = 0
Else
            total_volume = total_volume + ws.Cells(I, 7).Value
End If
Next I
Next ws
End Sub
