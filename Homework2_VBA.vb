Sub stocks()

Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Total Stock Volume"

Dim i As Double
Dim total_counter As Double
Dim Position As Integer
Position = 2
total_counter = 0

    For i = 2 To 797711
       If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            total_counter = total_counter + Cells(i, 7).Value
       Else
            total_counter = total_counter + Cells(i, 7).Value
            Cells(Position, 10).Value = Cells(i, 1).Value
            Cells(Position, 11).Value = total_counter
            total_counter = 0
            Position = Position + 1
        End If
    Next i

End Sub
