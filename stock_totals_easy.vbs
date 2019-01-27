Sub stock_totals_easy()
    
    ' create variable tick totals and total rows
    Dim tick As String
    Dim tick_total As Double
    Dim tick_total_row As Integer
    tick_total_row = 2

    'create variable for length of sheets
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'name the headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Ticker_Total"
    
    'loop through sheet and total until change
    For x = 2 To LastRow

        If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then

            tick = Cells(x, 1).Value

            tick_total = tick_total + Cells(x, 7).Value

            Range("I" & tick_total_row).Value = tick

            Range("J" & tick_total_row).Value = tick_total

            tick_total_row = tick_total_row + 1

            tick_total = 0

        Else

            tick_total = tick_total + Cells(x, 7).Value

        End If

    Next x

End Sub

    

