'I simply couldn;t get the code to slect the initial starting price for a stock and after searching and working on it, I'm giving up.  The easy version is done.

Sub stock_totals_moderat()
    
    ' create variable tick totals and total rows
    Dim tick As String
    Dim tick_total As Double
    Dim tick_total_row As Integer
    Dim change As Long
    Dim percent_change As Double
    Dim open_price As Long
    Dim close_price As Long
    tick_total_row = 2

    'create variable for length of sheets
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'name the headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Ticker_Total"
    
    'loop through sheet and total until change
    For x = 2 To LastRow
           
        
        If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then

            tick = Cells(x, 1).Value
            
            tick_total = tick_total + Cells(x, 7).Value

            Range("I" & tick_total_row).Value = tick

            Range("J" & tick_total_row).Value = change
            
            Range("L" & tick_total_row).Value = tick_total
            
            close_price = Cells(x, 6).Value
            
            tick_total_row = tick_total_row + 1

            tick_total = 0

        Else
                                        
            If open_price = 0 Then
            
             Cells(tick_total_row, 11).Value = Format(0, "Percent")
             
            Else
            
                Cells(tick_total_row, 11).Value = Format(((close_price - open_price) / open_price), "Percent")
            End If
   
            change = close_price - open_price
            
            tick_total = tick_total + Cells(x, 7).Value
            
            open_price = Cells(x, 3).Value


        End If

    Next x

End Sub