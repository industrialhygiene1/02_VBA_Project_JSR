Sub stock_totals_moderat()
    
    ' create variable tick totals and total rows
    Dim tick As String
    Dim tick_total As Double
    Dim tick_total_row As Integer
    Dim change As Double
    Dim percent_change As Double
    Dim open_price As Double
    open_price = Cells(2, 3).Value
    Dim close_price As Double
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
    
    'Close_price
    close_price = Cells(x, 6).Value
    
    ' I have not been able to get the correct open_price from the begginnning of the loop and have looked everytone for a solution
    ' once submitting, I'm hopeing the get a solution to this problem
    open_price = Cells(x, 3).Value
        
        'loop through to gather other data and do calcs while creating the new summary table
        If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then

            tick = Cells(x, 1).Value
            
            tick_total = tick_total + Cells(x, 7).Value
                      
            Range("I" & tick_total_row).Value = tick

            Range("J" & tick_total_row).Value = change
            
            Range("k" & tick_total_row).Value = percent_change
            
            Range("L" & tick_total_row).Value = tick_total
            
            tick_total_row = tick_total_row + 1

            tick_total = 0
            
            change = close_price - open_price
            
            percent_change = change / close_price
            
                              
        Else
        
                                                         
        tick_total = tick_total + Cells(x, 7).Value
       
        End If

    Next x

End Sub

    


