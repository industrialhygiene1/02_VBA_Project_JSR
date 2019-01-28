Sub stock_totals_moderate()

    'standard variables created
       Dim open_price As Double
       Dim close_price As Double
       Dim ychange As Double
       Dim tick As String
       Dim percent_change As Double
    'total volume variable
       Dim Volume As Double
       Volume = 0
    'line variable to allow creation of summary table lines
       Dim Line As Integer
       Line = 2
    'clumn variable to allow selection of open/close prices and create summary table
       Dim Column As Integer
       Column = 1
    'row looping variable
       Dim x As Long

    'allow code to cross sheets
       For Each WS In Worksheets

    'count rows
       LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    'create new table headers
       Cells(1, "I").Value = "Ticker"
       Cells(1, "J").Value = "Yearly Change"
       Cells(1, "K").Value = "Percent Change"
       Cells(1, "L").Value = "Total Stock Volume"
    'selection of open price from third column
       open_price = Cells(2, Column + 2).Value
    'row for loop start
       For x = 2 To LastRow
        'loop until difference between values in first column
           If Cells(x + 1, Column).Value <> Cells(x, Column).Value Then
           'collect ticker symbol
               tick = Cells(x, Column).Value
            'insert ticker symbol
               Cells(Line, Column + 8) = tick
            'collect close price from 6th column
               close_price = Cells(x, Column + 5).Value
            'calculate yearly stock price change
               ychange = close_price - open_price
            'insert yearly price change in new summary table
               Cells(Line, Column + 9).Value = ychange
            'calculate percent change without getting excel math errors
               If (open_price = 0 And close_price = 0) Then

                   percent_change = 0

               ElseIf (open_price = 0 And close_price <> 0) Then

                   percent_change = 1

               Else

                   percent_change = ychange / open_price
            'insert percent change calculation and format at %
                   Cells(Line, Column + 10).Value = percent_change

                   Cells(Line, Column + 10).NumberFormat = "0.00%"
            'end if statment based on first column difference
               End If
            'calculate total volume outside if statement but within loop to get totals by ticker symbol
               Volume = Volume + Cells(x, Column + 6).Value
            'insert total valume calculation in the 12th column
               Cells(Line, Column + 11).Value = Volume
            'advance rows of the new summary table
               Line = Line + 1
            'collect opn price outside of if delta statement from the first row in the new different tisker symbol rows
               open_price = Cells(x + 1, Column + 2)
            'reset total volumn to 0 before re-entering the x loop
               Volume = 0

           Else
            'continue to add volume numbers
               Volume = Volume + Cells(x, Column + 6).Value

           End If
    'continue loop
       Next x


    'second row counter for new summary table formatting changes
       LastRow2 = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row

    'second loop for formatting
       For y = 2 To LastRow2
        'green for greater than 0
           If (Cells(y, Column + 9).Value > 0 Or Cells(y, Column + 9).Value = 0) Then

               Cells(y, Column + 9).Interior.ColorIndex = 4
         'red for less than 0      
           ElseIf Cells(y, Column + 9).Value < 0 Then

               Cells(y, Column + 9).Interior.ColorIndex = 3

           End If

       Next y


   Next WS

End Sub
