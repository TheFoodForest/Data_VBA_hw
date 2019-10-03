Sub stocks_summary_table()
    Dim ticker As String
    Dim year As Long
    Dim open_price As Double
    Dim high_price As Double
    Dim low_price As Double
    Dim close_price As Double
    Dim stock_volume As Long
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_vol As Double
    Dim last_row As Long
    Dim sum_table_row As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet
    
    
    
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    sum_table_row = 2
    total_vol = 0

    
    'Getting last row of data for looping
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'creating a starter count for looping while appending to a
    'unique_index array
    Dim starter_count As Integer
    Dim unique_index(0 To 10000) As Long
    'setting starter count = 1 because loop below that fills doesnt get the
    'starting row for the first unique ticker row index
    starter_count = 1
    
    'setting the first unique index as 1, need to do for later looping reasons
    'the real start of the first uniqe ticker is on row index 2
    unique_index(0) = 1
    
     
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop gets the unique tickers from the first column, loops through every row
    'fills first column of summary table with the unique tickers

        For i = 2 To last_row
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                ticker = Cells(i, 1).Value
                'filling summary table with unique ticker
                Range("I" & sum_table_row).Value = ticker
                'making it so next unique ticker will be in next row by increasing sum_table_row by 1
                sum_table_row = sum_table_row + 1
                'when it finds a unique row by having the ticker change, it sets the unique_index(starter_count) = that row index
                'this gives us the last row index in the main data table (has 700k rows or so) of the unique ticker the loop has found
                unique_index(starter_count) = (i)
                'increasing starter_count by 1 so for the next unique ticker it finds, loop will be setting the index of unique_index to the next number
                '^ means it will append to the end of the list basically
                starter_count = starter_count + 1
                
            End If
        Next i
        
    'more variables
    Dim last_result_row As Integer
    'Dim unique_ticker As String
    'Dim date_check As Long
    'Dim tick_check As String
    
    'gets the last row index of the above filled first column of the summary table with all the unique tickers
    last_result_row = Cells(Rows.Count, 9).End(xlUp).Row
    
    
    'looping through the lenght of last row
    'doing this because the array, unique_index, has the same number of values as the tickers summary row
    'basically every unique ticker has a row index for the last row with that unique ticker in the main data table
    For i = 2 To last_result_row
        'below loop goes from starting and ending row indexes in the main data table for each unique ticker
        'continually sums the stock volume for the rows that have their <ticker> column value = a given unique ticker
        For k = (unique_index(i - 2) + 1) To unique_index(i - 1)
            
            total_vol = total_vol + Range("G" & k).Value
            
        Next k
        
        'back into above for loop with "i" counter
        'checking if open_price cell is 0, to avoid division by 0 errors
        If Range("c" & (unique_index(i - 2) + 1)).Value <> 0 Then
            open_price = Range("c" & (unique_index(i - 2) + 1)).Value
            close_price = Range("f" & unique_index(i - 1)).Value
            yearly_change = close_price - open_price
            percent_change = yearly_change / open_price
            
        Else
            close_price = Range("f" & unique_index(i - 1)).Value
            open_price = 0
            percent_change = 0
            yearly_change = close_price - open_price
           
        End If
        
        'filling results table with values
        'giving color and number format to cells
        Range("l" & i).Value = total_vol
        total_vol = 0
        Range("j" & i).Value = yearly_change
        If yearly_change < 0 Then
            Range("j" & i).Interior.ColorIndex = 3
        Else
            Range("j" & i).Interior.ColorIndex = 4
        End If
        
        Range("k" & i).Value = percent_change
        Range("K" & i).NumberFormat = "0.00%"
        yearly_change = 0
        percent_change = 0
        
    Next i
    'next i all the way down here because I do the loop for each unique ticker, thats why total_vol summation loop is for the row index's of each unique ticker
    
    
    Range("i1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"

    
    Range("n2").Value = "Greatest % Increase"
    Range("n3").Value = "Greatest % Decrease"
    Range("n4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    
    Dim max_percent As Double
    Dim max_percent_ticker As String
    Dim min_percent As Double
    Dim min_percent_ticker As String
    Dim max_vol As Double
    Dim max_vol_ticker As String
    
    
    max_percent = Application.WorksheetFunction.Max(Range("K2:K" & last_result_row))
    
    min_percent = Application.WorksheetFunction.Min(Range("k2:k" & last_result_row))
    
    max_vol = Application.WorksheetFunction.Max(Range("L2:L" & last_result_row))
    
    
    For h = 2 To last_result_row
    
        If Range("K" & h).Value = max_percent Then
        
            max_percent_ticker = Range("I" & h).Value
            
        End If
        
        If Range("k" & h).Value = min_percent Then
            min_percent_ticker = Range("I" & h).Value
        End If
        
        If Range("L" & h).Value = max_vol Then
            max_vol_ticker = Range("i" & h).Value
        End If
        
    Next h
    Range("O2").Value = max_percent_ticker
    Range("O3").Value = min_percent_ticker
    Range("O4").Value = max_vol_ticker
    Range("P2").Value = max_percent
    Range("P2").NumberFormat = "0.00%"
    Range("p3").Value = min_percent
    Range("P3").NumberFormat = "0.00%"
    Range("p4").Value = max_vol
    
    Next
    starting_ws.Activate
    
End Sub