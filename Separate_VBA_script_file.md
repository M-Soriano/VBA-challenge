
Sub tickerStock()

Dim ws As Worksheet

' Worksheet loop

    For Each ws In ThisWorkbook.Worksheets
    
    
       ' adding row and header variables
    'rows
    r = 1
    ' rows for counter
    
    Row = 2
    ' column
    c = 1
    
    
    ' setting closing price and opennig price as double
    Dim open_price As Double
    Dim closing_price As Double
    Dim gt_volume As Double
    
    'setting quarterly and percent variables as double
    Dim quarterly_change As Double
    Dim percent_change As Double
    
    ' Adding the header for need rows(hard coding)
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    
         'LastRow in Column A
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     'set initial price
     open_price = ws.Cells(2, 3).Value
     'set initial value of volume
     volume = 0
     
        'greatest increase, decrease, volume value set
        increase = ws.Cells(r + 1, 10).Value
        decrease = ws.Cells(r + 1, 10).Value
        gt_volume = ws.Cells(r + 1, 11).Value
        increase_ticker = ws.Cells(r + 1, 1).Value
        decrease_ticker = ws.Cells(r + 1, 1).Value
        volume_ticker = ws.Cells(r + 1, 1).Value
     
    ' Ticker name
     For r = 2 To LastRow
     
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
        ' Get the ticker symbol
        ticker = ws.Cells(r, 1).Value
        ' add the ticker symbol and row for this selection
        
        ws.Cells(Row, c + 8).Value = ticker
        
        'closing price
        
        closing_price = ws.Cells(r, c + 5).Value
        'Calculate quarterly change
           
        quarterly_change = closing_price - open_price
        'Calculation for precentage change
        percent_change = quarterly_change / open_price
        'which is the increase
            If percent_change > increase Then
                increase = percent_change
                increase_ticker = ws.Cells(r, 1).Value
                
                'drease option
                ElseIf percent_change < decrease Then
                    decrease = percent_change
                    decrease_ticker = ws.Cells(r, 1).Value
                    
                    End If
                      
                
        'vlolume caluclation and cell info insert
        
        volume = volume + ws.Cells(r, c + 6).Value
        ws.Cells(Row, c + 11).Value = volume
            If volume > gt_volume Then
            gt_volume = volume
            volume_ticker = ws.Cells(r, 1).Value
            End If
            
        'input result to cell
        ws.Cells(Row, c + 9).Value = quarterly_change
            If quarterly_change > 0 Then
            ws.Cells(Row, c + 9).Interior.ColorIndex = 4
                ElseIf quarterly_change < 0 Then
                ws.Cells(Row, c + 9).Interior.ColorIndex = 3
                End If
            
        ws.Cells(Row, c + 10).Value = percent_change '
        'Format percent change cells
        ws.Cells(Row, c + 10).NumberFormat = "0.00%"
        
        ' row counter
        Row = Row + 1
        
        'resetting openning price
        open_price = ws.Cells(r + 1, c + 2).Value
        
        'resetting volume
        volume = 0
        
        'set volume to catch at the volume of the sameticker
        ElseIf ws.Cells(r + 1, 1).Value = ws.Cells(r, 1).Value Then
            volume = volume + ws.Cells(r, c + 6).Value
     
        
            End If
        
            Next r
       
        ' setting new for look for Greastes increase, decrease, greastes total volume
        
        ws.Cells(2, 17).Value = increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 16).Value = increase_ticker
        ws.Cells(3, 17).Value = decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = decrease_ticker
        ws.Cells(4, 16).Value = volume_ticker
        ws.Cells(4, 17).Value = gt_volume
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
       
        'Format cell to auto fit
         ws.Cells.EntireColumn.AutoFit
       
            
    Next ws
    

End Sub

