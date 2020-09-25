Sub ticker_volume()

    'Dim Variables and Assign Values

    Dim Yearly_Change As Double
    Dim LastRow As Double
    Dim Open_Price As Double
    Dim Stock_Open As Double
    Dim Stock_Close As Double
    Dim Percent_Change As Double
    Dim ws As Worksheet

    'Insert Headers into Cells for Columns
    
For Each ws In ThisWorkbook.Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "% Change"
    ws.Cells(1, 12).Value = "Total Volume"

    'Set Row Count to work with
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop Through Stocks for Ticker, Yearly_Change, Percent_Change, and Total_Volume
        
    For i = 2 To LastRow
    
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
        LastRow3 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Grab Stock Ticker and Insert into Ticker Column
    
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(LastRow2, 9) = ws.Cells(i, 1).Value
                    
                'Reset Stock Values since we have triggered a new Ticker Value
            
                Stock_Vol = 0
                Yearly_Change = 0
                Percent_Change = 0
                Stock_Open = ws.Cells(i, 3).Value
            
            Else
            
                'Calculate Stock Volume, Yearly Change, Percent Change, then insert
            
                Stock_Close = ws.Cells(i, 6).Value
            
                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
                    ws.Cells(LastRow3, 12).Value = Stock_Vol
            
                Yearly_Change = Stock_Close - Stock_Open
                    ws.Cells(LastRow3, 10).Value = Yearly_Change
                    
                    'Format Cells Based on Yearly Change Performance
                    
                    If ws.Cells(LastRow3, 10).Value > 0 Then
                        ws.Cells(LastRow3, 10).Interior.ColorIndex = 4
                    End If
                    If ws.Cells(LastRow3, 10).Value < 0 Then
                        ws.Cells(LastRow3, 10).Interior.ColorIndex = 3
                    End If
                
                    If Stock_Open <> 0 Then
                        Percent_Change = (Yearly_Change / Stock_Open)
                        ws.Cells(LastRow3, 11).Value = Percent_Change
                        ws.Cells(LastRow3, 11).NumberFormat = "0.00%"
                    End If
                
            End If
        
        Next i
        
    Next ws

Call summary

End Sub

Sub summary()

    Dim ws As Worksheet
    Dim Low_Stock As Double
    Dim High_Stock As Double
    Dim Greatest_Volume As Double
    
    ' Add sheet and declare values for summary table
    
    Sheets.Add.Name = "Summary"
    
    Cells(1, 1).Value = "Summary Table"
    Cells(1, 2).Value = "Ticker"
    Cells(1, 3).Value = "Metric"
    Cells(2, 1).Value = "Greatest % Increase"
    Cells(3, 1).Value = "Greatest % Decrease"
    Cells(4, 1).Value = "Greatest Stock Volume"
    Cells(2, 3).NumberFormat = "0.00%"
    Cells(3, 3).NumberFormat = "0.00%"
    Cells(2, 3).Value = 0
    Cells(3, 3).Value = 0
    Cells(4, 3).Value = 0
    
    ' Loop Through Worksheets

    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Name <> "Summary" Then
        
        LastRow3 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        High_Stock = Sheets("Summary").Cells(2, 3).Value
        Low_Stock = Sheets("Summary").Cells(3, 3).Value
        Greatest_Volume = Sheets("Summary").Cells(4, 3).Value
        
        'Loop through Stocks Summaries Columns for Greatest Increase, Greatest Decrease, Greatest Volume
    
            For j = 2 To LastRow3
    
                'Find Greatest Percent Increase
                
                If ws.Cells(j, 11) > High_Stock Then
                    Sheets("Summary").Cells(2, 2).Value = ws.Cells(j, 9).Value
                    Sheets("Summary").Cells(2, 3).Value = ws.Cells(j, 11).Value
                    High_Stock = Sheets("Summary").Cells(2, 3).Value
                End If
                
                'Find Greatest Percent Decrease
                
                If ws.Cells(j, 11) < Low_Stock Then
                    Sheets("Summary").Cells(3, 2).Value = ws.Cells(j, 9).Value
                    Sheets("Summary").Cells(3, 3).Value = ws.Cells(j, 11).Value
                    Low_Stock = Sheets("Summary").Cells(3, 3).Value
                End If
                
                'Find Greatest Volume
                
                If ws.Cells(j, 12) > Greatest_Volume Then
                    Sheets("Summary").Cells(4, 2).Value = ws.Cells(j, 9).Value
                    Sheets("Summary").Cells(4, 3).Value = ws.Cells(j, 12).Value
                    Greatest_Volume = Sheets("Summary").Cells(4, 3).Value
                End If
            
            Next j
        
        End If
    
    Next ws

End Sub