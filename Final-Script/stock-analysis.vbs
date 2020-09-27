Sub ticker_volume()

    'Dim Variables and Assign Values

    Dim ws As Worksheet
    
For Each ws In ThisWorkbook.Worksheets

    ' Assign Variables and Values
    Dim Yearly_Change As Double
    Dim LastRow As Double
    Dim Stock_Open As Double
    Dim Stock_Close As Double
    Dim Percent_Change As Double
    Dim Stock_Vol As Double
    Dim Summary As Double
    Dim Ticker As String
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double

    'Insert Headers into Cells for Columns
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "% Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(1, 14).Value = "Greatest Increase"
    ws.Cells(2, 14).Value = "Greatest Decrease"
    ws.Cells(3, 14).Value = "Greatest Volume"

    'Set Row Count, Stock Open for first iteration on each sheet, results row for ranges, and autofit columns

    Results = 2
    Master_Results = 1
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Stock_Open = ws.Cells(2, 3).Value
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Volume = 0
    ws.Columns("A:P").AutoFit

    'Loop Through Stocks for Ticker, Yearly_Change, Percent_Change, and Total_Volume
        
    For i = 2 To LastRow

        'Grab Stock Ticker and Insert into Ticker Column
    
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'Set new stock ticker
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Results).Value = Ticker

                'Calculate and set yearly change
                Stock_Close = ws.Cells(i, 6).Value
                Yearly_Change = Stock_Close - Stock_Open
                ws.Range("J" & Results).Value = Yearly_Change

                'Check for division by 0, if not then calculate percent change

                If Stock_Open <> 0 Then
                    Percent_Change = (Yearly_Change / Stock_Open)
                End If

                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
                ws.Range("L" & Results).Value = Stock_Vol
                
                'Print Percent Change to its Column and set conditional formatting
                ws.Range("K" & Results).Value = Percent_Change
                ws.Range("K" & Results).NumberFormat = "0.00%"
                If Percent_Change > 0 Then
                    ws.Range("K" & Results).Interior.ColorIndex = 4
                End If
                If Percent_Change < 0 Then
                    ws.Range("K" & Results).Interior.ColorIndex = 3
                End If

               'Reset Stock Values since we have triggered a new Ticker Value
                Results = Results + 1
                Stock_Vol = 0
                Yearly_Change = 0
                Percent_Change = 0
                Stock_Open = ws.Cells(i + 1, 3)

            Else
            
                'Calculate Stock Volume
                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value

            End If

        Next i
        
        'Loop through new summary results for Master Results

        For j = 2 To LastRow2

            'Find Greatest Increase, Format
            If ws.Cells(j, 11).Value > Greatest_Increase Then
                Greatest_Increase = ws.Cells(j, 11).Value
                ws.Range("O" & Master_Results).Value = ws.Cells(j, 9)
                ws.Range("P" & Master_Results).Value = Greatest_Increase
                ws.Range("P" & Master_Results).NumberFormat = "0.00%"
            End If

            'Find Greatest Decrease, Format
            If ws.Cells(j, 11).Value < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(j, 11).Value
                ws.Range("O" & Master_Results + 1).Value = ws.Cells(j, 9)
                ws.Range("P" & Master_Results + 1).NumberFormat = "0.00%"
                ws.Range("P" & Master_Results + 1).Value = Greatest_Decrease
            End If

            'Find Greatest Volume
            If ws.Cells(j, 12).Value > Greatest_Volume Then
                Greatest_Volume = ws.Cells(j, 12).Value
                ws.Range("O" & Master_Results + 2).Value = ws.Cells(j, 9)
                ws.Range("P" & Master_Results + 2).Value = Greatest_Volume
            End If
            
        Next j
            
    Next ws

End Sub