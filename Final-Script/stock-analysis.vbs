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

    'Insert Headers into Cells for Columns
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "% Change"
    ws.Cells(1, 12).Value = "Total Volume"

    'Set Row Count, Stock Open for first iteration on each sheet, results row for ranges, and autofit columns

    Results = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Stock_Open = ws.Cells(2, 3).Value
    ws.Columns("A:L").AutoFit

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
        
    Next ws

End Sub