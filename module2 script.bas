Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()

    ' Loop through all worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        
        ' Set initial variables
        Dim ticker As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim totalVolume As Double
        Dim quarterlyChange As Double
        Dim percentChange As Double
        
        Dim lastRow As Long
        Dim summaryRow As Integer
        summaryRow = 2
        
        ' Add headers to the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        
        ' Find the last row of data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables
        openPrice = ws.Cells(2, 3).Value
        totalVolume = 0
        
        ' Loop through the rows of stock data
        For i = 2 To lastRow
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' If the ticker changes, record data
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set ticker name
                ticker = ws.Cells(i, 1).Value
                ' Closing price for the quarter
                closePrice = ws.Cells(i, 6).Value
                ' Calculate quarterly change
                quarterlyChange = closePrice - openPrice
                ' Avoid division by zero
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Add data to summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = totalVolume
                ws.Cells(summaryRow, 11).Value = quarterlyChange
                ws.Cells(summaryRow, 12).Value = percentChange
                
                ' Apply conditional formatting for quarterly change
                If quarterlyChange > 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(summaryRow, 11).Interior.Color = vbRed
                End If
                
                ' Apply conditional formatting for percent change
                If percentChange > 0 Then
                    ws.Cells(summaryRow, 12).Interior.Color = vbGreen
                Else
                    ws.Cells(summaryRow, 12).Interior.Color = vbRed
                End If
                
                ' Reset variables for the next ticker
                summaryRow = summaryRow + 1
                totalVolume = 0
                openPrice = ws.Cells(i + 1, 3).Value
            End If
        Next i
        
        ' Find greatest % increase, decrease, and total volume
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim tickerIncrease As String
        Dim tickerDecrease As String
        Dim tickerVolume As String
        
        greatestIncrease = WorksheetFunction.Max(ws.Range("L2:L" & summaryRow))
        greatestDecrease = WorksheetFunction.Min(ws.Range("L2:L" & summaryRow))
        greatestVolume = WorksheetFunction.Max(ws.Range("J2:J" & summaryRow))
        
        tickerIncrease = ws.Cells(WorksheetFunction.Match(greatestIncrease, ws.Range("L2:L" & summaryRow), 0) + 1, 9).Value
        tickerDecrease = ws.Cells(WorksheetFunction.Match(greatestDecrease, ws.Range("L2:L" & summaryRow), 0) + 1, 9).Value
        tickerVolume = ws.Cells(WorksheetFunction.Match(greatestVolume, ws.Range("J2:J" & summaryRow), 0) + 1, 9).Value
        
        ' Output results
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(2, 16).Value = tickerIncrease
        ws.Cells(2, 17).Value = greatestIncrease
        
        ws.Cells(3, 16).Value = tickerDecrease
        ws.Cells(3, 17).Value = greatestDecrease
        
        ws.Cells(4, 16).Value = tickerVolume
        ws.Cells(4, 17).Value = greatestVolume
        
    Next ws

End Sub

