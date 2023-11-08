Sub multiple_year_stock_data()
    Dim ws As Worksheet

    Dim ticker As String
    Dim rowend As Long
    
    Dim stockRow As Integer
    
    Dim i As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    
    For Each ws In ThisWorkbook.Sheets
    
        rowend = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Initialize
        stockRow = 2
        ticker = ws.Cells(2, 1).Value
        openPrice = ws.Cells(2, 3).Value
        totalVolume = 0
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To rowend
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                ws.Cells(stockRow, 9).Value = ws.Cells(i, 1).Value
                
                closePrice = ws.Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                ws.Cells(stockRow, 10).Value = yearlyChange

                if(yearlyChange >= 0) Then
                    ws.Cells(stockRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(stockRow, 10).Interior.ColorIndex = 3
                End If
                
                If (openPrice <> 0) Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If
                ws.Cells(stockRow, 11).Value = Format(percentChange, "0.00%")
                ws.Cells(stockRow, 12).Value = totalVolume + ws.Cells(i, 7).Value
                
                
                'reset values
                
                stockRow = stockRow + 1
                openPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
                
            Else
                closePrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        Dim maxPercentIncrease As Double
        Dim maxPercentDecrease As Double
        Dim maxTotalVolume As Double
        
        Dim stockRowEnd As Long
        Dim j As Long
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        stockRowEnd = ws.Range("I1").End(xlDown).Row
        
        maxPercentIncrease = WorksheetFunction.Max(ws.Range("K:K"))
        maxPercentDecrease = WorksheetFunction.Min(ws.Range("K:K"))
        maxTotalVolume = WorksheetFunction.Max(ws.Range("L:L"))
        
        ws.Range("Q2").Value = Format(maxPercentIncrease, "0.00%")
        ws.Range("Q3").Value = Format(maxPercentDecrease, "0.00%")
        ws.Range("Q4").Value = maxTotalVolume
        
        For j = 2 To stockRowEnd
            ticker = ws.Cells(j, 9).Value
            If (ws.Cells(j, 11).Value = maxPercentIncrease) Then
                ws.Range("P2").Value = ticker
                
            ElseIf (ws.Cells(j, 11).Value = maxPercentDecrease) Then
                ws.Range("P3").Value = ticker
              
            ElseIf (ws.Cells(j, 12).Value = maxTotalVolume) Then
                ws.Range("P4").Value = ticker
                
            End If
        Next j
        
    Next ws
    
End Sub
