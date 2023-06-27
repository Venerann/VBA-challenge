Sub Stock()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRow As Long
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    For Each ws In ThisWorkbook.Worksheets
    
        summaryTableRow = 2
        
        greatestIncrease = 0
        
        greatestDecrease = 0
        
        greatestVolume = 0
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        For i = 2 To lastRow
        
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
                If totalVolume > 0 Then
                
                    ws.Cells(summaryTableRow, 9).Value = tickerSymbol
                    ws.Cells(summaryTableRow, 10).Value = yearlyChange
                    ws.Cells(summaryTableRow, 11).Value = percentChange
                    ws.Cells(summaryTableRow, 12).Value = totalVolume
                    
                    If yearlyChange > 0 Then
                    
                        ws.Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0)
                        
                    ElseIf yearlyChange < 0 Then
                    
                        ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0)
                        
                    End If
                    
                    summaryTableRow = summaryTableRow + 1
                    
                End If
                
                tickerSymbol = ws.Cells(i, 1).Value
                
                openingPrice = ws.Cells(i, 3).Value
                
                yearlyChange = 0
                
                percentChange = 0
                
                totalVolume = 0
                
            End If
            
            closingPrice = ws.Cells(i, 6).Value
            
            yearlyChange = yearlyChange + closingPrice - openingPrice
            
            percentChange = yearlyChange / openingPrice
            
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            

            If percentChange > greatestIncrease Then
            
                greatestIncrease = percentChange
                
                greatestIncreaseTicker = tickerSymbol
                
            End If
            
            If percentChange < greatestDecrease Then
            
                greatestDecrease = percentChange
                
                greatestDecreaseTicker = tickerSymbol
                
            End If
            
            If totalVolume > greatestVolume Then
            
                greatestVolume = totalVolume
                
                greatestVolumeTicker = tickerSymbol
                
            End If
            
        Next i
        
        If totalVolume > 0 Then
        
            ws.Cells(summaryTableRow, 9).Value = tickerSymbol
            ws.Cells(summaryTableRow, 10).Value = yearlyChange
            ws.Cells(summaryTableRow, 11).Value = percentChange
            ws.Cells(summaryTableRow, 12).Value = totalVolume
          
        End If
        
    Next ws
    
    For Each ws In ThisWorkbook.Worksheets
    
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 17).Value = greatestVolume
        
    Next ws
    
End Sub


