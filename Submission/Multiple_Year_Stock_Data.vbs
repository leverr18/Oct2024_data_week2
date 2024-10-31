Sub StockData()
' define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim leaderboardRow As Integer
    Dim volume As Double
    Dim totalVolume As Double
    Dim nextTicker As String
    
    ' new variables for quarterly and percent changes
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterChange As Double
    Dim percentChange As Double
    
    'new variables for last leaderboard
    Dim maxPrice As Double
    Dim minPrice As Double
    Dim greatestVolume As Double
    Dim maxPriceTick As String
    Dim minPriceTick As String
    Dim maxVolumeTick As String
    
    'loop through worksheets
  For Each ws In ThisWorkbook.Worksheets
        
' Reset per ticker
    
    totalVolume = 0
    openPrice = ws.Cells(2, 3).Value
    leaderboardRow = 2
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'initialize leaderboard tracking variables
    maxPrice = 0
    minPrice = 0
    greatestVolume = 0
' set headers for output
With Sheets("Q1")
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
End With

    For i = 2 To lastRow
        ' extract all values from this workbook
        
        ticker = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        nextTicker = ws.Cells(i + 1, 1).Value
        
        'if statement
        If (ticker <> nextTicker) Then
        ' add total
        totalVolume = totalVolume + volume
        
        'get close price
        closePrice = ws.Cells(i, 6).Value
        quarterChange = closePrice - openPrice
        ' calculate percent change
        If openPrice <> 0 Then
            percentChange = (quarterChange / openPrice)
        Else
            percentChange = 0
        End If
        
        ' write to leaderboard
            ws.Cells(leaderboardRow, 12).Value = totalVolume
            ws.Cells(leaderboardRow, 11).Value = FormatPercent(percentChange)
            ws.Cells(leaderboardRow, 10).Value = quarterChange
            ws.Cells(leaderboardRow, 9).Value = ticker
            
        'Conditional formatting
        If (quarterChange > 0) Then
            ws.Cells(leaderboardRow, 10).Interior.ColorIndex = 4
        ElseIf (quarterChange < 0) Then
            ws.Cells(leaderboardRow, 10).Interior.ColorIndex = 3
        Else
            'do nothing default white

        End If
        
        ' reset total
        totalVolume = 0
        leaderboardRow = leaderboardRow + 1
       ' update open price for the next ticker
       If i + 1 <= lastRow Then
            openPrice = ws.Cells(i + 1, 3).Value
        End If
        
    Else
        ' add total
        totalVolume = totalVolume + volume
    End If
    
    ' if statement for second leaderboard to loop through first leaderboard
    
    If percentChange > maxPrice Then
    maxPrice = percentChange
    maxPriceTick = ticker
    
    End If
    
    If percentChange < minPrice Then
    minPrice = percentChange
    minPriceTick = ticker
    
    End If
    
    If totalVolume > greatestVolume Then
    greatestVolume = totalVolume
    maxVolumeTick = ticker
    
    End If
    
    
Next i

    'output greatest values and their ticker
    ws.Cells(2, 17).Value = FormatPercent(maxPrice)
    ws.Cells(2, 16).Value = maxPriceTick
    
    ws.Cells(3, 17).Value = FormatPercent(minPrice)
    ws.Cells(3, 16).Value = minPriceTick
    
    ws.Cells(4, 17).Value = greatestVolume
    ws.Cells(4, 16).Value = maxVolumeTick
    
Next ws
End Sub
