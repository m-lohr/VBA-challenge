'Attribute VB_Name = "Module1"
Sub stocks()


'Loop through all sheets
Dim ws As Worksheet
For Each ws In Worksheets

    'Define variables
    
    Dim ticker As String
    Dim tickerCount As Long
    
    Dim openPrice As Double
    Dim closePrice As Double
    Dim openRow As Long
    
    Dim priceChange As Double
    Dim percentChange As Double
    
    Dim totalStock As Double
    Dim openTotal As Range
    Dim closeTotal As Range
    
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxTicker As String
        
    'Define column headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Assign initial values to variables
    ticker = " "
    tickerCount = 1
    openPrice = 0
    closePrice = 0
    priceChange = 0
    percentChange = 0
    totalStock = 0
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    maxTicker = " "
    
    'Define last row
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop to find first and last row of each ticker symbol
    For i = 2 To lastRow
        
        'If statement to place ticker symbols in new table
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            tickerCount = tickerCount + 1
            ticker = ws.Cells(i, 1).Value
            ws.Cells(tickerCount, 9).Value = ticker
                        
            'Define locations of opening and closing prices
            openRow = ws.Range("A:A").Find(what:=ticker, after:=ws.Range("A1"), LookAt:=xlWhole).Row
            openPrice = ws.Cells(openRow, 3).Value
            closePrice = ws.Cells(i, 6).Value
            
            'Find difference in prices
            priceChange = closePrice - openPrice
            
            'If statement for conditional formatting of Yearly Change column
            If priceChange < 0 Then
            
                ws.Cells(tickerCount, 10).Value = priceChange
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 3
                
            ElseIf priceChange > 0 Then
                
                ws.Cells(tickerCount, 10).Value = priceChange
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 4
            
            ElseIf priceChange = 0 Then
                
                ws.Cells(tickerCount, 10).Value = priceChange
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 6
            
            End If
                        
            'Find and print percent change
            percentChange = (closePrice - openPrice) / openPrice
            ws.Cells(tickerCount, 11).Value = FormatPercent(percentChange)
            
            'Find Total Stock Volume rows and columns and sum
            Set openTotal = ws.Cells(openRow, 7)
            Set closeTotal = ws.Cells(i, 7)
            
            totalStock = WorksheetFunction.Sum(Range(openTotal, closeTotal))
            
            ws.Cells(tickerCount, 12).Value = totalStock
            
            'check if openRow is correct
            'ws.Cells(tickerCount, 13).Value = ws.Cells(openRow, 1)
            
        End If

    
    Next i
    
    For j = 2 To lastRow
        
        If maxIncrease < ws.Cells(j, 11).Value Then
            
            maxIncrease = ws.Cells(j, 11).Value
            ws.Cells(2, 17).Value = FormatPercent(maxIncrease)
            maxTicker = ws.Cells(j, 9).Value
            ws.Cells(2, 16).Value = maxTicker
        
        ElseIf maxIncrease > ws.Cells(j, 11).Value Then
            
            ws.Cells(2, 17).Value = FormatPercent(maxIncrease)
            
        End If
        
        If maxDecrease > ws.Cells(j, 11).Value Then
            
            maxDecrease = ws.Cells(j, 11).Value
            ws.Cells(3, 17).Value = FormatPercent(maxDecrease)
            maxTicker = ws.Cells(j, 9).Value
            ws.Cells(3, 16).Value = maxTicker
        
        ElseIf maxDecrease < ws.Cells(j, 11).Value Then
            
            ws.Cells(3, 17).Value = FormatPercent(maxDecrease)
        
        End If
        
        If maxVolume < ws.Cells(j, 12).Value Then
            
            maxVolume = ws.Cells(j, 12).Value
            ws.Cells(4, 17).Value = maxVolume
            maxTicker = ws.Cells(j, 9).Value
            ws.Cells(4, 16).Value = maxTicker
        
        ElseIf maxVolume > ws.Cells(j, 12).Value Then
            
            ws.Cells(4, 17).Value = maxVolume
        
        End If
    
    Next j
        
    'autofit column widths because i think it looks better
    ws.Columns("A:Z").AutoFit
    
Next ws

End Sub


