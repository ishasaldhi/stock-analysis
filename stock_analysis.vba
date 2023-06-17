Sub StockMarket():
    
    ' Set a worksheet
    Dim w As Worksheet
    
    ' Loop through all stocks
    For Each w In Worksheets
    
    ' Make column headings
    w.Range("I1") = "Ticker"
    w.Range("J1") = "Yearly Change"
    w.Range("K1") = "Percentage Change"
    w.Range("L1") = "Total Stock Volume"
    
    w.Range("P1").Value = "Ticker"
    w.Range("Q1").Value = "Value"
    w.Range("O2").Value = "Greatest % Increase"
    w.Range("O3").Value = "Greatest % Decrease"
    w.Range("O4").Value = "Greatest Total Volume"
    
    
    
    ' Define a variable for ticker
    Dim ticker As String
    ticker = " "
    Dim tickerVolume As Double
    tickerVolume = 0
    
    'Set first and last rows
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer
    
    ' Define variable for last row
    lastRow = w.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set price and percentage variables
    Dim openPrice As Double
    Dim closePrice As Double
    Dim priceChange As Double
    Dim percentChange As Double
    openPrice = 0
    closePrice = 0
    priceChange = 0
    percentChange = 0
    
    Dim stockVolume As Double
    stockVolume = 0
    Dim tickerRow As Double
    tickerRow = 1
    
    'Loop through the worksheet until the last row
    For i = 2 To lastRow
    
        If w.Cells(i - 1, 1).Value <> w.Cells(i, 1).Value Then
            tickerRow = tickerRow + 1
            ticker = w.Cells(i, 1).Value
            w.Cells(tickerRow, "I").Value = ticker
            End If
            
            ' Calculate Price change
            If w.Cells(i - 1, 1).Value <> ticker Then
                openPrice = w.Cells(i, 3).Value
            End If
            
            If w.Cells(i + 1, 1).Value <> ticker Then
                closePrice = w.Cells(i, 6).Value
            End If
            
            priceChange = closePrice - openPrice
            w.Cells(tickerRow, "J").Value = priceChange
            percentChange = (priceChange / openPrice)
            w.Cells(tickerRow, "K").Value = Format(percentChange, "0.00%")
            
            
            ' Format Yearly Change
            If w.Cells(i, "J").Value < 0 Then
                w.Cells(i, "J").Interior.ColorIndex = 3
            Else
                w.Cells(i, "J").Interior.ColorIndex = 4
            End If
            
            ' Update StockVolume
            If w.Cells(i + 1, 1).Value = ticker Then
                stockVolume = stockVolume + w.Cells(i, 7).Value
            ElseIf w.Cells(i + 1, 1) <> ticker Then
                stockVolume = stockVolume + w.Cells(i, 7).Value
                w.Cells(tickerRow, "L").Value = stockVolume
                stockVolume = 0
            End If
            
        Next i
   
        w.Cells(2, "Q").Value = Application.WorksheetFunction.Max(Columns("K"))
        w.Cells(3, "Q").Value = Application.WorksheetFunction.Min(Columns("K"))
        w.Cells(4, "Q").Value = Application.WorksheetFunction.Max(Columns("L"))
        w.Cells(2, "Q").Value = Format(w.Cells(2, "Q").Value, "0.00%")
        w.Cells(3, "Q").Value = Format(w.Cells(3, "Q").Value, "0.00%")
        w.Cells(4, "Q").Value = Format(w.Cells(4, "Q").Value, "0.00%")
        
    Next w
    
End Sub

