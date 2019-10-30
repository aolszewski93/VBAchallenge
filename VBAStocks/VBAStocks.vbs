Sub yearOfStockData()
    'create headers for collected data
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    'loop through data and collect ticker type, yearly change, percent change, and total stock volume
    'add conditional formatting
    
    Dim sheetLength As Long
    Dim tickerType As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim tickerCount As Integer
    Dim startOfStock As Double
    
    
    
    sheetLength = Cells(Rows.Count, "A").End(xlUp).Row
    'MsgBox (sheetLength)
    tickerCount = 2
    totalVolume = 0
    percentChange = 0
    yearlyChange = 0
    startOfStock = 2
    
    For i = 2 To sheetLength
        'grab data from row
        tickerType = Cells(i, 1).Value
        totalVolume = totalVolume + Cells(i, 7).Value
        
        
        'check to see if we are in the same ticker type
        If tickerType <> Cells(i + 1, 1).Value Then
            
            'handle stocks with zeros in all values
            If totalVolume = 0 Then
                yearlyChange = 0
                percentChange = 0
            Else
                'find first nonzero start value
                If Cells(startOfStock, 3).Value = 0 Then
                    For j = startOfStock To i
                        If Cells(startOfStock, 3).Value <> 0 Then
                            startOfStock = j
                            Exit For
                        End If
                    Next j
                End If
                'calculate yearlyChange and percentChange
                yearlyChange = Cells(i, 6).Value - Cells(startOfStock, 3).Value
                percentChange = yearlyChange / Cells(startOfStock, 3).Value * 100
            End If
            
            'log the data in sheet
            Range("I" & tickerCount).Value = tickerType
            Range("L" & tickerCount).Value = totalVolume
            Range("J" & tickerCount).Value = yearlyChange
            Range("K" & tickerCount).Value = "%" & percentChange
            
            ' colors positives green and negatives red
                Select Case yearlyChange
                    Case Is > 0
                        Range("J" & tickerCount).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & tickerCount).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & tickerCount).Interior.ColorIndex = 0
                End Select
                
            'add one to the ticker count
            tickerCount = tickerCount + 1
            'create new start value for next stock
            startOfStock = i + 1
            'reset Volume
            totalVolume = 0
        End If
    Next i
    
    'print out greatest percent increase, percent decrease, and total volume along with ticker
    Range("O2").Value = "Greatest Percent Increase"
    Range("O3").Value = "Greatest Percent Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "ticker"
    Range("Q1").Value = "value"
    'run a worksheet function to find values
    Range("Q2").Value = "%" & WorksheetFunction.Max(Range("K2:K" & tickerCount)) * 100
    Range("Q3").Value = "%" & WorksheetFunction.Min(Range("K2:K" & tickerCount)) * 100
    Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & tickerCount))
    'find tickers
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & tickerCount)), Range("K2:K" & tickerCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & tickerCount)), Range("K2:K" & tickerCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & tickerCount)), Range("L2:L" & tickerCount), 0)
    ' final ticker symbol for  total, greatest % of increase and decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)
    
End Sub
