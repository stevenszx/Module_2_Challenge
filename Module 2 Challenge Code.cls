Option Explicit

Sub RunStockData()
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim tickerSymbol As String
    Dim volumeStock As Double
    Dim openPrice As Double
    Dim closePrice As Double
    
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    Range("I:L").ClearContents
    Range("I1") = "Tickers"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    n = 1
    For i = 2 To lastRow
        If Cells(i, 1) <> Cells(i - 1, 1) Then
            n = n + 1
            tickerSymbol = Cells(i, 1)
            openPrice = Cells(i, 3)
            volumeStock = 0
            For j = i To i + 253
                If Cells(j, 1) = tickerSymbol Then
                    volumeStock = volumeStock + Cells(j, 7)
                    closePrice = Cells(j, 6)
                End If
            Next j
            Cells(n, 9) = tickerSymbol
            Cells(n, 10) = closePrice - openPrice
            Cells(n, 11) = (closePrice - openPrice) / openPrice
            Cells(n, 12) = volumeStock
        End If

    Next i
            
    
    lastRow = Range("L" & Rows.Count).End(xlUp).Row
    
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim tickerInc As String
    Dim tickerDec As String
    Dim tickerVol As String
    
    maxIncrease = Cells(2, 11)
    tickerInc = Cells(2, 9)
    maxDecrease = Cells(2, 11)
    tickerDec = Cells(2, 9)
    maxVolume = Cells(2, 12)
    tickerVol = Cells(2, 9)
    For i = 3 To lastRow
        If Cells(i, 11) > maxIncrease Then
            maxIncrease = Cells(i, 11)
            tickerInc = Cells(i, 9)
        End If
        If Cells(i, 11) < maxDecrease Then
            maxDecrease = Cells(i, 11)
            tickerDec = Cells(i, 9)
        End If
        If Cells(i, 12) > maxVolume Then
            maxVolume = Cells(i, 12)
            tickerVol = Cells(i, 9)
        End If
    Next
    
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    
    Range("P1") = "Tickers"
    Range("Q1") = "Value"
    
    Range("P2") = tickerInc
    Range("P3") = tickerDec
    Range("P4") = tickerVol
    Range("Q2") = maxIncrease
    Range("Q3") = maxDecrease
    Range("Q4") = maxVolume

End Sub
