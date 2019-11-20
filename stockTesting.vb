
'Stocks V2------------------- WORKING !
Sub Stocks()

    'Init Variables
    Dim currentTicker As String
    Dim newTicker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As LongLong
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Long
    Dim outputRow As Integer
    Dim LastRow As Integer
    
    outputRow = 1

    'Make Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    
    openPrice = Cells(2, 3).Value
    closePrice = Cells(2, 6).Value
    'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Dim LastRow As Long
    'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim i As Longs
    newTicker = Cells(i + 1, 1).Value
    volume = Cells(i, 7).Value


         If currentTicker <> newTicker Then

            If openPrice = 0 Then
                openPrice = null 
            End If

            closePrice = Cells(i,6).Value

            yearlyChange = closePrice - openPrice

            percentChange = yearlyChange / openPrice

            openPrice = Cells(i+1,3).Value

            
            'Add Volume 
            totalVolume = volume + totalVolume
            
            outputRow = outputRow + 1
            
            Range("I" & outputRow).Value = currentTicker
            Range("J" & outputRow).Value = yearlyChange
            Range("K" & outputRow).Value = percentChange
            Range("L" & outputRow).Value = totalVolume

            if yearlyChange > 0 Then
                Range("J" & outputRow).Interior.ColorIndex = 4
            elseif yearlyChange < 0 Then
                Range("J" & outputRow).Interior.ColorIndex = 3
            End if

            'Reset Volume
            totalVolume = 0
        Else
            'Just keep adding
            totalVolume = volume + totalVolume
        End If

       
    Next i

End Sub