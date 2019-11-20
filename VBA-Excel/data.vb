
' Cells(r,c)
'------------------------------------
Sub Stocks()

    'Init Variables
    Dim ticker As String
    Dim currentTicker As String
    Dim newTicker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Long
    Dim outputRow As Integer
    
    ticker = ""
    outputRow = 2
    'Make Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    
    openPrice = Cells(2, 3).Value
    closePrice = Cells(2, 6).Value

    Dim i As Integer
    
    For i = 2 To 70926

    currentTicker = Cells(i, 1).Value
    newTicker = Cells(i+1,1).Value
    totalVolume = 0 
    volume = cells(i,7).Value

         If currentTicker <> newTicker Then
            'Wrap it up

            'ticker = currentTicker.value
            'Add Vol
            totalVolume = volume + totalVolume
            outputrow = outputRow + 1
            'Output()
            Range("F",outputRow).Value = ticker
        Else
            'Just keep adding
            totalVolume = volume + totalVolume
        End if

    Next i

    sub output()
        Range("F",outputRow).Value = ticker
        'Range()
        'Range()
        'Range()
    End Sub

End Sub

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
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Dim LastRow As Long
    'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    
    For i = 2 To LastRow

    currentTicker = Cells(i, 1).Value
    newTicker = Cells(i + 1, 1).Value
    volume = Cells(i, 7).Value


         If currentTicker <> newTicker Then

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

