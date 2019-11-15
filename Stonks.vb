Sub Stonks()
    Dim ticker As String
    Dim oldTicker As String
    Dim newTicker As String
   
    ticker = Range("A2:A70926").Value
   
    'Input Variables
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Long
   
    openPrice = Cells(2, 3).Value
    closePrice = Cells(2, 6).Value
    volume = Cells(2, 7).Value
   
    'Output Variables
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Long
   
   
    '70926
    For i = 2 To 70926
        For t = 2 To 264
            If Cells(i, 1).Value <> ticker Then
             Cells(t, 9).Value = Cells(i, 1).Value
                
            End If
        'volume = volume + volume
            
        Next t
    Next i
'totalVolume = volume


    'Output in sheet
    Cells(2, 12).Value = totalVolume
       
       
       
    
End Sub