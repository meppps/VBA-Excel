'--- credit card checker
Sub Summarize()


    Dim cardBrand As String
    Dim nextBrand As String
    Dim cardRow As Integer
    Dim currentTotal As Double

    cardRow = 2
    cardBrand = ""
    BrandTotal = 0


   For i = 2 To 101
   
    cardBrand = Cells(i, 1).Value
    nextBrand = Cells(i + 1, 1).Value
    currentTotal = Cells(i, 3).Value
    
        If cardBrand <> nextBrand Then
        

            Range("G" & cardRow).Value = cardBrand
            Range("G" & cardRow).Interior.ColorIndex = 6
            Range("H" & cardRow).Value = BrandTotal
            Range("H" & cardRow).Interior.ColorIndex = 6
            Range("H" & cardRow).Style = "Currency"
            cardRow = cardRow + 1
        
            BrandTotal = 0

        Else
            BrandTotal = currentTotal + Cells(i, 3).Value
        End If
    Next i
End Sub