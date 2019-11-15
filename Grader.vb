' --- Grade calculator ---
Sub Calculate()

    Dim grade As Integer
    Dim result As Range
    Dim letter As Range

    grade = Range("B2").Value
    Set result = Range("C2").Value
    Set letter = Range("D2").Value
    
    
    If grade >= 90 Then
        result.Value = "Pass"
        letter.Value = "A"
        Range("C2").Interior.ColorIndex = 4

    ElseIf grade >= 80 Then
        result.Value = "Pass"
        letter.Value = "B"
        Range("C2").Interior.ColorIndex = 4

    ElseIf grade >= 70 Then
        result.Value = "Warning"
        letter.Value = "C"
        Range("C2").Interior.ColorIndex = 6

    ElseIf grade < 70 Then
        result.Value = "Fail"
        letter.Value = "F"
        Range("C2").Interior.ColorIndex = 3

    End If

End Sub

'--- Clear Values ---
Sub ResetValues()

Range("B2").Value = ""
Range("C2").Value = ""
Range("D2").Value = ""

End Sub