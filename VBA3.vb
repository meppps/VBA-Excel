'Nested Loop :)
Sub StarCounter()

  ' Create a variable to hold the StarCounter. We will repeatedly use this.
    Dim starCount As Integer
   ' Dim lastrow

    
    'lastrow = Cells(Rows.Count, 1) .End(xlup).Row
    'msgbox(lastrow)
  ' Loop through each row
  For i = 2 To 49

    ' Initially set the StarCounter to be 0 for each row
        starCount = 0

    ' While in each row, loop through each star column
        For j = 4 To 8

      ' If a column contains the word "Full-Star"...
            If Cells(i, j).Value = "Full-Star" Then

        ' Add 1 to the StarCounter
            starCount = starCount + 1

    ' Once we've iterated through each column in row i, print the value in the total column.
        
        End If
    Next j
    Cells(i, 9).Value = starCount
  Next i

End Sub



' ---- Formatter RGB -----
' Conditional formatting
' http://dmcritchie.mvps.org/excel/colors.htm

Range("A1:A3").font.colorindex = 3




' -- Grade calculator ---
Sub Calculate()

    Dim grade As Integer
    Dim result As String
    Dim letter As String

    grade = Range("B2").Value
    result = Range("C2").Value
    letter = Range("D2").Value
    
    
    If grade >= 90 Then
        result = "Pass"
        letter = "A"
        
    ElseIf grade >= 80 Then
        result = "Pass"
        letter = "B"
        
    ElseIf grade >= 70 Then
        result = "Warning"
        letter = "C"
        
    ElseIf grade < 70 Then
        result = "Fail"
        letter = "F"

    End If

End Sub

' --- Calculate V2 ---
' Use SET when assigning variables instead of doing the range manually
Sub Calculate()

    Dim grade As Integer
    Dim result As Range
    Dim letter As Range

    grade = Range("B2").Value
    Set result = Range("C2")
    Set letter = Range("D2")
    
    
    If grade >= 90 Then
        result.Value = "Pass"
        letter.Value = "A"
        
    ElseIf grade >= 80 Then
        result.Value = "Pass"
        letter = "B"
        
    ElseIf grade >= 70 Then
        result.Value = "Warning"
        letter = "C"
        
    ElseIf grade < 70 Then
        result.Value = "Fail"
        letter = "F"

    End If

Sub checkers()
Dim i As Integer
Dim j As Integer
i = 1
j = 1

For i = 1 To 8
    For j = 1 To 8
    
        If i Mod 2 = 1 And j Mod 2 = 1 Then
        Cells(i, j).Interior.ColorIndex = 3

        ElseIf i Mod 2 = 1 And j Mod 2 = 0 Then
        Cells(i, j).Interior.ColorIndex = 1

        ' If row is even
        ElseIf i Mod 2 = 0 And j Mod 2 = 1 Then
        Cells(i, j).Interior.ColorIndex = 1

        Else
        Cells(i, j).Interior.ColorIndex = 3

        

    End If
    Next j
Next i

' Alternate easier way

for i = 1 to 8
    for j = 1 to 8

if (cellnumber mod 2 = 0)

' Tally purchases for each brand



End Sub



'--- credit card checker
Sub Summarize()


    Dim cardBrand As String
    Dim nextBrand As String
    Dim cardRow As Integer
    Dim currentTotal As Double

    cardRow = 2
    cardBrand = ""
    BrandTotal = 0
'    currentTotal = Cells(2, 3).Value

   For i = 2 To 101
   
    cardBrand = Cells(i, 1).Value
    nextBrand = Cells(i + 1, 1).Value
    currentTotal = Cells(i, 3).Value
    
        If cardBrand <> nextBrand Then
        
'            currentTotal = currentTotal + Cells(i, 3).Value
            Range("G" & cardRow).Value = cardBrand
            Range("H" & cardRow).Value = BrandTotal
            cardRow = cardRow + 1
        
            BrandTotal = 0

        Else
            BrandTotal = currentTotal + Cells(i, 3).Value
        End If
    Next i
End Sub


' ------ Wells Fargo
'------- Loop thru wrksheets --------------------------
' --- !! BROKEN !! Needs fixed ---
Sub WellsFargo()

For Each ws In Worksheets

    ' Insert the state

    Dim Worksheetname As String

    'Get last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Worksheetname = ws.Name

    'Split the worksheetname
    State = Split(Worksheetname, "_")

    'Add the state to column
    ws.Range("A1").EntireColumn.Insert
    
    'Add the word state to the first column haeder
    ws.Cells(1, 1).Value = "State"

    'Add the state to all rows
    ws.Range("A2:A & LastRow") = State(0)

    ' Correct the year numbers

    ' !!!Determine the last COLUMN number
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    'RENAME THE YEAR COLUMNS BY LOOPING THRU AND RENAMING
    For i = 3 To LastColumn
        YearHeader = ws.Cells(1, i).Value
        YearSplit = Split(YearHeader, " ")

        ' Notice 0 based index
        ws.Cells(1, i).Value = YearSplit(3)

    Next i

    'Correct the currency format
    For i = 2 To LastRow
        For j = 2 To LastColumn

            ws.Cells(i, j).Style = "Currency"
        Next j
    Next i
Next ws

End Sub

'--- Follow up 
' --- Wells fargo 2
' INCOMPLETE, GET FROM SLACK

Sub WellsFargo2()

Sheets.Add.Name = "Combined_Data"

Sheets("Combined_Data").Move B


End Sub








