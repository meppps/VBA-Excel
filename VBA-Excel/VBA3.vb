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

' 1. Extract words before the phrase "_Wells_Fargo" to figure out which State.
' 2. Add the State to the first column of each spreadsheet.
' 3. Convert the headers of each row to simply say the year.
' 4. Convert the cells to currency format

Sub WellsFargo_PtI()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws in Worksheets

        ' --------------------------------------------
        ' INSERT THE STATE
        ' --------------------------------------------

        ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        ' MsgBox WorksheetName

        ' Split the WorksheetName
        State = Split(WorksheetName, "_")
        ' MsgBox State(0)

        ' Add the State to the Column
        ws.Range("A1").EntireColumn.Insert

        ' Add the word State to the First Column Header
        ws.Cells(1, 1).Value = "State"

        ' Add the State to all rows
        ws.Range("A2:A" & LastRow) = State(0)

        ' --------------------------------------------
        ' CORRECT THE YEAR NUMBERS
        ' --------------------------------------------

        ' Determine the Last Column Number
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        ' Rename the Year columns by looping through and renaming each
        For i = 3 To LastColumn
            YearHeader = ws.Cells(1, i).Value
            YearSplit = Split(YearHeader, " ")
            ' MsgBox YearSplit(0)
            ws.Cells(1, i).Value = YearSplit(3)
            ' MsgBox Cells(1, i)
            ' MsgBox YearSplit(3)

        Next i

        ' --------------------------------------------
        ' CORRECT THE CURRENCY FORMAT
        ' --------------------------------------------

        ' Add the currency
        For i = 2 To LastRow

            For j = 2 To LastColumn

                ws.Cells(i, j).Style = "Currency"

            Next j

        Next i

    ' --------------------------------------------
    ' FIXES COMPLETE
    ' --------------------------------------------
    Next ws

    MsgBox ("Fixes Complete")


End Sub



' --- Wells fargo 2 ------------------------


' Part II:

' 1. Loop through every worksheet and select the state contents.
' 2. Copy the state contents and paste it into the Combined_Data tab

Sub WellsFargo_PtII()
    
    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Combined_Data"
    'move created sheet to be first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Combined_Data")

    ' Loop through all sheets
    For Each ws In Worksheets

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

        ' Copy the contents of each state sheet into the combined sheet
        combined_sheet.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value

    Next ws

    ' Copy the headers from sheet 1
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    ' Autofit to display data
    combined_sheet.Columns("A:G").AutoFit
End Sub



