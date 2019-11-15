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