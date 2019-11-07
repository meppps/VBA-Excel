' Sub PeanutButter():

' dim ingredients(0 to 2) as String

' ingredients(0) = "Peanut Butter"
' ingredients(1) = "Jelly"
' ingredients(2) = "Bread"

' for i = 0 to 5

'     if(pbThickness >= 1.0)then{
'         stopSpreading()
'     }
'     else:
'         spreadMore()
'     end if

'     next i
' Sub SpreadMore():

' End Sub

' Add numbers

Sub AddNumbers():
    Dim itemTotal As Double
    Dim itemPrice As Double
    Dim taxRate As Double
    Dim quantity As Double
    
    itemPrice = Range("A2").Value
    taxRate = Range("B2").Value
    quantity = Range("C2").Value
    
    
    itemTotal = itemPrice * (1 + taxRate) * quantity
    
    MsgBox ("Your total is $" + Str(itemTotal))
    Range("D2").Value = itemTotal
    
    
End Sub



' Sentence Breaker

Sub SentenceBreaker()
Dim Sentnece As String
sentence = Cells(1, 2).Value
MsgBox (sentence)

Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer

num1 = Cells(4, 1).Value
num2 = Cells(5, 1).Value
num3 = Cells(6, 1).Value

MsgBox (num1)
MsgBox (num2)
MsgBox (num3)

Dim SentenceArray() As String
SentenceArray = Split(sentence, " ")

Cells(4, 2).Value = SentenceArray(num1 - 1)
Cells(5, 2).Value = SentenceArray(num2 - 1)
Cells(6, 2).Value = SentenceArray(num3 - 1)

End Sub


' If statements of doom
Sub Doom():
Dim path As Integer
path = Range("B1").Value

If path = 1 Then
    MsgBox ("You have entered the forrest of doom")
ElseIf path = 2 Then
    MsgBox ("You have entered the volcano of doom")
ElseIf path = 3 Then
    MsgBox ("You have enterd the bathroom")
ElseIf path > 3 Then
    MsgBox ("You need directions")
End If
End Sub

Sub budgetChecker():
    Dim Budget As Double
    Dim Price As Double
    Dim Fees As Double
    Dim Total As Double
    
    Budget = Range("C3").Value
    Price = Range("F3").Value
    Fees = Range("H3").Value
    Total = Price + Price * Fees
    
    Range("L3").Value = Total
    
    If Price < Budget Then
        MsgBox ("Under Budget")
    ElseIf Price > Budget Then
        MsgBox ("Over Budget")
    End If
    
End Sub

' Chicken Nuggs
Sub forLoop():
    
    For i = 1 To 10
        Cells(i, 1).Value = "I will eat"
        Cells(i, 2).Value = i + 10
        Cells(i, 3).Value = "Chicken Nuggets"
    Next i
           
End Sub

' Conditional Loop
Sub conditional_loop():

    For i = 1 To 10
    
        If Cells(i, 1).Value Mod 2 = 0 Then
        Cells(i, 2).Value = "Even Row"
        
        Else
        
        Cells(i, 2).Value = "Odd Row"
        
        End If
        Next i
End Sub

'Fizz buzz for loop
Sub fizzBuzz():
    For i = 2 To 101
    'next time init variable
    'num = Cells(i,1).value
    
    If Cells(i, 1).Value Mod 5 = 0 And Cells(i, 1).Value Mod 3 = 0 Then
        Cells(i, 2).Value = "FizzBuzz"
    ElseIf Cells(i, 1).Value Mod 5 = 0 Then
        Cells(i, 2).Value = "Fizz"
    ElseIf Cells(i, 1).Value Mod 3 = 0 Then
        Cells(i, 2).Value = "Buzz"
        
        
    End If
Next i

End Sub

Sub ClassScanner():
    Dim targetStudent As String
        For i = 1 to 3
            For j = 1 to 5
                'MsgBox("row " + i + " Column: " + j + " | " + Cells(i,j).Value)
            next j 
        next i
    End Sub

'Hornet count

Sub HornetsNest():
    Dim HornetsCount As Integer
    Dim BugsCount As Integer
    Dim BeesCount As Integer

    BugsCount = Range("L2").value
    BeesCount = Range("R2").value

    HornetsCount = 0

    For i = 1 To 6
        For j = 1 to 7

            If Cells(i,j).Value = "Hornets" Then
            HornetsCount = HornetsCount + 1

            If (BugsCount > 0) Then
            Cells(i,j).Value = "Bugs"
            BugsCount = BugsCount - 1

            Elseif(BeesCount > 0) Then
            Cells(i,j).Value = "Bees"
            Beescount = Beescount - 1

            End if
        End If
    Next j 
next i 

msgbox(HornetsCount + " Hornets Found")

If (Range("L2").Value + Range(""))

End Sub 