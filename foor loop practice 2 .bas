Sub for_loop_practice2()

Dim year As Integer
Dim number As Integer
Dim i As Integer
Dim x As Integer
'
Cells(3, 1).Interior.ColorIndex = 7
Cells(4, 1).Interior.ColorIndex = 7
Cells(3, 2).Value = "<-- Enter the year you want to check"
Cells(4, 2).Value = "<-- Enter the number of proceeding years you want to check"

year = Cells(3, 1).Value
number = Cells(4, 1).Value

If number < 1 Then
    MsgBox "Please enter a number in cell A4 greater than or equal to 1", vbOKOnly
Else
    x = 8
    For i = 1 To number
        
        If year Mod 4 = 0 And (year Mod 100 <> 0 Or year Mod 400 = 0) Then
            Cells(x, 3).Value = year & " is a leap year."
        Else
            Cells(x, 3).Value = year & " isn't a leap year."
        End If
        x = x + 1
        year = year + 1
    Next i
End If
End Sub