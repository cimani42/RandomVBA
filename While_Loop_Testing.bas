Option Explicit
Sub while_loops()
Dim userentry As Integer
Dim iterations As Integer
Dim columniter As Integer

Columns("B:B").Select
Selection.ClearContents
Columns("C:C").Select
Selection.ClearContents

Cells(1, 1).Interior.Color = vbCyan
Cells(1, 2).Interior.ColorIndex = 33
Cells(1, 3).Interior.ColorIndex = 34

MsgBox "enter a number between in cell A1 between 2 and 32767", vbOKOnly
userentry = Cells(1, 1).Value

If userentry > 32767 Or userentry < 2 Then
    MsgBox "please enter an whole number bewtween 2 and 32767", vbOKOnly
    Exit Sub
End If

iterations = 1
columniter = 1
'While userentry < 32767 'less than int value max
While iterations < userentry
    If iterations Mod 2 = 0 Then
        'columniter = 1
        Cells(columniter, 2) = "iteration " & iterations
        Cells(columniter, 3) = iterations + userentry
        columniter = columniter + 1
    Else
        Cells(1, 4) = "Difference =" & (32767 - userentry)
        Cells(1, 4).Interior.ColorIndex = 27
        Cells(1, 5).Interior.ColorIndex = 27
    End If
    iterations = iterations + 1
Wend

Cells(1, 7).Interior.ColorIndex = 7
Cells(1, 7).Value = Application.WorksheetFunction.Sum(Columns("C:C"))
Cells(1, 8) = "<-- Sum of Column C"
Debug.Print "Iterations = " & iterations
Debug.Print "User Entry = " & userentry

End Sub