Option Explicit
Sub while_loops()
Dim userentry As Integer
Dim iterations As Integer
Dim columniter As Integer

Columns("B:B").Select
Selection.ClearContents
Columns("C:C").Select
Selection.ClearContents

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
        Cells(1, 2) = 32767 - userentry
    End If
    iterations = iterations + 1
Wend

Debug.Print "Iterations = " & iterations
Debug.Print "User Entry = " & userentry

End Sub