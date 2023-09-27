Sub for_loop_practice1()

Dim iteration As Integer
Dim usernumber As Integer
Dim rowLoop As Integer
Dim i As Integer


usernumber = Cells(1, 1).Value
Cells(1, 1).Interior.ColorIndex = 33

'Finding the answers for the first twelve times table

Cells(3, 3) = "Multiple"
Cells(3, 4) = "Calculation"
Cells(3, 5) = "Output"

MsgBox "Please enter a number in cell A1", vbOKOnly

rowLoop = 4

For i = 1 To 12
    Cells(rowLoop, 3) = i
    Cells(rowLoop, 4) = usernumber & " * " & i
    Cells(rowLoop, 5) = usernumber * i
    rowLoop = rowLoop + 1
Next i
End Sub