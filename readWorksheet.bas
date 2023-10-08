Sub readWorksheet()
'A password protected xlsm workbook has VeryHidden excel worksheets.
'The content from the VerHidden sheets is required.
'This sub first prints out the name of the hidden worksheets to the immediate window.

'Using the name of the required sheets in the immediate window, the status will then be
'changed to become visible.


Dim folderpath As String
Dim filename As String
Dim wb As Workbook
Dim ws As Worksheet

folderpath = "C:\Users\Crispin Imani\Desktop\"
filename = "Book2.xlsx"

If Dir(folderpath & filename) <> "" Then
    Set wb = Workbooks.Open(folderpath & filename)
    
    For Each ws In wb.Worksheets
        Debug.Print ws.Name
    Next ws
    wb.Close savechanges:=False
Else
    MsgBox "no worksheets found."
End If

'A total of 6 worksheets with two currently visible.

If Dir(folderpath & filename) <> "" Then
    Set wb = Workbooks.Open(folderpath & filename)
    Set ws = wb.Sheets("Sheet2")
    ws.Visible = xlSheetVisible
    'wb.Close savechanges:=Save
End If
End Sub