Sub Data_Cleaner_Keep_Recent()
    Dim xRng As Range
    Dim xTxt As String
    On Error Resume Next
    xTxt = Application.ActiveWindow.RangeSelection.Address
    Set xRng = Application.InputBox("please select the data range:", "Data Cleaner", xTxt, , , , , 8)
    If xRng Is Nothing Then Exit Sub
    If (xRng.Columns.Count < 2) Or (xRng.Rows.Count < 2) Then
        MsgBox "the used range is invalid", , "Data Cleaner"
        Exit Sub
    End If
    xRng.Sort key1:=xRng.Cells(1, 1), Order1:=xlAscending, key2:=xRng.Cells(1, 2), Order2:=xlDescending, Header:=xlGuess
    xRng.RemoveDuplicates Columns:=1, Header:=xlGuess
End Sub