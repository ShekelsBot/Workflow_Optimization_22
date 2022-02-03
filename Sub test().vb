'Andrew Bruckbauer'
'2/1/2022'
'Excel Orginize and delete most concurrent occurance'
'Purpose of this script is to orginize data from most recent and A to Z'
'Then remove duplicates'

Sub Data_Cleaner_Keep_Recent()
    Dim xRng As Range
    Dim xTxt As String
    On Error Resume Next
    xTxt = Application.ActiveWindow.RangeSelection.Address
    Set xRng = Application.InputBox("please select the data range:", "Data Cleaner", xTxt, , , , , 8)
    'show box and give instructions'
    If xRng Is Nothing Then Exit Sub
    'if range is empty exit'
    If (xRng.Columns.Count < 2) Or (xRng.Rows.Count < 2) Then
        MsgBox "the used range is invalid", , "Data Cleaner"
        Exit Sub
    End If
    xRng.Sort key1:=xRng.Cells(1, 1), Order1:=xlAscending, key2:=xRng.Cells(1, 2), Order2:=xlDescending, Header:=xlGuess
    'order dates accending'
    xRng.RemoveDuplicates Columns:=1, Header:=xlGuess
    'remove dupes'
    'since the dates are from most recent and according to a-z the first result is kept as the original'
End Sub