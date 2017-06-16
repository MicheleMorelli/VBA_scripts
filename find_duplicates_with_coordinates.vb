Sub dup()
Dim i As Integer
Dim j As Integer
Dim lim As Integer
Dim ran As Range
Dim par As Range
'Finding duplicates and adding their coordinates - MM
lim = InputBox("enter the number of lines you want to check: ")

For i = 1 To lim Step 1
    Set ran = Cells(i, 1)
    For j = i + 1 To lim Step 1
        Set par = Cells(j, 1)
        If (ran = par) Then
            par.Interior.Color = RGB(0, 0, 0)
            par.Offset(0, 1) = "The Duplicate is in A" & i
        End If
        j = j + 1
    Next j
Next i
End Sub