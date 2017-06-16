' Various quick scripts to clean a customer list in Excel


Sub dup()
Dim i As Integer
Dim j As Integer
'removing duplicates in column A - MM
i = 1
Do While (i <= 60)
    j = i + 1
    Do While (j <= 60)
        If (Cells(i, 1) = Cells(j, 1)) Then
        Cells(j, 1) = ""
        End If
        j = j + 1
    Loop
    i = i + 1
Loop
End Sub


Sub ark()
' Ark. customers - MM
Dim i As Integer
i = 1
Do While (i <= 60)
    If (Cells(i, 2) <> "") Then
        Cells(i, 2) = "Yes"
    End If
    i = i + 1
Loop
End Sub


Sub col()
'colour customers MM
Dim i As Integer
Dim c As Integer
Dim ran As Range
 i = 3 
Do While (i <= 60)
    Set ran = Cells(i, 1)
    If (IsEmpty(ran)) Then
        ran.Interior.Color = ran.Offset(-1, 0).Interior.Color
        ran.Offset(0, 1).Interior.Color = ran.Offset(-1, 0).Interior.Color
        ran.Offset(0, 2).Interior.Color = ran.Offset(-1, 0).Interior.Color
    Else
        If (ran.Offset(-1, 0).Interior.Color = RGB(255, 255, 255)) Then
            c = 204
        Else
            c = 255
            End If
    End If
    ' colour the cells - MM
    ran.Interior.Color = RGB(c, c, c)
    ran.Offset(0, 1).Interior.Color = RGB(c, c, c)
    ran.Offset(0, 2).Interior.Color = RGB(c, c, c)
    i = i + 1
Loop
End Sub



Sub numb()
Dim i As Integer
Dim n As Integer
Dim ran As Range
 i = 3 
Do While (i <= 60)
    Set ran = Cells(i, 3)
    If (ran.Interior.Color = ran.Offset(-1, 0).Interior.Color) Then
        n = n + 1
    Else
        n = 1
    End If
    ' put count numb -MM
    ran = n
    i = i + 1
Loop
End Sub



Sub hyf()
' adds hyphen to col 3
Dim i As Integer
Dim ran As Range
 i = 2
Do While (i <= 60)
    Set ran = Cells(i, 3)
    ran = ran & "- "
    i = i + 1
Loop
End Sub
