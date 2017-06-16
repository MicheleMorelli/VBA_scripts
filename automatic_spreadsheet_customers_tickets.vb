Sub clean_the_sheet()
Dim i As Integer

'clear filters
Call clear_all_filters
i = 0
Worksheets(2).Activate
' DATE PARSER - last comment date
Call date_parser
'comment parser
'Call comment_parser
'remove unused colums

Do While (i < 5)
    Call remove_columns
    i = i + 1
Loop

' move the date of last comment on the right of the reference - NOT VERY EFFICIENT!
i = 1

Do While (i < 15)
    If (Cells(1, i) = "last_comment_date") Then
    Cells(1, i).EntireColumn.Cut
    Columns(2).Insert shift:=xlRight
   Exit Do
  End If
  i = i + 1
Loop

Application.CutCopyMode = False
'add blank column
Columns(2).Insert shift:=xlToRight
'append assigned to at the end
i = 1

Do While (i < 10)
    If (Cells(1, i) = "Assigned to Person") Then
    Cells(1, i).EntireColumn.Cut
    Cells(1, i).Offset(0, 2).Insert shift:=xlRight
   Exit Do
  End If
  i = i + 1
Loop

Application.CutCopyMode = False
'cleans the first row
Rows(1).Delete
'cut and paste to sheet 1
Call find_duplicates
Call remove_duplicates
Call cut_and_paste_in_main_sheet
'change format to date
Worksheets(1).Range("C:C").NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
'apply formula to column b
Worksheets(1).Activate
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B500"), Type:=xlFillDefault
Worksheets(3).Activate
End Sub


Sub remove_columns()
Dim col As Range
Dim i As Integer
For i = 1 To 15 Step 1
    If (Cells(1, i).Value <> "Reference" And Cells(1, i).Value <> "Assigned to Person" And Cells(1, i).Value <> "Company" And Cells(1, i).Value <> "Contact for this Call" And Cells(1, i).Value <> "last_comment_date") Then
    Cells(1, i).EntireColumn.Delete
    End If
Next i
End Sub


Sub find_duplicates()
Dim s1, s2 As String
Dim i, j, wsn As Integer
i = 1
Do While (i <= 50)
    j = 1
    s1 = Worksheets(2).Cells(i, 1)
    Do While (j <= 50)
        s2 = Worksheets(1).Cells(j, 1)
        If (s1 = s2 And Not IsEmpty(Worksheets(2).Cells(i, 1))) Then
            Worksheets(2).Cells(i, 1).Interior.Color = vbRed
        End If
        j = j + 1
    Loop
    j = 1
    i = i + 1
Loop
End Sub


Sub remove_duplicates()
Dim rng As Range
Dim i As Integer

For i = 1 To 100 Step 1
    If (Cells(i, 1).Interior.Color = vbRed) Then
        If (rng Is Nothing) Then
            Set rng = Cells(i, 1)
        Else
            Set rng = Union(rng, Cells(i, 1))
        End If
    End If
Next i
If (Not rng Is Nothing) Then
    rng.EntireRow.Select
    Selection.Delete
End If
End Sub


Sub cut_and_paste_in_main_sheet()
Dim pst_rng As Range
'first find the last used row in sheet 1
Set pst_rng = Worksheets(1).Range("A1").End(xlDown)
Worksheets(2).Range("A1", "F100").Cut pst_rng.Offset(1, 0)
End Sub

Sub clear_all_filters()
If Worksheets(1).FilterMode Then
Worksheets(1).ShowAllData
End If
End Sub

Sub date_parser()
Dim i As Integer
Dim j As Integer
Dim r As Range
i = 1
j = 2
' find the comments column
Do While (Not IsEmpty(Cells(1, i)))
    If (Cells(1, i) = "Comments and Work notes") Then
        Exit Do
    End If
    i = i + 1
Loop
'field title
Cells(1, i).Offset(0, 1) = "last_comment_date"
Cells(1, i).Offset(0, 1).Font.Bold = True
'parse the date
Do While (Not IsEmpty(Cells(j, i)))
    Cells(j, i).Offset(0, 1) = Mid(Cells(j, i), 7, 4) & "/" & Mid(Cells(j, i), 4, 2) & "/" & Left(Cells(j, i), 2)
    j = j + 1
Loop
End Sub


Sub comment_parser()
Dim i As Integer
Dim j As Integer
Dim r As Range
Dim comm As Comment
i = 1
j = 2
' find the comments column
Do While (Not IsEmpty(Cells(1, i)))
    If (Cells(1, i) = "Comments and Work notes") Then
        Exit Do
    End If
    i = i + 1
Loop

'field title
Cells(1, i).Offset(0, 2) = "Comment_body"
Cells(1, i).Offset(0, 2).Font.Bold = True

'parse the comments and put them in the spreadsheet - puts comment in the date

Do While (Not IsEmpty(Cells(j, i)))
    Cells(j, i).Offset(0, 1).AddComment Right(Cells(j, i), Len(Cells(j, i)) - 22)
    Set comm = Cells(j, i).Offset(0, 1).Comment
    comm.Shape.TextFrame.AutoSize = True
    j = j + 1
Loop

End Sub


Sub add_hyperlink()
Dim i As Integer

Worksheets(1).Activate

i = 2
Do While (Not IsEmpty(Cells(i, 1)))
    Worksheets(1).Hyperlinks.Add anchor:=Cells(i, 1), Address:="https://ThisIsATest.itwqwq/" & Cells(i, 1)
    i = i + 1
Loop
End Sub