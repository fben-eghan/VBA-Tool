Attribute VB_Name = "PreRectifier"

Sub AbsCol()
    Dim lastrow As Long, i As Long
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For i = lastrow To 2 Step -1
        Range("H" & i).Value = Abs(Range("G" & i).Value)
    Next i
End Sub

Sub RecSorter()
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Range("A1:H" & lastrow).Sort Key1:=Range("A1"), _
                     Order1:=xlAscending, _
                     Key2:=Range("H1"), _
                     Order2:=xlAscending, _
                     Key3:=Range("G1"), _
                     Order3:=xlDescending, _
                     Header:=xlYes

End Sub

Sub Inserter()
    Dim lastrow As Long, i As Long
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For i = lastrow To 2 Step -1
        If Not Cells(i, 1) = Cells(i - 1, 1) Then
            Rows(i).insert Shift:=xlShiftDown
        End If
    Next i
End Sub
Sub Aligner()
    Cells.EntireColumn.AutoFit
    Range("A1:G1").Select
    With Selection.Interior
        .Color = 5287936
    End With
    Selection.Font.Bold = True
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
End Sub

