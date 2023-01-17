Attribute VB_Name = "Module5"
Sub PfoTransform1()
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For i = lastrow To 2 Step -1
        If Left(Cells(i, 1).Value, 1) = "T" And Left(Cells(i, 1).Value, 3) <> "TFL" And Left(Cells(i, 1).Value, 3) <> "TST" Then
                Cells(i, 9).Value = Right(Cells(i, 1).Value, Len(Cells(i, 1).Value) - 1)
                
        ElseIf Left(Cells(i, 1).Value, 2) = "TT" And Left(Cells(i, 1).Value, 4) <> "TTFL" Then
                Cells(i, 9).Value = Right(Cells(i, 1).Value, Len(Cells(i, 1).Value) - 1)
        Else: Cells(i, 9).Value = Cells(i, 1).Value
        End If
    Next i
End Sub
