Attribute VB_Name = "Module5"
Sub PfoTransform()
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For i = lastrow To 2 Step -1
        If Left(Cells(i, 1).Value, 1) = "T" Then
                Cells(i, 1).Value = Right(Cells(i, 1).Value, Len(Cells(i, 1).Value) - 1)
        End If
    Next i
End Sub
