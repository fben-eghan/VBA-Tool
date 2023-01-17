Attribute VB_Name = "Module2"
Sub CashDiffer()
    Dim lastrow As Long, i As Long
    With ActiveSheet
            lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Range("A1:E" & lastrow).Sort Key1:=Range("A1"), _
                     Order1:=xlAscending, _
                     Key2:=Range("B1"), _
                     Order2:=xlAscending, _
                     Header:=xlYes
    Cells.EntireColumn.AutoFit
    Range("A1:E1").Select
    With Selection.Interior
        .Color = 5287936
    End With
    Selection.Font.Bold = True

    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
    End With
End Sub
