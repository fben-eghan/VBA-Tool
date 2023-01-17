Attribute VB_Name = "Module2"
Sub CashDifferNT()
    Dim lastrow As Long, i As Long
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For i = lastrow To 2 Step -1
        Range("F" & i).Value = Abs(Range("E" & i).Value)
    Next i


    With ActiveSheet
    End With
    Range("A1:F" & lastrow).Sort Key1:=Range("A1"), _
                     Order1:=xlAscending, _
                     Key2:=Range("F1"), _
                     Order2:=xlAscending, _
                     Key3:=Range("E1"), _
                     Order3:=xlDescending, _
                     Header:=xlYes

    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 6) <> Cells(i - 1, 6) And Cells(i + 1, 6) <> Cells(i, 6) And Cells(i, 1).Value <> "JOHGLO" And Cells(i, 6) <> 0 Then
            Cells(i, 5).Interior.Color = vbYellow
            Cells(i, 7).Value = "no"
        ElseIf Cells(i, 1).Value = "JOHGLO" Then
            Cells(i, 7).Value = "b/s"
        Else: Cells(i, 7).Value = "ok"
        End If
    Next i
    
    With ActiveSheet
    End With
    Range("A2:G" & lastrow).Sort Key1:=Range("E2"), _
                     Order1:=xlAscending, _
                     Key2:=Range("A2"), _
                     Order2:=xlAscending, _
                     Key3:=Range("F2"), _
                     Order3:=xlDescending, _
                     Header:=xlNo
    
    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 7) <> "ok" Then
            If Cells(i, 1) = "BARCIRE" Or Cells(i, 1) = "HLHI" Or Cells(i, 1) = "HLIG" Or Cells(i, 1) = "RUSSELLAPC" Or Cells(i, 1) = "SWIPUKO" Or Cells(i, 1) = "JOHUKDYN" Or Cells(i, 1) = "JOHUKEI" Or Cells(i, 1) = "JOHUKGR" Or Cells(i, 1) = "JOHUKOP" Or Cells(i, 1) = "IRUKDYN" Then
                Cells(i, 1).Interior.Color = vbCyan
            End If
        End If
    Next i
    
        With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 7) <> "ok" Then
            If Cells(i, 1) = "BTECV" Or Cells(i, 1) = "FFPEUR" Or Cells(i, 1) = "GIC" Or Cells(i, 1) = "JOHCON" Or Cells(i, 1) = "JOHECV" Or Cells(i, 1) = "JOHSEL" Then
                Cells(i, 1).Interior.Color = vbMagenta
            End If
        End If
    Next i
    
    Cells.EntireColumn.AutoFit
    Range("A1:E1").Select
    With Selection.Interior
        .Color = 5287936
    End With
    Selection.Font.Bold = True
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft

    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11

    End With
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
End Sub
