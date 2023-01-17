Attribute VB_Name = "Module1"


Sub RecDiffer()
    Dim lastrow As Long, i As Long
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For i = lastrow To 2 Step -1
        Range("H" & i).Value = Abs(Range("G" & i).Value)
    Next i


    With ActiveSheet
    End With
    Range("A1:H" & lastrow).Sort Key1:=Range("A1"), _
                     Order1:=xlAscending, _
                     Key2:=Range("H1"), _
                     Order2:=xlAscending, _
                     Key3:=Range("G1"), _
                     Order3:=xlDescending, _
                     Header:=xlYes

    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 8) <> Cells(i - 1, 8) And Cells(i + 1, 8) <> Cells(i, 8) And Cells(i, 1).Value <> "JOHGLO" And Cells(i, 8) <> 0 Then
            Cells(i, 7).Interior.Color = vbYellow
            Cells(i, 9).Value = "no"
        ElseIf Cells(i, 1).Value = "JOHGLO" Then
            Cells(i, 9).Value = "b/s"
        Else: Cells(i, 9).Value = "ok"
        End If
    Next i
    
    With ActiveSheet
    End With
    Range("A2:I" & lastrow).Sort Key1:=Range("I2"), _
                     Order1:=xlAscending, _
                     Key2:=Range("A2"), _
                     Order2:=xlAscending, _
                     Key3:=Range("H2"), _
                     Order3:=xlDescending, _
                     Header:=xlNo
    
    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 9) <> "ok" Then
            If Cells(i, 1) = "BARCIRE" Or Cells(i, 1) = "HLHI" Or Cells(i, 1) = "HLIG" Or Cells(i, 1) = "RUSSELLAPC" Or Cells(i, 1) = "SWIPUKO" Or Cells(i, 1) = "JOHUKDYN" Or Cells(i, 1) = "JOHUKEI" Or Cells(i, 1) = "JOHUKGR" Or Cells(i, 1) = "JOHUKOP" Or Cells(i, 1) = "IRUKDYN" Then
                Cells(i, 1).Interior.Color = vbCyan
            End If
        End If
    Next i
    
        With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 9) <> "ok" Then
            If Cells(i, 1) = "BTECV" Or Cells(i, 1) = "FFPEUR" Or Cells(i, 1) = "GIC" Or Cells(i, 1) = "JOHCON" Or Cells(i, 1) = "JOHECV" Or Cells(i, 1) = "JOHSEL" Then
                Cells(i, 1).Interior.Color = vbMagenta
            End If
        End If
    Next i
    
    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 1) <> Cells(i - 1, 1) And Cells(i, 4) <> Cells(i - 1, 4) Then
            Rows(i).insert Shift:=xlShiftDown
        End If
    Next i


    Cells.EntireColumn.AutoFit
    Range("A1:G1").Select
    With Selection.Interior
        .Color = 5287936
    End With
    Selection.Font.Bold = True
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft

    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11

    End With
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
End Sub

