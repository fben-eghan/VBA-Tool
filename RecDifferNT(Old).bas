Attribute VB_Name = "Module1"


Sub RecDifferNT()
    Dim lastrow As Long, i As Long
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

    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        Range("H" & i).Value = Abs(Range("G" & i).Value)
    Next i


    With ActiveSheet
    End With
    Range("A1:I" & lastrow).Sort Key1:=Range("I1"), _
                     Order1:=xlAscending, _
                     Key2:=Range("H1"), _
                     Order2:=xlAscending, _
                     Key3:=Range("G1"), _
                     Order3:=xlDescending, _
                     Header:=xlYes

    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 8) <> Cells(i - 1, 8) And Cells(i + 1, 8) <> Cells(i, 8) And Cells(i, 9).Value <> "JOHGLO" And Cells(i, 8) <> 0 Then
            Cells(i, 7).Interior.Color = vbYellow
            Cells(i, 10).Value = "no"
        ElseIf Cells(i, 9).Value = "JOHGLO" Then
            Cells(i, 10).Value = "b/s"
        Else: Cells(i, 10).Value = "ok"
        End If
    Next i
    
    With ActiveSheet
    End With
    Range("A2:J" & lastrow).Sort Key1:=Range("J2"), _
                     Order1:=xlAscending, _
                     Key2:=Range("I2"), _
                     Order2:=xlAscending, _
                     Key3:=Range("H2"), _
                     Order3:=xlDescending, _
                     Header:=xlNo
    
    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 10) <> "ok" Then
            If Cells(i, 9) = "BARCIRE" Or Cells(i, 9) = "HLHI" Or Cells(i, 9) = "HLIG" Or Cells(i, 9) = "RUSSELLAPC" Or Cells(i, 9) = "SWIPUKO" Or Cells(i, 9) = "JOHUKDYN" Or Cells(i, 9) = "JOHUKEI" Or Cells(i, 1) = "JOHUKGR" Or Cells(i, 9) = "JOHUKOP" Or Cells(i, 9) = "IRUKDYN" Then
                Cells(i, 1).Interior.Color = vbCyan
            End If
        End If
    Next i
    
        With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 10) <> "ok" Then
            If Cells(i, 9) = "BTECV" Or Cells(i, 9) = "FFPEUR" Or Cells(i, 9) = "GIC" Or Cells(i, 9) = "JOHCON" Or Cells(i, 9) = "JOHECV" Or Cells(i, 9) = "JOHSEL" Then
                Cells(i, 1).Interior.Color = vbMagenta
            End If
        End If
    Next i
    
    With ActiveSheet
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 9) <> Cells(i - 1, 9) And Cells(i, 4) <> Cells(i - 1, 4) Then
            Rows(i).insert Shift:=xlShiftDown
        End If
    Next i


    Cells.EntireColumn.AutoFit
    Range("A1:I1").Select
    With Selection.Interior
        .Color = 5287936
    End With
    Selection.Font.Bold = True
    Columns("H:I").Select
    Selection.Delete Shift:=xlToLeft

    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11

    End With
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
End Sub

