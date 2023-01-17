Attribute VB_Name = "Module3"
Sub UKs()
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 9) <> "ok" Then
            If Cells(i, 1) = "BARCIRE" Or Cells(i, 1) = "HLHI" Or Cells(i, 1) = "HLIG" Or Cells(i, 1) = "RUSSELLAPC" Or Cells(i, 1) = "SWIPUKO" Or Cells(i, 1) = "JOHUKDYN" Or Cells(i, 1) = "JOHUKEI" Or Cells(i, 1) = "JOHUKGR" Or Cells(i, 1) = "JOHUKOP" Or Cells(i, 1) = "IRUKDYN" Then
                Cells(i, 1).Interior.Color = vbCyan
            End If
        End If
    Next i
End Sub
