Attribute VB_Name = "Module4"
Sub Euros()
    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For i = lastrow To 2 Step -1
        If Cells(i, 9) <> "ok" Then
            If Cells(i, 1) = "BTECV" Or Cells(i, 1) = "FFPEUR" Or Cells(i, 1) = "GIC" Or Cells(i, 1) = "JOHCON" Or Cells(i, 1) = "JOHECV" Or Cells(i, 1) = "JOHSEL" Then
                Cells(i, 1).Interior.Color = vbMagenta
            End If
        End If
    Next i
End Sub
