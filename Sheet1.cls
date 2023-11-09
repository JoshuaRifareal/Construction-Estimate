Private Sub Worksheet_Change(ByVal Target As Range)

    ' Update Column Footing
    If Not Intersect(Target, Me.Range("K12:K14")) Is Nothing Then
        ColumnFooting.UpdateLength
        ColumnFooting.UpdateWidth
        ColumnFooting.UpdateThickness
    End If
End Sub

