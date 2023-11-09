Private Sub Worksheet_Change(ByVal Target As Range)
    Dim WatchRange As Range: Set WatchRange = Me.Range("K12:K14")

    ' Update Column Footing
    If Not Intersect(Target, WatchRange) Is Nothing Then
        ColumnFooting.UpdateLength
        ColumnFooting.UpdateWidth
        
    End If
End Sub

