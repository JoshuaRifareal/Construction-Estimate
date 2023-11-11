' WORKSHEET CHANGE
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Update Column Footing
    If Not Intersect(Target, Me.Range("ColFootLength:ColFootThick")) Is Nothing Then
        ColumnFooting.UpdateLength
        ColumnFooting.UpdateWidth
        ColumnFooting.UpdateThickness
    End If
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    HideDropdown "ColFoot Mix Chip", "ColFoot Mix Options", False
End Sub

' COLUMN FOOTING
Private Sub ColFoot_Mix_Option1Hover_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HighlightOption "ColFoot Mix Option 1", True
    HighlightOption "ColFoot Mix Option 2"
    HighlightOption "ColFoot Mix Option 3"
    HighlightOption "ColFoot Mix Option 4"
End Sub
Private Sub ColFoot_Mix_Option2Hover_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HighlightOption "ColFoot Mix Option 1"
    HighlightOption "ColFoot Mix Option 2", True
    HighlightOption "ColFoot Mix Option 3"
    HighlightOption "ColFoot Mix Option 4"
End Sub
Private Sub ColFoot_Mix_Option3Hover_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HighlightOption "ColFoot Mix Option 1"
    HighlightOption "ColFoot Mix Option 2"
    HighlightOption "ColFoot Mix Option 3", True
    HighlightOption "ColFoot Mix Option 4"
End Sub
Private Sub ColFoot_Mix_Option4Hover_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HighlightOption "ColFoot Mix Option 1"
    HighlightOption "ColFoot Mix Option 2"
    HighlightOption "ColFoot Mix Option 3"
    HighlightOption "ColFoot Mix Option 4", True
End Sub
Private Sub ColFoot_Mix_OptionsLeave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HighlightOption "ColFoot Mix Option 1"
    HighlightOption "ColFoot Mix Option 2"
    HighlightOption "ColFoot Mix Option 3"
    HighlightOption "ColFoot Mix Option 4"
End Sub