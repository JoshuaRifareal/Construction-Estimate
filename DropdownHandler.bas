Attribute VB_Name = "DropdownHandler"

Public Sub HideDropdown(chipName As String, dropdownName As String)
    If ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse Then
        ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoTrue
        ChangeShapeChip chipName, msoShapeRound2SameRectangle
    Else
        ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse
        ChangeShapeChip chipName, msoShapeRoundedRectangle
    End If
End Sub

Public Sub PositionDropdown(chipName As String, dropdownName As String, leaveName As String)
    ActiveSheet.Shapes.Range(Array(dropdownName)).Left = ActiveSheet.Shapes.Range(Array(chipName)).Left + (ActiveSheet.Shapes.Range(Array(chipName)).Width / 2) - (ActiveSheet.Shapes.Range(Array(dropdownName)).Width / 2) - 1
    ActiveSheet.Shapes.Range(Array(dropdownName)).Top = ActiveSheet.Shapes.Range(Array(chipName)).Top + ActiveSheet.Shapes.Range(Array(chipName)).Height - 1
    ActiveSheet.Shapes.Range(Array(leaveName)).Top = ActiveSheet.Shapes.Range(Array(dropdownName)).Top
End Sub

Public Sub PositionOptions(chipName As String, optionName As String, hoverName As String, Optional siblingName As String)
    If siblingName <> "" Then
        'Go below sibling
        ActiveSheet.Shapes.Range(Array(optionName)).Top = ActiveSheet.Shapes.Range(Array(siblingName)).Top + ActiveSheet.Shapes.Range(Array(siblingName)).Height
    Else
        'No sibling goes under chip
        ActiveSheet.Shapes.Range(Array(optionName)).Top = ActiveSheet.Shapes.Range(Array(chipName)).Top + ActiveSheet.Shapes.Range(Array(chipName)).Height
    End If

    'Hover always inherits
    ActiveSheet.Shapes.Range(Array(hoverName)).Top = ActiveSheet.Shapes.Range(Array(optionName)).Top
End Sub

Public Sub ChangeShapeChip(chipName As String, shapeType As MsoAutoShapeType)
    ActiveSheet.Shapes(chipName).AutoShapeType = shapeType
    ActiveSheet.Shapes(chipName).Adjustments(1) = 1
End Sub

Public Sub HighlightChip()
End Sub