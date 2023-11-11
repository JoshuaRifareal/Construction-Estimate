Attribute VB_Name = "DropdownHandler"

Public Sub HideDropdown(chipName As String, dropdownName As String, Optional isSwitch As Boolean = True)
    If isSwitch Then
        If ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse Then
            ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoTrue
            ChangeShapeChip chipName, msoShapeRound2SameRectangle
        Else
            ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse
            ChangeShapeChip chipName, msoShapeRoundedRectangle
        End If
    Else
        ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse
        ChangeShapeChip chipName, msoShapeRoundedRectangle
    End If
End Sub

Public Sub PositionDropdown(chipName As String, dropdownName As String, option1Name As String, panelName As String, leaveName As String)
    ActiveSheet.Shapes.Range(Array(dropdownName)).Left = ActiveSheet.Shapes.Range(Array(chipName)).Left + (ActiveSheet.Shapes.Range(Array(chipName)).Width / 2) - (ActiveSheet.Shapes.Range(Array(dropdownName)).Width / 2) - 1
    ActiveSheet.Shapes.Range(Array(dropdownName)).Top = ActiveSheet.Shapes.Range(Array(chipName)).Top + ActiveSheet.Shapes.Range(Array(chipName)).Height - 1
    ActiveSheet.Shapes.Range(Array(leaveName)).Top = ActiveSheet.Shapes.Range(Array(dropdownName)).Top
    ActiveSheet.Shapes.Range(Array(option1Name)).Left = ActiveSheet.Shapes.Range(Array(chipName)).Left
    ActiveSheet.Shapes.Range(Array(panelName)).Top = ActiveSheet.Shapes.Range(Array(leaveName)).Top + 1
    ActiveSheet.Shapes.Range(Array(panelName)).Left = ActiveSheet.Shapes.Range(Array(chipName)).Left

    ActiveSheet.Shapes.Range(Array(panelName)).Width = ActiveSheet.Shapes.Range(Array(option1Name)).Width
End Sub

Public Sub PositionOptions(chipName As String, optionName As String, hoverName As String, Optional siblingName As String)
    If siblingName <> "" Then
        'Go below sibling
        ActiveSheet.Shapes.Range(Array(optionName)).Top = ActiveSheet.Shapes.Range(Array(siblingName)).Top + ActiveSheet.Shapes.Range(Array(siblingName)).Height
        ActiveSheet.Shapes.Range(Array(optionName)).Left = ActiveSheet.Shapes.Range(Array(siblingName)).Left
        ActiveSheet.Shapes.Range(Array(optionName)).Width = ActiveSheet.Shapes.Range(Array(siblingName)).Width
    Else
        'No sibling goes under chip
        ActiveSheet.Shapes.Range(Array(optionName)).Top = ActiveSheet.Shapes.Range(Array(chipName)).Top + ActiveSheet.Shapes.Range(Array(chipName)).Height
        ActiveSheet.Shapes.Range(Array(optionName)).Width = ActiveSheet.Shapes.Range(Array(chipName)).Width
    End If

    'Hover always inherits
    ActiveSheet.Shapes.Range(Array(hoverName)).Top = ActiveSheet.Shapes.Range(Array(optionName)).Top
    ActiveSheet.Shapes.Range(Array(hoverName)).Left = ActiveSheet.Shapes.Range(Array(optionName)).Left
    ActiveSheet.Shapes.Range(Array(hoverName)).Width = ActiveSheet.Shapes.Range(Array(optionName)).Width
End Sub

Public Sub ChangeShapeChip(chipName As String, shapeType As MsoAutoShapeType)
    ActiveSheet.Shapes(chipName).AutoShapeType = shapeType
    ActiveSheet.Shapes(chipName).Adjustments(1) = 1
End Sub

Public Sub HighlightOption(optionName As String, Optional isHighlight As Boolean = False)
    Dim highglightColor As Long : highglightColor = RGB(221, 235, 247)
    Dim staticColor As Long : staticColor = RGB(255, 255, 255)

    If isHighlight =  True Then
        ActiveSheet.Shapes.Range(Array(optionName)).Fill.ForeColor.RGB = highglightColor
    Else
        ActiveSheet.Shapes.Range(Array(optionName)).Fill.ForeColor.RGB = staticColor
    End If
End Sub