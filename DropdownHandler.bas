Attribute VB_Name = "DropdownHandler"

Public Sub HideDropdown(chipName As String, dropdownName As String)
    If ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse Then
        ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoTrue
        ChangeShapeChip chipName, msoShapeRectangle
    Else
        ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse
        ChangeShapeChip chipName, msoShapeRoundedRectangle
    End If
End Sub

Public Sub PositionDropdown(chipName As String, dropdownName As String)
    ActiveSheet.Shapes.Range(Array(dropdownName)).Left = ActiveSheet.Shapes.Range(Array(chipName)).Left
    ActiveSheet.Shapes.Range(Array(dropdownName)).Top = ActiveSheet.Shapes.Range(Array(chipName)).Top + ActiveSheet.Shapes.Range(Array(chipName)).Height
End Sub

Public Sub ChangeShapeChip(chipName As String, shapeType As MsoAutoShapeType)
    ActiveSheet.Shapes(chipName).AutoShapeType = shapeType

    If shapeType = msoShapeRoundedRectangle Then
        ActiveSheet.Shapes(chipName).Adjustments(1) = 1
    End If
End Sub

Public Sub HighlightChip()
End Sub