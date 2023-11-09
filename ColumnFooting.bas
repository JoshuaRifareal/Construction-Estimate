Attribute VB_Name = "ColumnFooting"

Sub ResetProps()
    ChangeLength "Length 1", 2.5
    ChangeLength "Length 2", 2.5
    ChangeLength "Length 3", 2.5
    ChangeLength "Width 1", 2
    ChangeLength "Width 2", 2
    ChangeLength "Width 3", 2
    ChangeLength "Thickness 1", 0.6
    ChangeLength "Thickness 2", 0.6
    ChangeLength "Thickness 3", 0.6
    AttachPoints
End Sub

Sub AttachPoints()
    MoveRelativeShape "Width 1", "Thickness 1", True
    MoveRelativeShape "Width 3", "Length 1", True
    MoveRelativeShape "Thickness 2", "Width 2", True
    MoveRelativeShape "Thickness 3", "Length 2", True
    MoveRelativeShape "Thickness 3", "Width 3", True
    MoveRelativeShape "Length 2", "Width 2", True
    MoveRelativeShape "Length 3", "Thickness 2", True
End Sub

Public Sub UpdateLength()
    ChangeDimension "Length 1", "ColFoot Drawing", "ColFootLength", 2, "Width 1", "ColFootWidth", 2, 2.5, 2
    ChangeDimension "Length 2", "ColFoot Drawing", "ColFootLength", 2, "Width 2", "ColFootWidth", 2, 2.5, 2
    ChangeDimension "Length 3", "ColFoot Drawing", "ColFootLength", 2, "Width 3", "ColFootWidth", 2, 2.5, 2

    AttachPoints
End Sub

Public Sub UpdateWidth()
    ChangeDimension "Width 1", "ColFoot Drawing", "ColFootWidth", 2, "Length 1", "ColFootLength", 2, 2, 2.5
    ChangeDimension "Width 2", "ColFoot Drawing", "ColFootWidth", 2, "Length 2", "ColFootLength", 2, 2, 2.5
    ChangeDimension "Width 3", "ColFoot Drawing", "ColFootWidth", 2, "Length 3", "ColFootLength", 2, 2, 2.5
    
    AttachPoints
End Sub

Public Sub UpdateThickness()
    ChangeDimension "Thickness 1", "ColFoot Drawing", "ColFootThick", 0.6
    ChangeDimension "Thickness 2", "ColFoot Drawing", "ColFootThick", 0.6
    ChangeDimension "Thickness 3", "ColFoot Drawing", "ColFootThick", 0.6

    AttachPoints
End Sub