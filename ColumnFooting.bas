Attribute VB_Name = "ColumnFooting"

Public Sub UpdateLength()
    ChangeDimension "Length 1", "ColFoot Drawing", "ColFootLength", 2, "Width 1", "ColFootWidth", 2, 2.5, 2
    ChangeDimension "Length 2", "ColFoot Drawing", "ColFootLength", 2, "Width 2", "ColFootWidth", 2, 2.5, 2
    ChangeDimension "Length 3", "ColFoot Drawing", "ColFootLength", 2, "Width 3", "ColFootWidth", 2, 2.5, 2

    MoveRelativeShape "Width 3", "Length 1", True
    MoveRelativeShape "Thickness 3", "Length 2", True
End Sub

Public Sub UpdateWidth()
    ChangeDimension "Width 1", "ColFoot Drawing", "ColFootWidth", 2, "Length 1", "ColFootLength", 2, 2, 2.5
    ChangeDimension "Width 2", "ColFoot Drawing", "ColFootWidth", 2, "Length 2", "ColFootLength", 2, 2, 2.5
    ChangeDimension "Width 3", "ColFoot Drawing", "ColFootWidth", 2, "Length 3", "ColFootLength", 2, 2, 2.5

    MoveRelativeShape "Width 3", "Length 1", True
    MoveRelativeShape "Thickness 3", "Length 2", True
End Sub