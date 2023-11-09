Attribute VB_Name = "ColumnFooting"

Sub UpdateLength()
    ChangeDimension "Width", "Length 1",  "ColFoot Drawing", "ColFootLength", 2, "Width", "Width 1", "ColFootWidth", 2, 2.5, 2
    ChangeDimension "Width", "Length 2",  "ColFoot Drawing", "ColFootLength", 2, "Width", "Width 2", "ColFootWidth", 2, 2.5, 2
    ChangeDimension "Width", "Length 3",  "ColFoot Drawing", "ColFootLength", 2, "Width", "Width 3", "ColFootWidth", 2, 2.5, 2

    MoveRelativeShape "Width 3", "Length 1"
End Sub