Attribute VB_Name = "TransformMatrix"
Sub ChangeLength(shapeName As String, cellName As String, Optional cellRatioName As String)
    Dim workSheet As Worksheet : Set workSheet = ThisWorkbook.Worksheets("Sheet1")
    Dim cellValue As Variant : cellValue = Range(cellName).Value
    
    'Check if grouped or not
    For Each targetShape In workSheet.Shapes
        If targetShape.Name = shapeName Then
            If targetShape.Type = msoGroup Then
                'Iterate through members of the group
                MsgBox "The targetShape is a group."
                MsgBox cellValue
                For Each shapeMember In targetShape.GroupItems
                    If IsMissing(cellRatioName) Then
                        MsgBox "No ratio"
                    Else
                        MsgBox shapeMember.Name
                    End If
                Next shapeMember
                Exit For
            Else
                'Change length of single shape
                MsgBox "The targetShape is not a group."
            End If
            Exit For
        End If
    Next targetShape
End Sub

Sub Test()
    ChangeLength "Column footing Length", "ColFootLength", "ColFootLength"
End Sub

