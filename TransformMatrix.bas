Attribute VB_Name = "TransformMatrix"

Sub ChangeLength(shapeName As String)
    'Change length of shape based on input and ratio
    Dim workSheet As Worksheet : Set workSheet = ThisWorkbook.Worksheets("Sheet1")

    'Check if grouped or not
    For Each targetShape In workSheet.Shapes
        If targetShape.Name = shapeName Then
            If targetShape.Type = msoGroup Then
                'Iterate through members of the group
                For Each shapeMember In targetShape.GroupItems

                Next shapeMember
                Exit For
            Else
                'Change length of single shape
                
            End If
            Exit For
        End If
    Next targetShape
End Sub

Function CheckRatio(shapeName As String, cellName As String, Optional cellRatioName As String) As Variant
    Dim workSheet As Worksheet : Set workSheet = ThisWorkbook.Worksheets("Sheet1")
    Dim cellValue As Variant : cellValue = Range(cellName).Value
    Dim cellRatio As Variant
    Dim originalLength As Variant : originalLength = Null

    'Check if there is ratio
    If Not IsMissing(cellRatioName) Then
        If IsNull(originalLength) Then
            originalLength = targetShape.Width
            Debug.Print "Original length: " & originalLength
        End if

        cellRatio = Range(cellRatioName).Value
        CheckRatio = (cellValue/cellRatio)
    End If
End Function

Sub UpdateColFootLength()
    If ChangeLength("tryline", "ColFootLength", "ColFootWidth") < 1 Then

    End If
End Sub

