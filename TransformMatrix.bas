Attribute VB_Name = "TransformMatrix"

Sub ChangeLength(shapeName As String, inputCell As String, Optional parentName As String, Optional basisCell As String)
    Dim workSheet As Worksheet : Set workSheet = ThisWorkbook.Worksheets("Sheet1")
    Dim inputCellValue As Double : inputCellValue = CDbl(workSheet.Range(inputCell).Value)

    ' Step 1: Identify the target shape by its name
    Dim targetShape As Shape : Set targetShape = workSheet.Shapes(shapeName)
    If targetShape Is Nothing Then
        MsgBox "Shape not found."
        Exit Sub
    End If
    
    ' Step 2: Find child shapes with shapeName in their name
    If Not IsMissing(parentName) Then

        For Each groupShape In workSheet.Shapes

            If groupShape.Name = parentName And groupShape.Type = msoGroup Then
                MsgBox "Found parent! " & groupShape.Name

                For Each childShape In groupShape.GroupItems
                    If InStr(1, childShape.Name, shapeName, vbTextCompare) > 0 And InStr(1, childShape.Name, parentName, vbTextCompare) > 0 Then
                        MsgBox "Found child: " & childShape.Name
                    End If
                Next childShape
                Exit For
            End If

        Exit For
        Next groupShape
    End If
End Sub

Function NormalizedDimensions(maxX As Double, maxY As Double, originalX As Double, originalY As Double) As Collection
    Dim normalizedValues As New Collection
    
    ' Check if values exceed the maximum
    If originalX > maxX Or originalY > maxY Then
        ' Calculate the sum of original values
        Dim sumOriginal As Double
        sumOriginal = originalX + originalY
        
        ' Normalize values based on the sum and maximum values
        Dim normalizedX As Double
        Dim normalizedY As Double
        normalizedX = (originalX / sumOriginal) * (maxX + maxY)
        normalizedY = (originalY / sumOriginal) * (maxX + maxY)
        
        ' Add named elements to the collection
        normalizedValues.Add normalizedX, "NormalX"
        normalizedValues.Add normalizedY, "NormalY"
    Else
        ' Values are within limits, no need to normalize
        ' Add named elements to the collection
        normalizedValues.Add originalX, "NormalX"
        normalizedValues.Add originalY, "NormalY"
    End If
    
    Set NormalizedDimensions = normalizedValues
End Function

Sub UpdateDimensions()
    'Debug.Print "Normalized X: " & NormalizedDimensions(200, 200, 300, 150)("NormalX")
    ChangeLength "length"
End Sub

