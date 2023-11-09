Attribute VB_Name = "TransformMatrix"

Sub ChangeDimension(prop As String, shapeName As String, parentName As String, inputCell As String, maxValue As Double, Optional relativeProp As String, Optional relativeShapeName As String, Optional relativeCell As String, Optional maxRelative As Double, Optional mapToX As Double, Optional mapToY As Double)
    Dim ws As Worksheet : Set ws = ThisWorkbook.ActiveSheet
    Dim parentGroup As Shape
    Dim targetShape As Shape
    Dim relativeShape As Shape
    Dim inputValue As Double
    Dim inputValuePoints As Double
    Dim inputValueNormal As Double
    Dim relativeValue As Double
    Dim relativeValuePoints As Double
    Dim relativeValueNormal As Double
    Dim mappedDimension As Double


    ' Find the parent group by name
    For Each parentGroup In ws.Shapes
        If parentGroup.Name = parentName And parentGroup.Type = msoGroup Then
            Exit For
        End If
    Next parentGroup

    ' Look for child and relative shapes
    If Not parentGroup Is Nothing Then
        For Each targetShape In parentGroup.GroupItems
            If targetShape.Name = shapeName Then

                If IsMissing(relativeProp) Then
                    ' Modify child shape only
                    If IsNumeric(ws.Range(inputCell).Value) Then
                        inputValue =  CDbl(ws.Range(inputCell).Value)
                        
                        LengthenLineShape targetShape, inputValue, 30
                    
                    End If
                Else
                    ' Modify child and relative shapes
                    For Each relativeShape In parentGroup.GroupItems
                        If relativeShape.Name = relativeShapeName Then
                            If IsNumeric(ws.Range(relativeCell).Value) Then

                                inputValue =  CDbl(ws.Range(inputCell).Value)
                                inputValueNormal = NormalizeDimensions(maxValue, inputValue, maxRelative, relativeValue)("X")
                                inputValuePoints = Application.InchesToPoints(inputValueNormal)

                                relativeValue = CDbl(ws.Range(relativeCell).Value)
                                relativeValueNormal = NormalizeDimensions(maxValue, inputValue, maxRelative, relativeValue)("Y")
                                relativeValuePoints = Application.InchesToPoints(relativeValueNormal)

                                If Not IsMissing(mapToX) Then
                                    inputValueNormal = MapDimension(inputValueNormal, 0, maxValue, 0, mapToX)
                                    relativeValueNormal = MapDimension(relativeValueNormal, 0, maxValue, 0, mapToY)
                                    
                                    inputValuePoints = Application.InchesToPoints(inputValueNormal)
                                    relativeValuePoints = Application.InchesToPoints(relativeValueNormal)
                                End If

                                If prop="Width" Then 
                                    targetShape.Width = inputValuePoints 
                                End If

                                If prop="Height" Then 
                                    targetShape.Height = inputValuePoints 
                                End If

                                If relativeProp="Width" Then 
                                    relativeShape.Width = relativeValuePoints 
                                End If

                                If relativeProp="Height" Then 
                                    relativeShape.Height = relativeValuePoints 
                                End If
                            
                            End If
                        End If
                    Next relativeShape
                End If

                Exit For
            End If
        Next targetShape
    End If
End Sub

Function NormalizeDimensions(maxX As Double, originalX As Double, maxY As Double, originalY As Double) As Collection
    Dim normalizedValues As New Collection
    
    ' Check if values exceed the maximum
    If originalX > maxX Or originalY > maxY Then
        ' Calculate the sum of original values
        Dim sumOriginal As Double
        sumOriginal = originalX + originalY
        
        ' Normalize values based on the sum and maximum values
        Dim normalizedX As Double
        Dim normalizedY As Double
        normalizedX = (originalX / sumOriginal) * maxX
        normalizedY = (originalY / sumOriginal) * maxY
        
        ' Add named elements to the collection
        normalizedValues.Add normalizedX, "X"
        normalizedValues.Add normalizedY, "Y"
    Else
        ' Values are within limits, no need to normalize
        ' Add named elements to the collection
        normalizedValues.Add originalX, "X"
        normalizedValues.Add originalY, "Y"
    End If
    
    Set NormalizeDimensions = normalizedValues
End Function

Function MapDimension(inputValue As Double, inputMin As Double, inputMax As Double, outputMin As Double, outputMax As Double) As Double
    Dim ws As Worksheet : Set ws = ThisWorkbook.ActiveSheet
    
    If inputMin = inputMax Then
        ' Avoid division by zero
        MapValue = outputMin 
    Else
        ' Perform linear mapping
        MapDimension = ((inputValue - inputMin) / (inputMax - inputMin)) * (outputMax - outputMin) + outputMin
    End If
End Function

Sub MoveRelativeShape(shapeName As String, adjacentName As String)
    Dim ws As Worksheet : Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim targetShape As Shape : Set targetShape = ws.Shapes(shapeName)
    Dim adjacentShape As Shape : Set adjacentShape = ws.Shapes(adjacentName)

    targetShape.Left = adjacentShape.Left + adjacentShape.Width
    targetShape.Top = adjacentShape.Top + adjacentShape.Height
    MsgBox "Target: " & targetShape.Top & ", Adjacent: " & adjacentShape.Top
End Sub 

Sub LengthenLineShape(targetShape As Shape, newlength As Double, angleInDegrees As Double)
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change the worksheet name
    Dim angleInRadians As Double : angleInRadians = WorksheetFunction.Radians(angleInDegrees)

    ' Get the starting point coordinates
    Dim startX As Double : startX = targetShape.Left
    Dim startY As Double : startY = targetShape.Top
    
    ' Convert length to points
    newLength = Application.InchesToPoints(newlength)

    ' Calculate width and height
    Dim newWidth As Double : newWidth = newLength * Cos(angleInRadians)
    Dim newHeight As Double : newHeight = newLength * Sin(angleInRadians)

    '  Modidy length
    targetShape.Width = newWidth
    targetShape.Height = newHeight
    targetShape.Left = startX
    targetShape.Top = startY
End Sub








