Attribute VB_Name = "TransformMatrix"

Sub ChangeDimension(shapeName As String, parentName As String, inputCell As String, maxValue As Double, Optional relativeShapeName As String, Optional relativeCell As String, Optional maxRelative As Double, Optional mapToX As Double, Optional mapToY As Double)
    ' This routine is for changing any dimension in any 
    ' drawing or group that includes any relative 
    ' dimension specified using normalization 
    ' and mapping to keep values within range

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
            ' Shift group position back to original
            ' to compensate for any side effects

            parentGroupTop = parentGroup.Top
            parentGroupLeft = parentGroup.Left
            parentGroupWidth = parentGroup.Width
            parentGroupHeight = parentGroup.Height
            Exit For
        End If
    Next parentGroup

    ' Look for child and relative shapes
    If Not parentGroup Is Nothing Then
        For Each targetShape In parentGroup.GroupItems
            If targetShape.Name = shapeName Then

                If relativeShapeName = Missing Then
                    ' Modify child shape only
                    If IsNumeric(ws.Range(inputCell).Value) Then
                        inputValue =  CDbl(ws.Range(inputCell).Value)
                        
                        ChangeLength targetShape.Name, inputValue
                        parentGroup.Width = parentGroupWidth
                        parentGroup.Height = parentGroupHeight
                    End If
                Else
                    ' Modify child and relative shapes
                    For Each relativeShape In parentGroup.GroupItems
                        If relativeShape.Name = relativeShapeName Then
                            If IsNumeric(ws.Range(relativeCell).Value) Then

                                inputValue =  CDbl(ws.Range(inputCell).Value)
                                inputValueNormal = NormalizeDimensions(maxValue, inputValue, maxRelative, relativeValue)("X")
                                
                                relativeValue = CDbl(ws.Range(relativeCell).Value)
                                relativeValueNormal = NormalizeDimensions(maxValue, inputValue, maxRelative, relativeValue)("Y")

                                If Not IsMissing(mapToX) Then
                                    inputValueNormal = MapDimension(inputValueNormal, 0, maxValue, 0, mapToX)
                                    relativeValueNormal = MapDimension(relativeValueNormal, 0, maxValue, 0, mapToY)
                                End If

                                ChangeLength targetShape.Name, inputValueNormal
                                ChangeLength relativeShape.Name, relativeValueNormal
                                parentGroup.Width = parentGroupWidth
                                parentGroup.Height = parentGroupHeight
                            
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
    ' This function is used to normalize a value
    ' based on a relative value such that exceeding
    ' dimension transfers to another and vice versa
    ' to keep both within range

    Dim normalizedValues As New Collection
    Dim sumOriginal As Double
    Dim normalizedX As Double
    Dim normalizedY As Double
    
    ' Check if values exceed the maximum
    If originalX > maxX Or originalY > maxY Then
        ' Calculate the sum of original values
        sumOriginal = originalX + originalY
        
        ' Normalize values based on the sum and maximum values
        normalizedX = (originalX / sumOriginal) * maxX
        normalizedY = (originalY / sumOriginal) * maxY
        
        ' Bring back to max after normalization
        If (normalizedX > normalizedY) Then
            Dim scaleX : scaleX = maxValue/normalizedX
            normalizedX = maxValue
            normalizedY = normalizedY * scaleX
        End If
        If (normalizedY > normalizedX) Then
            Dim scaleY : scaleY = maxRelative/normalizedY
            normalizedX = normalizedX * scaleY
            normalizedY = maxRelative
        End If
        If (normalizedY = normalizedX) Then
            MsgBox "Equal sila"
            normalizedX = maxX
            normalizedY = maxY
        End If

        ' Add named elements to the collection
        normalizedValues.Add normalizedX, "X"
        normalizedValues.Add normalizedY, "Y"
    Else
        ' Values are within limits, no need to normalize
        normalizedValues.Add originalX, "X"
        normalizedValues.Add originalY, "Y"
    End If
    
    Set NormalizeDimensions = normalizedValues
End Function

Function MapDimension(inputValue As Double, inputMin As Double, inputMax As Double, outputMin As Double, outputMax As Double) As Double
    ' This function performs linear mapping
    ' of a given value within a given range
    ' onto a specified desired range. 
    ' Specifically, mapping a calculated
    ' dimension to a desired actual 
    ' measurement for shapes
    
    Dim ws As Worksheet : Set ws = ThisWorkbook.ActiveSheet
    
    If inputMin = inputMax Then
        MapValue = outputMin  ' Avoid division by zero
    Else
        ' Perform linear mapping
        MapDimension = ((inputValue - inputMin) / (inputMax - inputMin)) * (outputMax - outputMin) + outputMin
    End If
End Function

Sub ChangeLength(shapeName As String, newlength As Double)
    ' This routine changes the length of any given
    ' line with any slope or angle while 
    ' maintaining its starting point
    
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change the worksheet name
    Dim targetShape As Shape : Set targetShape = ws.Shapes(shapeName)

    'Get the starting point coordinates
    Dim startX As Double : startX = targetShape.Left
    Dim startY As Double : startY = targetShape.Top

    'Get angle
    Dim currentAngle As Double : currentAngle = GetAngle(targetShape)
    Dim angleInRadians As Double : angleInRadians = currentAngle * (WorksheetFunction.Pi / 180)

    'Modidy length
    targetShape.Width = Application.InchesToPoints(newlength) * Cos(angleInRadians)
    targetShape.Height = Application.InchesToPoints(newlength) * Sin(angleInRadians)
End Sub

Sub ChangeAngle(shapeName As String, angleInDegrees As Double)
    ' This routine changes the angle of any given line
    ' while maintaining its length and starting point
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Dim targetShape As Shape : Set targetShape = ws.Shapes(shapeName)
    Dim angleInRadians As Double : angleInRadians = WorksheetFunction.Radians(angleInDegrees)

    'Get the starting point coordinates
    Dim startX As Double : startX = targetShape.Left
    Dim startY As Double : startY = targetShape.Top
    Dim endX As Double : endX = targetShape.Left + targetShape.Width
    Dim endY As Double : endY = targetShape.Top + targetShape.Height

    ' Calculate the current length of the line
    Dim currentLength As Double : currentLength = Sqr(((endX - startX) ^ 2) + ((endY - startY) ^ 2))

    ' Modidy length
    targetShape.Width = currentLength * Cos(angleInRadians)
    targetShape.Height = currentLength * Sin(angleInRadians)
End Sub

Function GetAngle(lineShape As Shape) As Double
    ' This function returns the angle of 
    ' any given line in Degrees
    
    Dim x1 As Double : x1 = lineShape.Left
    Dim y1 As Double : y1 = lineShape.Top
    Dim x2 As Double : x2 = lineShape.Left + lineShape.Width
    Dim y2 As Double :  y2 = lineShape.Top + lineShape.Height
    
    ' Get angle in Degrees
    GetAngle = (WorksheetFunction.Atan2(lineShape.Width, IIf(lineShape.VerticalFlip, 1, 1) * lineShape.Height)) * (180 / WorksheetFunction.Pi)
End Function

Sub MoveRelativeShape(shapeName As String, adjacentName As String, moveToEnd as Boolean)
    Dim ws As Worksheet : Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim targetShape As Shape : Set targetShape = ws.Shapes(shapeName)
    Dim adjacentShape As Shape : Set adjacentShape = ws.Shapes(adjacentName)

    If moveToEnd Then
        targetShape.Left = adjacentShape.Left + adjacentShape.Width
        targetShape.Top = adjacentShape.Top + adjacentShape.Height
    Else
        targetShape.Left = adjacentShape.Left
        targetShape.Top = adjacentShape.Top
    End If
    
End Sub 

Sub TryFunctions()
    Dim ws As Worksheet : Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim targetShape As Shape : Set targetShape = ws.Shapes("Width 1")

    ' ChangeAngle "Width 1", 30
    ' ChangeLength "Width 1", 2
    ' MsgBox "Angle: " & GetAngle(targetShape)

    Dim normal As New Collection
    Set normal = NormalizeDimensions (2, 3, 2, 5)
    MsgBox "X: " & normal("X") & vbNewLine & "Y: " & normal("Y")
End Sub





