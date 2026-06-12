Sub Main()

    ' Ensure active document is a part
    If ThisApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("Please run this in a part document.")
        Exit Sub
    End If

    Dim partDoc As PartDocument = ThisApplication.ActiveDocument
    Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition

    ' Get values
    Dim length As Double = GetLongestEdgeLength(partDef)
    Dim width As Double = GetFlatWidth(partDef)

    ' Convert from cm → inches (optional)
    Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure
    Dim lengthIn As Double = uom.ConvertUnits(length, "cm", "in")
    Dim widthIn As Double = uom.ConvertUnits(width, "cm", "in")

    ' Output
    MessageBox.Show( _
        "Flat Length: " & Round(lengthIn, 3) & " in" & vbCrLf & _
        "Flat Width: " & Round(widthIn, 3) & " in")

End Sub


'========================
' LONGEST EDGE
'========================
Function GetLongestEdgeLength(partDef As PartComponentDefinition) As Double

    Dim maxLen As Double = 0.0

    For Each body As SurfaceBody In partDef.SurfaceBodies
        For Each edge As Edge In body.Edges
            
            ' FIX: use "edge", not "Edge"
            Dim eval As CurveEvaluator = Edge.Evaluator
            
            Dim minParam As Double
            Dim maxParam As Double
            eval.GetParamExtents(minParam, maxParam)
            
            Dim length As Double
            eval.GetLengthAtParam(minParam, maxParam, length)
            
            If length > maxLen Then
                maxLen = length
            End If
            
        Next
    Next

    Return maxLen

End Function


'========================
' FLAT WIDTH
'========================
Function GetFlatWidth(partDef As PartComponentDefinition) As Double

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

    ' --- Step 1: Get longest edge (tube axis) ---
    Dim longestEdge As Edge = Nothing
    Dim maxLen As Double = 0

    For Each body As SurfaceBody In partDef.SurfaceBodies
        For Each edge As Edge In body.Edges
            
            ' FIX: use "edge", not "Edge"
            Dim eval As CurveEvaluator = Edge.Evaluator
            
            Dim minP As Double
            Dim maxP As Double
            eval.GetParamExtents(minP, maxP)
            
            Dim len As Double
            eval.GetLengthAtParam(minP, maxP, len)
            
            If len > maxLen Then
                maxLen = len
                longestEdge = Edge   ' FIX: use edge
            End If

        Next
    Next

    If longestEdge Is Nothing Then Return 0

    ' --- Step 2: Get axis direction ---
    Dim startArr(2) As Double
	Dim endArr(2) As Double
	
	longestEdge.Evaluator.GetEndPoints(startArr, endArr)
	
	Dim startPt As Point = tg.CreatePoint(startArr(0), startArr(1), startArr(2))
	Dim endPt As Point = tg.CreatePoint(endArr(0), endArr(1), endArr(2))
    
    Dim axisDir As Vector = startPt.VectorTo(endPt)
    axisDir.Normalize()

    ' --- Step 3: Build perpendicular coordinate system ---
    Dim up As Vector = tg.CreateVector(0, 0, 1)

    ' Handle parallel case (IMPORTANT FIX)
    If Math.Abs(axisDir.DotProduct(up)) > 0.999 Then
        up = tg.CreateVector(1, 0, 0)
    End If

    Dim xDir As Vector = axisDir.CrossProduct(up)
    xDir.Normalize()

    Dim yDir As Vector = axisDir.CrossProduct(xDir)
    yDir.Normalize()

    ' --- Step 4: Project vertices into plane ---
    Dim refPt As Point = startPt

    Dim minX As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue
    Dim minY As Double = Double.MaxValue
    Dim maxY As Double = Double.MinValue

    For Each body As SurfaceBody In partDef.SurfaceBodies
        For Each v As Vertex In body.Vertices
            
            Dim vec As Vector = refPt.VectorTo(v.Point)
            
            Dim x As Double = vec.DotProduct(xDir)
            Dim y As Double = vec.DotProduct(yDir)
            
            If x < minX Then minX = x
            If x > maxX Then maxX = x
            If y < minY Then minY = y
            If y > maxY Then maxY = y

        Next
    Next

    Dim width As Double = maxX - minX
    Dim height As Double = maxY - minY

    ' --- Step 5: Flat pattern width ---
    Dim flatWidth As Double = 2 * (width + height)

    Return flatWidth

End Function

Function IsClosedSection(partDef As PartComponentDefinition) As Boolean

    For Each body As SurfaceBody In partDef.SurfaceBodies
        For Each f As Face In body.Faces
            
            If f.EdgeLoops.Count = 1 Then
                return False
            ElseIf f.EdgeLoops.Count = 2 Then
                return True
            Else
                return False
            End If
        '     For Each edgeLoop As EdgeLoop In f.EdgeLoops
                
        '         ' Inner loop found → hollow section
        '         If Not EdgeLoop.IsOuterEdgeLoop Then
        '             Return True
        '         End If
                
        '     Next
            
        Next
    Next

    ' No inner loops anywhere → open section
    Return False

End Function
