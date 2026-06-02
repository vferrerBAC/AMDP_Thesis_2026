'Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Math

Sub Main()

    Dim invApp As Inventor.Application = ThisApplication
    Dim asmDoc As AssemblyDocument = invApp.ActiveDocument

    If asmDoc Is Nothing Then
        MsgBox("Please open an Assembly document", MsgBoxStyle.Exclamation)
        Exit Sub
    End If

    Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
	
	Dim highlightSet1 As HighlightSet = asmDoc.CreateHighlightSet()
	Dim highlightSet2 As HighlightSet = asmDoc.CreateHighlightSet()
	
	' Colors (RGB)
	highlightSet1.Color = invApp.TransientObjects.CreateColor(255, 0, 0)   ' Red
	highlightSet2.Color = invApp.TransientObjects.CreateColor(0, 255, 0)   ' Green


    ' ==========================================================
    ' ✅ Excel setup
    ' ==========================================================
    Dim excelApp As Object = CreateObject("Excel.Application")
    excelApp.Visible = True

    Dim workbook As Object = excelApp.Workbooks.Add

    Dim sheet As Object = workbook.Sheets(1)
    sheet.Name = "Joints"

    Dim debugSheet As Object = workbook.Sheets.Add
    debugSheet.Name = "DebugLog"

    ' ==========================================================
    ' ✅ Headers (main sheet)
    ' ==========================================================
    sheet.Cells(1, 1).Value = "Joint ID"
    sheet.Cells(1, 2).Value = "X"
    sheet.Cells(1, 3).Value = "Y"
    sheet.Cells(1, 4).Value = "Z"
    sheet.Cells(1, 5).Value = "Part 1"
    sheet.Cells(1, 6).Value = "Material 1"
    sheet.Cells(1, 7).Value = "Part 2"
    sheet.Cells(1, 8).Value = "Material 2"

    ' ==========================================================
    ' ✅ Debug headers
    ' ==========================================================
    debugSheet.Cells(1, 1).Value = "Part 1"
    debugSheet.Cells(1, 2).Value = "Part 2"
    debugSheet.Cells(1, 3).Value = "Distance"
    debugSheet.Cells(1, 4).Value = "Angle"
    debugSheet.Cells(1, 5).Value = "Normal Check"
	debugSheet.Cells(1, 6).Value = "Overlap Check"
    debugSheet.Cells(1, 7).Value = "Result"

    Dim row As Integer = 2
    Dim debugRow As Integer = 2
    Dim jointID As Integer = 1

    ' ==========================================================
    ' ✅ Parameters
    ' ==========================================================
    Dim contactTol As Double = 0.0001

    Dim distances As New List(Of Double)

    ' ==========================================================
    ' ✅ Loop through occurrence pairs
    ' ==========================================================
    For i = 1 To compDef.Occurrences.Count - 1

        Dim occ1 As ComponentOccurrence = compDef.Occurrences.Item(i)

        For j = i + 1 To compDef.Occurrences.Count

            Dim occ2 As ComponentOccurrence = compDef.Occurrences.Item(j)

            If occ1.SurfaceBodies.Count = 0 Or occ2.SurfaceBodies.Count = 0 Then Continue For

            Dim body1 As SurfaceBody = occ1.SurfaceBodies.Item(1)
            Dim body2 As SurfaceBody = occ2.SurfaceBodies.Item(1)

            Dim part1Name As String = GetPartName(occ1)
            Dim part2Name As String = GetPartName(occ2)

            For Each face1 As Face In body1.Faces

                For Each face2 As Face In body2.Faces

                    ' ==========================================================
                    ' ✅ Run checks
                    ' ==========================================================
                    Dim minDist As Double = GetMinDistance_API(face1, face2)
                    Dim minAngle As Double = GetAngle(face1, face2)

                    Dim distPass As Boolean = (minDist < contactTol)
                    Dim anglePass As Boolean = (minAngle < contactTol)

                    Dim normalPass As Boolean = False
					Dim overlapPass As Boolean = False

                    If distPass And anglePass Then
'                        Dim faceProxy1 As FaceProxy
'						occ1.CreateGeometryProxy(face1, faceProxy1)
'						Dim faceProxy2 As FaceProxy
'						occ2.CreateGeometryProxy(face2, faceProxy2)
						
'						highlightSet1.AddItem(faceProxy1)
'						highlightSet2.AddItem(faceProxy2)
						
						normalPass = CheckNormal(face1, occ1, face2, occ2)
						overlapPass = DoFacesOverlap(face1, occ1, face2, occ2)
					
'						MsgBox("Coincident: " & normalPass)
'						MsgBox("Press OK to view next contact pair")
						
'						highlightSet1.Clear()
'			            highlightSet2.Clear()
						
						
                    End If

                    ' ==========================================================
                    ' ✅ Determine failure reason
                    ' ==========================================================
                    Dim result As String

                    If Not distPass Then
                        result = "FAIL: Distance"
                    ElseIf Not anglePass Then
                        result = "FAIL: Angle"
                    ElseIf Not normalPass Then
                        result = "FAIL: Normal"
					ElseIf Not overlapPass Then
                        result = "FAIL: Overlap"
                    Else
                        result = "PASS"
                    End If

                    ' ==========================================================
                    ' ✅ Write to debug sheet
                    ' ==========================================================
                    debugSheet.Cells(debugRow, 1).Value = part1Name
                    debugSheet.Cells(debugRow, 2).Value = part2Name
                    debugSheet.Cells(debugRow, 3).Value = Round(minDist, 4)
                    debugSheet.Cells(debugRow, 4).Value = Round(minAngle, 4)
                    debugSheet.Cells(debugRow, 5).Value = normalPass
					debugSheet.Cells(debugRow, 6).Value = overlapPass
                    debugSheet.Cells(debugRow, 7).Value = result

                    debugRow += 1

                    ' ==========================================================
                    ' ✅ If valid joint → process
                    ' ==========================================================
                    If result <> "PASS" Then Continue For

                    distances.Add(minDist)

                    Dim centroid As Point = GetContactCentroid(face1, occ1, face2, occ2)

                    ' Create work point
                    Dim workPt As WorkPoint = compDef.WorkPoints.AddFixed(centroid)
                    workPt.Name = "Joint_" & jointID

                    Dim mat1 As String = GetMaterial(occ1)
                    Dim mat2 As String = GetMaterial(occ2)

                    ' Write to main sheet
                    sheet.Cells(row, 1).Value = jointID
                    sheet.Cells(row, 2).Value = Round(centroid.X, 4)
                    sheet.Cells(row, 3).Value = Round(centroid.Y, 4)
                    sheet.Cells(row, 4).Value = Round(centroid.Z, 4)
                    sheet.Cells(row, 5).Value = part1Name
                    sheet.Cells(row, 6).Value = mat1
                    sheet.Cells(row, 7).Value = part2Name
                    sheet.Cells(row, 8).Value = mat2

                    row += 1
                    jointID += 1
					
'					Dim faceProxy1 As FaceProxy
'					occ1.CreateGeometryProxy(face1, faceProxy1)
'					Dim faceProxy2 As FaceProxy
'					occ2.CreateGeometryProxy(face2, faceProxy2)
					
'					highlightSet1.AddItem(faceProxy1)
'					highlightSet2.AddItem(faceProxy2)
					
'					normalPass = CheckNormal(face1, occ1, face2, occ2)
'					overlapPass = DoFacesOverlap(face1, occ1, face2, occ2)
				
'					MsgBox("Coincident: " & normalPass)
'					MsgBox("Press OK to view next contact pair")
					
'					highlightSet1.Clear()
'		            highlightSet2.Clear()

                Next
            Next
        Next
    Next

    ' ==========================================================
    ' ✅ Formatting
    ' ==========================================================
    sheet.Columns.AutoFit
    debugSheet.Columns.AutoFit

    MsgBox("There are " & distances.Count & " connections", MsgBoxStyle.Information)

End Sub



' ==========================================================
' ✅ Distance
' ==========================================================
Function GetMinDistance_API(face1 As Face, face2 As Face) As Double
    Try
'		Dim factor As Double = 10 ^ 3
        Return ThisApplication.MeasureTools.GetMinimumDistance(face1, face2)
    Catch
        Return 999999
    End Try
End Function

' ==========================================================
' ✅ Angle
' ==========================================================
Function GetAngle(face1 As Face, face2 As Face) As Double
    Try
        Return ThisApplication.MeasureTools.GetAngle(face1, face2)
    Catch
        Return -1
    End Try
End Function

' ==========================================================
' ✅ Part Name
' ==========================================================
Function GetPartName(occ As ComponentOccurrence) As String
    Return occ.Definition.Document.DisplayName
End Function

' ==========================================================
' ✅ Material
' ==========================================================
Function GetMaterial(occ As ComponentOccurrence) As String
    Try
        If TypeOf occ.Definition.Document Is PartDocument Then
            Dim partDoc As PartDocument = occ.Definition.Document
            Return partDoc.ComponentDefinition.Material.Name
        Else
            Return "Assembly"
        End If
    Catch
        Return "N/A"
    End Try
End Function

' ==========================================================
' ✅ Normal Check (same as yours, unchanged)
' ==========================================================
' (keep your existing CheckNormal function here)
Function CheckNormal(face1 As Face, occ1 As ComponentOccurrence, _
                     face2 As Face, occ2 As ComponentOccurrence) As Boolean

    Dim angTol As Double = 1E-3
    Dim tg As TransientGeometry = ThisApplication.TransientGeometry
    Dim mt As MeasureTools = ThisApplication.MeasureTools

    ' --- Create proxies ---
    Dim proxy1 As FaceProxy
    occ1.CreateGeometryProxy(face1, proxy1)

    Dim proxy2 As FaceProxy
    occ2.CreateGeometryProxy(face2, proxy2)

    ' --- Only planar faces ---
    If proxy1.SurfaceType <> kPlaneSurface Or proxy2.SurfaceType <> kPlaneSurface Then
        Return False
    End If

    ' --- Evaluators ---
    Dim eval1 As SurfaceEvaluator = proxy1.Evaluator
    Dim eval2 As SurfaceEvaluator = proxy2.Evaluator

    ' --- Get midpoint in UV space ---
    Dim center1(1) As Double
    center1(0) = (eval1.ParamRangeRect.MinPoint.X + eval1.ParamRangeRect.MaxPoint.X) / 2
    center1(1) = (eval1.ParamRangeRect.MinPoint.Y + eval1.ParamRangeRect.MaxPoint.Y) / 2

    Dim center2(1) As Double
    center2(0) = (eval2.ParamRangeRect.MinPoint.X + eval2.ParamRangeRect.MaxPoint.X) / 2
    center2(1) = (eval2.ParamRangeRect.MinPoint.Y + eval2.ParamRangeRect.MaxPoint.Y) / 2

    ' --- Get normals ---
    Dim n1(2) As Double
    Dim n2(2) As Double

    eval1.GetNormal(center1, n1)
    eval2.GetNormal(center2, n2)

    Dim v1 As Vector = tg.CreateVector(n1(0), n1(1), n1(2))
    Dim v2 As Vector = tg.CreateVector(n2(0), n2(1), n2(2))

    v1.Normalize
    v2.Normalize

    ' --- Apply parameter reversal ---
    If proxy1.IsParamReversed Then v1.ScaleBy(-1)
    If proxy2.IsParamReversed Then v2.ScaleBy(-1)

    ' ==========================================================
    ' ✅ METHOD 2: DISTANCE-BASED NORMAL ORIENTATION
    ' ==========================================================

    Dim eps As Double = 0.01   ' small offset (tune if needed)

    ' --- Get point on each face ---
    Dim pt1Arr(2) As Double
    Dim pt2Arr(2) As Double

    eval1.GetPointAtParam(center1, pt1Arr)
    eval2.GetPointAtParam(center2, pt2Arr)

    Dim p1 As Point = tg.CreatePoint(pt1Arr(0), pt1Arr(1), pt1Arr(2))
    Dim p2 As Point = tg.CreatePoint(pt2Arr(0), pt2Arr(1), pt2Arr(2))
	
	
'	MsgBox("V1=(" & v1.X.ToString("0.000") & "," & v1.Y.ToString("0.000") & "," & v1.Z.ToString("0.000") & ") | V2=(" & v2.X.ToString("0.000") & "," & v2.Y.ToString("0.000") & "," & v2.Z.ToString("0.000") & ")")


    ' ==========================================================
    ' ✅ Evaluate direction for Face 1
    ' ==========================================================

    Dim p1Forward As Point = tg.CreatePoint( _
        p1.X + v1.X * eps, _
        p1.Y + v1.Y * eps, _
        p1.Z + v1.Z * eps)

    Dim p1Backward As Point = tg.CreatePoint( _
        p1.X - v1.X * eps, _
        p1.Y - v1.Y * eps, _
        p1.Z - v1.Z * eps)

    Dim dFwd1 As Double = mt.GetMinimumDistance(p1Forward, occ1)
    Dim dBack1 As Double = mt.GetMinimumDistance(p1Backward, occ1)

    ' If forward is closer → inside → flip
    If dFwd1 < dBack1 Then
        v1.ScaleBy(-1)
'		MsgBox("V1 flipped")
    End If

    ' ==========================================================
    ' ✅ Evaluate direction for Face 2
    ' ==========================================================

    Dim p2Forward As Point = tg.CreatePoint( _
        p2.X + v2.X * eps, _
        p2.Y + v2.Y * eps, _
        p2.Z + v2.Z * eps)

    Dim p2Backward As Point = tg.CreatePoint( _
        p2.X - v2.X * eps, _
        p2.Y - v2.Y * eps, _
        p2.Z - v2.Z * eps)

    Dim dFwd2 As Double = mt.GetMinimumDistance(p2Forward, occ2)
    Dim dBack2 As Double = mt.GetMinimumDistance(p2Backward, occ2)

    If dFwd2 < dBack2 Then
        v2.ScaleBy(-1)
'		MsgBox("V2 flipped")
    End If
	
	
'	MsgBox("V1=(" & v1.X.ToString("0.000") & "," & v1.Y.ToString("0.000") & "," & v1.Z.ToString("0.000") & ") | V2=(" & v2.X.ToString("0.000") & "," & v2.Y.ToString("0.000") & "," & v2.Z.ToString("0.000") & ")")


    ' ==========================================================
    ' ✅ FINAL CHECK: Must already be opposing (NO artificial flipping)
    ' ==========================================================

    Dim dotRaw As Double = v1.DotProduct(v2)

    Return dotRaw < (-1.0 + angTol)

End Function


Function GetContactCentroid(face1 As Face, occ1 As ComponentOccurrence, _
                             face2 As Face, occ2 As ComponentOccurrence) As Point

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

    ' --- Create proxies (world space) ---
    Dim proxy1 As FaceProxy
    occ1.CreateGeometryProxy(face1, proxy1)

    Dim proxy2 As FaceProxy
    occ2.CreateGeometryProxy(face2, proxy2)

    ' ==========================================================
    ' ✅ STEP 1: Build a local 2D coordinate frame on the plane
    ' Use face1's normal as the plane normal
    ' ==========================================================
    Dim eval1 As SurfaceEvaluator = proxy1.Evaluator
    Dim eval2 As SurfaceEvaluator = proxy2.Evaluator

    ' Get midpoint UV of face1
    Dim centerUV(1) As Double
    centerUV(0) = (eval1.ParamRangeRect.MinPoint.X + eval1.ParamRangeRect.MaxPoint.X) / 2
    centerUV(1) = (eval1.ParamRangeRect.MinPoint.Y + eval1.ParamRangeRect.MaxPoint.Y) / 2

    ' Get origin point on the plane
    Dim originArr(2) As Double
    eval1.GetPointAtParam(centerUV, originArr)
    Dim origin As Point = tg.CreatePoint(originArr(0), originArr(1), originArr(2))

    ' Get normal (plane Z-axis)
    Dim nArr(2) As Double
    eval1.GetNormal(centerUV, nArr)
    Dim planeN As Vector = tg.CreateVector(nArr(0), nArr(1), nArr(2))
    planeN.Normalize()

    ' Build plane X-axis: pick a world axis not parallel to normal
    Dim worldX As Vector = tg.CreateVector(1, 0, 0)
    If Math.Abs(planeN.DotProduct(worldX)) > 0.9 Then
        worldX = tg.CreateVector(0, 1, 0)
    End If

    ' Gram-Schmidt: make planeX perpendicular to planeN
    Dim planeX As Vector = worldX.Copy()
    Dim proj As Vector = planeN.Copy()
    proj.ScaleBy(planeN.DotProduct(planeX))
    planeX.SubtractVector(proj)
    planeX.Normalize()

    ' planeY = planeN × planeX
    Dim planeY As Vector = planeN.CrossProduct(planeX)
    planeY.Normalize()

    ' ==========================================================
    ' ✅ STEP 2: Sample face tessellation points and project to 2D
    ' ==========================================================
    Dim pts1 As New List(Of Double())   ' list of (u, v) pairs
    Dim pts2 As New List(Of Double())

    SampleFacePoints(proxy1, origin, planeX, planeY, pts1)
    SampleFacePoints(proxy2, origin, planeX, planeY, pts2)

    ' ==========================================================
    ' ✅ STEP 3: Get 2D bounding boxes
    ' ==========================================================
    Dim min1U As Double, max1U As Double, min1V As Double, max1V As Double
    Dim min2U As Double, max2U As Double, min2V As Double, max2V As Double

    GetBounds2D(pts1, min1U, max1U, min1V, max1V)
    GetBounds2D(pts2, min2U, max2U, min2V, max2V)

    ' ==========================================================
    ' ✅ STEP 4: Compute 2D overlap
    ' ==========================================================
    Dim overlapMinU As Double = Math.Max(min1U, min2U)
    Dim overlapMaxU As Double = Math.Min(max1U, max2U)
    Dim overlapMinV As Double = Math.Max(min1V, min2V)
    Dim overlapMaxV As Double = Math.Min(max1V, max2V)

    Dim centerU As Double
    Dim centerV As Double

    If overlapMaxU >= overlapMinU And overlapMaxV >= overlapMinV Then
        ' Valid overlap → use overlap centroid
        centerU = (overlapMinU + overlapMaxU) / 2
        centerV = (overlapMinV + overlapMaxV) / 2
    Else
        ' Fallback: average of both face centroids in 2D
        centerU = (min1U + max1U + min2U + max2U) / 4
        centerV = (min1V + max1V + min2V + max2V) / 4
    End If

    ' ==========================================================
    ' ✅ STEP 5: Unproject 2D centroid back to 3D world space
    ' ==========================================================
    Dim wx As Double = origin.X + centerU * planeX.X + centerV * planeY.X
    Dim wy As Double = origin.Y + centerU * planeX.Y + centerV * planeY.Y
    Dim wz As Double = origin.Z + centerU * planeX.Z + centerV * planeY.Z

    Return tg.CreatePoint(wx, wy, wz)

End Function


' ==========================================================
' ✅ Helper: sample 3D face points and project into 2D plane coords
' ==========================================================
Sub SampleFacePoints(proxy As FaceProxy, origin As Point, _
                     planeX As Vector, planeY As Vector, _
                     pts As List(Of Double()))

    Dim eval As SurfaceEvaluator = proxy.Evaluator
    Dim rect As Box2d = eval.ParamRangeRect

    Dim uMin As Double = rect.MinPoint.X
    Dim uMax As Double = rect.MaxPoint.X
    Dim vMin As Double = rect.MinPoint.Y
    Dim vMax As Double = rect.MaxPoint.Y

    Dim steps As Integer = 5   ' 5×5 grid = 25 samples; increase for curved faces

    For iu As Integer = 0 To steps
        For iv As Integer = 0 To steps

            Dim u As Double = uMin + (uMax - uMin) * iu / steps
            Dim v As Double = vMin + (vMax - vMin) * iv / steps

            Dim uv(1) As Double
            uv(0) = u
            uv(1) = v

            Try
                Dim ptArr(2) As Double
                eval.GetPointAtParam(uv, ptArr)

                ' Project onto plane axes
                Dim dx As Double = ptArr(0) - origin.X
                Dim dy As Double = ptArr(1) - origin.Y
                Dim dz As Double = ptArr(2) - origin.Z

                Dim pu As Double = dx * planeX.X + dy * planeX.Y + dz * planeX.Z
                Dim pv As Double = dx * planeY.X + dy * planeY.Y + dz * planeY.Z

                pts.Add(New Double() {pu, pv})
            Catch
                ' skip invalid params (outside face boundary)
            End Try

        Next
    Next

End Sub


' ==========================================================
' ✅ Helper: get 2D bounding box from projected point list
' ==========================================================
Sub GetBounds2D(pts As List(Of Double()), _
                ByRef minU As Double, ByRef maxU As Double, _
                ByRef minV As Double, ByRef maxV As Double)

    minU = Double.MaxValue
    maxU = Double.MinValue
    minV = Double.MaxValue
    maxV = Double.MinValue

    For Each pt As Double() In pts
        If pt(0) < minU Then minU = pt(0)
        If pt(0) > maxU Then maxU = pt(0)
        If pt(1) < minV Then minV = pt(1)
        If pt(1) > maxV Then maxV = pt(1)
    Next

End Sub

Function DoFacesOverlap(face1 As Face, occ1 As ComponentOccurrence, _
                         face2 As Face, occ2 As ComponentOccurrence) As Boolean

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

    Dim proxy1 As FaceProxy
    occ1.CreateGeometryProxy(face1, proxy1)

    Dim proxy2 As FaceProxy
    occ2.CreateGeometryProxy(face2, proxy2)

    Dim box1 As Box = proxy1.Evaluator.RangeBox
    Dim box2 As Box = proxy2.Evaluator.RangeBox

    ' --- Check overlap on all 3 axes ---
    Dim overlapX As Boolean = (box1.MinPoint.X <= box2.MaxPoint.X) And (box1.MaxPoint.X >= box2.MinPoint.X)
    Dim overlapY As Boolean = (box1.MinPoint.Y <= box2.MaxPoint.Y) And (box1.MaxPoint.Y >= box2.MinPoint.Y)
    Dim overlapZ As Boolean = (box1.MinPoint.Z <= box2.MaxPoint.Z) And (box1.MaxPoint.Z >= box2.MinPoint.Z)

    ' --- For planar faces, one axis will be nearly equal (the normal axis) ---
    ' --- So we only require overlap on the 2 "in-plane" axes ---
    ' --- Determine which axis is the "thin" axis (normal direction) ---
    Dim spanX1 As Double = Math.Abs(box1.MaxPoint.X - box1.MinPoint.X)
    Dim spanY1 As Double = Math.Abs(box1.MaxPoint.Y - box1.MinPoint.Y)
    Dim spanZ1 As Double = Math.Abs(box1.MaxPoint.Z - box1.MinPoint.Z)

    Dim minSpan As Double = Math.Min(spanX1, Math.Min(spanY1, spanZ1))

    Dim flatTol As Double = 0.001  ' tune to your model scale

    If Math.Abs(spanX1 - minSpan) < flatTol Then
        ' X is normal axis → check Y and Z only
        Return overlapY And overlapZ
    ElseIf Math.Abs(spanY1 - minSpan) < flatTol Then
        ' Y is normal axis → check X and Z only
        Return overlapX And overlapZ
    Else
        ' Z is normal axis → check X and Y only
        Return overlapX And overlapY
    End If

End Function