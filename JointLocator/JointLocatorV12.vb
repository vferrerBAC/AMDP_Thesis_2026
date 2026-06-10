Imports Inventor
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
    excelApp.Visible = False   ' Keep hidden until done (faster)

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
    sheet.Cells(1, 9).Value = "Joint Length (in.)"

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

    ' ==========================================================
    ' ✅ Parameters
    ' ==========================================================
    Dim contactTol As Double = 0.254   ' 0.254 cm = 0.1 in → adjust as needed for tolerance

    ' ==========================================================
    ' ✅ Result accumulators (batch Excel writes at end)
    ' ==========================================================
    Dim jointRows As New List(Of Object())
    Dim debugRows As New List(Of Object())
    Dim jointID As Integer = 1

    ' ==========================================================
    ' ✅ Loop through occurrence pairs
    ' ==========================================================
    For i = 1 To compDef.Occurrences.Count - 1

        Dim occ1 As ComponentOccurrence = compDef.Occurrences.Item(i)

        For j = i + 1 To compDef.Occurrences.Count

            Dim occ2 As ComponentOccurrence = compDef.Occurrences.Item(j)

            ' ----------------------------------------------------------
            ' ✅ IMPROVEMENT 1: Occurrence-level bounding box pre-filter
            ' Skips entire part pairs that can't possibly be touching
            ' ----------------------------------------------------------
            If Not OccurrencesCouldTouch(occ1, occ2, contactTol) Then Continue For

            If occ1.SurfaceBodies.Count = 0 Or occ2.SurfaceBodies.Count = 0 Then Continue For

            Dim body1 As SurfaceBody = occ1.SurfaceBodies.Item(1)
            Dim body2 As SurfaceBody = occ2.SurfaceBodies.Item(1)

            Dim part1Name As String = occ1.Name
            Dim part2Name As String = occ2.Name

            For Each face1 As Face In body1.Faces

                For Each face2 As Face In body2.Faces

                    ' ----------------------------------------------------------
                    ' ✅ IMPROVEMENT 2: Face-level bounding box pre-filter
                    ' Skips face pairs whose boxes don't overlap — very cheap
                    ' ----------------------------------------------------------
                    If Not FaceBoxesCouldTouch(face1, occ1, face2, occ2, contactTol) Then Continue For

                    ' ----------------------------------------------------------
                    ' ✅ IMPROVEMENT 3: Angle first (cheaper), distance second
                    ' Each check exits early before running the next
                    ' ----------------------------------------------------------
                    Dim minAngle As Double = GetAngle(face1, face2)
                    Dim minDist As Double = 999999
                    Dim normalPass As Boolean = False
                    Dim overlapPass As Boolean = False

                    Dim result As String

                    If minAngle >= contactTol Then
                        result = "FAIL: Angle"
                    Else
                        minDist = GetMinDistance_API(face1, face2)

                        If minDist >= contactTol Then
                            result = "FAIL: Distance"
                        Else
                            ' ----------------------------------------------------------
                            ' ✅ IMPROVEMENT 4: Create proxies once, reuse across all
                            ' functions (CheckNormal, DoFacesOverlap, GetContactCentroid)
                            ' ----------------------------------------------------------
                            Dim proxy1 As FaceProxy
                            occ1.CreateGeometryProxy(face1, proxy1)
                            Dim proxy2 As FaceProxy
                            occ2.CreateGeometryProxy(face2, proxy2)

                            normalPass = CheckNormal(proxy1, occ1, proxy2, occ2)

                            If Not normalPass Then
                                result = "FAIL: Normal"
                            Else
                                overlapPass = DoFacesOverlap(proxy1, proxy2)

                                If Not overlapPass Then
                                    result = "FAIL: No Overlap"
                                Else
                                    result = "PASS"
                                End If
                            End If

                            ' Highlight passing or failing normal pairs (optional debug)
                            If result = "PASS" Then
                                highlightSet1.AddItem(proxy1)
                                highlightSet2.AddItem(proxy2)
                                highlightSet1.Clear()
                                highlightSet2.Clear()
                            End If
                        End If
                    End If

                    ' ----------------------------------------------------------
                    ' ✅ Accumulate debug row
                    ' ----------------------------------------------------------
                    debugRows.Add(New Object() { _
                        part1Name, part2Name, _
                        Round(minDist, 4), Round(minAngle, 4), _
                        normalPass, overlapPass, result})

                    ' ----------------------------------------------------------
                    ' ✅ If valid joint → compute centroid and accumulate
                    ' ----------------------------------------------------------
                    If result <> "PASS" Then Continue For

                    ' Proxies already created above — pass directly
                    Dim proxy1Final As FaceProxy
                    occ1.CreateGeometryProxy(face1, proxy1Final)
                    Dim proxy2Final As FaceProxy
                    occ2.CreateGeometryProxy(face2, proxy2Final)

                    Dim centroid As Point = GetContactCentroid(face1, face2)
                    Dim longestDim As Double = GetJointLength(face1, face2) / 2.54   ' Convert from cm to inches

                    ' Create work point
                    Dim workPt As WorkPoint = compDef.WorkPoints.AddFixed(centroid)
                    workPt.Name = "Joint_" & jointID

                    Dim mat1 As String = GetMaterial(occ1)
                    Dim mat2 As String = GetMaterial(occ2)

                    jointRows.Add(New Object() { _
                        jointID, _
                        Round(centroid.X, 4), Round(centroid.Y, 4), Round(centroid.Z, 4), _
                        part1Name, mat1, part2Name, mat2, Round(longestDim, 4)})

                    jointID += 1

                Next
            Next
        Next
    Next

    ' ==========================================================
    ' ✅ IMPROVEMENT 5: Batch write all rows to Excel at once
    ' ==========================================================
    For r As Integer = 0 To jointRows.Count - 1
        Dim data As Object() = jointRows(r)
        For c As Integer = 0 To data.Length - 1
            sheet.Cells(r + 2, c + 1).Value = data(c)
        Next
    Next

    For r As Integer = 0 To debugRows.Count - 1
        Dim data As Object() = debugRows(r)
        For c As Integer = 0 To data.Length - 1
            debugSheet.Cells(r + 2, c + 1).Value = data(c)
        Next
    Next

    ' ==========================================================
    ' ✅ Formatting
    ' ==========================================================
    sheet.Columns.AutoFit
    debugSheet.Columns.AutoFit

    ' Save file to specified location and close Excel
    Dim saveFolder As String = "C:\Users\SRosario\OneDrive - BAC\Desktop\Joint Catalog\JointIdentifierOutputs"
    Dim saveFileName As String = "Joints_" & Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
    Dim savePath As String = System.IO.Path.Combine(saveFolder, saveFileName)
    workbook.SaveAs(savePath)

    ' Cleanup and close Excel
    workbook.Close(False)
    excelApp.Quit()
    workbook = Nothing
    excelApp = Nothing

    MsgBox("There are " & (jointID - 1) & " connections" & vbCrLf & "Saved: " & savePath, MsgBoxStyle.Information)

End Sub


' ==========================================================
' ✅ IMPROVEMENT 1: Occurrence-level bounding box pre-filter
' ==========================================================
Function OccurrencesCouldTouch(occ1 As ComponentOccurrence, _
                                occ2 As ComponentOccurrence, _
                                tol As Double) As Boolean

    Dim box1 As Box = occ1.RangeBox
    Dim box2 As Box = occ2.RangeBox

    If box1.MaxPoint.X + tol < box2.MinPoint.X Then Return False
    If box2.MaxPoint.X + tol < box1.MinPoint.X Then Return False
    If box1.MaxPoint.Y + tol < box2.MinPoint.Y Then Return False
    If box2.MaxPoint.Y + tol < box1.MinPoint.Y Then Return False
    If box1.MaxPoint.Z + tol < box2.MinPoint.Z Then Return False
    If box2.MaxPoint.Z + tol < box1.MinPoint.Z Then Return False

    Return True

End Function


' ==========================================================
' ✅ IMPROVEMENT 2: Face-level bounding box pre-filter
' ==========================================================
Function FaceBoxesCouldTouch(face1 As Face, occ1 As ComponentOccurrence, _
                              face2 As Face, occ2 As ComponentOccurrence, _
                              tol As Double) As Boolean

    Dim proxy1 As FaceProxy
    occ1.CreateGeometryProxy(face1, proxy1)
    Dim proxy2 As FaceProxy
    occ2.CreateGeometryProxy(face2, proxy2)

    Dim box1 As Box = proxy1.Evaluator.RangeBox
    Dim box2 As Box = proxy2.Evaluator.RangeBox

    If box1.MaxPoint.X + tol < box2.MinPoint.X Then Return False
    If box2.MaxPoint.X + tol < box1.MinPoint.X Then Return False
    If box1.MaxPoint.Y + tol < box2.MinPoint.Y Then Return False
    If box2.MaxPoint.Y + tol < box1.MinPoint.Y Then Return False
    If box1.MaxPoint.Z + tol < box2.MinPoint.Z Then Return False
    If box2.MaxPoint.Z + tol < box1.MinPoint.Z Then Return False

    Return True

End Function


' ==========================================================
' ✅ Distance
' ==========================================================
Function GetMinDistance_API(face1 As Face, face2 As Face) As Double
    Try
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
' ✅ IMPROVEMENT 4: Normal Check — accepts pre-built proxies
' ==========================================================
Function CheckNormal(proxy1 As FaceProxy, occ1 As ComponentOccurrence, _
                     proxy2 As FaceProxy, occ2 As ComponentOccurrence) As Boolean

    Dim angTol As Double = 1E-3
    Dim tg As TransientGeometry = ThisApplication.TransientGeometry
    Dim mt As MeasureTools = ThisApplication.MeasureTools

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

    v1.Normalize()
    v2.Normalize()

    ' --- Apply parameter reversal ---
    If proxy1.IsParamReversed Then v1.ScaleBy(-1)
    If proxy2.IsParamReversed Then v2.ScaleBy(-1)

    ' --- Get points on each face ---
    Dim eps As Double = 0.01

    Dim pt1Arr(2) As Double
    Dim pt2Arr(2) As Double

    eval1.GetPointAtParam(center1, pt1Arr)
    eval2.GetPointAtParam(center2, pt2Arr)

    Dim p1 As Point = tg.CreatePoint(pt1Arr(0), pt1Arr(1), pt1Arr(2))
    Dim p2 As Point = tg.CreatePoint(pt2Arr(0), pt2Arr(1), pt2Arr(2))

    ' --- Evaluate direction for Face 1 ---
    Dim p1Forward As Point = tg.CreatePoint( _
        p1.X + v1.X * eps, p1.Y + v1.Y * eps, p1.Z + v1.Z * eps)
    Dim p1Backward As Point = tg.CreatePoint( _
        p1.X - v1.X * eps, p1.Y - v1.Y * eps, p1.Z - v1.Z * eps)

    Dim dFwd1 As Double = mt.GetMinimumDistance(p1Forward, occ1)
    Dim dBack1 As Double = mt.GetMinimumDistance(p1Backward, occ1)

    If dFwd1 < dBack1 Then v1.ScaleBy(-1)

    ' --- Evaluate direction for Face 2 ---
    Dim p2Forward As Point = tg.CreatePoint( _
        p2.X + v2.X * eps, p2.Y + v2.Y * eps, p2.Z + v2.Z * eps)
    Dim p2Backward As Point = tg.CreatePoint( _
        p2.X - v2.X * eps, p2.Y - v2.Y * eps, p2.Z - v2.Z * eps)

    Dim dFwd2 As Double = mt.GetMinimumDistance(p2Forward, occ2)
    Dim dBack2 As Double = mt.GetMinimumDistance(p2Backward, occ2)

    If dFwd2 < dBack2 Then v2.ScaleBy(-1)

    ' --- Final check: normals must be opposing ---
    Dim dotRaw As Double = v1.DotProduct(v2)
    Return dotRaw < (-1.0 + angTol)

End Function


' ==========================================================
' ✅ IMPROVEMENT 4: Overlap Check — accepts pre-built proxies
' ==========================================================
Function DoFacesOverlap(proxy1 As FaceProxy, proxy2 As FaceProxy) As Boolean

    Dim box1 As Box = proxy1.Evaluator.RangeBox
    Dim box2 As Box = proxy2.Evaluator.RangeBox

    Dim tol As Double = 0.001

    Dim overlapX As Boolean = (box1.MinPoint.X <= box2.MaxPoint.X + tol) And (box1.MaxPoint.X + tol >= box2.MinPoint.X)
    Dim overlapY As Boolean = (box1.MinPoint.Y <= box2.MaxPoint.Y + tol) And (box1.MaxPoint.Y + tol >= box2.MinPoint.Y)
    Dim overlapZ As Boolean = (box1.MinPoint.Z <= box2.MaxPoint.Z + tol) And (box1.MaxPoint.Z + tol >= box2.MinPoint.Z)

    ' Determine the "thin" (normal) axis of face1 and ignore it
    Dim spanX As Double = Math.Abs(box1.MaxPoint.X - box1.MinPoint.X)
    Dim spanY As Double = Math.Abs(box1.MaxPoint.Y - box1.MinPoint.Y)
    Dim spanZ As Double = Math.Abs(box1.MaxPoint.Z - box1.MinPoint.Z)

    Dim minSpan As Double = Math.Min(spanX, Math.Min(spanY, spanZ))
    Dim flatTol As Double = 0.001

    If Math.Abs(spanX - minSpan) < flatTol Then
        Return overlapY And overlapZ
    ElseIf Math.Abs(spanY - minSpan) < flatTol Then
        Return overlapX And overlapZ
    Else
        Return overlapX And overlapY
    End If

End Function


' ==========================================================
' ✅ IMPROVEMENT 4: Centroid — accepts pre-built proxies
'    Uses 2D plane projection for robustness with tilted faces
' ==========================================================
Function GetContactCentroid(face1 As Face, face2 As Face) As Point

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry
    ' Dim wpDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition

    If face1.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Or
       face2.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Then
        MessageBox.Show("Both faces must be planar.")
        Exit Function
    End If

    Dim obb1 As FaceOBB = BuildFaceOBB(face1, tg)
    Dim obb2 As FaceOBB = BuildFaceOBB(face2, tg)

    If obb1 Is Nothing Or obb2 Is Nothing Then
        MessageBox.Show("Could not build OBB for one or both faces.")
        Exit Function
    End If

    ' PlaceOBBWorkPoints(obb1, "Face1", wpDef)
    ' PlaceOBBWorkPoints(obb2, "Face2", wpDef)

    ' Project obb2 corners onto obb1's plane and convert directly to 2D — single pass
    Dim poly1 As List(Of Double()) = To2D(obb1.corners, obb1, False, tg)
    Dim poly2 As List(Of Double()) = To2D(obb2.corners, obb1, True, tg)   ' True = project first

    EnsureCCW(poly1)
    EnsureCCW(poly2)

    Dim clipped As List(Of Double()) = ClipPolygon(poly1, poly2)
    Dim centroid2D As Double() = ComputePolygonCentroid(clipped)

    If centroid2D IsNot Nothing Then
        return LocalToWorld(obb1.origin, obb1.xAxis, obb1.yAxis, centroid2D(0), centroid2D(1), tg)
    End If

End Function


Class FaceOBB
    Public origin As Point
    Public xAxis As Vector
    Public yAxis As Vector
    Public normal As Vector
    Public corners As New List(Of Point)
End Class


Function BuildFaceOBB(f As Face, tg As TransientGeometry) As FaceOBB

    Dim data As New FaceOBB()
    Dim plane As Plane = f.Geometry
    Dim n As UnitVector = plane.Normal
    data.normal = tg.CreateVector(n.X, n.Y, n.Z)
    data.origin = GetFaceCentroid(f, tg)

    Dim longestEdge As Edge = Nothing
    Dim maxLen As Double = 0
    For Each e As Edge In f.Edges
        Dim L As Double = e.StartVertex.Point.DistanceTo(e.StopVertex.Point)
        If L > maxLen Then maxLen = L : longestEdge = e
    Next
    If longestEdge Is Nothing Then Return Nothing

    Dim xAxis As Vector = longestEdge.StartVertex.Point.VectorTo(longestEdge.StopVertex.Point)
    xAxis.Normalize()
    Dim yAxis As Vector = data.normal.CrossProduct(xAxis)
    yAxis.Normalize()
    xAxis = yAxis.CrossProduct(data.normal)
    xAxis.Normalize()
    data.xAxis = xAxis
    data.yAxis = yAxis

    Dim minX As Double = Double.MaxValue, minY As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue, maxY As Double = Double.MinValue

    For Each v As Vertex In f.Vertices
        Dim vec As Vector = data.origin.VectorTo(v.Point)
        Dim px As Double = vec.DotProduct(xAxis)
        Dim py As Double = vec.DotProduct(yAxis)
        If px < minX Then minX = px
        If py < minY Then minY = py
        If px > maxX Then maxX = px
        If py > maxY Then maxY = py
    Next

    data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, minX, minY, tg))
    data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, maxX, minY, tg))
    data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, maxX, maxY, tg))
    data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, minX, maxY, tg))

    Return data

End Function

' Converts 3D points to 2D in obb1's frame.
' If project=True, each point is first dropped onto obb1's plane along its normal.
Function To2D(pts As List(Of Point), ref As FaceOBB, project As Boolean, tg As TransientGeometry) As List(Of Double())
    Dim output As New List(Of Double())
    For Each p As Point In pts
        Dim v As Vector = ref.origin.VectorTo(p)
        If project Then
            ' Remove normal component (projects onto plane)
            Dim dist As Double = v.DotProduct(ref.normal)
            v = tg.CreateVector(v.X - dist * ref.normal.X,
                                v.Y - dist * ref.normal.Y,
                                v.Z - dist * ref.normal.Z)
        End If
        output.Add(New Double() {v.DotProduct(ref.xAxis), v.DotProduct(ref.yAxis)})
    Next
    Return output
End Function


Function ClipPolygon(subject As List(Of Double()), clipper As List(Of Double())) As List(Of Double())
    Dim output As List(Of Double()) = subject
    For i As Integer = 0 To clipper.Count - 1
        If output.Count = 0 Then Return output
        Dim input As List(Of Double()) = output
        output = New List(Of Double())
        Dim A As Double() = clipper(i)
        Dim B As Double() = clipper((i + 1) Mod clipper.Count)
        For j As Integer = 0 To input.Count - 1
            Dim P As Double() = input(j)
            Dim Q As Double() = input((j + 1) Mod input.Count)
            Dim Pin As Boolean = IsInsideEdge(P, A, B)
            Dim Qin As Boolean = IsInsideEdge(Q, A, B)
            If Pin AndAlso Qin Then
                output.Add(Q)
            ElseIf Pin Then
                output.Add(EdgeIntersect(P, Q, A, B))
            ElseIf Qin Then
                output.Add(EdgeIntersect(P, Q, A, B))
                output.Add(Q)
            End If
        Next
    Next
    Return output
End Function


Function IsInsideEdge(p As Double(), a As Double(), b As Double()) As Boolean
    Return (b(0) - a(0)) * (p(1) - a(1)) - (b(1) - a(1)) * (p(0) - a(0)) >= 0
End Function


Function EdgeIntersect(p As Double(), q As Double(), a As Double(), b As Double()) As Double()
    Dim A1 As Double = q(1) - p(1), B1 As Double = p(0) - q(0)
    Dim C1 As Double = A1 * p(0) + B1 * p(1)
    Dim A2 As Double = b(1) - a(1), B2 As Double = a(0) - b(0)
    Dim C2 As Double = A2 * a(0) + B2 * a(1)
    Dim det As Double = A1 * B2 - A2 * B1
    If Math.Abs(det) < 1.0E-10 Then Return New Double() {p(0), p(1)}
    Return New Double() {(B2 * C1 - B1 * C2) / det, (A1 * C2 - A2 * C1) / det}
End Function


Function ComputePolygonCentroid(poly As List(Of Double())) As Double()
    If poly Is Nothing OrElse poly.Count = 0 Then Return Nothing
    If poly.Count = 1 Then Return poly(0)
    If poly.Count = 2 Then Return New Double() {(poly(0)(0) + poly(1)(0)) / 2.0,
                                                 (poly(0)(1) + poly(1)(1)) / 2.0}
    Dim cx As Double = 0, cy As Double = 0, area As Double = 0
    For i As Integer = 0 To poly.Count - 1
        Dim j As Integer = (i + 1) Mod poly.Count
        Dim cross As Double = poly(i)(0) * poly(j)(1) - poly(j)(0) * poly(i)(1)
        area += cross
        cx += (poly(i)(0) + poly(j)(0)) * cross
        cy += (poly(i)(1) + poly(j)(1)) * cross
    Next
    area *= 0.5
    If Math.Abs(area) < 1.0E-10 Then Return New Double() {(poly(0)(0) + poly(1)(0)) / 2.0,
                                                            (poly(0)(1) + poly(1)(1)) / 2.0}
    Return New Double() {cx / (6.0 * area), cy / (6.0 * area)}
End Function


Function SignedArea(poly As List(Of Double())) As Double
    Dim area As Double = 0
    For i As Integer = 0 To poly.Count - 1
        Dim j As Integer = (i + 1) Mod poly.Count
        area += poly(i)(0) * poly(j)(1) - poly(j)(0) * poly(i)(1)
    Next
    Return area * 0.5
End Function


Sub EnsureCCW(poly As List(Of Double()))
    If SignedArea(poly) < 0 Then poly.Reverse()
End Sub


Function GetFaceCentroid(f As Face, tg As TransientGeometry) As Point
    Dim sx As Double = 0, sy As Double = 0, sz As Double = 0, count As Integer = 0
    For Each v As Vertex In f.Vertices
        sx += v.Point.X : sy += v.Point.Y : sz += v.Point.Z : count += 1
    Next
    Return tg.CreatePoint(sx / count, sy / count, sz / count)
End Function


Function LocalToWorld(origin As Point, xAxis As Vector, yAxis As Vector,
                      lx As Double, ly As Double, tg As TransientGeometry) As Point
    Dim p As Point = origin.Copy()
    Dim vx As Vector = xAxis.Copy() : vx.ScaleBy(lx)
    Dim vy As Vector = yAxis.Copy() : vy.ScaleBy(ly)
    p.TranslateBy(vx)
    p.TranslateBy(vy)
    Return p
End Function


' ==========================================================
' ✅ Longest Dimension of the coincident overlap area
'    Uses the same 2D projection as GetContactCentroid
' ==========================================================
Function GetContactLongestDimension(proxy1 As FaceProxy, proxy2 As FaceProxy) As Double

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

    Dim eval1 As SurfaceEvaluator = proxy1.Evaluator
    Dim eval2 As SurfaceEvaluator = proxy2.Evaluator

    ' --- Build local 2D coordinate frame from face1's plane ---
    Dim centerUV(1) As Double
    centerUV(0) = (eval1.ParamRangeRect.MinPoint.X + eval1.ParamRangeRect.MaxPoint.X) / 2
    centerUV(1) = (eval1.ParamRangeRect.MinPoint.Y + eval1.ParamRangeRect.MaxPoint.Y) / 2

    Dim originArr(2) As Double
    eval1.GetPointAtParam(centerUV, originArr)
    Dim origin As Point = tg.CreatePoint(originArr(0), originArr(1), originArr(2))

    ' Plane normal
    Dim nArr(2) As Double
    eval1.GetNormal(centerUV, nArr)
    Dim planeN As Vector = tg.CreateVector(nArr(0), nArr(1), nArr(2))
    planeN.Normalize()

    ' Plane X axis (Gram-Schmidt vs world X or Y)
    Dim worldX As Vector = tg.CreateVector(1, 0, 0)
    If Math.Abs(planeN.DotProduct(worldX)) > 0.9 Then
        worldX = tg.CreateVector(0, 1, 0)
    End If

    Dim planeX As Vector = worldX.Copy()
    Dim proj As Vector = planeN.Copy()
    proj.ScaleBy(planeN.DotProduct(planeX))
    planeX.SubtractVector(proj)
    planeX.Normalize()

    ' Plane Y axis
    Dim planeY As Vector = planeN.CrossProduct(planeX)
    planeY.Normalize()

    ' --- Sample both faces into 2D ---
    Dim pts1 As New List(Of Double())
    Dim pts2 As New List(Of Double())

    SampleFacePoints(proxy1, origin, planeX, planeY, pts1)
    SampleFacePoints(proxy2, origin, planeX, planeY, pts2)

    ' --- Get 2D bounding boxes ---
    Dim min1U As Double, max1U As Double, min1V As Double, max1V As Double
    Dim min2U As Double, max2U As Double, min2V As Double, max2V As Double

    GetBounds2D(pts1, min1U, max1U, min1V, max1V)
    GetBounds2D(pts2, min2U, max2U, min2V, max2V)

    ' --- Compute 2D overlap extents ---
    Dim overlapMinU As Double = Math.Max(min1U, min2U)
    Dim overlapMaxU As Double = Math.Min(max1U, max2U)
    Dim overlapMinV As Double = Math.Max(min1V, min2V)
    Dim overlapMaxV As Double = Math.Min(max1V, max2V)

    If overlapMaxU >= overlapMinU And overlapMaxV >= overlapMinV Then
        Dim dimU As Double = overlapMaxU - overlapMinU
        Dim dimV As Double = overlapMaxV - overlapMinV
        Return Math.Max(dimU, dimV)
    Else
        ' Fallback: longest dimension of face1's bounding box
        Return Math.Max(max1U - min1U, max1V - min1V)
    End If

End Function


' ==========================================================
' ✅ Helper: sample face points and project into 2D plane coords
' ==========================================================
Sub SampleFacePoints(proxy As FaceProxy, origin As Point, _
                     planeX As Vector, planeY As Vector, _
                     pts As List(Of Double()))

    Dim eval As SurfaceEvaluator = proxy.Evaluator
    Dim rect As Box2D = eval.ParamRangeRect

    Dim uMin As Double = rect.MinPoint.X
    Dim uMax As Double = rect.MaxPoint.X
    Dim vMin As Double = rect.MinPoint.Y
    Dim vMax As Double = rect.MaxPoint.Y

    Dim steps As Integer = 5   ' 5×5 grid; increase for curved faces

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

                Dim dx As Double = ptArr(0) - origin.X
                Dim dy As Double = ptArr(1) - origin.Y
                Dim dz As Double = ptArr(2) - origin.Z

                Dim pu As Double = dx * planeX.X + dy * planeX.Y + dz * planeX.Z
                Dim pv As Double = dx * planeY.X + dy * planeY.Y + dz * planeY.Z

                pts.Add(New Double() {pu, pv})
            Catch
                ' Skip invalid params (outside face boundary)
            End Try

        Next
    Next

End Sub


' ==========================================================
' ✅ Helper: 2D bounding box from projected point list
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

Function GetJointLength(face1 As Face, face2 As Face) As Double
    Dim tg As TransientGeometry = ThisApplication.TransientGeometry
    ' Dim wpDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition

    If face1.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Or
       face2.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Then
        MessageBox.Show("Both faces must be planar.")
        Exit Function
    End If

    Dim obb1 As FaceOBB = BuildFaceOBB(face1, tg)
    Dim obb2 As FaceOBB = BuildFaceOBB(face2, tg)

    If obb1 Is Nothing Or obb2 Is Nothing Then
        MessageBox.Show("Could not build OBB for one or both faces.")
        Exit Function
    End If

    ' PlaceOBBWorkPoints(obb1, "Face1", wpDef)
    ' PlaceOBBWorkPoints(obb2, "Face2", wpDef)

    ' Project obb2 corners onto obb1's plane and convert directly to 2D — single pass
    Dim poly1 As List(Of Double()) = To2D(obb1.corners, obb1, False, tg)
    Dim poly2 As List(Of Double()) = To2D(obb2.corners, obb1, True, tg)   ' True = project first

    EnsureCCW(poly1)
    EnsureCCW(poly2)

    Dim clipped As List(Of Double()) = ClipPolygon(poly1, poly2)

    Dim maxDist As Double = GetMaxDistanceInPolygon(clipped)

    ' Dim result = GetMaxDistancePoints(clipped)
    ' If result IsNot Nothing Then
    '     Dim p2D As Double() = result.Item1
    '     Dim q2D As Double() = result.Item2

    '     Dim p3D As Point = LocalToWorld(obb1.origin, obb1.xAxis, obb1.yAxis, p2D(0), p2D(1), tg)
    '     Dim q3D As Point = LocalToWorld(obb1.origin, obb1.xAxis, obb1.yAxis, q2D(0), q2D(1), tg)

    '     wpDef.WorkPoints.AddFixed(p3D, False).Name = "MaxSpan_P1"
    '     wpDef.WorkPoints.AddFixed(q3D, False).Name = "MaxSpan_P2"
    ' End If

    ' MessageBox.Show("Max overlap span = " & maxDist/2.54 & " in.")
    return maxDist
End Function

Function GetMaxDistanceInPolygon(poly As List(Of Double())) As Double
    If poly Is Nothing OrElse poly.Count < 2 Then Return 0.0

    Dim maxDist As Double = 0.0

    For i As Integer = 0 To poly.Count - 1
        For j As Integer = i + 1 To poly.Count - 1
            Dim dx As Double = poly(i)(0) - poly(j)(0)
            Dim dy As Double = poly(i)(1) - poly(j)(1)
            Dim dist As Double = Math.Sqrt(dx * dx + dy * dy)

            If dist > maxDist Then maxDist = dist
        Next
    Next

    Return maxDist
End Function

Function GetMaxDistancePoints(poly As List(Of Double())) As Tuple(Of Double(), Double(), Double)
    If poly Is Nothing OrElse poly.Count < 2 Then Return Nothing

    Dim maxDist As Double = 0.0
    Dim bestP As Double() = Nothing
    Dim bestQ As Double() = Nothing

    For i As Integer = 0 To poly.Count - 1
        For j As Integer = i + 1 To poly.Count - 1
            Dim dx As Double = poly(i)(0) - poly(j)(0)
            Dim dy As Double = poly(i)(1) - poly(j)(1)
            Dim dist As Double = Math.Sqrt(dx * dx + dy * dy)

            If dist > maxDist Then
                maxDist = dist
                bestP = poly(i)
                bestQ = poly(j)
            End If
        Next
    Next

    Return Tuple.Create(bestP, bestQ, maxDist)
End Function