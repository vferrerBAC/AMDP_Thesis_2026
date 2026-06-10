Sub Main()

    Dim oAsmDoc As AssemblyDocument = TryCast(ThisApplication.ActiveDocument, AssemblyDocument)
    If oAsmDoc Is Nothing Then Exit Sub

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry
    Dim wpDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition

    Dim face1 As Face = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Select FIRST face (projection target)")
    If face1 Is Nothing Then Exit Sub

    Dim face2 As Face = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Select SECOND face (to project)")
    If face2 Is Nothing Then Exit Sub

    If face1.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Or
       face2.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Then
        MessageBox.Show("Both faces must be planar.")
        Exit Sub
    End If

    Dim obb1 As FaceOBB = BuildFaceOBB(face1, tg)
    Dim obb2 As FaceOBB = BuildFaceOBB(face2, tg)

    If obb1 Is Nothing Or obb2 Is Nothing Then
        MessageBox.Show("Could not build OBB for one or both faces.")
        Exit Sub
    End If

    PlaceOBBWorkPoints(obb1, "Face1", wpDef)
    PlaceOBBWorkPoints(obb2, "Face2", wpDef)

    ' Project obb2 corners onto obb1's plane and convert directly to 2D — single pass
    Dim poly1 As List(Of Double()) = To2D(obb1.corners, obb1, False, tg)
    Dim poly2 As List(Of Double()) = To2D(obb2.corners, obb1, True, tg)   ' True = project first

    EnsureCCW(poly1)
    EnsureCCW(poly2)

    Dim clipped As List(Of Double()) = ClipPolygon(poly1, poly2)
    Dim centroid2D As Double() = ComputePolygonCentroid(clipped)

    Dim maxDist As Double = GetMaxDistanceInPolygon(clipped)

    Dim result = GetMaxDistancePoints(clipped)
    If result IsNot Nothing Then
        Dim p2D As Double() = result.Item1
        Dim q2D As Double() = result.Item2

        Dim p3D As Point = LocalToWorld(obb1.origin, obb1.xAxis, obb1.yAxis, p2D(0), p2D(1), tg)
        Dim q3D As Point = LocalToWorld(obb1.origin, obb1.xAxis, obb1.yAxis, q2D(0), q2D(1), tg)

        wpDef.WorkPoints.AddFixed(p3D, False).Name = "MaxSpan_P1"
        wpDef.WorkPoints.AddFixed(q3D, False).Name = "MaxSpan_P2"
    End If

    MessageBox.Show("Max overlap span = " & maxDist/2.54 & " in.")

    If centroid2D IsNot Nothing Then
        wpDef.WorkPoints.AddFixed(
            LocalToWorld(obb1.origin, obb1.xAxis, obb1.yAxis, centroid2D(0), centroid2D(1), tg),
            False).Name = "Overlap_Centroid"
    End If

End Sub


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


Sub PlaceOBBWorkPoints(obb As FaceOBB, prefix As String, wpDef As AssemblyComponentDefinition)
    For i As Integer = 0 To obb.corners.Count - 1
        wpDef.WorkPoints.AddFixed(obb.corners(i), False).Name = prefix & "_OBB_Corner_" & (i + 1)
    Next
End Sub


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