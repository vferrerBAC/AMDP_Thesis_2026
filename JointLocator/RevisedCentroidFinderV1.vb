Sub Main()

    Dim oAsmDoc As AssemblyDocument = TryCast(ThisApplication.ActiveDocument, AssemblyDocument)
    If oAsmDoc Is Nothing Then Exit Sub

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry
    Dim wpDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition

    Dim face1 As Face = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Select FIRST face (projection target plane)")
    If face1 Is Nothing Then Exit Sub

    Dim face2 As Face = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Select SECOND face (to be projected)")
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

    Dim obb2Projected As New List(Of Point)
    For Each pt As Point In obb2.corners
        obb2Projected.Add(ProjectPointOntoPlane(pt, obb1, tg))
    Next

    Dim poly1 As List(Of Double()) = To2D(obb1.corners, obb1)
    Dim poly2 As List(Of Double()) = To2D(obb2Projected, obb1)
	
	EnsureCCW(poly1)
    EnsureCCW(poly2)
	
    Dim clipped As List(Of Double()) = ClipPolygon(poly1, poly2)
    Dim centroid2D As Double() = ComputePolygonCentroid(clipped)

    If centroid2D IsNot Nothing Then
        Dim centroid3D As Point = LocalToWorld(obb1.origin, obb1.xAxis, obb1.yAxis,
                                               centroid2D(0), centroid2D(1), tg)
        wpDef.WorkPoints.AddFixed(centroid3D, False).Name = "Overlap_Centroid"
    End If

End Sub


' =========================
' DATA CLASS
' =========================
Class FaceOBB
    Public origin As Point       ' Centroid of face (local frame origin)
    Public xAxis As Vector       ' Local X axis (along longest edge)
    Public yAxis As Vector       ' Local Y axis (perpendicular in plane)
    Public normal As Vector      ' Face normal
    Public corners As New List(Of Point)   ' 4 OBB corners in world space
End Class


' =========================
' BUILD OBB FOR A PLANAR FACE
' Returns the oriented bounding box aligned to the face's longest edge
' =========================
Function BuildFaceOBB(f As Face, tg As TransientGeometry) As FaceOBB

    Dim data As New FaceOBB()

    Dim plane As Plane = f.Geometry
    Dim n As UnitVector = plane.Normal
    data.normal = tg.CreateVector(n.X, n.Y, n.Z)

    data.origin = GetFaceCentroid(f, tg)

    ' Build local frame: xAxis along longest edge, yAxis in-plane perpendicular
    Dim longestEdge As Edge = Nothing
    Dim maxLen As Double = 0
    For Each e As Edge In f.Edges
        Dim L As Double = E.StartVertex.Point.DistanceTo(E.StopVertex.Point)
        If L > maxLen Then
            maxLen = L
            longestEdge = E
        End If
    Next

    If longestEdge Is Nothing Then Return Nothing

    Dim xAxis As Vector = longestEdge.StartVertex.Point.VectorTo(longestEdge.StopVertex.Point)
    xAxis.Normalize()

    Dim yAxis As Vector = data.normal.CrossProduct(xAxis)
    yAxis.Normalize()

    ' Re-orthogonalise xAxis to guarantee a clean frame
    xAxis = yAxis.CrossProduct(data.normal)
    xAxis.Normalize()

    data.xAxis = xAxis
    data.yAxis = yAxis

    ' Project all face vertices into local 2D to find extents
    Dim minX As Double = Double.MaxValue
    Dim minY As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue
    Dim maxY As Double = Double.MinValue

    For Each v As Vertex In f.Vertices
        Dim vec As Vector = data.origin.VectorTo(v.Point)
        Dim px As Double = vec.DotProduct(xAxis)
        Dim py As Double = vec.DotProduct(yAxis)
        minX = Math.Min(minX, px)
        minY = Math.Min(minY, py)
        maxX = Math.Max(maxX, px)
        maxY = Math.Max(maxY, py)
    Next

    ' Four OBB corners in world space (CCW winding)
    data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, minX, minY, tg))
    data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, maxX, minY, tg))
    data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, maxX, maxY, tg))
    data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, minX, maxY, tg))

    Return data

End Function


' =========================
' PLACE OBB WORK POINTS
' Creates a work point at every OBB corner for visual inspection
' =========================
Sub PlaceOBBWorkPoints(obb As FaceOBB, prefix As String, wpDef As AssemblyComponentDefinition)

    For i As Integer = 0 To obb.corners.Count - 1
        Dim wp As WorkPoint = wpDef.WorkPoints.AddFixed(obb.corners(i), False)
        wp.Name = prefix & "_OBB_Corner_" & (i + 1)
    Next

End Sub


' =========================
' PROJECT A 3D POINT ONTO FACE1'S PLANE
' Drops the point along face1's normal onto the plane defined by face1's origin + normal.
' This is a true orthographic projection: the result lies on face1's plane.
' =========================
Function ProjectPointOntoPlane(pt As Point, obb1 As FaceOBB, tg As TransientGeometry) As Point

    ' Vector from plane origin to the point
    Dim v As Vector = obb1.origin.VectorTo(pt)

    ' Signed distance along normal
    Dim dist As Double = v.DotProduct(obb1.normal)

    ' Subtract the normal component to land on the plane
    Dim projected As Point = pt.Copy()
    Dim offset As Vector = obb1.normal.Copy()
    offset.ScaleBy(-dist)
    projected.TranslateBy(offset)

    Return projected

End Function


' =========================
' CONVERT 3D WORLD POINTS TO 2D LOCAL COORDS (in obb1's frame)
' =========================
Function To2D(pts As List(Of Point), refFrame As FaceOBB) As List(Of Double())

    Dim output As New List(Of Double())
    For Each p As Point In pts
        Dim v As Vector = refFrame.origin.VectorTo(p)
        Dim arr(1) As Double
        arr(0) = v.DotProduct(refFrame.xAxis)
        arr(1) = v.DotProduct(refFrame.yAxis)
        output.Add(arr)
    Next
    Return output

End Function


' =========================
' SUTHERLAND-HODGMAN POLYGON CLIP
' Clips subject polygon against each edge of the clipper polygon.
' Both polygons must be in the same 2D plane (CCW winding assumed).
' =========================
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

            Dim Pinside As Boolean = IsInsideEdge(P, A, B)
            Dim Qinside As Boolean = IsInsideEdge(Q, A, B)

            If Pinside AndAlso Qinside Then
                ' Both inside: keep Q
                output.Add(Q)
            ElseIf Pinside AndAlso Not Qinside Then
                ' Exiting: add intersection
                output.Add(EdgeIntersect(P, Q, A, B))
            ElseIf Not Pinside AndAlso Qinside Then
                ' Entering: add intersection then Q
                output.Add(EdgeIntersect(P, Q, A, B))
                output.Add(Q)
            End If
            ' Both outside: add nothing
        Next

    Next

    Return output

End Function


' Left-of-edge test (inside = left side of directed edge A→B, CCW polygon)
Function IsInsideEdge(p As Double(), a As Double(), b As Double()) As Boolean
    Return (b(0) - a(0)) * (p(1) - a(1)) - (b(1) - a(1)) * (p(0) - a(0)) >= 0
End Function


' Line-line intersection of segment P→Q with infinite line A→B
Function EdgeIntersect(p As Double(), q As Double(), a As Double(), b As Double()) As Double()

    Dim res(1) As Double

    Dim A1 As Double = q(1) - p(1)
    Dim B1 As Double = p(0) - q(0)
    Dim C1 As Double = A1 * p(0) + B1 * p(1)

    Dim A2 As Double = b(1) - a(1)
    Dim B2 As Double = a(0) - b(0)
    Dim C2 As Double = A2 * a(0) + B2 * a(1)

    Dim det As Double = A1 * B2 - A2 * B1

    If Math.Abs(det) < 1.0E-10 Then
        ' Parallel: return P as fallback
        res(0) = p(0)
        res(1) = p(1)
    Else
        res(0) = (B2 * C1 - B1 * C2) / det
        res(1) = (A1 * C2 - A2 * C1) / det
    End If

    Return res

End Function


' =========================
' COMPUTE POLYGON CENTROID via shoelace formula
' Returns Nothing if polygon has zero area (degenerate)
' =========================
Function ComputePolygonCentroid(poly As List(Of Double())) As Double()

    ' --- Area case (3+ points forming a real polygon) ---
    If poly.Count >= 3 Then

        Dim cx As Double = 0, cy As Double = 0, area As Double = 0

        For i As Integer = 0 To poly.Count - 1
            Dim j As Integer = (i + 1) Mod poly.Count
            Dim xi As Double = poly(i)(0), yi As Double = poly(i)(1)
            Dim xj As Double = poly(j)(0), yj As Double = poly(j)(1)
            Dim cross As Double = xi * yj - xj * yi
            area += cross
            cx += (xi + xj) * cross
            cy += (yi + yj) * cross
        Next

        area *= 0.5

        If Math.Abs(area) > 1.0E-10 Then
            cx /= (6.0 * area)
            cy /= (6.0 * area)
            Return New Double() {cx, cy}
        End If

    End If

    ' --- Line segment case (exactly 2 points or degenerate area) ---
    If poly.Count = 2 Then
        Return New Double() {(poly(0)(0) + poly(1)(0)) / 2.0,
                             (poly(0)(1) + poly(1)(1)) / 2.0}
    End If

    ' --- Single point case ---
    If poly.Count = 1 Then
        Return New Double() {poly(0)(0), poly(0)(1)}
    End If

    Return Nothing

End Function


' =========================
' HELPERS
' =========================
Function GetFaceCentroid(f As Face, tg As TransientGeometry) As Point

    Dim sx As Double = 0, sy As Double = 0, sz As Double = 0, count As Integer = 0
    For Each v As Vertex In f.Vertices
        sx += v.Point.X : sy += v.Point.Y : sz += v.Point.Z
        count += 1
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

' =========================
' COMPUTE SIGNED AREA (positive = CCW, negative = CW)
' =========================
Function SignedArea(poly As List(Of Double())) As Double
    Dim area As Double = 0
    For i As Integer = 0 To poly.Count - 1
        Dim j As Integer = (i + 1) Mod poly.Count
        area += poly(i)(0) * poly(j)(1)
        area -= poly(j)(0) * poly(i)(1)
    Next
    Return area * 0.5
End Function


' =========================
' ENSURE CCW WINDING — reverse in place if CW
' =========================
Sub EnsureCCW(poly As List(Of Double()))
    If SignedArea(poly) < 0 Then
        poly.Reverse()
    End If
End Sub