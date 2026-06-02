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

    ' ==========================================================
    ' ✅ Excel setup
    ' ==========================================================
    Dim excelApp As Object = CreateObject("Excel.Application")
    excelApp.Visible = True

    Dim workbook As Object = excelApp.Workbooks.Add
    Dim sheet As Object = workbook.Sheets(1)

    ' Header row
    sheet.Cells(1, 1).Value = "Joint ID"
    sheet.Cells(1, 2).Value = "X"
    sheet.Cells(1, 3).Value = "Y"
    sheet.Cells(1, 4).Value = "Z"
    sheet.Cells(1, 5).Value = "Part 1"
    sheet.Cells(1, 6).Value = "Material 1"
    sheet.Cells(1, 7).Value = "Part 2"
    sheet.Cells(1, 8).Value = "Material 2"

    Dim row As Integer = 2
    Dim jointID As Integer = 1

    ' ==========================================================
    ' ✅ Parameters
    ' ==========================================================
    Dim contactTol As Double = 0.1

    Dim distances As New List(Of Double)

    ' ==========================================================
    ' ✅ Loop through all occurrence pairs
    ' ==========================================================
    For i = 1 To compDef.Occurrences.Count - 1

        Dim occ1 As ComponentOccurrence = compDef.Occurrences.Item(i)

        For j = i + 1 To compDef.Occurrences.Count

            Dim occ2 As ComponentOccurrence = compDef.Occurrences.Item(j)

            ' Skip if either has no bodies
            If occ1.SurfaceBodies.Count = 0 Or occ2.SurfaceBodies.Count = 0 Then Continue For

            Dim body1 As SurfaceBody = occ1.SurfaceBodies.Item(1)
            Dim body2 As SurfaceBody = occ2.SurfaceBodies.Item(1)

            ' ==========================================================
            ' ✅ Check ALL face pairs
            ' ==========================================================
            For Each face1 As Face In body1.Faces

                For Each face2 As Face In body2.Faces

                    ' --- Distance + angle checks ---
                    Dim minDist As Double = GetMinDistance_API(face1, face2)
                    If minDist >= contactTol Then Continue For

                    Dim minAngle As Double = GetAngle(face1, face2)
                    If minAngle >= contactTol Then Continue For

                    ' --- Check normals (true contact) ---
                    If Not CheckNormal(face1, occ1, face2, occ2) Then Continue For

                    ' ==========================================================
                    ' ✅ Valid joint found
                    ' ==========================================================

                    distances.Add(minDist)

                    ' --- Get centroid ---
                    Dim centroid As Point = GetContactCentroid(face1, occ1, face2, occ2)

                    ' --- Create work point (visual in CAD) ---
                    Dim workPt As WorkPoint = compDef.WorkPoints.AddFixed(centroid)
                    workPt.Name = "Joint_" & jointID

                    ' ==========================================================
                    ' ✅ Get part + material info
                    ' ==========================================================
                    Dim part1 As String = GetPartName(occ1)
                    Dim part2 As String = GetPartName(occ2)

                    Dim mat1 As String = GetMaterial(occ1)
                    Dim mat2 As String = GetMaterial(occ2)

                    ' ==========================================================
                    ' ✅ Write to Excel
                    ' ==========================================================
                    sheet.Cells(row, 1).Value = jointID
                    sheet.Cells(row, 2).Value = Round(centroid.X, 4)
                    sheet.Cells(row, 3).Value = Round(centroid.Y, 4)
                    sheet.Cells(row, 4).Value = Round(centroid.Z, 4)
                    sheet.Cells(row, 5).Value = part1
                    sheet.Cells(row, 6).Value = mat1
                    sheet.Cells(row, 7).Value = part2
                    sheet.Cells(row, 8).Value = mat2

                    row += 1
                    jointID += 1

                Next
            Next
        Next
    Next

    ' ==========================================================
    ' ✅ Final formatting
    ' ==========================================================
    sheet.Columns.AutoFit

    MsgBox("There are " & distances.Count & " connections", MsgBoxStyle.Information)

End Sub

' ==========================================================
' ✅ Get minimum distance between faces
' ==========================================================
Function GetMinDistance_API(face1 As Face, face2 As Face) As Double
    Try
        Return ThisApplication.MeasureTools.GetMinimumDistance(face1, face2)
    Catch
        Return 999999
    End Try
End Function

' ==========================================================
' ✅ Get angle between faces
' ==========================================================
Function GetAngle(face1 As Face, face2 As Face) As Double
    Try
        Return ThisApplication.MeasureTools.GetAngle(face1, face2)
    Catch
        Return -1
    End Try
End Function

' ==========================================================
' ✅ Get part name from occurrence
' ==========================================================
Function GetPartName(occ As ComponentOccurrence) As String
    Return occ.Definition.Document.DisplayName
End Function

' ==========================================================
' ✅ Get material from occurrence
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
' ✅ Check if normals are opposite (true contact)
' ==========================================================
Function CheckNormal(face1 As Face, occ1 As ComponentOccurrence, _
                     face2 As Face, occ2 As ComponentOccurrence) As Boolean

    Dim angTol As Double = 1E-3
    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

    Dim proxy1 As FaceProxy
    occ1.CreateGeometryProxy(face1, proxy1)

    Dim proxy2 As FaceProxy
    occ2.CreateGeometryProxy(face2, proxy2)

    If proxy1.SurfaceType <> kPlaneSurface Or proxy2.SurfaceType <> kPlaneSurface Then
        Return False
    End If

    Dim eval1 As SurfaceEvaluator = proxy1.Evaluator
    Dim eval2 As SurfaceEvaluator = proxy2.Evaluator

    ' --- Get midpoint in UV space ---
    Dim center1(1) As Double
    center1(0) = (eval1.ParamRangeRect.MinPoint.X + eval1.ParamRangeRect.MaxPoint.X) / 2
    center1(1) = (eval1.ParamRangeRect.MinPoint.Y + eval1.ParamRangeRect.MaxPoint.Y) / 2

    Dim center2(1) As Double
    center2(0) = (eval2.ParamRangeRect.MinPoint.X + eval2.ParamRangeRect.MaxPoint.X) / 2
    center2(1) = (eval2.ParamRangeRect.MinPoint.Y + eval2.ParamRangeRect.MaxPoint.Y) / 2

    Dim n1(2) As Double
    Dim n2(2) As Double

    eval1.GetNormal(center1, n1)
    eval2.GetNormal(center2, n2)

    Dim v1 As Vector = tg.CreateVector(n1(0), n1(1), n1(2))
    Dim v2 As Vector = tg.CreateVector(n2(0), n2(1), n2(2))

    v1.Normalize
    v2.Normalize

    If proxy1.IsParamReversed Then v1.ScaleBy(-1)
    If proxy2.IsParamReversed Then v2.ScaleBy(-1)

    ' --- Force outward normals ---
    Dim cog1 As Point = occ1.MassProperties.CenterOfMass
    Dim cog2 As Point = occ2.MassProperties.CenterOfMass

    Dim pt1(2) As Double
    Dim pt2(2) As Double

    eval1.GetPointAtParam(center1, pt1)
    eval2.GetPointAtParam(center2, pt2)

    Dim vec1 As Vector = tg.CreateVector(pt1(0) - cog1.X, pt1(1) - cog1.Y, pt1(2) - cog1.Z)
    Dim vec2 As Vector = tg.CreateVector(pt2(0) - cog2.X, pt2(1) - cog2.Y, pt2(2) - cog2.Z)

    vec1.Normalize
    vec2.Normalize

    If v1.DotProduct(vec1) < 0 Then v1.ScaleBy(-1)
    If v2.DotProduct(vec2) < 0 Then v2.ScaleBy(-1)

    ' --- Check if opposite ---
    Return v1.DotProduct(v2) < (-1.0 + angTol)

End Function

' ==========================================================
' ✅ Contact centroid (your current bounding-box approach)
' ==========================================================
Function GetContactCentroid(face1 As Face, occ1 As ComponentOccurrence, _
                             face2 As Face, occ2 As ComponentOccurrence) As Point

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

    Dim proxy1 As FaceProxy
    occ1.CreateGeometryProxy(face1, proxy1)

    Dim proxy2 As FaceProxy
    occ2.CreateGeometryProxy(face2, proxy2)

    Dim box1 As Box = proxy1.Evaluator.RangeBox
    Dim box2 As Box = proxy2.Evaluator.RangeBox

    Dim minX As Double = Math.Max(box1.MinPoint.X, box2.MinPoint.X)
    Dim minY As Double = Math.Max(box1.MinPoint.Y, box2.MinPoint.Y)
    Dim minZ As Double = Math.Max(box1.MinPoint.Z, box2.MinPoint.Z)

    Dim maxX As Double = Math.Min(box1.MaxPoint.X, box2.MaxPoint.X)
    Dim maxY As Double = Math.Min(box1.MaxPoint.Y, box2.MaxPoint.Y)
    Dim maxZ As Double = Math.Min(box1.MaxPoint.Z, box2.MaxPoint.Z)

    ' --- If no overlap, fallback ---
    If minX > maxX Or minY > maxY Or minZ > maxZ Then
        Return tg.CreatePoint( _
            (box1.MinPoint.X + box2.MinPoint.X) / 2, _
            (box1.MinPoint.Y + box2.MinPoint.Y) / 2, _
            (box1.MinPoint.Z + box2.MinPoint.Z) / 2)
    End If

    ' --- Intersection centroid ---
    Return tg.CreatePoint( _
        (minX + maxX) / 2, _
        (minY + maxY) / 2, _
        (minZ + maxZ) / 2)

End Function