' ==========================================
' CONTACT CENTROID TOOL (FINAL STABLE VERSION)
' NO AREA DEPENDENCIES
' ==========================================

Imports Inventor

Sub Main()

    Dim invApp As Inventor.Application = ThisApplication
    Dim asmDoc As AssemblyDocument = invApp.ActiveDocument
    Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    
    Dim results As New List(Of String)

    Dim tol As Double = 0.01          ' Contact tolerance
    Dim clusterTol As Double = 0.5    ' Clustering tolerance

    For i = 1 To compDef.Occurrences.Count - 1
        
        Dim occ1 As ComponentOccurrence = compDef.Occurrences.Item(i)
        
        For j = i + 1 To compDef.Occurrences.Count
            
            Dim occ2 As ComponentOccurrence = compDef.Occurrences.Item(j)
            
            ProcessConnection(occ1, occ2, tol, clusterTol, results)
        
        Next
    Next

    ExportToExcel(results)

    MessageBox.Show("Complete. Found " & results.Count & " connections.")

End Sub

' ==========================================
' PROCESS COMPONENT PAIR
' ==========================================
Sub ProcessConnection(occ1 As ComponentOccurrence, occ2 As ComponentOccurrence, tol As Double, clusterTol As Double, results As List(Of String))

    Dim rawPoints As New List(Of Double()) ' just points now

    Dim body1 As SurfaceBody = occ1.SurfaceBodies.Item(1)
    Dim body2 As SurfaceBody = occ2.SurfaceBodies.Item(1)

    For Each face1 As Face In body1.Faces
        
        Dim eval1 As SurfaceEvaluator = face1.Evaluator
        
        ' ✅ SAFE POINT
        Dim pt1(2) As Double
        Dim pnt1 As Point = face1.PointOnFace

        pt1(0) = pnt1.X
        pt1(1) = pnt1.Y
        pt1(2) = pnt1.Z

        ' ✅ SAFE NORMAL
        Dim normal1(2) As Double
        eval1.GetNormalAtPoint(pt1, normal1)

        For Each face2 As Face In body2.Faces
            
            Dim eval2 As SurfaceEvaluator = face2.Evaluator
            
            ' ✅ SAFE POINT
            Dim pt2(2) As Double
            Dim pnt2 As Point = face2.PointOnFace

            pt2(0) = pnt2.X
            pt2(1) = pnt2.Y
            pt2(2) = pnt2.Z

            ' ✅ SAFE NORMAL
            Dim normal2(2) As Double
            eval2.GetNormalAtPoint(pt2, normal2)

            ' Check opposing normals
            Dim dot As Double = DotProduct(normal1, normal2)

            If dot < -0.95 Then

                Dim dist As Double = Distance(pt1, pt2)

                If dist < tol Then

                    Dim pt() As Double = GetFacePoint(face1)
                    rawPoints.Add(pt)

                    Exit For

                End If

            End If

        Next
    Next

    ' ✅ Cluster into one centroid
    Dim clustered As List(Of Double()) = ClusterCentroids(rawPoints, clusterTol)

    For Each c In clustered

        Dim result As String = _
            occ1.Name & "," & occ2.Name & "," & _
            c(0) & "," & c(1) & "," & c(2)

        results.Add(result)

    Next

End Sub

' ==========================================
' CLUSTER CENTROIDS (EQUAL WEIGHT)
' ==========================================
Function ClusterCentroids(points As List(Of Double()), tol As Double) As List(Of Double())

    Dim clusters As New List(Of List(Of Double()))

    For Each pt In points
        
        Dim added As Boolean = False

        For Each cluster In clusters

            If Distance(pt, cluster(0)) < tol Then
                cluster.Add(pt)
                added = True
                Exit For
            End If

        Next

        If Not added Then
            Dim newCluster As New List(Of Double())
            newCluster.Add(pt)
            clusters.Add(newCluster)
        End If

    Next

    Dim results As New List(Of Double())

    For Each cluster In clusters
        
        Dim sumX As Double = 0
        Dim sumY As Double = 0
        Dim sumZ As Double = 0

        For Each pt In cluster
            
            sumX += pt(0)
            sumY += pt(1)
            sumZ += pt(2)

        Next

        Dim centroid(2) As Double
        centroid(0) = sumX / cluster.Count
        centroid(1) = sumY / cluster.Count
        centroid(2) = sumZ / cluster.Count

        results.Add(centroid)

    Next

    Return results

End Function

' ==========================================
' FACE POINT
' ==========================================
Function GetFacePoint(face As Face) As Double()

    Dim pt(2) As Double
    Dim pnt As Point = face.PointOnFace

    pt(0) = pnt.X
    pt(1) = pnt.Y
    pt(2) = pnt.Z

    Return pt

End Function

' ==========================================
' HELPERS
' ==========================================
Function Distance(p1() As Double, p2() As Double) As Double
    Return Math.Sqrt((p1(0)-p2(0))^2 + (p1(1)-p2(1))^2 + (p1(2)-p2(2))^2)
End Function

Function DotProduct(a() As Double, b() As Double) As Double
    Return a(0)*b(0) + a(1)*b(1) + a(2)*b(2)
End Function

' ==========================================
' EXPORT TO EXCEL
' ==========================================
Sub ExportToExcel(results As List(Of String))

    Dim excelApp As Object = CreateObject("Excel.Application")
    excelApp.Visible = True

    Dim workbook = excelApp.Workbooks.Add()
    Dim sheet = workbook.Sheets(1)

    sheet.Cells(1, 1).Value = "Component 1"
    sheet.Cells(1, 2).Value = "Component 2"
    sheet.Cells(1, 3).Value = "X"
    sheet.Cells(1, 4).Value = "Y"
    sheet.Cells(1, 5).Value = "Z"

    Dim row As Integer = 2

    For Each line As String In results
        
        Dim parts() As String = Split(line, ",")

        sheet.Cells(row, 1).Value = parts(0)
        sheet.Cells(row, 2).Value = parts(1)
        sheet.Cells(row, 3).Value = CDbl(parts(2))
        sheet.Cells(row, 4).Value = CDbl(parts(3))
        sheet.Cells(row, 5).Value = CDbl(parts(4))

        row += 1

    Next

    sheet.Columns.AutoFit()

End Sub