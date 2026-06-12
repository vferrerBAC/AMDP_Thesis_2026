Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Math

Sub Main()
    Dim doc As AssemblyDocument = ThisApplication.ActiveDocument
    Dim asmDef As AssemblyComponentDefinition = doc.ComponentDefinition
    ' User defined inputs
    Dim touchTolerance As Double = 0.15 ' Unit lengths to consider parts as touching
    Dim minJointLength As Double = 0.5  ' Minimum length to consider a joint valid
    Dim unitLength As Double = 2.54     ' Conversion factor (from cm)

    ' Dictionary to store contact pairs and contact lengths
    Dim contactLengths As New Dictionary(Of String, Double)

    ' Dictionary to store contact pair centroids (for work point creation)
    Dim contactCentroids As New Dictionary(Of String, Double())

    ' Get all component occurrences
    Dim Components As List(Of ComponentOccurrence) = GetAllLeafOccurrences(asmDef.Occurrences)

    ' Compare each part with every other part
    For i As Integer = 0 To Components.Count - 2
        For j As Integer = i + 1 To Components.Count - 1
            Try
                Dim comp1 As ComponentOccurrence = Components(i)
                Dim comp2 As ComponentOccurrence = Components(j)
                Dim name1 As String = GetComponentName(comp1)
                Dim name2 As String = GetComponentName(comp2)

                ' Check if parts are touching using bounding box proximity
                If AreBoundingBoxesClose(comp1, comp2, touchTolerance, unitLength) Then
                    Dim centroid() As Double = Nothing
                    Dim jointLength As Double = GetContactLengthBetweenParts(comp1, comp2, touchTolerance, unitLength, centroid)

                    ' Store pair key sorted alphabetically to avoid duplicates
                    If jointLength >= minJointLength Then
                        Dim pairKey As String
                        If String.Compare(name1, name2) < 0 Then
                            pairKey = name1 & "," & name2
                        Else
                            pairKey = name2 & "," & name1
                        End If

                        If Not contactLengths.ContainsKey(pairKey) Then
                            contactLengths.Add(pairKey, jointLength)
                            ' Store the centroid for this pair
                            If centroid IsNot Nothing Then
                                contactCentroids.Add(pairKey, centroid)
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
    Next

    ' ---------------------------------------------------------------
    ' Create a Work Point at the centroid of each intersection region
    ' ---------------------------------------------------------------
    Dim workPointsCreated As Integer = 0
    Dim transGeom As TransientGeometry = ThisApplication.TransientGeometry

    For Each pairKey As String In contactCentroids.Keys
        Try
            Dim c() As Double = contactCentroids(pairKey)
            ' c(0)=X, c(1)=Y, c(2)=Z  (in Inventor internal units, cm)
            Dim pt As Point = transGeom.CreatePoint(c(0), c(1), c(2))

            ' Create a fixed work point (grounded, no constraint dependencies)
            Dim wp As WorkPoint = asmDef.WorkPoints.AddFixed(pt)

            ' Name the work point after the pair so it is easy to identify
            ' Replace commas with " x " for a cleaner label
            wp.Name = "WP_" & pairKey.Replace(",", " x ")

            workPointsCreated += 1
        Catch ex As Exception
            ' Skip if work point creation fails for this pair
        End Try
    Next

    ' ---------------------------------------------------------------
    ' Write CSV report (same as before)
    ' ---------------------------------------------------------------
    Try
        Dim docNameNoExt As String = System.IO.Path.GetFileNameWithoutExtension(doc.FullDocumentName)
        Dim desktopPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)
        Dim csvPath As String = System.IO.Path.Combine(
            desktopPath, _
            docNameNoExt & "_JointPairs_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".csv")

        Dim csvContent As New System.Text.StringBuilder()
        csvContent.AppendLine("Part A,Part B,Joint Length Inches")
        For Each pairKey As String In contactLengths.Keys
            Dim parts() As String = pairKey.Split(","c)
            Dim lengthVal As Double = contactLengths(pairKey)
            csvContent.AppendLine(parts(0) & "," & parts(1) & "," & lengthVal.ToString("F3"))
        Next

        System.IO.File.WriteAllText(csvPath, csvContent.ToString())
    Catch ex As Exception
        MessageBox.Show("Failed to write CSV file: " & ex.Message, "Export Error")
    End Try

    MessageBox.Show(
        "Done." & System.Environment.NewLine &
        "Contact pairs found : " & contactLengths.Count & System.Environment.NewLine &
        "Work points created : " & workPointsCreated,
        "Intersection Work Points")
End Sub

' ---------------------------------------------------------------
' Returns True when the two axis-aligned bounding boxes overlap
' or are within tolerance of each other.
' ---------------------------------------------------------------
Function AreBoundingBoxesClose(comp1 As ComponentOccurrence,
                               comp2 As ComponentOccurrence,
                               tolerance As Double,
                               unitLength As Double) As Boolean
    Try
        tolerance = tolerance * unitLength
        Dim box1 As Box = comp1.RangeBox
        Dim box2 As Box = comp2.RangeBox

        Dim xOverlap As Boolean = (box1.MinPoint.X - tolerance <= box2.MaxPoint.X) AndAlso (box2.MinPoint.X - tolerance <= box1.MaxPoint.X)
        Dim yOverlap As Boolean = (box1.MinPoint.Y - tolerance <= box2.MaxPoint.Y) AndAlso (box2.MinPoint.Y - tolerance <= box1.MaxPoint.Y)
        Dim zOverlap As Boolean = (box1.MinPoint.Z - tolerance <= box2.MaxPoint.Z) AndAlso (box2.MinPoint.Z - tolerance <= box1.MaxPoint.Z)
        Return xOverlap AndAlso yOverlap AndAlso zOverlap
    Catch
        Return False
    End Try
End Function

' ---------------------------------------------------------------
' Returns the maximum contact length between two components and
' populates 'centroid' with the XYZ centroid (in Inventor internal
' units, cm) of the best-matching face-pair intersection box.
' ---------------------------------------------------------------
Function GetContactLengthBetweenParts(comp1 As ComponentOccurrence,
                                      comp2 As ComponentOccurrence,
                                      tolerance As Double,
                                      unitLength As Double,
                                      ByRef centroid() As Double) As Double
    Try
        Dim maxContactLength As Double = 0.0
        tolerance = tolerance * unitLength
        centroid = Nothing

        For Each body1 As SurfaceBody In comp1.SurfaceBodies
            For Each face1 As Face In body1.Faces
                Dim bbox1 As Box = face1.Evaluator.RangeBox
                For Each body2 As SurfaceBody In comp2.SurfaceBodies
                    For Each face2 As Face In body2.Faces
                        If face1.SurfaceType = SurfaceTypeEnum.kPlaneSurface AndAlso
                           face2.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then

                            Dim bbox2 As Box = face2.Evaluator.RangeBox

                            ' Compute the extents of the intersection region
                            Dim intMinX As Double = Math.Max(bbox1.MinPoint.X, bbox2.MinPoint.X)
                            Dim intMinY As Double = Math.Max(bbox1.MinPoint.Y, bbox2.MinPoint.Y)
                            Dim intMinZ As Double = Math.Max(bbox1.MinPoint.Z, bbox2.MinPoint.Z)
                            Dim intMaxX As Double = Math.Min(bbox1.MaxPoint.X, bbox2.MaxPoint.X)
                            Dim intMaxY As Double = Math.Min(bbox1.MaxPoint.Y, bbox2.MaxPoint.Y)
                            Dim intMaxZ As Double = Math.Min(bbox1.MaxPoint.Z, bbox2.MaxPoint.Z)

                            Dim xLen As Double = intMaxX - intMinX
                            Dim yLen As Double = intMaxY - intMinY
                            Dim zLen As Double = intMaxZ - intMinZ

                            If xLen >= -tolerance AndAlso yLen >= -tolerance AndAlso zLen >= -tolerance Then
                                Dim candidateLength As Double = Math.Max(xLen, Math.Max(yLen, zLen))
                                If candidateLength > maxContactLength Then
                                    maxContactLength = candidateLength

                                    ' Centroid of the intersection bounding box
                                    ' Clamp degenerate (zero-thickness) axes to midpoint
                                    centroid = New Double(2) {
                                        (intMinX + intMaxX) / 2.0,
                                        (intMinY + intMaxY) / 2.0,
                                        (intMinZ + intMaxZ) / 2.0
                                    }
                                End If
                            End If
                        End If
                    Next
                Next
            Next
        Next

        ' Convert length from internal units (cm) to inches
        Return maxContactLength / unitLength
    Catch
        Return 0.0
    End Try
End Function

Function GetComponentName(comp As ComponentOccurrence) As String
    If comp Is Nothing Then Return "Unknown"
    Try
        Dim fullName As String = comp.Name
        If Not String.IsNullOrEmpty(fullName) Then
            Return fullName
        End If

        Try
            Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(comp.Definition.Document.DisplayName)
            For i As Integer = 1 To comp.Parent.Occurrences.Count
                If comp.Parent.Occurrences.Item(i) Is comp Then
                    Return baseName & ":" & i.ToString()
                End If
            Next
            Return baseName & ":?"
        Catch
            Return fullName
        End Try
    Catch ex As Exception
        Return "Unknown_Instance"
    End Try
End Function

Function GetAllLeafOccurrences(occurrences As ComponentOccurrences) As List(Of ComponentOccurrence)
    Dim allParts As New List(Of ComponentOccurrence)
    For Each occ As ComponentOccurrence In occurrences
        If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim subAsmDef As AssemblyComponentDefinition = CType(occ.Definition, AssemblyComponentDefinition)
            Dim subParts As List(Of ComponentOccurrence) = GetAllLeafOccurrences(subAsmDef.Occurrences)
            For Each subOcc As ComponentOccurrence In subParts
                allParts.Add(subOcc)
            Next
        Else
            If occ.Visible And Not occ.Suppressed Then
                allParts.Add(occ)
            End If
        End If
    Next
    Return allParts
End Function