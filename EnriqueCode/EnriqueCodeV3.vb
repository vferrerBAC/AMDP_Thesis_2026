Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Math

Sub Main()
    Dim doc As AssemblyDocument = ThisApplication.ActiveDocument
    Dim asmDef As AssemblyComponentDefinition = doc.ComponentDefinition

    ' User defined inputs
    Dim touchTolerance As Double = 0.15
    Dim minJointLength As Double = 0.5
    Dim unitLength As Double = 2.54

    ' List to store ALL contact pairs (duplicates allowed)
    Dim contactLengths As New List(Of Tuple(Of String, String, Double))

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

                ' Broad phase check
                If AreBoundingBoxesClose(comp1, comp2, touchTolerance, unitLength) Then

                    Dim jointLength As Double = GetContactLengthBetweenParts(comp1, comp2, touchTolerance, unitLength)

                    If jointLength >= minJointLength Then
                        ' ✅ ADD EVERY RESULT (NO DUPLICATE FILTER)
                        contactLengths.Add(Tuple.Create(name1, name2, jointLength))
                    End If
                End If

            Catch ex As Exception
                Continue For
            End Try
        Next
    Next

    ' Export CSV
    Try
        Dim docNameNoExt As String = System.IO.Path.GetFileNameWithoutExtension(doc.FullDocumentName)
        Dim desktopPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)

        Dim csvPath As String = System.IO.Path.Combine(
            desktopPath,
            docNameNoExt & "_JointPairs_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".csv")

        Dim csvContent As New System.Text.StringBuilder()
        csvContent.AppendLine("Part A,Part B,Joint Length Inches")

        ' ✅ LOOP THROUGH LIST INSTEAD OF DICTIONARY
        For Each entry In contactLengths
            csvContent.AppendLine(entry.Item1 & "," & entry.Item2 & "," & entry.Item3.ToString("F3"))
        Next

        System.IO.File.WriteAllText(csvPath, csvContent.ToString())

        ' Optional confirmation
        MessageBox.Show("CSV saved to: " & csvPath)

    Catch ex As Exception
        MessageBox.Show("Failed to write CSV file: " & ex.Message, "Export Error")
    End Try
End Sub


Function AreBoundingBoxesClose(comp1 As ComponentOccurrence, comp2 As ComponentOccurrence, tolerance As Double, unitLength As Double) As Boolean
    Try
        tolerance = tolerance * unitLength

        Dim box1 As Box = comp1.RangeBox
        Dim box2 As Box = comp2.RangeBox

        Dim xOverlap As Boolean =
            (box1.MinPoint.X - tolerance <= box2.MaxPoint.X) And
            (box2.MinPoint.X - tolerance <= box1.MaxPoint.X)

        Dim yOverlap As Boolean =
            (box1.MinPoint.Y - tolerance <= box2.MaxPoint.Y) And
            (box2.MinPoint.Y - tolerance <= box1.MaxPoint.Y)

        Dim zOverlap As Boolean =
            (box1.MinPoint.Z - tolerance <= box2.MaxPoint.Z) And
            (box2.MinPoint.Z - tolerance <= box1.MaxPoint.Z)

        Return xOverlap And yOverlap And zOverlap

    Catch
        Return False
    End Try
End Function


Function GetContactLengthBetweenParts(comp1 As ComponentOccurrence, comp2 As ComponentOccurrence, tolerance As Double, unitLength As Double) As Double
    Try
        Dim maxContactLength As Double = 0.0
        tolerance = tolerance * unitLength

        For Each body1 As SurfaceBody In comp1.SurfaceBodies
            For Each face1 As Face In body1.Faces

                Dim bbox1 As Box = face1.Evaluator.RangeBox

                For Each body2 As SurfaceBody In comp2.SurfaceBodies
                    For Each face2 As Face In body2.Faces

                        Dim bbox2 As Box = face2.Evaluator.RangeBox

                        If face1.SurfaceType = SurfaceTypeEnum.kPlaneSurface AndAlso _
                           face2.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then

                            Dim xLen As Double = Math.Min(bbox1.MaxPoint.X, bbox2.MaxPoint.X) - Math.Max(bbox1.MinPoint.X, bbox2.MinPoint.X)
                            Dim yLen As Double = Math.Min(bbox1.MaxPoint.Y, bbox2.MaxPoint.Y) - Math.Max(bbox1.MinPoint.Y, bbox2.MinPoint.Y)
                            Dim zLen As Double = Math.Min(bbox1.MaxPoint.Z, bbox2.MaxPoint.Z) - Math.Max(bbox1.MinPoint.Z, bbox2.MinPoint.Z)

                            If xLen >= -tolerance AndAlso yLen >= -tolerance AndAlso zLen >= -tolerance Then
                                Dim candidateLength As Double = Math.Max(xLen, Math.Max(yLen, zLen))

                                If candidateLength > maxContactLength Then
                                    maxContactLength = candidateLength
                                End If
                            End If

                        End If
                    Next
                Next
            Next
        Next

        Return maxContactLength / unitLength

    Catch
        Return 0.0
    End Try
End Function


Function GetComponentName(comp As ComponentOccurrence) As String
    If comp Is Nothing Then Return "Unknown"

    Try
        Dim fullName As String = comp.Name
        If Not String.IsNullOrEmpty(fullName) Then Return fullName

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

    Catch
        Return "Unknown_Instance"
    End Try
End Function


Function GetAllLeafOccurrences(occurrences As ComponentOccurrences) As List(Of ComponentOccurrence)
    Dim allParts As New List(Of ComponentOccurrence)

    For Each occ As ComponentOccurrence In occurrences

        If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then

            Dim subAsmDef As AssemblyComponentDefinition =
                CType(occ.Definition, AssemblyComponentDefinition)

            Dim subParts As List(Of ComponentOccurrence) =
                GetAllLeafOccurrences(subAsmDef.Occurrences)

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