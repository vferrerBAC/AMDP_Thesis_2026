' ============================================================
' CheckCrossSection_iLogic.iLogicVb
' Autodesk Inventor iLogic Rule (VB.NET)
'
' Iterates over all part documents in the current assembly
' (or the active part document) and reports whether each
' tube profile has an OPEN or CLOSED cross-section.
'
' Detection method:
'   A closed cross-section has every sketch loop flagged as
'   closed by Inventor's geometry kernel.  An open profile
'   contains at least one open loop.
'
' Works with:
'   - Part documents (.ipt)     — single-body tubes
'   - Assembly documents (.iam) — checks every component part
' ============================================================

Dim oApp As Inventor.Application = ThisApplication
Dim oDoc As Inventor.Document   = oApp.ActiveDocument

Dim results As String = "Cross-Section Analysis" & vbCrLf & _
                        New String("="c, 50)    & vbCrLf

Select Case oDoc.DocumentType

    Case Inventor.DocumentTypeEnum.kPartDocumentObject
        results &= AnalysePart(CType(oDoc, Inventor.PartDocument))

    Case Inventor.DocumentTypeEnum.kAssemblyDocumentObject
        Dim oAsm     As Inventor.AssemblyDocument            = CType(oDoc, Inventor.AssemblyDocument)
        Dim oCompDef As Inventor.AssemblyComponentDefinition = oAsm.ComponentDefinition

        For Each oOcc As Inventor.ComponentOccurrence In oCompDef.Occurrences
            If oOcc.DefinitionDocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
                results &= AnalysePart(CType(oOcc.Definition.Document, Inventor.PartDocument))
            End If
        Next

    Case Else
        MessageBox.Show("Please run this rule from a Part or Assembly document.", _
                        "Unsupported Document Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return

End Select

MessageBox.Show(results, "Cross-Section Results", MessageBoxButtons.OK, MessageBoxIcon.Information)


' ----------------------------------------------------------
' AnalysePart
' Examines all sketches referenced by extrusions / sweeps
' inside a Part document and returns a result string.
' ----------------------------------------------------------
Function AnalysePart(oPartDoc As Inventor.PartDocument) As String

    Dim output As String = vbCrLf & "Part: " & oPartDoc.DisplayName & vbCrLf & _
                           New String("-"c, 40) & vbCrLf

    Dim oCompDef As Inventor.PartComponentDefinition = oPartDoc.ComponentDefinition

    ' List to track sketches already reported
    Dim checkedSketches As New List(Of Inventor.Sketch)

    For Each feat As Inventor.PartFeature In oCompDef.Features

        Dim profileSketch As Inventor.Sketch = Nothing

        Try
            ' --- Extrusion ---
            If TypeOf feat Is Inventor.ExtrudeFeature Then
                Dim oExtrude As Inventor.ExtrudeFeature = CType(feat, Inventor.ExtrudeFeature)
                profileSketch = CType(oExtrude.Profile.Parent, Inventor.Sketch)

            ' --- Sweep (most tubes use a sweep) ---
            ElseIf TypeOf feat Is Inventor.SweepFeature Then
                Dim oSweep As Inventor.SweepFeature = CType(feat, Inventor.SweepFeature)
                profileSketch = CType(oSweep.Profile.Parent, Inventor.Sketch)

            ' --- Coil ---
            ElseIf TypeOf feat Is Inventor.CoilFeature Then
                Dim oCoil As Inventor.CoilFeature = CType(feat, Inventor.CoilFeature)
                profileSketch = CType(oCoil.Profile.Parent, Inventor.Sketch)
            End If
        Catch
            ' Profile not accessible on this feature — skip it
        End Try

        ' Skip if no sketch found or already reported
        If profileSketch IsNot Nothing Then
            If Not checkedSketches.Contains(profileSketch) Then
                checkedSketches.Add(profileSketch)

                Dim sectionType As String = GetSectionType(profileSketch)

                output &= "  Feature : " & feat.Name          & vbCrLf & _
                          "  Sketch  : " & profileSketch.Name & vbCrLf & _
                          "  Section : " & sectionType        & vbCrLf & vbCrLf
            End If
        End If

    Next

    If checkedSketches.Count = 0 Then
        output &= "  (No extrusion/sweep profiles found)" & vbCrLf
    End If

    Return output

End Function


' ----------------------------------------------------------
' GetSectionType
' Returns "CLOSED", "OPEN", "MIXED", or "UNKNOWN" based on
' whether every profile loop in the sketch is closed.
' ----------------------------------------------------------
Function GetSectionType(oSketch As Inventor.Sketch) As String

    Dim totalLoops  As Integer = 0
    Dim closedLoops As Integer = 0

    Dim oProfile As Inventor.Profile = Nothing

    Try
        oProfile = oSketch.Profiles.AddForSolid()
    Catch
        ' Sketch cannot form a solid profile — likely open geometry
    End Try

    ' Primary method: inspect Profile loops (most reliable)
    ' ProfilePath has no IsClosed property — closure is determined by checking
    ' whether the end point of the last segment matches the start of the first.
    If oProfile IsNot Nothing Then
        For Each oPath As Inventor.ProfilePath In oProfile
            totalLoops += 1

            Dim segCount As Integer = oPath.Count
            If segCount > 0 Then
                Try
                    Dim firstSeg As Inventor.ProfileEntity = oPath.Item(1)
                    Dim lastSeg  As Inventor.ProfileEntity = oPath.Item(segCount)

                    ' Each ProfileEntity exposes its underlying SketchEntity curve.
                    ' We evaluate the start/end of the trimmed curve via the
                    ' StartPoint and EndPoint of the sketch entity geometry.
                    Dim firstCurve As Inventor.SketchEntity = firstSeg.SketchEntity
                    Dim lastCurve  As Inventor.SketchEntity = lastSeg.SketchEntity

                    Dim startPt As Inventor.Point2d = Nothing
                    Dim endPt   As Inventor.Point2d = Nothing

                    ' Extract start of first segment
                    If TypeOf firstCurve Is Inventor.SketchLine Then
                        startPt = CType(firstCurve, Inventor.SketchLine).StartSketchPoint.Geometry
                    ElseIf TypeOf firstCurve Is Inventor.SketchArc Then
                        startPt = CType(firstCurve, Inventor.SketchArc).StartSketchPoint.Geometry
                    ElseIf TypeOf firstCurve Is Inventor.SketchSpline Then
                        startPt = CType(firstCurve, Inventor.SketchSpline).SplinePoints.Item(1).Geometry
                    End If

                    ' Extract end of last segment
                    If TypeOf lastCurve Is Inventor.SketchLine Then
                        endPt = CType(lastCurve, Inventor.SketchLine).EndSketchPoint.Geometry
                    ElseIf TypeOf lastCurve Is Inventor.SketchArc Then
                        endPt = CType(lastCurve, Inventor.SketchArc).EndSketchPoint.Geometry
                    ElseIf TypeOf lastCurve Is Inventor.SketchSpline Then
                        Dim sp As Inventor.SketchSpline = CType(lastCurve, Inventor.SketchSpline)
                        endPt = sp.SplinePoints.Item(sp.SplinePoints.Count).Geometry
                    End If

                    ' Compare — tolerance of 0.0001 cm covers floating-point snap
                    If startPt IsNot Nothing AndAlso endPt IsNot Nothing Then
                        If startPt.IsEqualTo(endPt, 0.0001) Then
                            closedLoops += 1
                        End If
                    End If
                Catch
                    ' Could not evaluate this path — skip
                End Try
            End If
        Next
    End If

    ' Fallback: walk sketch lines directly if AddForSolid failed
    If totalLoops = 0 Then
        Dim hasOpenCurve As Boolean = False

        For Each oLine As Inventor.SketchLine In oSketch.SketchLines
            ' If start and end points are different objects the loop is open
            If oLine.StartSketchPoint IsNot oLine.EndSketchPoint Then
                ' Check whether the endpoints share the same position
                Dim sp As Inventor.Point2d = oLine.StartSketchPoint.Geometry
                Dim ep As Inventor.Point2d = oLine.EndSketchPoint.Geometry
                If Not sp.IsEqualTo(ep) Then
                    hasOpenCurve = True
                End If
            End If
        Next

        If hasOpenCurve Then
            Return "OPEN    (fallback check)"
        Else
            Return "CLOSED  (fallback check)"
        End If
    End If

    ' Evaluate primary result
    If totalLoops = 0 Then
        Return "UNKNOWN (no loops detected)"
    ElseIf closedLoops = totalLoops Then
        Return "CLOSED  [" & closedLoops & "/" & totalLoops & " loops closed]"
    ElseIf closedLoops = 0 Then
        Return "OPEN    [0/" & totalLoops & " loops closed]"
    Else
        Return "MIXED   [" & closedLoops & "/" & totalLoops & " loops closed]"
    End If

End Function