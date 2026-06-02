' =====================================================
' CONTACT CENTROID CALCULATOR (AREA-WEIGHTED)
' Calculates the centroid of contact area between
' two coincident-constrained components in Inventor
' =====================================================

Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Math

Sub Main()
    
    Dim invApp As Inventor.Application = ThisApplication
    
    Try
        Dim asmDoc As AssemblyDocument = invApp.ActiveDocument
        If asmDoc Is Nothing Then
            MsgBox("Please open an Assembly document", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        
        Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim results As New List(Of ContactInfo)
        
        ' Parameters
        Dim contactTol As Double = 0.01      ' 0.01 cm tolerance for contact
        Dim sampleDensity As Integer = 50    ' Points per face dimension
        
        ' Get all component pairs
        For i = 1 To compDef.Occurrences.Count - 1
            Dim occ1 As ComponentOccurrence = compDef.Occurrences.Item(i)
            
            For j = i + 1 To compDef.Occurrences.Count
                Dim occ2 As ComponentOccurrence = compDef.Occurrences.Item(j)
                
                Dim contact = FindContactCentroid(occ1, occ2, contactTol, sampleDensity)
                If contact IsNot Nothing Then
                    results.Add(contact)
                End If
            Next
        Next
        
        If results.Count = 0 Then
            MsgBox("No contacts found between components", MsgBoxStyle.Information)
        Else
            DisplayResults(results)
            ExportToExcel(results)
        End If
        
    Catch ex As Exception
        MsgBox("Error: " & ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical)
    End Try
    
End Sub

' =====================================================
' CONTACT INFO CLASS
' =====================================================
Public Class ContactInfo
    Public Comp1Name As String
    Public Comp2Name As String
    Public CentroidX As Double
    Public CentroidY As Double
    Public CentroidZ As Double
    Public TotalContactArea As Double
    Public ContactPointCount As Integer
End Class

' =====================================================
' FIND CONTACT CENTROID BETWEEN TWO COMPONENTS
' =====================================================
Function FindContactCentroid(occ1 As ComponentOccurrence, occ2 As ComponentOccurrence, _
                             contactTol As Double, sampleDensity As Integer) As ContactInfo
    
    Dim body1 As SurfaceBody = Nothing
    Dim body2 As SurfaceBody = Nothing
    
    Try
        body1 = occ1.SurfaceBodies.Item(1)
    Catch
        body1 = Nothing
    End Try
    
    Try
        body2 = occ2.SurfaceBodies.Item(1)
    Catch
        body2 = Nothing
    End Try
    
    If body1 Is Nothing Or body2 Is Nothing Then
        Return Nothing
    End If
    
    Dim contactPoints As New List(Of ContactPoint)
    
    ' Check all face pairs
    For Each face1 As Face In body1.Faces
        For Each face2 As Face In body2.Faces
            
            ' Sample both faces and find contact points
            Dim faceContactPoints = SampleFaceContact(face1, face2, contactTol, sampleDensity)
            
            For Each cp In faceContactPoints
                contactPoints.Add(cp)
            Next
        Next
    Next
    
    If contactPoints.Count = 0 Then
        Return Nothing
    End If
    
    ' Calculate area-weighted centroid
    Dim totalArea As Double = 0
    Dim sumX As Double = 0
    Dim sumY As Double = 0
    Dim sumZ As Double = 0
    
    For Each cp In contactPoints
        totalArea += cp.Area
        sumX += cp.X * cp.Area
        sumY += cp.Y * cp.Area
        sumZ += cp.Z * cp.Area
    Next
    
    If totalArea = 0 Then
        Return Nothing
    End If
    
    Dim info As New ContactInfo()
    info.Comp1Name = occ1.Name
    info.Comp2Name = occ2.Name
    info.CentroidX = sumX / totalArea
    info.CentroidY = sumY / totalArea
    info.CentroidZ = sumZ / totalArea
    info.TotalContactArea = totalArea
    info.ContactPointCount = contactPoints.Count
    
    Return info
    
End Function

' =====================================================
' CONTACT POINT CLASS
' =====================================================
Public Class ContactPoint
    Public X As Double
    Public Y As Double
    Public Z As Double
    Public Area As Double
    Public Normal1X As Double
    Public Normal1Y As Double
    Public Normal1Z As Double
    Public Normal2X As Double
    Public Normal2Y As Double
    Public Normal2Z As Double
End Class

' =====================================================
' SAMPLE FACE CONTACT REGION
' =====================================================
Function SampleFaceContact(face1 As Face, face2 As Face, contactTol As Double, _
                          sampleDensity As Integer) As List(Of ContactPoint)
    
    Dim contactPoints As New List(Of ContactPoint)
    
    Try
        Dim eval1 As SurfaceEvaluator = face1.Evaluator
        Dim eval2 As SurfaceEvaluator = face2.Evaluator
        
        ' Get UV bounds for face1
        Dim uMin As Double, uMax As Double, vMin As Double, vMax As Double
        eval1.GetUVBounds(uMin, uMax, vMin, vMax)
        
        Dim uStep As Double = (uMax - uMin) / sampleDensity
        Dim vStep As Double = (vMax - vMin) / sampleDensity
        Dim pointArea As Double = (uStep * vStep) * 0.01 ' Approximate area per point
        
        ' Sample face1 in UV grid
        Dim u As Double = uMin
        While u < uMax
            Dim v As Double = vMin
            While v < vMax
                
                ' Get point on face1
                Dim pt1(2) As Double
                eval1.GetPointAtUV(u, v, pt1)
                
                ' Get normal at face1
                Dim norm1(2) As Double
                eval1.GetNormalAtUV(u, v, norm1)
                
                ' Check contact with face2
                Dim closestPt(2) As Double
                Dim closestU As Double, closestV As Double
                Dim distance As Double = contactTol + 1
                
                Try
                    ' Find closest point on face2
                    eval2.GetClosestPointTo(pt1(0), pt1(1), pt1(2), closestU, closestV)
                    eval2.GetPointAtUV(closestU, closestV, closestPt)
                    
                    distance = Sqrt((pt1(0) - closestPt(0))^2 + _
                                   (pt1(1) - closestPt(1))^2 + _
                                   (pt1(2) - closestPt(2))^2)
                    
                    ' Check if in contact and normals oppose
                    If distance < contactTol Then
                        Dim norm2(2) As Double
                        eval2.GetNormalAtUV(closestU, closestV, norm2)
                        
                        Dim dot As Double = DotProduct(norm1, norm2)
                        Dim absDot As Double = Abs(dot)
                        
                        ' Accept both opposing (-0.85) and perpendicular (~0) normals
                        ' dot < -0.85: Opposing (e.g., flat-on-flat contact)
                        ' absDot < 0.25: Nearly perpendicular (e.g., 90-degree edge/face contact)
                        If dot < -0.85 Or absDot < 0.25 Then
                            Dim cp As New ContactPoint()
                            cp.X = (pt1(0) + closestPt(0)) / 2
                            cp.Y = (pt1(1) + closestPt(1)) / 2
                            cp.Z = (pt1(2) + closestPt(2)) / 2
                            cp.Area = pointArea
                            cp.Normal1X = norm1(0)
                            cp.Normal1Y = norm1(1)
                            cp.Normal1Z = norm1(2)
                            cp.Normal2X = norm2(0)
                            cp.Normal2Y = norm2(1)
                            cp.Normal2Z = norm2(2)
                            
                            contactPoints.Add(cp)
                        End If
                    End If
                Catch
                    ' Skip this point if eval fails
                End Try
                
                v += vStep
            End While
            u += uStep
        End While
        
    Catch ex As Exception
        ' Return what we have
    End Try
    
    Return contactPoints
    
End Function

' =====================================================
' HELPER FUNCTIONS
' =====================================================
Function DotProduct(a() As Double, b() As Double) As Double
    Return a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
End Function

Function Magnitude(v() As Double) As Double
    Return Sqrt(v(0)^2 + v(1)^2 + v(2)^2)
End Function

Function Distance(p1() As Double, p2() As Double) As Double
    Return Sqrt((p1(0) - p2(0))^2 + (p1(1) - p2(1))^2 + (p1(2) - p2(2))^2)
End Function

' =====================================================
' DISPLAY RESULTS IN MESSAGE BOX
' =====================================================
Sub DisplayResults(results As List(Of ContactInfo))
    
    Dim msg As String = "Contact Centroids Found:" & vbCrLf & vbCrLf
    
    For Each contact In results
        msg &= contact.Comp1Name & " <-> " & contact.Comp2Name & vbCrLf
        msg &= String.Format("  Centroid: ({0:F3}, {1:F3}, {2:F3})", _
                            contact.CentroidX, contact.CentroidY, contact.CentroidZ) & vbCrLf
        msg &= String.Format("  Area: {0:F4} cm²", contact.TotalContactArea) & vbCrLf
        msg &= String.Format("  Points: {0}", contact.ContactPointCount) & vbCrLf & vbCrLf
    Next
    
    MsgBox(msg, MsgBoxStyle.Information, "Contact Analysis Results")
    
End Sub

' =====================================================
' EXPORT TO EXCEL
' =====================================================
Sub ExportToExcel(results As List(Of ContactInfo))
    
    Try
        Dim excelApp As Object = CreateObject("Excel.Application")
        excelApp.Visible = True
        
        Dim workbook = excelApp.Workbooks.Add()
        Dim sheet = workbook.Sheets(1)
        
        ' Headers
        sheet.Cells(1, 1).Value = "Component 1"
        sheet.Cells(1, 2).Value = "Component 2"
        sheet.Cells(1, 3).Value = "Centroid X (cm)"
        sheet.Cells(1, 4).Value = "Centroid Y (cm)"
        sheet.Cells(1, 5).Value = "Centroid Z (cm)"
        sheet.Cells(1, 6).Value = "Contact Area (cm²)"
        sheet.Cells(1, 7).Value = "Contact Points"
        
        Dim row As Integer = 2
        For Each contact In results
            sheet.Cells(row, 1).Value = contact.Comp1Name
            sheet.Cells(row, 2).Value = contact.Comp2Name
            sheet.Cells(row, 3).Value = contact.CentroidX
            sheet.Cells(row, 4).Value = contact.CentroidY
            sheet.Cells(row, 5).Value = contact.CentroidZ
            sheet.Cells(row, 6).Value = contact.TotalContactArea
            sheet.Cells(row, 7).Value = contact.ContactPointCount
            row += 1
        Next
        
        sheet.Columns.AutoFit()
        
    Catch ex As Exception
        MsgBox("Could not export to Excel: " & ex.Message, MsgBoxStyle.Critical)
    End Try
    
End Sub
