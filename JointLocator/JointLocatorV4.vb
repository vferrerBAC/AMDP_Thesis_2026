' =====================================================
' CONTACT CENTROID CALCULATOR V4 (DISTANCE-BASED)
' Finds contact area centroid by detecting faces
' with zero distance between them, then overlapping
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

        MsgBox("There are " & compDef.Occurrences.Count & " components in the assembly.", MsgBoxStyle.Information)
        
        ' Parameters
        Dim contactTol As Double = 0.01      ' Tolerance for "touching"
        Dim sampleDensity As Integer = 30    ' Grid points per face
        
        ' Get all component pairs
        For i = 1 To compDef.Occurrences.Count - 1
            Dim occ1 As ComponentOccurrence = compDef.Occurrences.Item(i)
            
            For j = i + 1 To compDef.Occurrences.Count
                Dim occ2 As ComponentOccurrence = compDef.Occurrences.Item(j)
                
                Dim contacts = FindAllContacts(occ1, occ2, contactTol, sampleDensity)
                For Each contact In contacts
                    results.Add(contact)
                Next
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
' FIND ALL CONTACTS BETWEEN TWO COMPONENTS
' =====================================================
Function FindAllContacts(occ1 As ComponentOccurrence, occ2 As ComponentOccurrence, _
                         contactTol As Double, sampleDensity As Integer) As List(Of ContactInfo)
    
    Dim results As New List(Of ContactInfo)
    
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
        ' Not exiting here
        Return results
    End If
    
    ' Check all face pairs for zero distance
    For Each face1 As Face In body1.Faces
        For Each face2 As Face In body2.Faces
            
            ' Check if faces are touching (minimum distance ≈ 0)
            Dim minDist As Double = GetMinimumFaceDistance(face1, face2)
            
            If minDist < contactTol Then
                ' These faces are in contact
                Dim contact = CalculateContactCentroid(occ1, occ2, face1, face2, sampleDensity)
                If contact IsNot Nothing Then
                    results.Add(contact)
                End If
            End If
        Next
    Next
    
    Return results
    
End Function

' =====================================================
' GET MINIMUM DISTANCE BETWEEN TWO FACES
' =====================================================
Function GetMinimumFaceDistance(face1 As Face, face2 As Face) As Double
    
    Dim minDist As Double = Double.MaxValue
    
    Try
        Dim eval1 As SurfaceEvaluator = face1.Evaluator
        Dim eval2 As SurfaceEvaluator = face2.Evaluator
        
        ' Get UV bounds for face1
        Dim uMin As Double, uMax As Double, vMin As Double, vMax As Double
        eval1.GetUVBounds(uMin, uMax, vMin, vMax)
        
        ' Sample a coarse grid (10x10) to find minimum distance quickly
        Dim gridSize As Integer = 10
        Dim uStep As Double = (uMax - uMin) / gridSize
        Dim vStep As Double = (vMax - vMin) / gridSize
        
        Dim u As Double = uMin
        While u < uMax And minDist > 0.001
            Dim v As Double = vMin
            While v < vMax And minDist > 0.001
                
                Dim pt1(2) As Double
                eval1.GetPointAtUV(u, v, pt1)
                
                ' Find closest point on face2
                Dim closestU As Double, closestV As Double
                Dim pt2(2) As Double
                
                Try
                    eval2.GetClosestPointTo(pt1(0), pt1(1), pt1(2), closestU, closestV)
                    eval2.GetPointAtUV(closestU, closestV, pt2)
                    
                    Dim dist As Double = Distance(pt1, pt2)
                    If dist < minDist Then
                        minDist = dist
                    End If
                Catch
                End Try
                
                v += vStep
            End While
            u += uStep
        End While
        
    Catch ex As Exception
    End Try
    
    Return minDist
    
End Function

' =====================================================
' CALCULATE CONTACT CENTROID FOR FACE PAIR
' =====================================================
Function CalculateContactCentroid(occ1 As ComponentOccurrence, occ2 As ComponentOccurrence, _
                                  face1 As Face, face2 As Face, sampleDensity As Integer) As ContactInfo
    
    Dim contactPoints As New List(Of Double())
    
    Try
        Dim eval1 As SurfaceEvaluator = face1.Evaluator
        Dim eval2 As SurfaceEvaluator = face2.Evaluator
        
        ' Get UV bounds for face1
        Dim uMin As Double, uMax As Double, vMin As Double, vMax As Double
        eval1.GetUVBounds(uMin, uMax, vMin, vMax)
        
        Dim uStep As Double = (uMax - uMin) / sampleDensity
        Dim vStep As Double = (vMax - vMin) / sampleDensity
        
        ' Sample face1
        Dim u As Double = uMin
        While u < uMax
            Dim v As Double = vMin
            While v < vMax
                
                Dim pt1(2) As Double
                eval1.GetPointAtUV(u, v, pt1)
                
                ' Find closest point on face2
                Dim closestU As Double, closestV As Double
                Dim pt2(2) As Double
                
                Try
                    eval2.GetClosestPointTo(pt1(0), pt1(1), pt1(2), closestU, closestV)
                    eval2.GetPointAtUV(closestU, closestV, pt2)
                    
                    Dim dist As Double = Distance(pt1, pt2)
                    
                    ' Point is in contact if distance is very small
                    If dist < 0.001 Then
                        ' Midpoint between the two surfaces
                        Dim contactPt(2) As Double
                        contactPt(0) = (pt1(0) + pt2(0)) / 2
                        contactPt(1) = (pt1(1) + pt2(1)) / 2
                        contactPt(2) = (pt1(2) + pt2(2)) / 2
                        
                        contactPoints.Add(contactPt)
                    End If
                Catch
                End Try
                
                v += vStep
            End While
            u += uStep
        End While
        
    Catch ex As Exception
        Return Nothing
    End Try
    
    If contactPoints.Count < 2 Then
        ' Not enough contact points to form meaningful area
        Return Nothing
    End If
    
    ' Calculate centroid (equal weight for each sample point)
    Dim sumX As Double = 0
    Dim sumY As Double = 0
    Dim sumZ As Double = 0
    
    For Each pt In contactPoints
        sumX += pt(0)
        sumY += pt(1)
        sumZ += pt(2)
    Next
    
    Dim info As New ContactInfo()
    info.Comp1Name = occ1.Name
    info.Comp2Name = occ2.Name
    info.CentroidX = sumX / contactPoints.Count
    info.CentroidY = sumY / contactPoints.Count
    info.CentroidZ = sumZ / contactPoints.Count
    info.TotalContactArea = contactPoints.Count * 0.001  ' Rough estimate
    info.ContactPointCount = contactPoints.Count
    
    Return info
    
End Function

' =====================================================
' HELPER FUNCTIONS
' =====================================================
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
        msg &= String.Format("  Contact Points: {0}", contact.ContactPointCount) & vbCrLf & vbCrLf
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
        sheet.Cells(1, 6).Value = "Contact Points"
        
        Dim row As Integer = 2
        For Each contact In results
            sheet.Cells(row, 1).Value = contact.Comp1Name
            sheet.Cells(row, 2).Value = contact.Comp2Name
            sheet.Cells(row, 3).Value = contact.CentroidX
            sheet.Cells(row, 4).Value = contact.CentroidY
            sheet.Cells(row, 5).Value = contact.CentroidZ
            sheet.Cells(row, 6).Value = contact.ContactPointCount
            row += 1
        Next
        
        sheet.Columns.AutoFit()
        
    Catch ex As Exception
        MsgBox("Could not export to Excel: " & ex.Message, MsgBoxStyle.Critical)
    End Try
    
End Sub
