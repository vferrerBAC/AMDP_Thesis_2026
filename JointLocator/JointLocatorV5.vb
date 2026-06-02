Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Math
Imports System.Threading
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


' Parameters
Dim contactTol As Double = 0.1      ' Tolerance for "touching"
Dim sampleDensity As Integer = 30    ' Grid points per face

Dim distances As New List(Of Double)
Dim face1 As Face
Dim face2 As Face
'MsgBox("There are " & compDef.Occurrences.Count & " components in the assembly.", MsgBoxStyle.Information)

' Get all component pairs
For i = 1 To compDef.Occurrences.Count - 1
	Dim occ1 As ComponentOccurrence = compDef.Occurrences.Item(i)

	For j = i + 1 To compDef.Occurrences.Count
	    Dim occ2 As ComponentOccurrence = compDef.Occurrences.Item(j)


		'Dim occ1 As ComponentOccurrence = compDef.Occurrences.Item(1)
		'Dim occ2 As ComponentOccurrence = compDef.Occurrences.Item(2)
		
		Dim body1 As SurfaceBody = Nothing
		Dim body2 As SurfaceBody = Nothing
		
		body1 = occ1.SurfaceBodies.Item(1)
		body2 = occ2.SurfaceBodies.Item(1)
		
		'MsgBox("There are " & occ1.SurfaceBodies.Count & " SurfaceBodies in the red member.", MsgBoxStyle.Information)
		'MsgBox("There are " & occ2.SurfaceBodies.Count & " SurfaceBodies in the green member.", MsgBoxStyle.Information)
		
		'MsgBox("There are " & body1.Faces.Count & " faces in the red member.", MsgBoxStyle.Information)
		'MsgBox("There are " & body2.Faces.Count & " faces in the green member.", MsgBoxStyle.Information)
		
		' Check all face pairs for zero distance
		
		
		For Each face1 In body1.Faces
		    For Each face2 In body2.Faces
		        
		        ' Check if faces are touching (minimum distance ≈ 0)
		        Dim minDist As Double = GetMinDistance_API(face1, face2)
				Dim minAngle As Double = GetAngle(face1, face2)
					
		        
		        If minDist < contactTol Then
					If minAngle < contactTol Then
						
			'			Dim faceProxy1 As FaceProxy
						occ1.CreateGeometryProxy(face1, faceProxy1)
						
			'			Dim faceProxy2 As FaceProxy
						occ2.CreateGeometryProxy(face2, faceProxy2)
						
						highlightSet1.AddItem(faceProxy1)
						highlightSet2.AddItem(faceProxy2)
						
						Dim coincident As Boolean = AreFacesCoincidentAndOpposite(face1, occ1, face2, occ2)
						
						MsgBox("Coincident: " & coincident)
						
						MsgBox("Press OK to view next contact pair")
						
						highlightSet1.Clear()
			            highlightSet2.Clear()
						If coincident Then
							distances.Add(minDist)
						End If
					End If
		        End If
		    Next
		Next
	Next
Next

'Dim faceProxy1 As FaceProxy
'occ1.CreateGeometryProxy(face1, faceProxy1)

'Dim faceProxy2 As FaceProxy
'occ2.CreateGeometryProxy(face2, faceProxy2)

'highlightSet1.AddItem(faceProxy1)
'highlightSet2.AddItem(faceProxy2)

MsgBox("There are " & distances.Count & " connections", MsgBoxStyle.Information)
'MsgBox("Press OK to view next contact pair")

'highlightSet1.Clear()
'highlightSet2.Clear()

End Sub

Function GetMinDistance_API(face1 As Face, face2 As Face) As Double
    
    Dim invApp As Application = ThisApplication
    Dim measureTools As MeasureTools = invApp.MeasureTools
    
    Try
        ' This returns the minimum distance directly
        Dim minDist As Double = measureTools.GetMinimumDistance(face1, face2)
        
        Return minDist
        
    Catch ex As Exception
        MsgBox("Error computing distance: " & ex.Message)
        Return 999999
    End Try
    
End Function

Function GetAngle(face1 As Face, face2 As Face) As Double
    
    Dim invApp As Application = ThisApplication
    Dim measureTools As MeasureTools = invApp.MeasureTools
    
    Try
        ' This returns the minimum distance directly
        Dim minAngle As Double = measureTools.GetAngle(face1, face2)
        
        Return minAngle
        
    Catch ex As Exception
        MsgBox("Error computing angle: " & ex.Message)
        Return -1
    End Try
    
End Function

Function Distance(p1() As Double, p2() As Double) As Double
    Return Sqrt((p1(0) - p2(0))^2 + (p1(1) - p2(1))^2 + (p1(2) - p2(2))^2)
End Function

Function AreFacesCoincidentAndOpposite(face1 As Face, occ1 As ComponentOccurrence, _
                                       face2 As Face, occ2 As ComponentOccurrence) As Boolean

    Dim angTol As Double = 1E-3     ' relaxed slightly for stability
    Dim distTol As Double = 1E-4

    Dim invApp As Application = ThisApplication
    Dim measureTools As MeasureTools = invApp.MeasureTools

    ' --- Create proxies ---
    Dim proxy1 As FaceProxy
    If TypeOf face1 Is FaceProxy Then
        proxy1 = face1
    Else
        occ1.CreateGeometryProxy(face1, proxy1)
    End If

    Dim proxy2 As FaceProxy
    If TypeOf face2 Is FaceProxy Then
        proxy2 = face2
    Else
        occ2.CreateGeometryProxy(face2, proxy2)
    End If

    ' --- Only planar faces ---
    If proxy1.SurfaceType <> kPlaneSurface Or proxy2.SurfaceType <> kPlaneSurface Then
        Return False
    End If

    ' --- Get normals ---
    Dim plane1 As Plane = proxy1.Geometry
    Dim plane2 As Plane = proxy2.Geometry
	

    Dim v1 As Vector = plane1.Normal.AsVector
    Dim v2 As Vector = plane2.Normal.AsVector
	

    ' Apply face orientation
    If proxy1.IsParamReversed Then v1.ScaleBy(-1)
    If proxy2.IsParamReversed Then v2.ScaleBy(-1)

    v1.Normalize()
    v2.Normalize()
	
'	MsgBox("V1: (" & v1.X & ", " & v1.Y & ", " & v1.Z & ")")
'	MsgBox("V2: (" & v2.X & ", " & v2.Y & ", " & v2.Z & ")")
    
	' --- Check normals ---
    Dim dot As Double = v1.DotProduct(v2)
'	MsgBox("Dot Product: " & dot)
	

    If dot > (-1.0 + angTol) Then
        Return False
    End If

    ' --- ✅ Use REAL distance (not plane math) ---
    Dim minDist As Double = measureTools.GetMinimumDistance(proxy1, proxy2)

    If minDist > distTol Then
        Return False
    End If

    ' ✅ Final: true opposing + touching
    Return True

End Function
