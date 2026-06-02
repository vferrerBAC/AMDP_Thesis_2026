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
Dim centroids As New List(Of Point)
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
						
						Dim coincident As Boolean = CheckNormal(face1, occ1, face2, occ2)
						
'						MsgBox("Coincident: " & coincident)
						
'						MsgBox("Press OK to view next contact pair")
						
						highlightSet1.Clear()
			            highlightSet2.Clear()
						If coincident Then
							distances.Add(minDist)
							Dim centroid As Point = GetContactCentroid(face1, occ1, face2, occ2)
							
							MsgBox("Contact Centroid:" & vbCrLf & _
						       "X = " & Round(centroid.X, 4) & vbCrLf & _
						       "Y = " & Round(centroid.Y, 4) & vbCrLf & _
						       "Z = " & Round(centroid.Z, 4))
							   
							workPt = compDef.WorkPoints.AddFixed(centroid)
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

' Checks the minimum distance between faces
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

' Gets the angle between faces
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

' Helper function to calculate distance between two points
Function Distance(p1() As Double, p2() As Double) As Double
    Return Sqrt((p1(0) - p2(0))^2 + (p1(1) - p2(1))^2 + (p1(2) - p2(2))^2)
End Function

' Checks if two faces are normal to each other (parallel and opposite normal vectors)
Function CheckNormal(face1 As Face, occ1 As ComponentOccurrence, _
                     face2 As Face, occ2 As ComponentOccurrence) As Boolean

    Dim angTol As Double = 1E-3

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

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

    ' --- Evaluators ---
    Dim eval1 As SurfaceEvaluator = proxy1.Evaluator
    Dim eval2 As SurfaceEvaluator = proxy2.Evaluator
	
	' Gets center coordinate of face
	Dim center1(1) As Double
	center1(0) = (eval1.ParamRangeRect.MinPoint.X + eval1.ParamRangeRect.MaxPoint.X) / 2
	center1(1) = (eval1.ParamRangeRect.MinPoint.Y + eval1.ParamRangeRect.MaxPoint.Y) / 2
	Dim center2(1) As Double
	center2(0) = (eval2.ParamRangeRect.MinPoint.X + eval2.ParamRangeRect.MaxPoint.X) / 2
	center2(1) = (eval2.ParamRangeRect.MinPoint.Y + eval2.ParamRangeRect.MaxPoint.Y) / 2

    ' --- Get normals ---
    Dim normal1(2) As Double
    eval1.GetNormal(center1, normal1)

    Dim normal2(2) As Double
    eval2.GetNormal(center2, normal2)

    Dim v1 As Vector = tg.CreateVector(normal1(0), normal1(1), normal1(2))
    Dim v2 As Vector = tg.CreateVector(normal2(0), normal2(1), normal2(2))

    v1.Normalize
    v2.Normalize

    ' --- Handle param reversal ---
    If proxy1.IsParamReversed Then v1.ScaleBy(-1)
    If proxy2.IsParamReversed Then v2.ScaleBy(-1)

    ' ==========================================================
    ' FORCE OUTWARD NORMALS USING CENTROID
    ' ==========================================================

    ' --- Get center of mass (already in assembly space) ---
    Dim cog1 As Point = occ1.MassProperties.CenterOfMass
    Dim cog2 As Point = occ2.MassProperties.CenterOfMass

    ' --- Get actual point on each face ---
    Dim pt1Arr(2) As Double
    eval1.GetPointAtParam(center1, pt1Arr)

    Dim pt2Arr(2) As Double
    eval2.GetPointAtParam(center2, pt2Arr)

    Dim pt1 As Point = tg.CreatePoint(pt1Arr(0), pt1Arr(1), pt1Arr(2))
    Dim pt2 As Point = tg.CreatePoint(pt2Arr(0), pt2Arr(1), pt2Arr(2))

    ' --- Create vectors from centroid → face ---
    Dim vecToFace1 As Vector = tg.CreateVector( _
        pt1.X - cog1.X, _
        pt1.Y - cog1.Y, _
        pt1.Z - cog1.Z)

    Dim vecToFace2 As Vector = tg.CreateVector( _
        pt2.X - cog2.X, _
        pt2.Y - cog2.Y, _
        pt2.Z - cog2.Z)

    vecToFace1.Normalize
    vecToFace2.Normalize

    ' --- Flip normals if they point inward ---
    If v1.DotProduct(vecToFace1) < 0 Then
        v1.ScaleBy(-1)
    End If

    If v2.DotProduct(vecToFace2) < 0 Then
        v2.ScaleBy(-1)
    End If
	
'	MsgBox("Normal Vector 1:" & vbCrLf & _
'       "X1 = " & v1.X & vbCrLf & _
'       "Y1 = " & v1.Y & vbCrLf & _
'       "Z1 = " & v1.Z)
	   
'	MsgBox("Normal Vector 2:" & vbCrLf & _
'       "X2 = " & v2.X & vbCrLf & _
'       "Y2 = " & v2.Y & vbCrLf & _
'       "Z2 = " & v2.Z)
	
    ' ==========================================================
    ' FINAL CHECK: normals should be opposite for contact
    ' ==========================================================

    Dim dotProduct As Double = v1.DotProduct(v2)

    ' Debug (optional)
    ' MsgBox("Dot Product = " & dotProduct)

    If dotProduct < (-1.0 + angTol) Then
        Return True
    End If

    Return False

End Function


Function GetContactCentroid(face1 As Face, occ1 As ComponentOccurrence, _
                             face2 As Face, occ2 As ComponentOccurrence) As Point

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

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

    ' --- Get evaluators ---
    Dim eval1 As SurfaceEvaluator = proxy1.Evaluator
    Dim eval2 As SurfaceEvaluator = proxy2.Evaluator

    ' ==========================================================
    ' ✅ Get bounding boxes (IN ASSEMBLY SPACE)
    ' ==========================================================

    Dim box1 As Box = eval1.RangeBox
    Dim box2 As Box = eval2.RangeBox

    ' ==========================================================
    ' ✅ Compute intersection box
    ' ==========================================================

    Dim minX As Double = Math.Max(box1.MinPoint.X, box2.MinPoint.X)
    Dim minY As Double = Math.Max(box1.MinPoint.Y, box2.MinPoint.Y)
    Dim minZ As Double = Math.Max(box1.MinPoint.Z, box2.MinPoint.Z)

    Dim maxX As Double = Math.Min(box1.MaxPoint.X, box2.MaxPoint.X)
    Dim maxY As Double = Math.Min(box1.MaxPoint.Y, box2.MaxPoint.Y)
    Dim maxZ As Double = Math.Min(box1.MaxPoint.Z, box2.MaxPoint.Z)
	
	' Build intersection box
	

	Dim intersectBox As Box = tg.CreateBox
	
	intersectBox.MinPoint = tg.CreatePoint(minX, minY, minZ)
	intersectBox.MaxPoint = tg.CreatePoint(maxX, maxY, maxZ)


    ' ==========================================================
    ' ✅ Check if they actually intersect
    ' ==========================================================

    If minX > maxX Or minY > maxY Or minZ > maxZ Then
        ' No overlap → fallback (midpoint between faces)
        Dim midPt As Point = tg.CreatePoint( _
            (box1.MinPoint.X + box2.MinPoint.X) / 2, _
            (box1.MinPoint.Y + box2.MinPoint.Y) / 2, _
            (box1.MinPoint.Z + box2.MinPoint.Z) / 2)
        Return midPt
    End If

    ' ==========================================================
    ' ✅ Compute centroid of intersection box
    ' ==========================================================

    Dim centroid As Point = tg.CreatePoint( _
        (minX + maxX) / 2, _
        (minY + maxY) / 2, _
        (minZ + maxZ) / 2)

    Return centroid

End Function