# Auto-generated starter conversion from VB to Python
# REVIEW REQUIRED: this is a mechanical translation scaffold, not a guaranteed
# drop-in replacement for Autodesk Inventor.
import win32com.client
from math import *

# Imports Inventor
# Imports System
# Imports System.Collections.Generic
# Imports System.Math

# # ==================================================================================
# # COMBINED SCRIPT: Joint Locator + Cost Estimation
# # ----------------------------------------------------------------------------------
# # This merges JointLocatorV12.vb and CostEstimationV2.vb into a single macro.
# # Joint data is computed in memory and written directly into the cost-estimation
# # workbook's "Joints List" sheet (and a detailed "Joints" + "DebugLog" sheet, for
# # parity with the old standalone tool) — there is no longer an intermediate Excel
# # file that gets exported by one script and re-imported by the other.
# # ==================================================================================

# Sub Main()

# # ==========================================================
# # ✅ Inventor / Assembly setup
# # ==========================================================
#     invApp = None  # As Inventor.Application = ThisApplication
#     oDoc = None  # As Document = invApp.ActiveDocument

#     If oDoc Is Nothing OrElse oDoc.DocumentType <> kAssemblyDocumentObject :
#         MsgBox("Please open an Assembly document first.", MsgBoxStyle.Exclamation, "Wrong Document Type")
#         Exit Sub
#     End If

#     asmDoc = None  # As AssemblyDocument = oDoc
#     compDef = None  # As AssemblyComponentDefinition = asmDoc.ComponentDefinition

# # ==========================================================
# # ✅ USER INPUT — only the cost template is needed now.
# #    (No more InputBox asking for a JointDetector output file.)
# # ==========================================================
#     templatePath = None  # As String = "C:\Users\SRosario\OneDrive - BAC\Documents\GitHub\AMDP_Thesis_2026\CostEstimation\cost_calculator - Clean - 12FEB26.xlsx"

#     outputPath = None  # As String = "C:\Users\SRosario\OneDrive - BAC\Desktop\Joint Catalog\CostEstimationOutputs\Output_" & _
#         Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"

# # Column mapping (BAC Part List sheet)
#     Const PART_SET As Integer = 1
#     Const PART_IDENTIFIER As Integer = 2
#     Const PART_QUANTITY As Integer = 3
#     Const NCX_MATERIAL As Integer = 6
#     Const GAUGE As Integer = 7
#     Const COST_DATA_PIERCE_COUNT As Integer = 8
#     Const COST_DATA_CUT_DISTANCE_INCHES As Integer = 9
#     Const COST_DATA_UNIQUE_BENDS As Integer = 10
#     Const CORNER_WELD As Integer = 11
#     Const COST_DATA_FLAT_LENGTH_INCHES As Integer = 12
#     Const COST_DATA_FLAT_WIDTH_INCHES As Integer = 13
#     Const COST_DATA_ASSEMBLY_CATEGORY As Integer = 14

# # ==========================================================
# # ✅ STEP 1: Copy template, open Excel
# # ==========================================================
#     System.IO.File.Copy(templatePath, outputPath, True)

#     excelApp = None  # As Object = CreateObject("Excel.Application")
#     excelApp.Visible = False
#     excelApp.ScreenUpdating = False

#     workbook = None  # As Object = excelApp.Workbooks.Open(outputPath)
#     sheetBACPartList = None  # As Object = workbook.Sheets("BAC Part List")
#     sheetSummary = None  # As Object = workbook.Sheets("Summary")
#     sheetJointsList = None  # As Object = workbook.Sheets("Joints List")

# # Extra sheets kept from the old JointLocator output, for reference/debugging.
# # (Not strictly required by the cost calculator, but preserves old functionality.)
#     jointsDetailSheet = None  # As Object = workbook.Sheets.Add
#     jointsDetailSheet.Name = "Joints"
#     jointsDetailSheet.Cells(1, 1).Value = "Joint ID"
#     jointsDetailSheet.Cells(1, 2).Value = "X"
#     jointsDetailSheet.Cells(1, 3).Value = "Y"
#     jointsDetailSheet.Cells(1, 4).Value = "Z"
#     jointsDetailSheet.Cells(1, 5).Value = "Part 1"
#     jointsDetailSheet.Cells(1, 6).Value = "Material 1"
#     jointsDetailSheet.Cells(1, 7).Value = "Part 2"
#     jointsDetailSheet.Cells(1, 8).Value = "Material 2"
#     jointsDetailSheet.Cells(1, 9).Value = "Joint Length (in.)"
#     jointsDetailSheet.Cells(1, 10).Value = "Welded Joint Length (in.)"

#     debugSheet = None  # As Object = workbook.Sheets.Add
#     debugSheet.Name = "DebugLog"
#     debugSheet.Cells(1, 1).Value = "Part 1"
#     debugSheet.Cells(1, 2).Value = "Part 2"
#     debugSheet.Cells(1, 3).Value = "Distance"
#     debugSheet.Cells(1, 4).Value = "Angle"
#     debugSheet.Cells(1, 5).Value = "Normal Check"
#     debugSheet.Cells(1, 6).Value = "Overlap Check"
#     debugSheet.Cells(1, 7).Value = "Result"

# # ==========================================================
# # ✅ Assembly sanity checks (kept from CostEstimationV2)
# # ==========================================================
#     oAsmDoc = None  # As AssemblyDocument = asmDoc
#     row = None  # As Integer = 4

#     partCounts = None  # As New Dictionary(Of String, Integer)
#     partOrder = None  # As New List(Of String)
#     oProcessedDocs = None  # As New Collection

# # ==========================================================
# # ✅ STEP 2: BAC Part List population (from CostEstimationV2)
# # ==========================================================
#     oOccurrence = None  # As ComponentOccurrence
#     For Each oOccurrence In oAsmDoc.ComponentDefinition.Occurrences
#         oPartDoc = None  # As Document
#         oPartDoc = oOccurrence.Definition.Document

# # Only process Part documents (skip sub-assemblies)
#         If oPartDoc.DocumentType = kPartDocumentObject :
#             bAlreadyProcessed = None  # As Boolean
#             bAlreadyProcessed = False
#             oCheckedDoc = None  # As Document
#             For Each oCheckedDoc In oProcessedDocs
#                 If oCheckedDoc.FullDocumentName = oPartDoc.FullDocumentName :
#                     oDesignTrackingProps = None  # As PropertySet = oPartDoc.PropertySets.Item("Design Tracking Properties")
#                     partIdentifier = None  # As String = oDesignTrackingProps.Item("Part Number").Value
#                     If partCounts.ContainsKey(partIdentifier) :
#                         partCounts(partIdentifier) += 1
#                     End If
#                     bAlreadyProcessed = True
#                     Exit For
#                 End If
#             Next oCheckedDoc

#             If Not bAlreadyProcessed :

#                 oProcessedDocs.Add(oPartDoc)

#                 oDesignTrackingProps = None  # As PropertySet = oPartDoc.PropertySets.Item("Design Tracking Properties")
#                 partIdentifier = None  # As String = oDesignTrackingProps.Item("Part Number").Value

#                 If Not partCounts.ContainsKey(partIdentifier) :
#                     partCounts.Add(partIdentifier, 0)
#                     partOrder.Add(partIdentifier)
#                 End If
#                 partCounts(partIdentifier) += 1

#                 ncxMaterial = None  # As String = GetCustomiProp(oPartDoc, "NCx_Material")
#                 gaugeValue = None  # As String = GetCustomiProp(oPartDoc, "Gauge")
#                 costDataPierceCount = None  # As String = GetCustomiProp(oPartDoc, "CostDataPierceCount")
#                 costDataCutDistanceInches = None  # As String = GetCustomiProp(oPartDoc, "CostDataCutDistanceInches")
#                 costDataUniqueBends = None  # As String = GetCustomiProp(oPartDoc, "CostDataUniqueBends")
#                 cornerWeld = None  # As String = GetCustomiProp(oPartDoc, "Corner Weld")
#                 costDataFlatLengthInches = None  # As String = GetCustomiProp(oPartDoc, "CostDataFlatLengthInches")
#                 costDataFlatWidthInches = None  # As String = GetCustomiProp(oPartDoc, "CostDataFlatWidthInches")
#                 costDataAssemblyCategory = None  # As String = GetCustomiProp(oPartDoc, "CostDataAssemblyCategory")

#                 sheetBACPartList.Cells(row, PART_SET).Value = oAsmDoc.DisplayName
#                 sheetBACPartList.Cells(row, PART_IDENTIFIER).Value = partIdentifier

#                 sheetBACPartList.Cells(row, NCX_MATERIAL).Value = ncxMaterial
#                 sheetBACPartList.Cells(row, GAUGE).Value = TryParseDoubleSafe(gaugeValue)
#                 sheetBACPartList.Cells(row, COST_DATA_PIERCE_COUNT).Value = TryParseDoubleSafe(costDataPierceCount)
#                 sheetBACPartList.Cells(row, COST_DATA_CUT_DISTANCE_INCHES).Value = TryParseDoubleSafe(costDataCutDistanceInches)
#                 sheetBACPartList.Cells(row, COST_DATA_UNIQUE_BENDS).Value = TryParseDoubleSafe(costDataUniqueBends)
#                 sheetBACPartList.Cells(row, CORNER_WELD).Value = TryParseDoubleSafe(cornerWeld)
#                 sheetBACPartList.Cells(row, COST_DATA_FLAT_LENGTH_INCHES).Value = TryParseDoubleSafe(costDataFlatLengthInches)
#                 sheetBACPartList.Cells(row, COST_DATA_FLAT_WIDTH_INCHES).Value = TryParseDoubleSafe(costDataFlatWidthInches)
#                 sheetBACPartList.Cells(row, COST_DATA_ASSEMBLY_CATEGORY).Value = costDataAssemblyCategory

#                 row += 1
#             End If
#         End If
#     Next

#     quantityRow = None  # As Integer = 4
#     partName = None  # As String
#     For Each partName In partOrder
#         sheetBACPartList.Cells(quantityRow, PART_QUANTITY).Value = partCounts(partName)
#         quantityRow += 1
#     Next partName

#     sheetSummary.Cells(2, 1).Value = oAsmDoc.DisplayName

# # ==========================================================
# # ✅ STEP 3: Joint Detection (from JointLocatorV12), in-memory
# # ==========================================================
#     highlightSet1 = None  # As HighlightSet = asmDoc.CreateHighlightSet()
#     highlightSet2 = None  # As HighlightSet = asmDoc.CreateHighlightSet()

#     highlightSet1.Color = invApp.TransientObjects.CreateColor(255, 0, 0)   ' Red
#     highlightSet2.Color = invApp.TransientObjects.CreateColor(0, 255, 0)   ' Green

#     contactTol = None  # As Double = 0.254   ' 0.254 cm = 0.1 in → adjust as needed for tolerance

#     jointRows = None  # As New List(Of Object())
#     debugRows = None  # As New List(Of Object())
#     jointID = None  # As Integer = 1

#     For i = 1 To compDef.Occurrences.Count - 1

#         occ1 = None  # As ComponentOccurrence = compDef.Occurrences.Item(i)

#         For j = i + 1 To compDef.Occurrences.Count

#             occ2 = None  # As ComponentOccurrence = compDef.Occurrences.Item(j)

#             If Not OccurrencesCouldTouch(occ1, occ2, contactTol) : Continue For

#             If occ1.SurfaceBodies.Count = 0 Or occ2.SurfaceBodies.Count = 0 : Continue For

#             body1 = None  # As SurfaceBody = occ1.SurfaceBodies.Item(1)
#             body2 = None  # As SurfaceBody = occ2.SurfaceBodies.Item(1)

#             part1Name = None  # As String = occ1.Name
#             part2Name = None  # As String = occ2.Name

#             For Each face1 As Face In body1.Faces

#                 For Each face2 As Face In body2.Faces

#                     If Not FaceBoxesCouldTouch(face1, occ1, face2, occ2, contactTol) : Continue For

#                     minAngle = None  # As Double = GetAngle(face1, face2)
#                     minDist = None  # As Double = 999999
#                     normalPass = None  # As Boolean = False
#                     overlapPass = None  # As Boolean = False

#                     result = None  # As String

#                     If minAngle >= contactTol :
#                         result = "FAIL: Angle"
#                     Else
#                         minDist = GetMinDistance_API(face1, face2)

#                         If minDist >= contactTol :
#                             result = "FAIL: Distance"
#                         Else
#                             proxy1 = None  # As FaceProxy
#                             occ1.CreateGeometryProxy(face1, proxy1)
#                             proxy2 = None  # As FaceProxy
#                             occ2.CreateGeometryProxy(face2, proxy2)

#                             normalPass = CheckNormal(proxy1, occ1, proxy2, occ2)

#                             If Not normalPass :
#                                 result = "FAIL: Normal"
#                             Else
#                                 overlapPass = DoFacesOverlap(proxy1, proxy2)

#                                 If Not overlapPass :
#                                     result = "FAIL: No Overlap"
#                                 Else
#                                     result = "PASS"
#                                 End If
#                             End If

#                             If result = "PASS" :
#                                 highlightSet1.AddItem(proxy1)
#                                 highlightSet2.AddItem(proxy2)
#                                 highlightSet1.Clear()
#                                 highlightSet2.Clear()
#                             End If
#                         End If
#                     End If

#                     debugRows.Add(New Object() { _
#                         part1Name, part2Name, _
#                         Round(minDist, 4), Round(minAngle, 4), _
#                         normalPass, overlapPass, result})

#                     If result <> "PASS" : Continue For

#                     centroid = None  # As Point = GetContactCentroid(face1, face2)
#                     jointLength = None  # As Double = GetJointLength(face1, face2) / 2.54   ' cm -> in
#                     weldedJointLength = None  # As Double = GetWeldedJointLength(face1, face2) / 2.54   ' cm -> in

#                     workPt = None  # As WorkPoint = compDef.WorkPoints.AddFixed(centroid)
#                     workPt.Name = "Joint_" & jointID

#                     mat1 = None  # As String = GetMaterial(occ1)
#                     mat2 = None  # As String = GetMaterial(occ2)

#                     jointRows.Add(New Object() { _
#                         jointID, _
#                         Round(centroid.X, 4), Round(centroid.Y, 4), Round(centroid.Z, 4), _
#                         part1Name, mat1, part2Name, mat2, Round(jointLength, 4), Round(weldedJointLength, 4)})

#                     jointID += 1

#                 Next
#             Next
#         Next
#     Next

# # ==========================================================
# # ✅ STEP 4: Write joint results directly into the cost workbook
# #    (this replaces the old "export Joints.xlsx, then re-open it
# #    in CostEstimationV2" round trip)
# # ==========================================================

# # 4a. Full detail + debug log — batch write (kept for reference/debugging)
#     For r As Integer = 0 To jointRows.Count - 1
#         data = None  # As Object() = jointRows(r)
#         For c As Integer = 0 To data.Length - 1
#             jointsDetailSheet.Cells(r + 2, c + 1).Value = data(c)
#         Next
#     Next

#     For r As Integer = 0 To debugRows.Count - 1
#         data = None  # As Object() = debugRows(r)
#         For c As Integer = 0 To data.Length - 1
#             debugSheet.Cells(r + 2, c + 1).Value = data(c)
#         Next
#     Next

#     jointsDetailSheet.Columns.AutoFit
#     debugSheet.Columns.AutoFit

# # 4b. Simplified Part1 / Part2 / Joint Length mapping consumed by the
# #     cost calculator's "Joints List" sheet — written straight from the
# #     in-memory jointRows list instead of re-reading a saved Excel file.
#     jlRow = None  # As Integer = 4
#     For Each jr As Object() In jointRows
#         p1 = None  # As String = CStr(jr(4)).Split("."c)(0)
#         p2 = None  # As String = CStr(jr(6)).Split("."c)(0)
#         jointLenIn = None  # As Double = CDbl(jr(8))

#         sheetJointsList.Cells(jlRow, 1).Value = p1
#         sheetJointsList.Cells(jlRow, 2).Value = p2
#         sheetJointsList.Cells(jlRow, 3).Value = jointLenIn
#         jlRow += 1
#     Next

# # ==========================================================
# # ✅ SAVE + CLEANUP
# # ==========================================================
#     sheetBACPartList.Columns.AutoFit

#     workbook.Save()
#     workbook.Close()
#     excelApp.Quit()

#     workbook = Nothing
#     excelApp = Nothing

#     MsgBox("New file created:" & vbCrLf & outputPath & vbCrLf & _
#            (jointID - 1) & " joint connections found.", MsgBoxStyle.Information)

# End Sub


# # ==========================================================
# # ✅ Custom iProperty lookup (from CostEstimationV2)
# # ==========================================================
# Function GetCustomiProp(oPartDoc As Document, propName As String) As String
#     value = None  # As String
#     value = ""
#     Try
#         oUserProps = None  # As PropertySet
#         oUserProps = oPartDoc.PropertySets.Item("User Defined Properties")
#         value = oUserProps.Item(propName).Value
#         return value
#     Catch
#         MsgBox("Warning: Custom property '" & propName & "' not found in document '" & oPartDoc.DisplayName & "'. Defaulting to 1.")
#         return "1"
#     End Try
# End Function


# Function TryParseDoubleSafe(input As String) As Double
#     result = None  # As Double
#     If Double.TryParse(input, result) :
#         Return result
#     Else
#         Return 1.0
#     End If
# End Function


# # ==========================================================
# # ✅ IMPROVEMENT 1: Occurrence-level bounding box pre-filter
# # ==========================================================
# Function OccurrencesCouldTouch(occ1 As ComponentOccurrence, _
#                                 occ2 As ComponentOccurrence, _
#                                 tol As Double) As Boolean

#     box1 = None  # As Box = occ1.RangeBox
#     box2 = None  # As Box = occ2.RangeBox

#     If box1.MaxPoint.X + tol < box2.MinPoint.X : Return False
#     If box2.MaxPoint.X + tol < box1.MinPoint.X : Return False
#     If box1.MaxPoint.Y + tol < box2.MinPoint.Y : Return False
#     If box2.MaxPoint.Y + tol < box1.MinPoint.Y : Return False
#     If box1.MaxPoint.Z + tol < box2.MinPoint.Z : Return False
#     If box2.MaxPoint.Z + tol < box1.MinPoint.Z : Return False

#     Return True

# End Function


# # ==========================================================
# # ✅ IMPROVEMENT 2: Face-level bounding box pre-filter
# # ==========================================================
# Function FaceBoxesCouldTouch(face1 As Face, occ1 As ComponentOccurrence, _
#                               face2 As Face, occ2 As ComponentOccurrence, _
#                               tol As Double) As Boolean

#     proxy1 = None  # As FaceProxy
#     occ1.CreateGeometryProxy(face1, proxy1)
#     proxy2 = None  # As FaceProxy
#     occ2.CreateGeometryProxy(face2, proxy2)

#     box1 = None  # As Box = proxy1.Evaluator.RangeBox
#     box2 = None  # As Box = proxy2.Evaluator.RangeBox

#     If box1.MaxPoint.X + tol < box2.MinPoint.X : Return False
#     If box2.MaxPoint.X + tol < box1.MinPoint.X : Return False
#     If box1.MaxPoint.Y + tol < box2.MinPoint.Y : Return False
#     If box2.MaxPoint.Y + tol < box1.MinPoint.Y : Return False
#     If box1.MaxPoint.Z + tol < box2.MinPoint.Z : Return False
#     If box2.MaxPoint.Z + tol < box1.MinPoint.Z : Return False

#     Return True

# End Function


# # ==========================================================
# # ✅ Distance
# # ==========================================================
# Function GetMinDistance_API(face1 As Face, face2 As Face) As Double
#     Try
#         Return ThisApplication.MeasureTools.GetMinimumDistance(face1, face2)
#     Catch
#         Return 999999
#     End Try
# End Function


# # ==========================================================
# # ✅ Angle
# # ==========================================================
# Function GetAngle(face1 As Face, face2 As Face) As Double
#     Try
#         Return ThisApplication.MeasureTools.GetAngle(face1, face2)
#     Catch
#         Return -1
#     End Try
# End Function


# # ==========================================================
# # ✅ Material
# # ==========================================================
# Function GetMaterial(occ As ComponentOccurrence) As String
#     Try
#         If TypeOf occ.Definition.Document Is PartDocument :
#             partDoc = None  # As PartDocument = occ.Definition.Document
#             Return partDoc.ComponentDefinition.Material.Name
#         Else
#             Return "Assembly"
#         End If
#     Catch
#         Return "N/A"
#     End Try
# End Function


# # ==========================================================
# # ✅ IMPROVEMENT 4: Normal Check — accepts pre-built proxies
# # ==========================================================
# Function CheckNormal(proxy1 As FaceProxy, occ1 As ComponentOccurrence, _
#                      proxy2 As FaceProxy, occ2 As ComponentOccurrence) As Boolean

#     angTol = None  # As Double = 1E-3
#     tg = None  # As TransientGeometry = ThisApplication.TransientGeometry
#     mt = None  # As MeasureTools = ThisApplication.MeasureTools

#     If proxy1.SurfaceType <> kPlaneSurface Or proxy2.SurfaceType <> kPlaneSurface :
#         Return False
#     End If

#     eval1 = None  # As SurfaceEvaluator = proxy1.Evaluator
#     eval2 = None  # As SurfaceEvaluator = proxy2.Evaluator

#     Dim center1(1) As Double
#     center1(0) = (eval1.ParamRangeRect.MinPoint.X + eval1.ParamRangeRect.MaxPoint.X) / 2
#     center1(1) = (eval1.ParamRangeRect.MinPoint.Y + eval1.ParamRangeRect.MaxPoint.Y) / 2

#     Dim center2(1) As Double
#     center2(0) = (eval2.ParamRangeRect.MinPoint.X + eval2.ParamRangeRect.MaxPoint.X) / 2
#     center2(1) = (eval2.ParamRangeRect.MinPoint.Y + eval2.ParamRangeRect.MaxPoint.Y) / 2

#     Dim n1(2) As Double
#     Dim n2(2) As Double

#     eval1.GetNormal(center1, n1)
#     eval2.GetNormal(center2, n2)

#     v1 = None  # As Vector = tg.CreateVector(n1(0), n1(1), n1(2))
#     v2 = None  # As Vector = tg.CreateVector(n2(0), n2(1), n2(2))

#     v1.Normalize()
#     v2.Normalize()

#     If proxy1.IsParamReversed : v1.ScaleBy(-1)
#     If proxy2.IsParamReversed : v2.ScaleBy(-1)

#     eps = None  # As Double = 0.01

#     Dim pt1Arr(2) As Double
#     Dim pt2Arr(2) As Double

#     eval1.GetPointAtParam(center1, pt1Arr)
#     eval2.GetPointAtParam(center2, pt2Arr)

#     p1 = None  # As Point = tg.CreatePoint(pt1Arr(0), pt1Arr(1), pt1Arr(2))
#     p2 = None  # As Point = tg.CreatePoint(pt2Arr(0), pt2Arr(1), pt2Arr(2))

#     p1Forward = None  # As Point = tg.CreatePoint( _
#         p1.X + v1.X * eps, p1.Y + v1.Y * eps, p1.Z + v1.Z * eps)
#     p1Backward = None  # As Point = tg.CreatePoint( _
#         p1.X - v1.X * eps, p1.Y - v1.Y * eps, p1.Z - v1.Z * eps)

#     dFwd1 = None  # As Double = mt.GetMinimumDistance(p1Forward, occ1)
#     dBack1 = None  # As Double = mt.GetMinimumDistance(p1Backward, occ1)

#     If dFwd1 < dBack1 : v1.ScaleBy(-1)

#     p2Forward = None  # As Point = tg.CreatePoint( _
#         p2.X + v2.X * eps, p2.Y + v2.Y * eps, p2.Z + v2.Z * eps)
#     p2Backward = None  # As Point = tg.CreatePoint( _
#         p2.X - v2.X * eps, p2.Y - v2.Y * eps, p2.Z - v2.Z * eps)

#     dFwd2 = None  # As Double = mt.GetMinimumDistance(p2Forward, occ2)
#     dBack2 = None  # As Double = mt.GetMinimumDistance(p2Backward, occ2)

#     If dFwd2 < dBack2 : v2.ScaleBy(-1)

#     dotRaw = None  # As Double = v1.DotProduct(v2)
#     Return dotRaw < (-1.0 + angTol)

# End Function


# # ==========================================================
# # ✅ IMPROVEMENT 4: Overlap Check — accepts pre-built proxies
# # ==========================================================
# Function DoFacesOverlap(proxy1 As FaceProxy, proxy2 As FaceProxy) As Boolean

#     box1 = None  # As Box = proxy1.Evaluator.RangeBox
#     box2 = None  # As Box = proxy2.Evaluator.RangeBox

#     tol = None  # As Double = 0.001

#     overlapX = None  # As Boolean = (box1.MinPoint.X <= box2.MaxPoint.X + tol) And (box1.MaxPoint.X + tol >= box2.MinPoint.X)
#     overlapY = None  # As Boolean = (box1.MinPoint.Y <= box2.MaxPoint.Y + tol) And (box1.MaxPoint.Y + tol >= box2.MinPoint.Y)
#     overlapZ = None  # As Boolean = (box1.MinPoint.Z <= box2.MaxPoint.Z + tol) And (box1.MaxPoint.Z + tol >= box2.MinPoint.Z)

#     spanX = None  # As Double = Math.Abs(box1.MaxPoint.X - box1.MinPoint.X)
#     spanY = None  # As Double = Math.Abs(box1.MaxPoint.Y - box1.MinPoint.Y)
#     spanZ = None  # As Double = Math.Abs(box1.MaxPoint.Z - box1.MinPoint.Z)

#     minSpan = None  # As Double = Math.Min(spanX, Math.Min(spanY, spanZ))
#     flatTol = None  # As Double = 0.001

#     If Math.Abs(spanX - minSpan) < flatTol :
#         Return overlapY And overlapZ
#     ElseIf Math.Abs(spanY - minSpan) < flatTol :
#         Return overlapX And overlapZ
#     Else
#         Return overlapX And overlapY
#     End If

# End Function


# # ==========================================================
# # ✅ IMPROVEMENT 4: Centroid — accepts pre-built proxies
# #    Uses 2D plane projection for robustness with tilted faces
# # ==========================================================
# Function GetContactCentroid(face1 As Face, face2 As Face) As Point

#     tg = None  # As TransientGeometry = ThisApplication.TransientGeometry

#     If face1.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Or
#        face2.SurfaceType <> SurfaceTypeEnum.kPlaneSurface :
#         MessageBox.Show("Both faces must be planar.")
#         Exit Function
#     End If

#     obb1 = None  # As FaceOBB = BuildFaceOBB(face1, tg)
#     obb2 = None  # As FaceOBB = BuildFaceOBB(face2, tg)

#     If obb1 Is Nothing Or obb2 Is Nothing :
#         MessageBox.Show("Could not build OBB for one or both faces.")
#         Exit Function
#     End If

#     poly1 = None  # As List(Of Double()) = To2D(obb1.corners, obb1, False, tg)
#     poly2 = None  # As List(Of Double()) = To2D(obb2.corners, obb1, True, tg)

#     EnsureCCW(poly1)
#     EnsureCCW(poly2)

#     clipped = None  # As List(Of Double()) = ClipPolygon(poly1, poly2)
#     centroid2D = None  # As Double() = ComputePolygonCentroid(clipped)

#     If centroid2D IsNot Nothing :
#         return LocalToWorld(obb1.origin, obb1.xAxis, obb1.yAxis, centroid2D(0), centroid2D(1), tg)
#     End If

# End Function


# Class FaceOBB
#     Public origin As Point
#     Public xAxis As Vector
#     Public yAxis As Vector
#     Public normal As Vector
#     Public corners As New List(Of Point)
# End Class


# Function BuildFaceOBB(f As Face, tg As TransientGeometry) As FaceOBB

#     data = None  # As New FaceOBB()
#     plane = None  # As Plane = f.Geometry
#     n = None  # As UnitVector = plane.Normal
#     data.normal = tg.CreateVector(n.X, n.Y, n.Z)
#     data.origin = GetFaceCentroid(f, tg)

#     longestEdge = None  # As Edge = Nothing
#     maxLen = None  # As Double = 0
#     For Each e As Edge In f.Edges
#         L = None  # As Double = e.StartVertex.Point.DistanceTo(e.StopVertex.Point)
#         If L > maxLen : maxLen = L : longestEdge = e
#     Next
#     If longestEdge Is Nothing : Return Nothing

#     xAxis = None  # As Vector = longestEdge.StartVertex.Point.VectorTo(longestEdge.StopVertex.Point)
#     xAxis.Normalize()
#     yAxis = None  # As Vector = data.normal.CrossProduct(xAxis)
#     yAxis.Normalize()
#     xAxis = yAxis.CrossProduct(data.normal)
#     xAxis.Normalize()
#     data.xAxis = xAxis
#     data.yAxis = yAxis

#     minX = None  # As Double = Double.MaxValue, minY As Double = Double.MaxValue
#     maxX = None  # As Double = Double.MinValue, maxY As Double = Double.MinValue

#     For Each v As Vertex In f.Vertices
#         vec = None  # As Vector = data.origin.VectorTo(v.Point)
#         px = None  # As Double = vec.DotProduct(xAxis)
#         py = None  # As Double = vec.DotProduct(yAxis)
#         If px < minX : minX = px
#         If py < minY : minY = py
#         If px > maxX : maxX = px
#         If py > maxY : maxY = py
#     Next

#     data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, minX, minY, tg))
#     data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, maxX, minY, tg))
#     data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, maxX, maxY, tg))
#     data.corners.Add(LocalToWorld(data.origin, xAxis, yAxis, minX, maxY, tg))

#     Return data

# End Function

# # Converts 3D points to 2D in obb1's frame.
# # If project=True, each point is first dropped onto obb1's plane along its normal.
# Function To2D(pts As List(Of Point), ref As FaceOBB, project As Boolean, tg As TransientGeometry) As List(Of Double())
#     output = None  # As New List(Of Double())
#     For Each p As Point In pts
#         v = None  # As Vector = ref.origin.VectorTo(p)
#         If project :
#             dist = None  # As Double = v.DotProduct(ref.normal)
#             v = tg.CreateVector(v.X - dist * ref.normal.X,
#                                 v.Y - dist * ref.normal.Y,
#                                 v.Z - dist * ref.normal.Z)
#         End If
#         output.Add(New Double() {v.DotProduct(ref.xAxis), v.DotProduct(ref.yAxis)})
#     Next
#     Return output
# End Function


# Function ClipPolygon(subject As List(Of Double()), clipper As List(Of Double())) As List(Of Double())
#     output = None  # As List(Of Double()) = subject
#     For i As Integer = 0 To clipper.Count - 1
#         If output.Count = 0 : Return output
#         input = None  # As List(Of Double()) = output
#         output = New List(Of Double())
#         A = None  # As Double() = clipper(i)
#         B = None  # As Double() = clipper((i + 1) Mod clipper.Count)
#         For j As Integer = 0 To input.Count - 1
#             P = None  # As Double() = input(j)
#             Q = None  # As Double() = input((j + 1) Mod input.Count)
#             Pin = None  # As Boolean = IsInsideEdge(P, A, B)
#             Qin = None  # As Boolean = IsInsideEdge(Q, A, B)
#             If Pin AndAlso Qin :
#                 output.Add(Q)
#             ElseIf Pin :
#                 output.Add(EdgeIntersect(P, Q, A, B))
#             ElseIf Qin :
#                 output.Add(EdgeIntersect(P, Q, A, B))
#                 output.Add(Q)
#             End If
#         Next
#     Next
#     Return output
# End Function


# Function IsInsideEdge(p As Double(), a As Double(), b As Double()) As Boolean
#     Return (b(0) - a(0)) * (p(1) - a(1)) - (b(1) - a(1)) * (p(0) - a(0)) >= 0
# End Function


# Function EdgeIntersect(p As Double(), q As Double(), a As Double(), b As Double()) As Double()
#     A1 = None  # As Double = q(1) - p(1), B1 As Double = p(0) - q(0)
#     C1 = None  # As Double = A1 * p(0) + B1 * p(1)
#     A2 = None  # As Double = b(1) - a(1), B2 As Double = a(0) - b(0)
#     C2 = None  # As Double = A2 * a(0) + B2 * a(1)
#     det = None  # As Double = A1 * B2 - A2 * B1
#     If Math.Abs(det) < 1.0E-10 : Return New Double() {p(0), p(1)}
#     Return New Double() {(B2 * C1 - B1 * C2) / det, (A1 * C2 - A2 * C1) / det}
# End Function


# Function ComputePolygonCentroid(poly As List(Of Double())) As Double()
#     If poly Is Nothing OrElse poly.Count = 0 : Return Nothing
#     If poly.Count = 1 : Return poly(0)
#     If poly.Count = 2 : Return New Double() {(poly(0)(0) + poly(1)(0)) / 2.0,
#                                                  (poly(0)(1) + poly(1)(1)) / 2.0}
#     cx = None  # As Double = 0, cy As Double = 0, area As Double = 0
#     For i As Integer = 0 To poly.Count - 1
#         j = None  # As Integer = (i + 1) Mod poly.Count
#         cross = None  # As Double = poly(i)(0) * poly(j)(1) - poly(j)(0) * poly(i)(1)
#         area += cross
#         cx += (poly(i)(0) + poly(j)(0)) * cross
#         cy += (poly(i)(1) + poly(j)(1)) * cross
#     Next
#     area *= 0.5
#     If Math.Abs(area) < 1.0E-10 : Return New Double() {(poly(0)(0) + poly(1)(0)) / 2.0,
#                                                             (poly(0)(1) + poly(1)(1)) / 2.0}
#     Return New Double() {cx / (6.0 * area), cy / (6.0 * area)}
# End Function


# Function SignedArea(poly As List(Of Double())) As Double
#     area = None  # As Double = 0
#     For i As Integer = 0 To poly.Count - 1
#         j = None  # As Integer = (i + 1) Mod poly.Count
#         area += poly(i)(0) * poly(j)(1) - poly(j)(0) * poly(i)(1)
#     Next
#     Return area * 0.5
# End Function


# Sub EnsureCCW(poly As List(Of Double()))
#     If SignedArea(poly) < 0 : poly.Reverse()
# End Sub


# Function GetFaceCentroid(f As Face, tg As TransientGeometry) As Point
#     sx = None  # As Double = 0, sy As Double = 0, sz As Double = 0, count As Integer = 0
#     For Each v As Vertex In f.Vertices
#         sx += v.Point.X : sy += v.Point.Y : sz += v.Point.Z : count += 1
#     Next
#     Return tg.CreatePoint(sx / count, sy / count, sz / count)
# End Function


# Function LocalToWorld(origin As Point, xAxis As Vector, yAxis As Vector,
#                       lx As Double, ly As Double, tg As TransientGeometry) As Point
#     p = None  # As Point = origin.Copy()
#     vx = None  # As Vector = xAxis.Copy() : vx.ScaleBy(lx)
#     vy = None  # As Vector = yAxis.Copy() : vy.ScaleBy(ly)
#     p.TranslateBy(vx)
#     p.TranslateBy(vy)
#     Return p
# End Function


# # ==========================================================
# # ✅ Longest Dimension of the coincident overlap area
# #    Uses the same 2D projection as GetContactCentroid
# # ==========================================================
# Function GetContactLongestDimension(proxy1 As FaceProxy, proxy2 As FaceProxy) As Double

#     tg = None  # As TransientGeometry = ThisApplication.TransientGeometry

#     eval1 = None  # As SurfaceEvaluator = proxy1.Evaluator
#     eval2 = None  # As SurfaceEvaluator = proxy2.Evaluator

#     Dim centerUV(1) As Double
#     centerUV(0) = (eval1.ParamRangeRect.MinPoint.X + eval1.ParamRangeRect.MaxPoint.X) / 2
#     centerUV(1) = (eval1.ParamRangeRect.MinPoint.Y + eval1.ParamRangeRect.MaxPoint.Y) / 2

#     Dim originArr(2) As Double
#     eval1.GetPointAtParam(centerUV, originArr)
#     origin = None  # As Point = tg.CreatePoint(originArr(0), originArr(1), originArr(2))

#     Dim nArr(2) As Double
#     eval1.GetNormal(centerUV, nArr)
#     planeN = None  # As Vector = tg.CreateVector(nArr(0), nArr(1), nArr(2))
#     planeN.Normalize()

#     worldX = None  # As Vector = tg.CreateVector(1, 0, 0)
#     If Math.Abs(planeN.DotProduct(worldX)) > 0.9 :
#         worldX = tg.CreateVector(0, 1, 0)
#     End If

#     planeX = None  # As Vector = worldX.Copy()
#     proj = None  # As Vector = planeN.Copy()
#     proj.ScaleBy(planeN.DotProduct(planeX))
#     planeX.SubtractVector(proj)
#     planeX.Normalize()

#     planeY = None  # As Vector = planeN.CrossProduct(planeX)
#     planeY.Normalize()

#     pts1 = None  # As New List(Of Double())
#     pts2 = None  # As New List(Of Double())

#     SampleFacePoints(proxy1, origin, planeX, planeY, pts1)
#     SampleFacePoints(proxy2, origin, planeX, planeY, pts2)

#     min1U = None  # As Double, max1U As Double, min1V As Double, max1V As Double
#     min2U = None  # As Double, max2U As Double, min2V As Double, max2V As Double

#     GetBounds2D(pts1, min1U, max1U, min1V, max1V)
#     GetBounds2D(pts2, min2U, max2U, min2V, max2V)

#     overlapMinU = None  # As Double = Math.Max(min1U, min2U)
#     overlapMaxU = None  # As Double = Math.Min(max1U, max2U)
#     overlapMinV = None  # As Double = Math.Max(min1V, min2V)
#     overlapMaxV = None  # As Double = Math.Min(max1V, max2V)

#     If overlapMaxU >= overlapMinU And overlapMaxV >= overlapMinV :
#         dimU = None  # As Double = overlapMaxU - overlapMinU
#         dimV = None  # As Double = overlapMaxV - overlapMinV
#         Return Math.Max(dimU, dimV)
#     Else
#         Return Math.Max(max1U - min1U, max1V - min1V)
#     End If

# End Function


# # ==========================================================
# # ✅ Helper: sample face points and project into 2D plane coords
# # ==========================================================
# Sub SampleFacePoints(proxy As FaceProxy, origin As Point, _
#                      planeX As Vector, planeY As Vector, _
#                      pts As List(Of Double()))

#     eval = None  # As SurfaceEvaluator = proxy.Evaluator
#     rect = None  # As Box2D = eval.ParamRangeRect

#     uMin = None  # As Double = rect.MinPoint.X
#     uMax = None  # As Double = rect.MaxPoint.X
#     vMin = None  # As Double = rect.MinPoint.Y
#     vMax = None  # As Double = rect.MaxPoint.Y

#     steps = None  # As Integer = 5   ' 5×5 grid; increase for curved faces

#     For iu As Integer = 0 To steps
#         For iv As Integer = 0 To steps

#             u = None  # As Double = uMin + (uMax - uMin) * iu / steps
#             v = None  # As Double = vMin + (vMax - vMin) * iv / steps

#             Dim uv(1) As Double
#             uv(0) = u
#             uv(1) = v

#             Try
#                 Dim ptArr(2) As Double
#                 eval.GetPointAtParam(uv, ptArr)

#                 dx = None  # As Double = ptArr(0) - origin.X
#                 dy = None  # As Double = ptArr(1) - origin.Y
#                 dz = None  # As Double = ptArr(2) - origin.Z

#                 pu = None  # As Double = dx * planeX.X + dy * planeX.Y + dz * planeX.Z
#                 pv = None  # As Double = dx * planeY.X + dy * planeY.Y + dz * planeY.Z

#                 pts.Add(New Double() {pu, pv})
#             Catch
# # Skip invalid params (outside face boundary)
#             End Try

#         Next
#     Next

# End Sub


# # ==========================================================
# # ✅ Helper: 2D bounding box from projected point list
# # ==========================================================
# Sub GetBounds2D(pts As List(Of Double()), _
#                 ByRef minU As Double, ByRef maxU As Double, _
#                 ByRef minV As Double, ByRef maxV As Double)

#     minU = Double.MaxValue
#     maxU = Double.MinValue
#     minV = Double.MaxValue
#     maxV = Double.MinValue

#     For Each pt As Double() In pts
#         If pt(0) < minU : minU = pt(0)
#         If pt(0) > maxU : maxU = pt(0)
#         If pt(1) < minV : minV = pt(1)
#         If pt(1) > maxV : maxV = pt(1)
#     Next

# End Sub

# Function GetJointLength(face1 As Face, face2 As Face) As Double
#     tg = None  # As TransientGeometry = ThisApplication.TransientGeometry

#     If face1.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Or
#        face2.SurfaceType <> SurfaceTypeEnum.kPlaneSurface :
#         MessageBox.Show("Both faces must be planar.")
#         Exit Function
#     End If

#     obb1 = None  # As FaceOBB = BuildFaceOBB(face1, tg)
#     obb2 = None  # As FaceOBB = BuildFaceOBB(face2, tg)

#     If obb1 Is Nothing Or obb2 Is Nothing :
#         MessageBox.Show("Could not build OBB for one or both faces.")
#         Exit Function
#     End If

#     poly1 = None  # As List(Of Double()) = To2D(obb1.corners, obb1, False, tg)
#     poly2 = None  # As List(Of Double()) = To2D(obb2.corners, obb1, True, tg)

#     EnsureCCW(poly1)
#     EnsureCCW(poly2)

#     clipped = None  # As List(Of Double()) = ClipPolygon(poly1, poly2)

#     maxDist = None  # As Double = GetMaxDistanceInPolygon(clipped)

#     return maxDist
# End Function

# Function GetMaxDistanceInPolygon(poly As List(Of Double())) As Double
#     If poly Is Nothing OrElse poly.Count < 2 : Return 0.0

#     maxDist = None  # As Double = 0.0

#     For i As Integer = 0 To poly.Count - 1
#         For j As Integer = i + 1 To poly.Count - 1
#             dx = None  # As Double = poly(i)(0) - poly(j)(0)
#             dy = None  # As Double = poly(i)(1) - poly(j)(1)
#             dist = None  # As Double = Math.Sqrt(dx * dx + dy * dy)

#             If dist > maxDist : maxDist = dist
#         Next
#     Next

#     Return maxDist
# End Function

# Function GetMaxDistancePoints(poly As List(Of Double())) As Tuple(Of Double(), Double(), Double)
#     If poly Is Nothing OrElse poly.Count < 2 : Return Nothing

#     maxDist = None  # As Double = 0.0
#     bestP = None  # As Double() = Nothing
#     bestQ = None  # As Double() = Nothing

#     For i As Integer = 0 To poly.Count - 1
#         For j As Integer = i + 1 To poly.Count - 1
#             dx = None  # As Double = poly(i)(0) - poly(j)(0)
#             dy = None  # As Double = poly(i)(1) - poly(j)(1)
#             dist = None  # As Double = Math.Sqrt(dx * dx + dy * dy)

#             If dist > maxDist :
#                 maxDist = dist
#                 bestP = poly(i)
#                 bestQ = poly(j)
#             End If
#         Next
#     Next

#     Return Tuple.Create(bestP, bestQ, maxDist)
# End Function

# Function GetWeldedJointLength(face1 As Face, face2 As Face) As Double

#     tg = None  # As TransientGeometry = ThisApplication.TransientGeometry

#     If face1.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Or _
#        face2.SurfaceType <> SurfaceTypeEnum.kPlaneSurface :
#         MessageBox.Show("Both faces must be planar.")
#         Exit Function
#     End If

#     obb1 = None  # As FaceOBB = BuildFaceOBB(face1, tg)
#     obb2 = None  # As FaceOBB = BuildFaceOBB(face2, tg)

#     If obb1 Is Nothing Or obb2 Is Nothing :
#         MessageBox.Show("Could not build OBB for one or both faces.")
#         Exit Function
#     End If

#     poly1 = None  # As List(Of Double()) = To2D(obb1.corners, obb1, False, tg)
#     poly2 = None  # As List(Of Double()) = To2D(obb2.corners, obb1, True, tg)

#     EnsureCCW(poly1)
#     EnsureCCW(poly2)

#     clipped = None  # As List(Of Double()) = ClipPolygon(poly1, poly2)

#     perimeter = None  # As Double = GetPolygonPerimeter(clipped)

#     Return perimeter

# End Function

# Function GetPolygonPerimeter(poly As List(Of Double())) As Double

#     If poly Is Nothing Or poly.Count < 2 : Return 0

#     perimeter = None  # As Double = 0.0

#     For i As Integer = 0 To poly.Count - 1

#         p1 = None  # As Double() = poly(i)
#         p2 = None  # As Double() = poly((i + 1) Mod poly.Count)

#         dx = None  # As Double = p2(0) - p1(0)
#         dy = None  # As Double = p2(1) - p1(1)

#         perimeter += Math.Sqrt(dx * dx + dy * dy)

#     Next

#     Return perimeter

# End Function