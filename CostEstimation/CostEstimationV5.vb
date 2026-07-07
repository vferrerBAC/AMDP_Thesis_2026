Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Math

' =====================================================================
'  Cost Estimation V3
'  - Sheet-metal parts: read iProperties directly (unchanged behavior).
'  - Non-sheet-metal (tube) parts: measure geometry and derive the six
'    props that aren't populated on tube stock.
'  - Batched Excel writes for speed.
' =====================================================================

' Shared record so the sheet-metal and tube paths write identically.
Public Structure PartCostData
    Public NCxMaterial As String
    Public Gauge As Double
    Public PierceCount As Double
    Public CutDistanceInches As Double
    Public UniqueBends As Double
    Public CornerWeld As Double
    Public FlatLengthInches As Double
    Public FlatWidthInches As Double
    Public AssemblyCategory As String
End Structure

Public Structure TubeGeometry
    Public LongestAxisInches As Double         ' bounding-box longest dim
    Public CrossSectionPerimeter As Double     ' outer loop of ONE end face
    Public EndFaceOuterPerimeterSum As Double  ' outer loops of BOTH end faces
    Public InteriorCutEdgeSum As Double        ' hole/notch edges on outer walls
    Public InteriorCutLoopCount As Integer     ' distinct interior loops on outer walls
End Structure


Sub Main()

    ' ===== USER INPUT =====
    Dim templatePath As String = "C:\Users\SRosario\OneDrive - BAC\Documents\GitHub\AMDP_Thesis_2026\CostEstimation\cost_calculator - Clean - 12FEB26.xlsx"

    Dim jointDetectorPath As String = InputBox( _
        "Enter the file path to the JointDetector output Excel sheet " & _
        "(MAKE SURE TO DELETE ANY QUOTES AT THE BEGINNING AND END OF FILE PATH):", _
        "Joint Detector File Path")
    If jointDetectorPath = "" Then
        MsgBox("No file path provided. Exiting.", vbExclamation)
        Exit Sub
    End If

    Dim outputPath As String = "C:\Users\SRosario\OneDrive - BAC\Desktop\Joint Catalog\CostEstimationOutputs\Output_" & _
        Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"

    ' Column mapping (Part List sheet)
    Const PART_SET As Integer = 1
    Const PART_IDENTIFIER As Integer = 2
    Const PART_QUANTITY As Integer = 3
    Const NCX_MATERIAL As Integer = 6
    Const GAUGE As Integer = 7
    Const COST_DATA_PIERCE_COUNT As Integer = 8
    Const COST_DATA_CUT_DISTANCE_INCHES As Integer = 9
    Const COST_DATA_UNIQUE_BENDS As Integer = 10
    Const CORNER_WELD As Integer = 11
    Const COST_DATA_FLAT_LENGTH_INCHES As Integer = 12
    Const COST_DATA_FLAT_WIDTH_INCHES As Integer = 13
    Const COST_DATA_ASSEMBLY_CATEGORY As Integer = 14
    Const FIRST_DATA_ROW As Integer = 4

    ' ===== VERIFY ACTIVE DOCUMENT IS ASSEMBLY =====
    Dim oAsmDoc As AssemblyDocument
    If ThisApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MsgBox("Please open an Assembly document first.", vbExclamation, "Wrong Document Type")
        Exit Sub
    End If
    oAsmDoc = CType(ThisApplication.ActiveDocument, AssemblyDocument)

    ' ===== COPY TEMPLATE AND OPEN =====
    System.IO.File.Copy(templatePath, outputPath, True)

    Dim excelApp As Object = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.ScreenUpdating = False
    excelApp.DisplayAlerts = False

    Dim workbook As Object = excelApp.Workbooks.Open(outputPath)

    ' Calculation can only be accessed once a workbook is open, AND Excel has
    ' to be in a "ready" state — right after Workbooks.Open it sometimes isn't,
    ' which raises HRESULT 0x800AC472 ("application is busy"). Retry briefly.
    Dim originalCalc As Object = GetCalcWithRetry(excelApp)
    SetCalcWithRetry(excelApp, -4135)   ' xlCalculationManual

    Dim sheetBACPartList As Object = workbook.Sheets("BAC Part List")

    ' ===== WALK ASSEMBLY, BUILD PER-PART DATA IN MEMORY =====
    Dim processedDocs As New Dictionary(Of String, String)   ' FullDocName -> partIdentifier
    Dim partCounts As New Dictionary(Of String, Integer)
    Dim partOrder As New List(Of String)
    Dim partDataByIdentifier As New Dictionary(Of String, PartCostData)

    For Each oOccurrence As ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences
        Dim oPartDoc As Document = oOccurrence.Definition.Document

        ' Skip sub-assemblies
        If oPartDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            Continue For
        End If

        Dim docKey As String = oPartDoc.FullDocumentName

        ' Duplicate? Just bump the count using the cached identifier.
        If processedDocs.ContainsKey(docKey) Then
            Dim cachedId As String = processedDocs(docKey)
            partCounts(cachedId) += 1
            Continue For
        End If

        ' --- First time seeing this part document ---
        Dim oDesignTrackingProps As PropertySet = oPartDoc.PropertySets.Item("Design Tracking Properties")
        Dim partIdentifier As String = oDesignTrackingProps.Item("Part Number").Value.ToString()

        processedDocs.Add(docKey, partIdentifier)

        If Not partCounts.ContainsKey(partIdentifier) Then
            partCounts.Add(partIdentifier, 0)
            partOrder.Add(partIdentifier)
        End If
        partCounts(partIdentifier) += 1

        ' Route: sheet-metal path or tube path.
        ' Capability check rather than type check — we probe for one of the
        ' sheet-metal-only iProperties. If it exists, the sheet-metal path
        ' can read the rest; if not, we compute from geometry. This is
        ' immune to Inventor's SubType quirks (e.g. a part authored from a
        ' sheet metal template but with no SM features still reports as SM).
        Dim probe As String = ""
        Dim data As PartCostData
        If TryGetCustomiProp(oPartDoc, "CostDataFlatLengthInches", probe) Then
            data = BuildSheetMetalData(oPartDoc)
        Else
            data = BuildTubePartData(oPartDoc)
        End If

        partDataByIdentifier(partIdentifier) = data
    Next

    ' ===== BATCH WRITE PART LIST =====
    ' Build a 2D array sized [rows, columns] and drop it into the sheet in a
    ' single Range assignment. Massively faster than per-cell writes.
    Dim rowCount As Integer = partOrder.Count
    If rowCount > 0 Then
        Dim colCount As Integer = COST_DATA_ASSEMBLY_CATEGORY  ' widest column we write
        Dim buffer(rowCount - 1, colCount - 1) As Object
        Dim assemblyName As String = oAsmDoc.DisplayName

        For i As Integer = 0 To rowCount - 1
            Dim pid As String = partOrder(i)
            Dim d As PartCostData = partDataByIdentifier(pid)

            buffer(i, PART_SET - 1) = assemblyName
            buffer(i, PART_IDENTIFIER - 1) = pid
            buffer(i, PART_QUANTITY - 1) = partCounts(pid)
            buffer(i, NCX_MATERIAL - 1) = d.NCxMaterial
            buffer(i, GAUGE - 1) = d.Gauge
            buffer(i, COST_DATA_PIERCE_COUNT - 1) = d.PierceCount
            buffer(i, COST_DATA_CUT_DISTANCE_INCHES - 1) = d.CutDistanceInches
            buffer(i, COST_DATA_UNIQUE_BENDS - 1) = d.UniqueBends
            buffer(i, CORNER_WELD - 1) = d.CornerWeld
            buffer(i, COST_DATA_FLAT_LENGTH_INCHES - 1) = d.FlatLengthInches
            buffer(i, COST_DATA_FLAT_WIDTH_INCHES - 1) = d.FlatWidthInches
            buffer(i, COST_DATA_ASSEMBLY_CATEGORY - 1) = d.AssemblyCategory
        Next

        Dim topLeft As Object = sheetBACPartList.Cells(FIRST_DATA_ROW, 1)
        Dim bottomRight As Object = sheetBACPartList.Cells(FIRST_DATA_ROW + rowCount - 1, colCount)
        Dim writeRange As Object = sheetBACPartList.Range(topLeft, bottomRight)
        writeRange.Value = buffer
    End If

    ShrinkListObject(sheetBACPartList, "PartsListTable", rowCount)

    ' ===== SUMMARY SHEET =====
    Dim sheetSummary As Object = workbook.Sheets("Summary")
    sheetSummary.Cells(2, 1).Value = oAsmDoc.DisplayName

    ShrinkListObject(sheetSummary, "SummaryTable", 1)

    ' ===== JOINTS LIST (from JointDetector workbook) =====
    Dim sheetJointsList As Object = workbook.Sheets("Joints List")
    Dim jointOutputWorkbook As Object = excelApp.Workbooks.Open(jointDetectorPath)
    Dim jointsSheet As Object = jointOutputWorkbook.Sheets("Joints")

    Const MEMBER1 As Integer = 5
    Const MEMBER2 As Integer = 7
    Const JOINT_DISTANCE As Integer = 9
    Const CONNECTION_TYPE As Integer = 11
    Const WELD_CHECK_COL As Integer = 15

    ' Read the joints source in one shot via UsedRange, then process the array.
    ' Per-cell COM reads have the same round-trip cost as writes.
    Dim jointsUsed As Object = jointsSheet.UsedRange.Value
    Dim jointRows As New List(Of Object())

    If jointsUsed IsNot Nothing Then
        Dim jointRowCount As Integer = jointsSheet.UsedRange.Rows.Count
        For r As Integer = 1 To jointRowCount
            Dim m1 As Object = jointsUsed(r, MEMBER1)
            Dim m2 As Object = jointsUsed(r, MEMBER2)
            Dim jd As Object = jointsUsed(r, JOINT_DISTANCE)
            Dim ct As Object = jointsUsed(r, CONNECTION_TYPE)

            If m1 Is Nothing AndAlso m2 Is Nothing AndAlso jd Is Nothing Then
                Exit For
            End If

            Dim rowOut(3) As Object
            rowOut(0) = If(m1 Is Nothing, "", m1.ToString().Split("."c)(0))
            rowOut(1) = If(m2 Is Nothing, "", m2.ToString().Split("."c)(0))
            rowOut(2) = TryParseDoubleSafe(If(jd Is Nothing, "", jd.ToString()))
            ' Weld Check: TRUE iff Connection Type reads as "welded"
            ' (case-insensitive, trimmed).
            Dim ctStr As String = If(ct Is Nothing, "", ct.ToString().Trim().ToLower())
            rowOut(3) = (ctStr = "welded")
            jointRows.Add(rowOut)
        Next
    End If

    If jointRows.Count > 0 Then
        ' Cols A-C: Part A, Part B, Joint Length Inches
        Dim jbuf(jointRows.Count - 1, 2) As Object
        ' Col O: Weld Check
        Dim wbuf(jointRows.Count - 1, 0) As Object
        For i As Integer = 0 To jointRows.Count - 1
            jbuf(i, 0) = jointRows(i)(0)
            jbuf(i, 1) = jointRows(i)(1)
            jbuf(i, 2) = jointRows(i)(2)
            wbuf(i, 0) = jointRows(i)(3)
        Next

        Dim jTL As Object = sheetJointsList.Cells(FIRST_DATA_ROW, 1)
        Dim jBR As Object = sheetJointsList.Cells(FIRST_DATA_ROW + jointRows.Count - 1, 3)
        sheetJointsList.Range(jTL, jBR).Value = jbuf

        Dim wTL As Object = sheetJointsList.Cells(FIRST_DATA_ROW, WELD_CHECK_COL)
        Dim wBR As Object = sheetJointsList.Cells(FIRST_DATA_ROW + jointRows.Count - 1, WELD_CHECK_COL)
        sheetJointsList.Range(wTL, wBR).Value = wbuf
    End If

    ShrinkListObject(sheetJointsList, "JointsListTable", jointRows.Count)

    ' ===== SAVE + CLEANUP =====
    SetCalcWithRetry(excelApp, CInt(originalCalc))   ' restore, forces recalc   ' restore, forces recalc
    jointOutputWorkbook.Close(False)
    workbook.Save()
    workbook.Close()
    excelApp.Quit()

    workbook = Nothing
    jointOutputWorkbook = Nothing
    excelApp = Nothing

    MsgBox("New file created:" & vbCrLf & outputPath)
End Sub


' =====================================================================
'  SHEET METAL PATH
' =====================================================================
Function BuildSheetMetalData(oPartDoc As Document) As PartCostData
    Dim d As New PartCostData
    d.NCxMaterial       = GetRequiredCustomiProp(oPartDoc, "NCx_Material")
    d.Gauge             = TryParseDoubleSafe(GetRequiredCustomiProp(oPartDoc, "Gauge"))
    d.AssemblyCategory  = GetRequiredCustomiProp(oPartDoc, "CostDataAssemblyCategory")
    d.PierceCount       = TryParseDoubleSafe(GetRequiredCustomiProp(oPartDoc, "CostDataPierceCount"))
    d.CutDistanceInches = TryParseDoubleSafe(GetRequiredCustomiProp(oPartDoc, "CostDataCutDistanceInches"))
    d.UniqueBends       = TryParseDoubleSafe(GetRequiredCustomiProp(oPartDoc, "CostDataUniqueBends"))
    d.CornerWeld        = TryParseDoubleSafe(GetRequiredCustomiProp(oPartDoc, "Corner Weld"))
    d.FlatLengthInches  = TryParseDoubleSafe(GetRequiredCustomiProp(oPartDoc, "CostDataFlatLengthInches"))
    d.FlatWidthInches   = TryParseDoubleSafe(GetRequiredCustomiProp(oPartDoc, "CostDataFlatWidthInches"))
    Return d
End Function


' =====================================================================
'  TUBE (NON-SHEET-METAL) PATH
' =====================================================================
Function BuildTubePartData(oPartDoc As Document) As PartCostData
    Dim d As New PartCostData

    ' Props that DO exist on tube parts.
    d.NCxMaterial      = GetRequiredCustomiProp(oPartDoc, "NCx_Material")
    d.Gauge            = TryParseDoubleSafe(GetRequiredCustomiProp(oPartDoc, "Gauge"))
    d.AssemblyCategory = GetRequiredCustomiProp(oPartDoc, "CostDataAssemblyCategory")

    ' One geometry pass, reused below.
    Dim g As TubeGeometry = MeasureTubeGeometry(oPartDoc)

    d.FlatLengthInches  = g.LongestAxisInches
    d.FlatWidthInches   = g.CrossSectionPerimeter
    d.CutDistanceInches = g.EndFaceOuterPerimeterSum + g.InteriorCutEdgeSum
    ' Option B (confirmed): every cut needs a pierce. 2 end pierces + one per interior loop.
    d.PierceCount       = 2 + g.InteriorCutLoopCount
    ' Confirmed: substitute 1 for both on non-sheet-metal parts.
    d.UniqueBends       = 1
    d.CornerWeld        = 1

    Return d
End Function


' ---------------------------------------------------------------------
'  Geometry pass: one walk of the body's faces + edges.
'  Handles rectangular and round tubes, straight or mitered ends.
'  Assumes a single body (SurfaceBodies.Item(1)).
' ---------------------------------------------------------------------
Function MeasureTubeGeometry(oPartDoc As Document) As TubeGeometry
    Const CM_TO_IN As Double = 1.0 / 2.54
    Dim g As New TubeGeometry

    Dim compDef As PartComponentDefinition = CType(oPartDoc.ComponentDefinition, PartComponentDefinition)
    Dim body As SurfaceBody = compDef.SurfaceBodies.Item(1)

    ' --- 1. Longest axis from the range box ---
    Dim rb As Box = body.RangeBox
    Dim dxCm As Double = rb.MaxPoint.X - rb.MinPoint.X
    Dim dyCm As Double = rb.MaxPoint.Y - rb.MinPoint.Y
    Dim dzCm As Double = rb.MaxPoint.Z - rb.MinPoint.Z

    Dim axisIdx As Integer = 0     ' 0=X, 1=Y, 2=Z
    Dim longestCm As Double = dxCm
    If dyCm > longestCm Then axisIdx = 1 : longestCm = dyCm
    If dzCm > longestCm Then axisIdx = 2 : longestCm = dzCm
    g.LongestAxisInches = longestCm * CM_TO_IN

    ' --- 2. Identify the two end faces ---
    ' Miter-safe: from all planar faces, pick the two whose centroids sit at
    ' the extremes of the longest axis. Works for square-cut or angled ends.
    Dim endFaces As Face() = FindEndFaces(body, axisIdx)

    ' --- 3. Cross-section perimeter (outer edge loop of one end face) ---
    If endFaces(0) IsNot Nothing Then
        g.CrossSectionPerimeter = OuterLoopLengthInches(endFaces(0))
    End If

    ' --- 4. Both end faces' outer perimeters ---
    For i As Integer = 0 To 1
        If endFaces(i) IsNot Nothing Then
            g.EndFaceOuterPerimeterSum += OuterLoopLengthInches(endFaces(i))
        End If
    Next

    ' --- 5. Interior cuts: inner edge loops on faces that AREN'T end faces ---
    ' Every hole/notch in a tube wall shows up as a non-outer EdgeLoop on the
    ' wall face(s) it pierces. For each such loop, count 1 pierce and sum
    ' its edge lengths as cut distance.
    Dim endFaceSet As New HashSet(Of Integer)
    For Each ef As Face In endFaces
        If ef IsNot Nothing Then endFaceSet.Add(GetFaceKey(ef))
    Next

    For Each f As Face In body.Faces
        If endFaceSet.Contains(GetFaceKey(f)) Then Continue For

        For Each loopObj As EdgeLoop In f.EdgeLoops
            If Not loopObj.IsOuterEdgeLoop Then
                g.InteriorCutLoopCount += 1
                For Each e As Edge In loopObj.Edges
                    g.InteriorCutEdgeSum += EdgeLengthCm(e) * CM_TO_IN
                Next
            End If
        Next
    Next

    ' Each interior loop is shared by an inner cut face and (typically) an
    ' outer wall face, so it will be visited from both sides. Halve to
    ' avoid double-counting.
    g.InteriorCutEdgeSum = g.InteriorCutEdgeSum / 2.0
    g.InteriorCutLoopCount = CInt(Math.Ceiling(g.InteriorCutLoopCount / 2.0))

    Return g
End Function


' Pick the two planar faces whose centroids are furthest apart along the
' chosen axis. Robust to miter cuts. Returns a length-2 array; entries may
' be Nothing if the body has fewer than 2 planar faces.
Function FindEndFaces(body As SurfaceBody, axisIdx As Integer) As Face()
    Dim result(1) As Face
    Dim minCoord As Double = Double.MaxValue
    Dim maxCoord As Double = Double.MinValue

    For Each f As Face In body.Faces
        If f.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Then Continue For

        Dim c As Point = FaceCentroid(f)
        If c Is Nothing Then Continue For

        Dim coord As Double = AxisComponent(c, axisIdx)
        If coord < minCoord Then
            minCoord = coord
            result(0) = f
        End If
        If coord > maxCoord Then
            maxCoord = coord
            result(1) = f
        End If
    Next

    Return result
End Function


' Sum of the outer edge loop of a face, in inches.
Function OuterLoopLengthInches(f As Face) As Double
    Const CM_TO_IN As Double = 1.0 / 2.54
    For Each loopObj As EdgeLoop In f.EdgeLoops
        If loopObj.IsOuterEdgeLoop Then
            Dim total As Double = 0
            For Each e As Edge In loopObj.Edges
                total += EdgeLengthCm(e)
            Next
            Return total * CM_TO_IN
        End If
    Next
    Return 0.0
End Function


' Edge arc length in cm. Uses the EdgeEvaluator so it works for lines,
' circular arcs, splines — anything Inventor puts on an edge.
Function EdgeLengthCm(e As Edge) As Double
    Try
        Dim minP As Double, maxP As Double
        e.Evaluator.GetParamExtents(minP, maxP)
        Dim length As Double
        e.Evaluator.GetLengthAtParam(minP, maxP, length)
        Return length
    Catch
        Return 0.0
    End Try
End Function


' Approximate centroid via evaluator midpoint. Adequate for end-face ranking.
Function FaceCentroid(f As Face) As Point
    Try
        Dim eval As SurfaceEvaluator = f.Evaluator
        Dim rb As Box = f.Evaluator.RangeBox
        Dim midCm As Point = ThisApplication.TransientGeometry.CreatePoint( _
            (rb.MinPoint.X + rb.MaxPoint.X) / 2.0, _
            (rb.MinPoint.Y + rb.MaxPoint.Y) / 2.0, _
            (rb.MinPoint.Z + rb.MaxPoint.Z) / 2.0)
        Return midCm
    Catch
        Return Nothing
    End Try
End Function


Function AxisComponent(p As Point, axisIdx As Integer) As Double
    Select Case axisIdx
        Case 0 : Return p.X
        Case 1 : Return p.Y
        Case Else : Return p.Z
    End Select
End Function


' Faces don't have a stable numeric ID everywhere in the API, but hashing on
' the underlying .NET object reference works for our in-loop dedup.
Function GetFaceKey(f As Face) As Integer
    Return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(f)
End Function


' =====================================================================
'  iProperty helpers
' =====================================================================

' Loud version: for props that MUST exist on the part. Missing = real bug.
Function GetRequiredCustomiProp(oPartDoc As Document, propName As String) As String
    Try
        Return oPartDoc.PropertySets.Item("User Defined Properties").Item(propName).Value.ToString()
    Catch
        MsgBox("Warning: Required custom property '" & propName & _
               "' not found in '" & oPartDoc.DisplayName & "'. Defaulting to '1'.")
        Return "1"
    End Try
End Function


' Silent version: for props that may legitimately be absent (tube parts).
Function TryGetCustomiProp(oPartDoc As Document, propName As String, ByRef value As String) As Boolean
    Try
        value = oPartDoc.PropertySets.Item("User Defined Properties").Item(propName).Value.ToString()
        Return True
    Catch
        value = ""
        Return False
    End Try
End Function


Function TryParseDoubleSafe(input As String) As Double
    Dim result As Double
    If Double.TryParse(input, result) Then
        Return result
    Else
        Return 1.0
    End If
End Function

' ---------------------------------------------------------------------
'  Shrink a ListObject to exactly `neededDataRows` data rows and wipe
'  any formulas/values in the rows that fall outside the new body.
'  Called with Calculation = Manual so orphaned formulas cost nothing
'  in the moment between resize and clear.
' ---------------------------------------------------------------------
Sub ShrinkListObject(sheet As Object, tableName As String, neededDataRows As Integer)
    ' Excel requires >= 1 data row in a table. Guard against 0 or negative.
    If neededDataRows < 1 Then neededDataRows = 1

    Dim lo As Object = sheet.ListObjects(tableName)
    Dim headerRow  As Integer = lo.HeaderRowRange.Row
    Dim firstCol   As Integer = lo.HeaderRowRange.Column
    Dim colCount   As Integer = lo.ListColumns.Count
    Dim lastCol    As Integer = firstCol + colCount - 1

    ' Table body ends at this absolute sheet row today.
    Dim currentLastRow As Integer = lo.Range.Row + lo.Range.Rows.Count - 1
    Dim newLastRow     As Integer = headerRow + neededDataRows

    If newLastRow >= currentLastRow Then Exit Sub    ' nothing to shrink

    ' Resize the table down. Structured-reference formulas outside the
    ' new bounds become plain cells, then we clear them.
    Dim newTL As Object = sheet.Cells(headerRow, firstCol)
    Dim newBR As Object = sheet.Cells(newLastRow, lastCol)
    lo.Resize(sheet.Range(newTL, newBR))

    Dim clrTL As Object = sheet.Cells(newLastRow + 1, firstCol)
    Dim clrBR As Object = sheet.Cells(currentLastRow, lastCol)
    sheet.Range(clrTL, clrBR).ClearContents()
End Sub

Function GetCalcWithRetry(excelApp As Object) As Object
    For attempt As Integer = 1 To 10
        Try
            Return excelApp.Calculation
        Catch
            System.Threading.Thread.Sleep(150)
        End Try
    Next
    ' Fall back to Automatic (-4105 = xlCalculationAutomatic) so restore still works.
    Return -4105
End Function

Sub SetCalcWithRetry(excelApp As Object, mode As Integer)
    For attempt As Integer = 1 To 10
        Try
            excelApp.Calculation = mode
            Exit Sub
        Catch
            System.Threading.Thread.Sleep(150)
        End Try
    Next
    ' If we still can't set it after 1.5 s, let the exception surface next call.
    excelApp.Calculation = mode
End Sub