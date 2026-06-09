Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Math

Sub Main()

    ' ===== USER INPUT =====
    Dim templatePath As String = "C:\Users\SRosario\OneDrive - BAC\Documents\GitHub\AMDP_Thesis_2026\CostEstimation\cost_calculator - Clean - 12FEB26.xlsx"

    ' Create a unique filename (timestamp helps avoid overwrite issues)
    Dim outputPath As String = "C:\Users\SRosario\OneDrive - BAC\Desktop\Output_" & _
        Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"

    ' Column mapping
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

    ' ======================

    ' ===== STEP 1: COPY TEMPLATE =====
    System.IO.File.Copy(templatePath, outputPath, True)

    ' ===== STEP 2: OPEN COPIED FILE =====
    Dim excelApp As Object = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.ScreenUpdating = False

    Dim workbook As Object = excelApp.Workbooks.Open(outputPath)
    Dim sheet As Object = workbook.Sheets("BAC Part List")

    ' ===== INVENTOR SETUP =====
    Dim asmDoc As AssemblyDocument = CType(ThisApplication.ActiveDocument, AssemblyDocument)
    Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    Dim bomView As BOMView = compDef.BOM.BOMViews.Item(1)

    Dim row As Integer = 4

    Dim processed As New System.Collections.Generic.HashSet(Of String)

     ' ===== LOOP THROUGH PARTS =====
     
    For Each bomRow As BOMRow In bomView.BOMRows

        ' --- Get document ---
        Dim compDefRow As ComponentDefinition = bomRow.ComponentDefinitions.Item(1)
        Dim doc As Document = compDefRow.Document

        ' --- Read "columns" ---
        Dim partSet As String = asmDoc.DisplayName
        Dim partIdentifier As String = GetiProp(doc, "Project", "Part Number")
        Dim partQuantity As String = bomRow.ItemQuantity.ToString()

        Dim ncxMaterial As String = GetiProp(doc, "Custom", "NCx_Material")
        Dim gaugeValue As String = GetiProp(doc, "Custom", "Gauge")
        Dim costDataPierceCount As String = GetiProp(doc, "Custom", "CostDataPierceCount")
        Dim costDataCutDistanceInches As String = GetiProp(doc, "Custom", "CostDataCutDistanceInches")
        Dim costDataUniqueBends As String = GetiProp(doc, "Custom", "CostDataUniqueBends")
        Dim cornerWeld As String = GetiProp(doc, "Custom", "Corner Weld")
        Dim costDataFlatLengthInches As String = GetiProp(doc, "Custom", "CostDataFlatLengthInches")
        Dim costDataFlatWidthInches As String = GetiProp(doc, "Custom", "CostDataFlatWidthInches")
        Dim costDataAssemblyCategory As String = GetiProp(doc, "Custom", "CostDataAssemblyCategory")

        ' --- Use values ---
        sheet.Cells(row, PART_SET).Value = partSet
        sheet.Cells(row, PART_IDENTIFIER).Value = partIdentifier
        sheet.Cells(row, PART_QUANTITY).Value = partQuantity

        sheet.Cells(row, NCX_MATERIAL).Value = ncxMaterial
        sheet.Cells(row, GAUGE).Value = TryParseDoubleSafe(gaugeValue)
        sheet.Cells(row, COST_DATA_PIERCE_COUNT).Value = TryParseDoubleSafe(costDataPierceCount)
        sheet.Cells(row, COST_DATA_CUT_DISTANCE_INCHES).Value = TryParseDoubleSafe(costDataCutDistanceInches)
        sheet.Cells(row, COST_DATA_UNIQUE_BENDS).Value = TryParseDoubleSafe(costDataUniqueBends)
        sheet.Cells(row, CORNER_WELD).Value = TryParseDoubleSafe(cornerWeld)
        sheet.Cells(row, COST_DATA_FLAT_LENGTH_INCHES).Value = TryParseDoubleSafe(costDataFlatLengthInches)
        sheet.Cells(row, COST_DATA_FLAT_WIDTH_INCHES).Value = TryParseDoubleSafe(costDataFlatWidthInches)
        sheet.Cells(row, COST_DATA_ASSEMBLY_CATEGORY).Value = costDataAssemblyCategory

        row += 1

    Next

    ' ===== SAVE + CLEANUP =====
    workbook.Save()
    workbook.Close()
    excelApp.Quit()

    workbook = Nothing
    excelApp = Nothing

    MsgBox("New file created:" & vbCrLf & outputPath)

End Sub

' ===== HELPER FUNCTION =====
Function GetiProp(doc As Document, setName As String, propName As String) As String
    Try
        Dim propSet As PropertySet = doc.PropertySets.Item(setName)
        Dim prop As [Property] = propSet.Item(propName)
        If prop.Value Is Nothing Then Return "1"
        Return prop.Value.ToString()
    Catch ex As Exception
        MsgBox("Property lookup failed for '" & setName & "." & propName & "'" &
           vbCrLf & "Doc=" & doc.DisplayName &
           vbCrLf & "Error=" & ex.Message)
        Return "1"
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