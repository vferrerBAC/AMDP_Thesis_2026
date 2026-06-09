Imports Inventor
Imports System
Imports System.Collections.Generic
Imports System.Math

Sub Main()

    ' ===== USER INPUT =====
    Dim templatePath As String = "C:\Users\SRosario\OneDrive - BAC\Documents\GitHub\AMDP_Thesis_2026\CostEstimation\cost_calculator - Clean - 12FEB26.xlsx"

    ' Prompt user for JointDetector output file path
    Dim jointDetectorPath As String = InputBox("Enter the file path to the JointDetector output Excel sheet:", "Joint Detector File Path")
    If jointDetectorPath = "" Then
        MsgBox("No file path provided. Exiting.", vbExclamation)
        Exit Sub
    End If

    ' Create a unique filename (timestamp helps avoid overwrite issues)
    Dim outputPath As String = "C:\Users\SRosario\OneDrive - BAC\Desktop\Joint Catalog\CostEstimationOutputs\Output_" & _
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

    ' ==== Assembly Setup ====

    ' Get the active document and verify it's an assembly
    Dim oDoc As Document
    oDoc = ThisApplication.ActiveDocument

    If oDoc.DocumentType <> kAssemblyDocumentObject Then
        MsgBox("Please open an Assembly document first.", vbExclamation, "Wrong Document Type")
        Exit Sub
    End If

    Dim oAsmDoc As AssemblyDocument
    oAsmDoc = oDoc

    ' Collection to track already-processed part documents
    Dim oProcessedDocs As New Collection

    ' Iterate through each component occurrence in the assembly
    Dim oOccurrence As ComponentOccurrence
    For Each oOccurrence In oAsmDoc.ComponentDefinition.Occurrences
        ' Get the referenced document for this occurrence
        Dim oPartDoc As Document
        oPartDoc = oOccurrence.Definition.Document

        ' Only process Part documents (skip sub-assemblies)
        If oPartDoc.DocumentType = kPartDocumentObject Then
            ' Check if this part document has already been processed
            Dim bAlreadyProcessed As Boolean
            bAlreadyProcessed = False
            Dim oCheckedDoc As Document
            For Each oCheckedDoc In oProcessedDocs
                If oCheckedDoc.FullDocumentName = oPartDoc.FullDocumentName Then
                    bAlreadyProcessed = True
                    Exit For
                End If
            Next oCheckedDoc

            If Not bAlreadyProcessed Then

                ' Mark this document as processed
                oProcessedDocs.Add(oPartDoc)
                
                ' Access the Design Tracking iProperties
                Dim oDesignTrackingProps As PropertySet = oPartDoc.PropertySets.Item("Design Tracking Properties")

                ' Get the Part Number property
                Dim partIdentifier As String = oDesignTrackingProps.Item("Part Number").Value

                Dim ncxMaterial As String = GetCustomiProp(oPartDoc, "NCx_Material")
                Dim gaugeValue As String = GetCustomiProp(oPartDoc, "Gauge")
                Dim costDataPierceCount As String = GetCustomiProp(oPartDoc, "CostDataPierceCount")
                Dim costDataCutDistanceInches As String = GetCustomiProp(oPartDoc, "CostDataCutDistanceInches")
                Dim costDataUniqueBends As String = GetCustomiProp(oPartDoc, "CostDataUniqueBends")
                Dim cornerWeld As String = GetCustomiProp(oPartDoc, "Corner Weld")
                Dim costDataFlatLengthInches As String = GetCustomiProp(oPartDoc, "CostDataFlatLengthInches")
                Dim costDataFlatWidthInches As String = GetCustomiProp(oPartDoc, "CostDataFlatWidthInches")
                Dim costDataAssemblyCategory As String = GetCustomiProp(oPartDoc, "CostDataAssemblyCategory")

                ' --- Use values ---
                sheet.Cells(row, PART_SET).Value = oAsmDoc.DisplayName
                sheet.Cells(row, PART_IDENTIFIER).Value = partIdentifier
                ' sheet.Cells(row, PART_QUANTITY).Value = partQuantity

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
            End If
        End If
    Next
    
    sheet = workbook.Sheets("Summary")
    sheet.Cells(2, 1).Value = oAsmDoc.DisplayName

    sheet = workbook.Sheets("Joints List")

    Dim jointOutputWorkbook As Object = excelApp.Workbooks.Open(jointDetectorPath)
    Dim jointsSheet As Object = jointOutputWorkbook.Sheets("Joints")

    row = 4
    Dim jointsSheetRow As Integer = 1
    Const MEMBER1 As Integer = 5
    Const MEMBER2 As Integer = 7 
    CONST JOINT_DISTANCE As Integer = 9

    Dim reachedLastRow As Boolean = False
    While Not reachedLastRow
        Dim member1Value As Object = jointsSheet.Cells(jointsSheetRow, MEMBER1).Value
        Dim member2Value As Object = jointsSheet.Cells(jointsSheetRow, MEMBER2).Value
        Dim jointDistanceValue As Object = jointsSheet.Cells(jointsSheetRow, JOINT_DISTANCE).Value

        If member1Value Is Nothing And member2Value Is Nothing And jointDistanceValue Is Nothing Then
            reachedLastRow = True
        Else
            sheet.Cells(row, 1).Value = member1Value.Split("."c)(0)
            sheet.Cells(row, 2).Value = member2Value.Split("."c)(0)
            sheet.Cells(row, 3).Value = TryParseDoubleSafe(jointDistanceValue)
            row += 1
            jointsSheetRow += 1
        End If
    End While
    ' ===== SAVE + CLEANUP =====
    jointOutputWorkbook.Save()
    jointOutputWorkbook.Close()
    workbook.Save()
    workbook.Close()
    excelApp.Quit()

    workbook = Nothing
    excelApp = Nothing

    MsgBox("New file created:" & vbCrLf & outputPath)

End Sub

Function GetCustomiProp(oPartDoc As Document, propName As String) As String
    ' Access the User Defined iProperties for custom properties
    Dim value As String
    value = ""
    Try
        Dim oUserProps As PropertySet
        oUserProps = oPartDoc.PropertySets.Item("User Defined Properties")
        value = oUserProps.Item(propName).Value
        return value
    Catch
        MsgBox("Warning: Custom property '" & propName & "' not found in document '" & oPartDoc.DisplayName & "'. Defaulting to 1.")
        return "1"
    End Try

    ' On Error Resume Next
    ' value = ""
    ' Dim oUserProps As PropertySet
    ' oUserProps = oPartDoc.PropertySets.Item("User Defined Properties")
    ' value = oUserProps.Item(propName).Value
    ' On Error GoTo 0

    ' If value = "" Then value = "(not set)"
End Function

Function TryParseDoubleSafe(input As String) As Double
    Dim result As Double
    If Double.TryParse(input, result) Then
        Return result
    Else
        Return 1.0
    End If
End Function