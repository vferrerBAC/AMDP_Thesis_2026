Option Explicit

'======================================================================
' CONNECTIONS OPTIMIZER INTERFACE  --  Joint Catalog Connection Optimizer
' -----------------------------------------------------------------------
' Run SetupIOInterface() ONCE (Alt+F8 -> Run) to build the UI on the
' "Connection Optimizer" tab. After setup, use the buttons on that tab.
'======================================================================
Private Const IO_SH_DATA  As String = "Bolted Vs. Welded"
Private Const IO_SH_BOLT  As String = "Bolted Connections"
Private Const IO_SH_WELD  As String = "Welded Connections"
Private Const IO_SH_IO    As String = "Connection Optimizer"
Private Const IO_SH_LISTS As String = "_JC_Lists"    ' very-hidden helper sheet

' Calculator tabs the "Move Selected" button targets
Private Const IO_SH_BOLTCALC As String = "Bolted Connection Calculator"
Private Const IO_SH_WELDCALC As String = "Welded Connection Calculator"

' Shape-name prefix used for the per-row selection checkboxes
Private Const IO_CB_PREFIX   As String = "chkCand_"

' "Bolted Vs. Welded" summary tab columns
Private Const IO_DC_MAT   As Integer = 1
Private Const IO_DC_BRTHK As Integer = 2
Private Const IO_DC_CHTHK As Integer = 3
Private Const IO_DC_SHR   As Integer = 4
Private Const IO_DC_CST   As Integer = 5
Private Const IO_DC_TYPE  As Integer = 6

' Detail tab columns
Private Const BOLT_MAT_COL  As Integer = 1
Private Const BOLT_SHR_COL  As Integer = 18
Private Const BOLT_LAST_COL As Integer = 22   ' JointCost

Private Const WELD_MAT_COL  As Integer = 1
Private Const WELD_SHR_COL  As Integer = 20
Private Const WELD_LAST_COL As Integer = 22   ' Joint Cost

' Dropdown source ranges on the hidden _JC_Lists helper sheet
Private Const LIST_MAT_RNG   As String = "A1:A4"
Private Const LIST_GAUGE_RNG As String = "B1:B6"

' ---- "Inputs Outputs" sheet layout ----
Private Const IO_RW_TTL    As Long = 1
Private Const IO_RW_BAR    As Long = 2

' OVERALL
Private Const IO_RW_OSEC   As Long = 3
Private Const IO_RW_I1     As Long = 4    ' Min shear
Private Const IO_RW_I2     As Long = 5    ' Min tensile

' BRANCH MEMBER
Private Const IO_RW_BSEC   As Long = 6
Private Const IO_RW_BMAT   As Long = 7
Private Const IO_RW_BGAU   As Long = 8
Private Const IO_RW_BLEN   As Long = 9
Private Const IO_RW_BWID   As Long = 10

' CHORD MEMBER
Private Const IO_RW_CSEC   As Long = 11
Private Const IO_RW_CMAT   As Long = 12
Private Const IO_RW_CGAU   As Long = 13
Private Const IO_RW_CLEN   As Long = 14
Private Const IO_RW_CWID   As Long = 15

' Buttons & spacer
Private Const IO_RW_BTN    As Long = 17   ' row 16 = spacer
Private Const IO_RW_SP1    As Long = 18

' RECOMMENDED CONNECTION
Private Const IO_RW_RSEC   As Long = 19
Private Const IO_RW_RO1    As Long = 20   ' Connection Type
Private Const IO_RW_RO2    As Long = 21   ' Material

' MANUFACTURING RECOMMENDATIONS
Private Const IO_RW_MSEC   As Long = 22
Private Const IO_RW_MBR    As Long = 23   ' Branch member method
Private Const IO_RW_MCH    As Long = 24   ' Chord member method

' Status / message
Private Const IO_RW_MSG    As Long = 26   ' row 25 = spacer

' Candidate parameter list
Private Const IO_RW_LSEC   As Long = 28   ' row 27 = spacer
Private Const IO_RW_LHDR   As Long = 29
Private Const IO_RW_LSTART As Long = 30

' Column anchors (input/output block)
Private Const IO_CL As Long = 2   ' B  -- label
Private Const IO_CV As Long = 5   ' E  -- value / input
Private Const IO_CH As Long = 7   ' G  -- hint
Private Const IO_CE As Long = 8   ' H  -- right edge

'======================================================================
'  SetupIOInterface  --  Run ONCE to (re)build the UI
'======================================================================
Sub SetupIOInterface()
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(IO_SH_IO)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add( _
                 After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = IO_SH_IO
    End If

    ' Make this sheet active so Form Control text/font manipulation works
    ' on re-runs (Shapes.Add can fail with 1004 on an inactive sheet).
    ws.Activate

    ws.Cells.Clear
    Dim sp As Shape
    For Each sp In ws.Shapes: sp.Delete: Next sp

    ' Un-hide columns J and K if a prior version of this module hid them
    ' (dropdown sources have moved to a separate very-hidden sheet).
    ws.Columns("J:K").Hidden = False

    ' Build / refresh the very-hidden helper sheet that holds the
    ' dropdown source lists.
    IO_EnsureListSheet

    '--- Column widths ---
    ws.Columns(1).ColumnWidth = 5   ' left margin / room for row-select checkbox
    ws.Columns(2).ColumnWidth = 34
    ws.Columns(3).ColumnWidth = 2
    ws.Columns(4).ColumnWidth = 2
    ws.Columns(5).ColumnWidth = 28
    ws.Columns(6).ColumnWidth = 2
    ws.Columns(7).ColumnWidth = 32
    ws.Columns(8).ColumnWidth = 3

    '--- Row heights ---
    ws.Rows(IO_RW_TTL).RowHeight = 34
    ws.Rows(IO_RW_BAR).RowHeight = 6
    Dim ir As Variant
    For Each ir In Array(IO_RW_OSEC, IO_RW_BSEC, IO_RW_CSEC, _
                         IO_RW_RSEC, IO_RW_MSEC, IO_RW_LSEC)
        ws.Rows(ir).RowHeight = 22
    Next
    For Each ir In Array(IO_RW_I1, IO_RW_I2, _
                         IO_RW_BMAT, IO_RW_BGAU, IO_RW_BLEN, IO_RW_BWID, _
                         IO_RW_CMAT, IO_RW_CGAU, IO_RW_CLEN, IO_RW_CWID, _
                         IO_RW_RO1, IO_RW_RO2, IO_RW_MBR, IO_RW_MCH)
        ws.Rows(ir).RowHeight = 24
    Next
    ws.Rows(16).RowHeight = 8
    ws.Rows(IO_RW_BTN).RowHeight = 32
    ws.Rows(IO_RW_SP1).RowHeight = 12
    ws.Rows(25).RowHeight = 8
    ws.Rows(IO_RW_MSG).RowHeight = 26
    ws.Rows(27).RowHeight = 12
    ws.Rows(IO_RW_LHDR).RowHeight = 30

    '--- Colour palette ---
    Dim cNavy As Long: cNavy = RGB(28, 78, 148)
    Dim cHdrB As Long: cHdrB = RGB(68, 114, 196)
    Dim cHdrG As Long: cHdrG = RGB(56, 142, 60)
    Dim cHdrO As Long: cHdrO = RGB(180, 95, 6)
    Dim cHdrP As Long: cHdrP = RGB(112, 48, 160)
    Dim cYel  As Long: cYel = RGB(255, 255, 204)
    Dim cGrn  As Long: cGrn = RGB(220, 240, 220)

    ' ================ TITLE ================
    IO_SafeMerge ws, IO_RW_TTL, IO_CL, IO_RW_TTL, IO_CE
    With ws.Cells(IO_RW_TTL, IO_CL)
        .Value = "JOINT CATALOG  --  CONNECTION OPTIMIZER"
        .Font.Bold = True: .Font.Size = 16: .Font.Color = cNavy
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' ================ COLOUR BAR ================
    ws.Range(ws.Cells(IO_RW_BAR, IO_CL), _
             ws.Cells(IO_RW_BAR, IO_CE)).Interior.Color = cNavy

    ' ================ OVERALL SECTION ================
    IO_BuildSectionHeader ws, IO_RW_OSEC, "  OVERALL  (optional)", cHdrB

    IO_BuildInputRow ws, IO_RW_I1, _
        "Minimum Shear Strength (kips):", _
        cYel, "Required for connection search", RGB(180, 0, 0)

    IO_BuildInputRow ws, IO_RW_I2, _
        "Minimum Tensile Strength (kips):", _
        cYel, "Optional  --  bolted only, noted", RGB(100, 100, 100)

    ' ================ BRANCH MEMBER SECTION ================
    IO_BuildSectionHeader ws, IO_RW_BSEC, _
        "  BRANCH MEMBER  (optional  --  fill all 4 to trigger mfg recommendation)", cHdrP

    IO_BuildInputRow ws, IO_RW_BMAT, "Part Material:", cYel, _
        "Select from dropdown", RGB(100, 100, 100)
    IO_AddListValidation ws, IO_RW_BMAT, LIST_MAT_RNG

    IO_BuildInputRow ws, IO_RW_BGAU, "Part Gauge:", cYel, _
        "Select from dropdown", RGB(100, 100, 100)
    IO_AddListValidation ws, IO_RW_BGAU, LIST_GAUGE_RNG

    IO_BuildInputRow ws, IO_RW_BLEN, "Part Flat Length, x (in):", cYel, _
        "Numeric", RGB(100, 100, 100)

    IO_BuildInputRow ws, IO_RW_BWID, "Part Flat Width, y (in):", cYel, _
        "Numeric", RGB(100, 100, 100)

    ' ================ CHORD MEMBER SECTION ================
    IO_BuildSectionHeader ws, IO_RW_CSEC, _
        "  CHORD MEMBER  (optional  --  fill all 4 to trigger mfg recommendation)", cHdrP

    IO_BuildInputRow ws, IO_RW_CMAT, "Part Material:", cYel, _
        "Select from dropdown", RGB(100, 100, 100)
    IO_AddListValidation ws, IO_RW_CMAT, LIST_MAT_RNG

    IO_BuildInputRow ws, IO_RW_CGAU, "Part Gauge:", cYel, _
        "Select from dropdown", RGB(100, 100, 100)
    IO_AddListValidation ws, IO_RW_CGAU, LIST_GAUGE_RNG

    IO_BuildInputRow ws, IO_RW_CLEN, "Part Flat Length, x (in):", cYel, _
        "Numeric", RGB(100, 100, 100)

    IO_BuildInputRow ws, IO_RW_CWID, "Part Flat Width, y (in):", cYel, _
        "Numeric", RGB(100, 100, 100)

    ' ================ BUTTONS ================
    Dim la As Double: la = ws.Cells(IO_RW_BTN, IO_CL).Left + 4
    Dim ta As Double: ta = ws.Cells(IO_RW_BTN, IO_CL).Top + 4

    Dim btnFind As Shape
    Set btnFind = ws.Shapes.AddFormControl(xlButtonControl, la, ta, 200, 24)
    With btnFind
        .Name = "BtnFindOptimal"
        .TextFrame.Characters.Text = "Find Optimal Connection"
        .OnAction = "FindOptimalConnection"
        .TextFrame.Characters.Font.Bold = True
    End With

    Dim btnClr As Shape
    Set btnClr = ws.Shapes.AddFormControl(xlButtonControl, la + 210, ta, 100, 24)
    With btnClr
        .Name = "BtnClearIO"
        .TextFrame.Characters.Text = "Clear All"
        .OnAction = "ClearIO"
        .TextFrame.Characters.Font.Bold = True
    End With

    Dim btnRst As Shape
    Set btnRst = ws.Shapes.AddFormControl(xlButtonControl, la + 320, ta, 130, 24)
    With btnRst
        .Name = "BtnResetDisplay"
        .TextFrame.Characters.Text = "Reset Display"
        .OnAction = "SetupIOInterface"
        .TextFrame.Characters.Font.Bold = True
    End With

    Dim btnMov As Shape
    Set btnMov = ws.Shapes.AddFormControl(xlButtonControl, la + 460, ta, 220, 24)
    With btnMov
        .Name = "BtnMoveSelected"
        .TextFrame.Characters.Text = "Move Selected to Calculator"
        .OnAction = "MoveSelectedToCalculator"
        .TextFrame.Characters.Font.Bold = True
    End With

    ' ================ RECOMMENDED CONNECTION ================
    IO_BuildSectionHeader ws, IO_RW_RSEC, "  RECOMMENDED CONNECTION", cHdrG

    IO_BuildOutputRow ws, IO_RW_RO1, "Connection Type:", cGrn
    IO_BuildOutputRow ws, IO_RW_RO2, "Material:", cGrn

    ' ================ MANUFACTURING RECOMMENDATIONS ================
    IO_BuildSectionHeader ws, IO_RW_MSEC, "  MANUFACTURING RECOMMENDATION", cHdrO

    IO_BuildOutputRow ws, IO_RW_MBR, "Branch Member Method:", cGrn
    IO_BuildOutputRow ws, IO_RW_MCH, "Chord Member Method:", cGrn

    ' ================ STATUS MESSAGE ================
    IO_SafeMerge ws, IO_RW_MSG, IO_CL, IO_RW_MSG, IO_CE
    With ws.Cells(IO_RW_MSG, IO_CL)
        .Value = "Enter inputs above and click  Find Optimal Connection."
        .Font.Italic = True: .Font.Color = RGB(140, 140, 140)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' ================ CANDIDATE LIST SECTION HEADER ================
    IO_SafeMerge ws, IO_RW_LSEC, IO_CL, IO_RW_LSEC, IO_CE
    With ws.Cells(IO_RW_LSEC, IO_CL)
        .Value = "  CANDIDATE JOINT PARAMETERS  (sorted by cost, utilization-colored)"
        .Font.Bold = True: .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(180, 95, 6)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With

    Application.ScreenUpdating = True
    MsgBox "Interface created on the """ & IO_SH_IO & """ tab.", _
           vbInformation, "Setup Complete"
End Sub

'======================================================================
'  FindOptimalConnection  --  assigned to the Find button
'======================================================================
Sub FindOptimalConnection()

    Dim wsIO   As Worksheet
    Dim wsData As Worksheet
    On Error GoTo EH

    Set wsIO = ThisWorkbook.Sheets(IO_SH_IO)
    Application.ScreenUpdating = False

    ' ---------------- STEP 1: Read all inputs ----------------
    Dim minShr As Double, hasShr As Boolean
    Dim minTen As Double, hasTen As Boolean
    With wsIO
        If IO_IsNumericCell(.Cells(IO_RW_I1, IO_CV).Value) Then
            minShr = CDbl(.Cells(IO_RW_I1, IO_CV).Value): hasShr = True
        End If
        If IO_IsNumericCell(.Cells(IO_RW_I2, IO_CV).Value) Then
            minTen = CDbl(.Cells(IO_RW_I2, IO_CV).Value): hasTen = True
        End If
    End With

    Dim brComplete As Boolean: brComplete = IO_MemberComplete(wsIO, IO_RW_BMAT, IO_RW_BGAU, IO_RW_BLEN, IO_RW_BWID)
    Dim chComplete As Boolean: chComplete = IO_MemberComplete(wsIO, IO_RW_CMAT, IO_RW_CGAU, IO_RW_CLEN, IO_RW_CWID)

    ' ---------------- STEP 2: Decide which case applies ----------------
    '   Case A: hasShr=T, no member complete  -> connection results only
    '   Case B: hasShr=T, branch and/or chord complete  -> connection + mfg
    '   Case C: hasShr=F, branch and/or chord complete  -> mfg only
    '   Error : nothing useful provided
    Dim runConnection As Boolean, runMfg As Boolean
    runConnection = hasShr
    runMfg = (brComplete Or chComplete)

    If Not runConnection And Not runMfg Then
        Application.ScreenUpdating = True
        MsgBox "Please enter at least:" & vbCrLf & _
               "  - Minimum Shear Strength (for a connection recommendation), or" & vbCrLf & _
               "  - All 4 fields for the Branch and/or Chord member (for a " & _
               "manufacturing recommendation).", _
               vbExclamation, "Input Required"
        Exit Sub
    End If

    ' ---------------- STEP 3: Clear all outputs before writing ----------------
    IO_ClearConnOutputs wsIO
    IO_ClearMfgOutputs wsIO
    IO_ClearListArea wsIO

    ' ---------------- STEP 4: Connection search (Case A or B) ----------------
    Dim outType As String, outMat As String
    Dim listCount As Long: listCount = 0
    Dim foundConn As Boolean: foundConn = False

    If runConnection Then
        If minShr <= 0 Then
            Application.ScreenUpdating = True
            MsgBox "Minimum Shear Strength must be greater than zero.", _
                   vbExclamation, "Invalid Input"
            Exit Sub
        End If

        On Error Resume Next
        Set wsData = ThisWorkbook.Sheets(IO_SH_DATA)
        On Error GoTo EH

        If wsData Is Nothing Then
            MsgBox "Data sheet not found: """ & IO_SH_DATA & """", _
                   vbCritical, "Sheet Missing"
            Application.ScreenUpdating = True
            Exit Sub
        End If

        Dim lr As Long
        lr = wsData.Cells(wsData.Rows.Count, IO_DC_MAT).End(xlUp).Row

        Dim bestRow  As Long:   bestRow = -1
        Dim bestCost As Double: bestCost = 1E+308
        Dim util     As Double
        Dim rShr     As Double, rCst As Double
        Dim r        As Long

        For r = 2 To lr
            If Not IO_IsNumericCell(wsData.Cells(r, IO_DC_SHR).Value) Then GoTo NextRow
            If Not IO_IsNumericCell(wsData.Cells(r, IO_DC_CST).Value) Then GoTo NextRow

            rShr = CDbl(wsData.Cells(r, IO_DC_SHR).Value)
            rCst = CDbl(wsData.Cells(r, IO_DC_CST).Value)

            If rShr < minShr Then GoTo NextRow
            util = (minShr / rShr) * 100#
            If util < 40# Or util > 60# Then GoTo NextRow

            If rCst < bestCost Then
                bestCost = rCst
                bestRow = r
            End If
NextRow:
        Next r

        If bestRow = -1 Then
            ' No connection found in 40-60 band; still allow mfg block to run
            IO_PaintCell wsIO, IO_RW_RO1, "Not found", RGB(255, 230, 200), False, ""
            IO_PaintCell wsIO, IO_RW_RO2, "Not found", RGB(255, 230, 200), False, ""
        Else
            foundConn = True
            outType = CStr(wsData.Cells(bestRow, IO_DC_TYPE).Value)
            outMat = CStr(wsData.Cells(bestRow, IO_DC_MAT).Value)
            Dim bgRec As Long: bgRec = RGB(220, 240, 220)
            IO_PaintCell wsIO, IO_RW_RO1, outType, bgRec, True, ""
            IO_PaintCell wsIO, IO_RW_RO2, outMat, bgRec, True, ""

            listCount = IO_PopulateDetailList(wsIO, outType, outMat, minShr)
        End If
    End If

    ' ---------------- STEP 5: Manufacturing recommendations (Case B or C) ----------------
    Dim brMethod As String, chMethod As String
    If brComplete Then
        brMethod = IO_MfgMethod( _
            CStr(wsIO.Cells(IO_RW_BMAT, IO_CV).Value), _
            CInt(wsIO.Cells(IO_RW_BGAU, IO_CV).Value), _
            CDbl(wsIO.Cells(IO_RW_BLEN, IO_CV).Value), _
            CDbl(wsIO.Cells(IO_RW_BWID, IO_CV).Value))
        IO_PaintCell wsIO, IO_RW_MBR, brMethod, IO_MfgColor(brMethod), True, ""
    End If

    If chComplete Then
        chMethod = IO_MfgMethod( _
            CStr(wsIO.Cells(IO_RW_CMAT, IO_CV).Value), _
            CInt(wsIO.Cells(IO_RW_CGAU, IO_CV).Value), _
            CDbl(wsIO.Cells(IO_RW_CLEN, IO_CV).Value), _
            CDbl(wsIO.Cells(IO_RW_CWID, IO_CV).Value))
        IO_PaintCell wsIO, IO_RW_MCH, chMethod, IO_MfgColor(chMethod), True, ""
    End If

    ' ---------------- STEP 6: Status message ----------------
    Dim msg As String
    If runConnection And runMfg Then
        If foundConn Then
            msg = ">> Connection: " & outType & " / " & outMat & _
                  "  |  " & listCount & " candidate(s) listed  |  Mfg method shown above."
        Else
            msg = ">> No connection found in 40-60% band for " & _
                  Format(minShr, "0.000") & " kips. Mfg recommendation shown above."
        End If
    ElseIf runConnection Then
        If foundConn Then
            msg = ">> Connection: " & outType & " / " & outMat & _
                  "  |  " & listCount & " candidate(s) listed below."
        Else
            msg = ">> No connection found within the 40-60% utilization band for " & _
                  Format(minShr, "0.000") & " kips."
        End If
    Else
        msg = ">> Manufacturing recommendation generated.  " & _
              "(Enter a Minimum Shear Strength to also get a connection recommendation.)"
    End If
    If hasTen Then msg = msg & "  |  Tensile input: " & Format(minTen, "0.000") & " kips noted"

    IO_SafeMerge wsIO, IO_RW_MSG, IO_CL, IO_RW_MSG, IO_CE
    With wsIO.Cells(IO_RW_MSG, IO_CL)
        .Value = msg
        .Font.Italic = True: .Font.Bold = True
        .Font.Color = RGB(0, 120, 0)
        .HorizontalAlignment = xlCenter
    End With

    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "FindOptimalConnection"
End Sub

'======================================================================
'  ClearIO  --  assigned to the Clear button
'======================================================================
Sub ClearIO()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(IO_SH_IO)

    ' Clear all input cells
    Dim r As Variant
    For Each r In Array(IO_RW_I1, IO_RW_I2, _
                       IO_RW_BMAT, IO_RW_BGAU, IO_RW_BLEN, IO_RW_BWID, _
                       IO_RW_CMAT, IO_RW_CGAU, IO_RW_CLEN, IO_RW_CWID)
        IO_ClearCell ws, CLng(r), IO_CV
    Next r

    IO_ClearConnOutputs ws
    IO_ClearMfgOutputs ws
    IO_ClearListArea ws
    Application.ScreenUpdating = True
End Sub

'======================================================================
'  MoveSelectedToCalculator  --  assigned to the Move Selected button
'
'  For each candidate row whose selection checkbox is ticked, copies
'  the row's values to the appropriate calculator tab, mapping source
'  columns to their corresponding calculator columns. Target tab is
'  chosen based on the current Connection Type recommendation:
'      Bolted  ->  "Bolted Connection Calculator"
'      Welded  ->  "Welded Connection Calculator"
'  New rows are appended below any existing data in column A of the
'  target sheet. Copied checkboxes are un-ticked afterwards.
'======================================================================
Sub MoveSelectedToCalculator()
    Dim wsIO As Worksheet
    On Error GoTo EH
    Set wsIO = ThisWorkbook.Sheets(IO_SH_IO)

    ' Which calculator do we target?
    Dim connType As String
    connType = LCase(Trim(CStr(wsIO.Cells(IO_RW_RO1, IO_CV).Value)))

    Dim calcName As String, srcLastCol As Integer
    Select Case connType
        Case "bolted"
            calcName = IO_SH_BOLTCALC
            srcLastCol = BOLT_LAST_COL
        Case "welded"
            calcName = IO_SH_WELDCALC
            srcLastCol = WELD_LAST_COL
        Case Else
            MsgBox "No connection recommendation is currently displayed." & vbCrLf & _
                   "Run  Find Optimal Connection  first so a Connection Type " & _
                   "(Bolted or Welded) is available.", _
                   vbInformation, "No Recommendation"
            Exit Sub
    End Select

    Dim wsCalc As Worksheet
    On Error Resume Next
    Set wsCalc = ThisWorkbook.Sheets(calcName)
    On Error GoTo EH
    If wsCalc Is Nothing Then
        MsgBox "Calculator sheet not found:" & vbCrLf & "   " & calcName, _
               vbCritical, "Sheet Missing"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Walk the shapes on the IO sheet, look for our selection checkboxes.
    Dim shp As Shape
    Dim checkedCount As Long: checkedCount = 0
    Dim copiedCount As Long: copiedCount = 0

    For Each shp In wsIO.Shapes
        If InStr(shp.Name, IO_CB_PREFIX) = 1 Then
            Dim cbVal As Long: cbVal = 0
            On Error Resume Next
            cbVal = shp.OLEFormat.Object.Value
            On Error GoTo EH
            If cbVal = 1 Then           ' 1 == xlOn == ticked
                checkedCount = checkedCount + 1

                Dim srcRow As Long
                srcRow = shp.TopLeftCell.Row

                ' Append below any existing data in col A. Calculator
                ' sheets now use a two-row header (row 1 = section
                ' labels, row 2 = column headers), so data starts on
                ' row 3. If End(xlUp) somehow returns above row 2,
                ' fall through to row 3.
                Dim tgtRow As Long
                tgtRow = wsCalc.Cells(wsCalc.Rows.Count, 1).End(xlUp).Row + 1
                If tgtRow < 3 Then tgtRow = 3

                Dim srcCol As Integer, tgtCol As Integer
                For srcCol = 1 To srcLastCol
                    tgtCol = IO_MapCalcColumn(connType, srcCol)
                    If tgtCol > 0 Then
                        ' On wsIO the source data starts at column B, i.e.
                        ' source col N lives in wsIO column N + 1.
                        wsCalc.Cells(tgtRow, tgtCol).Value = _
                            wsIO.Cells(srcRow, srcCol + 1).Value
                    End If
                Next srcCol

                copiedCount = copiedCount + 1

                ' Un-tick this checkbox so the user can immediately see
                ' which rows have already been moved.
                On Error Resume Next
                shp.OLEFormat.Object.Value = xlOff
                On Error GoTo EH
            End If
        End If
    Next shp

    Application.ScreenUpdating = True

    If checkedCount = 0 Then
        MsgBox "No candidate rows are selected." & vbCrLf & _
               "Tick the checkbox in column A next to each row you want to send.", _
               vbInformation, "Nothing Selected"
        Exit Sub
    End If

    MsgBox copiedCount & " row(s) copied to  """ & calcName & """.", _
           vbInformation, "Copy Complete"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "MoveSelectedToCalculator"
End Sub

'----------------------------------------------------------------------
'  IO_MapCalcColumn
'  Given a source column from the Bolted or Welded detail tab, returns
'  the matching INPUT column on the corresponding calculator tab, or 0
'  if the source column has no target. Only columns the calculator
'  treats as user inputs are populated -- formula-driven columns on
'  the calculator are left alone so their existing formulas keep
'  computing based on the pasted inputs.
'
'  Calculator layouts start at row 3 (row 1 = section labels, row 2 =
'  column headers). MoveSelectedToCalculator handles the row offset.
'
'  BOLTED calculator inputs (cols 1..6):
'      1 Connection Material   2 Gauge   3 Number of Fasteners
'      4 Bolt Type             5 Bolt Diameter (d)
'      6 Distance to closest end (Le)
'
'  WELDED calculator inputs (cols 1..7):
'      1 Branch Material   2 Branch Gauge   3 Branch Width
'      4 Branch Height     5 Chord Material 6 Chord Gauge
'      7 Electrode Specification
'----------------------------------------------------------------------
Private Function IO_MapCalcColumn(connType As String, srcCol As Integer) As Integer
    Select Case connType
        Case "bolted"
            '  Bolted Connections source layout:
            '   1 ConnectionMaterial       -> calc 1
            '   2 Gauge                    -> calc 2
            '   3 Number of Fasteners      -> calc 3
            '   4 BoltUsed                 -> calc 4  (Bolt Type)
            '   5 Bolt Diameter (d)        -> calc 5
            '   6 Le                       -> calc 6
            '   7..22  formula-driven / skipped
            Select Case srcCol
                Case 1 To 6:    IO_MapCalcColumn = srcCol
                Case Else:      IO_MapCalcColumn = 0
            End Select

        Case "welded"
            '  Welded Connections source layout:
            '   1  Branch Material          -> calc 1
            '   2  Branch Gauge             -> calc 2
            '   3  Branch Thickness         -> skip (calc col 9 is a formula)
            '   4  Branch Width             -> calc 3
            '   5  Branch Height            -> calc 4
            '   6  Branch Shape             -> skip (not in calc)
            '   7  Chord Material           -> calc 5
            '   8  Chord Gauge              -> calc 6
            '   9  Chord Thickness          -> skip (calc col 10 is a formula)
            '   10 Chord Shape              -> skip
            '   11 Electrode Specification  -> calc 7
            '   12..22  formula-driven / skipped
            Select Case srcCol
                Case 1:         IO_MapCalcColumn = 1
                Case 2:         IO_MapCalcColumn = 2
                Case 4:         IO_MapCalcColumn = 3
                Case 5:         IO_MapCalcColumn = 4
                Case 7:         IO_MapCalcColumn = 5
                Case 8:         IO_MapCalcColumn = 6
                Case 11:        IO_MapCalcColumn = 7
                Case Else:      IO_MapCalcColumn = 0
            End Select

        Case Else
            IO_MapCalcColumn = 0
    End Select
End Function

'======================================================================
'  IO_MfgMethod
'  Returns the recommended manufacturing method given the part's
'  material, gauge, flat length (x) and flat width (y).
'
'  Priority when more than one method is feasible:  TL > MPB > APB
'  Carbon/Stainless Steel Tube -> automatic Tube Laser regardless.
'======================================================================
Private Function IO_MfgMethod(material As String, gauge As Integer, _
                              x As Double, y As Double) As String
    Dim mat As String: mat = LCase(Trim(material))

    ' Tube materials always go to Tube Laser
    If InStr(mat, "tube") > 0 Then
        IO_MfgMethod = "Tube Laser (TL)"
        Exit Function
    End If

    Dim isGLV As Boolean: isGLV = (InStr(mat, "galvanized") > 0)
    Dim isSST As Boolean: isSST = (InStr(mat, "stainless") > 0 And InStr(mat, "sheet") > 0)

    If Not (isGLV Or isSST) Then
        IO_MfgMethod = "N/A (unrecognized material)"
        Exit Function
    End If

    ' Universal MPB / TL formulas (no gauge dependence in spec)
    Dim mpbOK As Boolean
    mpbOK = (x < 60 And y < 168) Or (x < 168 And y < 60)

    Dim tlOK As Boolean
    tlOK = (x < 12.5 And y < 334.65) Or (x < 334.65 And y < 12.5)

    ' APB Y-bound varies by gauge & material.
    ' apbY <= 0 means "APB not available for this gauge/material combo".
    Dim apbY As Double: apbY = -1
    Select Case gauge
        Case 10
            If isGLV Then apbY = 118.11
            If isSST Then apbY = 82.67
        Case 12
            If isGLV Then apbY = 149.6
            If isSST Then apbY = 108.26
        Case 14
            If isGLV Then apbY = 149.6
            If isSST Then apbY = 118.11
        Case 16
            If isGLV Then apbY = 149.6
            If isSST Then apbY = 149.6
        ' 8 and 18 gauge: no APB envelope provided -> APB unavailable
    End Select

    Dim apbOK As Boolean: apbOK = False
    If apbY > 0 Then
        Dim rectOK As Boolean, circOK As Boolean
        rectOK = (x > 18.3 And x < 60 And y > 27.75 And y < apbY) Or _
                 (x > 27.75 And x < apbY And y > 18.3 And y < 60)
        circOK = (x * x + y * y) < (157.48 * 157.48)
        apbOK = rectOK And circOK
    End If

    ' Priority: TL > MPB > APB
    If tlOK Then
        IO_MfgMethod = "Tube Laser (TL)"
    ElseIf mpbOK Then
        IO_MfgMethod = "Manual Press Brake (MPB)"
    ElseIf apbOK Then
        IO_MfgMethod = "Automatic Panel Bender (APB)"
    Else
        IO_MfgMethod = "No suitable manufacturing method"
    End If
End Function

'----------------------------------------------------------------------
'  IO_MfgColor  --  background tint for the mfg output cell
'----------------------------------------------------------------------
Private Function IO_MfgColor(method As String) As Long
    Select Case True
        Case InStr(method, "Tube Laser") > 0
            IO_MfgColor = RGB(120, 210, 100)   ' green - best
        Case InStr(method, "Manual Press") > 0
            IO_MfgColor = RGB(180, 220, 250)   ' light blue - second
        Case InStr(method, "Automatic Panel") > 0
            IO_MfgColor = RGB(255, 230, 100)   ' yellow - last resort
        Case Else
            IO_MfgColor = RGB(255, 180, 180)   ' red - none / unrecognized
    End Select
End Function

'======================================================================
'  IO_PopulateDetailList  (unchanged behavior from prior version --
'  filters detail tab by material and minShr, sorts by cost ascending,
'  paints rows by utilization color)
'======================================================================
Private Function IO_PopulateDetailList(wsIO As Worksheet, _
                                       connType As String, _
                                       matName As String, _
                                       minShr As Double) As Long
    Dim wsSrc As Worksheet
    Dim matCol As Integer, shrCol As Integer, lastCol As Integer
    Dim srcName As String

    Select Case LCase(Trim(connType))
        Case "bolted"
            srcName = IO_SH_BOLT
            matCol = BOLT_MAT_COL: shrCol = BOLT_SHR_COL: lastCol = BOLT_LAST_COL
        Case "welded"
            srcName = IO_SH_WELD
            matCol = WELD_MAT_COL: shrCol = WELD_SHR_COL: lastCol = WELD_LAST_COL
        Case Else
            IO_PopulateDetailList = 0: Exit Function
    End Select

    On Error Resume Next
    Set wsSrc = ThisWorkbook.Sheets(srcName)
    On Error GoTo 0
    If wsSrc Is Nothing Then IO_PopulateDetailList = 0: Exit Function

    Dim c As Integer
    For c = 1 To lastCol
        With wsIO.Cells(IO_RW_LHDR, c + 1)
            .Value = wsSrc.Cells(1, c).Value
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(80, 80, 80)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
        End With
    Next c
    With wsIO.Cells(IO_RW_LHDR, lastCol + 2)
        .Value = "Utilization"
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(40, 40, 40)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With

    Dim lastSrcRow As Long
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, matCol).End(xlUp).Row

    Dim cap As Long: cap = lastSrcRow + 5
    Dim arr() As Variant
    ReDim arr(1 To cap, 1 To lastCol + 1)
    Dim cnt As Long: cnt = 0

    Dim r As Long, shrVal As Double
    For r = 2 To lastSrcRow
        If StrComp(CStr(wsSrc.Cells(r, matCol).Value), matName, vbTextCompare) = 0 Then
            If IO_IsNumericCell(wsSrc.Cells(r, shrCol).Value) Then
                shrVal = CDbl(wsSrc.Cells(r, shrCol).Value)
                If shrVal > 0 And shrVal >= minShr Then
                    cnt = cnt + 1
                    For c = 1 To lastCol
                        arr(cnt, c) = wsSrc.Cells(r, c).Value
                    Next c
                    arr(cnt, lastCol + 1) = (minShr / shrVal) * 100#
                End If
            End If
        End If
    Next r

    If cnt = 0 Then IO_PopulateDetailList = 0: Exit Function

    ' Sort by joint cost ascending (cost = source col lastCol)
    IO_SortAsc arr, cnt, lastCol

    Dim writeRow As Long: writeRow = IO_RW_LSTART
    Dim i As Long, bg As Long, util As Double
    For i = 1 To cnt
        util = CDbl(arr(i, lastCol + 1))
        bg = IO_UtilColor(util)
        For c = 1 To lastCol
            With wsIO.Cells(writeRow, c + 1)
                .Value = arr(i, c)
                .Interior.Color = bg
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200): .Borders.Weight = xlThin
                .Font.Color = RGB(0, 0, 0): .Font.Bold = False
            End With
        Next c
        With wsIO.Cells(writeRow, lastCol + 2)
            .Value = util / 100#
            .NumberFormat = "0.0%"
            .Interior.Color = bg: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200): .Borders.Weight = xlThin
        End With

        ' Add a per-row selection checkbox in column A. TopLeftCell.Row
        ' anchors the checkbox to writeRow, so MoveSelectedToCalculator
        ' can read the associated data back off wsIO by that row number.
        Dim cellA As Range: Set cellA = wsIO.Cells(writeRow, 1)
        Dim cbW As Double: cbW = 15
        Dim cbH As Double: cbH = 15
        Dim cbL As Double: cbL = cellA.Left + (cellA.Width - cbW) / 2
        Dim cbT As Double: cbT = cellA.Top + (cellA.Height - cbH) / 2

        Dim cb As Shape
        Set cb = wsIO.Shapes.AddFormControl(xlCheckBox, cbL, cbT, cbW, cbH)
        cb.Name = IO_CB_PREFIX & i
        On Error Resume Next
        cb.OLEFormat.Object.Caption = ""
        cb.OLEFormat.Object.Value = xlOff
        On Error GoTo 0

        writeRow = writeRow + 1
    Next i

    Dim listRng As Range
    Set listRng = wsIO.Range(wsIO.Cells(IO_RW_LHDR, 2), _
                              wsIO.Cells(IO_RW_LHDR, lastCol + 2))
    listRng.EntireColumn.AutoFit
    For c = 2 To lastCol + 2
        If wsIO.Columns(c).ColumnWidth > 22 Then wsIO.Columns(c).ColumnWidth = 22
        If wsIO.Columns(c).ColumnWidth < 10 Then wsIO.Columns(c).ColumnWidth = 10
    Next c

    IO_PopulateDetailList = cnt
End Function

'======================================================================
'  Private helpers
'======================================================================

Private Sub IO_BuildSectionHeader(ws As Worksheet, rw As Long, _
                                   txt As String, bg As Long)
    IO_SafeMerge ws, rw, IO_CL, rw, IO_CE
    With ws.Cells(rw, IO_CL)
        .Value = txt
        .Font.Bold = True: .Font.Size = 11: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = bg
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub IO_BuildInputRow(ws As Worksheet, rw As Long, lbl As String, _
                              bgColor As Long, hint As String, hintColor As Long)
    With ws.Cells(rw, IO_CL)
        .Value = lbl
        .Font.Bold = True: .VerticalAlignment = xlCenter
    End With
    With ws.Cells(rw, IO_CV)
        .Interior.Color = bgColor
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(160, 160, 0): .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With ws.Cells(rw, IO_CH)
        .Value = hint
        .Font.Italic = True: .Font.Color = hintColor: .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub IO_BuildOutputRow(ws As Worksheet, rw As Long, lbl As String, _
                               bgColor As Long)
    With ws.Cells(rw, IO_CL)
        .Value = lbl
        .Font.Bold = True: .VerticalAlignment = xlCenter
    End With
    With ws.Cells(rw, IO_CV)
        .Value = "--"
        .Interior.Color = bgColor
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 130, 50): .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Font.Italic = True: .Font.Color = RGB(150, 150, 150)
    End With
End Sub

Private Sub IO_AddListValidation(ws As Worksheet, rw As Long, rngAddr As String)
    Dim cel As Range
    Set cel = ws.Cells(rw, IO_CV)

    ' Reference cells on the very-hidden _JC_Lists helper sheet.
    ' Using a range reference (rather than a comma-separated literal)
    ' avoids the comma-vs-semicolon locale issue that can make
    ' Validation.Add fail with a 1004 on some regional settings.
    Dim srcFormula As String
    srcFormula = "=" & IO_SH_LISTS & "!" & rngAddr

    On Error Resume Next
    cel.Validation.Delete
    cel.Validation.Add Type:=xlValidateList, _
                       AlertStyle:=xlValidAlertStop, _
                       Formula1:=srcFormula
    With cel.Validation
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Invalid Selection"
        .ErrorMessage = "Please pick a value from the dropdown list."
    End With
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------
'  IO_EnsureListSheet
'  Creates (or refreshes) a helper worksheet that holds the source
'  values for the material and gauge dropdowns, then sets it to
'  xlSheetVeryHidden so it cannot be revealed through the usual
'  right-click Unhide menu.
'----------------------------------------------------------------------
Private Sub IO_EnsureListSheet()
    Dim wsL As Worksheet
    On Error Resume Next
    Set wsL = ThisWorkbook.Sheets(IO_SH_LISTS)
    On Error GoTo 0

    If wsL Is Nothing Then
        ' Add a new sheet without changing which sheet is currently active.
        Set wsL = ThisWorkbook.Sheets.Add
        wsL.Name = IO_SH_LISTS
    Else
        ' Make sure it is visible while we edit its contents.
        wsL.Visible = xlSheetVisible
    End If

    With wsL
        .Cells.Clear
        .Range("A1").Value = "Galvanized Sheet Steel"
        .Range("A2").Value = "Stainless Sheet Steel"
        .Range("A3").Value = "Carbon Steel Tube"
        .Range("A4").Value = "Stainless Steel Tube"
        .Range("B1").Value = 8
        .Range("B2").Value = 10
        .Range("B3").Value = 12
        .Range("B4").Value = 14
        .Range("B5").Value = 16
        .Range("B6").Value = 18
    End With

    ' Move focus back to the main IO sheet before hiding, so we do
    ' not end up with the lists sheet as the last-active sheet.
    On Error Resume Next
    ThisWorkbook.Sheets(IO_SH_IO).Activate
    On Error GoTo 0

    wsL.Visible = xlSheetVeryHidden
End Sub

Private Sub IO_ClearConnOutputs(ws As Worksheet)
    Dim oRows(1 To 2) As Long
    oRows(1) = IO_RW_RO1: oRows(2) = IO_RW_RO2
    Dim i As Integer
    For i = 1 To 2
        With ws.Cells(oRows(i), IO_CV)
            .NumberFormat = "General"
            .Value = "--"
            .Interior.Color = RGB(220, 240, 220)
            .Font.Italic = True: .Font.Bold = False
            .Font.Color = RGB(150, 150, 150)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(0, 130, 50): .Borders.Weight = xlThin
        End With
    Next i

    IO_SafeMerge ws, IO_RW_MSG, IO_CL, IO_RW_MSG, IO_CE
    With ws.Cells(IO_RW_MSG, IO_CL)
        .Value = "Enter inputs above and click  Find Optimal Connection."
        .Font.Italic = True: .Font.Bold = False
        .Font.Color = RGB(140, 140, 140)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub IO_ClearMfgOutputs(ws As Worksheet)
    Dim oRows(1 To 2) As Long
    oRows(1) = IO_RW_MBR: oRows(2) = IO_RW_MCH
    Dim i As Integer
    For i = 1 To 2
        With ws.Cells(oRows(i), IO_CV)
            .NumberFormat = "General"
            .Value = "--"
            .Interior.Color = RGB(220, 240, 220)
            .Font.Italic = True: .Font.Bold = False
            .Font.Color = RGB(150, 150, 150)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(0, 130, 50): .Borders.Weight = xlThin
        End With
    Next i
End Sub

Private Sub IO_ClearListArea(ws As Worksheet)
    ' Delete any candidate-row selection checkboxes from previous runs.
    ' Iterate backwards because deleting shifts the collection indices.
    Dim k As Long
    For k = ws.Shapes.Count To 1 Step -1
        If InStr(ws.Shapes(k).Name, IO_CB_PREFIX) = 1 Then
            ws.Shapes(k).Delete
        End If
    Next k

    Dim lastListCol As Integer
    lastListCol = WorksheetFunction.Max(BOLT_LAST_COL, WELD_LAST_COL) + 2

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < IO_RW_LSTART Then lastRow = IO_RW_LSTART + 50

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(IO_RW_LHDR, 1), _
                       ws.Cells(lastRow + 5, lastListCol + 2))
    With rng
        .Clear
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
    End With
End Sub

Private Sub IO_PaintCell(ws As Worksheet, rw As Long, val As Variant, _
                          bg As Long, isBold As Boolean, fmt As String)
    With ws.Cells(rw, IO_CV)
        If fmt <> "" And IsNumeric(val) Then
            .NumberFormat = fmt
            .Value = CDbl(val)
        Else
            .NumberFormat = "@"
            .Value = CStr(val)
        End If
        .Interior.Color = bg
        .Font.Bold = isBold: .Font.Italic = False
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 130, 50): .Borders.Weight = xlThin
    End With
End Sub

Private Sub IO_SafeMerge(ws As Worksheet, r1 As Long, c1 As Long, _
                          r2 As Long, c2 As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    On Error Resume Next: rng.UnMerge: rng.Merge: On Error GoTo 0
End Sub

Private Function IO_IsNumericCell(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If CStr(v) = "" Then Exit Function
    IO_IsNumericCell = IsNumeric(v)
End Function

Private Function IO_HasTextValue(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    Dim s As String: s = Trim(CStr(v))
    If s = "" Or s = "--" Then Exit Function
    IO_HasTextValue = True
End Function

Private Function IO_MemberComplete(ws As Worksheet, matRow As Long, _
                                   gauRow As Long, lenRow As Long, _
                                   widRow As Long) As Boolean
    IO_MemberComplete = _
        IO_HasTextValue(ws.Cells(matRow, IO_CV).Value) And _
        IO_IsNumericCell(ws.Cells(gauRow, IO_CV).Value) And _
        IO_IsNumericCell(ws.Cells(lenRow, IO_CV).Value) And _
        IO_IsNumericCell(ws.Cells(widRow, IO_CV).Value)
End Function

Private Sub IO_ClearCell(ws As Worksheet, r As Long, c As Long)
    Dim rng As Range: Set rng = ws.Cells(r, c)
    If rng.MergeCells Then
        rng.MergeArea.ClearContents
    Else
        rng.ClearContents
    End If
End Sub

'----------------------------------------------------------------------
'  IO_UtilColor  --  utilization gradient
'   >=90%       -> Red
'   80 -<90%    -> Yellow
'   40 -<80%    -> Green
'   20 -<40%    -> Yellow
'   <20%        -> Red
'----------------------------------------------------------------------
Private Function IO_UtilColor(pct As Double) As Long
    Select Case True
        Case pct >= 90#:                  IO_UtilColor = RGB(255, 120, 120)
        Case pct >= 80# And pct < 90#:    IO_UtilColor = RGB(255, 230, 100)
        Case pct >= 40# And pct < 80#:    IO_UtilColor = RGB(120, 210, 100)
        Case pct >= 20# And pct < 40#:    IO_UtilColor = RGB(255, 230, 100)
        Case Else:                        IO_UtilColor = RGB(255, 120, 120)
    End Select
End Function

'----------------------------------------------------------------------
'  IO_SortAsc  --  in-place insertion sort by one column (ascending)
'----------------------------------------------------------------------
Private Sub IO_SortAsc(arr As Variant, n As Long, sortCol As Integer)
    Dim i As Long, j As Long, k As Integer
    Dim cols As Integer: cols = UBound(arr, 2)
    Dim tmp() As Variant
    ReDim tmp(1 To cols)
    Dim keyVal As Double

    For i = 2 To n
        For k = 1 To cols
            tmp(k) = arr(i, k)
        Next k
        keyVal = IO_AsSortDouble(tmp(sortCol))
        j = i - 1
        Do While j >= 1
            If IO_AsSortDouble(arr(j, sortCol)) > keyVal Then
                For k = 1 To cols
                    arr(j + 1, k) = arr(j, k)
                Next k
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        For k = 1 To cols
            arr(j + 1, k) = tmp(k)
        Next k
    Next i
End Sub

Private Function IO_AsSortDouble(v As Variant) As Double
    If IsNumeric(v) Then
        IO_AsSortDouble = CDbl(v)
    Else
        IO_AsSortDouble = 1E+308
    End If
End Function