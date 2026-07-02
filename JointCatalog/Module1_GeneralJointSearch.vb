Option Explicit

'======================================================================
' LOOKUP INTERFACE  –  Joint Catalog Chapter 2
' Run SetupLookupInterface() ONCE (Alt+F8 ? Run) to build the UI.
' After setup, use the Search / Clear buttons on the Look Up tab.
'
' ?? VERIFY / UPDATE THE THREE SHEET NAMES BELOW TO MATCH YOUR FILE ??
'======================================================================
Private Const SH_BOLTED As String = "Bolted Connections"
Private Const SH_WELDED As String = "Welded Connections"
Private Const SH_LOOKUP As String = "General Joint Search"

'----------------------------------------------------------------------
' BOLTED data column indices  (1-based, row 1 = header)
'----------------------------------------------------------------------
Private Const BC_MAT  As Integer = 1    ' Connection Material
Private Const BC_GAU  As Integer = 2    ' Gauge
Private Const BC_NFT  As Integer = 3    ' Number of Fasteners
Private Const BC_BOLT As Integer = 4    ' Bolt Used
Private Const BC_DIAM As Integer = 5    ' Bolt Diameter (in)
Private Const BC_SHR  As Integer = 18   ' MaxShearStrength (kips)
Private Const BC_GOV  As Integer = 19   ' Governing mode (Bolt Shear / Bolt Bearing)
Private Const BC_TEN  As Integer = 21   ' MaxBoltTension (kips)
Private Const BC_CST  As Integer = 22   ' JointCost ($)

'----------------------------------------------------------------------
' WELDED data column indices
'----------------------------------------------------------------------
Private Const WC_BRMAT As Integer = 1   ' Branch Material
Private Const WC_BRGAU As Integer = 2   ' Branch Gauge
Private Const WC_BRSHP As Integer = 6   ' Branch Shape
Private Const WC_CHMAT As Integer = 7   ' Chord Material
Private Const WC_CHGAU As Integer = 8   ' Chord Gauge
Private Const WC_CHSHP As Integer = 10  ' Chord Shape
Private Const WC_ELEC  As Integer = 11  ' Electrode Specification
Private Const WC_WLEG  As Integer = 17  ' Weld Leg Size (in)
Private Const WC_SHR   As Integer = 20  ' Max Shear Strength (kips)
Private Const WC_CST   As Integer = 22  ' Joint Cost ($)

'----------------------------------------------------------------------
' LOOK UP sheet layout  (both panels share the same row numbers;
' Bolted occupies cols A-J, Welded occupies cols M-W)
'----------------------------------------------------------------------
Private Const RW_TTL As Long = 1    ' Title row
Private Const RW_BAR As Long = 2    ' Colour bar
Private Const RW_I1  As Long = 3    ' Input row 1
Private Const RW_I2  As Long = 4    ' Input row 2
Private Const RW_I3  As Long = 5    ' Input row 3
Private Const RW_I4  As Long = 6    ' Input row 4
Private Const RW_I5  As Long = 7    ' Input row 5  (Bolted only; welded leaves blank)
Private Const RW_BTN As Long = 8    ' Buttons row
Private Const RW_HDR As Long = 10   ' Results header
Private Const RW_DAT As Long = 11   ' First results row

' Column anchors
Private Const BL As Long = 1    ' Bolted label start  (col A)
Private Const BI As Long = 4    ' Bolted input cell   (col D)
Private Const WL As Long = 13   ' Welded label start  (col M)
Private Const WI As Long = 16   ' Welded input cell   (col P)

'======================================================================
' SetupLookupInterface  –  run ONCE to create the UI
'======================================================================
Sub SetupLookupInterface()
232    Application.ScreenUpdating = False

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_LOOKUP)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add( _
                 After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SH_LOOKUP
    End If

    ws.Cells.Clear
    Dim sp As Shape
    For Each sp In ws.Shapes: sp.Delete: Next sp

    Build_BoltedPanel ws
    Build_WeldedPanel ws

    ' ---- Column widths  (Bolted side  A-J) ---
    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 24
    ws.Columns(3).ColumnWidth = 7
    ws.Columns(4).ColumnWidth = 10
    ws.Columns(5).ColumnWidth = 22
    ws.Columns(6).ColumnWidth = 11
    ws.Columns(7).ColumnWidth = 13
    ws.Columns(8).ColumnWidth = 14
    ws.Columns(9).ColumnWidth = 17
    ws.Columns(10).ColumnWidth = 13
    ws.Columns(11).ColumnWidth = 2      ' spacer
    ws.Columns(12).ColumnWidth = 2      ' spacer
    ' ---- Column widths  (Welded side  M-W) ---
    ws.Columns(13).ColumnWidth = 5
    ws.Columns(14).ColumnWidth = 20
    ws.Columns(15).ColumnWidth = 9
    ws.Columns(16).ColumnWidth = 11
    ws.Columns(17).ColumnWidth = 20
    ws.Columns(18).ColumnWidth = 9
    ws.Columns(19).ColumnWidth = 11
    ws.Columns(20).ColumnWidth = 10
    ws.Columns(21).ColumnWidth = 10
    ws.Columns(22).ColumnWidth = 14
    ws.Columns(23).ColumnWidth = 13

    Application.ScreenUpdating = True

    MsgBox "Look Up interface created!" & vbCrLf & vbCrLf & _
           "Verify these tab names exist in your workbook:" & vbCrLf & _
           "  Bolted data  ->  """ & SH_BOLTED & """" & vbCrLf & _
           "  Welded data  ->  """ & SH_WELDED & """" & vbCrLf & vbCrLf & _
           "Edit SH_BOLTED / SH_WELDED at the top of the module if needed.", _
           vbInformation, "Setup Complete"
End Sub

'======================================================================
'  Build_BoltedPanel
'======================================================================
Private Sub Build_BoltedPanel(ws As Worksheet)
    Dim cBlue  As Long: cBlue = RGB(28, 78, 148)
    Dim cYel   As Long: cYel = RGB(255, 255, 204)
    Dim cHdr   As Long: cHdr = RGB(68, 114, 196)

    ' Title
    SafeMerge ws, RW_TTL, BL, RW_TTL, BL + 9
    ws.Rows(RW_TTL).RowHeight = 28
    With ws.Cells(RW_TTL, BL)
        .Value = "BOLTED CONNECTIONS  —  SEARCH"
        .Font.Bold = True:  .Font.Size = 14:  .Font.Color = cBlue
        .HorizontalAlignment = xlCenter
    End With

    ' Colour bar
    With ws.Range(ws.Cells(RW_BAR, BL), ws.Cells(RW_BAR, BL + 9))
        .Interior.Color = cBlue:  .RowHeight = 5
    End With

    ' Input rows
    Dim lbl(1 To 5) As String, hint(1 To 5) As String, iRow(1 To 5) As Long
    lbl(1) = "Min. Shear Capacity (kips):":     hint(1) = "":  iRow(1) = RW_I1
    lbl(2) = "Min. Tensile Capacity (kips):":   hint(2) = "":  iRow(2) = RW_I2
    lbl(3) = "Connection Material (opt.):":     iRow(3) = RW_I3
    hint(3) = "partial text, e.g.  Galvanized  or  Stainless"
    lbl(4) = "Bolt Type Filter (opt.):":        iRow(4) = RW_I4
    hint(4) = "partial text, e.g.  A307  /  A325  /  Grade 8"
    lbl(5) = "Max. Joint Cost $ (opt.):":       hint(5) = "":  iRow(5) = RW_I5

    Dim i As Integer
    For i = 1 To 5
        ws.Rows(iRow(i)).RowHeight = 22
        ' Label  (cols A-C)
        SafeMerge ws, iRow(i), BL, iRow(i), BL + 2
        With ws.Cells(iRow(i), BL)
            .Value = lbl(i):  .Font.Bold = True:  .VerticalAlignment = xlCenter
        End With
        ' Input cell  (cols D-F)
        SafeMerge ws, iRow(i), BI, iRow(i), BI + 2
        With ws.Cells(iRow(i), BI)
            .Interior.Color = cYel
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(160, 160, 0):  .Borders.Weight = xlThin
        End With
        ' Hint  (cols G-J)
        If hint(i) <> "" Then
            SafeMerge ws, iRow(i), BI + 3, iRow(i), BL + 9
            With ws.Cells(iRow(i), BI + 3)
                .Value = ">> " & hint(i)
                .Font.Italic = True:  .Font.Color = RGB(120, 120, 120):  .Font.Size = 8
            End With
        End If
    Next i

    ' Buttons
    ws.Rows(RW_BTN).RowHeight = 30
    Dim la As Double: la = ws.Cells(RW_BTN, BL).Left
    Dim ta As Double: ta = ws.Cells(RW_BTN, BL).Top + 4

    Dim bs As Shape
    Set bs = ws.Shapes.AddFormControl(xlButtonControl, la + 4, ta, 125, 22)
    bs.Name = "BtnBoltSearch"
    bs.TextFrame.Characters.Text = "Search Bolted Connections"
    bs.OnAction = "SearchBolted"
    bs.TextFrame.Characters.Font.Bold = True
    bs.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    bs.Fill.ForeColor.RGB = cBlue

    Dim bc As Shape
    Set bc = ws.Shapes.AddFormControl(xlButtonControl, la + 139, ta, 95, 22)
    bc.Name = "BtnBoltClear"
    bc.TextFrame.Characters.Text = "Clear Inputs & Results"
    bc.OnAction = "ClearBolted"
    bc.TextFrame.Characters.Font.Bold = True
    bc.Fill.ForeColor.RGB = RGB(200, 200, 200)

    ws.Rows(RW_BTN + 1).RowHeight = 6

    ' Results header
    ws.Rows(RW_HDR).RowHeight = 24
    Dim bh(0 To 9) As String
    bh(0) = "#":              bh(1) = "Connection Material"
    bh(2) = "Gauge":          bh(3) = "# Fasteners"
    bh(4) = "Bolt Type":      bh(5) = "Diam. (in)"
    bh(6) = "Shear Cap (kips)":  bh(7) = "Tension Cap (kips)"
    bh(8) = "Gov. Mode":      bh(9) = "Joint Cost ($)"
    For i = 0 To 9
        With ws.Cells(RW_HDR, BL + i)
            .Value = bh(i):  .Font.Bold = True
            .Font.Color = RGB(255, 255, 255):  .Interior.Color = cHdr
            .HorizontalAlignment = xlCenter:  .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous:  .Borders.Color = RGB(255, 255, 255)
        End With
    Next i

    ' Placeholder row
    ws.Rows(RW_DAT).RowHeight = 22
    SafeMerge ws, RW_DAT, BL, RW_DAT, BL + 9
    With ws.Cells(RW_DAT, BL)
        .Value = "Enter criteria above and click  Search Bolted Connections."
        .Font.Italic = True:  .Font.Color = RGB(140, 140, 140)
        .HorizontalAlignment = xlCenter
    End With
End Sub

'======================================================================
'  Build_WeldedPanel
'======================================================================
Private Sub Build_WeldedPanel(ws As Worksheet)
    Dim cRed  As Long: cRed = RGB(148, 28, 28)
    Dim cPink As Long: cPink = RGB(255, 218, 218)
    Dim cHdr  As Long: cHdr = RGB(196, 68, 68)

    ' Title
    SafeMerge ws, RW_TTL, WL, RW_TTL, WL + 10
    With ws.Cells(RW_TTL, WL)
        .Value = "WELDED CONNECTIONS  —  SEARCH"
        .Font.Bold = True:  .Font.Size = 14:  .Font.Color = cRed
        .HorizontalAlignment = xlCenter
    End With

    ' Colour bar
    With ws.Range(ws.Cells(RW_BAR, WL), ws.Cells(RW_BAR, WL + 10))
        .Interior.Color = cRed:  .RowHeight = 5
    End With

    ' Input rows  (4 inputs; row RW_I5 left blank for alignment with Bolted panel)
    Dim lbl(1 To 4) As String, hint(1 To 4) As String, iRow(1 To 4) As Long
    lbl(1) = "Min. Shear Capacity (kips):":   hint(1) = "":  iRow(1) = RW_I1
    lbl(2) = "Branch Material (optional):":   iRow(2) = RW_I2
    hint(2) = "partial text, e.g.  Stainless  or  Carbon"
    lbl(3) = "Chord Material (optional):":    iRow(3) = RW_I3
    hint(3) = "partial text, e.g.  Stainless  or  Carbon"
    lbl(4) = "Max. Joint Cost $ (opt.):":     hint(4) = "":  iRow(4) = RW_I4

    Dim i As Integer
    For i = 1 To 4
        ' Label  (cols M-O)
        SafeMerge ws, iRow(i), WL, iRow(i), WL + 2
        With ws.Cells(iRow(i), WL)
            .Value = lbl(i):  .Font.Bold = True:  .VerticalAlignment = xlCenter
        End With
        ' Input cell  (cols P-R)
        SafeMerge ws, iRow(i), WI, iRow(i), WI + 2
        With ws.Cells(iRow(i), WI)
            .Interior.Color = cPink
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(160, 80, 80):  .Borders.Weight = xlThin
        End With
        ' Hint  (cols S-W)
        If hint(i) <> "" Then
            SafeMerge ws, iRow(i), WI + 3, iRow(i), WL + 10
            With ws.Cells(iRow(i), WI + 3)
                .Value = ">> " & hint(i)
                .Font.Italic = True:  .Font.Color = RGB(120, 120, 120):  .Font.Size = 8
            End With
        End If
    Next i

    ' Buttons
    Dim la As Double: la = ws.Cells(RW_BTN, WL).Left
    Dim ta As Double: ta = ws.Cells(RW_BTN, WL).Top + 4

    Dim ws2 As Shape
    Set ws2 = ws.Shapes.AddFormControl(xlButtonControl, la + 4, ta, 125, 22)
    ws2.Name = "BtnWeldSearch"
    ws2.TextFrame.Characters.Text = "Search Welded Connections"
    ws2.OnAction = "SearchWelded"
    ws2.TextFrame.Characters.Font.Bold = True
    ws2.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    ws2.Fill.ForeColor.RGB = cRed

    Dim wc2 As Shape
    Set wc2 = ws.Shapes.AddFormControl(xlButtonControl, la + 139, ta, 95, 22)
    wc2.Name = "BtnWeldClear"
    wc2.TextFrame.Characters.Text = "Clear Inputs & Results"
    wc2.OnAction = "ClearWelded"
    wc2.TextFrame.Characters.Font.Bold = True
    wc2.Fill.ForeColor.RGB = RGB(200, 200, 200)

    ' Results header
    ws.Rows(RW_HDR).RowHeight = 24
    Dim wh(0 To 10) As String
    wh(0) = "#":                wh(1) = "Branch Material"
    wh(2) = "Br. Gauge":       wh(3) = "Br. Shape"
    wh(4) = "Chord Material": wh(5) = "Ch. Gauge"
    wh(6) = "Ch. Shape":       wh(7) = "Electrode"
    wh(8) = "Weld Leg (in)":   wh(9) = "Shear Cap (kips)"
    wh(10) = "Joint Cost ($)"
    For i = 0 To 10
        With ws.Cells(RW_HDR, WL + i)
            .Value = wh(i):  .Font.Bold = True
            .Font.Color = RGB(255, 255, 255):  .Interior.Color = cHdr
            .HorizontalAlignment = xlCenter:  .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous:  .Borders.Color = RGB(255, 255, 255)
        End With
    Next i

    ' Placeholder row
    ws.Rows(RW_DAT).RowHeight = 22
    SafeMerge ws, RW_DAT, WL, RW_DAT, WL + 10
    With ws.Cells(RW_DAT, WL)
        .Value = "Enter criteria above and click  Search Welded Connections."
        .Font.Italic = True:  .Font.Color = RGB(140, 140, 140)
        .HorizontalAlignment = xlCenter
    End With
End Sub

'======================================================================
'  SearchBolted  (assigned to Search button)
'======================================================================
Sub SearchBolted()
    
    Dim wsLU As Worksheet, wsBT As Worksheet
    On Error GoTo EH
    Set wsLU = ThisWorkbook.Sheets(SH_LOOKUP)
    On Error Resume Next: Set wsBT = ThisWorkbook.Sheets(SH_BOLTED): On Error GoTo EH
    If wsBT Is Nothing Then
        MsgBox "Sheet not found: """ & SH_BOLTED & """" & vbCrLf & _
               "Update SH_BOLTED in the VBA module.", vbCritical, "Sheet Missing": Exit Sub
    End If

    ' Read criteria
    Dim minShr As Double, hasShr As Boolean
    Dim minTen As Double, hasTen As Boolean
    Dim matFlt As String, hasMat As Boolean
    Dim bltFlt As String, hasBlt As Boolean
    Dim maxCst As Double, hasCst As Boolean

    With wsLU
        If IsNE(.Cells(RW_I1, BI)) Then minShr = CDbl(.Cells(RW_I1, BI)): hasShr = True
        If IsNE(.Cells(RW_I2, BI)) Then minTen = CDbl(.Cells(RW_I2, BI)): hasTen = True
        matFlt = Trim(CStr(.Cells(RW_I3, BI))): hasMat = (matFlt <> "")
        bltFlt = Trim(CStr(.Cells(RW_I4, BI))): hasBlt = (bltFlt <> "")
        If IsNE(.Cells(RW_I5, BI)) Then maxCst = CDbl(.Cells(RW_I5, BI)): hasCst = True
    End With

    If Not hasShr And Not hasTen And Not hasMat And Not hasBlt And Not hasCst Then
        MsgBox "Enter at least one criterion.", vbExclamation, "No Criteria": Exit Sub
    End If

    ClearBoltedArea wsLU   ' clear previous results only

    ' Scan and collect
    Dim lr As Long: lr = wsBT.Cells(wsBT.Rows.Count, 1).End(xlUp).Row
    Dim mR() As Long, mC() As Double, n As Long
    ReDim mR(1 To lr): ReDim mC(1 To lr)

    Dim r As Long, rs As Double, rt As Double, rc As Double, rm As String, rb As String
    For r = 2 To lr
        If Not IsNE(wsBT.Cells(r, BC_SHR)) Then GoTo SkB
        If Not IsNE(wsBT.Cells(r, BC_TEN)) Then GoTo SkB
        If Not IsNE(wsBT.Cells(r, BC_CST)) Then GoTo SkB
        rs = CDbl(wsBT.Cells(r, BC_SHR)): rt = CDbl(wsBT.Cells(r, BC_TEN))
        rc = CDbl(wsBT.Cells(r, BC_CST))
        rm = CStr(wsBT.Cells(r, BC_MAT)): rb = CStr(wsBT.Cells(r, BC_BOLT))
        If hasShr And rs < minShr Then GoTo SkB
        If hasTen And rt < minTen Then GoTo SkB
        If hasMat And InStr(1, rm, matFlt, vbTextCompare) = 0 Then GoTo SkB
        If hasBlt And InStr(1, rb, bltFlt, vbTextCompare) = 0 Then GoTo SkB
        If hasCst And rc > maxCst Then GoTo SkB
        n = n + 1: mR(n) = r: mC(n) = rc
SkB:
    Next r

    If n = 0 Then
        SafeMerge wsLU, RW_DAT, BL, RW_DAT, BL + 9
        With wsLU.Cells(RW_DAT, BL)
            .Value = "No matches found.  Try relaxing your criteria."
            .Font.Italic = True: .Font.Color = RGB(180, 0, 0): .HorizontalAlignment = xlCenter
        End With
        Exit Sub
    End If

    If n > 1 Then QuickSort mR, mC, 1, n

    ' Write results
    Application.ScreenUpdating = False
    Dim outR As Long: outR = RW_DAT
    Dim k As Long, bg As Long, c As Integer
    Dim util As Double, capShr As Double                         ' << NEW
    For k = 1 To n
        r = mR(k)
        ' --- utilisation gradient -----------------------------------------
        If hasShr And minShr > 0 Then
            capShr = CDbl(wsBT.Cells(r, BC_SHR).Value)
            util = (minShr / capShr) * 100#
            bg = UtilColor(util)
        Else
            bg = IIf(k Mod 2 = 1, RGB(235, 241, 252), RGB(255, 255, 255))
        End If
        ' ------------------------------------------------------------------
        wsLU.Rows(outR).RowHeight = 18
        With wsLU
            .Cells(outR, BL).Value = k
            .Cells(outR, BL + 1).Value = wsBT.Cells(r, BC_MAT).Value
            .Cells(outR, BL + 2).Value = wsBT.Cells(r, BC_GAU).Value
            .Cells(outR, BL + 3).Value = wsBT.Cells(r, BC_NFT).Value
            .Cells(outR, BL + 4).Value = ClnTxt(CStr(wsBT.Cells(r, BC_BOLT).Value))
            .Cells(outR, BL + 5).Value = wsBT.Cells(r, BC_DIAM).Value
            .Cells(outR, BL + 6).Value = wsBT.Cells(r, BC_SHR).Value
            .Cells(outR, BL + 7).Value = wsBT.Cells(r, BC_TEN).Value
            .Cells(outR, BL + 8).Value = wsBT.Cells(r, BC_GOV).Value
            .Cells(outR, BL + 9).Value = wsBT.Cells(r, BC_CST).Value
            .Cells(outR, BL + 5).NumberFormat = "0.000"
            .Cells(outR, BL + 6).NumberFormat = "0.000"
            .Cells(outR, BL + 7).NumberFormat = "0.000"
            .Cells(outR, BL + 9).NumberFormat = "$#,##0.00"
            For c = 0 To 9
                With .Cells(outR, BL + c)
                    .Interior.Color = bg
                    .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(200, 200, 215)
                    .VerticalAlignment = xlCenter: .WrapText = False
                End With
            Next c
            .Cells(outR, BL).HorizontalAlignment = xlCenter
            .Cells(outR, BL + 2).HorizontalAlignment = xlCenter
            .Cells(outR, BL + 3).HorizontalAlignment = xlCenter
            .Cells(outR, BL + 5).HorizontalAlignment = xlCenter
            .Cells(outR, BL + 6).HorizontalAlignment = xlRight
            .Cells(outR, BL + 7).HorizontalAlignment = xlRight
            .Cells(outR, BL + 9).HorizontalAlignment = xlRight
        End With
        outR = outR + 1
    Next k
    ' Summary bar
    SafeMerge wsLU, outR, BL, outR, BL + 9
    wsLU.Rows(outR).RowHeight = 20
    With wsLU.Cells(outR, BL)
        .Value = n & " result(s) found  |  NOT sorted by cost  (lowest to highest)"
        .Font.Italic = True: .Font.Bold = True: .Font.Color = RGB(28, 78, 148)
        .Interior.Color = RGB(210, 224, 245): .HorizontalAlignment = xlCenter
    End With
    Application.ScreenUpdating = True
    Exit Sub
EH: MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "SearchBolted"
End Sub

'======================================================================
'  SearchWelded  (assigned to Search button)
'======================================================================
Sub SearchWelded()
    Dim wsLU As Worksheet, wsWD As Worksheet
    On Error GoTo EH
    Set wsLU = ThisWorkbook.Sheets(SH_LOOKUP)
    On Error Resume Next: Set wsWD = ThisWorkbook.Sheets(SH_WELDED): On Error GoTo EH
    If wsWD Is Nothing Then
        MsgBox "Sheet not found: """ & SH_WELDED & """" & vbCrLf & _
               "Update SH_WELDED in the VBA module.", vbCritical, "Sheet Missing": Exit Sub
    End If

    ' Read criteria
    Dim minShr As Double, hasShr As Boolean
    Dim brFlt  As String, hasBr  As Boolean
    Dim chFlt  As String, hasCh  As Boolean
    Dim maxCst As Double, hasCst As Boolean

    With wsLU
        If IsNE(.Cells(RW_I1, WI)) Then minShr = CDbl(.Cells(RW_I1, WI)): hasShr = True
        brFlt = Trim(CStr(.Cells(RW_I2, WI))): hasBr = (brFlt <> "")
        chFlt = Trim(CStr(.Cells(RW_I3, WI))): hasCh = (chFlt <> "")
        If IsNE(.Cells(RW_I4, WI)) Then maxCst = CDbl(.Cells(RW_I4, WI)): hasCst = True
    End With

    If Not hasShr And Not hasBr And Not hasCh And Not hasCst Then
        MsgBox "Enter at least one criterion.", vbExclamation, "No Criteria": Exit Sub
    End If

    ClearWeldedArea wsLU

    Dim lr As Long: lr = wsWD.Cells(wsWD.Rows.Count, 1).End(xlUp).Row
    Dim mR() As Long, mC() As Double, n As Long
    ReDim mR(1 To lr): ReDim mC(1 To lr)

    Dim r As Long, rs As Double, rc As Double, rbm As String, rcm As String
    For r = 2 To lr
        If Not IsNE(wsWD.Cells(r, WC_SHR)) Then GoTo SkW
        If Not IsNE(wsWD.Cells(r, WC_CST)) Then GoTo SkW
        rs = CDbl(wsWD.Cells(r, WC_SHR)): rc = CDbl(wsWD.Cells(r, WC_CST))
        rbm = CStr(wsWD.Cells(r, WC_BRMAT)): rcm = CStr(wsWD.Cells(r, WC_CHMAT))
        If hasShr And rs < minShr Then GoTo SkW
        If hasBr And InStr(1, rbm, brFlt, vbTextCompare) = 0 Then GoTo SkW
        If hasCh And InStr(1, rcm, chFlt, vbTextCompare) = 0 Then GoTo SkW
        If hasCst And rc > maxCst Then GoTo SkW
        n = n + 1: mR(n) = r: mC(n) = rc
SkW:
    Next r

    If n = 0 Then
        SafeMerge wsLU, RW_DAT, WL, RW_DAT, WL + 10
        With wsLU.Cells(RW_DAT, WL)
            .Value = "No matches found.  Try relaxing your criteria."
            .Font.Italic = True: .Font.Color = RGB(180, 0, 0): .HorizontalAlignment = xlCenter
        End With
        Exit Sub
    End If

    If n > 1 Then QuickSort mR, mC, 1, n

    Application.ScreenUpdating = False
    Dim outR As Long: outR = RW_DAT
    Dim k As Long, bg As Long, c As Integer
    Dim util As Double, capShr As Double                         ' << NEW
    For k = 1 To n
        r = mR(k)
        ' --- utilisation gradient -----------------------------------------
        If hasShr And minShr > 0 Then
            capShr = CDbl(wsWD.Cells(r, WC_SHR).Value)
            util = (minShr / capShr) * 100#
            bg = UtilColor(util)
        Else
            bg = IIf(k Mod 2 = 1, RGB(252, 235, 235), RGB(255, 255, 255))
        End If
        ' ------------------------------------------------------------------
        wsLU.Rows(outR).RowHeight = 18
        With wsLU
            .Cells(outR, WL).Value = k
            .Cells(outR, WL + 1).Value = wsWD.Cells(r, WC_BRMAT).Value
            .Cells(outR, WL + 2).Value = wsWD.Cells(r, WC_BRGAU).Value
            .Cells(outR, WL + 3).Value = wsWD.Cells(r, WC_BRSHP).Value
            .Cells(outR, WL + 4).Value = wsWD.Cells(r, WC_CHMAT).Value
            .Cells(outR, WL + 5).Value = wsWD.Cells(r, WC_CHGAU).Value
            .Cells(outR, WL + 6).Value = wsWD.Cells(r, WC_CHSHP).Value
            .Cells(outR, WL + 7).Value = wsWD.Cells(r, WC_ELEC).Value
            .Cells(outR, WL + 8).Value = wsWD.Cells(r, WC_WLEG).Value
            .Cells(outR, WL + 9).Value = wsWD.Cells(r, WC_SHR).Value
            .Cells(outR, WL + 10).Value = wsWD.Cells(r, WC_CST).Value
            .Cells(outR, WL + 8).NumberFormat = "0.000"
            .Cells(outR, WL + 9).NumberFormat = "0.000"
            .Cells(outR, WL + 10).NumberFormat = "$#,##0.00"
            For c = 0 To 10
                With .Cells(outR, WL + c)
                    .Interior.Color = bg
                    .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(215, 200, 200)
                    .VerticalAlignment = xlCenter: .WrapText = False
                End With
            Next c
            .Cells(outR, WL).HorizontalAlignment = xlCenter
            .Cells(outR, WL + 2).HorizontalAlignment = xlCenter
            .Cells(outR, WL + 3).HorizontalAlignment = xlCenter
            .Cells(outR, WL + 5).HorizontalAlignment = xlCenter
            .Cells(outR, WL + 6).HorizontalAlignment = xlCenter
            .Cells(outR, WL + 7).HorizontalAlignment = xlCenter
            .Cells(outR, WL + 8).HorizontalAlignment = xlCenter
            .Cells(outR, WL + 9).HorizontalAlignment = xlRight
            .Cells(outR, WL + 10).HorizontalAlignment = xlRight
        End With
        outR = outR + 1
    Next k
    SafeMerge wsLU, outR, WL, outR, WL + 10
    wsLU.Rows(outR).RowHeight = 20
    With wsLU.Cells(outR, WL)
        .Value = n & " result(s) found  |  NOT sorted by cost  (lowest to highest)"
        .Font.Italic = True: .Font.Bold = True: .Font.Color = RGB(148, 28, 28)
        .Interior.Color = RGB(245, 210, 210): .HorizontalAlignment = xlCenter
    End With
    Application.ScreenUpdating = True
    Exit Sub
EH: MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "SearchWelded"
End Sub

'======================================================================
'  ClearBolted / ClearWelded  (public – assigned to Clear buttons)
'======================================================================
Sub ClearBolted()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_LOOKUP)
    ClearMergedCell ws, RW_I1, BI
    ClearMergedCell ws, RW_I2, BI
    ClearMergedCell ws, RW_I3, BI
    ClearMergedCell ws, RW_I4, BI
    ClearMergedCell ws, RW_I5, BI
    ClearBoltedArea ws
End Sub

Sub ClearWelded()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_LOOKUP)
    ClearMergedCell ws, RW_I1, WI
    ClearMergedCell ws, RW_I2, WI
    ClearMergedCell ws, RW_I3, WI
    ClearMergedCell ws, RW_I4, WI
    ClearWeldedArea ws
End Sub

'----------------------------------------------------------------------
Private Sub ClearBoltedArea(ws As Worksheet)
    Application.ScreenUpdating = False
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, BL).End(xlUp).Row
    If lr >= RW_DAT Then
        ws.Range(ws.Cells(RW_DAT, BL), ws.Cells(lr, BL + 9)).Clear
    End If
    SafeMerge ws, RW_DAT, BL, RW_DAT, BL + 9
    ws.Rows(RW_DAT).RowHeight = 22
    With ws.Cells(RW_DAT, BL)
        .Value = "Enter criteria above and click  Search Bolted Connections."
        .Font.Italic = True: .Font.Color = RGB(140, 140, 140): .HorizontalAlignment = xlCenter
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub ClearWeldedArea(ws As Worksheet)
    Application.ScreenUpdating = False
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, WL).End(xlUp).Row
    If lr >= RW_DAT Then
        ws.Range(ws.Cells(RW_DAT, WL), ws.Cells(lr, WL + 10)).Clear
    End If
    SafeMerge ws, RW_DAT, WL, RW_DAT, WL + 10
    ws.Rows(RW_DAT).RowHeight = 22
    With ws.Cells(RW_DAT, WL)
        .Value = "Enter criteria above and click  Search Welded Connections."
        .Font.Italic = True: .Font.Color = RGB(140, 140, 140): .HorizontalAlignment = xlCenter
    End With
    Application.ScreenUpdating = True
End Sub

'======================================================================
'  Helpers
'======================================================================
Private Sub SafeMerge(ws As Worksheet, r1 As Long, c1 As Long, r2 As Long, c2 As Long)
    Dim rng As Range: Set rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    On Error Resume Next: rng.UnMerge: rng.Merge: On Error GoTo 0
End Sub

Private Function IsNE(v As Variant) As Boolean
    ' True if the cell/value is numeric and non-empty
    If IsEmpty(v) Then Exit Function
    If CStr(v) = "" Then Exit Function
    IsNE = IsNumeric(v)
End Function

Private Function ClnTxt(s As String) As String
    Dim r As String: r = s
    r = Replace(r, Chr(10), " "): r = Replace(r, Chr(13), " "): r = Replace(r, "\n", " ")
    Do While InStr(r, "  ") > 0: r = Replace(r, "  ", " "): Loop
    ClnTxt = Trim(r)
End Function

Private Sub ClearMergedCell(ws As Worksheet, r As Long, c As Long)
    Dim rng As Range: Set rng = ws.Cells(r, c)
    If rng.MergeCells Then
        rng.MergeArea.ClearContents
    Else
        rng.ClearContents
    End If
End Sub

'----------------------------------------------------------------------
' UtilColor – row background based on shear utilisation %
'   >= 90 %  ?  Red     (near/at capacity)
'   80–90 %  ?  Yellow
'   40–80 %  ?  Green   (well matched)
'   20–40 %  ?  Yellow
'   <  20 %  ?  Red     (very over-specified)
'----------------------------------------------------------------------
Private Function UtilColor(pct As Double) As Long
    Select Case True
        Case pct >= 90: UtilColor = RGB(255, 120, 120)  ' Red
        Case pct >= 80: UtilColor = RGB(255, 230, 100)  ' Yellow
        Case pct >= 40: UtilColor = RGB(120, 210, 100)  ' Green
        Case pct >= 20: UtilColor = RGB(255, 230, 100)  ' Yellow
        Case Else:      UtilColor = RGB(255, 120, 120)  ' Red
    End Select
End Function

'======================================================================
' QuickSort – sorts by cost ascending (lowest first)
' Keeps row and cost arrays aligned
'======================================================================
Private Sub QuickSort(ByRef rws() As Long, ByRef csts() As Double, _
                      ByVal first As Long, ByVal last As Long)

    Dim i As Long, j As Long
    Dim pivot As Double
    Dim tempCost As Double
    Dim tempRow As Long

    i = first
    j = last
    pivot = csts((first + last) \ 2)   ' middle value as pivot

    Do While i <= j

        ' Move i forward
        Do While csts(i) < pivot
            i = i + 1
        Loop

        ' Move j backward
        Do While csts(j) > pivot
            j = j - 1
        Loop

        ' Swap
        If i <= j Then
            tempCost = csts(i)
            csts(i) = csts(j)
            csts(j) = tempCost

            tempRow = rws(i)
            rws(i) = rws(j)
            rws(j) = tempRow

            i = i + 1
            j = j - 1
        End If

    Loop

    ' Recursive calls
    If first < j Then QuickSort rws, csts, first, j
    If i < last Then QuickSort rws, csts, i, last

End Sub
