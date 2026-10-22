Attribute VB_Name = "ac2fNest"
'=====================================================================
'  ac2f pack  --  ac2fNest
'
'  True shape nesting: packs the selected parts onto sheets by their real
'  outline, not their bounding box. Concave parts interlock, small parts
'  can drop into the holes of big ones, and a drawn offcut of any shape
'  can serve as a sheet.
'
'  Macros:
'    ac2fNest           - opens the nesting window, then nests
'    ac2fNestReport     - same window, works it out, moves nothing
'    ac2fNestAgain      - nests with the last settings, no window
'    ac2fNestMarkSheets - names the selected shapes as offcut sheets
'
'  The engine is ac2fNestCore (pure arithmetic, tested outside CorelDRAW
'  by tools/test_nest.py). This module talks to the document:
'
'    1  sort the selection into parts and drawn offcut sheets
'    2  shapes lying inside another shape join it as one part (optional)
'    3  identical parts are found and share one set of rasters
'    4  every allowed turn (and mirror) of a part is made by turning the
'       real object in CorelDRAW, reading its outline, and turning it back
'       - so no assumption about rotation centres or curve maths is made
'    5  the engine packs, repeating with other orders while time allows
'    6  each part is turned the same way again and moved by its bounding
'       box, which is exact whatever the object's rotation centre
'
'  Outlines are read point by point along each curve
'  (Segment.GetPointPositionAt). If that call is missing or answers
'  nonsense, every curved segment is covered instead by a box that holds
'  ANY curve of its length between its two nodes (|AP| + |PB| <= L), so
'  a part can come out a little loose but never overlapping.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Nesting"

Private Const MAX_PARTS   As Long = 2000
Private Const MAX_SHEETS  As Long = 200
Private Const MAX_INST    As Long = 60
Private Const CELL_BUDGET As Double = 8000000#    ' raster cells planned for (2 bytes x 2)
Private Const MAX_SAMPLES As Long = 200           ' points per curved segment
Private Const PRESET_SEC  As String = "NestPresets"

'---------------------------------------------------------------------
' Registry keys (internal, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_NS_ML     As String = "NSKenarSolMM"
Public Const AC2F_K_NS_MR     As String = "NSKenarSagMM"
Public Const AC2F_K_NS_MT     As String = "NSKenarUstMM"
Public Const AC2F_K_NS_MB     As String = "NSKenarAltMM"
Public Const AC2F_K_NS_GAP    As String = "NSParcaAraligiMM"
Public Const AC2F_K_NS_ROT    As String = "NSDonmeAdimiDerece"
Public Const AC2F_K_NS_SW     As String = "NSPlakaGenisligiMM"
Public Const AC2F_K_NS_SH     As String = "NSPlakaYuksekligiMM"
Public Const AC2F_K_NS_SGAP   As String = "NSPlakaAraligiMM"
Public Const AC2F_K_NS_RES    As String = "NSHassasiyetMM"
Public Const AC2F_K_NS_TIME   As String = "NSOptimizeSaniye"
Public Const AC2F_K_NS_HOLES  As String = "NSDelikleriKullan"
Public Const AC2F_K_NS_INNER  As String = "NSIcDisBirlikte"
Public Const AC2F_K_NS_GROUP  As String = "NSGrupla"
Public Const AC2F_K_NS_MIRROR As String = "NSAynaSerbest"
Public Const AC2F_K_NS_CORNER As String = "NSBaslangicKosesi"
Public Const AC2F_K_NS_DIR    As String = "NSDoldurmaYonu"
Public Const AC2F_K_NS_STRAT  As String = "NSStrateji"
Public Const AC2F_K_NS_COPIES As String = "NSKopyaAdedi"
Public Const AC2F_K_NS_OUTLINE As String = "NSTabakaCizgisi"
Public Const AC2F_K_NS_START  As String = "NSBaslangicNoktasi"
Public Const AC2F_K_NS_OMARGIN As String = "NSArtikKenarMM"
Public Const AC2F_K_NS_SHEETS As String = "NSTabakaListesi"

Public Const AC2F_DEF_NS_ML     As Double = 10#
Public Const AC2F_DEF_NS_MR     As Double = 10#
Public Const AC2F_DEF_NS_MT     As Double = 10#
Public Const AC2F_DEF_NS_MB     As Double = 10#
Public Const AC2F_DEF_NS_GAP    As Double = 3#
Public Const AC2F_DEF_NS_ROT    As Double = 90#
Public Const AC2F_DEF_NS_SW     As Double = 0#    ' 0 = take the page width
Public Const AC2F_DEF_NS_SH     As Double = 0#    ' 0 = take the page height
Public Const AC2F_DEF_NS_SGAP   As Double = 20#
Public Const AC2F_DEF_NS_RES    As Double = 1#    ' raster cell, mm
Public Const AC2F_DEF_NS_TIME   As Double = 10#   ' seconds of optimising
Public Const AC2F_DEF_NS_HOLES  As Long = 0
Public Const AC2F_DEF_NS_INNER  As Long = 1
Public Const AC2F_DEF_NS_GROUP  As Long = 1       ' 0 none, 1 part, 2 sheet, 3 all
Public Const AC2F_DEF_NS_MIRROR As Long = 0
Public Const AC2F_DEF_NS_CORNER As Long = 1       ' 1 BL, 2 BR, 3 TL, 4 TR
Public Const AC2F_DEF_NS_DIR    As Long = 0       ' 0 rows, 1 columns
Public Const AC2F_DEF_NS_STRAT  As Long = 0       ' 0 auto, 1 top, 2 bottom-left
Public Const AC2F_DEF_NS_COPIES As Long = 1
Public Const AC2F_DEF_NS_OUTLINE As Long = 1
Public Const AC2F_DEF_NS_START  As Long = 0
Public Const AC2F_DEF_NS_OMARGIN As Double = 5#

'---------------------------------------------------------------------
' Settings of this run
'---------------------------------------------------------------------
Private m_gap As Double, m_mL As Double, m_mR As Double, m_mT As Double, m_mB As Double
Private m_om As Double, m_sgap As Double, m_rot As Double, m_res As Double, m_time As Double
Private m_mirror As Boolean, m_inner As Boolean, m_holes As Boolean
Private m_outline As Boolean, m_startPts As Boolean
Private m_copies As Long, m_corner As Long, m_dir As Long, m_strat As Long, m_group As Long

'---------------------------------------------------------------------
' Parts
'---------------------------------------------------------------------
Private m_shp() As Shape
Private m_own() As Boolean          ' a group this macro made (inner + outer)
Private m_cls() As Long             ' class per part, -1 = nothing to measure
Private m_np As Long
Private m_csig() As String          ' fingerprint per class
Private m_carea() As Double         ' material area per class, mm2
Private m_nc As Long
Private m_vAng() As Double, m_vMir() As Boolean     ' per variant tag
Private m_vBX() As Double, m_vBY() As Double        ' bounding box corner then
Private m_nv As Long
Private m_piPart() As Long, m_piCopy() As Long      ' engine part -> part, copy
Private m_piShp() As Shape
Private m_npi As Long

'---------------------------------------------------------------------
' Sheets: the list rows, saved ones and drawn offcuts mixed, in order
'---------------------------------------------------------------------
Private m_shName() As String
Private m_shW() As Double, m_shH() As Double
Private m_shQty() As Long
Private m_shUse() As Boolean
Private m_shDrawn() As Shape        ' Nothing for a sheet from the list
Private m_shType() As Long          ' engine type, -1 = not used
Private m_shN As Long
Private m_typeRow() As Long         ' engine type -> list row
Private m_nDrawn As Long

' sheet instances placed on the page
Private m_kCX() As Double, m_kCY() As Double        ' grid origin, engine frame
Private m_kSX() As Double, m_kSY() As Double        ' sheet corner on the page

'---------------------------------------------------------------------
' Engine frame: x' = a x + b y, y' = c x + e y (start corner, direction)
'---------------------------------------------------------------------
Private m_ta As Double, m_tb As Double, m_tc As Double, m_te As Double

' outline reading
Private m_flat As Long              ' 0 untested, 1 sampling works, -1 safe model
Private m_step As Double            ' sampling step along curves, mm
Private m_px() As Double, m_py() As Double
Private m_bAx() As Double, m_bAy() As Double, m_bBx() As Double, m_bBy() As Double
Private m_bL() As Double, m_bn As Long

' run bookkeeping
Private m_changed As Boolean        ' something was turned, grouped or copied
Private m_steps As Long
Private m_elapsed As Double
Private m_resUsed As Double
Private m_noMeasure As Long
Private m_nLocked As Long           ' locked objects left alone
Private m_x0 As Double, m_y0 As Double   ' lower left of the parts: sheet 1 goes there

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fNest()
Attribute ac2fNest.VB_Description = "ac2f pack: Nest the selection onto sheets by true shape"
    ac2fNSRun True, True
End Sub

Public Sub ac2fNestReport()
Attribute ac2fNestReport.VB_Description = "ac2f pack: Nesting report, moves nothing"
    ac2fNSRun False, True
End Sub

Public Sub ac2fNestAgain()
Attribute ac2fNestAgain.VB_Description = "ac2f pack: Nest with the last settings, no window"
    ac2fNSRun True, False
End Sub

Public Sub ac2fNestMarkSheets()
Attribute ac2fNestMarkSheets.VB_Description = "ac2f pack: Mark the selected shapes as offcut sheets"
    Dim sr As ShapeRange
    Dim i As Long, n As Long

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then Exit Sub
    If sr.Count = 0 Then
        ac2fWarn "Select the offcut shapes to use as sheets.", CAPTION_
        Exit Sub
    End If
    ActiveDocument.BeginCommandGroup ac2fTitle("mark offcuts")
    On Error Resume Next
    For i = 1 To sr.Count
        n = n + 1
        sr(i).Name = "offcut " & n
    Next i
    On Error GoTo 0
    ActiveDocument.EndCommandGroup
    ac2fInfo n & " shape(s) are now offcut sheets." & vbCrLf & vbCrLf & _
             "Select them together with the parts and run Nesting: each one " & _
             "is used as a sheet of its own shape, where it lies.", CAPTION_
End Sub

'=====================================================================
' MAIN FLOW
'=====================================================================

Private Sub ac2fNSRun(ByVal apply As Boolean, ByVal ask As Boolean)
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean, optOn As Boolean, inCmd As Boolean
    Dim i As Long, rc As Long, used As Long, choice As Long
    Dim preview As Boolean, shown As Boolean
    Dim msg As String

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Select the parts to nest.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Select the parts to nest.", CAPTION_
        Exit Sub
    End If

    On Error GoTo Fail
    ' every size is read in millimetres: the page, the offcuts, the parts
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    ac2fNSLoad
    ac2fNSLoadSheets
    ac2fNSClassify sr, True
    msg = ac2fNSPartsProblem()
    If Len(msg) > 0 Then GoTo Warn

    Do
        preview = False
        If ask Then
            rc = ac2fNSDialog(apply, preview)
            If rc = -2 Then
                If Not ac2fNSAskText(apply) Then GoTo Quit
            ElseIf rc = 0 Then
                GoTo Quit
            End If
            ac2fNSLoad
        End If

        used = 0
        For i = 0 To m_shN - 1
            If m_shUse(i) Then used = used + 1
        Next i
        If used = 0 Then
            msg = "No sheet is ticked to use."
            GoTo Warn
        End If

        oldOpt = Application.Optimization
        Application.Optimization = True
        optOn = True
        ActiveDocument.BeginCommandGroup ac2fTitle("nesting")
        inCmd = True

        choice = 0
        msg = ""
        If Not ac2fNSCompute() Then
            msg = "None of the ticked sheets leaves room inside its margins."
        ElseIf preview Then
            choice = ac2fNSPreview()
            If choice = -2 Then
                ' no preview window: do what the button said
                If apply Then choice = 1 Else choice = 0
            Else
                shown = True
            End If
        ElseIf apply Then
            choice = 1
        End If
        If choice = 1 Then ac2fNSApply
        If Len(msg) = 0 Then
            If choice = 1 Then msg = ac2fNSReport(1) Else msg = ac2fNSReport(0)
        End If

        ActiveDocument.EndCommandGroup
        inCmd = False
        ' anything but applying leaves the document exactly as it was
        If choice <> 1 And m_changed Then ActiveDocument.Undo
        Application.Optimization = oldOpt
        optOn = False
        Application.Refresh

        If choice <> 2 Then Exit Do
        ' back to the settings, with the parts as they were
        shown = False
        ac2fNSClassify sr, False
    Loop

    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    On Error GoTo 0
    ' after the preview window the report has been seen already
    If Not shown Then ac2fInfo msg, CAPTION_
    Exit Sub

Quit:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Exit Sub

Warn:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    On Error GoTo 0
    ac2fWarn msg, CAPTION_
    Exit Sub

Fail:
    msg = Err.Description
    On Error Resume Next
    If inCmd Then ActiveDocument.EndCommandGroup
    If optOn Then Application.Optimization = oldOpt
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Refresh
    ac2fWarn "The operation failed: " & msg & vbCrLf & vbCrLf & _
             "Ctrl+Z undoes whatever was done before the failure.", CAPTION_
End Sub

' Why the selection cannot be nested, or "".
Private Function ac2fNSPartsProblem() As String
    If m_np = 0 And m_nLocked > 0 Then
        ac2fNSPartsProblem = "The selected parts are locked. Unlock them (Arrange > Unlock) " & _
                             "and run Nesting again."
    ElseIf m_np = 0 Then
        ac2fNSPartsProblem = "Only sheets are selected. Select the parts as well."
    ElseIf m_np > MAX_PARTS Then
        ac2fNSPartsProblem = "That is " & m_np & " parts; the limit is " & MAX_PARTS & "."
    End If
End Function

' Works the layout out: inner shapes join their outer one, identical
' parts and their turns, the sheets, then the optimiser for the time
' allowed. False when no ticked sheet has room inside its margins.
Private Function ac2fNSCompute() As Boolean
    Dim i As Long
    Dim bx As Double, by As Double, bw As Double, bh As Double

    m_changed = False
    m_flat = 0
    ' where the parts are now: the first sheet goes there
    For i = 0 To m_np - 1
        m_shp(i).GetBoundingBox bx, by, bw, bh
        If i = 0 Then
            m_x0 = bx: m_y0 = by
        Else
            If bx < m_x0 Then m_x0 = bx
            If by < m_y0 Then m_y0 = by
        End If
    Next i

    ac2fNSFrame
    If m_inner And m_np > 1 Then ac2fNSMerge

    m_resUsed = ac2fNSPickRes()
    m_step = m_resUsed
    ac2fNCInit m_resUsed, m_gap / 2# + m_resUsed * 0.15 + 0.05, m_holes
    ac2fNCSetStrategy m_strat

    ac2fNSBuildClasses
    ac2fNSBuildParts
    If Not ac2fNSBuildSheets() Then Exit Function
    ac2fNCPrepare ac2fNSMaxInst()
    ac2fNSOptimise
    ac2fNSLayout
    ac2fNSCompute = True
End Function

'---------------------------------------------------------------------
' Settings
'---------------------------------------------------------------------
Private Sub ac2fNSLoad()
    m_gap = ac2fNSPos(AC2F_K_NS_GAP, AC2F_DEF_NS_GAP)
    m_mL = ac2fNSPos(AC2F_K_NS_ML, AC2F_DEF_NS_ML)
    m_mR = ac2fNSPos(AC2F_K_NS_MR, AC2F_DEF_NS_MR)
    m_mT = ac2fNSPos(AC2F_K_NS_MT, AC2F_DEF_NS_MT)
    m_mB = ac2fNSPos(AC2F_K_NS_MB, AC2F_DEF_NS_MB)
    m_om = ac2fNSPos(AC2F_K_NS_OMARGIN, AC2F_DEF_NS_OMARGIN)
    m_sgap = ac2fNSPos(AC2F_K_NS_SGAP, AC2F_DEF_NS_SGAP)
    m_rot = ac2fNSPos(AC2F_K_NS_ROT, AC2F_DEF_NS_ROT)
    If m_rot > 0 And m_rot < 15 Then m_rot = 15
    m_res = ac2fGetNum(AC2F_K_NS_RES, AC2F_DEF_NS_RES)
    If m_res < 0.2 Then m_res = 0.2
    If m_res > 20 Then m_res = 20
    m_time = ac2fNSPos(AC2F_K_NS_TIME, AC2F_DEF_NS_TIME)
    If m_time > 3600 Then m_time = 3600
    m_holes = (ac2fGetLng(AC2F_K_NS_HOLES, AC2F_DEF_NS_HOLES) <> 0)
    m_inner = (ac2fGetLng(AC2F_K_NS_INNER, AC2F_DEF_NS_INNER) <> 0)
    m_mirror = (ac2fGetLng(AC2F_K_NS_MIRROR, AC2F_DEF_NS_MIRROR) <> 0)
    m_outline = (ac2fGetLng(AC2F_K_NS_OUTLINE, AC2F_DEF_NS_OUTLINE) <> 0)
    m_startPts = (ac2fGetLng(AC2F_K_NS_START, AC2F_DEF_NS_START) <> 0)
    m_group = ac2fNSClampL(ac2fGetLng(AC2F_K_NS_GROUP, AC2F_DEF_NS_GROUP), 0, 3)
    m_corner = ac2fNSClampL(ac2fGetLng(AC2F_K_NS_CORNER, AC2F_DEF_NS_CORNER), 1, 4)
    m_dir = ac2fNSClampL(ac2fGetLng(AC2F_K_NS_DIR, AC2F_DEF_NS_DIR), 0, 1)
    m_strat = ac2fNSClampL(ac2fGetLng(AC2F_K_NS_STRAT, AC2F_DEF_NS_STRAT), 0, 2)
    m_copies = ac2fNSClampL(ac2fGetLng(AC2F_K_NS_COPIES, AC2F_DEF_NS_COPIES), 1, 500)
End Sub

Private Function ac2fNSPos(ByVal key As String, ByVal def As Double) As Double
    Dim v As Double
    v = ac2fGetNum(key, def)
    If v < 0 Then v = 0
    ac2fNSPos = v
End Function

Private Function ac2fNSClampL(ByVal v As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If v < lo Then v = lo
    If v > hi Then v = hi
    ac2fNSClampL = v
End Function

'---------------------------------------------------------------------
' Sheet list. Saved as "use;name;w;h;qty|..." - the drawn offcuts are
' never saved, they come from the selection each time.
'---------------------------------------------------------------------
Private Sub ac2fNSLoadSheets()
    Dim raw As String
    Dim rows() As String, f() As String
    Dim i As Long, j As Long
    Dim w As Double, h As Double, jw As Double, jh As Double
    Dim found As Boolean

    m_shN = 0
    m_nDrawn = 0
    ReDim m_shName(0 To 15): ReDim m_shW(0 To 15): ReDim m_shH(0 To 15)
    ReDim m_shQty(0 To 15): ReDim m_shUse(0 To 15): ReDim m_shDrawn(0 To 15)
    ReDim m_shType(0 To 15)

    raw = ac2fGetStr(AC2F_K_NS_SHEETS, "")
    If Len(raw) > 0 Then
        rows = Split(raw, "|")
        For i = 0 To UBound(rows)
            f = Split(rows(i) & ";;;;", ";")
            w = ac2fParseNum(f(2), 0)
            h = ac2fParseNum(f(3), 0)
            If w > 0 And h > 0 Then
                ac2fNSAddSheet (Trim$(f(0)) <> "0"), f(1), w, h, CLng(ac2fParseNum(f(4), 0)), Nothing
            End If
        Next i
    End If

    ' The customer job sets the sheet it took from stock; make sure it is
    ' in the list, on top.
    jw = ac2fGetNum(AC2F_K_NS_SW, AC2F_DEF_NS_SW)
    jh = ac2fGetNum(AC2F_K_NS_SH, AC2F_DEF_NS_SH)
    If jw > 0 And jh > 0 Then
        For i = 0 To m_shN - 1
            If Abs(m_shW(i) - jw) < 0.5 And Abs(m_shH(i) - jh) < 0.5 Then found = True
        Next i
        If Not found Then
            ac2fNSAddSheet True, "Job sheet " & ac2fNSDims(jw, jh), jw, jh, 0, Nothing
            ' move it to the top
            For j = m_shN - 1 To 1 Step -1
                ac2fNSSwapRows j, j - 1
            Next j
        End If
    End If

    If m_shN = 0 Then
        ac2fNSAddSheet True, "Page " & ac2fNSDims(ActivePage.SizeWidth, ActivePage.SizeHeight), _
                       ActivePage.SizeWidth, ActivePage.SizeHeight, 0, Nothing
    End If
End Sub

Private Sub ac2fNSSaveSheets()
    Dim i As Long
    Dim s As String, nm As String
    For i = 0 To m_shN - 1
        If m_shDrawn(i) Is Nothing Then
            nm = Replace$(Replace$(m_shName(i), "|", " "), ";", " ")
            If Len(s) > 0 Then s = s & "|"
            s = s & IIf(m_shUse(i), "1", "0") & ";" & nm & ";" & ac2fNumStr(m_shW(i)) & ";" & _
                ac2fNumStr(m_shH(i)) & ";" & CStr(m_shQty(i))
        End If
    Next i
    ac2fSetStr AC2F_K_NS_SHEETS, s
End Sub

Private Sub ac2fNSAddSheet(ByVal use As Boolean, ByVal nm As String, ByVal w As Double, _
                           ByVal h As Double, ByVal qty As Long, ByVal drawn As Shape)
    If m_shN >= MAX_SHEETS Then Exit Sub
    If m_shN > UBound(m_shName) Then
        ReDim Preserve m_shName(0 To m_shN * 2): ReDim Preserve m_shW(0 To m_shN * 2)
        ReDim Preserve m_shH(0 To m_shN * 2): ReDim Preserve m_shQty(0 To m_shN * 2)
        ReDim Preserve m_shUse(0 To m_shN * 2): ReDim Preserve m_shDrawn(0 To m_shN * 2)
        ReDim Preserve m_shType(0 To m_shN * 2)
    End If
    If Len(Trim$(nm)) = 0 Then nm = "Sheet " & ac2fNSDims(w, h)
    m_shName(m_shN) = Trim$(nm)
    m_shW(m_shN) = w
    m_shH(m_shN) = h
    If qty < 0 Then qty = 0
    m_shQty(m_shN) = qty
    m_shUse(m_shN) = use
    Set m_shDrawn(m_shN) = drawn
    m_shType(m_shN) = -1
    m_shN = m_shN + 1
End Sub

Private Sub ac2fNSSwapRows(ByVal i As Long, ByVal j As Long)
    Dim s As String, d As Double, q As Long, b As Boolean, sh As Shape
    s = m_shName(i): m_shName(i) = m_shName(j): m_shName(j) = s
    d = m_shW(i): m_shW(i) = m_shW(j): m_shW(j) = d
    d = m_shH(i): m_shH(i) = m_shH(j): m_shH(j) = d
    q = m_shQty(i): m_shQty(i) = m_shQty(j): m_shQty(j) = q
    b = m_shUse(i): m_shUse(i) = m_shUse(j): m_shUse(j) = b
    Set sh = m_shDrawn(i): Set m_shDrawn(i) = m_shDrawn(j): Set m_shDrawn(j) = sh
End Sub

Private Function ac2fNSDims(ByVal w As Double, ByVal h As Double) As String
    ac2fNSDims = Format$(w, "0") & " x " & Format$(h, "0")
End Function

'---------------------------------------------------------------------
' Parts and drawn sheets out of the selection
'---------------------------------------------------------------------
' withSheets False collects only the parts again: the sheet rows, drawn
' offcuts included, stay as the window left them.
Private Sub ac2fNSClassify(ByVal sr As ShapeRange, ByVal withSheets As Boolean)
    Dim i As Long
    Dim s As Shape
    Dim x As Double, y As Double, w As Double, h As Double

    m_np = 0
    m_nLocked = 0
    ReDim m_shp(0 To sr.Count)
    ReDim m_own(0 To sr.Count)
    For i = 1 To sr.Count
        Set s = sr(i)
        If ac2fNSIsOwnOutline(s) Then
            ' an outline drawn by an earlier run: neither part nor sheet
        ElseIf ac2fNSIsSheetShape(s) Then
            If withSheets Then
                On Error Resume Next
                s.GetBoundingBox x, y, w, h
                On Error GoTo 0
                m_nDrawn = m_nDrawn + 1
                ac2fNSAddSheet True, "Offcut " & m_nDrawn & ": " & ac2fNSShapeName(s) & " (" & _
                               ac2fNSDims(w, h) & ")", w, h, 1, s
            End If
        ElseIf ac2fNSIsLocked(s) Then
            ' CorelDRAW will not move it: left where it is, and said so
            m_nLocked = m_nLocked + 1
        Else
            Set m_shp(m_np) = s
            m_own(m_np) = False
            m_np = m_np + 1
        End If
    Next i
End Sub

Private Function ac2fNSIsLocked(ByVal s As Shape) As Boolean
    On Error Resume Next
    ac2fNSIsLocked = s.Locked
End Function

Private Function ac2fNSShapeName(ByVal s As Shape) As String
    Dim n As String
    On Error Resume Next
    n = Trim$(s.Name)
    If Len(n) = 0 Then n = "shape"
    ac2fNSShapeName = n
End Function

Private Function ac2fNSIsOwnOutline(ByVal s As Shape) As Boolean
    Dim n As String
    On Error Resume Next
    n = LCase$(s.Name)
    ac2fNSIsOwnOutline = (Left$(n, 10) = "ac2f sheet")
End Function

' A shape named like a sheet, or lying on a layer named like one.
Private Function ac2fNSIsSheetShape(ByVal s As Shape) As Boolean
    Dim n As String, ly As String
    On Error Resume Next
    n = LCase$(Trim$(s.Name))
    ly = LCase$(s.Layer.Name)
    On Error GoTo 0
    ac2fNSIsSheetShape = ac2fNSSheetWord(n, True) Or ac2fNSSheetWord(ly, False)
End Function

Private Function ac2fNSSheetWord(ByVal t As String, ByVal atStart As Boolean) As Boolean
    Dim w As Variant
    For Each w In Array("offcut", "sheet", "remnant", "artik", "plaka", "tabaka", "levha")
        If atStart Then
            If Left$(t, Len(w)) = w Then ac2fNSSheetWord = True
        Else
            If InStr(1, t, w) > 0 Then ac2fNSSheetWord = True
        End If
    Next w
End Function

'---------------------------------------------------------------------
' Engine frame from the start corner and the fill direction. The engine
' always packs from its own bottom left along its own rows; flipping and
' swapping the axes here makes that any corner and either direction.
'---------------------------------------------------------------------
Private Sub ac2fNSFrame()
    Dim fx As Double, fy As Double
    fx = 1: fy = 1
    If m_corner = 2 Or m_corner = 4 Then fx = -1
    If m_corner = 3 Or m_corner = 4 Then fy = -1
    If m_dir = 1 Then
        m_ta = 0: m_tb = fy: m_tc = fx: m_te = 0
    Else
        m_ta = fx: m_tb = 0: m_tc = 0: m_te = fy
    End If
End Sub

' Lower left corner, in the engine frame, of a box given on the page.
Private Sub ac2fNSFrameMin(ByVal x0 As Double, ByVal y0 As Double, _
                           ByVal x1 As Double, ByVal y1 As Double, _
                           ByRef ux As Double, ByRef uy As Double)
    Dim xs(0 To 3) As Double, ys(0 To 3) As Double
    Dim i As Long, tx As Double, ty As Double
    xs(0) = x0: ys(0) = y0: xs(1) = x1: ys(1) = y0
    xs(2) = x1: ys(2) = y1: xs(3) = x0: ys(3) = y1
    For i = 0 To 3
        tx = m_ta * xs(i) + m_tb * ys(i)
        ty = m_tc * xs(i) + m_te * ys(i)
        If i = 0 Or tx < ux Then ux = tx
        If i = 0 Or ty < uy Then uy = ty
    Next i
End Sub

'---------------------------------------------------------------------
' Inner shapes join the shape they lie in: grouped into one part
'---------------------------------------------------------------------
Private Sub ac2fNSMerge()
    Dim i As Long, j As Long, n As Long, cnt As Long
    Dim par() As Long
    Dim nshp() As Shape, nown() As Boolean
    Dim g As Shape
    Dim rng As ShapeRange

    ac2fNCPoolClear
    For i = 0 To m_np - 1
        ac2fNCGeoClear
        ac2fNSFlatten m_shp(i)
        ac2fNCPoolAdd
    Next i
    ac2fNCPoolResolve

    ReDim par(0 To m_np - 1)
    For i = 0 To m_np - 1
        par(i) = ac2fNCPoolParent(i)
    Next i
    ReDim nshp(0 To m_np): ReDim nown(0 To m_np)
    For i = 0 To m_np - 1
        If par(i) < 0 Then
            cnt = 0
            For j = 0 To m_np - 1
                If par(j) = i Then cnt = cnt + 1
            Next j
            If cnt = 0 Then
                Set nshp(n) = m_shp(i)
                nown(n) = m_own(i)
            Else
                ' a range, not the selection: CorelDRAW does not always keep
                ' the selection up to date while Optimization is on
                Set rng = New ShapeRange
                rng.Add m_shp(i)
                For j = 0 To m_np - 1
                    If par(j) = i Then rng.Add m_shp(j)
                Next j
                Set g = rng.Group()
                m_changed = True
                Set nshp(n) = g
                nown(n) = True
            End If
            n = n + 1
        End If
    Next i
    m_np = n
    m_shp = nshp
    m_own = nown
End Sub

'---------------------------------------------------------------------
' Grid size: the precision asked for, coarser only when the sheets the
' job will need would not fit the raster budget.
'---------------------------------------------------------------------
Private Function ac2fNSPickRes() As Double
    Dim p As Double, area As Double, cells As Double, u As Double, need As Double
    Dim i As Long
    Dim x As Double, y As Double, w As Double, h As Double

    p = m_res
    For i = 0 To m_np - 1
        m_shp(i).GetBoundingBox x, y, w, h
        area = area + (w + m_gap) * (h + m_gap) * m_copies
    Next i
    For i = 0 To m_shN - 1
        If m_shUse(i) Then
            If m_shDrawn(i) Is Nothing Then
                u = (m_shW(i) - m_mL - m_mR) * (m_shH(i) - m_mT - m_mB)
                If u > 0 And area > 0 Then
                    need = Int(area * 1.15 / u) + 1
                    If m_shQty(i) > 0 And need > m_shQty(i) Then need = m_shQty(i)
                    cells = cells + need * (m_shW(i) / p) * (m_shH(i) / p)
                    area = area - need * u
                    If area < 0 Then area = 0
                End If
            Else
                cells = cells + (m_shW(i) / p) * (m_shH(i) / p)
            End If
        End If
    Next i
    If cells > CELL_BUDGET Then p = p * Sqr(cells / CELL_BUDGET)
    ac2fNSPickRes = p
End Function

' How many sheets the engine may open, so memory stays bounded.
Private Function ac2fNSMaxInst() As Long
    Dim t As Long, w As Long, h As Long, fc As Long
    Dim ox As Double, oy As Double, mx As Double, n As Long

    For t = 0 To ac2fNCSheetTypeCount() - 1
        If ac2fNCSheetCells(t) > mx Then mx = ac2fNCSheetCells(t)
    Next t
    n = MAX_INST
    If mx > 0 Then
        If 2# * CELL_BUDGET / mx < n Then n = Int(2# * CELL_BUDGET / mx)
    End If
    If n < m_nDrawn + 1 Then n = m_nDrawn + 1
    ac2fNSMaxInst = n
End Function

'---------------------------------------------------------------------
' Classes and their variants
'---------------------------------------------------------------------
Private Sub ac2fNSBuildClasses()
    Dim i As Long, c As Long, a As Long, mi As Long, na As Long
    Dim angs(0 To 23) As Double
    Dim ang As Double
    Dim sig As String
    Dim bx As Double, by As Double, bw As Double, bh As Double
    Dim mir As Boolean, okMirror As Boolean

    na = 1
    angs(0) = 0
    If m_rot > 0 Then
        ang = m_rot
        Do While ang < 359.999 And na < 24
            angs(na) = ang
            na = na + 1
            ang = ang + m_rot
        Loop
    End If
    okMirror = m_mirror

    m_nc = 0: m_nv = 0: m_noMeasure = 0
    ReDim m_cls(0 To m_np): ReDim m_csig(0 To m_np): ReDim m_carea(0 To m_np)
    ReDim m_vAng(0 To 63): ReDim m_vMir(0 To 63)
    ReDim m_vBX(0 To 63): ReDim m_vBY(0 To 63)

    For i = 0 To m_np - 1
        ac2fNCGeoClear
        ac2fNSFlatten m_shp(i)
        sig = ""
        If ac2fNCGeoContours() > 0 Then sig = ac2fNCGeoSig()
        m_cls(i) = -1
        If Len(sig) = 0 Then
            m_noMeasure = m_noMeasure + 1
        Else
            For c = 0 To m_nc - 1
                If m_csig(c) = sig Then
                    m_cls(i) = c
                    Exit For
                End If
            Next c
            If m_cls(i) < 0 Then
                c = ac2fNCAddClass()
                m_csig(c) = sig
                m_carea(c) = ac2fNCGeoArea()
                m_nc = m_nc + 1
                m_cls(i) = c
                For mi = 0 To 1
                    mir = (mi = 1)
                    If mi = 0 Or okMirror Then
                        For a = 0 To na - 1
                            If ac2fNSOps(m_shp(i), mir, angs(a)) Then
                                m_shp(i).GetBoundingBox bx, by, bw, bh
                                ac2fNCGeoClear
                                ac2fNSFlatten m_shp(i)
                                ac2fNCGeoTransform m_ta, m_tb, m_tc, m_te
                                ac2fNSAddVariantTag mir, angs(a), bx, by
                                ac2fNSUndoOps m_shp(i), mir, angs(a)
                            ElseIf mir Then
                                okMirror = False      ' this CorelDRAW will not flip
                                Exit For
                            End If
                        Next a
                    End If
                Next mi
            End If
        End If
    Next i
End Sub

Private Sub ac2fNSAddVariantTag(ByVal mir As Boolean, ByVal ang As Double, _
                                ByVal bx As Double, ByVal by As Double)
    If m_nv > UBound(m_vAng) Then
        ReDim Preserve m_vAng(0 To m_nv * 2): ReDim Preserve m_vMir(0 To m_nv * 2)
        ReDim Preserve m_vBX(0 To m_nv * 2): ReDim Preserve m_vBY(0 To m_nv * 2)
    End If
    m_vAng(m_nv) = ang: m_vMir(m_nv) = mir
    m_vBX(m_nv) = bx: m_vBY(m_nv) = by
    ac2fNCAddVariant m_nv
    m_nv = m_nv + 1
End Sub

' Mirror then turn. False when the object refused.
Private Function ac2fNSOps(ByVal s As Shape, ByVal mir As Boolean, ByVal ang As Double) As Boolean
    On Error GoTo Fail
    If mir Then
        s.Flip cdrFlipHorizontal
        m_changed = True
    End If
    If ang <> 0 Then
        s.Rotate ang
        m_changed = True
    End If
    ac2fNSOps = True
    Exit Function
Fail:
End Function

Private Sub ac2fNSUndoOps(ByVal s As Shape, ByVal mir As Boolean, ByVal ang As Double)
    On Error Resume Next
    If ang <> 0 Then s.Rotate -ang
    If mir Then s.Flip cdrFlipHorizontal
End Sub

Private Sub ac2fNSBuildParts()
    Dim i As Long, c As Long
    m_npi = 0
    ReDim m_piPart(0 To m_np * m_copies + 1)
    ReDim m_piCopy(0 To m_np * m_copies + 1)
    ReDim m_piShp(0 To m_np * m_copies + 1)
    For i = 0 To m_np - 1
        If m_cls(i) >= 0 Then
            For c = 0 To m_copies - 1
                ac2fNCAddPart m_cls(i)
                m_piPart(m_npi) = i
                m_piCopy(m_npi) = c
                m_npi = m_npi + 1
            Next c
        End If
    Next i
End Sub

' One engine sheet type per ticked row, in list order.
Private Function ac2fNSBuildSheets() As Boolean
    Dim i As Long, t As Long, q As Long
    Dim uw As Double, uh As Double

    ReDim m_typeRow(0 To m_shN + 1)
    For i = 0 To m_shN - 1
        m_shType(i) = -1
        If m_shUse(i) Then
            t = -1
            If m_shDrawn(i) Is Nothing Then
                uw = m_shW(i) - m_mL - m_mR
                uh = m_shH(i) - m_mT - m_mB
                If uw > 0 And uh > 0 Then
                    q = m_shQty(i)
                    If q <= 0 Then q = -1
                    If m_dir = 1 Then
                        t = ac2fNCAddRectSheet(uh, uw, q)
                    Else
                        t = ac2fNCAddRectSheet(uw, uh, q)
                    End If
                End If
            Else
                ac2fNCGeoClear
                ac2fNSFlatten m_shDrawn(i)
                If ac2fNCGeoContours() > 0 Then
                    ac2fNCGeoTransform m_ta, m_tb, m_tc, m_te
                    t = ac2fNCAddShapeSheet(m_om)
                End If
            End If
            m_shType(i) = t
            If t >= 0 Then
                m_typeRow(t) = i
                ac2fNSBuildSheets = True
            End If
        End If
    Next i
End Function

Private Sub ac2fNSOptimise()
    Dim t0 As Double, el As Double

    ac2fNCOptBegin 1
    m_steps = 0
    t0 = Timer
    Do
        ac2fNCOptStep
        m_steps = m_steps + 1
        DoEvents
        el = Timer - t0
        If el < 0 Then el = el + 86400#
        If m_time <= 0 Then Exit Do
        If el >= m_time Then Exit Do
        If m_steps >= 100000 Then Exit Do
    Loop
    m_elapsed = el
End Sub

'=====================================================================
' OUTLINES INTO THE ENGINE
'=====================================================================

' Adds every outline of s to the engine's geometry buffer.
Private Sub ac2fNSFlatten(ByVal s As Shape)
    Dim i As Long
    Dim cv As Curve
    Dim dup As Shape
    Dim x As Double, y As Double, w As Double, h As Double

    If s Is Nothing Then Exit Sub
    On Error GoTo Skip
    Select Case s.Type
        Case cdrBitmapShape, cdrOLEObjectShape
            ' no path: its frame, solid
            s.GetBoundingBox x, y, w, h
            ac2fNCGeoBegin AC2F_NC_SOLID
            ac2fNCGeoPt x, y
            ac2fNCGeoPt x + w, y
            ac2fNCGeoPt x + w, y + h
            ac2fNCGeoPt x, y + h
            ac2fNCGeoEnd
            Exit Sub
        Case cdrGroupShape
            For i = 1 To s.Shapes.Count
                ac2fNSFlatten s.Shapes(i)
            Next i
            Exit Sub
    End Select

    Set cv = Nothing
    On Error Resume Next
    Set cv = s.DisplayCurve
    On Error GoTo Skip
    If Not cv Is Nothing Then
        ac2fNSFlattenCurve cv
        Exit Sub
    End If

    ' no DisplayCurve: read a throwaway copy converted to curves
    Set dup = s.Duplicate(0, 0)
    m_changed = True
    On Error Resume Next
    dup.ConvertToCurves
    On Error GoTo SkipDup
    If dup.Type = cdrGroupShape Then
        ac2fNSFlatten dup
    Else
        Set cv = Nothing
        On Error Resume Next
        Set cv = dup.Curve
        On Error GoTo SkipDup
        If Not cv Is Nothing Then ac2fNSFlattenCurve cv
    End If
    dup.Delete
    Exit Sub

SkipDup:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
    Exit Sub
Skip:
End Sub

Private Sub ac2fNSFlattenCurve(ByVal cv As Curve)
    Dim i As Long, k As Long, nSub As Long, nSeg As Long
    Dim sp As SubPath, sg As Segment
    Dim isClosed As Boolean
    Dim ex As Double, ey As Double

    On Error Resume Next
    nSub = cv.SubPaths.Count
    On Error GoTo 0
    For i = 1 To nSub
        Set sp = Nothing
        nSeg = 0
        isClosed = False
        On Error Resume Next
        Set sp = cv.SubPaths(i)
        nSeg = sp.Segments.Count
        isClosed = sp.Closed
        On Error GoTo 0
        If nSeg > 0 Then
            m_bn = 0
            If isClosed Then
                ac2fNCGeoBegin AC2F_NC_CLOSED
            Else
                ac2fNCGeoBegin AC2F_NC_OPEN
            End If
            For k = 1 To nSeg
                Set sg = sp.Segments(k)
                ac2fNSSegment sg
                If k = nSeg And Not isClosed Then
                    ex = sg.EndNode.PositionX
                    ey = sg.EndNode.PositionY
                    ac2fNCGeoPt ex, ey
                End If
            Next k
            ac2fNCGeoEnd
            ac2fNSFlushBlobs
        End If
    Next i
End Sub

' One segment: its start point, then points along it (not its end, which
' is the next segment's start).
Private Sub ac2fNSSegment(ByVal sg As Segment)
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim ln As Double, chord As Double
    Dim steps As Long, j As Long

    ax = sg.StartNode.PositionX: ay = sg.StartNode.PositionY
    bx = sg.EndNode.PositionX: by = sg.EndNode.PositionY
    ln = 0
    On Error Resume Next
    ln = sg.Length
    On Error GoTo 0
    chord = Sqr((bx - ax) * (bx - ax) + (by - ay) * (by - ay))
    ac2fNCGeoPt ax, ay
    If ln <= chord + 0.0005 * ln + 0.001 Then Exit Sub      ' straight

    steps = CLng(ac2fCeil(ln / m_step))
    If steps < 2 Then steps = 2
    If steps > MAX_SAMPLES Then steps = MAX_SAMPLES

    If m_flat >= 0 Then
        If ac2fNSSample(sg, steps, ax, ay, bx, by, ln) Then
            For j = 1 To steps - 1
                ac2fNCGeoPt m_px(j), m_py(j)
            Next j
            Exit Sub
        End If
        m_flat = -1
    End If

    ' safe model: the box that holds every curve of this length
    If m_bn = 0 Then
        ReDim m_bAx(0 To 63): ReDim m_bAy(0 To 63)
        ReDim m_bBx(0 To 63): ReDim m_bBy(0 To 63): ReDim m_bL(0 To 63)
    ElseIf m_bn > UBound(m_bAx) Then
        ReDim Preserve m_bAx(0 To m_bn * 2): ReDim Preserve m_bAy(0 To m_bn * 2)
        ReDim Preserve m_bBx(0 To m_bn * 2): ReDim Preserve m_bBy(0 To m_bn * 2)
        ReDim Preserve m_bL(0 To m_bn * 2)
    End If
    m_bAx(m_bn) = ax: m_bAy(m_bn) = ay
    m_bBx(m_bn) = bx: m_bBy(m_bn) = by
    m_bL(m_bn) = ln
    m_bn = m_bn + 1
End Sub

' Points j = 1 .. steps-1 along the segment into m_px/m_py. Checks the
' answers: on the first call that the ends come back as the nodes, and
' always that a point lies within reach of both ends (|AP| + |PB| <= L).
Private Function ac2fNSSample(ByVal sg As Segment, ByVal steps As Long, _
                              ByVal ax As Double, ByVal ay As Double, _
                              ByVal bx As Double, ByVal by As Double, _
                              ByVal ln As Double) As Boolean
    Dim j As Long
    Dim px As Double, py As Double, reach As Double

    On Error GoTo Fail
    ReDim m_px(0 To steps): ReDim m_py(0 To steps)
    If m_flat = 0 Then
        sg.GetPointPositionAt px, py, 0#, cdrRelativeSegmentOffset
        If Abs(px - ax) > 0.05 Or Abs(py - ay) > 0.05 Then Exit Function
        sg.GetPointPositionAt px, py, 1#, cdrRelativeSegmentOffset
        If Abs(px - bx) > 0.05 Or Abs(py - by) > 0.05 Then Exit Function
        m_flat = 1
    End If
    For j = 1 To steps - 1
        sg.GetPointPositionAt px, py, j / steps, cdrRelativeSegmentOffset
        reach = Sqr((px - ax) * (px - ax) + (py - ay) * (py - ay)) + _
                Sqr((px - bx) * (px - bx) + (py - by) * (py - by))
        If reach > ln * 1.002 + 0.05 Then Exit Function
        m_px(j) = px
        m_py(j) = py
    Next j
    ac2fNSSample = True
    Exit Function
Fail:
End Function

' Writes the safety boxes of the subpath just read as solid contours.
Private Sub ac2fNSFlushBlobs()
    Dim i As Long
    Dim mx As Double, my As Double, ux As Double, uy As Double, c As Double
    Dim hl As Double, hw As Double, vx As Double, vy As Double

    For i = 0 To m_bn - 1
        mx = (m_bAx(i) + m_bBx(i)) / 2#
        my = (m_bAy(i) + m_bBy(i)) / 2#
        c = Sqr((m_bBx(i) - m_bAx(i)) ^ 2 + (m_bBy(i) - m_bAy(i)) ^ 2)
        If c > 0.000001 Then
            ux = (m_bBx(i) - m_bAx(i)) / c
            uy = (m_bBy(i) - m_bAy(i)) / c
        Else
            ux = 1: uy = 0
        End If
        hl = m_bL(i) / 2#
        hw = (m_bL(i) / 2#) ^ 2 - (c / 2#) ^ 2
        If hw < 0 Then hw = 0
        hw = Sqr(hw)
        vx = -uy: vy = ux
        ac2fNCGeoBegin AC2F_NC_SOLID
        ac2fNCGeoPt mx - ux * hl - vx * hw, my - uy * hl - vy * hw
        ac2fNCGeoPt mx + ux * hl - vx * hw, my + uy * hl - vy * hw
        ac2fNCGeoPt mx + ux * hl + vx * hw, my + uy * hl + vy * hw
        ac2fNCGeoPt mx - ux * hl + vx * hw, my - uy * hl + vy * hw
        ac2fNCGeoEnd
    Next i
    m_bn = 0
End Sub

'=====================================================================
' APPLY
'=====================================================================

' Where each used sheet goes on the page: list sheets left to right from
' where the parts were, drawn offcuts stay where they lie. Also the grid
' origin of every sheet in the engine frame.
Private Sub ac2fNSLayout()
    Dim nk As Long, k As Long, t As Long, row As Long
    Dim parts As Long, rows As Long, cells As Long, seq As Long
    Dim cursor As Double, d As Double
    Dim sw As Long, sh As Long, fc As Long
    Dim tox As Double, toy As Double, ux As Double, uy As Double

    nk = ac2fNCInstanceCount()
    ReDim m_kCX(0 To nk): ReDim m_kCY(0 To nk)
    ReDim m_kSX(0 To nk): ReDim m_kSY(0 To nk)
    If nk = 0 Then Exit Sub
    d = ac2fNCGrow()
    cursor = m_x0
    For k = 0 To nk - 1
        ac2fNCInstanceUse k, parts, rows, cells, seq
        If parts > 0 Then
            t = ac2fNCInstanceType(k)
            row = m_typeRow(t)
            If m_shDrawn(row) Is Nothing Then
                m_kSX(k) = cursor
                m_kSY(k) = m_y0
                cursor = cursor + m_shW(row) + m_sgap
                ac2fNSFrameMin m_kSX(k) + m_mL, m_kSY(k) + m_mB, _
                               m_kSX(k) + m_shW(row) - m_mR, m_kSY(k) + m_shH(row) - m_mT, ux, uy
                m_kCX(k) = ux - d
                m_kCY(k) = uy - d
            Else
                ac2fNCSheetTypeInfo t, sw, sh, tox, toy, fc
                m_kCX(k) = tox
                m_kCY(k) = toy
            End If
        End If
    Next k
End Sub

Private Sub ac2fNSApply()
    Dim nk As Long, k As Long, row As Long, p As Long
    Dim parts As Long, rows As Long, cells As Long, seq As Long
    Dim res As Double
    Dim v As Long, cx As Long, cy As Long, tag As Long
    Dim vw As Long, vh As Long, vcells As Long
    Dim vox As Double, voy As Double
    Dim bx As Double, by As Double, bw As Double, bh As Double
    Dim dbx As Double, dby As Double, curX As Double, curY As Double
    Dim mx As Double, my As Double
    Dim s As Shape
    Dim outl() As Shape

    nk = ac2fNCInstanceCount()
    If nk = 0 Then Exit Sub
    ReDim outl(0 To nk)
    res = ac2fNCRes()

    ' copies first, from the parts as they still are
    For p = 0 To m_npi - 1
        Set m_piShp(p) = Nothing
        k = ac2fNCResult(p, v, cx, cy)
        If k >= 0 Then
            If m_piCopy(p) = 0 Then
                Set m_piShp(p) = m_shp(m_piPart(p))
            Else
                Set m_piShp(p) = m_shp(m_piPart(p)).Duplicate(0, 0)
                m_changed = True
            End If
        End If
    Next p

    ' turn like the variant, then move by the bounding box
    For p = 0 To m_npi - 1
        k = ac2fNCResult(p, v, cx, cy)
        If k >= 0 Then
            Set s = m_piShp(p)
            ac2fNCVariantInfo v, vw, vh, vox, voy, vcells, tag
            If ac2fNSOps(s, m_vMir(tag), m_vAng(tag)) Then
                s.GetBoundingBox bx, by, bw, bh
                dbx = bx - m_vBX(tag)
                dby = by - m_vBY(tag)
                curX = vox + m_ta * dbx + m_tb * dby
                curY = voy + m_tc * dbx + m_te * dby
                mx = m_kCX(k) + cx * res - curX
                my = m_kCY(k) + cy * res - curY
                s.Move m_ta * mx + m_tc * my, m_tb * mx + m_te * my
                m_changed = True
            End If
        End If
    Next p

    ' outlines of the list sheets used
    If m_outline Then
        For k = 0 To nk - 1
            ac2fNCInstanceUse k, parts, rows, cells, seq
            If parts > 0 Then
                row = m_typeRow(ac2fNCInstanceType(k))
                If m_shDrawn(row) Is Nothing Then
                    Set s = Nothing
                    On Error Resume Next
                    Set s = ActiveLayer.CreateRectangle(m_kSX(k), m_kSY(k) + m_shH(row), _
                                                        m_kSX(k) + m_shW(row), m_kSY(k))
                    If Not s Is Nothing Then
                        s.Fill.ApplyNoFill
                        s.Outline.Width = 0.25
                        s.Outline.Color.RGBAssign 150, 150, 150
                        s.Name = "ac2f sheet " & (k + 1) & " (" & ac2fNSDims(m_shW(row), m_shH(row)) & ")"
                        Set outl(k) = s
                        m_changed = True
                    End If
                    On Error GoTo 0
                End If
            End If
        Next k
    End If

    ac2fNSGroupResult nk
    ac2fNSGroupOutlines nk, outl

    If m_startPts Then ac2fNSStartPoints
End Sub

' Outlines stay out of the part groups, in one group "ac2f sheets" as
' before: the customer job, the check and the area price find them there.
Private Sub ac2fNSGroupOutlines(ByVal nk As Long, ByRef outl() As Shape)
    Dim k As Long
    Dim g As Shape
    Dim rng As ShapeRange

    On Error Resume Next
    Set rng = New ShapeRange
    For k = 0 To nk - 1
        If Not outl(k) Is Nothing Then rng.Add outl(k)
    Next k
    If rng.Count > 1 Then
        Set g = rng.Group()
        g.Name = "ac2f sheets"
    End If
    On Error GoTo 0
End Sub

' Groups the result the way the setting asks.
Private Sub ac2fNSGroupResult(ByVal nk As Long)
    Dim p As Long, k As Long, v As Long, cx As Long, cy As Long
    Dim g As Shape
    Dim rng As ShapeRange

    On Error Resume Next
    Select Case m_group
        Case 0
            ' no grouping: the inner + outer groups this run made go again
            For p = 0 To m_npi - 1
                If Not m_piShp(p) Is Nothing Then
                    If m_own(m_piPart(p)) Then m_piShp(p).Ungroup
                End If
            Next p
            For p = 0 To m_np - 1
                If m_own(p) And ac2fNSUnplaced(p) Then m_shp(p).Ungroup
            Next p
        Case 2
            For k = 0 To nk - 1
                Set rng = New ShapeRange
                For p = 0 To m_npi - 1
                    If ac2fNCResult(p, v, cx, cy) = k Then rng.Add m_piShp(p)
                Next p
                If rng.Count > 1 Then
                    Set g = rng.Group()
                    g.Name = "ac2f nest sheet " & (k + 1)
                End If
            Next k
        Case 3
            Set rng = New ShapeRange
            For p = 0 To m_npi - 1
                If Not m_piShp(p) Is Nothing Then rng.Add m_piShp(p)
            Next p
            If rng.Count > 1 Then
                Set g = rng.Group()
                g.Name = "ac2f nesting"
            End If
    End Select
    On Error GoTo 0
End Sub

Private Function ac2fNSUnplaced(ByVal part As Long) As Boolean
    Dim p As Long, v As Long, cx As Long, cy As Long
    ac2fNSUnplaced = True
    For p = 0 To m_npi - 1
        If m_piPart(p) = part Then
            If ac2fNCResult(p, v, cx, cy) >= 0 Then
                ac2fNSUnplaced = False
                Exit Function
            End If
        End If
    Next p
End Function

' Start points on everything placed, with the saved start point settings.
Private Sub ac2fNSStartPoints()
    Dim arr() As Shape
    Dim p As Long, n As Long, v As Long, cx As Long, cy As Long

    ReDim arr(0 To m_npi)
    For p = 0 To m_npi - 1
        If ac2fNCResult(p, v, cx, cy) >= 0 Then
            If Not m_piShp(p) Is Nothing Then
                Set arr(n) = m_piShp(p)
                n = n + 1
            End If
        End If
    Next p
    On Error Resume Next
    If n > 0 Then ac2fStartApplyShapes arr, n
    On Error GoTo 0
End Sub

'=====================================================================
' REPORT
'=====================================================================

' mode: 0 report only, 1 nested, 2 preview (nothing moved yet)
Private Function ac2fNSReport(ByVal mode As Long) As String
    Dim s As String
    Dim p As Long, k As Long, nk As Long, t As Long, row As Long
    Dim v As Long, cx As Long, cy As Long
    Dim parts As Long, rows As Long, cells As Long, seq As Long, last As Long, lastSeq As Long
    Dim placed As Long, total As Long, sheetsUsed As Long
    Dim sw As Long, sh As Long, fc As Long
    Dim tox As Double, toy As Double
    Dim partArea As Double, sheetArea As Double, usedArea As Double
    Dim kArea As Double, fill As Double, usedLen As Double
    Dim names As String, cornerTxt As String

    total = m_npi
    nk = ac2fNCInstanceCount()
    For p = 0 To m_npi - 1
        If ac2fNCResult(p, v, cx, cy) >= 0 Then placed = placed + 1
    Next p
    last = -1
    For k = 0 To nk - 1
        ac2fNCInstanceUse k, parts, rows, cells, seq
        If parts > 0 Then
            sheetsUsed = sheetsUsed + 1
            If seq > lastSeq Then
                lastSeq = seq
                last = k
            End If
        End If
    Next k

    Select Case mode
        Case 1: s = "NESTED" & vbCrLf
        Case 2: s = "PREVIEW - nothing has moved yet" & vbCrLf
        Case Else: s = "REPORT ONLY - nothing moved" & vbCrLf
    End Select
    s = s & "   Parts             : " & placed & " of " & total & " placed"
    If m_copies > 1 Then s = s & "  (" & m_copies & " of each)"
    s = s & vbCrLf
    If placed < total Then
        s = s & "   Did not fit       : " & (total - placed) & "  (left where they were)" & vbCrLf
    End If
    If m_noMeasure > 0 Then
        s = s & "   Nothing to measure: " & m_noMeasure & " object(s), left alone" & vbCrLf
    End If
    If m_nLocked > 0 Then
        s = s & "   Locked            : " & m_nLocked & " object(s), left alone" & vbCrLf
    End If
    s = s & "   Sheets used       : " & sheetsUsed & vbCrLf

    For k = 0 To nk - 1
        ac2fNCInstanceUse k, parts, rows, cells, seq
        If parts > 0 Then
            t = ac2fNCInstanceType(k)
            row = m_typeRow(t)
            ac2fNCSheetTypeInfo t, sw, sh, tox, toy, fc
            kArea = 0
            For p = 0 To m_npi - 1
                If ac2fNCResult(p, v, cx, cy) = k Then kArea = kArea + m_carea(m_cls(m_piPart(p)))
            Next p
            partArea = partArea + kArea
            If m_shDrawn(row) Is Nothing Then
                sheetArea = (m_shW(row) - m_mL - m_mR) * (m_shH(row) - m_mT - m_mB)
            Else
                sheetArea = fc * m_resUsed * m_resUsed
            End If
            fill = 0
            If sheetArea > 0 Then fill = kArea / sheetArea * 100#
            s = s & "     " & ac2fRPad(CStr(k + 1), 2) & "  " & ac2fPad(m_shName(row), 24) & _
                ac2fRPad(CStr(parts), 4) & " parts   fill " & ac2fFmt(fill, 1) & " %"
            If k = last And m_shDrawn(row) Is Nothing Then
                usedLen = rows * m_resUsed - 2# * ac2fNCGrow()
                If usedLen < 0 Then usedLen = 0
                s = s & "   used " & ac2fFmt(usedLen, 0) & " mm"
                If m_dir = 1 Then s = s & " of the width" Else s = s & " of the height"
                usedArea = usedArea + sheetArea * usedLen / IIf(m_dir = 1, m_shW(row) - m_mL - m_mR, m_shH(row) - m_mT - m_mB)
            Else
                usedArea = usedArea + sheetArea
            End If
            s = s & vbCrLf
        End If
    Next k
    If usedArea > 0 Then
        s = s & "   Fill (all sheets) : " & ac2fFmt(partArea / usedArea * 100#, 1) & _
                " %  (the last sheet up to its highest part)" & vbCrLf
    End If
    s = s & "   Layouts tried     : " & m_steps & " in " & ac2fFmt(m_elapsed, 1) & " s; best: " & _
            IIf(ac2fNCBestStrategy() = 1, "lowest top edge", "bottom-left") & _
            IIf(ac2fNCBestRotates(), "", ", no rotation") & vbCrLf
    s = s & vbCrLf

    Select Case m_corner
        Case 1: cornerTxt = "bottom left"
        Case 2: cornerTxt = "bottom right"
        Case 3: cornerTxt = "top left"
        Case Else: cornerTxt = "top right"
    End Select
    s = s & "SETTINGS USED" & vbCrLf
    s = s & "   Gap " & ac2fFmt(m_gap) & " mm, margins L R T B " & ac2fFmt(m_mL, 0) & " " & _
            ac2fFmt(m_mR, 0) & " " & ac2fFmt(m_mT, 0) & " " & ac2fFmt(m_mB, 0) & " mm" & vbCrLf
    s = s & "   Rotation " & IIf(m_rot <= 0, "off", ac2fFmt(m_rot, 0) & " deg") & _
            ", mirror " & IIf(m_mirror, "on", "off") & ", precision " & ac2fFmt(m_resUsed, 2) & " mm"
    If m_resUsed > m_res + 0.0001 Then s = s & " (coarsened to fit the job)"
    s = s & vbCrLf
    s = s & "   Inner shapes with outer " & IIf(m_inner, "on", "off") & ", holes used " & _
            IIf(m_holes, "on", "off") & ", from the " & cornerTxt & IIf(m_dir = 1, " in columns", " in rows") & vbCrLf
    If m_flat < 0 Then
        s = s & "   Curves were read with the safe model (this CorelDRAW does not report" & vbCrLf & _
                "   points along a curve), so parts sit a little looser." & vbCrLf
    End If
    If mode = 1 Then s = s & vbCrLf & "Ctrl+Z puts everything back."
    ac2fNSReport = s
End Function

'=====================================================================
' THE WINDOW
'=====================================================================

' Returns 1 to go on, 0 when cancelled, -2 when the window could not be
' shown (the caller falls back to the text prompt). apply turns True for
' Nest, preview for Preview; both stay False for the old Report only.
Private Function ac2fNSDialog(ByRef apply As Boolean, ByRef preview As Boolean) As Long
    Dim inf As String, body As String
    Dim lines() As String, f() As String
    Dim i As Long, n As Long, idx As Long
    Dim act As String, tag As String
    Dim nName() As String, nW() As Double, nH() As Double, nQ() As Long
    Dim nUse() As Boolean, nDr() As Shape
    Dim presets As Variant

    ac2fNSDialog = -2
    If ac2fGetLng(AC2F_K_JB_LIST, AC2F_DEF_JB_LIST) = 0 Then Exit Function
    inf = ac2fNSTemp("ac2f_nest_in.txt")

    body = "T|" & ac2fTitle(CAPTION_) & vbCrLf
    body = body & "P|" & m_np & " part(s) selected" & _
           IIf(m_nDrawn > 0, ", " & m_nDrawn & " drawn offcut sheet(s)", "") & _
           IIf(m_nLocked > 0, ", " & m_nLocked & " locked (left alone)", "") & _
           ". Preview shows the layout before anything moves; hover over a setting for what it does." & vbCrLf
    body = body & "G|" & ac2fNumStr(ActivePage.SizeWidth) & "|" & ac2fNumStr(ActivePage.SizeHeight) & vbCrLf
    body = body & ac2fNSKV("gap", m_gap) & ac2fNSKV("mL", m_mL) & ac2fNSKV("mR", m_mR) & _
           ac2fNSKV("mT", m_mT) & ac2fNSKV("mB", m_mB) & ac2fNSKV("om", m_om) & _
           ac2fNSKV("rot", m_rot) & ac2fNSKV("mirror", IIf(m_mirror, 1, 0)) & _
           ac2fNSKV("copies", m_copies) & ac2fNSKV("inner", IIf(m_inner, 1, 0)) & _
           ac2fNSKV("holes", IIf(m_holes, 1, 0)) & ac2fNSKV("corner", m_corner) & _
           ac2fNSKV("dir", m_dir) & ac2fNSKV("strat", m_strat) & ac2fNSKV("res", m_res) & _
           ac2fNSKV("time", m_time) & ac2fNSKV("group", m_group) & _
           ac2fNSKV("outline", IIf(m_outline, 1, 0)) & ac2fNSKV("start", IIf(m_startPts, 1, 0)) & _
           ac2fNSKV("sgap", m_sgap)
    For i = 0 To m_shN - 1
        If m_shDrawn(i) Is Nothing Then
            tag = "L"
        Else
            tag = "D" & i
        End If
        body = body & "S|" & IIf(m_shUse(i), "1", "0") & "|" & ac2fNSClean(m_shName(i)) & "|" & _
               ac2fNumStr(m_shW(i)) & "|" & ac2fNumStr(m_shH(i)) & "|" & m_shQty(i) & "|" & tag & vbCrLf
    Next i
    On Error Resume Next
    presets = GetAllSettings(AC2F_REG_APP, PRESET_SEC)
    If IsArray(presets) Then
        For i = LBound(presets, 1) To UBound(presets, 1)
            body = body & "R|" & ac2fNSClean(CStr(presets(i, 0))) & "|" & ac2fNSClean(CStr(presets(i, 1))) & vbCrLf
        Next i
    End If
    On Error GoTo Fail

    If Not ac2fNSWriteUtf8(inf, body) Then Exit Function
    body = ac2fNSShowWindow(1, inf)
    If Len(body) = 0 Then Exit Function
    lines = Split(body, vbLf)

    ' presets come back whatever the button, so a saved one is kept
    On Error Resume Next
    DeleteSetting AC2F_REG_APP, PRESET_SEC
    On Error GoTo Fail
    For i = 0 To UBound(lines)
        f = Split(lines(i) & "|||", "|")
        If f(0) = "R" And Len(f(1)) > 0 Then SaveSetting AC2F_REG_APP, PRESET_SEC, f(1), f(2)
        If f(0) = "A" Then act = f(1)
    Next i
    If act <> "NEST" And act <> "PREVIEW" And act <> "REPORT" Then
        ac2fNSDialog = 0
        Exit Function
    End If
    apply = (act = "NEST")
    preview = (act = "PREVIEW")

    ' settings, then the sheet rows in their new order
    ReDim nName(0 To m_shN + 64): ReDim nW(0 To m_shN + 64): ReDim nH(0 To m_shN + 64)
    ReDim nQ(0 To m_shN + 64): ReDim nUse(0 To m_shN + 64): ReDim nDr(0 To m_shN + 64)
    For i = 0 To UBound(lines)
        f = Split(lines(i) & "||||||||", "|")
        If f(0) = "K" Then ac2fNSStoreKey f(1), f(2)
        If f(0) = "S" And n <= UBound(nName) Then
            nUse(n) = (f(1) = "1")
            nName(n) = f(2)
            nW(n) = ac2fParseNum(f(3), 0)
            nH(n) = ac2fParseNum(f(4), 0)
            nQ(n) = CLng(ac2fParseNum(f(5), 0))
            Set nDr(n) = Nothing
            If Left$(f(6), 1) = "D" Then
                idx = CLng(Val(Mid$(f(6), 2)))
                If idx >= 0 And idx < m_shN Then
                    Set nDr(n) = m_shDrawn(idx)
                    nW(n) = m_shW(idx): nH(n) = m_shH(idx): nQ(n) = 1
                End If
            End If
            If nW(n) > 0 And nH(n) > 0 Then n = n + 1
        End If
    Next i
    m_shN = 0
    For i = 0 To n - 1
        ac2fNSAddSheet nUse(i), nName(i), nW(i), nH(i), nQ(i), nDr(i)
    Next i
    ac2fNSSaveSheets
    ac2fNSDialog = 1
    Exit Function

Fail:
    ac2fNSDialog = -2
End Function

' Runs window script which (1 settings, 2 preview) on an input file that
' is already written; returns what the window wrote back (lines split by
' vbLf), or "" when it could not be shown.
Private Function ac2fNSShowWindow(ByVal which As Long, ByVal inf As String) As String
    Dim ps As String, outf As String, cmd As String, body As String
    Dim rc As Long, ff As Integer
    Dim ok As Boolean
    Dim sh As Object

    On Error GoTo Fail
    If which = 2 Then
        ps = ac2fNSTemp("ac2f_nest_preview.ps1")
        outf = ac2fNSTemp("ac2f_nest_preview_out.txt")
    Else
        ps = ac2fNSTemp("ac2f_nest.ps1")
        outf = ac2fNSTemp("ac2f_nest_out.txt")
    End If
    If Len(Dir$(outf)) > 0 Then Kill outf

    ff = FreeFile
    Open ps For Output As #ff
    ok = ac2fNSWriteScript(which, ff)
    Close #ff
    ff = 0
    If Not ok Then Exit Function

    Set sh = CreateObject("WScript.Shell")
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden" & _
          " -File """ & ps & """ -In """ & inf & """ -Out """ & outf & """"
    rc = sh.Run(cmd, 0, True)
    If rc <> 0 Then Exit Function
    If Len(Dir$(outf)) = 0 Then Exit Function
    body = ac2fNSReadUtf8(outf)
    If Len(body) > 0 Then
        If AscW(Left$(body, 1)) = &HFEFF Then body = Mid$(body, 2)
    End If
    ac2fNSShowWindow = Replace$(body, vbCrLf, vbLf)
    Exit Function

Fail:
    On Error Resume Next
    If ff <> 0 Then Close #ff
End Function

'---------------------------------------------------------------------
' Preview: the layout drawn in a window before anything moves
'---------------------------------------------------------------------

' 1 apply, 2 back to the settings, 0 cancel, -2 no window.
Private Function ac2fNSPreview() As Long
    Dim inf As String, body As String
    Dim lines() As String
    Dim i As Long

    ac2fNSPreview = -2
    If ac2fGetLng(AC2F_K_JB_LIST, AC2F_DEF_JB_LIST) = 0 Then Exit Function
    inf = ac2fNSTemp("ac2f_nest_preview_in.txt")
    If Not ac2fNSPreviewWrite(inf) Then Exit Function
    body = ac2fNSShowWindow(2, inf)
    If Len(body) = 0 Then Exit Function
    lines = Split(body, vbLf)
    ac2fNSPreview = 0
    For i = 0 To UBound(lines)
        If lines(i) = "A|APPLY" Then ac2fNSPreview = 1
        If lines(i) = "A|SETTINGS" Then ac2fNSPreview = 2
    Next i
End Function

' Input of the preview window, written straight to the file (it can run to
' megabytes; building it as one string would be slow in VBA). Coordinates
' are on the page in tenths of a millimetre, whole numbers, so nothing
' depends on the decimal separator. ANSI like every Print #; the window
' reads it with the same code page.
'   T|title            R|report line
'   S|n|name|parts|x0|y0|x1|y1|ux0|uy0|ux1|uy1     a sheet and its usable box
'   O|n|kind|x y x y ...                           outline of a drawn offcut
'   P|n|part|class|kind|x y x y ...                one contour of a placed part
Private Function ac2fNSPreviewWrite(ByVal path As String) As Boolean
    Dim ff As Integer
    Dim rep As String, ln As String
    Dim lines() As String, xs() As Double, ys() As Double
    Dim nk As Long, k As Long, t As Long, row As Long, p As Long, c As Long, i As Long
    Dim parts As Long, rows As Long, cells As Long, seq As Long
    Dim v As Long, cx As Long, cy As Long, vw As Long, vh As Long, vcells As Long, tag As Long
    Dim n As Long, kind As Long, s0 As Long, stp As Long
    Dim vox As Double, voy As Double, tx As Double, ty As Double, ex As Double, ey As Double
    Dim x0 As Double, y0 As Double, w As Double, h As Double, res As Double

    On Error GoTo Fail
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "T|" & ac2fTitle(CAPTION_) & " - preview"
    rep = ac2fNSReport(2)
    lines = Split(Replace$(rep, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        Print #ff, "R|" & ac2fNSClean(lines(i))
    Next i

    nk = ac2fNCInstanceCount()
    res = ac2fNCRes()
    For k = 0 To nk - 1
        ac2fNCInstanceUse k, parts, rows, cells, seq
        If parts > 0 Then
            t = ac2fNCInstanceType(k)
            row = m_typeRow(t)
            If m_shDrawn(row) Is Nothing Then
                x0 = m_kSX(k): y0 = m_kSY(k): w = m_shW(row): h = m_shH(row)
                Print #ff, "S|" & (k + 1) & "|" & ac2fNSClean(m_shName(row)) & "|" & parts & "|" & _
                    ac2fNSI(x0) & "|" & ac2fNSI(y0) & "|" & ac2fNSI(x0 + w) & "|" & ac2fNSI(y0 + h) & "|" & _
                    ac2fNSI(x0 + m_mL) & "|" & ac2fNSI(y0 + m_mB) & "|" & _
                    ac2fNSI(x0 + w - m_mR) & "|" & ac2fNSI(y0 + h - m_mT)
            Else
                m_shDrawn(row).GetBoundingBox x0, y0, w, h
                Print #ff, "S|" & (k + 1) & "|" & ac2fNSClean(m_shName(row)) & "|" & parts & "|" & _
                    ac2fNSI(x0) & "|" & ac2fNSI(y0) & "|" & ac2fNSI(x0 + w) & "|" & ac2fNSI(y0 + h) & "|" & _
                    ac2fNSI(x0) & "|" & ac2fNSI(y0) & "|" & ac2fNSI(x0 + w) & "|" & ac2fNSI(y0 + h)
                ac2fNCGeoClear
                ac2fNSFlatten m_shDrawn(row)
                For c = 0 To ac2fNCGeoContours() - 1
                    ac2fNCGeoContourInfo c, s0, n, kind
                    stp = (n + 199) \ 200
                    If stp < 1 Then stp = 1
                    ln = "O|" & (k + 1) & "|" & kind & "|"
                    For i = s0 To s0 + n - 1 Step stp
                        If i > s0 Then ln = ln & " "
                        ln = ln & ac2fNSI(ac2fNCGeoX(i)) & " " & ac2fNSI(ac2fNCGeoY(i))
                    Next i
                    Print #ff, ln
                Next c
            End If
        End If
    Next k

    ' every placed part: its variant's outline, moved to where the raster
    ' went (engine frame), then back to the page frame
    For p = 0 To m_npi - 1
        k = ac2fNCResult(p, v, cx, cy)
        If k >= 0 Then
            ac2fNCVariantInfo v, vw, vh, vox, voy, vcells, tag
            tx = m_kCX(k) + cx * res - vox
            ty = m_kCY(k) + cy * res - voy
            For c = 0 To ac2fNCVariantContourCount(v) - 1
                n = ac2fNCVariantContour(v, c, xs, ys, kind)
                ln = "P|" & (k + 1) & "|" & p & "|" & m_cls(m_piPart(p)) & "|" & kind & "|"
                For i = 0 To n - 1
                    ex = xs(i) + tx
                    ey = ys(i) + ty
                    If i > 0 Then ln = ln & " "
                    ln = ln & ac2fNSI(m_ta * ex + m_tc * ey) & " " & ac2fNSI(m_tb * ex + m_te * ey)
                Next i
                Print #ff, ln
            Next c
        End If
    Next p
    Close #ff
    ac2fNSPreviewWrite = True
    Exit Function

Fail:
    On Error Resume Next
    Close #ff
End Function

' Tenths of a millimetre as a whole number.
Private Function ac2fNSI(ByVal x As Double) As String
    ac2fNSI = CStr(CLng(x * 10#))
End Function

Private Function ac2fNSKV(ByVal k As String, ByVal v As Double) As String
    ac2fNSKV = "K|" & k & "|" & ac2fNumStr(v) & vbCrLf
End Function

Private Function ac2fNSClean(ByVal s As String) As String
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    ac2fNSClean = Replace$(s, "|", " ")
End Function

' One value from the window into the registry.
Private Sub ac2fNSStoreKey(ByVal k As String, ByVal v As String)
    Dim x As Double
    x = ac2fParseNum(v, -1)
    If x < 0 Then Exit Sub
    Select Case k
        Case "gap": ac2fSetNum AC2F_K_NS_GAP, x
        Case "mL": ac2fSetNum AC2F_K_NS_ML, x
        Case "mR": ac2fSetNum AC2F_K_NS_MR, x
        Case "mT": ac2fSetNum AC2F_K_NS_MT, x
        Case "mB": ac2fSetNum AC2F_K_NS_MB, x
        Case "om": ac2fSetNum AC2F_K_NS_OMARGIN, x
        Case "rot": ac2fSetNum AC2F_K_NS_ROT, x
        Case "mirror": ac2fSetLng AC2F_K_NS_MIRROR, CLng(x)
        Case "copies": ac2fSetLng AC2F_K_NS_COPIES, CLng(x)
        Case "inner": ac2fSetLng AC2F_K_NS_INNER, CLng(x)
        Case "holes": ac2fSetLng AC2F_K_NS_HOLES, CLng(x)
        Case "corner": ac2fSetLng AC2F_K_NS_CORNER, CLng(x)
        Case "dir": ac2fSetLng AC2F_K_NS_DIR, CLng(x)
        Case "strat": ac2fSetLng AC2F_K_NS_STRAT, CLng(x)
        Case "res": ac2fSetNum AC2F_K_NS_RES, x
        Case "time": ac2fSetNum AC2F_K_NS_TIME, x
        Case "group": ac2fSetLng AC2F_K_NS_GROUP, CLng(x)
        Case "outline": ac2fSetLng AC2F_K_NS_OUTLINE, CLng(x)
        Case "start": ac2fSetLng AC2F_K_NS_START, CLng(x)
        Case "sgap": ac2fSetNum AC2F_K_NS_SGAP, x
    End Select
End Sub

'---------------------------------------------------------------------
' Text fallback when the window cannot open
'---------------------------------------------------------------------
Private Function ac2fNSAskText(ByRef apply As Boolean) As Boolean
    Dim cmd As String, s As String, t As String
    Dim i As Long

    Do
        s = "NESTING  -  " & m_np & " part(s)"
        If m_nDrawn > 0 Then s = s & ", " & m_nDrawn & " drawn offcut(s)"
        s = s & vbCrLf & vbCrLf & "SHEETS  (top first, qty 0 = as needed)" & vbCrLf
        For i = 0 To m_shN - 1
            If i < 8 Then
                s = s & "  " & IIf(m_shUse(i), "[x] ", "[ ] ") & ac2fPad(m_shName(i), 26) & _
                    IIf(m_shDrawn(i) Is Nothing, " qty " & m_shQty(i), "") & vbCrLf
            End If
        Next i
        If m_shN > 8 Then s = s & "  ... and " & (m_shN - 8) & " more" & vbCrLf
        s = s & vbCrLf & "Gap " & ac2fNumStr(m_gap) & "  margins " & ac2fNumStr(m_mL) & "/" & _
            ac2fNumStr(m_mR) & "/" & ac2fNumStr(m_mT) & "/" & ac2fNumStr(m_mB) & _
            "  rotation " & ac2fNumStr(m_rot) & "  precision " & ac2fNumStr(m_res) & _
            "  optimise " & ac2fNumStr(m_time) & " s" & vbCrLf & vbCrLf
        s = s & "Enter = nest     R = report only     S = all settings" & vbCrLf
        s = s & "T = type the sheets, e.g.  T 3050x2050x0; 2440x1220x2"
        cmd = InputBox(s, ac2fTitle(CAPTION_), "")
        If StrPtr(cmd) = 0 Then Exit Function
        cmd = Trim$(cmd)
        t = UCase$(Left$(cmd, 1))
        If Len(cmd) = 0 Then
            apply = True
            ac2fNSAskText = True
            Exit Function
        ElseIf t = "R" And Len(cmd) = 1 Then
            apply = False
            ac2fNSAskText = True
            Exit Function
        ElseIf t = "S" And Len(cmd) = 1 Then
            ac2fSettings.ac2fSettings
            ac2fNSLoad
        ElseIf t = "T" Then
            ac2fNSTypedSheets Mid$(cmd, 2)
        End If
    Loop
End Function

' "3050x2050x0; 2440x1220x2" replaces the list sheets; offcuts stay.
Private Sub ac2fNSTypedSheets(ByVal txt As String)
    Dim parts() As String, f() As String
    Dim i As Long, n As Long
    Dim keepD() As Shape, kName() As String, kW() As Double, kH() As Double
    Dim w As Double, h As Double, q As Long

    ReDim keepD(0 To m_shN): ReDim kName(0 To m_shN)
    ReDim kW(0 To m_shN): ReDim kH(0 To m_shN)
    For i = 0 To m_shN - 1
        If Not m_shDrawn(i) Is Nothing Then
            Set keepD(n) = m_shDrawn(i)
            kName(n) = m_shName(i): kW(n) = m_shW(i): kH(n) = m_shH(i)
            n = n + 1
        End If
    Next i
    m_shN = 0
    parts = Split(Replace$(LCase$(txt), "*", "x"), ";")
    For i = 0 To UBound(parts)
        f = Split(Trim$(parts(i)) & "xx", "x")
        w = ac2fParseNum(f(0), 0)
        h = ac2fParseNum(f(1), 0)
        q = CLng(ac2fParseNum(f(2), 0))
        If w > 0 And h > 0 Then ac2fNSAddSheet True, "Sheet " & ac2fNSDims(w, h), w, h, q, Nothing
    Next i
    For i = 0 To n - 1
        ac2fNSAddSheet True, kName(i), kW(i), kH(i), 1, keepD(i)
    Next i
    ac2fNSSaveSheets
End Sub

'---------------------------------------------------------------------
' Files
'---------------------------------------------------------------------
Private Function ac2fNSTemp(ByVal nm As String) As String
    Dim t As String
    t = Environ$("TEMP")
    If Len(t) = 0 Then t = Environ$("TMP")
    If Len(t) = 0 Then t = "C:\Windows\Temp"
    If Right$(t, 1) = "\" Then t = Left$(t, Len(t) - 1)
    ac2fNSTemp = t & "\" & nm
End Function

Private Function ac2fNSWriteUtf8(ByVal path As String, ByVal txt As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
    ac2fNSWriteUtf8 = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
End Function

Private Function ac2fNSReadUtf8(ByVal path As String) As String
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ac2fNSReadUtf8 = st.ReadText(-1)
    st.Close
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
End Function

'---------------------------------------------------------------------
' The window scripts, written out each run so they always match this
' module (1 settings, 2 preview). Generated from tools/nest_window.ps1
' and tools/nest_preview.ps1 by tools/ps2vba.py.
'---------------------------------------------------------------------
Private Function ac2fNSWriteScript(ByVal which As Long, ByVal f As Integer) As Boolean
    Dim ok As Boolean
    If which = 1 Then
        ok = ac2fNSPsSet1(f)
        If ok Then ok = ac2fNSPsSet2(f)
        If ok Then ok = ac2fNSPsSet3(f)
    ElseIf which = 2 Then
        ok = ac2fNSPsPrev1(f)
        If ok Then ok = ac2fNSPsPrev2(f)
    End If
    ac2fNSWriteScript = ok
End Function

Private Function ac2fNSPsSet1(ByVal f As Integer) As Boolean
    On Error GoTo Fail
    Print #f, "param([string]$In, [string]$Out)"
    Print #f, "$ErrorActionPreference = 'Stop'"
    Print #f, "Add-Type -AssemblyName System.Windows.Forms"
    Print #f, "Add-Type -AssemblyName System.Drawing"
    Print #f, "[System.Windows.Forms.Application]::EnableVisualStyles()"
    Print #f, "$inv = [System.Globalization.CultureInfo]::InvariantCulture"
    Print #f, "$all = [System.IO.File]::ReadAllLines($In, [System.Text.Encoding]::UTF8)"
    Print #f, "$vals = @{}"
    Print #f, "$sheets = New-Object System.Collections.Generic.List[object]"
    Print #f, "$presets = New-Object System.Collections.Specialized.OrderedDictionary"
    Print #f, "$title = 'Nesting'"
    Print #f, "$summary = ''"
    Print #f, "$pageW = ''"
    Print #f, "$pageH = ''"
    Print #f, "foreach ($ln in $all) {"
    Print #f, "  $p = $ln.Split('|')"
    Print #f, "  switch ($p[0]) {"
    Print #f, "    'T' { $title = $p[1] }"
    Print #f, "    'P' { $summary = $p[1] }"
    Print #f, "    'G' { if ($p.Length -ge 3) { $pageW = $p[1]; $pageH = $p[2] } }"
    Print #f, "    'K' { $vals[$p[1]] = $p[2] }"
    Print #f, "    'S' { if ($p.Length -ge 7) { $sheets.Add(@($p[1], $p[2], $p[3], $p[4], $p[5], $p[6])) } }"
    Print #f, "    'R' { if ($p.Length -ge 3) { $presets[$p[1]] = $p[2] } }"
    Print #f, "  }"
    Print #f, "}"
    Print #f, ""
    Print #f, "# key, label, kind, options, tab, help"
    Print #f, "$fields = @("
    Print #f, "  @('gap', 'Part gap (mm)', 'num', '', 0, 'Clear space between neighbouring parts. A router needs at least its bit diameter, a laser its kerf plus a little.'),"
    Print #f, "  @('mL', 'Margin left (mm)', 'num', '', 0, 'Clear strip along the left edge of every sheet.'),"
    Print #f, "  @('mR', 'Margin right (mm)', 'num', '', 0, 'Clear strip along the right edge of every sheet.'),"
    Print #f, "  @('mT', 'Margin top (mm)', 'num', '', 0, 'Clear strip along the top edge of every sheet.'),"
    Print #f, "  @('mB', 'Margin bottom (mm)', 'num', '', 0, 'Clear strip along the bottom edge of every sheet.'),"
    Print #f, "  @('rot', 'Rotation', 'combo', '0=Keep as drawn|180=Half turn only|90=Every 90 degrees|45=Every 45 degrees|30=Every 30 degrees|15=Every 15 degrees', 0, 'The turns parts may take. Finer steps nest tighter but take longer.'),"
    Print #f, "  @('copies', 'Copies of each part', 'int', '', 0, 'Every selected part is nested this many times. Copies are only created when they find room.'),"
    Print #f, "  @('time', 'Optimise for (seconds)', 'num', '', 0, 'Keeps trying other orders and strategies this long and keeps the best layout. 0 = a single quick pass.'),"
    Print #f, "  @('inner', 'Shapes inside a shape move with it', 'chk', '', 0, 'A shape lying inside another (a counter, a hole, an engraving) is kept with the outer shape as one part.'),"
    Print #f, "  @('holes', 'Small parts may go into holes', 'chk', '', 0, 'Inner contours are cut out, so other parts may be nested inside them. Leave off when inner shapes are engravings.'),"
    Print #f, "  @('om', 'Offcut edge (mm)', 'num', '', 1, 'Clear space kept from the edge of a drawn offcut sheet.'),"
    Print #f, "  @('mirror', 'Parts may be mirrored', 'chk', '', 1, 'Allows mirrored copies. Leave off for printed, coated or filmed material.'),"
    Print #f, "  @('corner', 'Start corner', 'combo', '1=Bottom left|2=Bottom right|3=Top left|4=Top right', 1, 'The corner of the sheet the packing starts from.'),"
    Print #f, "  @('dir', 'Fill direction', 'combo', '0=Rows (fill along the width)|1=Columns (fill along the height)', 1, 'Rows keep the used material short in height; columns keep it short in width.'),"
    Print #f, "  @('strat', 'Placement', 'combo', '0=Automatic (tries both)|1=Lowest top edge|2=Bottom-left', 1, 'How each part chooses its place. Automatic tries both and keeps the better layout.'),"
    Print #f, "  @('res', 'Precision (mm)', 'combo', '0.5=0.5 mm (finest, slowest)|1=1 mm|2=2 mm|3=3 mm|5=5 mm (fastest)', 1, 'Grid size used to test shapes. Finer is tighter and slower. The gap is always kept.'),"
    Print #f, "  @('group', 'Group the result', 'combo', '0=No grouping|1=Each part|2=Each sheet|3=Everything', 1, 'How placed parts are grouped once nesting is done.'),"
    Print #f, "  @('outline', 'Draw sheet outlines', 'chk', '', 1, 'Draws a grey outline for every sheet used.'),"
    Print #f, "  @('start', 'Set start points afterwards', 'chk', '', 1, 'Applies the saved start point settings (main menu) to every placed part.'),"
    Print #f, "  @('sgap', 'Space between sheets (mm)', 'num', '', 1, 'Gap between the sheets drawn on the page. Layout only.')"
    Print #f, ")"
    Print #f, ""
    Print #f, "$f = New-Object System.Windows.Forms.Form"
    Print #f, "$f.Text = $title"
    Print #f, "$f.StartPosition = 'CenterScreen'"
    Print #f, "$f.Size = New-Object System.Drawing.Size(1060, 680)"
    Print #f, "$f.MinimumSize = New-Object System.Drawing.Size(880, 560)"
    Print #f, "$f.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)"
    Print #f, "$f.TopMost = $true"
    Print #f, "$f.KeyPreview = $true"
    Print #f, "$tip = New-Object System.Windows.Forms.ToolTip"
    Print #f, "$tip.AutoPopDelay = 20000"
    Print #f, ""
    Print #f, "$root = New-Object System.Windows.Forms.TableLayoutPanel"
    Print #f, "$root.Dock = 'Fill'"
    Print #f, "$root.ColumnCount = 2"
    Print #f, "$root.RowCount = 3"
    Print #f, "$root.Padding = New-Object System.Windows.Forms.Padding(10)"
    Print #f, "[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 57)))"
    Print #f, "[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 43)))"
    Print #f, "[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))"
    Print #f, "[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))"
    Print #f, "[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))"
    Print #f, ""
    Print #f, "$head = New-Object System.Windows.Forms.Label"
    Print #f, "$head.Text = $summary"
    Print #f, "$head.AutoSize = $true"
    Print #f, "$head.MaximumSize = New-Object System.Drawing.Size(1000, 0)"
    Print #f, "$head.Margin = New-Object System.Windows.Forms.Padding(2, 0, 2, 8)"
    Print #f, "$root.Controls.Add($head, 0, 0)"
    Print #f, "$root.SetColumnSpan($head, 2)"
    Print #f, ""
    Print #f, "# ---------------- sheets ----------------"
    Print #f, "$gs = New-Object System.Windows.Forms.GroupBox"
    Print #f, "$gs.Text = 'Sheets (used from the top down)'"
    Print #f, "$gs.Dock = 'Fill'"
    Print #f, "$sl = New-Object System.Windows.Forms.TableLayoutPanel"
    Print #f, "$sl.Dock = 'Fill'"
    Print #f, "$sl.ColumnCount = 1"
    Print #f, "$sl.RowCount = 4"
    Print #f, "[void]$sl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))"
    Print #f, "[void]$sl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))"
    Print #f, "[void]$sl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))"
    Print #f, "[void]$sl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))"
    Print #f, "$gs.Controls.Add($sl)"
    Print #f, ""
    Print #f, "$grid = New-Object System.Windows.Forms.DataGridView"
    Print #f, "$grid.Dock = 'Fill'"
    Print #f, "$grid.AllowUserToAddRows = $false"
    Print #f, "$grid.AllowUserToResizeRows = $false"
    Print #f, "$grid.RowHeadersVisible = $false"
    Print #f, "$grid.SelectionMode = 'FullRowSelect'"
    Print #f, "$grid.MultiSelect = $false"
    Print #f, "$grid.AutoSizeColumnsMode = 'Fill'"
    Print #f, "$grid.BackgroundColor = [System.Drawing.SystemColors]::Window"
    Print #f, "$cUse = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn"
    Print #f, "$cUse.HeaderText = 'Use'"
    Print #f, "$cUse.FillWeight = 12"
    Print #f, "$cName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn"
    Print #f, "$cName.HeaderText = 'Sheet'"
    Print #f, "$cName.FillWeight = 46"
    Print #f, "$cW = New-Object System.Windows.Forms.DataGridViewTextBoxColumn"
    Print #f, "$cW.HeaderText = 'Width'"
    Print #f, "$cW.FillWeight = 16"
    Print #f, "$cH = New-Object System.Windows.Forms.DataGridViewTextBoxColumn"
    Print #f, "$cH.HeaderText = 'Height'"
    Print #f, "$cH.FillWeight = 16"
    Print #f, "$cQ = New-Object System.Windows.Forms.DataGridViewTextBoxColumn"
    Print #f, "$cQ.HeaderText = 'Qty'"
    Print #f, "$cQ.FillWeight = 14"
    Print #f, "$cQ.ToolTipText = '0 = as many as needed'"
    Print #f, "[void]$grid.Columns.AddRange(@($cUse, $cName, $cW, $cH, $cQ))"
    Print #f, "$sl.Controls.Add($grid, 0, 0)"
    Print #f, ""
    ac2fNSPsSet1 = True
Fail:
End Function

Private Function ac2fNSPsSet2(ByVal f As Integer) As Boolean
    On Error GoTo Fail
    Print #f, "function AddRow($use, $name, $w, $h, $q, $src) {"
    Print #f, "  $i = $grid.Rows.Add(@(($use -eq '1'), $name, $w, $h, $q))"
    Print #f, "  $r = $grid.Rows[$i]"
    Print #f, "  $r.Tag = $src"
    Print #f, "  if (([string]$src).StartsWith('D')) {"
    Print #f, "    $r.Cells[2].ReadOnly = $true"
    Print #f, "    $r.Cells[3].ReadOnly = $true"
    Print #f, "    $r.Cells[4].ReadOnly = $true"
    Print #f, "    $r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkSlateBlue"
    Print #f, "  }"
    Print #f, "  return $i"
    Print #f, "}"
    Print #f, "foreach ($s in $sheets) { [void](AddRow $s[0] $s[1] $s[2] $s[3] $s[4] $s[5]) }"
    Print #f, ""
    Print #f, "function Btn([string]$t, $panel) {"
    Print #f, "  $b = New-Object System.Windows.Forms.Button"
    Print #f, "  $b.Text = $t"
    Print #f, "  $b.AutoSize = $true"
    Print #f, "  $b.Margin = New-Object System.Windows.Forms.Padding(0, 4, 6, 0)"
    Print #f, "  $panel.Controls.Add($b)"
    Print #f, "  return $b"
    Print #f, "}"
    Print #f, "function Lbl([string]$t, $panel) {"
    Print #f, "  $l = New-Object System.Windows.Forms.Label"
    Print #f, "  $l.Text = $t"
    Print #f, "  $l.AutoSize = $true"
    Print #f, "  $l.Margin = New-Object System.Windows.Forms.Padding(0, 9, 4, 0)"
    Print #f, "  $panel.Controls.Add($l)"
    Print #f, "  return $l"
    Print #f, "}"
    Print #f, ""
    Print #f, "# a common size in one click"
    Print #f, "$sa = New-Object System.Windows.Forms.FlowLayoutPanel"
    Print #f, "$sa.Dock = 'Fill'"
    Print #f, "$sa.AutoSize = $true"
    Print #f, "$sa.WrapContents = $true"
    Print #f, "[void](Lbl 'Add a sheet:' $sa)"
    Print #f, "$cSize = New-Object System.Windows.Forms.ComboBox"
    Print #f, "$cSize.DropDownStyle = 'DropDownList'"
    Print #f, "$cSize.Width = 150"
    Print #f, "$cSize.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)"
    Print #f, "foreach ($z in @('3050 x 2050', '3050 x 1500', '3000 x 1500', '2440 x 1220', '2500 x 1250', '2000 x 1000', '1220 x 2440')) { [void]$cSize.Items.Add($z) }"
    Print #f, "if ($pageW -ne '' -and $pageH -ne '') {"
    Print #f, "  $pw = 0.0; $ph = 0.0"
    Print #f, "  if ([double]::TryParse($pageW, [System.Globalization.NumberStyles]::Float, $inv, [ref]$pw) -and"
    Print #f, "      [double]::TryParse($pageH, [System.Globalization.NumberStyles]::Float, $inv, [ref]$ph)) {"
    Print #f, "    [void]$cSize.Items.Add(('Page ' + [Math]::Round($pw).ToString($inv) + ' x ' + [Math]::Round($ph).ToString($inv)))"
    Print #f, "  }"
    Print #f, "}"
    Print #f, "$cSize.SelectedIndex = 0"
    Print #f, "$sa.Controls.Add($cSize)"
    Print #f, "$bAdd = Btn 'Add' $sa"
    Print #f, "$tip.SetToolTip($cSize, 'Pick a common sheet size and press Add. The size can be edited in the list afterwards.')"
    Print #f, "$sl.Controls.Add($sa, 0, 1)"
    Print #f, ""
    Print #f, "$sb = New-Object System.Windows.Forms.FlowLayoutPanel"
    Print #f, "$sb.Dock = 'Fill'"
    Print #f, "$sb.AutoSize = $true"
    Print #f, "$sb.WrapContents = $true"
    Print #f, "$bDel = Btn 'Remove' $sb"
    Print #f, "$bUp = Btn 'Up' $sb"
    Print #f, "$bDown = Btn 'Down' $sb"
    Print #f, "$bAllOn = Btn 'Use all' $sb"
    Print #f, "$bAllOff = Btn 'Use none' $sb"
    Print #f, "[void](Lbl 'Qty for all:' $sb)"
    Print #f, "$tq = New-Object System.Windows.Forms.TextBox"
    Print #f, "$tq.Width = 50"
    Print #f, "$tq.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)"
    Print #f, "$sb.Controls.Add($tq)"
    Print #f, "$bQ = Btn 'Apply' $sb"
    Print #f, "$sl.Controls.Add($sb, 0, 2)"
    Print #f, ""
    Print #f, "$hint = New-Object System.Windows.Forms.Label"
    Print #f, "$hint.AutoSize = $true"
    Print #f, "$hint.MaximumSize = New-Object System.Drawing.Size(560, 0)"
    Print #f, "$hint.ForeColor = [System.Drawing.Color]::DimGray"
    Print #f, "$hint.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)"
    Print #f, "$hint.Text = 'Qty 0 = as many as needed. Drawn offcuts (blue) are shapes named ""offcut"" or ""sheet"" selected with the parts; the main menu can mark them. Untick a sheet to leave it out.'"
    Print #f, "$sl.Controls.Add($hint, 0, 3)"
    Print #f, "$root.Controls.Add($gs, 0, 1)"
    Print #f, ""
    Print #f, "$bAdd.Add_Click({"
    Print #f, "  $t = [string]$cSize.SelectedItem"
    Print #f, "  $m = [regex]::Match($t, '(\d+) x (\d+)')"
    Print #f, "  if ($m.Success) {"
    Print #f, "    $w = $m.Groups[1].Value"
    Print #f, "    $h = $m.Groups[2].Value"
    Print #f, "    $i = AddRow '1' ('Sheet ' + $w + ' x ' + $h) $w $h '0' 'L'"
    Print #f, "    $grid.CurrentCell = $grid.Rows[$i].Cells[1]"
    Print #f, "  }"
    Print #f, "})"
    Print #f, "$bDel.Add_Click({"
    Print #f, "  if ($grid.CurrentRow -ne $null) { $grid.Rows.RemoveAt($grid.CurrentRow.Index) }"
    Print #f, "})"
    Print #f, "function MoveRow([int]$d) {"
    Print #f, "  if ($grid.CurrentRow -eq $null) { return }"
    Print #f, "  $i = $grid.CurrentRow.Index"
    Print #f, "  $j = $i + $d"
    Print #f, "  if ($j -lt 0 -or $j -ge $grid.Rows.Count) { return }"
    Print #f, "  $r = $grid.Rows[$i]"
    Print #f, "  $grid.Rows.RemoveAt($i)"
    Print #f, "  $grid.Rows.Insert($j, $r)"
    Print #f, "  $grid.CurrentCell = $grid.Rows[$j].Cells[1]"
    Print #f, "}"
    Print #f, "$bUp.Add_Click({ MoveRow -1 })"
    Print #f, "$bDown.Add_Click({ MoveRow 1 })"
    Print #f, "$bAllOn.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells[0].Value = $true } })"
    Print #f, "$bAllOff.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells[0].Value = $false } })"
    Print #f, "$bQ.Add_Click({"
    Print #f, "  $v = 0"
    Print #f, "  if ([int]::TryParse($tq.Text.Trim(), [ref]$v) -and $v -ge 0) {"
    Print #f, "    [void]$grid.EndEdit()"
    Print #f, "    foreach ($r in $grid.Rows) { if (-not ([string]$r.Tag).StartsWith('D')) { $r.Cells[4].Value = [string]$v } }"
    Print #f, "  } else { [void]$tq.Focus() }"
    Print #f, "})"
    Print #f, ""
    Print #f, "# ---------------- settings: two tabs ----------------"
    Print #f, "$tabs = New-Object System.Windows.Forms.TabControl"
    Print #f, "$tabs.Dock = 'Fill'"
    Print #f, "$tls = @()"
    Print #f, "foreach ($tn in @('Main', 'More')) {"
    Print #f, "  $tp = New-Object System.Windows.Forms.TabPage"
    Print #f, "  $tp.Text = $tn"
    Print #f, "  $tp.Padding = New-Object System.Windows.Forms.Padding(6)"
    Print #f, "  $tp.AutoScroll = $true"
    Print #f, "  $tl = New-Object System.Windows.Forms.TableLayoutPanel"
    Print #f, "  $tl.ColumnCount = 2"
    Print #f, "  $tl.AutoSize = $true"
    Print #f, "  $tl.Dock = 'Top'"
    Print #f, "  [void]$tl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 55)))"
    Print #f, "  [void]$tl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 45)))"
    Print #f, "  $tp.Controls.Add($tl)"
    Print #f, "  $tabs.TabPages.Add($tp)"
    Print #f, "  $tls += $tl"
    Print #f, "}"
    Print #f, "$ctl = @{}"
    Print #f, "$rows = @(0, 0)"
    Print #f, "foreach ($fd in $fields) {"
    Print #f, "  $key = $fd[0]"
    Print #f, "  $tl = $tls[$fd[4]]"
    Print #f, "  $row = $rows[$fd[4]]"
    Print #f, "  $v = ''"
    Print #f, "  if ($vals.ContainsKey($key)) { $v = $vals[$key] }"
    Print #f, "  if ($fd[2] -eq 'chk') {"
    Print #f, "    $c = New-Object System.Windows.Forms.CheckBox"
    Print #f, "    $c.Text = $fd[1]"
    Print #f, "    $c.AutoSize = $true"
    Print #f, "    $c.Checked = ($v -eq '1')"
    Print #f, "    $c.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 2)"
    Print #f, "    $tl.Controls.Add($c, 0, $row)"
    Print #f, "    $tl.SetColumnSpan($c, 2)"
    Print #f, "    $tip.SetToolTip($c, $fd[5])"
    Print #f, "  } else {"
    Print #f, "    $l = New-Object System.Windows.Forms.Label"
    Print #f, "    $l.Text = $fd[1]"
    Print #f, "    $l.AutoSize = $true"
    Print #f, "    $l.Margin = New-Object System.Windows.Forms.Padding(3, 7, 3, 2)"
    Print #f, "    $tl.Controls.Add($l, 0, $row)"
    Print #f, "    $tip.SetToolTip($l, $fd[5])"
    Print #f, "    if ($fd[2] -eq 'combo') {"
    Print #f, "      $c = New-Object System.Windows.Forms.ComboBox"
    Print #f, "      $c.DropDownStyle = 'DropDownList'"
    Print #f, "      $c.Dock = 'Fill'"
    Print #f, "      $codes = New-Object System.Collections.Generic.List[string]"
    Print #f, "      foreach ($o in $fd[3].Split('|')) {"
    Print #f, "        $k = $o.IndexOf('=')"
    Print #f, "        $codes.Add($o.Substring(0, $k))"
    Print #f, "        [void]$c.Items.Add($o.Substring($k + 1))"
    Print #f, "      }"
    Print #f, "      $c.Tag = $codes"
    Print #f, "      $sel = 0"
    Print #f, "      for ($i = 0; $i -lt $codes.Count; $i++) {"
    Print #f, "        $a = 0.0; $b = 0.0"
    Print #f, "        if ($codes[$i] -eq $v) { $sel = $i }"
    Print #f, "        elseif ([double]::TryParse($codes[$i], [System.Globalization.NumberStyles]::Float, $inv, [ref]$a) -and"
    Print #f, "                [double]::TryParse($v, [System.Globalization.NumberStyles]::Float, $inv, [ref]$b) -and"
    Print #f, "                [math]::Abs($a - $b) -lt 0.0001) { $sel = $i }"
    Print #f, "      }"
    Print #f, "      $c.SelectedIndex = $sel"
    Print #f, "    } else {"
    Print #f, "      $c = New-Object System.Windows.Forms.TextBox"
    Print #f, "      $c.Text = $v"
    Print #f, "      $c.Dock = 'Fill'"
    Print #f, "    }"
    Print #f, "    $c.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 2)"
    Print #f, "    $tl.Controls.Add($c, 1, $row)"
    Print #f, "    $tip.SetToolTip($c, $fd[5])"
    Print #f, "  }"
    Print #f, "  $ctl[$key] = $c"
    Print #f, "  $rows[$fd[4]] = $row + 1"
    Print #f, "}"
    Print #f, "$tls[0].RowCount = $rows[0]"
    Print #f, "$tls[1].RowCount = $rows[1]"
    Print #f, "$root.Controls.Add($tabs, 1, 1)"
    Print #f, ""
    ac2fNSPsSet2 = True
Fail:
End Function

Private Function ac2fNSPsSet3(ByVal f As Integer) As Boolean
    On Error GoTo Fail
    Print #f, "function GetVal([string]$key) {"
    Print #f, "  $c = $ctl[$key]"
    Print #f, "  if ($c -is [System.Windows.Forms.CheckBox]) { if ($c.Checked) { return '1' } else { return '0' } }"
    Print #f, "  if ($c -is [System.Windows.Forms.ComboBox]) { return $c.Tag[$c.SelectedIndex] }"
    Print #f, "  return $c.Text.Trim().Replace(',', '.')"
    Print #f, "}"
    Print #f, "function SetVal([string]$key, [string]$v) {"
    Print #f, "  $c = $ctl[$key]"
    Print #f, "  if ($c -eq $null) { return }"
    Print #f, "  if ($c -is [System.Windows.Forms.CheckBox]) { $c.Checked = ($v -eq '1'); return }"
    Print #f, "  if ($c -is [System.Windows.Forms.ComboBox]) {"
    Print #f, "    for ($i = 0; $i -lt $c.Tag.Count; $i++) { if ($c.Tag[$i] -eq $v) { $c.SelectedIndex = $i } }"
    Print #f, "    return"
    Print #f, "  }"
    Print #f, "  $c.Text = $v"
    Print #f, "}"
    Print #f, "function Blob() {"
    Print #f, "  $parts = @()"
    Print #f, "  foreach ($fd in $fields) { $parts += ($fd[0] + '=' + (GetVal $fd[0])) }"
    Print #f, "  return ($parts -join ';')"
    Print #f, "}"
    Print #f, ""
    Print #f, "# ---------------- bottom bar ----------------"
    Print #f, "$bar = New-Object System.Windows.Forms.TableLayoutPanel"
    Print #f, "$bar.Dock = 'Fill'"
    Print #f, "$bar.AutoSize = $true"
    Print #f, "$bar.ColumnCount = 2"
    Print #f, "[void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))"
    Print #f, "[void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))"
    Print #f, "$left = New-Object System.Windows.Forms.FlowLayoutPanel"
    Print #f, "$left.AutoSize = $true"
    Print #f, "$left.Dock = 'Fill'"
    Print #f, "[void](Lbl 'Preset:' $left)"
    Print #f, "$cp = New-Object System.Windows.Forms.ComboBox"
    Print #f, "$cp.Width = 170"
    Print #f, "$cp.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)"
    Print #f, "foreach ($k in $presets.Keys) { [void]$cp.Items.Add($k) }"
    Print #f, "$left.Controls.Add($cp)"
    Print #f, "$bLoad = Btn 'Load' $left"
    Print #f, "$bSave = Btn 'Save' $left"
    Print #f, "$bPDel = Btn 'Delete' $left"
    Print #f, "$tip.SetToolTip($cp, 'Type a name and press Save to keep every setting under it, e.g. Laser 3mm or CNC 6mm.')"
    Print #f, "$right = New-Object System.Windows.Forms.FlowLayoutPanel"
    Print #f, "$right.AutoSize = $true"
    Print #f, "$right.Dock = 'Fill'"
    Print #f, "$right.FlowDirection = 'RightToLeft'"
    Print #f, "$bCancel = Btn 'Cancel' $right"
    Print #f, "$bNest = Btn 'Nest now' $right"
    Print #f, "$bPreview = Btn 'Preview' $right"
    Print #f, "$bPreview.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)"
    Print #f, "$tip.SetToolTip($bPreview, 'Works the layout out and shows it before anything moves. Apply it there, or come back and change a setting.')"
    Print #f, "$tip.SetToolTip($bNest, 'Works the layout out and places the parts straight away. Ctrl+Z undoes it.')"
    Print #f, "$bar.Controls.Add($left, 0, 0)"
    Print #f, "$bar.Controls.Add($right, 1, 0)"
    Print #f, "$root.Controls.Add($bar, 0, 2)"
    Print #f, "$root.SetColumnSpan($bar, 2)"
    Print #f, "$f.Controls.Add($root)"
    Print #f, "$f.AcceptButton = $bPreview"
    Print #f, "$f.CancelButton = $bCancel"
    Print #f, ""
    Print #f, "$bLoad.Add_Click({"
    Print #f, "  $n = $cp.Text.Trim()"
    Print #f, "  if ($presets.Contains($n)) {"
    Print #f, "    foreach ($kv in ([string]$presets[$n]).Split(';')) {"
    Print #f, "      $i = $kv.IndexOf('=')"
    Print #f, "      if ($i -gt 0) { SetVal $kv.Substring(0, $i) $kv.Substring($i + 1) }"
    Print #f, "    }"
    Print #f, "  }"
    Print #f, "})"
    Print #f, "$bSave.Add_Click({"
    Print #f, "  $n = $cp.Text.Trim().Replace('|', ' ').Replace(';', ' ').Replace('=', ' ')"
    Print #f, "  if ($n.Length -eq 0) { [void]$cp.Focus(); return }"
    Print #f, "  if (-not $presets.Contains($n)) { [void]$cp.Items.Add($n) }"
    Print #f, "  $presets[$n] = Blob"
    Print #f, "  $cp.Text = $n"
    Print #f, "})"
    Print #f, "$bPDel.Add_Click({"
    Print #f, "  $n = $cp.Text.Trim()"
    Print #f, "  if ($presets.Contains($n)) { $presets.Remove($n); $cp.Items.Remove($n); $cp.Text = '' }"
    Print #f, "})"
    Print #f, ""
    Print #f, "$script:act = 'CANCEL'"
    Print #f, "function Check() {"
    Print #f, "  [void]$grid.EndEdit()"
    Print #f, "  foreach ($fd in $fields) {"
    Print #f, "    if ($fd[2] -eq 'num' -or $fd[2] -eq 'int') {"
    Print #f, "      $t = GetVal $fd[0]"
    Print #f, "      $x = 0.0"
    Print #f, "      if (-not [double]::TryParse($t, [System.Globalization.NumberStyles]::Float, $inv, [ref]$x) -or $x -lt 0) {"
    Print #f, "        $tabs.SelectedIndex = $fd[4]"
    Print #f, "        [void][System.Windows.Forms.MessageBox]::Show($f, ($fd[1] + ': enter a number of 0 or more.'), $title)"
    Print #f, "        [void]$ctl[$fd[0]].Focus()"
    Print #f, "        return $false"
    Print #f, "      }"
    Print #f, "    }"
    Print #f, "  }"
    Print #f, "  $used = 0"
    Print #f, "  foreach ($r in $grid.Rows) {"
    Print #f, "    if ($r.Cells[0].Value -eq $true) {"
    Print #f, "      $w = 0.0; $h = 0.0; $q = 0"
    Print #f, "      if (-not [double]::TryParse(([string]$r.Cells[2].Value).Replace(',', '.'), [System.Globalization.NumberStyles]::Float, $inv, [ref]$w) -or $w -le 0 -or"
    Print #f, "          -not [double]::TryParse(([string]$r.Cells[3].Value).Replace(',', '.'), [System.Globalization.NumberStyles]::Float, $inv, [ref]$h) -or $h -le 0 -or"
    Print #f, "          -not [int]::TryParse(([string]$r.Cells[4].Value).Trim(), [ref]$q) -or $q -lt 0) {"
    Print #f, "        [void][System.Windows.Forms.MessageBox]::Show($f, ('Row ' + ($r.Index + 1) + ': width and height must be above 0, qty 0 or more.'), $title)"
    Print #f, "        $grid.CurrentCell = $r.Cells[2]"
    Print #f, "        return $false"
    Print #f, "      }"
    Print #f, "      $used++"
    Print #f, "    }"
    Print #f, "  }"
    Print #f, "  if ($used -eq 0) {"
    Print #f, "    [void][System.Windows.Forms.MessageBox]::Show($f, 'Tick at least one sheet.', $title)"
    Print #f, "    return $false"
    Print #f, "  }"
    Print #f, "  return $true"
    Print #f, "}"
    Print #f, "$bPreview.Add_Click({ if (Check) { $script:act = 'PREVIEW'; $f.Close() } })"
    Print #f, "$bNest.Add_Click({ if (Check) { $script:act = 'NEST'; $f.Close() } })"
    Print #f, "$bCancel.Add_Click({ $script:act = 'CANCEL'; $f.Close() })"
    Print #f, ""
    Print #f, "[void]$f.ShowDialog()"
    Print #f, ""
    Print #f, "$o = New-Object System.Collections.Generic.List[string]"
    Print #f, "$o.Add('A|' + $script:act)"
    Print #f, "if ($script:act -ne 'CANCEL') {"
    Print #f, "  foreach ($fd in $fields) { $o.Add('K|' + $fd[0] + '|' + (GetVal $fd[0])) }"
    Print #f, "  foreach ($r in $grid.Rows) {"
    Print #f, "    $u = '0'"
    Print #f, "    if ($r.Cells[0].Value -eq $true) { $u = '1' }"
    Print #f, "    $nm = ([string]$r.Cells[1].Value).Replace('|', ' ')"
    Print #f, "    $o.Add('S|' + $u + '|' + $nm + '|' + ([string]$r.Cells[2].Value).Replace(',', '.') + '|' +"
    Print #f, "           ([string]$r.Cells[3].Value).Replace(',', '.') + '|' + ([string]$r.Cells[4].Value) + '|' + [string]$r.Tag)"
    Print #f, "  }"
    Print #f, "}"
    Print #f, "foreach ($k in $presets.Keys) { $o.Add('R|' + $k + '|' + $presets[$k]) }"
    Print #f, "[System.IO.File]::WriteAllLines($Out, $o, (New-Object System.Text.UTF8Encoding($true)))"
    ac2fNSPsSet3 = True
Fail:
End Function

Private Function ac2fNSPsPrev1(ByVal f As Integer) As Boolean
    On Error GoTo Fail
    Print #f, "param([string]$In, [string]$Out)"
    Print #f, "$ErrorActionPreference = 'Stop'"
    Print #f, "Add-Type -AssemblyName System.Windows.Forms"
    Print #f, "Add-Type -AssemblyName System.Drawing"
    Print #f, "[System.Windows.Forms.Application]::EnableVisualStyles()"
    Print #f, "# written by VBA Print #, so in the code page of this machine"
    Print #f, "$all = [System.IO.File]::ReadAllLines($In, [System.Text.Encoding]::Default)"
    Print #f, "$title = 'Nesting - preview'"
    Print #f, "$report = New-Object System.Collections.Generic.List[string]"
    Print #f, "$sheets = New-Object System.Collections.Generic.List[object]"
    Print #f, "$byNo = @{}"
    Print #f, "$byPart = @{}"
    Print #f, ""
    Print #f, "# ""x y x y ..."" in tenths of a millimetre -> PointF[] in millimetres"
    Print #f, "function Pts([string]$s) {"
    Print #f, "  $a = $s.Split(' ')"
    Print #f, "  $n = [int][Math]::Floor($a.Length / 2)"
    Print #f, "  $r = New-Object 'System.Drawing.PointF[]' $n"
    Print #f, "  for ($i = 0; $i -lt $n; $i++) {"
    Print #f, "    $r[$i] = [System.Drawing.PointF]::new(([single]([int]$a[2 * $i]) / 10), ([single]([int]$a[2 * $i + 1]) / 10))"
    Print #f, "  }"
    Print #f, "  return ,$r"
    Print #f, "}"
    Print #f, ""
    Print #f, "foreach ($ln in $all) {"
    Print #f, "  $p = $ln.Split('|')"
    Print #f, "  switch ($p[0]) {"
    Print #f, "    'T' { $title = $p[1] }"
    Print #f, "    'R' { $report.Add($p[1]) }"
    Print #f, "    'S' {"
    Print #f, "      if ($p.Length -ge 12) {"
    Print #f, "        $s = New-Object PSObject -Property @{"
    Print #f, "          No = [int]$p[1]; Name = $p[2]; Parts = [int]$p[3]"
    Print #f, "          X0 = [int]$p[4] / 10; Y0 = [int]$p[5] / 10; X1 = [int]$p[6] / 10; Y1 = [int]$p[7] / 10"
    Print #f, "          UX0 = [int]$p[8] / 10; UY0 = [int]$p[9] / 10; UX1 = [int]$p[10] / 10; UY1 = [int]$p[11] / 10"
    Print #f, "          Outl = (New-Object System.Collections.Generic.List[object])"
    Print #f, "          Items = (New-Object System.Collections.Generic.List[object])"
    Print #f, "        }"
    Print #f, "        $sheets.Add($s)"
    Print #f, "        $byNo[$s.No] = $s"
    Print #f, "      }"
    Print #f, "    }"
    Print #f, "    'O' {"
    Print #f, "      if ($p.Length -ge 4 -and $byNo.ContainsKey([int]$p[1])) {"
    Print #f, "        $pts = Pts $p[3]"
    Print #f, "        if ($pts.Length -ge 3) { $byNo[[int]$p[1]].Outl.Add($pts) }"
    Print #f, "      }"
    Print #f, "    }"
    Print #f, "    'P' {"
    Print #f, "      if ($p.Length -ge 6 -and $byNo.ContainsKey([int]$p[1])) {"
    Print #f, "        $pk = [int]$p[2]"
    Print #f, "        $it = $byPart[$pk]"
    Print #f, "        if ($it -eq $null) {"
    Print #f, "          $it = New-Object PSObject -Property @{"
    Print #f, "            Cls = [int]$p[3]"
    Print #f, "            Closed = (New-Object System.Collections.Generic.List[object])"
    Print #f, "            Open = (New-Object System.Collections.Generic.List[object])"
    Print #f, "          }"
    Print #f, "          $byPart[$pk] = $it"
    Print #f, "          $byNo[[int]$p[1]].Items.Add($it)"
    Print #f, "        }"
    Print #f, "        $pts = Pts $p[5]"
    Print #f, "        if ($p[4] -eq '1') {"
    Print #f, "          if ($pts.Length -ge 2) { $it.Open.Add($pts) }"
    Print #f, "        } elseif ($pts.Length -ge 3) {"
    Print #f, "          $it.Closed.Add($pts)"
    Print #f, "        }"
    Print #f, "      }"
    Print #f, "    }"
    Print #f, "  }"
    Print #f, "}"
    Print #f, ""
    Print #f, "$f = New-Object System.Windows.Forms.Form"
    Print #f, "$f.Text = $title"
    Print #f, "$f.StartPosition = 'CenterScreen'"
    Print #f, "$f.Size = New-Object System.Drawing.Size(1180, 760)"
    Print #f, "$f.MinimumSize = New-Object System.Drawing.Size(820, 520)"
    Print #f, "$f.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)"
    Print #f, "$f.TopMost = $true"
    Print #f, "$f.KeyPreview = $true"
    Print #f, ""
    Print #f, "$root = New-Object System.Windows.Forms.TableLayoutPanel"
    Print #f, "$root.Dock = 'Fill'"
    Print #f, "$root.ColumnCount = 2"
    Print #f, "$root.RowCount = 2"
    Print #f, "$root.Padding = New-Object System.Windows.Forms.Padding(10)"
    Print #f, "[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 68)))"
    Print #f, "[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 32)))"
    Print #f, "[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))"
    Print #f, "[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))"
    Print #f, ""
    Print #f, "$pb = New-Object System.Windows.Forms.PictureBox"
    Print #f, "$pb.Dock = 'Fill'"
    Print #f, "$pb.BorderStyle = 'FixedSingle'"
    Print #f, "$pb.BackColor = [System.Drawing.Color]::White"
    Print #f, "$root.Controls.Add($pb, 0, 0)"
    Print #f, ""
    Print #f, "$side = New-Object System.Windows.Forms.TableLayoutPanel"
    Print #f, "$side.Dock = 'Fill'"
    Print #f, "$side.ColumnCount = 1"
    Print #f, "$side.RowCount = 4"
    Print #f, "[void]$side.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))"
    Print #f, "[void]$side.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 35)))"
    Print #f, "[void]$side.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))"
    Print #f, "[void]$side.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 65)))"
    Print #f, "$l1 = New-Object System.Windows.Forms.Label"
    Print #f, "$l1.Text = 'Show'"
    Print #f, "$l1.AutoSize = $true"
    Print #f, "$l1.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 3)"
    Print #f, "$side.Controls.Add($l1, 0, 0)"
    Print #f, "$lst = New-Object System.Windows.Forms.ListBox"
    Print #f, "$lst.Dock = 'Fill'"
    Print #f, "$lst.IntegralHeight = $false"
    Print #f, "[void]$lst.Items.Add('All sheets')"
    Print #f, "foreach ($s in $sheets) { [void]$lst.Items.Add(('Sheet ' + $s.No + ': ' + $s.Name + ' (' + $s.Parts + ' parts)')) }"
    Print #f, "$side.Controls.Add($lst, 0, 1)"
    Print #f, "$l2 = New-Object System.Windows.Forms.Label"
    Print #f, "$l2.Text = 'Report'"
    Print #f, "$l2.AutoSize = $true"
    Print #f, "$l2.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 3)"
    Print #f, "$side.Controls.Add($l2, 0, 2)"
    Print #f, "$txt = New-Object System.Windows.Forms.TextBox"
    Print #f, "$txt.Multiline = $true"
    Print #f, "$txt.ReadOnly = $true"
    Print #f, "$txt.ScrollBars = 'Both'"
    Print #f, "$txt.WordWrap = $false"
    Print #f, "$txt.Dock = 'Fill'"
    Print #f, "$txt.Font = New-Object System.Drawing.Font('Consolas', 9)"
    Print #f, "$txt.Text = [string]::Join([Environment]::NewLine, $report.ToArray())"
    Print #f, "$side.Controls.Add($txt, 0, 3)"
    Print #f, "$root.Controls.Add($side, 1, 0)"
    Print #f, ""
    Print #f, "$bar = New-Object System.Windows.Forms.TableLayoutPanel"
    Print #f, "$bar.Dock = 'Fill'"
    Print #f, "$bar.AutoSize = $true"
    Print #f, "$bar.ColumnCount = 2"
    Print #f, "[void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))"
    Print #f, "[void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))"
    Print #f, "$note = New-Object System.Windows.Forms.Label"
    Print #f, "$note.AutoSize = $true"
    Print #f, "$note.ForeColor = [System.Drawing.Color]::DimGray"
    Print #f, "$note.Margin = New-Object System.Windows.Forms.Padding(2, 10, 2, 0)"
    Print #f, "$note.Text = 'Nothing has moved yet. Apply places the parts as shown (Ctrl+Z undoes it).'"
    Print #f, "$bar.Controls.Add($note, 0, 0)"
    Print #f, "$right = New-Object System.Windows.Forms.FlowLayoutPanel"
    Print #f, "$right.AutoSize = $true"
    Print #f, "$right.Dock = 'Fill'"
    Print #f, "$right.FlowDirection = 'RightToLeft'"
    Print #f, "function Btn([string]$t) {"
    Print #f, "  $b = New-Object System.Windows.Forms.Button"
    Print #f, "  $b.Text = $t"
    Print #f, "  $b.AutoSize = $true"
    Print #f, "  $b.Margin = New-Object System.Windows.Forms.Padding(0, 4, 6, 0)"
    Print #f, "  $right.Controls.Add($b)"
    Print #f, "  return $b"
    Print #f, "}"
    Print #f, "$bCancel = Btn 'Cancel'"
    Print #f, "$bBack = Btn 'Change settings'"
    Print #f, "$bApply = Btn 'Apply'"
    Print #f, "$bApply.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)"
    Print #f, "$bar.Controls.Add($right, 1, 0)"
    Print #f, "$root.Controls.Add($bar, 0, 1)"
    Print #f, "$root.SetColumnSpan($bar, 2)"
    Print #f, "$f.Controls.Add($root)"
    Print #f, "$f.AcceptButton = $bApply"
    Print #f, "$f.CancelButton = $bCancel"
    Print #f, ""
    ac2fNSPsPrev1 = True
Fail:
End Function

Private Function ac2fNSPsPrev2(ByVal f As Integer) As Boolean
    On Error GoTo Fail
    Print #f, "# ---------------- drawing ----------------"
    Print #f, "$colors = @("
    Print #f, "  @(166, 206, 227), @(178, 223, 138), @(251, 180, 174), @(253, 205, 172),"
    Print #f, "  @(203, 213, 232), @(244, 202, 228), @(230, 245, 201), @(255, 242, 174),"
    Print #f, "  @(204, 235, 197), @(222, 203, 228)"
    Print #f, ")"
    Print #f, "$brushes = @()"
    Print #f, "foreach ($c in $colors) { $brushes += (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, $c[0], $c[1], $c[2]))) }"
    Print #f, "$brSheet = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 236, 238, 241))"
    Print #f, "$brUse = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 250, 250, 250))"
    Print #f, "$brText = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 40, 40, 40))"
    Print #f, "$penSheet = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 120, 120, 120), 1)"
    Print #f, "$penUse = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 150, 180, 220), 1)"
    Print #f, "$penPart = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 45, 60, 85), 1)"
    Print #f, "$fLabel = New-Object System.Drawing.Font('Segoe UI', 9)"
    Print #f, ""
    Print #f, "function Render {"
    Print #f, "  try {"
    Print #f, "    $w = $pb.ClientSize.Width"
    Print #f, "    $h = $pb.ClientSize.Height"
    Print #f, "    if ($w -lt 40 -or $h -lt 40) { return }"
    Print #f, "    $bmp = New-Object System.Drawing.Bitmap($w, $h)"
    Print #f, "    $g = [System.Drawing.Graphics]::FromImage($bmp)"
    Print #f, "    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias"
    Print #f, "    $g.Clear([System.Drawing.Color]::White)"
    Print #f, "    $show = New-Object System.Collections.Generic.List[object]"
    Print #f, "    if ($lst.SelectedIndex -le 0) {"
    Print #f, "      foreach ($s in $sheets) { $show.Add($s) }"
    Print #f, "    } else {"
    Print #f, "      $show.Add($sheets[$lst.SelectedIndex - 1])"
    Print #f, "    }"
    Print #f, "    if ($show.Count -eq 0) {"
    Print #f, "      $g.DrawString('Nothing was placed.', $fLabel, $brText, [single]12, [single]12)"
    Print #f, "    } else {"
    Print #f, "      $minx = $show[0].X0; $miny = $show[0].Y0; $maxx = $show[0].X1; $maxy = $show[0].Y1"
    Print #f, "      foreach ($s in $show) {"
    Print #f, "        if ($s.X0 -lt $minx) { $minx = $s.X0 }"
    Print #f, "        if ($s.Y0 -lt $miny) { $miny = $s.Y0 }"
    Print #f, "        if ($s.X1 -gt $maxx) { $maxx = $s.X1 }"
    Print #f, "        if ($s.Y1 -gt $maxy) { $maxy = $s.Y1 }"
    Print #f, "      }"
    Print #f, "      $bw = [Math]::Max(1.0, $maxx - $minx)"
    Print #f, "      $bh = [Math]::Max(1.0, $maxy - $miny)"
    Print #f, "      $sc = [Math]::Min(($w - 30) / $bw, ($h - 50) / $bh)"
    Print #f, "      $ox = ($w - $bw * $sc) / 2 - $minx * $sc"
    Print #f, "      $oy = ($h + $bh * $sc) / 2 + 10 + $miny * $sc"
    Print #f, "      # page y points up, the screen down"
    Print #f, "      $g.TranslateTransform([single]$ox, [single]$oy)"
    Print #f, "      $g.ScaleTransform([single]$sc, [single](-$sc))"
    Print #f, "      $thin = [single](1.0 / $sc)"
    Print #f, "      $penSheet.Width = $thin"
    Print #f, "      $penUse.Width = $thin"
    Print #f, "      $penPart.Width = $thin"
    Print #f, "      foreach ($s in $show) {"
    Print #f, "        if ($s.Outl.Count -gt 0) {"
    Print #f, "          $gp = New-Object System.Drawing.Drawing2D.GraphicsPath"
    Print #f, "          $gp.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate"
    Print #f, "          foreach ($o in $s.Outl) { $gp.AddPolygon($o) }"
    Print #f, "          $g.FillPath($brUse, $gp)"
    Print #f, "          $g.DrawPath($penSheet, $gp)"
    Print #f, "          $gp.Dispose()"
    Print #f, "        } else {"
    Print #f, "          $g.FillRectangle($brSheet, [single]$s.X0, [single]$s.Y0, [single]($s.X1 - $s.X0), [single]($s.Y1 - $s.Y0))"
    Print #f, "          $g.DrawRectangle($penSheet, [single]$s.X0, [single]$s.Y0, [single]($s.X1 - $s.X0), [single]($s.Y1 - $s.Y0))"
    Print #f, "          if ($s.UX1 -gt $s.UX0 -and $s.UY1 -gt $s.UY0) {"
    Print #f, "            $g.FillRectangle($brUse, [single]$s.UX0, [single]$s.UY0, [single]($s.UX1 - $s.UX0), [single]($s.UY1 - $s.UY0))"
    Print #f, "            $g.DrawRectangle($penUse, [single]$s.UX0, [single]$s.UY0, [single]($s.UX1 - $s.UX0), [single]($s.UY1 - $s.UY0))"
    Print #f, "          }"
    Print #f, "        }"
    Print #f, "        foreach ($it in $s.Items) {"
    Print #f, "          if ($it.Closed.Count -gt 0) {"
    Print #f, "            $gp = New-Object System.Drawing.Drawing2D.GraphicsPath"
    Print #f, "            $gp.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate"
    Print #f, "            foreach ($c in $it.Closed) { $gp.AddPolygon($c) }"
    Print #f, "            $g.FillPath($brushes[$it.Cls % $brushes.Count], $gp)"
    Print #f, "            $g.DrawPath($penPart, $gp)"
    Print #f, "            $gp.Dispose()"
    Print #f, "          }"
    Print #f, "          foreach ($c in $it.Open) { $g.DrawLines($penPart, $c) }"
    Print #f, "        }"
    Print #f, "      }"
    Print #f, "      $g.ResetTransform()"
    Print #f, "      foreach ($s in $show) {"
    Print #f, "        $tx = $ox + $s.X0 * $sc"
    Print #f, "        $ty = $oy - $s.Y1 * $sc - 18"
    Print #f, "        $g.DrawString(([string]$s.No + '  ' + $s.Name), $fLabel, $brText, [single]$tx, [single]$ty)"
    Print #f, "      }"
    Print #f, "    }"
    Print #f, "    $g.Dispose()"
    Print #f, "    $old = $pb.Image"
    Print #f, "    $pb.Image = $bmp"
    Print #f, "    if ($old -ne $null) { $old.Dispose() }"
    Print #f, "  } catch {"
    Print #f, "    # a drawing problem must not cost the user the window"
    Print #f, "  }"
    Print #f, "}"
    Print #f, ""
    Print #f, "$lst.Add_SelectedIndexChanged({ Render })"
    Print #f, "$pb.Add_Resize({ Render })"
    Print #f, "$f.Add_Shown({ Render })"
    Print #f, "$lst.SelectedIndex = 0"
    Print #f, ""
    Print #f, "$script:act = 'CANCEL'"
    Print #f, "$bApply.Add_Click({ $script:act = 'APPLY'; $f.Close() })"
    Print #f, "$bBack.Add_Click({ $script:act = 'SETTINGS'; $f.Close() })"
    Print #f, "$bCancel.Add_Click({ $script:act = 'CANCEL'; $f.Close() })"
    Print #f, ""
    Print #f, "[void]$f.ShowDialog()"
    Print #f, ""
    Print #f, "$o = New-Object System.Collections.Generic.List[string]"
    Print #f, "$o.Add('A|' + $script:act)"
    Print #f, "[System.IO.File]::WriteAllLines($Out, $o, (New-Object System.Text.UTF8Encoding($true)))"
    ac2fNSPsPrev2 = True
Fail:
End Function
