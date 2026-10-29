Attribute VB_Name = "ac2fPanel"
'=====================================================================
'  ac2f pack  --  ac2fPanel
'
'  V-groove layout for aluminium composite panel (ACP) boxes.
'
'  Takes a rectangle, works out the flat sheet and draws the four
'  V-groove lines a CAM program can follow, in clockwise order with
'  the start point of each line placed so the tool enters where the
'  operator expects.
'
'  A panel longer than the plates in stock is SPLIT: the macro offers
'  ways to share the length between plates, lets you type the length
'  used from each plate, and draws one sheet per plate. Every plate
'  keeps a joint allowance on the sides that meet its neighbour, so
'  after the folds and joints are taken off, the faces add up to the
'  finished size exactly.
'
'  Macros:
'    ac2fPanelGroove - draw the sheet and its four V-grooves; splits
'                      the panel when it is longer than a plate
'    ac2fPanelSplit  - the same, but always asks about plates
'
'  LAYOUT
'  With a 100x200 rectangle and a 5 (cm) fold, "out" gives a 110x210
'  sheet whose groove lines intersect exactly on the original 100x200.
'  Lines run clockwise and span the full sheet:
'
'      1  left    x = X0+d      bottom -> top     length H
'      2  top     y = Y0+H-d    left   -> right   length W
'      3  right   x = X0+W-d    top    -> bottom  length H
'      4  bottom  y = Y0+d      right  -> left    length W
'
'  Each line is created in that order and named, so the object order
'  in the Object Manager is the cutting order.
'
'  SPLIT
'  A 4000 sign, fold 50, joint 30, plates 3200: two plates.
'
'      plate 1   cut 3200   allowances 50 | 30   face 3120
'      plate 2   cut  960   allowances 30 | 50   face  880
'                                               ---------
'                                          net  4000
'
'  The ends of the sign carry the fold, every joint carries the joint
'  allowance on BOTH plates, and the two long sides carry the fold on
'  every plate. Cut length = face + the allowance on each end.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "ACP Panel Grooves"

Private Const MAX_PANELS As Long = 100
Private Const MAX_PIECES As Long = 12        ' plates one panel may be split into
Private Const MAX_CANDS  As Long = 40        ' split plans worked out per panel
Private Const MAX_OPTS   As Long = 6         ' plans offered to the user
Private Const MIN_FACE   As Double = 100#    ' narrowest face a suggestion makes (mm)
Private Const PIECE_GAP  As Double = 100#    ' space between pieces on the page (mm)
Private Const PROMPT_ROOM As Long = 700     ' room for the option lines in a prompt

' The cut outline and the V-grooves each get a layer of their own.
Private Const LAYER_CUT    As String = "ACP Cut"
Private Const LAYER_GROOVE As String = "ACP Grooves"

'---------------------------------------------------------------------
' Registry keys (internal identifiers, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_ACP_FOLD   As String = "ACPDerzMM"
Public Const AC2F_K_ACP_DIR    As String = "ACPYon"
Public Const AC2F_K_ACP_KEEP   As String = "ACPKaynagiKoru"
Public Const AC2F_K_ACP_JOINT  As String = "ACPBirlesimMM"
Public Const AC2F_K_ACP_PLATES As String = "ACPPlakaListesi"

Public Const AC2F_DEF_ACP_FOLD   As Double = 50#   ' fold size (mm) = 5 cm
Public Const AC2F_DEF_ACP_DIR    As Long = 1       ' 1 = out, 2 = in
Public Const AC2F_DEF_ACP_KEEP   As Long = 1       ' 1 = keep source rectangle
Public Const AC2F_DEF_ACP_JOINT  As Double = 30#   ' joint allowance (mm)
Public Const AC2F_DEF_ACP_PLATES As String = "3200" ' plate lengths in stock (mm)

Public Const AC2F_ACP_OUT As Long = 1
Public Const AC2F_ACP_IN  As Long = 2

'---------------------------------------------------------------------
' Split plan. The questions are asked BEFORE anything is drawn, so
' Cancel leaves the page untouched. One row per panel, one column per
' plate; m_pn = 0 means "draw as a single sheet".
'---------------------------------------------------------------------
Private m_pn() As Long
Private m_pplate() As Double       ' plate length, index panel * MAX_PIECES + piece
Private m_pused() As Double        ' length cut from it, same index

' Layers the panel is drawn on, and whether the plates of a split panel
' are laid apart (True) instead of in place on the finished size.
Private m_lyCut As Layer
Private m_lyGrv As Layer
Private m_sep As Boolean

' Candidate plans for the panel being asked about, and the ones shown.
Private m_cc As Long
Private m_cn() As Long
Private m_cp() As Double
Private m_cu() As Double           ' cut lengths, index plan * MAX_PIECES + piece
Private m_ct() As String
Private m_oc As Long
Private m_oi() As Long

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fPanelGroove()
Attribute ac2fPanelGroove.VB_Description = "ac2f pack: Draw ACP panel V-groove lines around a rectangle (splits it across plates when it is longer than a plate)"
    ac2fPanelRun False
End Sub

Public Sub ac2fPanelSplit()
Attribute ac2fPanelSplit.VB_Description = "ac2f pack: ACP panel V-grooves, always asking how to share the length between plates"
    ac2fPanelRun True
End Sub

Private Sub ac2fPanelRun(ByVal forceSplit As Boolean)
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean
    Dim n As Long, i As Long
    Dim src() As Shape
    Dim bx() As Double, by() As Double, bw() As Double, bh() As Double
    Dim fold As Double, joint As Double
    Dim dir_ As Long
    Dim keepSrc As Boolean
    Dim made As Long, skipped As Long, ungrouped As Long
    Dim splits As Long, pieces As Long, loose As Long
    Dim rep As String
    Dim planned As Boolean, grouped As Boolean
    Dim why As String
    Dim oldLayer As Layer

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Select the rectangle to groove.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Select the rectangle to groove.", CAPTION_
        Exit Sub
    End If

    If Not ac2fAskFold(fold, dir_) Then Exit Sub
    keepSrc = (ac2fGetLng(AC2F_K_ACP_KEEP, AC2F_DEF_ACP_KEEP) <> 0)

    On Error GoTo Fail
    oldOpt = Application.Optimization
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    ' Read every bounding box up front. Creating shapes and grouping
    ' them changes the selection, so the source range must not be
    ' walked once drawing has started.
    n = sr.Count
    If n > MAX_PANELS Then n = MAX_PANELS
    ReDim src(0 To n - 1)
    ReDim bx(0 To n - 1): ReDim by(0 To n - 1)
    ReDim bw(0 To n - 1): ReDim bh(0 To n - 1)
    For i = 0 To n - 1
        Set src(i) = sr(i + 1)
        src(i).GetBoundingBox bx(i), by(i), bw(i), bh(i)
    Next i

    ' Plates and joints. Asks, draws nothing.
    planned = ac2fPanelPlanAll(n, bw, bh, fold, dir_, forceSplit, joint)
    If Not planned Then
        On Error Resume Next
        If unitChanged Then ActiveDocument.Unit = oldUnit
        On Error GoTo 0
        Exit Sub
    End If

    Application.Optimization = True
    ActiveDocument.BeginCommandGroup ac2fTitle("ACP grooves")
    grouped = True

    ' Cut outline and grooves on layers of their own (grooves on top).
    Set oldLayer = ActiveLayer
    Set m_lyCut = ac2fPanelLayer(LAYER_CUT)
    If m_lyCut Is Nothing Then Set m_lyCut = ActiveLayer
    Set m_lyGrv = ac2fPanelLayer(LAYER_GROOVE)
    If m_lyGrv Is Nothing Then Set m_lyGrv = ActiveLayer

    For i = 0 To n - 1
        If m_pn(i) > 0 Then
            pieces = pieces + ac2fPanelDrawSplit(i, bx(i), by(i), bw(i), bh(i), _
                                fold, joint, dir_, rep, loose)
            splits = splits + 1
            made = made + 1
            ungrouped = ungrouped + loose
            loose = 0
        Else
            Select Case ac2fPanelOne(bx(i), by(i), bw(i), bh(i), fold, dir_, rep)
                Case 1: made = made + 1
                Case 2: made = made + 1: ungrouped = ungrouped + 1
                Case Else: skipped = skipped + 1
            End Select
        End If
    Next i

    If Not keepSrc Then
        For i = 0 To n - 1
            On Error Resume Next
            src(i).Delete
            On Error GoTo Fail
        Next i
    End If

    ActiveDocument.EndCommandGroup
    grouped = False

    On Error Resume Next
    If Not oldLayer Is Nothing Then oldLayer.Activate
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    On Error GoTo 0

    ac2fInfo ac2fPanelReport(made, skipped, ungrouped, fold, dir_, keepSrc, _
                             rep, splits, pieces, joint), CAPTION_
    Exit Sub

Fail:
    why = Err.Description
    On Error Resume Next
    If grouped Then ActiveDocument.EndCommandGroup
    If Not oldLayer Is Nothing Then oldLayer.Activate
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    ac2fWarn "The operation failed: " & why, CAPTION_
End Sub

'=====================================================================
' INPUT
'=====================================================================

' Asks for the fold size and direction in one line, for example
' "5cm out" or "50mm in" or just "5cm". A bare number is millimetres.
Private Function ac2fAskFold(ByRef fold As Double, ByRef dir_ As Long) As Boolean
    Dim answer As String
    Dim sizeTok As String, dirTok As String
    Dim p As Long
    Dim d As Double
    Dim curFold As Double
    Dim curDir As Long

    curFold = ac2fGetNum(AC2F_K_ACP_FOLD, AC2F_DEF_ACP_FOLD)
    curDir = ac2fGetLng(AC2F_K_ACP_DIR, AC2F_DEF_ACP_DIR)

    answer = InputBox( _
        "Fold size and direction." & vbCrLf & vbCrLf & _
        "   5cm out    5 cm fold, sheet grows outwards" & vbCrLf & _
        "   50mm out   the same, written in millimetres" & vbCrLf & _
        "   5cm in     grooves set inwards, sheet size kept" & vbCrLf & vbCrLf & _
        "A bare number is read as millimetres." & vbCrLf & _
        "out = the selected rectangle is the FINISHED size" & vbCrLf & _
        "in  = the selected rectangle is the SHEET size", _
        ac2fTitle(CAPTION_), _
        ac2fNumStr(curFold) & "mm " & IIf(curDir = AC2F_ACP_IN, "in", "out"))

    If StrPtr(answer) = 0 Then Exit Function
    answer = Trim$(answer)
    If Len(answer) = 0 Then Exit Function

    p = InStr(answer, " ")
    If p > 0 Then
        sizeTok = Left$(answer, p - 1)
        dirTok = LCase$(Trim$(Mid$(answer, p + 1)))
    Else
        sizeTok = answer
        dirTok = ""
    End If

    d = ac2fParseLen(sizeTok, -1)
    If d <= 0 Then
        ac2fWarn "Could not read a fold size from """ & sizeTok & """." & vbCrLf & _
                 "Write it as 5cm, 50mm or 50.", CAPTION_
        Exit Function
    End If

    Select Case dirTok
        Case "": dir_ = curDir
        Case "in", "inner", "i", "ic": dir_ = AC2F_ACP_IN
        Case "out", "outer", "o", "dis": dir_ = AC2F_ACP_OUT
        Case Else
            ac2fWarn "Do not understand the direction """ & dirTok & """." & vbCrLf & _
                     "Use out or in.", CAPTION_
            Exit Function
    End Select

    fold = d
    ac2fSetNum AC2F_K_ACP_FOLD, d
    ac2fSetLng AC2F_K_ACP_DIR, dir_
    ac2fAskFold = True
End Function

' Reads a length written as "5cm", "50mm" or "50" and returns mm.
Public Function ac2fParseLen(ByVal s As String, ByVal defMM As Double) As Double
    Dim t As String
    Dim mult As Double

    t = LCase$(Trim$(s))
    mult = 1#
    If Len(t) > 2 Then
        If Right$(t, 2) = "cm" Then
            mult = 10#
            t = Trim$(Left$(t, Len(t) - 2))
        ElseIf Right$(t, 2) = "mm" Then
            t = Trim$(Left$(t, Len(t) - 2))
        End If
    End If

    If Len(t) = 0 Then
        ac2fParseLen = defMM
        Exit Function
    End If
    ac2fParseLen = ac2fParseNum(t, defMM) * mult
End Function

'=====================================================================
' SPLIT: PLANNING AND QUESTIONS
'=====================================================================

' The panel is split along its longer side.
Private Sub ac2fPanelAxis(ByVal bw As Double, ByVal bh As Double, _
                          ByRef along As Double, ByRef across As Double)
    If bw >= bh Then
        along = bw
        across = bh
    Else
        along = bh
        across = bw
    End If
End Sub

' Works out, for every rectangle, whether it has to be split and how.
' Asks the plate question once and the "how" question once per panel
' that needs it. Draws nothing, so Cancel leaves the page untouched.
' Returns False when the user cancelled.
Private Function ac2fPanelPlanAll(ByVal n As Long, ByRef bw() As Double, _
                                  ByRef bh() As Double, ByVal fold As Double, _
                                  ByVal dir_ As Long, ByVal forceSplit As Boolean, _
                                  ByRef joint As Double) As Boolean
    Dim plates As String
    Dim pv(0 To 12) As Double
    Dim pc As Long
    Dim i As Long, first As Long
    Dim along As Double, across As Double
    Dim maxP As Double
    Dim off As Boolean
    Dim res As Long

    m_sep = False
    ReDim m_pn(0 To n - 1)
    ReDim m_pplate(0 To n * MAX_PIECES - 1)
    ReDim m_pused(0 To n * MAX_PIECES - 1)

    joint = ac2fGetNum(AC2F_K_ACP_JOINT, AC2F_DEF_ACP_JOINT)
    plates = ac2fGetStr(AC2F_K_ACP_PLATES, AC2F_DEF_ACP_PLATES)
    ac2fPanelPlateList plates, pv, pc
    maxP = 0#
    If pc > 0 Then maxP = pv(pc - 1)

    first = -1
    For i = 0 To n - 1
        ac2fPanelAxis bw(i), bh(i), along, across
        If forceSplit Then
            If first < 0 Then first = i
        ElseIf maxP > 0# Then
            If ac2fPanelFlat(along, fold, joint, 1, dir_) > maxP + 0.001 Then
                If first < 0 Then first = i
            End If
        End If
    Next i
    If first < 0 Then
        ac2fPanelPlanAll = True
        Exit Function
    End If

    ac2fPanelAxis bw(first), bh(first), along, across
    If Not ac2fAskPlates(plates, joint, along, across, fold, dir_, off) Then Exit Function
    If off Then
        ac2fPanelPlanAll = True
        Exit Function
    End If
    ac2fPanelPlateList plates, pv, pc
    If pc = 0 Then
        ac2fPanelPlanAll = True
        Exit Function
    End If
    maxP = pv(pc - 1)

    For i = 0 To n - 1
        ac2fPanelAxis bw(i), bh(i), along, across
        If forceSplit Or ac2fPanelFlat(along, fold, joint, 1, dir_) > maxP + 0.001 Then
            res = ac2fAskSplit(i, n, along, across, fold, joint, dir_, pv, pc)
            If res < 0 Then Exit Function
        End If
    Next i
    ac2fPanelPlanAll = True
End Function

' Plate lengths and joint allowance in one line, for example
' "3200 j30" or "3200 2440 j20". Both are remembered. "0" means: do not
' split, draw one sheet.
Private Function ac2fAskPlates(ByRef plates As String, ByRef joint As Double, _
                               ByVal along As Double, ByVal across As Double, _
                               ByVal fold As Double, ByVal dir_ As Long, _
                               ByRef off As Boolean) As Boolean
    Dim answer As String, defText As String, msg As String, why As String
    Dim pv(0 To 12) As Double
    Dim pc As Long, i As Long
    Dim newPlates As String, newJoint As Double

    ac2fPanelPlateList plates, pv, pc
    For i = 0 To pc - 1
        defText = defText & ac2fPanelMm(pv(i)) & " "
    Next i
    If pc = 0 Then defText = AC2F_DEF_ACP_PLATES & " "
    defText = defText & "j" & ac2fPanelMm(joint)

    msg = "Plate lengths and joint allowance." & vbCrLf & vbCrLf & _
          "   3200 j30       one plate length, joint 30" & vbCrLf & _
          "   3200 2440 j20  two plate lengths, joint 20" & vbCrLf & _
          "   3.2m j3cm      metres and cm are read too" & vbCrLf & _
          "   0              no split, draw one sheet" & vbCrLf & _
          "   sep            lay the plates apart, not in place" & vbCrLf & vbCrLf
    msg = msg & "Panel " & ac2fPanelMm(along) & " x " & ac2fPanelMm(across) & " mm " & _
          IIf(dir_ = AC2F_ACP_IN, "(sheet size)", "(finished size)") & _
          ", fold " & ac2fPanelMm(fold) & "." & vbCrLf & _
          "j = allowance a plate keeps on each side that meets" & vbCrLf & _
          "another plate (both plates of a joint have it)."

    Do
        answer = InputBox(msg, ac2fTitle(CAPTION_), defText)
        If StrPtr(answer) = 0 Then Exit Function
        If ac2fPanelParsePlates(answer, plates, joint, newPlates, newJoint, off, m_sep, why) Then Exit Do
        ac2fWarn why, CAPTION_
        defText = Trim$(answer)
    Loop

    If Not off Then
        plates = newPlates
        joint = newJoint
        ac2fSetStr AC2F_K_ACP_PLATES, plates
        ac2fSetNum AC2F_K_ACP_JOINT, joint
    End If
    ac2fAskPlates = True
End Function

' Reads the plate/joint line. On success outPlates is the cleaned list
' (ascending, no duplicates) and outJoint the joint allowance.
Private Function ac2fPanelParsePlates(ByVal ans As String, ByVal curPlates As String, _
                                      ByVal curJoint As Double, ByRef outPlates As String, _
                                      ByRef outJoint As Double, ByRef off As Boolean, _
                                      ByRef sep As Boolean, ByRef why As String) As Boolean
    Dim tok() As String
    Dim t As String, pre As String, valText As String, lst As String
    Dim k As Long, c As Long, i As Long, pc As Long
    Dim v As Double
    Dim pv(0 To 12) As Double

    off = False
    sep = False
    outPlates = curPlates
    outJoint = curJoint
    t = LCase$(Trim$(ans))
    t = Replace$(t, ", ", " ")       ' "3200, 3050"; a lone comma is a decimal one
    t = Replace$(t, ";", " ")
    tok = Split(t, " ")

    k = 0
    Do While k <= UBound(tok)
        t = tok(k)
        If Len(t) > 0 Then
            If t = "0" Or t = "no" Or t = "off" Or t = "none" Then
                off = True
            ElseIf t = "sep" Or t = "apart" Or t = "ayri" Then
                sep = True
            Else
                pre = ""
                For c = 1 To Len(t)
                    If InStr("0123456789.,", Mid$(t, c, 1)) > 0 Then Exit For
                    pre = pre & Mid$(t, c, 1)
                Next c
                If pre = "j" Or pre = "joint" Or pre = "d" Or pre = "derz" Or pre = "pay" Then
                    valText = Mid$(t, Len(pre) + 1)
                    If Len(valText) = 0 And k < UBound(tok) Then
                        k = k + 1
                        valText = tok(k)
                    End If
                    v = ac2fPanelLen(valText, 0#)
                    If v < 0# Or v > 500# Then
                        why = "Could not read a joint allowance from """ & t & """." & _
                              vbCrLf & "Write it like j30, j3cm or joint 30 (up to 500 mm)."
                        Exit Function
                    End If
                    outJoint = v
                Else
                    v = ac2fPanelLen(t, 50#)
                    If v <= 0# Then
                        why = "Could not read a plate length from """ & t & """." & _
                              vbCrLf & "Write it as 3200, 320cm or 3.2m."
                        Exit Function
                    End If
                    lst = lst & ac2fPanelMm(v) & " "
                End If
            End If
        End If
        k = k + 1
    Loop

    If off Then
        ac2fPanelParsePlates = True
        Exit Function
    End If
    If Len(Trim$(lst)) = 0 Then lst = curPlates
    ac2fPanelPlateList lst, pv, pc
    If pc = 0 Then
        why = "Give at least one plate length, for example 3200."
        Exit Function
    End If
    outPlates = ""
    For i = 0 To pc - 1
        outPlates = outPlates & ac2fPanelMm(pv(i)) & " "
    Next i
    outPlates = Trim$(outPlates)
    ac2fPanelParsePlates = True
End Function

' Turns a list such as "3200 2440" into pv(0..pc-1): ascending, no
' duplicates. Anything that is not a length is skipped.
Private Sub ac2fPanelPlateList(ByVal s As String, ByRef pv() As Double, ByRef pc As Long)
    Dim tok() As String
    Dim t As String
    Dim i As Long, j As Long
    Dim v As Double
    Dim dup As Boolean

    pc = 0
    t = LCase$(s)
    t = Replace$(t, ", ", " ")
    t = Replace$(t, ";", " ")
    tok = Split(t, " ")
    For i = 0 To UBound(tok)
        v = ac2fPanelLen(tok(i), 50#)
        If v > 0# And pc < MAX_PIECES Then
            dup = False
            For j = 0 To pc - 1
                If Abs(pv(j) - v) < 0.5 Then dup = True
            Next j
            If Not dup Then
                pv(pc) = v
                pc = pc + 1
            End If
        End If
    Next i

    For i = 1 To pc - 1
        v = pv(i)
        j = i - 1
        Do While j >= 0
            If pv(j) <= v Then Exit Do
            pv(j + 1) = pv(j)
            j = j - 1
        Loop
        pv(j + 1) = v
    Next i
End Sub

' A length written as 3200, 320cm, 3.2m or 3200mm, in mm. A bare number
' below metreBelow is read as metres (3.2 is 3.20 m). -1 when it cannot
' be read.
Private Function ac2fPanelLen(ByVal tok As String, ByVal metreBelow As Double) As Double
    Dim t As String
    Dim mult As Double, v As Double
    Dim unitGiven As Boolean

    ac2fPanelLen = -1#
    t = LCase$(Trim$(tok))
    If Len(t) = 0 Then Exit Function
    mult = 1#
    If Len(t) > 2 And Right$(t, 2) = "mm" Then
        t = Trim$(Left$(t, Len(t) - 2))
        unitGiven = True
    ElseIf Len(t) > 2 And Right$(t, 2) = "cm" Then
        t = Trim$(Left$(t, Len(t) - 2))
        mult = 10#
        unitGiven = True
    ElseIf Len(t) > 1 And Right$(t, 1) = "m" Then
        t = Trim$(Left$(t, Len(t) - 1))
        mult = 1000#
        unitGiven = True
    End If
    If Len(t) = 0 Then Exit Function
    v = ac2fParseNum(t, -1#)
    If v < 0# Then Exit Function
    If Not unitGiven And v > 0# And v < metreBelow Then mult = 1000#
    ac2fPanelLen = v * mult
End Function

' Millimetres as plain text with at most one decimal and no thousands
' separator, so a figure shown in a prompt can be typed back in.
Private Function ac2fPanelMm(ByVal v As Double) As String
    ac2fPanelMm = ac2fNumStr(Int(v * 10# + 0.5) / 10#)
End Function

Private Function ac2fPanelIsInt(ByVal s As String) As Boolean
    Dim c As Long
    If Len(s) = 0 Or Len(s) > 6 Then Exit Function
    For c = 1 To Len(s)
        If InStr("0123456789", Mid$(s, c, 1)) = 0 Then Exit Function
    Next c
    ac2fPanelIsInt = True
End Function

'---------------------------------------------------------------------
' The sums. Everything is measured along the split axis.
'
'   out:  finished length is what you selected
'         flat  = finished + 2 folds + 2 joints per joint
'   in:   the selection is the flat sheet
'         finished = flat - 2 folds - 2 joints per joint
'---------------------------------------------------------------------

' Flat length of the whole panel when it is cut into n plates.
Private Function ac2fPanelFlat(ByVal sel As Double, ByVal fold As Double, _
                               ByVal joint As Double, ByVal n As Long, _
                               ByVal dir_ As Long) As Double
    If dir_ = AC2F_ACP_IN Then
        ac2fPanelFlat = sel
    Else
        ac2fPanelFlat = sel + 2# * fold + 2# * (n - 1) * joint
    End If
End Function

' Finished (face) length when the panel is cut into n plates.
Private Function ac2fPanelNet(ByVal sel As Double, ByVal fold As Double, _
                              ByVal joint As Double, ByVal n As Long, _
                              ByVal dir_ As Long) As Double
    If dir_ = AC2F_ACP_IN Then
        ac2fPanelNet = sel - 2# * fold - 2# * (n - 1) * joint
    Else
        ac2fPanelNet = sel
    End If
End Function

' Allowance at the start (side 0) or the end (side 1) of piece k of n.
' The two ends of the sign carry the fold, every joint carries the joint.
Private Function ac2fPanelAllow(ByVal k As Long, ByVal n As Long, ByVal side As Long, _
                                ByVal fold As Double, ByVal joint As Double) As Double
    If side = 0 Then
        If k = 0 Then
            ac2fPanelAllow = fold
        Else
            ac2fPanelAllow = joint
        End If
    Else
        If k = n - 1 Then
            ac2fPanelAllow = fold
        Else
            ac2fPanelAllow = joint
        End If
    End If
End Function

' "Full plates": every plate is used to its full length and the last
' one takes the rest, so the leftover is one big reusable piece.
' Returns the plate count, 0 when it cannot be done.
Private Function ac2fPanelFill(ByVal sel As Double, ByVal fold As Double, _
                               ByVal joint As Double, ByVal dir_ As Long, _
                               ByVal p As Double, ByRef used() As Double) As Long
    Dim n As Long, i As Long
    Dim lastFace As Double, shift As Double

    n = 1
    Do While n * p < ac2fPanelFlat(sel, fold, joint, n, dir_) - 0.001
        n = n + 1
        If n > MAX_PIECES Then Exit Function
    Loop
    If ac2fPanelNet(sel, fold, joint, n, dir_) <= 0# Then Exit Function

    For i = 0 To n - 2
        used(i) = p
    Next i
    used(n - 1) = ac2fPanelFlat(sel, fold, joint, n, dir_) - (n - 1) * p
    ' A sliver of a last plate is no use: take a little from the one before.
    If n > 1 Then
        lastFace = used(n - 1) - fold - joint
        If lastFace < MIN_FACE Then
            shift = MIN_FACE - lastFace
            used(n - 1) = used(n - 1) + shift
            used(n - 2) = used(n - 2) - shift
        End If
    End If
    ac2fPanelFill = n
End Function

' "Even joints": every plate carries the same face length, so the joints
' sit at equal spacing. Starts at minN plates and goes up until each cut
' fits its plate. Returns the plate count, 0 when it cannot be done.
Private Function ac2fPanelEqual(ByVal sel As Double, ByVal fold As Double, _
                                ByVal joint As Double, ByVal dir_ As Long, _
                                ByVal p As Double, ByVal minN As Long, _
                                ByRef used() As Double) As Long
    Dim n As Long, i As Long
    Dim net As Double, face As Double, sumUsed As Double
    Dim fits As Boolean

    n = minN
    If n < 1 Then n = 1
    Do While n <= MAX_PIECES
        net = ac2fPanelNet(sel, fold, joint, n, dir_)
        If net <= 0# Then Exit Function
        face = Int(net / n * 10# + 0.5) / 10#
        fits = True
        sumUsed = 0#
        For i = 0 To n - 1
            If i = n - 1 Then
                used(i) = ac2fPanelFlat(sel, fold, joint, n, dir_) - sumUsed
            Else
                used(i) = face + ac2fPanelAllow(i, n, 0, fold, joint) + _
                          ac2fPanelAllow(i, n, 1, fold, joint)
                sumUsed = sumUsed + used(i)
            End If
            If used(i) > p + 0.001 Then fits = False
        Next i
        If fits Then
            ac2fPanelEqual = n
            Exit Function
        End If
        n = n + 1
    Loop
End Function

' "" when the plan is usable, otherwise what is wrong with it.
Private Function ac2fPanelCheck(ByVal n As Long, ByRef plate() As Double, _
                                ByRef used() As Double, ByVal sel As Double, _
                                ByVal fold As Double, ByVal joint As Double, _
                                ByVal dir_ As Long) As String
    Dim k As Long
    Dim a0 As Double, a1 As Double, face As Double, sumFace As Double

    For k = 0 To n - 1
        a0 = ac2fPanelAllow(k, n, 0, fold, joint)
        a1 = ac2fPanelAllow(k, n, 1, fold, joint)
        If used(k) > plate(k) + 0.001 Then
            ac2fPanelCheck = "Plate " & (k + 1) & " is " & ac2fPanelMm(plate(k)) & _
                " long but " & ac2fPanelMm(used(k)) & " is asked of it."
            Exit Function
        End If
        face = used(k) - a0 - a1
        If face < 1# Then
            ac2fPanelCheck = "Plate " & (k + 1) & " is cut " & ac2fPanelMm(used(k)) & _
                ", no more than its own folds and joints (" & ac2fPanelMm(a0 + a1) & _
                "); no face is left."
            Exit Function
        End If
        sumFace = sumFace + face
    Next k
    If Abs(sumFace - ac2fPanelNet(sel, fold, joint, n, dir_)) > 0.5 Then
        ac2fPanelCheck = "The faces add up to " & ac2fPanelMm(sumFace) & ", not " & _
            ac2fPanelMm(ac2fPanelNet(sel, fold, joint, n, dir_)) & "."
    End If
End Function

' "3200 + 960"
Private Function ac2fPanelUsedText(ByVal n As Long, ByRef used() As Double) As String
    Dim k As Long
    Dim s As String
    For k = 0 To n - 1
        If k > 0 Then s = s & " + "
        s = s & ac2fPanelMm(used(k))
    Next k
    ac2fPanelUsedText = s
End Function

'---------------------------------------------------------------------
' Candidate plans for one panel
'---------------------------------------------------------------------

Private Sub ac2fPanelAddCand(ByVal p As Double, ByVal n As Long, ByRef used() As Double, _
                             ByVal sel As Double, ByVal fold As Double, _
                             ByVal joint As Double, ByVal dir_ As Long)
    Dim j As Long, k As Long
    Dim pl(0 To 11) As Double
    Dim key As String

    If n < 1 Or m_cc >= MAX_CANDS Then Exit Sub
    For k = 0 To n - 1
        pl(k) = p
    Next k
    If Len(ac2fPanelCheck(n, pl, used, sel, fold, joint, dir_)) > 0 Then Exit Sub
    ' The same cut lengths on a longer plate are the same plan, only
    ' wasteful: keep the shortest plate that fits.
    key = ac2fPanelUsedText(n, used)
    For j = 0 To m_cc - 1
        If m_ct(j) = key Then
            If p < m_cp(j) Then m_cp(j) = p
            Exit Sub
        End If
    Next j
    m_cn(m_cc) = n
    m_cp(m_cc) = p
    m_ct(m_cc) = key
    For k = 0 To n - 1
        m_cu(m_cc * MAX_PIECES + k) = used(k)
    Next k
    m_cc = m_cc + 1
End Sub

' Works out the plans worth offering, for every plate length in the
' list, and picks the first MAX_OPTS of them by plate count (fewest
' plates first; for the same count, full plates before even joints).
Private Sub ac2fPanelCandidates(ByVal sel As Double, ByVal fold As Double, _
                                ByVal joint As Double, ByVal dir_ As Long, _
                                ByRef pv() As Double, ByVal pc As Long)
    Dim k As Long, j As Long, n As Long, best As Long
    Dim nF As Long, nE As Long
    Dim p As Double
    Dim used(0 To 11) As Double

    m_cc = 0
    m_oc = 0
    ReDim m_cn(0 To MAX_CANDS - 1)
    ReDim m_cp(0 To MAX_CANDS - 1)
    ReDim m_cu(0 To MAX_CANDS * MAX_PIECES - 1)
    ReDim m_ct(0 To MAX_CANDS - 1)
    ReDim m_oi(0 To MAX_OPTS - 1)
    For k = pc - 1 To 0 Step -1
        p = pv(k)
        nF = ac2fPanelFill(sel, fold, joint, dir_, p, used)
        If nF > 0 Then ac2fPanelAddCand p, nF, used, sel, fold, joint, dir_
        If nF < 1 Then nF = 1
        nE = ac2fPanelEqual(sel, fold, joint, dir_, p, nF, used)
        If nE > 0 Then
            ac2fPanelAddCand p, nE, used, sel, fold, joint, dir_
            nE = ac2fPanelEqual(sel, fold, joint, dir_, p, nE + 1, used)
            If nE > 0 Then ac2fPanelAddCand p, nE, used, sel, fold, joint, dir_
        End If
    Next k

    ' Fewest plates first; for the same count the least wasted plate
    ' first, and on a tie full plates before even joints.
    For n = 1 To MAX_PIECES
        Do While m_oc < MAX_OPTS
            best = -1
            For j = 0 To m_cc - 1
                If m_cn(j) = n And Not ac2fPanelPicked(j) Then
                    If best < 0 Then
                        best = j
                    ElseIf ac2fPanelWaste(j) < ac2fPanelWaste(best) - 0.5 Then
                        best = j
                    End If
                End If
            Next j
            If best < 0 Then Exit Do
            m_oi(m_oc) = best
            m_oc = m_oc + 1
        Loop
    Next n
End Sub

' Plate length left over once the plan's cuts are taken from it.
Private Function ac2fPanelWaste(ByVal ci As Long) As Double
    Dim k As Long
    Dim w As Double
    For k = 0 To m_cn(ci) - 1
        w = w + m_cp(ci) - m_cu(ci * MAX_PIECES + k)
    Next k
    ac2fPanelWaste = w
End Function

' True when plan ci is already in the option list.
Private Function ac2fPanelPicked(ByVal ci As Long) As Boolean
    Dim o As Long
    For o = 0 To m_oc - 1
        If m_oi(o) = ci Then
            ac2fPanelPicked = True
            Exit Function
        End If
    Next o
End Function

' One line of the option list, with the cut length taken from each plate
' and the face each plate then gives on the finished size:
' "1. 3200 + 960  faces 3120 + 880  (left 2240)".
Private Function ac2fPanelOptLine(ByVal o As Long, ByVal pc As Long, _
                                  ByVal fold As Double, ByVal joint As Double) As String
    Dim ci As Long, k As Long, n As Long
    Dim used(0 To 11) As Double, face(0 To 11) As Double
    Dim lft As String, s As String

    ci = m_oi(o)
    n = m_cn(ci)
    For k = 0 To n - 1
        used(k) = m_cu(ci * MAX_PIECES + k)
        face(k) = used(k) - ac2fPanelAllow(k, n, 0, fold, joint) - _
                  ac2fPanelAllow(k, n, 1, fold, joint)
        If m_cp(ci) - used(k) > 0.5 Then
            If Len(lft) > 0 Then lft = lft & "+"
            lft = lft & ac2fPanelMm(m_cp(ci) - used(k))
        End If
    Next k
    s = (o + 1) & ". " & ac2fPanelUsedText(n, used) & "  faces " & ac2fPanelUsedText(n, face) & "  ("
    If pc > 1 Then s = s & "of " & ac2fPanelMm(m_cp(ci)) & ", "
    If Len(lft) > 0 Then s = s & "left " & lft Else s = s & "no offcut"
    ac2fPanelOptLine = s & ")"
End Function

' The question for one panel: the options, then how to type your own.
Private Function ac2fPanelSplitPrompt(ByVal i As Long, ByVal np As Long, _
                                      ByVal along As Double, ByVal across As Double, _
                                      ByVal fold As Double, ByVal joint As Double, _
                                      ByVal dir_ As Long, ByVal pc As Long) As String
    Dim msg As String, ln As String
    Dim o As Long, n1 As Long, shown As Long

    msg = "Split across plates."
    If np > 1 Then msg = msg & "  Panel " & (i + 1) & " of " & np
    msg = msg & vbCrLf & "Panel " & ac2fPanelMm(along) & " x " & ac2fPanelMm(across) & _
          ", fold " & ac2fPanelMm(fold) & ", joint " & ac2fPanelMm(joint) & vbCrLf

    If m_oc > 0 Then
        n1 = m_cn(m_oi(0))
        If dir_ = AC2F_ACP_IN Then
            msg = msg & "Sheet " & ac2fPanelMm(along) & " - folds " & ac2fPanelMm(2# * fold) & _
                  " - joints " & ac2fPanelMm(2# * (n1 - 1) * joint) & " = " & _
                  ac2fPanelMm(ac2fPanelNet(along, fold, joint, n1, dir_)) & " finished"
        Else
            msg = msg & "Finished " & ac2fPanelMm(along) & " + folds " & ac2fPanelMm(2# * fold) & _
                  " + joints " & ac2fPanelMm(2# * (n1 - 1) * joint) & " = " & _
                  ac2fPanelMm(ac2fPanelFlat(along, fold, joint, n1, dir_)) & " flat"
        End If
        msg = msg & "  (" & n1 & IIf(n1 = 1, " plate)", " plates)") & vbCrLf & vbCrLf
        ' An InputBox prompt tops out near 1024 characters: show as many
        ' options as fit and forget the rest, so a number always matches
        ' a line on screen.
        shown = 0
        For o = 0 To m_oc - 1
            ln = ac2fPanelOptLine(o, pc, fold, joint) & vbCrLf
            If Len(msg) + Len(ln) > PROMPT_ROOM Then Exit For
            msg = msg & ln
            shown = shown + 1
        Next o
        m_oc = shown
    Else
        msg = msg & vbCrLf & "No ready-made plan fits these plates." & vbCrLf
    End If

    msg = msg & vbCrLf & "Pick a number, or type the length cut from each" & vbCrLf & _
          "plate (folds and joints included):" & vbCrLf & _
          "   3200 960           the shortest plate that fits" & vbCrLf & _
          "   3050:3000 3200:*   plate:length, * = what is left" & vbCrLf & _
          "   0                  no split, one sheet"
    ac2fPanelSplitPrompt = msg
End Function

' Asks how to split panel i. Returns the plate count, 0 for "one sheet",
' -1 when cancelled.
Private Function ac2fAskSplit(ByVal i As Long, ByVal np As Long, ByVal along As Double, _
                              ByVal across As Double, ByVal fold As Double, _
                              ByVal joint As Double, ByVal dir_ As Long, _
                              ByRef pv() As Double, ByVal pc As Long) As Long
    Dim answer As String, msg As String, why As String, defText As String
    Dim tp(0 To 11) As Double, tu(0 To 11) As Double
    Dim cnt As Long, k As Long, o As Long

    ac2fAskSplit = -1
    ac2fPanelCandidates along, fold, joint, dir_, pv, pc
    msg = ac2fPanelSplitPrompt(i, np, along, across, fold, joint, dir_, pc)
    If m_oc > 0 Then defText = "1" Else defText = "0"

    Do
        answer = InputBox(msg, ac2fTitle(CAPTION_), defText)
        If StrPtr(answer) = 0 Then Exit Function
        answer = LCase$(Trim$(answer))
        If Len(answer) = 0 Then answer = defText

        If answer = "0" Or answer = "no" Or answer = "off" Or answer = "none" Then
            m_pn(i) = 0
            ac2fAskSplit = 0
            Exit Function
        End If

        If ac2fPanelIsInt(answer) And Val(answer) <= 20# Then
            o = CLng(Val(answer))
            If o >= 1 And o <= m_oc Then
                cnt = m_cn(m_oi(o - 1))
                m_pn(i) = cnt
                For k = 0 To cnt - 1
                    m_pplate(i * MAX_PIECES + k) = m_cp(m_oi(o - 1))
                    m_pused(i * MAX_PIECES + k) = m_cu(m_oi(o - 1) * MAX_PIECES + k)
                Next k
                ac2fAskSplit = cnt
                Exit Function
            End If
            why = "There is no option " & o & "."
        ElseIf ac2fPanelParseCustom(answer, along, fold, joint, dir_, pv, pc, tp, tu, cnt, why) Then
            m_pn(i) = cnt
            For k = 0 To cnt - 1
                m_pplate(i * MAX_PIECES + k) = tp(k)
                m_pused(i * MAX_PIECES + k) = tu(k)
            Next k
            ac2fAskSplit = cnt
            Exit Function
        End If
        ac2fWarn why, CAPTION_
        defText = answer
    Loop
End Function

' Reads "3200 960", "3050:3000 3200:*" and the like: one item per plate,
' each the length cut from it, optionally preceded by the plate length.
' A single "*" takes whatever the others leave. On success tp/tu hold
' plate and cut length per piece and cnt the count.
Private Function ac2fPanelParseCustom(ByVal ans As String, ByVal sel As Double, _
                                      ByVal fold As Double, ByVal joint As Double, _
                                      ByVal dir_ As Long, ByRef pv() As Double, _
                                      ByVal pc As Long, ByRef tp() As Double, _
                                      ByRef tu() As Double, ByRef cnt As Long, _
                                      ByRef why As String) As Boolean
    Dim tok() As String
    Dim t As String, lp As String, rp As String
    Dim k As Long, c As Long, starAt As Long, j As Long
    Dim total As Double, sumKnown As Double, need As Double
    Dim bad As String

    t = LCase$(Trim$(ans))
    t = Replace$(t, ", ", " ")
    t = Replace$(t, ";", " ")
    t = Replace$(t, "+", " ")
    tok = Split(t, " ")

    cnt = 0
    starAt = -1
    For k = 0 To UBound(tok)
        If Len(tok(k)) > 0 Then
            If cnt >= MAX_PIECES Then
                why = "At most " & MAX_PIECES & " plates per panel."
                Exit Function
            End If
            c = InStr(tok(k), ":")
            If c > 0 Then
                lp = Left$(tok(k), c - 1)
                rp = Mid$(tok(k), c + 1)
            Else
                lp = ""
                rp = tok(k)
            End If
            tp(cnt) = 0#
            If Len(lp) > 0 Then
                tp(cnt) = ac2fPanelLen(lp, 50#)
                If tp(cnt) <= 0# Then
                    why = "Could not read the plate length in """ & tok(k) & """."
                    Exit Function
                End If
            End If
            If rp = "*" Then
                If starAt >= 0 Then
                    why = "Only one plate can take ""the rest"" (*)."
                    Exit Function
                End If
                starAt = cnt
                tu(cnt) = 0#
            Else
                tu(cnt) = ac2fPanelLen(rp, 50#)
                If tu(cnt) <= 0# Then
                    why = "Could not read """ & tok(k) & """ as a length." & vbCrLf & _
                          "Type a number from the list, or lengths like 3200 960."
                    Exit Function
                End If
                sumKnown = sumKnown + tu(cnt)
            End If
            cnt = cnt + 1
        End If
    Next k
    If cnt = 0 Then
        why = "Type an option number or the length cut from each plate."
        Exit Function
    End If

    total = ac2fPanelFlat(sel, fold, joint, cnt, dir_)
    If starAt >= 0 Then
        tu(starAt) = total - sumKnown
        If tu(starAt) <= 0# Then
            why = "The other plates already take " & ac2fPanelMm(sumKnown) & ", and " & cnt & _
                  " plates need " & ac2fPanelMm(total) & " in all."
            Exit Function
        End If
    ElseIf Abs(sumKnown - total) > 0.5 Then
        why = "These lengths add up to " & ac2fPanelMm(sumKnown) & " but " & cnt & _
              " plates need " & ac2fPanelMm(total) & vbCrLf & "(" & _
              ac2fPanelMm(ac2fPanelNet(sel, fold, joint, cnt, dir_)) & " finished + folds and joints; " & _
              IIf(sumKnown < total, "short by ", "over by ") & _
              ac2fPanelMm(Abs(total - sumKnown)) & ")." & vbCrLf & _
              "End the list with * and the last plate takes the rest."
        Exit Function
    End If

    ' Plates left out are picked: the shortest one in the list that fits.
    For k = 0 To cnt - 1
        If tp(k) <= 0# Then
            For j = 0 To pc - 1
                If pv(j) >= tu(k) - 0.001 Then
                    tp(k) = pv(j)
                    Exit For
                End If
            Next j
            If tp(k) <= 0# Then
                why = "No plate in your list is long enough for " & ac2fPanelMm(tu(k)) & _
                      " (the longest is " & ac2fPanelMm(pv(pc - 1)) & ")." & vbCrLf & _
                      "Write it as plate:length, for example " & ac2fPanelMm(tu(k)) & ":" & _
                      ac2fPanelMm(tu(k)) & "."
                Exit Function
            End If
        End If
    Next k

    bad = ac2fPanelCheck(cnt, tp, tu, sel, fold, joint, dir_)
    If Len(bad) > 0 Then
        why = bad
        Exit Function
    End If
    ac2fPanelParseCustom = True
End Function


'=====================================================================
' LAYERS
'=====================================================================

' The cut outline and the V-grooves live on layers of their own, so the
' CAM program (or a cut-out against the artwork) can switch either off.
' Found by name on the active page, created when missing. Nothing when
' the layer cannot be had; the caller then falls back to the active one.
Private Function ac2fPanelLayer(ByVal nm As String) As Layer
    Dim i As Long
    Dim ly As Layer

    On Error Resume Next
    For i = 1 To ActivePage.Layers.Count
        If ActivePage.Layers(i).Name = nm Then
            Set ly = ActivePage.Layers(i)
            Exit For
        End If
    Next i
    If ly Is Nothing Then Set ly = ActivePage.CreateLayer(nm)
    If Not ly Is Nothing Then
        ly.Visible = True
        ly.Editable = True
    End If
    Set ac2fPanelLayer = ly
End Function

'=====================================================================
' ONE PANEL
'=====================================================================

' Returns 1 when drawn and grouped, 2 when drawn but grouping failed,
' 0 when the rectangle was skipped.
Private Function ac2fPanelOne(ByVal bx As Double, ByVal by As Double, _
                              ByVal bw As Double, ByVal bh As Double, _
                              ByVal fold As Double, ByVal dir_ As Long, _
                              ByRef rep As String) As Long
    Dim x0 As Double, y0 As Double, w As Double, h As Double
    Dim why As String
    Dim r As Long
    Dim sheet As Shape, grv As Shape

    If bw <= 0 Or bh <= 0 Then Exit Function

    If dir_ = AC2F_ACP_IN Then
        x0 = bx: y0 = by: w = bw: h = bh
    Else
        x0 = bx - fold: y0 = by - fold
        w = bw + 2# * fold: h = bh + 2# * fold
    End If

    r = ac2fPanelDraw(x0, y0, w, h, fold, fold, fold, fold, "", sheet, grv, why)
    If r = 0 Then
        If Len(why) > 0 Then
            rep = rep & "   skipped " & ac2fFmt(bw, 0) & "x" & ac2fFmt(bh, 0) & _
                  " - " & why & vbCrLf
        End If
        Exit Function
    End If

    rep = rep & "   sheet " & ac2fFmt(w, 0) & "x" & ac2fFmt(h, 0) & " mm" & _
          "   folded " & ac2fFmt(w - 2# * fold, 0) & "x" & ac2fFmt(h - 2# * fold, 0) & _
          " mm" & vbCrLf
    ac2fPanelOne = r
End Function

' Draws one sheet outline on the cut layer and its four V-grooves, as a
' group, on the groove layer. The fold on each side is its own number,
' so a plate of a split panel can carry the sign's fold on one end and
' the joint allowance on the other:
'   fl left, fb bottom, fr right, ft top (all in mm, measured inwards).
' tag names the piece ("1 of 2"); "" for an ordinary single panel.
' sheet and grv return the shapes made (grv is Nothing when the four
' lines could not be grouped). Returns 1 drawn and grouped, 2 drawn but
' not grouped, 0 skipped (why says why when there is something to tell).
Private Function ac2fPanelDraw(ByVal x0 As Double, ByVal y0 As Double, _
                               ByVal w As Double, ByVal h As Double, _
                               ByVal fl As Double, ByVal fb As Double, _
                               ByVal fr As Double, ByVal ft As Double, _
                               ByVal tag As String, ByRef sheet As Shape, _
                               ByRef grv As Shape, ByRef why As String) As Long
    Dim g(0 To 3) As Shape
    Dim i As Long
    Dim dims As String, pre As String

    Set sheet = Nothing
    Set grv = Nothing
    If fl + fr >= w Or fb + ft >= h Then
        why = "fold too large for an inward layout"
        Exit Function
    End If

    dims = ac2fPanelMm(w) & "x" & ac2fPanelMm(h)
    pre = ""
    If Len(tag) > 0 Then pre = tag & " "

    ' --- sheet outline, cut layer ---------------------------------------
    Set sheet = ac2fMakeRect(m_lyCut, x0, y0, w, h)
    If sheet Is Nothing Then Exit Function
    sheet.Name = "ac2f ACP sheet " & pre & dims

    ' --- four groove lines, groove layer, clockwise, in cutting order ----
    ' 1 left: starts at the bottom so the tool enters low
    Set g(0) = ac2fMakeLine(m_lyGrv, x0 + fl, y0, x0 + fl, y0 + h, "1 left")
    ' 2 top: starts at the left
    Set g(1) = ac2fMakeLine(m_lyGrv, x0, y0 + h - ft, x0 + w, y0 + h - ft, "2 top")
    ' 3 right: starts at the top
    Set g(2) = ac2fMakeLine(m_lyGrv, x0 + w - fr, y0 + h, x0 + w - fr, y0, "3 right")
    ' 4 bottom: starts at the right
    Set g(3) = ac2fMakeLine(m_lyGrv, x0 + w, y0 + fb, x0, y0 + fb, "4 bottom")

    For i = 0 To 3
        If g(i) Is Nothing Then
            ac2fPanelDraw = 2
            Exit Function
        End If
    Next i

    ' --- the four grooves are one group --------------------------------
    Set grv = ac2fGroupShapes(g, 4, "ac2f ACP grooves " & pre & dims)
    If grv Is Nothing Then
        ac2fPanelDraw = 2
        Exit Function
    End If
    ac2fPanelDraw = 1
End Function

' Draws the plates of panel i. Normally IN PLACE: the faces tile the
' finished size exactly, every groove lies on the original outline and
' the joint allowances stick out past the faces, overlapping the next
' plate's, so the whole sign can be viewed (and cut out against artwork)
' at its real size. With "sep" the plates are laid apart instead, one
' under the other below the source for a wide panel, side by side to its
' right for a tall one, so they do not overlap.
' The sheet outlines are one group on the cut layer and the grooves one
' group on the groove layer (each plate's four lines a group inside it).
' Returns how many plates were drawn; loose counts groups that could
' not be made.
Private Function ac2fPanelDrawSplit(ByVal i As Long, ByVal bx As Double, ByVal by As Double, _
                                    ByVal bw As Double, ByVal bh As Double, _
                                    ByVal fold As Double, ByVal joint As Double, _
                                    ByVal dir_ As Long, ByRef rep As String, _
                                    ByRef loose As Long) As Long
    Dim isX As Boolean
    Dim n As Long, k As Long, r As Long, made As Long
    Dim sel As Double, across As Double
    Dim netAlong As Double, netAcross As Double, acrossFlat As Double
    Dim nx0 As Double, ny0 As Double, nTop As Double
    Dim w As Double, h As Double, x0 As Double, y0 As Double
    Dim fl As Double, fb As Double, fr As Double, ft As Double
    Dim xCur As Double, yCur As Double
    Dim a0 As Double, a1 As Double, face As Double, offs As Double
    Dim sheets(0 To 11) As Shape, grvs(0 To 11) As Shape
    Dim ns As Long, ng As Long
    Dim sh As Shape, gr As Shape
    Dim grp As Shape
    Dim why As String, line_ As String, dims As String

    n = m_pn(i)
    isX = (bw >= bh)
    If isX Then
        sel = bw
        across = bh
    Else
        sel = bh
        across = bw
    End If

    ' The finished size and where it sits. "out": the selection. "in":
    ' the selection is the sheet, so the finished size is inset by a fold
    ' and (along the split) shortened by the joints.
    netAlong = ac2fPanelNet(sel, fold, joint, n, dir_)
    netAcross = across
    nx0 = bx
    ny0 = by
    nTop = by + bh
    If dir_ = AC2F_ACP_IN Then
        netAcross = across - 2# * fold
        nx0 = bx + fold
        ny0 = by + fold
        nTop = by + bh - fold
    End If
    acrossFlat = netAcross + 2# * fold

    If isX Then
        dims = ac2fPanelMm(netAlong) & "x" & ac2fPanelMm(netAcross)
    Else
        dims = ac2fPanelMm(netAcross) & "x" & ac2fPanelMm(netAlong)
    End If

    ' Start of the first plate when they are laid apart.
    If m_sep Then
        If isX Then
            If dir_ = AC2F_ACP_IN Then
                xCur = bx
                yCur = by - PIECE_GAP
            Else
                xCur = bx - fold
                yCur = by - fold - PIECE_GAP
            End If
        Else
            If dir_ = AC2F_ACP_IN Then
                xCur = bx + bw + PIECE_GAP
                yCur = by + bh
            Else
                xCur = bx + bw + fold + PIECE_GAP
                yCur = by + bh + fold
            End If
        End If
    End If

    rep = rep & "   panel " & ac2fPanelMm(netAlong) & " x " & ac2fPanelMm(netAcross) & _
          " finished -> " & n & " plates" & _
          IIf(m_sep, ", laid apart", ", in place") & vbCrLf

    For k = 0 To n - 1
        a0 = ac2fPanelAllow(k, n, 0, fold, joint)
        a1 = ac2fPanelAllow(k, n, 1, fold, joint)
        face = m_pused(i * MAX_PIECES + k) - a0 - a1

        If isX Then
            w = m_pused(i * MAX_PIECES + k)
            h = acrossFlat
            If m_sep Then
                x0 = xCur
                y0 = yCur - h
            Else
                x0 = nx0 + offs - a0
                y0 = ny0 - fold
            End If
            fl = a0
            fr = a1
            fb = fold
            ft = fold
        Else
            w = acrossFlat
            h = m_pused(i * MAX_PIECES + k)
            If m_sep Then
                x0 = xCur
                y0 = yCur - h
            Else
                x0 = nx0 - fold
                y0 = nTop - offs + a0 - h
            End If
            ft = a0
            fb = a1
            fl = fold
            fr = fold
        End If

        r = ac2fPanelDraw(x0, y0, w, h, fl, fb, fr, ft, (k + 1) & " of " & n, sh, gr, why)
        If r > 0 Then made = made + 1
        If r = 2 Then loose = loose + 1
        If Not sh Is Nothing Then
            Set sheets(ns) = sh
            ns = ns + 1
        End If
        If Not gr Is Nothing Then
            Set grvs(ng) = gr
            ng = ng + 1
        End If

        line_ = "    " & (k + 1) & "/" & n & "  " & ac2fPanelMm(w) & " x " & ac2fPanelMm(h) & _
                "  face " & ac2fPanelMm(offs) & "-" & ac2fPanelMm(offs + face)
        If m_pplate(i * MAX_PIECES + k) - m_pused(i * MAX_PIECES + k) > 0.5 Then
            line_ = line_ & "  left " & _
                    ac2fPanelMm(m_pplate(i * MAX_PIECES + k) - m_pused(i * MAX_PIECES + k))
        End If
        If face < MIN_FACE Then line_ = line_ & "  NARROW"
        If r = 0 Then line_ = line_ & "  NOT DRAWN"
        rep = rep & line_ & vbCrLf

        offs = offs + face
        If m_sep Then
            If isX Then
                yCur = y0 - PIECE_GAP
            Else
                xCur = x0 + w + PIECE_GAP
            End If
        End If
    Next k

    ' One group of outlines on the cut layer, one of grooves on the
    ' groove layer. A single shape cannot be grouped and needs no group.
    If ns >= 2 Then
        Set grp = ac2fGroupShapes(sheets, ns, "ac2f ACP cut " & dims & " (" & n & " plates)")
        If grp Is Nothing Then loose = loose + 1
    End If
    If ng >= 2 Then
        Set grp = ac2fGroupShapes(grvs, ng, "ac2f ACP grooves " & dims & " (" & n & " plates)")
        If grp Is Nothing Then loose = loose + 1
    End If

    rep = rep & "    faces add up to " & ac2fPanelMm(offs) & " mm" & vbCrLf
    ac2fPanelDrawSplit = made
End Function

'---------------------------------------------------------------------
' Drawing helpers. Geometry must succeed; styling is best effort.
'---------------------------------------------------------------------

Private Function ac2fMakeRect(ByVal ly As Layer, ByVal x0 As Double, ByVal y0 As Double, _
                              ByVal w As Double, ByVal h As Double) As Shape
    Dim s As Shape
    On Error GoTo Fail
    Set s = ly.CreateRectangle(x0, y0 + h, x0 + w, y0)
    On Error Resume Next
    s.Fill.ApplyNoFill
    s.Outline.Width = 0.2
    s.Outline.Color.RGBAssign 0, 0, 0            ' sheet / cut contour
    On Error GoTo 0
    Set ac2fMakeRect = s
    Exit Function
Fail:
End Function

Private Function ac2fMakeLine(ByVal ly As Layer, ByVal x1 As Double, ByVal y1 As Double, _
                              ByVal x2 As Double, ByVal y2 As Double, _
                              ByVal nm As String) As Shape
    Dim s As Shape
    On Error GoTo Fail
    ' The first pair is the start node, which is the point the CAM
    ' program enters the cut at.
    Set s = ly.CreateLineSegment(x1, y1, x2, y2)
    On Error Resume Next
    s.Name = "ac2f groove " & nm
    s.Outline.Width = 0.15
    s.Outline.Color.RGBAssign 0, 160, 220        ' V-groove
    On Error GoTo 0
    Set ac2fMakeLine = s
    Exit Function
Fail:
End Function

' Groups shapes in the order given. Returns Nothing when grouping is
' not possible, in which case the shapes are still on the page.
Private Function ac2fGroupShapes(ByRef arr() As Shape, ByVal n As Long, _
                                 ByVal nm As String) As Shape
    Dim i As Long
    Dim g As Shape

    If n <= 0 Then Exit Function

    On Error GoTo Fail
    ActiveDocument.ClearSelection
    For i = 0 To n - 1
        If arr(i) Is Nothing Then Exit Function
        arr(i).AddToSelection
    Next i

    Set g = ActiveSelectionRange.Group()
    On Error Resume Next
    g.Name = nm
    On Error GoTo 0
    Set ac2fGroupShapes = g
    Exit Function
Fail:
End Function

'=====================================================================
' REPORT
'=====================================================================

Private Function ac2fPanelReport(ByVal made As Long, ByVal skipped As Long, _
                                 ByVal ungrouped As Long, ByVal fold As Double, _
                                 ByVal dir_ As Long, ByVal keepSrc As Boolean, _
                                 ByVal rep As String, ByVal splits As Long, _
                                 ByVal pieces As Long, ByVal joint As Double) As String
    Dim s As String

    s = "RESULT" & vbCrLf
    s = s & "   Panels drawn        : " & made
    If splits > 0 Then
        s = s & "  (" & splits & " split into " & pieces & " plates)"
    End If
    s = s & vbCrLf
    If skipped > 0 Then s = s & "   Skipped             : " & skipped & vbCrLf
    s = s & "   Fold size           : " & ac2fFmt(fold) & " mm  (" & _
            ac2fFmt(fold / 10#) & " cm)" & vbCrLf
    If splits > 0 Then
        s = s & "   Joint allowance     : " & ac2fFmt(joint) & " mm" & vbCrLf
    End If
    s = s & "   Direction           : " & ac2fPanelDirName(dir_) & vbCrLf
    s = s & "   Source rectangle    : " & IIf(keepSrc, "kept", "deleted") & vbCrLf

    If Len(rep) > 0 Then
        s = s & vbCrLf & "PANELS" & vbCrLf & rep
    End If

    If splits > 0 Then
        If m_sep Then
            s = s & vbCrLf & "   Plates laid apart: below the source (wide panel)" & vbCrLf
            s = s & "   or to its right (tall panel)." & vbCrLf
        Else
            s = s & vbCrLf & "   Plates sit in place on the finished size; the joint" & vbCrLf
            s = s & "   allowances overlap. Type sep to lay them apart." & vbCrLf
        End If
    Else
        s = s & vbCrLf & "GROOVE ORDER  (clockwise, object order = cut order)" & vbCrLf
        s = s & "   1 left    starts at the bottom, full height" & vbCrLf
        s = s & "   2 top     starts at the left,   full width" & vbCrLf
        s = s & "   3 right   starts at the top,    full height" & vbCrLf
        s = s & "   4 bottom  starts at the right,  full width" & vbCrLf
    End If
    s = s & vbCrLf & "   layer " & LAYER_CUT & ": sheet outline (black)" & vbCrLf
    s = s & "   layer " & LAYER_GROOVE & ": V-grooves (blue)" & vbCrLf
    s = s & "   The grooves intersect on the folded size." & vbCrLf

    If ungrouped > 0 Then
        s = s & vbCrLf & "WARNING: " & ungrouped & " group(s) could not be made." & _
                vbCrLf & "The lines are on the page but left loose."
    End If

    ' A message box shows about 1024 characters and cuts the rest off.
    If Len(s) > 1000 Then s = Left$(s, 990) & vbCrLf & "..."
    ac2fPanelReport = s
End Function

Public Function ac2fPanelDirName(ByVal d As Long) As String
    If d = AC2F_ACP_IN Then
        ac2fPanelDirName = "in - selection is the sheet"
    Else
        ac2fPanelDirName = "out - selection is the folded size"
    End If
End Function
