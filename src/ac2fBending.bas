Attribute VB_Name = "ac2fBending"
'=====================================================================
'  ac2f pack  --  ac2fBending
'
'  Works out the order channel letters should be bent in, and either
'  shows it, applies it to the drawing, or hands it to whatever sends
'  the job out.
'
'  Macros:
'    ac2fBendOrder   - puts the selected objects in bending order
'    ac2fBendPreview - numbers them on the page, changes nothing else
'    ac2fBendReport  - lists the order, touches nothing
'
'  WHY THE ORDER MATTERS
'  A bender is set up per shape. Taken as they fall, NOBOBOXB is eight
'  setups. Bent in a working order it is four, because the three O go
'  one after another and so do the three B.
'
'  HOW IT WORKS
'
'  1  FLATTENING.  Every selected object is reduced to contours of
'     straight steps whose corners lie ON the real curve. CorelDRAW is
'     asked for the points (Segment.GetPointPositionAt) and the step is
'     refined until the curve stays within Curve tol of it, so an S, a
'     loop or a hand-lettered script comes out as drawn. See CURVE
'     SAMPLING in ac2fCore. (Version 1.22 and earlier guessed a circular
'     arc from chord and length; on anything but a true arc that was
'     wrong by millimetres.)
'
'  2  CONTAINMENT.  Each closed contour is tested against every other:
'     does its first point fall inside, and is the other one bigger?
'     A contour with nothing around it is an OUTSIDE, one inside an
'     outside is a COUNTER (the hole in an A), one inside a counter is
'     an outside again. An outside and everything it encloses make one
'     LETTER.
'
'  3  WORKING ORDER.  Letters go left to right. Letters that overlap in
'     X form a column and are taken top down inside it, which is what
'     "left to right, top first" means on a sheet with rows.
'
'  4  IDENTICAL FIRST.  When a letter is reached, every later letter of
'     the same shape is pulled up behind it, so one setup does them all.
'
'  STRAIGHT BEFORE CURVED
'  Order 2 runs the whole thing twice: first the letters with no curve
'  in them (N, A, K, X, E, M), then the ones with a curve (O, B, P, S,
'  C, G). On a bender with rollers the straight work is finished before
'  the rollers are engaged at all. Order 3 goes further and holds every
'  counter back to the very end.
'
'  THE BUFFER IS SHARED
'  The flattened geometry is public, because ac2fPlot writes HPGL
'  straight out of it. Flattening once and ordering once serves both.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Bending"

Private Const MAX_PTS As Long = 500000

'---------------------------------------------------------------------
' Registry keys (internal, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_BN_ORDER  As String = "BNSiralama"
Public Const AC2F_K_BN_GROUP  As String = "BNAyniParcalariGrupla"
Public Const AC2F_K_BN_COLGAP As String = "BNSutunBindirmesiMM"
Public Const AC2F_K_BN_CURVE  As String = "BNEgrilikEsigiMM"
Public Const AC2F_K_BN_MATCH  As String = "BNEslesmeToleransiMM"
Public Const AC2F_K_BN_ZORDER As String = "BNIstifYonu"
Public Const AC2F_K_BN_RENAME As String = "BNNesneleriAdlandir"

Public Const AC2F_DEF_BN_ORDER  As Long = 1      ' 0 off, 1 left to right, 2 straight first, 3 outsides first
Public Const AC2F_DEF_BN_GROUP  As Long = 1      ' 1 = bend identical letters together
Public Const AC2F_DEF_BN_COLGAP As Double = 5#   ' overlap that keeps two letters in one column (mm)
Public Const AC2F_DEF_BN_CURVE  As Double = 0.5  ' bulge that makes a letter count as curved (mm)
Public Const AC2F_DEF_BN_MATCH  As Double = 0.5  ' how close two letters must be to count as the same (mm)
Public Const AC2F_DEF_BN_ZORDER As Long = 0      ' 0 = first on top, 1 = first at the bottom
Public Const AC2F_DEF_BN_RENAME As Long = 1      ' 1 = number the objects as well

'---------------------------------------------------------------------
' Flattened geometry. Public so ac2fPlot can write HPGL from it without
' a function call per point.
'---------------------------------------------------------------------
Public ac2fGX() As Double          ' point X, mm
Public ac2fGY() As Double          ' point Y, mm
Public ac2fGN As Long              ' how many points
Public ac2fGS() As Long            ' first point of contour i
Public ac2fGC() As Long            ' how many points contour i has
Public ac2fGCl() As Boolean        ' is contour i closed
Public ac2fGSub As Long            ' how many contours

' Members of a group are normally taken top first, the way the Object
' Manager lists them. A caller sending bottom first sets this after
' ac2fBNReset so groups follow the same direction.
Public ac2fBNBackToFront As Boolean

' Per contour, filled as it is flattened and used by the ordering.
Private m_sx0() As Double, m_sy0() As Double
Private m_sx1() As Double, m_sy1() As Double
Private m_sar() As Double          ' signed area, mm2 (its sign spots a mirror)
Private m_slen() As Double         ' outline length, mm
Private m_sbulge() As Double       ' largest departure from a straight line, mm
Private m_sgl() As Long            ' which letter this contour belongs to
Private m_sin() As Boolean         ' True for a counter (the hole in an A)
Private m_sord() As Long           ' contours sorted left to right, top first
Private m_sshp() As Long           ' which drawing object this contour came from

' The objects themselves, so the drawing can be reordered.
Private m_shp() As Shape, m_shpN As Long
Private m_curShape As Long
Private m_redo As Boolean           ' sampling turned out unavailable: read this shape again

' Per letter. A letter is one outside contour plus whatever it encloses.
Private Type ac2fBNGlyph
    X0      As Double
    Y0      As Double
    X1      As Double
    Y1      As Double
    PathLen As Double
    SignArea As Double
    Parts   As Long
    Curved  As Boolean
    Col     As Long
End Type

Private m_g() As ac2fBNGlyph, m_gn As Long
Private m_gsetup() As Long         ' setup number of each letter
Private m_gseq() As Long, m_gseqN As Long
Private m_groups As Long           ' distinct shapes among the letters
Private m_ordered As Long          ' which algorithm this buffer already ran

'=====================================================================
' MACROS
'=====================================================================

' Puts the selected objects into bending order in the drawing itself:
' the stacking order, and their names. Nothing is moved on the page.
Public Sub ac2fBendOrder()
Attribute ac2fBendOrder.VB_Description = "ac2f pack: Put the selected objects in bending order"
    Dim alg As Long
    Dim ord() As Long, n As Long
    Dim i As Long, renamed As Long
    Dim zdir As Long, doName As Long
    Dim s As String

    alg = ac2fBNPrepare()
    If alg = 0 Then Exit Sub

    n = ac2fBNShapeOrder(ord)
    If n = 0 Then
        ac2fWarn "No object could be matched to an outline.", CAPTION_
        Exit Sub
    End If

    zdir = ac2fGetLng(AC2F_K_BN_ZORDER, AC2F_DEF_BN_ZORDER)
    doName = ac2fGetLng(AC2F_K_BN_RENAME, AC2F_DEF_BN_RENAME)

    On Error GoTo Fail
    ActiveDocument.BeginCommandGroup ac2fTitle("bending order")

    ' OrderToFront puts a shape at the top of the stack, so walking the
    ' list backwards leaves the first one to bend on top. Walking it
    ' forwards leaves the first one at the bottom.
    If zdir = 1 Then
        For i = 0 To n - 1
            ac2fBNToFront ord(i)
        Next i
    Else
        For i = n - 1 To 0 Step -1
            ac2fBNToFront ord(i)
        Next i
    End If

    If doName <> 0 Then
        For i = 0 To n - 1
            If ac2fBNRename(ord(i), i + 1) Then renamed = renamed + 1
        Next i
    End If

    ActiveDocument.EndCommandGroup

    s = "BENDING ORDER APPLIED" & vbCrLf
    s = s & "   Rule                : " & ac2fBNOrderName(alg) & vbCrLf
    s = s & "   Objects ordered     : " & n & vbCrLf
    s = s & "   Letters             : " & m_gn & vbCrLf
    s = s & "   Setups              : " & m_groups & vbCrLf
    If doName <> 0 Then
        s = s & "   Objects renamed     : " & renamed & vbCrLf
    End If
    s = s & "   Stack               : " & _
            IIf(zdir = 1, "first one at the bottom", "first one on top") & vbCrLf
    s = s & vbCrLf
    s = s & "The names read [007 S03], meaning seventh to bend, third " & _
            "setup. Objects sharing a setup number are the same shape " & _
            "and are bent one after another." & vbCrLf & vbCrLf
    s = s & "One Undo puts everything back."
    If alg >= 3 Then
        s = s & vbCrLf & vbCrLf & _
            "Order 3 holds counters back to the end. A counter that " & _
            "sits in the SAME object as its outside cannot be pulled " & _
            "out by stacking order. Break such letters apart first, " & _
            "or send with Send order 2, which writes the contours out " & _
            "one by one."
    End If
    ac2fInfo s, CAPTION_
    Exit Sub

Fail:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    ac2fWarn "Ordering the objects failed: " & Err.Description, CAPTION_
End Sub

' Numbers the selection in bending order, on the page, so the order can
' be checked before anything is cut. One Undo clears it.
Public Sub ac2fBendPreview()
Attribute ac2fBendPreview.VB_Description = "ac2f pack: Number the selection in bending order"
    Dim alg As Long
    Dim i As Long, drawn As Long
    Dim t As Shape
    Dim cx As Double, cy As Double, sz As Double
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim numText As String

    alg = ac2fBNPrepare()
    If alg = 0 Then Exit Sub

    On Error GoTo Fail
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    ActiveDocument.BeginCommandGroup ac2fTitle("bending order")
    For i = 0 To ac2fGSub - 1
        cx = (m_sx0(i) + m_sx1(i)) / 2#
        cy = (m_sy0(i) + m_sy1(i)) / 2#
        numText = CStr(i + 1)
        If ac2fGetLng(AC2F_K_BN_GROUP, AC2F_DEF_BN_GROUP) <> 0 Then
            numText = numText & " S" & m_gsetup(m_sgl(i))
        End If
        Set t = Nothing
        On Error Resume Next
        Set t = ActiveLayer.CreateArtisticText(cx, cy, numText)
        If Not t Is Nothing Then
            ' Sized against the outline it belongs to, so the numbers
            ' stay readable on a 2 m letter and on a 30 mm counter.
            sz = (m_sy1(i) - m_sy0(i)) / 5# / 0.3527778
            If sz < 8# Then sz = 8#
            If sz > 200# Then sz = 200#
            t.Text.Story.Size = sz
            drawn = drawn + 1
        End If
        On Error GoTo Fail
    Next i
    ActiveDocument.EndCommandGroup

    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    ac2fInfo "BENDING ORDER" & vbCrLf & _
             "   Rule                : " & ac2fBNOrderName(alg) & vbCrLf & _
             "   Letters             : " & m_gn & vbCrLf & _
             "   Contours            : " & ac2fGSub & "  (" & _
             ac2fBNInnerCount() & " counters)" & vbCrLf & _
             "   Setups              : " & m_groups & vbCrLf & _
             "   Numbers placed      : " & drawn & vbCrLf & vbCrLf & _
             "Each number starts at the middle of the outline it belongs " & _
             "to. S is the setup: same number, same shape, bent one " & _
             "after another." & vbCrLf & vbCrLf & _
             "Nothing else was touched. One Undo removes the numbers.", CAPTION_
    Exit Sub

Fail:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    ac2fWarn "The preview failed: " & Err.Description, CAPTION_
End Sub

' Lists the order without touching the drawing at all.
Public Sub ac2fBendReport()
Attribute ac2fBendReport.VB_Description = "ac2f pack: List the bending order, change nothing"
    Dim alg As Long
    Dim i As Long, shown As Long
    Dim s As String

    alg = ac2fBNPrepare()
    If alg = 0 Then Exit Sub

    s = "BENDING ORDER" & vbCrLf
    s = s & "   Rule                : " & ac2fBNOrderName(alg) & vbCrLf
    s = s & "   Letters             : " & m_gn & vbCrLf
    s = s & "   Contours            : " & ac2fGSub & "  (" & _
            ac2fBNInnerCount() & " counters)" & vbCrLf
    s = s & "   Setups              : " & m_groups & vbCrLf & vbCrLf
    s = s & "  No  Setup  Kind      Size            Length" & vbCrLf

    For i = 0 To ac2fGSub - 1
        If shown >= 30 Then
            s = s & "  ... and " & (ac2fGSub - shown) & " more" & vbCrLf
            Exit For
        End If
        s = s & ac2fPad(CStr(i + 1), 4) & "  " & _
                ac2fPad("S" & m_gsetup(m_sgl(i)), 5) & "  " & _
                ac2fRPad(IIf(m_sin(i), "counter", "outside"), 8) & "  " & _
                ac2fPad(ac2fFmt(m_sx1(i) - m_sx0(i), 1) & " x " & _
                        ac2fFmt(m_sy1(i) - m_sy0(i), 1), 14) & "  " & _
                ac2fPad(ac2fFmt(m_slen(i), 1), 8) & vbCrLf
        shown = shown + 1
    Next i

    s = s & vbCrLf & "Nothing was changed."
    ac2fInfo s, CAPTION_
End Sub

'=====================================================================
' SHARED FRONT END
'=====================================================================

' Checks the selection, flattens it and orders it. Returns the
' algorithm that ran, or 0 after explaining why nothing happened.
Private Function ac2fBNPrepare() As Long
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim tol As Double
    Dim alg As Long

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Function
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Select the letters you want to put in order.", CAPTION_
        Exit Function
    End If
    If sr.Count = 0 Then
        ac2fWarn "Select the letters you want to put in order.", CAPTION_
        Exit Function
    End If

    If Not ac2fBNAskAlgorithm() Then Exit Function

    ' One flattening tolerance serves the whole package; it lives with
    ' the plotter settings because that is where it decides the output.
    tol = ac2fGetNum(AC2F_K_PL_TOL, AC2F_DEF_PL_TOL)
    If tol <= 0 Then tol = AC2F_DEF_PL_TOL

    On Error GoTo Fail
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If
    ac2fBNCollect sr, tol
    alg = ac2fBNOrderNow()
    If unitChanged Then ActiveDocument.Unit = oldUnit

    If ac2fGSub = 0 Then
        ac2fWarn "Nothing with an outline in the selection." & vbCrLf & vbCrLf & _
                 "Bitmaps and objects with no path are skipped.", CAPTION_
        Exit Function
    End If
    If alg = 0 Then
        If ac2fGetLng(AC2F_K_BN_ORDER, AC2F_DEF_BN_ORDER) <= 0 Then
            ac2fWarn "Ordering is switched off." & vbCrLf & vbCrLf & _
                     "Set Order to 1, 2 or 3 in the settings.", CAPTION_
        ElseIf ac2fGSub < 2 Then
            ac2fWarn "There is one outline in the selection, so there is " & _
                     "no order to work out.", CAPTION_
        Else
            ac2fWarn "The order could not be worked out for this " & _
                     "selection, so nothing was changed.", CAPTION_
        End If
        Exit Function
    End If

    ac2fBNPrepare = alg
    Exit Function

Fail:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    ac2fWarn "Reading the selection failed: " & Err.Description, CAPTION_
End Function

' Asks which algorithm to run, offering the one used last. The answer
' is kept, so the next run offers it again. False on Cancel.
Private Function ac2fBNAskAlgorithm() As Boolean
    Dim cur As Long
    Dim answer As String
    Dim v As Long
    Dim m As String

    cur = ac2fGetLng(AC2F_K_BN_ORDER, AC2F_DEF_BN_ORDER)
    If cur < 1 Or cur > 3 Then cur = 1

    m = "Bending order - pick the algorithm:" & vbCrLf & vbCrLf & _
        "  1  -  Left to right, top first" & vbCrLf & vbCrLf & _
        "  2  -  Straight letters first (N, A, K, X ...)," & vbCrLf & _
        "         then curved ones (O, P, B, S ...)" & vbCrLf & vbCrLf & _
        "  3  -  As 2, but every counter (the hole in A," & vbCrLf & _
        "         O, B ...) is left to the very end" & vbCrLf & vbCrLf
    If ac2fGetLng(AC2F_K_BN_GROUP, AC2F_DEF_BN_GROUP) <> 0 Then
        m = m & "In all three, identical letters are bent one " & _
                "after another on the same setup." & vbCrLf & vbCrLf
    End If
    m = m & "The choice is remembered for next time."

    Do
        answer = InputBox(m, ac2fTitle(CAPTION_), CStr(cur))
        If StrPtr(answer) = 0 Then Exit Function
        answer = Trim$(answer)
        If Len(answer) = 0 Then Exit Function
        v = CLng(Val(answer))
        If v >= 1 And v <= 3 Then Exit Do
        ac2fWarn "Enter 1, 2 or 3.", CAPTION_
    Loop

    ac2fSetLng AC2F_K_BN_ORDER, v
    ac2fBNAskAlgorithm = True
End Function

' Objects in bending order: each one sits where its earliest contour
' does. Returns how many, 0 when none could be matched.
Private Function ac2fBNShapeOrder(ByRef ord() As Long) As Long
    Dim key() As Long
    Dim i As Long, k As Long, n As Long
    Dim v As Long, j As Long

    If m_shpN = 0 Then Exit Function
    ReDim key(0 To m_shpN)
    For i = 0 To m_shpN - 1
        key(i) = -1
    Next i
    For i = 0 To ac2fGSub - 1
        k = m_sshp(i)
        If k >= 0 And k < m_shpN Then
            If key(k) < 0 Then key(k) = i
        End If
    Next i

    ReDim ord(0 To m_shpN)
    n = 0
    For i = 0 To m_shpN - 1
        If key(i) >= 0 Then
            ord(n) = i: n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function

    ' insertion sort on the earliest contour
    For i = 1 To n - 1
        v = ord(i)
        j = i - 1
        Do While j >= 0
            If key(ord(j)) > key(v) Then
                ord(j + 1) = ord(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        ord(j + 1) = v
    Next i

    ac2fBNShapeOrder = n
End Function

Private Sub ac2fBNToFront(ByVal k As Long)
    On Error Resume Next
    m_shp(k).OrderToFront
End Sub

' Names an object "[007 S03] whatever it was called". An earlier tag is
' stripped first, so running this twice does not stack them up.
Private Function ac2fBNRename(ByVal k As Long, ByVal seq As Long) As Boolean
    Dim nm As String
    Dim p As Long
    Dim setupNo As Long

    On Error Resume Next
    nm = m_shp(k).Name
    If Left$(nm, 1) = "[" Then
        p = InStr(nm, "] ")
        If p > 0 Then nm = Mid$(nm, p + 2)
    End If

    setupNo = ac2fBNSetupOfShape(k)
    nm = "[" & Format$(seq, "000") & " S" & Format$(setupNo, "00") & "] " & nm
    nm = Left$(nm, 100)
    m_shp(k).Name = nm
    ac2fBNRename = (m_shp(k).Name = nm)
End Function

' The setup number of the first contour that came from this object.
Private Function ac2fBNSetupOfShape(ByVal k As Long) As Long
    Dim i As Long
    For i = 0 To ac2fGSub - 1
        If m_sshp(i) = k Then
            ac2fBNSetupOfShape = m_gsetup(m_sgl(i))
            Exit Function
        End If
    Next i
End Function

Private Sub ac2fBNAddShape(ByVal s As Shape)
    If m_shpN > UBound(m_shp) Then
        ReDim Preserve m_shp(0 To (UBound(m_shp) + 1) * 2 - 1)
    End If
    Set m_shp(m_shpN) = s
    m_curShape = m_shpN
    m_shpN = m_shpN + 1
End Sub

'=====================================================================
' WHAT THE OTHER MODULES ASK FOR
'=====================================================================

Public Function ac2fBNLetters() As Long
    ac2fBNLetters = m_gn
End Function

Public Function ac2fBNSetups() As Long
    ac2fBNSetups = m_groups
End Function

'=====================================================================
' ENGINE
'=====================================================================

'---------------------------------------------------------------------
' Flatten one shape into the point buffer.
'---------------------------------------------------------------------
Private Sub ac2fBNShape(ByVal s As Shape, ByVal tol As Double)
    Dim i As Long
    Dim cv As Curve
    Dim dup As Shape
    Dim gn0 As Long, gs0 As Long

    If s Is Nothing Then Exit Sub
    On Error GoTo Skip

    Select Case s.Type
        Case cdrBitmapShape, cdrOLEObjectShape
            Exit Sub
    End Select

    If s.Type = cdrGroupShape Then
        If ac2fBNBackToFront Then
            For i = s.Shapes.Count To 1 Step -1
                ac2fBNShape s.Shapes(i), tol
            Next i
        Else
            For i = 1 To s.Shapes.Count
                ac2fBNShape s.Shapes(i), tol
            Next i
        End If
        Exit Sub
    End If

    ' From here on this is a leaf object, and every contour it produces
    ' belongs to it. Reordering the drawing needs that link; a group is
    ' not recorded, because what moves in the stack is the leaf.
    ac2fBNAddShape s

    gn0 = ac2fGN: gs0 = ac2fGSub

    ' Normal route: read the displayed curve as it is. It is not changed.
    Set cv = Nothing
    If ac2fSegMode() >= 0 Then
        On Error Resume Next
        Set cv = s.DisplayCurve
        On Error GoTo Skip
        If Not cv Is Nothing Then
            m_redo = False
            ac2fBNCurve cv, tol
            If Not m_redo Then Exit Sub
            ' CorelDRAW gave no usable points along a segment: forget
            ' what this shape produced and read it again by cutting a copy.
            ac2fGN = gn0: ac2fGSub = gs0
        End If
    End If

    ' Copy route: the object is not a plain curve (text, rectangle ...), or
    ' this CorelDRAW cannot give points along a segment. Either way the
    ' work is done on a duplicate, which is deleted afterwards.
    Set dup = s.Duplicate(0, 0)
    On Error Resume Next
    dup.ConvertToCurves
    Set cv = dup.Curve
    On Error GoTo SkipDup
    If cv Is Nothing Then GoTo SkipDup
    m_redo = False
    ac2fBNCurve cv, tol
    If m_redo Then
        ' The mode has just turned to cutting: go through the copy again.
        ac2fGN = gn0: ac2fGSub = gs0
        ac2fBNCurve cv, tol
    End If
    dup.Delete
    Exit Sub

SkipDup:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
    Exit Sub
Skip:
End Sub

'---------------------------------------------------------------------
' Flatten one curve into contours of points that lie on the real curve
' (see CURVE SAMPLING in ac2fCore). tol is how far the polyline may
' stray from the drawing. Besides the points, the largest departure of
' each contour from its own chords is kept (m_sbulge): it tells a round
' letter from a straight one.
'---------------------------------------------------------------------
Private Sub ac2fBNCurve(ByVal cv As Curve, ByVal tol As Double)
    Dim i As Long, k As Long, j As Long, m As Long, n As Long
    Dim nSub As Long, nSeg As Long
    Dim sp As SubPath, sg As Segment
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim chord As Double, arcL As Double
    Dim bulge As Double, maxBulge As Double

    On Error Resume Next
    nSub = cv.SubPaths.Count
    On Error GoTo 0
    If nSub = 0 Then Exit Sub

    For i = 1 To nSub
        Set sp = Nothing
        On Error Resume Next
        Set sp = cv.SubPaths(i)
        nSeg = sp.Segments.Count
        On Error GoTo 0
        If Not sp Is Nothing Then
            If nSeg > 0 Then
                ac2fBNEnsure
                m_sshp(ac2fGSub) = m_curShape
                ac2fGS(ac2fGSub) = ac2fGN
                ac2fGCl(ac2fGSub) = False
                On Error Resume Next
                ac2fGCl(ac2fGSub) = sp.Closed
                On Error GoTo 0
                maxBulge = 0

                k = 1
                Do While k <= nSeg
                    Set sg = sp.Segments(k)

                    If ac2fSegMode() < 0 Then
                        ' Cutting mode, on a copy: halve the segment until
                        ' every piece is within tolerance of its chord, then
                        ' take the pieces' start nodes. The bulge is taken
                        ' from the segment as it was, before it is cut.
                        arcL = sg.Length
                        ax = sg.StartNode.PositionX: ay = sg.StartNode.PositionY
                        bx = sg.EndNode.PositionX: by = sg.EndNode.PositionY
                        chord = Sqr((bx - ax) * (bx - ax) + (by - ay) * (by - ay))
                        bulge = ac2fSegBulge(arcL, chord)
                        m = ac2fSegSplit(sp, k, tol)
                        For j = 0 To m - 1
                            Set sg = sp.Segments(k + j)
                            ac2fBNAddPt sg.StartNode.PositionX, sg.StartNode.PositionY
                        Next j
                        nSeg = sp.Segments.Count
                        k = k + m
                    Else
                        n = ac2fSegmentPoints(sg, tol)
                        If n < 0 Then
                            m_redo = True
                            Exit Sub
                        End If
                        For j = 0 To n - 1
                            ac2fBNAddPt ac2fSgX(j), ac2fSgY(j)
                        Next j
                        bulge = ac2fSgDev
                        k = k + 1
                    End If
                    If bulge > maxBulge Then maxBulge = bulge
                    If ac2fGN >= MAX_PTS Then Exit Do
                Loop

                ' last node of the subpath
                Set sg = sp.Segments(nSeg)
                ac2fBNAddPt sg.EndNode.PositionX, sg.EndNode.PositionY

                ac2fGC(ac2fGSub) = ac2fGN - ac2fGS(ac2fGSub)
                If ac2fGC(ac2fGSub) >= 2 Then
                    m_sbulge(ac2fGSub) = maxBulge
                    ac2fBNFinish ac2fGSub
                    ac2fGSub = ac2fGSub + 1
                Else
                    ac2fGN = ac2fGS(ac2fGSub)
                End If
            End If
        End If
        If ac2fGN >= MAX_PTS Then Exit For
    Next i
End Sub

Public Sub ac2fBNAddPt(ByVal x As Double, ByVal y As Double)
    If ac2fGN > UBound(ac2fGX) Then
        ReDim Preserve ac2fGX(0 To (UBound(ac2fGX) + 1) * 2 - 1)
        ReDim Preserve ac2fGY(0 To UBound(ac2fGX))
    End If
    ac2fGX(ac2fGN) = x
    ac2fGY(ac2fGN) = y
    ac2fGN = ac2fGN + 1
End Sub

'=====================================================================
' ORDER
'
' The order outlines leave in decides how much the bender is set up
' and reset. Three steps do the work.
'
' 1  CONTAINMENT.  Every closed contour is tested against every other:
'    does its first point fall inside, and is the other one bigger?
'    A contour with no container is an outside, one inside an outside
'    is a counter, one inside a counter is an island again. An outside
'    and everything it encloses make one letter.
'
' 2  WORKING ORDER.  Letters are sorted left to right. Letters that
'    overlap in X form a column and are taken top down inside it, which
'    is what "left to right, top first" means on a sheet with rows.
'
' 3  IDENTICAL FIRST.  When a letter is reached, every later letter
'    that is the same shape is pulled up behind it. Three identical O
'    in a word are bent one after another, on one setup, instead of
'    three times with B and X in between.
'
' A letter counts as the same when it has the same number of contours,
' the same length, the same enclosed area and the same bounding box,
' each within the match tolerance. Width and height may swap over, so a
' copy turned a quarter turn still counts. The sign of the area has to
' match as well, which keeps a mirrored copy out: it bends the other
' way and is not the same job.
'
' Algorithm 2 runs the whole thing twice, straight-sided letters first
' and curved ones after, so the rollers are engaged once. Algorithm 3
' goes further and holds every counter back to the end, so all the
' outsides are finished before the first hole is started.
'=====================================================================

Public Sub ac2fBNReset()
    ac2fGN = 0: ac2fGSub = 0: m_gn = 0: m_gseqN = 0: m_groups = 0
    m_shpN = 0: m_curShape = -1: m_ordered = 0
    ac2fBNBackToFront = False
    ac2fSegReset
    ReDim m_shp(0 To 255)
    ReDim m_sshp(0 To 255)
    ReDim ac2fGX(0 To 8191): ReDim ac2fGY(0 To 8191)
    ReDim ac2fGS(0 To 255): ReDim ac2fGC(0 To 255): ReDim ac2fGCl(0 To 255)
    ReDim m_sx0(0 To 255): ReDim m_sy0(0 To 255)
    ReDim m_sx1(0 To 255): ReDim m_sy1(0 To 255)
    ReDim m_sar(0 To 255): ReDim m_slen(0 To 255)
    ReDim m_sbulge(0 To 255): ReDim m_sgl(0 To 255)
    ReDim m_sin(0 To 255): ReDim m_sord(0 To 255)
End Sub

Private Sub ac2fBNEnsure()
    Dim nu As Long
    If ac2fGSub <= UBound(ac2fGS) Then Exit Sub
    nu = (UBound(ac2fGS) + 1) * 2 - 1
    ReDim Preserve ac2fGS(0 To nu)
    ReDim Preserve ac2fGC(0 To nu)
    ReDim Preserve ac2fGCl(0 To nu)
    ReDim Preserve m_sx0(0 To nu)
    ReDim Preserve m_sy0(0 To nu)
    ReDim Preserve m_sx1(0 To nu)
    ReDim Preserve m_sy1(0 To nu)
    ReDim Preserve m_sar(0 To nu)
    ReDim Preserve m_slen(0 To nu)
    ReDim Preserve m_sbulge(0 To nu)
    ReDim Preserve m_sgl(0 To nu)
    ReDim Preserve m_sin(0 To nu)
    ReDim Preserve m_sord(0 To nu)
    ReDim Preserve m_sshp(0 To nu)
End Sub

' Box, length and signed area of one finished contour. The area is
' taken round the closed ring either way, since an open path still has
' to be told from its neighbours; only the length leaves out the
' closing step when the path is open.
Private Sub ac2fBNFinish(ByVal i As Long)
    Dim s As Long, n As Long, k As Long
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim acc As Double, plen As Double

    s = ac2fGS(i): n = ac2fGC(i)
    m_sx0(i) = ac2fGX(s): m_sx1(i) = ac2fGX(s)
    m_sy0(i) = ac2fGY(s): m_sy1(i) = ac2fGY(s)

    For k = 0 To n - 1
        ax = ac2fGX(s + k): ay = ac2fGY(s + k)
        If ax < m_sx0(i) Then m_sx0(i) = ax
        If ax > m_sx1(i) Then m_sx1(i) = ax
        If ay < m_sy0(i) Then m_sy0(i) = ay
        If ay > m_sy1(i) Then m_sy1(i) = ay

        If k = n - 1 Then
            bx = ac2fGX(s): by = ac2fGY(s)
        Else
            bx = ac2fGX(s + k + 1): by = ac2fGY(s + k + 1)
        End If
        acc = acc + (ax * by - bx * ay)
        If k < n - 1 Or ac2fGCl(i) Then
            plen = plen + Sqr((bx - ax) * (bx - ax) + (by - ay) * (by - ay))
        End If
    Next k

    m_sar(i) = acc / 2#
    m_slen(i) = plen
End Sub

Public Sub ac2fBNCollect(ByVal sr As ShapeRange, ByVal tol As Double)
    Dim i As Long
    ac2fBNReset
    For i = 1 To sr.Count
        ac2fBNShape sr(i), tol
    Next i
End Sub

' Flattens one more object onto the end of the buffer, for callers that
' decide the order of the objects themselves. Call ac2fBNReset first.
Public Sub ac2fBNAddObject(ByVal s As Shape, ByVal tol As Double)
    ac2fBNShape s, tol
End Sub

' Containment, letters, and the contour list sorted left to right.
Private Sub ac2fBNAnalyse()
    Dim i As Long, j As Long, g As Long
    Dim best As Long, bestA As Double, a As Double
    Dim par() As Long
    Dim depth As Long, q As Long, guard As Long
    Dim curveMin As Double

    If ac2fGSub = 0 Then Exit Sub
    curveMin = ac2fGetNum(AC2F_K_BN_CURVE, AC2F_DEF_BN_CURVE)
    If curveMin < 0 Then curveMin = 0

    ReDim par(0 To ac2fGSub)
    For i = 0 To ac2fGSub - 1
        best = -1: bestA = 0
        For j = 0 To ac2fGSub - 1
            If j <> i Then
                If ac2fGCl(j) Then
                    ' only a strictly larger contour can be the parent,
                    ' which also stops two coincident copies pointing at
                    ' each other and making a loop
                    If Abs(m_sar(j)) > Abs(m_sar(i)) Then
                        If m_sx0(j) <= m_sx0(i) And m_sy0(j) <= m_sy0(i) And _
                           m_sx1(j) >= m_sx1(i) And m_sy1(j) >= m_sy1(i) Then
                            If ac2fBNInside(i, j) Then
                                a = Abs(m_sar(j))
                                If best < 0 Or a < bestA Then
                                    best = j: bestA = a
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        par(i) = best
    Next i

    For i = 0 To ac2fGSub - 1
        depth = 0: q = par(i): guard = 0
        Do While q >= 0 And guard <= ac2fGSub
            depth = depth + 1: q = par(q): guard = guard + 1
        Loop
        m_sin(i) = ((depth Mod 2) = 1)
        m_sgl(i) = -1
    Next i

    m_gn = 0
    For i = 0 To ac2fGSub - 1
        If par(i) < 0 Then
            m_sgl(i) = m_gn
            m_gn = m_gn + 1
        End If
    Next i
    For i = 0 To ac2fGSub - 1
        If m_sgl(i) < 0 Then
            q = i: guard = 0
            Do While par(q) >= 0 And guard <= ac2fGSub
                q = par(q): guard = guard + 1
            Loop
            ' The walk always ends on a root, which already has its
            ' number. Anything else would mean a loop in the chain, and
            ' that contour is better off standing alone than indexing
            ' past the end of the letter table.
            If m_sgl(q) >= 0 Then
                m_sgl(i) = m_sgl(q)
            Else
                m_sgl(i) = m_gn
                m_gn = m_gn + 1
            End If
        End If
    Next i

    ReDim m_g(0 To m_gn)
    ReDim m_gsetup(0 To m_gn)
    For i = 0 To ac2fGSub - 1
        g = m_sgl(i)
        If m_g(g).Parts = 0 Then
            m_g(g).X0 = m_sx0(i): m_g(g).X1 = m_sx1(i)
            m_g(g).Y0 = m_sy0(i): m_g(g).Y1 = m_sy1(i)
        Else
            If m_sx0(i) < m_g(g).X0 Then m_g(g).X0 = m_sx0(i)
            If m_sx1(i) > m_g(g).X1 Then m_g(g).X1 = m_sx1(i)
            If m_sy0(i) < m_g(g).Y0 Then m_g(g).Y0 = m_sy0(i)
            If m_sy1(i) > m_g(g).Y1 Then m_g(g).Y1 = m_sy1(i)
        End If
        m_g(g).Parts = m_g(g).Parts + 1
        m_g(g).PathLen = m_g(g).PathLen + m_slen(i)
        m_g(g).SignArea = m_g(g).SignArea + m_sar(i)
        If m_sbulge(i) >= curveMin Then m_g(g).Curved = True
    Next i

    ac2fBNSortContours
End Sub

' Is the first point of contour i inside contour j? Even-odd ray cast.
Private Function ac2fBNInside(ByVal i As Long, ByVal j As Long) As Boolean
    Dim px As Double, py As Double
    Dim s As Long, n As Long, k As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim inside As Boolean

    px = ac2fGX(ac2fGS(i)): py = ac2fGY(ac2fGS(i))
    s = ac2fGS(j): n = ac2fGC(j)
    For k = 0 To n - 1
        x1 = ac2fGX(s + k): y1 = ac2fGY(s + k)
        If k = n - 1 Then
            x2 = ac2fGX(s): y2 = ac2fGY(s)
        Else
            x2 = ac2fGX(s + k + 1): y2 = ac2fGY(s + k + 1)
        End If
        If (y1 > py) <> (y2 > py) Then
            If px < x1 + (py - y1) * (x2 - x1) / (y2 - y1) Then inside = Not inside
        End If
    Next k
    ac2fBNInside = inside
End Function

' Contours left to right, top first. Used inside a letter so that two
' identical letters always have their counters in the same order.
Private Sub ac2fBNSortContours()
    Dim i As Long, j As Long, v As Long
    For i = 0 To ac2fGSub - 1
        m_sord(i) = i
    Next i
    For i = 1 To ac2fGSub - 1
        v = m_sord(i)
        j = i - 1
        Do While j >= 0
            If m_sx0(m_sord(j)) > m_sx0(v) Or _
               (m_sx0(m_sord(j)) = m_sx0(v) And m_sy1(m_sord(j)) < m_sy1(v)) Then
                m_sord(j + 1) = m_sord(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        m_sord(j + 1) = v
    Next i
End Sub

'---------------------------------------------------------------------
' Letter order
'---------------------------------------------------------------------
Private Sub ac2fBNBuildOrder(ByVal alg As Long)
    ReDim m_gseq(0 To m_gn)
    m_gseqN = 0
    m_groups = 0
    If alg >= 2 Then
        ac2fBNPhase 0          ' straight sided
        ac2fBNPhase 1          ' curved
    Else
        ac2fBNPhase 2          ' all of them, in one pass
    End If
End Sub

' One pass of the ordering over the letters that match 'want':
' 0 the straight sided ones, 1 the curved ones, 2 all of them.
Private Sub ac2fBNPhase(ByVal want As Long)
    Dim idx() As Long, n As Long, i As Long
    Dim takeIt As Boolean

    If m_gn = 0 Then Exit Sub
    ReDim idx(0 To m_gn)
    n = 0
    For i = 0 To m_gn - 1
        If want = 2 Then
            takeIt = True
        ElseIf want = 1 Then
            takeIt = m_g(i).Curved
        Else
            takeIt = Not m_g(i).Curved
        End If
        If takeIt Then
            idx(n) = i: n = n + 1
        End If
    Next i
    If n = 0 Then Exit Sub

    ac2fBNSortLR idx, n
    ac2fBNColumns idx, n
    ac2fBNSortCols idx, n
    If ac2fGetLng(AC2F_K_BN_GROUP, AC2F_DEF_BN_GROUP) <> 0 Then
        ac2fBNGroupSame idx, n
    Else
        For i = 0 To n - 1
            m_groups = m_groups + 1
            m_gsetup(idx(i)) = m_groups
        Next i
    End If

    For i = 0 To n - 1
        m_gseq(m_gseqN) = idx(i)
        m_gseqN = m_gseqN + 1
    Next i
End Sub

Private Sub ac2fBNSortLR(ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, j As Long, v As Long
    For i = 1 To n - 1
        v = idx(i)
        j = i - 1
        Do While j >= 0
            If m_g(idx(j)).X0 > m_g(v).X0 Or _
               (m_g(idx(j)).X0 = m_g(v).X0 And m_g(idx(j)).Y1 < m_g(v).Y1) Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = v
    Next i
End Sub

' Letters that overlap in X belong to the same column. A new column
' starts where a letter's left edge clears the column's right edge by
' more than the overlap allowance.
Private Sub ac2fBNColumns(ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, col As Long
    Dim rightEdge As Double, gap As Double

    gap = ac2fGetNum(AC2F_K_BN_COLGAP, AC2F_DEF_BN_COLGAP)
    col = 0
    rightEdge = m_g(idx(0)).X1
    m_g(idx(0)).Col = 0
    For i = 1 To n - 1
        If m_g(idx(i)).X0 >= rightEdge - gap Then
            col = col + 1
            rightEdge = m_g(idx(i)).X1
        ElseIf m_g(idx(i)).X1 > rightEdge Then
            rightEdge = m_g(idx(i)).X1
        End If
        m_g(idx(i)).Col = col
    Next i
End Sub

Private Sub ac2fBNSortCols(ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, j As Long, v As Long
    For i = 1 To n - 1
        v = idx(i)
        j = i - 1
        Do While j >= 0
            If ac2fBNColAfter(idx(j), v) Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = v
    Next i
End Sub

' True when letter a should come after letter b: later column, or the
' same column and lower down, or level and further right.
Private Function ac2fBNColAfter(ByVal a As Long, ByVal b As Long) As Boolean
    If m_g(a).Col <> m_g(b).Col Then
        ac2fBNColAfter = (m_g(a).Col > m_g(b).Col)
        Exit Function
    End If
    If Abs(m_g(a).Y1 - m_g(b).Y1) > 0.001 Then
        ac2fBNColAfter = (m_g(a).Y1 < m_g(b).Y1)
        Exit Function
    End If
    ac2fBNColAfter = (m_g(a).X0 > m_g(b).X0)
End Function

' Pulls every later copy of a letter up behind the first one.
Private Sub ac2fBNGroupSame(ByRef idx() As Long, ByVal n As Long)
    Dim out() As Long, used() As Boolean
    Dim i As Long, j As Long, k As Long

    ReDim out(0 To n)
    ReDim used(0 To n)
    k = 0
    For i = 0 To n - 1
        If Not used(i) Then
            used(i) = True
            out(k) = idx(i): k = k + 1
            m_groups = m_groups + 1
            m_gsetup(idx(i)) = m_groups
            For j = i + 1 To n - 1
                If Not used(j) Then
                    If ac2fBNSameGlyph(idx(i), idx(j)) Then
                        used(j) = True
                        out(k) = idx(j): k = k + 1
                        m_gsetup(idx(j)) = m_groups
                    End If
                End If
            Next j
        End If
    Next i
    For i = 0 To n - 1
        idx(i) = out(i)
    Next i
End Sub

' Same shape, so the same setup on the bender.
Private Function ac2fBNSameGlyph(ByVal a As Long, ByVal b As Long) As Boolean
    Dim tol As Double, t As Double
    Dim w1 As Double, h1 As Double, w2 As Double, h2 As Double

    tol = ac2fGetNum(AC2F_K_BN_MATCH, AC2F_DEF_BN_MATCH)
    If tol < 0 Then tol = 0

    If m_g(a).Parts <> m_g(b).Parts Then Exit Function
    If m_g(a).Curved <> m_g(b).Curved Then Exit Function

    t = tol
    If 0.002 * m_g(a).PathLen > t Then t = 0.002 * m_g(a).PathLen
    If Abs(m_g(a).PathLen - m_g(b).PathLen) > t Then Exit Function

    ' A mirrored copy has the same length and the same box but the other
    ' sign of area, and it bends the other way, so it is not the same job.
    If (m_g(a).SignArea < 0) <> (m_g(b).SignArea < 0) Then Exit Function

    t = tol * tol * 10#
    If 0.004 * Abs(m_g(a).SignArea) > t Then t = 0.004 * Abs(m_g(a).SignArea)
    If Abs(Abs(m_g(a).SignArea) - Abs(m_g(b).SignArea)) > t Then Exit Function

    w1 = m_g(a).X1 - m_g(a).X0: h1 = m_g(a).Y1 - m_g(a).Y0
    w2 = m_g(b).X1 - m_g(b).X0: h2 = m_g(b).Y1 - m_g(b).Y0
    If w1 > h1 Then t = w1: w1 = h1: h1 = t
    If w2 > h2 Then t = w2: w2 = h2: h2 = t
    If Abs(w1 - w2) > tol Then Exit Function
    If Abs(h1 - h2) > tol Then Exit Function

    ac2fBNSameGlyph = True
End Function

'---------------------------------------------------------------------
' Turning the letter order into a contour order
'---------------------------------------------------------------------
Private Function ac2fBNApplyOrder(ByVal alg As Long) As Boolean
    Dim perm() As Long, np As Long
    Dim i As Long

    If ac2fGSub = 0 Then Exit Function
    ReDim perm(0 To ac2fGSub)
    np = 0

    If alg >= 3 Then
        ' every outside first, then every counter in the same order
        For i = 0 To m_gseqN - 1
            ac2fBNTake m_gseq(i), False, perm, np
        Next i
        For i = 0 To m_gseqN - 1
            ac2fBNTake m_gseq(i), True, perm, np
        Next i
    Else
        For i = 0 To m_gseqN - 1
            ac2fBNTake m_gseq(i), False, perm, np
            ac2fBNTake m_gseq(i), True, perm, np
        Next i
    End If

    ' Anything short of a complete permutation is left alone rather than
    ' sent half ordered and half not.
    If np <> ac2fGSub Then Exit Function

    ac2fBNPermute perm
    ac2fBNApplyOrder = True
End Function

Private Sub ac2fBNTake(ByVal g As Long, ByVal wantInner As Boolean, _
                       ByRef perm() As Long, ByRef np As Long)
    Dim k As Long, c As Long
    For k = 0 To ac2fGSub - 1
        c = m_sord(k)
        If m_sgl(c) = g Then
            If m_sin(c) = wantInner Then
                perm(np) = c
                np = np + 1
            End If
        End If
    Next k
End Sub

' Only the contour index arrays move. The points stay where they are.
Private Sub ac2fBNPermute(ByRef perm() As Long)
    Dim i As Long, c As Long
    Dim ss() As Long, sc() As Long, scl() As Boolean
    Dim x0() As Double, y0() As Double, x1() As Double, y1() As Double
    Dim ar() As Double, ln() As Double, bg() As Double
    Dim gl() As Long, inn() As Boolean
    Dim sh() As Long

    ReDim ss(0 To ac2fGSub): ReDim sc(0 To ac2fGSub): ReDim scl(0 To ac2fGSub)
    ReDim x0(0 To ac2fGSub): ReDim y0(0 To ac2fGSub)
    ReDim x1(0 To ac2fGSub): ReDim y1(0 To ac2fGSub)
    ReDim ar(0 To ac2fGSub): ReDim ln(0 To ac2fGSub): ReDim bg(0 To ac2fGSub)
    ReDim gl(0 To ac2fGSub): ReDim inn(0 To ac2fGSub)
    ReDim sh(0 To ac2fGSub)

    For i = 0 To ac2fGSub - 1
        c = perm(i)
        ss(i) = ac2fGS(c): sc(i) = ac2fGC(c): scl(i) = ac2fGCl(c)
        x0(i) = m_sx0(c): y0(i) = m_sy0(c)
        x1(i) = m_sx1(c): y1(i) = m_sy1(c)
        ar(i) = m_sar(c): ln(i) = m_slen(c): bg(i) = m_sbulge(c)
        gl(i) = m_sgl(c): inn(i) = m_sin(c): sh(i) = m_sshp(c)
    Next i
    For i = 0 To ac2fGSub - 1
        ac2fGS(i) = ss(i): ac2fGC(i) = sc(i): ac2fGCl(i) = scl(i)
        m_sx0(i) = x0(i): m_sy0(i) = y0(i)
        m_sx1(i) = x1(i): m_sy1(i) = y1(i)
        m_sar(i) = ar(i): m_slen(i) = ln(i): m_sbulge(i) = bg(i)
        m_sgl(i) = gl(i): m_sin(i) = inn(i): m_sshp(i) = sh(i)
    Next i
End Sub

' Orders what has just been collected and reports which algorithm ran.
' 0 means the contours were left in the order the document gave them.
Public Function ac2fBNOrderNow() As Long
    Dim alg As Long
    ' Ordering an already ordered buffer would reorder a reordering.
    ' One collection, one ordering; a second call just says what ran.
    If m_ordered <> 0 Then
        ac2fBNOrderNow = m_ordered
        Exit Function
    End If
    alg = ac2fGetLng(AC2F_K_BN_ORDER, AC2F_DEF_BN_ORDER)
    If alg <= 0 Or ac2fGSub < 2 Then Exit Function
    If alg > 3 Then alg = 3
    ac2fBNAnalyse
    ac2fBNBuildOrder alg
    If Not ac2fBNApplyOrder(alg) Then Exit Function
    m_ordered = alg
    ac2fBNOrderNow = alg
End Function

Public Function ac2fBNInnerCount() As Long
    Dim i As Long, n As Long
    For i = 0 To ac2fGSub - 1
        If m_sin(i) Then n = n + 1
    Next i
    ac2fBNInnerCount = n
End Function

Public Function ac2fBNOrderName(ByVal alg As Long) As String
    Select Case alg
        Case 1:    ac2fBNOrderName = "left to right, top first"
        Case 2:    ac2fBNOrderName = "straight sided first, then curved"
        Case 3:    ac2fBNOrderName = "outsides first, counters last"
        Case Else: ac2fBNOrderName = "as the document has them"
    End Select
End Function
