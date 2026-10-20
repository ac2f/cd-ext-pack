Attribute VB_Name = "ac2fStart"
'=====================================================================
'  ac2f pack  --  ac2fStart
'
'  Start points: moves the first node of every path to a chosen place -
'  a corner, the middle of a side or the centre of its box - so the
'  laser or router enters each part where you want it. Optionally sets
'  the cutting direction too.
'
'  Macros:
'    ac2fStartPoints      - asks where, then sets the selection
'    ac2fStartPointsAgain - same with the last answer, no question
'    ac2fStartPointsShow  - draws (or removes) markers at every start
'
'  POSITIONS follow the numeric keypad:   7 8 9
'                                          4 5 6
'                                          1 2 3
'  measured on each contour's own box, each object's box or the box of
'  the whole selection.
'
'  HOW
'  The point of the path nearest that position is found by reading
'  points along every segment. In "exact" mode a node is added there
'  (the shape does not change) unless a node already sits on it. The
'  path is then broken at that node and joined again, which makes it the
'  start node - the method CorelDRAW users have long used by hand.
'
'  Rectangles, ellipses and polygons have no nodes to move and are
'  converted to curves first; text is left alone.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Start Points"
Private Const MARK_LAYER As String = "ac2f start points"

'---------------------------------------------------------------------
' Registry keys (internal, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_SP_POS   As String = "SPKonum"
Public Const AC2F_K_SP_REF   As String = "SPReferans"
Public Const AC2F_K_SP_EXACT As String = "SPTamNokta"
Public Const AC2F_K_SP_DIR   As String = "SPYon"
Public Const AC2F_K_SP_OPEN  As String = "SPAcikYol"
Public Const AC2F_K_SP_MARK  As String = "SPIsaretle"

Public Const AC2F_DEF_SP_POS   As Long = 7    ' keypad: top left
Public Const AC2F_DEF_SP_REF   As Long = 1    ' 1 contour, 2 object, 3 selection
Public Const AC2F_DEF_SP_EXACT As Long = 1    ' 1 exact point, 0 nearest node
Public Const AC2F_DEF_SP_DIR   As Long = 0    ' 0 keep, 1..4 see ac2fSPDirText
Public Const AC2F_DEF_SP_OPEN  As Long = 1    ' open paths start at the nearer end
Public Const AC2F_DEF_SP_MARK  As Long = 0    ' draw markers afterwards

' settings of this run
Private m_pos As Long, m_ref As Long, m_dir As Long
Private m_exact As Boolean, m_open As Boolean, m_mark As Boolean

' sampled points of the subpaths of one curve: per subpath a range
Private m_x() As Double, m_y() As Double
Private m_seg() As Long, m_t() As Double
Private m_n As Long
Private m_s0() As Long, m_sc() As Long, m_closed() As Boolean
Private m_ns As Long
Private m_noSample As Boolean

' counters for the report
Private m_paths As Long, m_moved As Long, m_added As Long, m_reversed As Long
Private m_failed As Long, m_noReverse As Long, m_converted As Long, m_text As Long

' the boxes the positions are measured on
Private m_selX0 As Double, m_selY0 As Double, m_selX1 As Double, m_selY1 As Double
Private m_objX0 As Double, m_objY0 As Double, m_objX1 As Double, m_objY1 As Double

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fStartPoints()
Attribute ac2fStartPoints.VB_Description = "ac2f pack: Set where every path starts (corner, side, centre)"
    Dim ans As String

    If Not ac2fSPCheck() Then Exit Sub
    ac2fSPLoad
    ans = InputBox(ac2fSPPrompt(), ac2fTitle(CAPTION_), ac2fSPCode())
    If StrPtr(ans) = 0 Then Exit Sub
    If Len(Trim$(ans)) = 0 Then Exit Sub
    If Not ac2fSPParse(ans) Then
        ac2fWarn "Type a number from 1 to 9, optionally followed by letters." & vbCrLf & _
                 "For example:  7   or   8 S W", CAPTION_
        Exit Sub
    End If
    ac2fSPSave
    ac2fSPRunSelection
End Sub

Public Sub ac2fStartPointsAgain()
Attribute ac2fStartPointsAgain.VB_Description = "ac2f pack: Set start points with the last settings"
    If Not ac2fSPCheck() Then Exit Sub
    ac2fSPLoad
    ac2fSPRunSelection
End Sub

Public Sub ac2fStartPointsShow()
Attribute ac2fStartPointsShow.VB_Description = "ac2f pack: Show or hide markers at every start point"
    Dim ly As Layer
    Dim sr As ShapeRange
    Dim n As Long

    If ActiveDocument Is Nothing Then Exit Sub
    Set ly = ac2fSPFindLayer(False)
    If Not ly Is Nothing Then
        If ly.Shapes.Count > 0 Then
            ActiveDocument.BeginCommandGroup ac2fTitle("start markers")
            ly.Shapes.All.Delete
            ActiveDocument.EndCommandGroup
            ac2fInfo "Start point markers removed.", CAPTION_
            Exit Sub
        End If
    End If
    If Not ac2fSPCheck() Then Exit Sub
    Set sr = ActiveSelectionRange
    ActiveDocument.BeginCommandGroup ac2fTitle("start markers")
    n = ac2fSPMarkRange(sr)
    ActiveDocument.EndCommandGroup
    ac2fInfo n & " start point(s) marked on the layer """ & MARK_LAYER & """." & vbCrLf & _
             "The circle is the start, the short line shows the direction." & vbCrLf & _
             "Run this again to remove them.", CAPTION_
End Sub

' For other modules (nesting): sets the start points of the given shapes
' with the saved settings, no questions. Returns the paths changed.
Public Function ac2fStartApplyShapes(ByRef shp() As Shape, ByVal n As Long) As Long
    Dim i As Long
    Dim first As Boolean
    Dim x As Double, y As Double, w As Double, h As Double

    ac2fSPLoad
    ac2fSPResetCounts
    first = True
    For i = 0 To n - 1
        If Not shp(i) Is Nothing Then
            shp(i).GetBoundingBox x, y, w, h
            ac2fSPGrow first, x, y, w, h
        End If
    Next i
    For i = 0 To n - 1
        If Not shp(i) Is Nothing Then ac2fSPTop shp(i)
    Next i
    ac2fStartApplyShapes = m_moved + m_reversed
End Function

'=====================================================================
' FLOW
'=====================================================================

Private Function ac2fSPCheck() As Boolean
    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Function
    End If
    If ActiveSelectionRange Is Nothing Then
        ac2fWarn "Select the paths first.", CAPTION_
        Exit Function
    End If
    If ActiveSelectionRange.Count = 0 Then
        ac2fWarn "Select the paths first.", CAPTION_
        Exit Function
    End If
    ac2fSPCheck = True
End Function

Private Sub ac2fSPRunSelection()
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean
    Dim i As Long, marks As Long
    Dim x As Double, y As Double, w As Double, h As Double
    Dim first As Boolean
    Dim s As String

    Set sr = ActiveSelectionRange
    ac2fSPResetCounts
    On Error GoTo Fail
    oldOpt = Application.Optimization
    Application.Optimization = True
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If
    ActiveDocument.BeginCommandGroup ac2fTitle("start points")

    first = True
    For i = 1 To sr.Count
        sr(i).GetBoundingBox x, y, w, h
        ac2fSPGrow first, x, y, w, h
    Next i
    For i = 1 To sr.Count
        ac2fSPTop sr(i)
    Next i
    If m_mark Then marks = ac2fSPMarkRange(sr)

    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh

    s = "Paths looked at   : " & m_paths & vbCrLf
    s = s & "Start moved       : " & m_moved & vbCrLf
    If m_added > 0 Then s = s & "Nodes added       : " & m_added & "  (exact points; shapes unchanged)" & vbCrLf
    If m_reversed > 0 Then s = s & "Direction turned  : " & m_reversed & vbCrLf
    If m_converted > 0 Then s = s & "Converted to curves: " & m_converted & " object(s)" & vbCrLf
    If m_text > 0 Then s = s & "Text left alone   : " & m_text & " object(s)" & vbCrLf
    If m_noReverse > 0 Then
        s = s & "Direction NOT changed on " & m_noReverse & " path(s): this CorelDRAW" & vbCrLf & _
                "does not let a macro reverse a path." & vbCrLf
    End If
    If m_failed > 0 Then s = s & "Could not move    : " & m_failed & " path(s)" & vbCrLf
    If marks > 0 Then s = s & "Markers drawn     : " & marks & vbCrLf
    s = s & vbCrLf & "Position " & m_pos & " (" & ac2fSPPosText(m_pos) & "), " & ac2fSPRefText(m_ref) & _
            ", " & IIf(m_exact, "exact point", "nearest node") & vbCrLf & _
            "Direction: " & ac2fSPDirText(m_dir) & vbCrLf & vbCrLf & "Ctrl+Z undoes it."
    ac2fInfo s, CAPTION_
    Exit Sub

Fail:
    s = Err.Description
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    ac2fWarn "The operation failed: " & s, CAPTION_
End Sub

Private Sub ac2fSPResetCounts()
    m_paths = 0: m_moved = 0: m_added = 0: m_reversed = 0
    m_failed = 0: m_noReverse = 0: m_converted = 0: m_text = 0
End Sub

Private Sub ac2fSPGrow(ByRef first As Boolean, ByVal x As Double, ByVal y As Double, _
                       ByVal w As Double, ByVal h As Double)
    If first Then
        m_selX0 = x: m_selY0 = y: m_selX1 = x + w: m_selY1 = y + h
        first = False
    Else
        If x < m_selX0 Then m_selX0 = x
        If y < m_selY0 Then m_selY0 = y
        If x + w > m_selX1 Then m_selX1 = x + w
        If y + h > m_selY1 Then m_selY1 = y + h
    End If
End Sub

' A top level object: its box is the "object" box for everything in it.
Private Sub ac2fSPTop(ByVal s As Shape)
    Dim x As Double, y As Double, w As Double, h As Double
    s.GetBoundingBox x, y, w, h
    m_objX0 = x: m_objY0 = y: m_objX1 = x + w: m_objY1 = y + h
    ac2fSPShape s
End Sub

Private Sub ac2fSPShape(ByVal s As Shape)
    Dim i As Long
    Dim cv As Curve

    On Error GoTo Skip
    Select Case s.Type
        Case cdrGroupShape
            For i = 1 To s.Shapes.Count
                ac2fSPShape s.Shapes(i)
            Next i
            Exit Sub
        Case cdrTextShape
            m_text = m_text + 1
            Exit Sub
        Case cdrRectangleShape, cdrEllipseShape, cdrPolygonShape
            s.ConvertToCurves
            m_converted = m_converted + 1
        Case cdrCurveShape
        Case Else
            Exit Sub
    End Select
    Set cv = s.Curve
    If cv Is Nothing Then Exit Sub
    ac2fSPCurve cv
Skip:
End Sub

'=====================================================================
' ONE CURVE
'=====================================================================

Private Sub ac2fSPCurve(ByVal cv As Curve)
    Dim i As Long, nSub As Long, dep As Long, j As Long
    Dim area As Double
    Dim wantCW As Boolean
    Dim sp As SubPath
    Dim ax As Double, ay As Double

    nSub = cv.SubPaths.Count
    If nSub = 0 Then Exit Sub

    ' outer or hole is decided once, on the curve as it is
    ac2fSPSampleAll cv
    For i = 1 To nSub
        m_paths = m_paths + 1
        Set sp = cv.SubPaths(i)
        If m_closed(i) Then
            If m_dir > 0 Then
                dep = 0
                For j = 1 To m_ns
                    If j <> i And m_closed(j) Then
                        If ac2fSPInside(m_x(m_s0(i)), m_y(m_s0(i)), j) Then dep = dep + 1
                    End If
                Next j
                area = ac2fSPArea(i)
                Select Case m_dir
                    Case 1: wantCW = ((dep Mod 2) = 0)
                    Case 2: wantCW = ((dep Mod 2) = 1)
                    Case 3: wantCW = True
                    Case Else: wantCW = False
                End Select
                ' positive area = counter-clockwise (y up)
                If (area < 0) <> wantCW Then ac2fSPReverse sp
            End If
            ac2fSPSampleOne cv.SubPaths(i), i
            ac2fSPAnchor i, ax, ay
            ac2fSPMoveStart cv, i, ax, ay
        ElseIf m_open Then
            ac2fSPAnchor i, ax, ay
            If ac2fSPDist(sp.EndNode.PositionX, sp.EndNode.PositionY, ax, ay) < _
               ac2fSPDist(sp.StartNode.PositionX, sp.StartNode.PositionY, ax, ay) - 0.001 Then
                ac2fSPReverse sp
            End If
        End If
    Next i
End Sub

Private Sub ac2fSPReverse(ByVal sp As SubPath)
    On Error GoTo Fail
    sp.ReverseDirection
    m_reversed = m_reversed + 1
    Exit Sub
Fail:
    m_noReverse = m_noReverse + 1
End Sub

' The point of the reference box for the chosen position.
Private Sub ac2fSPAnchor(ByVal i As Long, ByRef ax As Double, ByRef ay As Double)
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim k As Long, col As Long, rw As Long

    Select Case m_ref
        Case 2
            x0 = m_objX0: y0 = m_objY0: x1 = m_objX1: y1 = m_objY1
        Case 3
            x0 = m_selX0: y0 = m_selY0: x1 = m_selX1: y1 = m_selY1
        Case Else
            x0 = m_x(m_s0(i)): x1 = x0: y0 = m_y(m_s0(i)): y1 = y0
            For k = m_s0(i) To m_s0(i) + m_sc(i) - 1
                If m_x(k) < x0 Then x0 = m_x(k)
                If m_x(k) > x1 Then x1 = m_x(k)
                If m_y(k) < y0 Then y0 = m_y(k)
                If m_y(k) > y1 Then y1 = m_y(k)
            Next k
    End Select
    col = (m_pos - 1) Mod 3
    rw = (m_pos - 1) \ 3
    ax = x0 + col * (x1 - x0) / 2#
    ay = y0 + rw * (y1 - y0) / 2#
End Sub

' Makes the point of subpath i nearest (ax, ay) its start node.
Private Sub ac2fSPMoveStart(ByVal cv As Curve, ByVal i As Long, ByVal ax As Double, ByVal ay As Double)
    Dim k As Long, best As Long, segK As Long
    Dim d As Double, bestD As Double, t As Double
    Dim tx As Double, ty As Double
    Dim sp As SubPath
    Dim nd As Node

    best = -1
    For k = m_s0(i) To m_s0(i) + m_sc(i) - 1
        If m_exact Or m_t(k) = 0 Then
            d = ac2fSPDist(m_x(k), m_y(k), ax, ay)
            If best < 0 Or d < bestD - 0.000001 Then
                best = k
                bestD = d
            End If
        End If
    Next k
    If best < 0 Then Exit Sub
    tx = m_x(best): ty = m_y(best)
    segK = m_seg(best): t = m_t(best)

    On Error GoTo Fail
    Set sp = cv.SubPaths(i)
    If t > 0 Then
        ' between nodes: refine along the segment, then add a node there
        If ac2fSPStraight(sp.Segments(segK)) Then
            ac2fSPRefineLine sp.Segments(segK), ax, ay, t, tx, ty
        Else
            ac2fSPRefine sp.Segments(segK), ax, ay, t, tx, ty
        End If
        If ac2fSPDist(tx, ty, sp.Segments(segK).StartNode.PositionX, sp.Segments(segK).StartNode.PositionY) > 0.01 And _
           ac2fSPDist(tx, ty, sp.Segments(segK).EndNode.PositionX, sp.Segments(segK).EndNode.PositionY) > 0.01 Then
            sp.Segments(segK).AddNodeAt t, cdrRelativeSegmentOffset
            m_added = m_added + 1
            Set sp = cv.SubPaths(i)
        End If
    End If

    Set nd = ac2fSPNodeAt(sp, tx, ty)
    If nd Is Nothing Then GoTo Fail
    tx = nd.PositionX: ty = nd.PositionY
    If ac2fSPDist(sp.StartNode.PositionX, sp.StartNode.PositionY, tx, ty) <= 0.01 Then Exit Sub

    nd.BreakApart
    Set sp = cv.SubPaths(i)
    sp.StartNode.JoinWith sp.EndNode
    Set sp = cv.SubPaths(i)
    If Not sp.Closed Then sp.Closed = True
    If ac2fSPDist(sp.StartNode.PositionX, sp.StartNode.PositionY, tx, ty) <= 0.05 Then
        m_moved = m_moved + 1
    Else
        m_failed = m_failed + 1
    End If
    Exit Sub
Fail:
    m_failed = m_failed + 1
End Sub

' The node at (x, y); if none sits within 0.05 mm (an added node can land
' a hair off where it was asked for), the nearest one.
Private Function ac2fSPNodeAt(ByVal sp As SubPath, ByVal x As Double, ByVal y As Double) As Node
    Dim j As Long
    Dim d As Double, bestD As Double
    Dim nd As Node
    bestD = -1
    For j = 1 To sp.Nodes.Count
        Set nd = sp.Nodes(j)
        d = ac2fSPDist(nd.PositionX, nd.PositionY, x, y)
        If bestD < 0 Or d < bestD Then
            bestD = d
            Set ac2fSPNodeAt = nd
        End If
    Next j
End Function

Private Function ac2fSPStraight(ByVal sg As Segment) As Boolean
    Dim ln As Double, chord As Double
    ln = sg.Length
    chord = ac2fSPDist(sg.StartNode.PositionX, sg.StartNode.PositionY, _
                       sg.EndNode.PositionX, sg.EndNode.PositionY)
    ac2fSPStraight = (ln <= chord + 0.0005 * ln + 0.001)
End Function

' Nearest point of a straight segment, exactly.
Private Sub ac2fSPRefineLine(ByVal sg As Segment, ByVal ax As Double, ByVal ay As Double, _
                             ByRef t As Double, ByRef tx As Double, ByRef ty As Double)
    Dim x0 As Double, y0 As Double, dx As Double, dy As Double, l2 As Double
    x0 = sg.StartNode.PositionX: y0 = sg.StartNode.PositionY
    dx = sg.EndNode.PositionX - x0: dy = sg.EndNode.PositionY - y0
    l2 = dx * dx + dy * dy
    If l2 <= 0 Then Exit Sub
    t = ((ax - x0) * dx + (ay - y0) * dy) / l2
    If t < 0 Then t = 0
    If t > 1 Then t = 1
    tx = x0 + t * dx
    ty = y0 + t * dy
End Sub

' Golden section on t around the best sample: the nearest point of the
' segment to (ax, ay).
Private Sub ac2fSPRefine(ByVal sg As Segment, ByVal ax As Double, ByVal ay As Double, _
                         ByRef t As Double, ByRef tx As Double, ByRef ty As Double)
    Dim lo As Double, hi As Double, a As Double, b As Double
    Dim fa As Double, fb As Double, xa As Double, ya As Double, xb As Double, yb As Double
    Dim k As Long
    Const GR As Double = 0.618033988749895

    If m_noSample Then Exit Sub
    On Error GoTo Done
    lo = t - 0.1: hi = t + 0.1
    If lo < 0 Then lo = 0
    If hi > 1 Then hi = 1
    a = hi - GR * (hi - lo): b = lo + GR * (hi - lo)
    sg.GetPointPositionAt xa, ya, a, cdrRelativeSegmentOffset
    sg.GetPointPositionAt xb, yb, b, cdrRelativeSegmentOffset
    fa = ac2fSPDist(xa, ya, ax, ay): fb = ac2fSPDist(xb, yb, ax, ay)
    For k = 1 To 24
        If fa < fb Then
            hi = b: b = a: fb = fa: xb = xa: yb = ya
            a = hi - GR * (hi - lo)
            sg.GetPointPositionAt xa, ya, a, cdrRelativeSegmentOffset
            fa = ac2fSPDist(xa, ya, ax, ay)
        Else
            lo = a: a = b: fa = fb: xa = xb: ya = yb
            b = lo + GR * (hi - lo)
            sg.GetPointPositionAt xb, yb, b, cdrRelativeSegmentOffset
            fb = ac2fSPDist(xb, yb, ax, ay)
        End If
    Next k
    If fa < fb Then
        t = a: tx = xa: ty = ya
    Else
        t = b: tx = xb: ty = yb
    End If
Done:
End Sub

'=====================================================================
' SAMPLING
'=====================================================================

Private Sub ac2fSPSampleAll(ByVal cv As Curve)
    Dim i As Long, nSub As Long
    nSub = cv.SubPaths.Count
    m_n = 0
    m_ns = nSub
    ReDim m_s0(0 To nSub): ReDim m_sc(0 To nSub): ReDim m_closed(0 To nSub)
    ReDim m_x(0 To 1023): ReDim m_y(0 To 1023)
    ReDim m_seg(0 To 1023): ReDim m_t(0 To 1023)
    For i = 1 To nSub
        ac2fSPSampleOne cv.SubPaths(i), i
    Next i
End Sub

' Samples one subpath into the shared arrays (appended; the range of
' subpath i is replaced when it is sampled again).
Private Sub ac2fSPSampleOne(ByVal sp As SubPath, ByVal i As Long)
    Dim k As Long, j As Long, nSeg As Long, steps As Long
    Dim sg As Segment
    Dim ln As Double, chord As Double, x As Double, y As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double

    m_closed(i) = sp.Closed
    m_s0(i) = m_n
    nSeg = sp.Segments.Count
    For k = 1 To nSeg
        Set sg = sp.Segments(k)
        ax = sg.StartNode.PositionX: ay = sg.StartNode.PositionY
        bx = sg.EndNode.PositionX: by = sg.EndNode.PositionY
        ac2fSPAdd ax, ay, k, 0
        ln = sg.Length
        chord = ac2fSPDist(ax, ay, bx, by)
        steps = 1
        If ln > chord + 0.0005 * ln + 0.001 Then steps = 24
        If chord > 50 Then steps = steps * 2
        If steps < 2 And m_exact Then steps = CLng(chord / 5#) + 1
        If steps > 64 Then steps = 64
        For j = 1 To steps - 1
            If ln <= chord + 0.0005 * ln + 0.001 Then
                ' a straight segment needs no help from CorelDRAW
                ac2fSPAdd ax + (bx - ax) * j / steps, ay + (by - ay) * j / steps, k, j / steps
            ElseIf ac2fSPPoint(sg, j / steps, x, y) Then
                ac2fSPAdd x, y, k, j / steps
            End If
        Next j
        If k = nSeg And Not m_closed(i) Then ac2fSPAdd bx, by, k, 1
    Next k
    m_sc(i) = m_n - m_s0(i)
End Sub

' A point at t along a curved segment. False when this CorelDRAW cannot
' say; the nodes alone are then used for curved segments.
Private Function ac2fSPPoint(ByVal sg As Segment, ByVal t As Double, _
                             ByRef x As Double, ByRef y As Double) As Boolean
    If Not m_noSample Then
        On Error GoTo NoSample
        sg.GetPointPositionAt x, y, t, cdrRelativeSegmentOffset
        ac2fSPPoint = True
        Exit Function
    End If
NoSample:
    m_noSample = True
End Function

Private Sub ac2fSPAdd(ByVal x As Double, ByVal y As Double, ByVal k As Long, ByVal t As Double)
    If m_n > UBound(m_x) Then
        ReDim Preserve m_x(0 To m_n * 2): ReDim Preserve m_y(0 To m_n * 2)
        ReDim Preserve m_seg(0 To m_n * 2): ReDim Preserve m_t(0 To m_n * 2)
    End If
    m_x(m_n) = x: m_y(m_n) = y: m_seg(m_n) = k: m_t(m_n) = t
    m_n = m_n + 1
End Sub

Private Function ac2fSPArea(ByVal i As Long) As Double
    Dim k As Long, j As Long, s As Long, n As Long
    Dim a As Double
    s = m_s0(i): n = m_sc(i)
    If n < 3 Then Exit Function
    j = s + n - 1
    For k = s To s + n - 1
        a = a + (m_x(j) * m_y(k) - m_x(k) * m_y(j))
        j = k
    Next k
    ac2fSPArea = a / 2#
End Function

Private Function ac2fSPInside(ByVal px As Double, ByVal py As Double, ByVal i As Long) As Boolean
    Dim k As Long, j As Long, s As Long, n As Long
    Dim inside As Boolean
    s = m_s0(i): n = m_sc(i)
    If n < 3 Then Exit Function
    j = s + n - 1
    For k = s To s + n - 1
        If (m_y(k) > py) <> (m_y(j) > py) Then
            If px < m_x(k) + (py - m_y(k)) * (m_x(j) - m_x(k)) / (m_y(j) - m_y(k)) Then inside = Not inside
        End If
        j = k
    Next k
    ac2fSPInside = inside
End Function

Private Function ac2fSPDist(ByVal x1 As Double, ByVal y1 As Double, _
                            ByVal x2 As Double, ByVal y2 As Double) As Double
    ac2fSPDist = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
End Function

'=====================================================================
' MARKERS
'=====================================================================

Private Function ac2fSPMarkRange(ByVal sr As ShapeRange) As Long
    Dim ly As Layer, old As Layer
    Dim i As Long, n As Long
    Dim r As Double
    Dim x As Double, y As Double, w As Double, h As Double

    Set ly = ac2fSPFindLayer(True)
    If ly Is Nothing Then Exit Function
    sr.GetBoundingBox x, y, w, h
    r = Sqr(w * w + h * h) * 0.008
    If r < 0.8 Then r = 0.8
    If r > 6 Then r = 6
    Set old = ActiveLayer
    For i = 1 To sr.Count
        n = n + ac2fSPMarkShape(sr(i), ly, r)
    Next i
    On Error Resume Next
    old.Activate
    ac2fSPMarkRange = n
End Function

Private Function ac2fSPMarkShape(ByVal s As Shape, ByVal ly As Layer, ByVal r As Double) As Long
    Dim i As Long, n As Long
    Dim cv As Curve
    Dim sp As SubPath
    Dim sx As Double, sy As Double, dx As Double, dy As Double, d As Double
    Dim m As Shape

    On Error GoTo Done
    If s.Type = cdrGroupShape Then
        For i = 1 To s.Shapes.Count
            n = n + ac2fSPMarkShape(s.Shapes(i), ly, r)
        Next i
        ac2fSPMarkShape = n
        Exit Function
    End If
    If s.Type <> cdrCurveShape Then Exit Function
    Set cv = s.Curve
    For i = 1 To cv.SubPaths.Count
        Set sp = cv.SubPaths(i)
        sx = sp.StartNode.PositionX: sy = sp.StartNode.PositionY
        ' direction: towards a point a little along the first segment
        dx = sp.Segments(1).EndNode.PositionX - sx
        dy = sp.Segments(1).EndNode.PositionY - sy
        If Not m_noSample Then
            On Error Resume Next
            sp.Segments(1).GetPointPositionAt dx, dy, 0.1, cdrRelativeSegmentOffset
            If Err.Number = 0 Then
                dx = dx - sx: dy = dy - sy
            Else
                m_noSample = True
                dx = sp.Segments(1).EndNode.PositionX - sx
                dy = sp.Segments(1).EndNode.PositionY - sy
            End If
            On Error GoTo Done
        End If
        d = Sqr(dx * dx + dy * dy)
        Set m = ly.CreateEllipse2(sx, sy, r, r)
        m.Fill.ApplyNoFill
        m.Outline.Width = 0.2
        m.Outline.Color.RGBAssign 230, 30, 30
        If d > 0.000001 Then
            Set m = ly.CreateLineSegment(sx, sy, sx + dx / d * r * 3#, sy + dy / d * r * 3#)
            m.Outline.Width = 0.2
            m.Outline.Color.RGBAssign 230, 30, 30
        End If
        n = n + 1
    Next i
Done:
    ac2fSPMarkShape = n
End Function

Private Function ac2fSPFindLayer(ByVal create As Boolean) As Layer
    Dim i As Long
    On Error Resume Next
    For i = 1 To ActivePage.Layers.Count
        If ActivePage.Layers(i).Name = MARK_LAYER Then
            Set ac2fSPFindLayer = ActivePage.Layers(i)
            Exit Function
        End If
    Next i
    If create Then Set ac2fSPFindLayer = ActivePage.CreateLayer(MARK_LAYER)
End Function

'=====================================================================
' SETTINGS AND THE QUESTION
'=====================================================================

Private Sub ac2fSPLoad()
    m_pos = ac2fGetLng(AC2F_K_SP_POS, AC2F_DEF_SP_POS)
    If m_pos < 1 Or m_pos > 9 Then m_pos = AC2F_DEF_SP_POS
    m_ref = ac2fGetLng(AC2F_K_SP_REF, AC2F_DEF_SP_REF)
    If m_ref < 1 Or m_ref > 3 Then m_ref = AC2F_DEF_SP_REF
    m_dir = ac2fGetLng(AC2F_K_SP_DIR, AC2F_DEF_SP_DIR)
    If m_dir < 0 Or m_dir > 4 Then m_dir = AC2F_DEF_SP_DIR
    m_exact = (ac2fGetLng(AC2F_K_SP_EXACT, AC2F_DEF_SP_EXACT) <> 0)
    m_open = (ac2fGetLng(AC2F_K_SP_OPEN, AC2F_DEF_SP_OPEN) <> 0)
    m_mark = (ac2fGetLng(AC2F_K_SP_MARK, AC2F_DEF_SP_MARK) <> 0)
    m_noSample = False
End Sub

Private Sub ac2fSPSave()
    ac2fSetLng AC2F_K_SP_POS, m_pos
    ac2fSetLng AC2F_K_SP_REF, m_ref
    ac2fSetLng AC2F_K_SP_DIR, m_dir
    ac2fSetLng AC2F_K_SP_EXACT, IIf(m_exact, 1, 0)
    ac2fSetLng AC2F_K_SP_MARK, IIf(m_mark, 1, 0)
End Sub

' "7", "8 s w", "3EXA" ... a digit for the position, letters for options.
Private Function ac2fSPParse(ByVal ans As String) As Boolean
    Dim i As Long
    Dim c As String
    Dim gotPos As Boolean

    ans = UCase$(ans)
    For i = 1 To Len(ans)
        c = Mid$(ans, i, 1)
        Select Case c
            Case "1" To "9"
                m_pos = CLng(c)
                gotPos = True
            Case "E": m_ref = 1
            Case "O": m_ref = 2
            Case "S": m_ref = 3
            Case "X": m_exact = True
            Case "N": m_exact = False
            Case "K": m_dir = 0
            Case "W": m_dir = 1
            Case "C": m_dir = 2
            Case "A": m_dir = 3
            Case "Z": m_dir = 4
            Case "M": m_mark = True
            Case "Q": m_mark = False
            Case " ", ",", ";", "-"
            Case Else
                Exit Function
        End Select
    Next i
    ac2fSPParse = gotPos
End Function

Private Function ac2fSPCode() As String
    Dim s As String
    s = CStr(m_pos) & " " & Mid$("EOS", m_ref, 1) & " " & IIf(m_exact, "X", "N") & " " & _
        Mid$("KWCAZ", m_dir + 1, 1) & " " & IIf(m_mark, "M", "Q")
    ac2fSPCode = s
End Function

Private Function ac2fSPPrompt() As String
    Dim s As String
    s = "Where should every path start?" & vbCrLf & vbCrLf
    s = s & "        7 ----- 8 ----- 9" & vbCrLf
    s = s & "        |               |" & vbCrLf
    s = s & "        4       5       6" & vbCrLf
    s = s & "        |               |" & vbCrLf
    s = s & "        1 ----- 2 ----- 3" & vbCrLf & vbCrLf
    s = s & "A number, then letters to change options:" & vbCrLf
    s = s & "  E each contour's box   O each object   S whole selection" & vbCrLf
    s = s & "  X exact point (adds a node)   N nearest existing node" & vbCrLf
    s = s & "  K keep direction   W outer clockwise, holes counter" & vbCrLf
    s = s & "  C outer counter-clockwise, holes clockwise" & vbCrLf
    s = s & "  A all clockwise    Z all counter-clockwise" & vbCrLf
    s = s & "  M draw markers after   Q no markers" & vbCrLf & vbCrLf
    s = s & "Example: 8 S W = top centre of the selection, outer clockwise."
    ac2fSPPrompt = s
End Function

Private Function ac2fSPPosText(ByVal p As Long) As String
    Select Case p
        Case 7: ac2fSPPosText = "top left"
        Case 8: ac2fSPPosText = "top centre"
        Case 9: ac2fSPPosText = "top right"
        Case 4: ac2fSPPosText = "left centre"
        Case 5: ac2fSPPosText = "centre"
        Case 6: ac2fSPPosText = "right centre"
        Case 1: ac2fSPPosText = "bottom left"
        Case 2: ac2fSPPosText = "bottom centre"
        Case Else: ac2fSPPosText = "bottom right"
    End Select
End Function

Private Function ac2fSPRefText(ByVal r As Long) As String
    Select Case r
        Case 2: ac2fSPRefText = "each object's box"
        Case 3: ac2fSPRefText = "the whole selection's box"
        Case Else: ac2fSPRefText = "each contour's own box"
    End Select
End Function

Private Function ac2fSPDirText(ByVal d As Long) As String
    Select Case d
        Case 1: ac2fSPDirText = "outer clockwise, holes counter-clockwise"
        Case 2: ac2fSPDirText = "outer counter-clockwise, holes clockwise"
        Case 3: ac2fSPDirText = "all clockwise"
        Case 4: ac2fSPDirText = "all counter-clockwise"
        Case Else: ac2fSPDirText = "kept as drawn"
    End Select
End Function
