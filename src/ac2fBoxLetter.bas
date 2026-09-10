Attribute VB_Name = "ac2fBoxLetter"
'=====================================================================
'  ac2f pack  --  ac2fBoxLetter
'
'  Box letter (channel letter) return development and groove layout.
'
'  Takes the selected letter outlines, works out the developed length
'  from the material thickness and flexibility ratio, computes the
'  groove positions and draws flat strips ready to cut and groove.
'
'  Macros:
'    ac2fBoxLetterStrip    - calculates and draws the strips
'    ac2fBoxLetterReport   - calculates only, draws nothing
'    ac2fBoxLetterSettings - material and groove settings
'
'  GEOMETRY MODEL
'  Every segment is treated as a circular arc. The turn angle is solved
'  from the chord/arc ratio (c/L = 2*sin(t/2)/t) and the radius is
'  R = L/t. That removes any need to read bezier control points; node
'  positions and segment lengths are enough. On real cubic bezier arcs
'  the radius error is around 0.05%.
'
'  The development comes from offsetting the neutral axis. Total turning
'  of a simple closed curve is 2*pi, so the developed length is
'  P -/+ 2*pi*g (Steiner). That total is exact regardless of the per
'  segment turn signs; the signs only shift intermediate groove
'  positions, and the worst measured drift is 0.06 mm.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Box Letter Strip"

' Drawing safety limits
Private Const MAX_GROOVE   As Long = 5000     ' grooves per strip
Private Const MAX_STRIPS   As Long = 200      ' strips per job
Private Const STRAIGHT_EPS As Double = 0.0002 ' straight if (arc-chord)/arc < this

'---------------------------------------------------------------------
' Registry key names. Left unchanged across releases so that saved
' settings survive an upgrade (see ac2fCore).
'---------------------------------------------------------------------
Public Const AC2F_K_BL_THICK   As String = "KHKalinlikMM"
Public Const AC2F_K_BL_KFAC    As String = "KHKFaktoru"
Public Const AC2F_K_BL_REF     As String = "KHReferansYuzey"
Public Const AC2F_K_BL_HEIGHT  As String = "KHSeritYuksekligiMM"
Public Const AC2F_K_BL_DEPTH   As String = "KHDerzDerinlikOrani"
Public Const AC2F_K_BL_MOUTH   As String = "KHDerzAgziMM"
Public Const AC2F_K_BL_FLEX    As String = "KHEsneklikOrani"
Public Const AC2F_K_BL_TOL     As String = "KHYuzeyToleransiMM"
Public Const AC2F_K_BL_SMIN    As String = "KHDerzEnAzAralikMM"
Public Const AC2F_K_BL_SMAX    As String = "KHDerzEnCokAralikMM"
Public Const AC2F_K_BL_CORNER  As String = "KHKoseEsigiDerece"
Public Const AC2F_K_BL_JOINT   As String = "KHEkPayiMM"
Public Const AC2F_K_BL_COIL    As String = "KHRuloBoyuMM"
Public Const AC2F_K_BL_GAP     As String = "KHSeritAraligiMM"

'---------------------------------------------------------------------
' Defaults
'---------------------------------------------------------------------
Public Const AC2F_DEF_BL_THICK  As Double = 1#     ' material thickness (mm)
Public Const AC2F_DEF_BL_KFAC   As Double = 0.44   ' neutral axis K factor
Public Const AC2F_DEF_BL_REF    As Long = 1        ' 1=outer 2=inner 3=neutral
Public Const AC2F_DEF_BL_HEIGHT As Double = 80#    ' strip height = letter depth
Public Const AC2F_DEF_BL_DEPTH  As Double = 0.7    ' groove depth / thickness
Public Const AC2F_DEF_BL_MOUTH  As Double = 1.2    ' max groove mouth (mm)
Public Const AC2F_DEF_BL_FLEX   As Double = 1#     ' flexibility ratio
Public Const AC2F_DEF_BL_TOL    As Double = 0.15   ' surface (chord) tolerance mm
Public Const AC2F_DEF_BL_SMIN   As Double = 3#     ' minimum groove spacing
Public Const AC2F_DEF_BL_SMAX   As Double = 60#    ' maximum groove spacing
Public Const AC2F_DEF_BL_CORNER As Double = 5#     ' corner threshold (degrees)
Public Const AC2F_DEF_BL_JOINT  As Double = 20#    ' joint allowance (mm)
Public Const AC2F_DEF_BL_COIL   As Double = 3000#  ' coil length (0 = no split)
Public Const AC2F_DEF_BL_GAP    As Double = 10#    ' gap between drawn strips

'---------------------------------------------------------------------
' Analysed outline (module level arrays, to avoid copying large UDTs)
'---------------------------------------------------------------------
Private m_n    As Long        ' segment count
Private m_L()  As Double      ' arc length
Private m_Th() As Double      ' turn angle, unsigned (radians)
Private m_Sg() As Double      ' turn sign (+1 / -1)
Private m_R()  As Double      ' radius (0 = straight)
Private m_Dir() As Double     ' chord direction
Private m_T0() As Double      ' start tangent
Private m_T1() As Double      ' end tangent
Private m_Cor() As Double     ' corner turn after this segment
Private m_Tau  As Double      ' total signed turning

' Groove positions in developed coordinates (mm)
Private m_gN   As Long
Private m_gPos() As Double
Private m_gCor() As Boolean   ' True = corner groove

' Result for one strip
Private Type ac2fStrip
    Name     As String
    DevLen   As Double        ' developed length
    RawLen   As Double        ' raw outline length
    Grooves  As Long
    Corners  As Long
    MinR     As Double
    IsHole   As Boolean
End Type

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fBoxLetterStrip()
Attribute ac2fBoxLetterStrip.VB_Description = "ac2f pack: Develop the box letter return and draw the grooved strip"
    ac2fBLRun True
End Sub

Public Sub ac2fBoxLetterReport()
Attribute ac2fBoxLetterReport.VB_Description = "ac2f pack: Box letter strip report (draws nothing)"
    ac2fBLRun False
End Sub

'=====================================================================
' MAIN FLOW
'=====================================================================

Private Sub ac2fBLRun(ByVal draw As Boolean)
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean
    Dim strips() As ac2fStrip
    Dim nStrip As Long
    Dim x0 As Double, y0 As Double
    Dim bx As Double, by As Double, bw As Double, bh As Double
    Dim i As Long
    Dim warning As String

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "No outline selected to develop.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "No outline selected to develop.", CAPTION_
        Exit Sub
    End If

    If ac2fGetNum(AC2F_K_BL_THICK, AC2F_DEF_BL_THICK) <= 0 Then
        ac2fWarn "The material thickness must be greater than zero." & vbCrLf & _
                 "Correct it under Settings - Box letter.", CAPTION_
        Exit Sub
    End If

    ReDim strips(0 To MAX_STRIPS - 1)

    On Error GoTo Fail
    oldOpt = Application.Optimization
    Application.Optimization = True
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    sr.GetBoundingBox bx, by, bw, bh
    x0 = bx + bw + 20#
    y0 = by + bh

    ActiveDocument.BeginCommandGroup ac2fTitle("box letter strip")

    For i = 1 To sr.Count
        ac2fBLShape sr(i), strips, nStrip, draw, x0, y0
    Next i

    ActiveDocument.EndCommandGroup

    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    On Error GoTo 0

    If nStrip = 0 Then
        ac2fWarn "No closed outline found in the selection." & vbCrLf & vbCrLf & _
                 "A box letter strip can only be developed from a closed path.", CAPTION_
        Exit Sub
    End If

    If nStrip >= MAX_STRIPS Then
        warning = vbCrLf & "WARNING: the limit of " & MAX_STRIPS & _
                  " strips was reached, the remaining outlines were skipped." & vbCrLf
    End If

    ac2fInfo ac2fBLReport(strips, nStrip, draw) & warning, CAPTION_
    Exit Sub

Fail:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    ac2fWarn "The operation failed: " & Err.Description, CAPTION_
End Sub

'---------------------------------------------------------------------
' Handles one shape (recursive for groups).
'---------------------------------------------------------------------
Private Sub ac2fBLShape(ByVal s As Shape, ByRef strips() As ac2fStrip, _
                        ByRef nStrip As Long, ByVal draw As Boolean, _
                        ByRef x0 As Double, ByRef y0 As Double)
    Dim i As Long
    Dim cv As Curve
    Dim dup As Shape

    If s Is Nothing Then Exit Sub
    On Error GoTo Skip

    Select Case s.Type
        Case cdrBitmapShape, cdrOLEObjectShape
            Exit Sub
    End Select

    If s.Type = cdrGroupShape Then
        For i = 1 To s.Shapes.Count
            ac2fBLShape s.Shapes(i), strips, nStrip, draw, x0, y0
        Next i
        Exit Sub
    End If

    Set cv = Nothing
    On Error Resume Next
    Set cv = s.DisplayCurve
    On Error GoTo Skip

    If Not cv Is Nothing Then
        ac2fBLCurve cv, ac2fBLName(s), strips, nStrip, draw, x0, y0
        Exit Sub
    End If

    Set dup = s.Duplicate(0, 0)
    On Error Resume Next
    dup.ConvertToCurves
    On Error GoTo SkipDup
    If dup.Type = cdrGroupShape Then
        ac2fBLShape dup, strips, nStrip, draw, x0, y0
        dup.Delete
        Exit Sub
    End If
    Set cv = Nothing
    On Error Resume Next
    Set cv = dup.Curve
    On Error GoTo SkipDup
    If cv Is Nothing Then GoTo SkipDup
    ac2fBLCurve cv, ac2fBLName(s), strips, nStrip, draw, x0, y0
    dup.Delete
    Exit Sub

SkipDup:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
    Exit Sub
Skip:
End Sub

Private Function ac2fBLName(ByVal s As Shape) As String
    Dim n As String
    On Error Resume Next
    n = s.Name
    If Len(Trim$(n)) = 0 Then n = "Outline " & CStr(s.StaticID)
    ac2fBLName = n
End Function

'---------------------------------------------------------------------
' Turns every closed sub-path of a curve into a strip.
'---------------------------------------------------------------------
Private Sub ac2fBLCurve(ByVal cv As Curve, ByVal nm As String, _
                        ByRef strips() As ac2fStrip, ByRef nStrip As Long, _
                        ByVal draw As Boolean, ByRef x0 As Double, ByRef y0 As Double)
    Dim i As Long, nSub As Long
    Dim isHole() As Boolean

    On Error Resume Next
    nSub = cv.SubPaths.Count
    On Error GoTo 0
    If nSub = 0 Then Exit Sub

    ' Counter detection: a sub-path whose bounding box sits inside another
    ' one is a hole. Reliable and fast on letter outlines.
    ReDim isHole(0 To nSub - 1)
    ac2fBLMarkHoles cv, nSub, isHole

    For i = 1 To nSub
        If nStrip >= MAX_STRIPS Then Exit Sub
        Call ac2fBLSubPath(cv.SubPaths(i), nm & " #" & i, isHole(i - 1), _
                           strips, nStrip, draw, x0, y0)
    Next i
End Sub

' Marks holes by comparing sub-path bounding boxes.
Private Sub ac2fBLMarkHoles(ByVal cv As Curve, ByVal nSub As Long, _
                            ByRef isHole() As Boolean)
    Dim i As Long, j As Long
    Dim x1() As Double, y1() As Double, x2() As Double, y2() As Double

    ReDim x1(0 To nSub - 1): ReDim y1(0 To nSub - 1)
    ReDim x2(0 To nSub - 1): ReDim y2(0 To nSub - 1)

    For i = 0 To nSub - 1
        If Not ac2fBLSubBox(cv.SubPaths(i + 1), x1(i), y1(i), x2(i), y2(i)) Then
            x1(i) = 0: y1(i) = 0: x2(i) = 0: y2(i) = 0
        End If
    Next i

    For i = 0 To nSub - 1
        For j = 0 To nSub - 1
            If i <> j Then
                If x1(i) >= x1(j) And x2(i) <= x2(j) And _
                   y1(i) >= y1(j) And y2(i) <= y2(j) And _
                   (x2(j) - x1(j)) > (x2(i) - x1(i)) Then
                    isHole(i) = Not isHole(i)   ' role flips at every nesting level
                End If
            End If
        Next j
    Next i
End Sub

Private Function ac2fBLSubBox(ByVal sp As SubPath, ByRef x1 As Double, ByRef y1 As Double, _
                              ByRef x2 As Double, ByRef y2 As Double) As Boolean
    Dim k As Long, nn As Long
    Dim px As Double, py As Double

    On Error GoTo Fail
    nn = sp.Nodes.Count
    If nn = 0 Then Exit Function

    For k = 1 To nn
        px = sp.Nodes(k).PositionX
        py = sp.Nodes(k).PositionY
        If k = 1 Then
            x1 = px: x2 = px: y1 = py: y2 = py
        Else
            If px < x1 Then x1 = px
            If px > x2 Then x2 = px
            If py < y1 Then y1 = py
            If py > y2 Then y2 = py
        End If
    Next k
    ac2fBLSubBox = True
    Exit Function
Fail:
End Function

'=====================================================================
' GEOMETRY ANALYSIS
'=====================================================================

' Solves c/L = 2*sin(t/2)/t for t by bisection.
' The right hand side is strictly decreasing on (0, 2*pi).
Private Function ac2fBLTheta(ByVal chord As Double, ByVal arc As Double) As Double
    Dim r As Double, lo As Double, hi As Double, mid As Double, f As Double
    Dim i As Long
    Const PI2 As Double = 6.28318530717959

    If arc <= 0.000000001 Then Exit Function
    r = chord / arc
    If r >= 0.999999 Then Exit Function              ' straight
    If r <= 0# Then
        ac2fBLTheta = PI2 - 0.000001
        Exit Function
    End If

    lo = 0.000001: hi = PI2 - 0.000001
    For i = 1 To 60
        mid = (lo + hi) / 2#
        f = 2# * Sin(mid / 2#) / mid
        If f > r Then lo = mid Else hi = mid
    Next i
    ac2fBLTheta = (lo + hi) / 2#
End Function

Private Function ac2fBLWrap(ByVal a As Double) As Double
    Const PI2 As Double = 6.28318530717959
    Const PI_ As Double = 3.14159265358979
    Dim v As Double
    v = a
    Do While v > PI_
        v = v - PI2
    Loop
    Do While v <= -PI_
        v = v + PI2
    Loop
    ac2fBLWrap = v
End Function

' Analyses a sub-path: segment arcs, turns, radii and corners.
Private Function ac2fBLAnalyse(ByVal sp As SubPath) As Boolean
    Dim i As Long, nx As Long, pv As Long
    Dim ax As Double, ay As Double, bx2 As Double, by2 As Double
    Dim cx As Double, cy As Double, chord As Double
    Dim a As Double, b As Double, h As Double
    Dim sg As Segment

    On Error GoTo Fail
    m_n = sp.Segments.Count
    If m_n < 1 Then Exit Function

    ReDim m_L(0 To m_n - 1): ReDim m_Th(0 To m_n - 1)
    ReDim m_Sg(0 To m_n - 1): ReDim m_R(0 To m_n - 1)
    ReDim m_Dir(0 To m_n - 1): ReDim m_T0(0 To m_n - 1)
    ReDim m_T1(0 To m_n - 1): ReDim m_Cor(0 To m_n - 1)

    For i = 0 To m_n - 1
        Set sg = sp.Segments(i + 1)
        m_L(i) = sg.Length
        ax = sg.StartNode.PositionX: ay = sg.StartNode.PositionY
        bx2 = sg.EndNode.PositionX: by2 = sg.EndNode.PositionY
        cx = bx2 - ax: cy = by2 - ay
        chord = Sqr(cx * cx + cy * cy)
        m_Dir(i) = ac2fBLAtan2(cy, cx)

        ' Straightness is decided geometrically, so no dependency on the
        ' segment type constants.
        If m_L(i) <= 0 Then
            m_Th(i) = 0
        ElseIf (m_L(i) - chord) / m_L(i) < STRAIGHT_EPS Then
            m_Th(i) = 0
        Else
            m_Th(i) = ac2fBLTheta(chord, m_L(i))
        End If

        If m_Th(i) > 0.000000001 Then m_R(i) = m_L(i) / m_Th(i) Else m_R(i) = 0
    Next i

    ' Turn sign: the single node chord turns on either side of the segment.
    ' Spanning two segments instead would be ambiguous at +/-180 degrees and
    ' would flip the sign on clockwise (hole) outlines.
    For i = 0 To m_n - 1
        nx = i + 1: If nx > m_n - 1 Then nx = 0
        pv = i - 1: If pv < 0 Then pv = m_n - 1
        a = ac2fBLWrap(m_Dir(nx) - m_Dir(i))
        b = ac2fBLWrap(m_Dir(i) - m_Dir(pv))
        If (a + b) >= 0 Then m_Sg(i) = 1# Else m_Sg(i) = -1#
    Next i

    ' Tangents: chord direction -/+ half the turn.
    For i = 0 To m_n - 1
        h = m_Sg(i) * m_Th(i) / 2#
        m_T0(i) = m_Dir(i) - h
        m_T1(i) = m_Dir(i) + h
    Next i

    ' Corner turns and total turning.
    m_Tau = 0
    For i = 0 To m_n - 1
        nx = i + 1: If nx > m_n - 1 Then nx = 0
        m_Cor(i) = ac2fBLWrap(m_T0(nx) - m_T1(i))
        m_Tau = m_Tau + m_Sg(i) * m_Th(i) + m_Cor(i)
    Next i

    ac2fBLAnalyse = True
    Exit Function
Fail:
    m_n = 0
End Function

Private Function ac2fBLAtan2(ByVal y As Double, ByVal x As Double) As Double
    Const PI_ As Double = 3.14159265358979
    If x > 0 Then
        ac2fBLAtan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then ac2fBLAtan2 = Atn(y / x) + PI_ Else ac2fBLAtan2 = Atn(y / x) - PI_
    Else
        If y > 0 Then
            ac2fBLAtan2 = PI_ / 2#
        ElseIf y < 0 Then
            ac2fBLAtan2 = -PI_ / 2#
        Else
            ac2fBLAtan2 = 0
        End If
    End If
End Function

'=====================================================================
' GROOVE LAYOUT
'=====================================================================

' Produces the developed length and the groove positions.
Private Function ac2fBLGrooves(ByVal isHole As Boolean, ByRef devLen As Double, _
                               ByRef minR As Double, ByRef nCorner As Long) As Boolean
    Dim i As Long, k As Long, nGr As Long
    Dim t As Double, kf As Double, refFace As Long, g As Double, m As Double
    Dim depth As Double, mouth As Double, flex As Double, tol As Double
    Dim sMin As Double, sMax As Double, cornerLimit As Double
    Dim dPhiMax As Double
    Dim sig As Double, thHat As Double, corHat As Double
    Dim segDev As Double, corDev As Double
    Dim s As Double, s1 As Double, s2 As Double
    Dim acc As Double
    Const PI_ As Double = 3.14159265358979

    If m_n = 0 Then Exit Function

    t = ac2fGetNum(AC2F_K_BL_THICK, AC2F_DEF_BL_THICK)
    kf = ac2fGetNum(AC2F_K_BL_KFAC, AC2F_DEF_BL_KFAC)
    refFace = ac2fGetLng(AC2F_K_BL_REF, AC2F_DEF_BL_REF)
    depth = ac2fGetNum(AC2F_K_BL_DEPTH, AC2F_DEF_BL_DEPTH)
    mouth = ac2fGetNum(AC2F_K_BL_MOUTH, AC2F_DEF_BL_MOUTH)
    flex = ac2fGetNum(AC2F_K_BL_FLEX, AC2F_DEF_BL_FLEX)
    tol = ac2fGetNum(AC2F_K_BL_TOL, AC2F_DEF_BL_TOL)
    sMin = ac2fGetNum(AC2F_K_BL_SMIN, AC2F_DEF_BL_SMIN)
    sMax = ac2fGetNum(AC2F_K_BL_SMAX, AC2F_DEF_BL_SMAX)
    cornerLimit = ac2fGetNum(AC2F_K_BL_CORNER, AC2F_DEF_BL_CORNER) * PI_ / 180#

    If kf < 0 Then kf = 0
    If kf > 1 Then kf = 1
    If depth <= 0 Then depth = AC2F_DEF_BL_DEPTH
    If depth > 0.95 Then depth = 0.95
    If flex <= 0 Then flex = 1#
    If tol <= 0 Then tol = AC2F_DEF_BL_TOL
    If sMin <= 0 Then sMin = 0.5
    If sMax < sMin Then sMax = sMin

    ' Offset of the neutral axis from the vector, into the material
    Select Case refFace
        Case 2:    g = -kf * t          ' vector is the inner face
        Case 3:    g = 0#               ' vector is already the neutral axis
        Case Else: g = (1# - kf) * t    ' vector is the outer face (default)
    End Select
    If isHole Then m = -1# Else m = 1#

    ' Largest turn one groove can absorb:
    ' mouth width = groove depth * angle  ->  angle = mouth / depth
    If t * depth <= 0 Then Exit Function
    dPhiMax = (mouth / (t * depth)) * flex
    If dPhiMax <= 0.0001 Then dPhiMax = 0.0001

    sig = 1#
    If m_Tau < 0 Then sig = -1#

    ReDim m_gPos(0 To 255)
    ReDim m_gCor(0 To 255)
    m_gN = 0
    nCorner = 0
    minR = 0
    acc = 0

    For i = 0 To m_n - 1
        thHat = sig * m_Sg(i) * m_Th(i)          ' normalised turn
        segDev = m_L(i) - g * m * thHat

        If m_Th(i) > 0.000000001 And m_R(i) > 0 Then
            If minR = 0 Or m_R(i) < minR Then minR = m_R(i)

            ' Spacing: (a) groove mouth limit  (b) surface tolerance  (c) ceiling
            s1 = m_R(i) * dPhiMax
            s2 = Sqr(8# * m_R(i) * tol)
            s = s1
            If s2 < s Then s = s2
            If sMax < s Then s = sMax
            If s < sMin Then s = sMin

            nGr = CLng(ac2fCeil(m_L(i) / s))
            If nGr < 1 Then nGr = 1

            For k = 0 To nGr - 1
                ac2fBLAddGroove acc + (k + 0.5) * segDev / nGr, False
            Next k
        End If

        acc = acc + segDev

        ' Corner
        corHat = sig * m_Cor(i)
        corDev = -g * m * corHat
        If Abs(corHat) > cornerLimit Then
            nGr = CLng(ac2fCeil(Abs(corHat) / dPhiMax))
            If nGr < 1 Then nGr = 1
            For k = 0 To nGr - 1
                ac2fBLAddGroove acc + corDev / 2# + (k - (nGr - 1) / 2#) * sMin, True
            Next k
            nCorner = nCorner + nGr
        End If
        acc = acc + corDev
    Next i

    devLen = acc
    ac2fBLGrooves = (devLen > 0)
End Function

Private Sub ac2fBLAddGroove(ByVal pos As Double, ByVal isCorner As Boolean)
    If m_gN >= MAX_GROOVE Then Exit Sub
    If m_gN > UBound(m_gPos) Then
        ReDim Preserve m_gPos(0 To (UBound(m_gPos) + 1) * 2 - 1)
        ReDim Preserve m_gCor(0 To UBound(m_gPos))
    End If
    m_gPos(m_gN) = pos
    m_gCor(m_gN) = isCorner
    m_gN = m_gN + 1
End Sub

'=====================================================================
' STRIP GENERATION
'=====================================================================

Private Function ac2fBLSubPath(ByVal sp As SubPath, ByVal nm As String, _
                               ByVal isHole As Boolean, ByRef strips() As ac2fStrip, _
                               ByRef nStrip As Long, ByVal draw As Boolean, _
                               ByRef x0 As Double, ByRef y0 As Double) As Boolean
    Dim devLen As Double, minR As Double
    Dim nCorner As Long
    Dim rawLen As Double
    Dim closed As Boolean
    Dim h As Double, gap As Double

    On Error GoTo Fail

    closed = False
    On Error Resume Next
    closed = sp.Closed
    rawLen = sp.Length
    On Error GoTo Fail
    If Not closed Then Exit Function      ' an open path cannot become a return
    If rawLen <= 0 Then Exit Function

    If Not ac2fBLAnalyse(sp) Then Exit Function
    If Not ac2fBLGrooves(isHole, devLen, minR, nCorner) Then Exit Function

    strips(nStrip).Name = nm
    strips(nStrip).DevLen = devLen
    strips(nStrip).RawLen = rawLen
    strips(nStrip).Grooves = m_gN
    strips(nStrip).Corners = nCorner
    strips(nStrip).MinR = minR
    strips(nStrip).IsHole = isHole
    nStrip = nStrip + 1

    If draw Then
        h = ac2fGetNum(AC2F_K_BL_HEIGHT, AC2F_DEF_BL_HEIGHT)
        gap = ac2fGetNum(AC2F_K_BL_GAP, AC2F_DEF_BL_GAP)
        If h <= 0 Then h = AC2F_DEF_BL_HEIGHT
        If gap < 0 Then gap = 0
        ac2fBLDraw x0, y0, devLen, h, nm
        y0 = y0 - (h + gap + 6#)
    End If

    ac2fBLSubPath = True
    Exit Function
Fail:
End Function

' Draws the flat strip, the groove lines and the cut marks.
Private Sub ac2fBLDraw(ByVal x As Double, ByVal yTop As Double, _
                       ByVal L As Double, ByVal h As Double, ByVal nm As String)
    Dim i As Long
    Dim sh As Shape
    Dim yBot As Double
    Dim coil As Double, joint As Double
    Dim cut As Double

    yBot = yTop - h

    On Error Resume Next

    ' Strip outline
    Set sh = ActiveLayer.CreateRectangle(x, yTop, x + L, yBot)
    If Not sh Is Nothing Then
        sh.Fill.ApplyNoFill
        sh.Outline.Width = 0.2
        sh.Outline.Color.RGBAssign 0, 0, 0
    End If

    ' Groove lines
    For i = 0 To m_gN - 1
        If m_gPos(i) >= 0 And m_gPos(i) <= L Then
            Set sh = ActiveLayer.CreateLineSegment(x + m_gPos(i), yTop, _
                                                   x + m_gPos(i), yBot)
            If Not sh Is Nothing Then
                sh.Outline.Width = 0.1
                If m_gCor(i) Then
                    sh.Outline.Color.RGBAssign 230, 0, 120     ' corner groove
                Else
                    sh.Outline.Color.RGBAssign 0, 160, 220     ' curve groove
                End If
            End If
        End If
    Next i

    ' Cut / joint marks when the coil length is exceeded
    coil = ac2fGetNum(AC2F_K_BL_COIL, AC2F_DEF_BL_COIL)
    joint = ac2fGetNum(AC2F_K_BL_JOINT, AC2F_DEF_BL_JOINT)
    ' The joint allowance must be smaller than the coil length, otherwise
    ' the step would be zero or negative and the loop would never advance.
    If coil > 0 And L > coil And joint < coil Then
        cut = coil
        Do While cut < L
            Set sh = ActiveLayer.CreateLineSegment(x + cut, yTop + 4#, x + cut, yBot - 4#)
            If Not sh Is Nothing Then
                sh.Outline.Width = 0.4
                sh.Outline.Color.RGBAssign 255, 0, 0           ' cut here
            End If
            cut = cut + coil - joint
        Loop
    End If

    ' Label
    Set sh = ActiveLayer.CreateArtisticText(x, yTop + 2#, _
             nm & "  |  developed " & ac2fFmt(L) & " mm  |  height " & _
             ac2fFmt(h, 0) & " mm  |  " & m_gN & " grooves")
    If Not sh Is Nothing Then sh.Text.Story.Size = 8
End Sub

'=====================================================================
' REPORT
'=====================================================================

Private Function ac2fBLReport(ByRef strips() As ac2fStrip, ByVal n As Long, _
                              ByVal draw As Boolean) As String
    Dim s As String
    Dim i As Long
    Dim totDev As Double
    Dim totGr As Long, totCor As Long
    Dim smallestR As Double
    Dim coil As Double, joint As Double
    Dim pieces As Long
    Dim t As Double
    Dim refFace As Long

    For i = 0 To n - 1
        totDev = totDev + strips(i).DevLen
        totGr = totGr + strips(i).Grooves
        totCor = totCor + strips(i).Corners
        If strips(i).MinR > 0 Then
            If smallestR = 0 Or strips(i).MinR < smallestR Then smallestR = strips(i).MinR
        End If
    Next i

    t = ac2fGetNum(AC2F_K_BL_THICK, AC2F_DEF_BL_THICK)
    refFace = ac2fGetLng(AC2F_K_BL_REF, AC2F_DEF_BL_REF)
    coil = ac2fGetNum(AC2F_K_BL_COIL, AC2F_DEF_BL_COIL)
    joint = ac2fGetNum(AC2F_K_BL_JOINT, AC2F_DEF_BL_JOINT)

    s = "RESULT" & vbCrLf
    s = s & "   Strips              : " & n & vbCrLf
    s = s & "   Total developed     : " & ac2fFmtLength(totDev) & vbCrLf
    s = s & "   Total grooves       : " & totGr & "  (" & totCor & " corner)" & vbCrLf
    If smallestR > 0 Then
        s = s & "   Smallest radius     : " & ac2fFmt(smallestR) & " mm" & vbCrLf
    End If

    If coil > 0 And totDev > 0 And joint < coil Then
        pieces = CLng(ac2fCeil(totDev / (coil - joint)))
        s = s & "   Coil required       : ~" & pieces & " x " & _
                ac2fFmt(coil, 0) & " mm (" & ac2fFmt(joint, 0) & " mm joint)" & vbCrLf
    End If
    s = s & vbCrLf

    s = s & "STRIPS" & vbCrLf
    For i = 0 To n - 1
        If i >= 12 Then
            s = s & "   ... (first 12 of " & n & ")" & vbCrLf
            Exit For
        End If
        s = s & "   " & (i + 1) & ". " & strips(i).Name
        If strips(i).IsHole Then s = s & " [hole]"
        s = s & vbCrLf & "      developed " & ac2fFmt(strips(i).DevLen) & " mm" & _
                " (raw " & ac2fFmt(strips(i).RawLen) & ")" & _
                "  grooves " & strips(i).Grooves
        If strips(i).MinR > 0 Then s = s & "  min R " & ac2fFmt(strips(i).MinR, 1)
        s = s & vbCrLf
    Next i
    s = s & vbCrLf

    s = s & "MATERIAL" & vbCrLf
    s = s & "   Thickness           : " & ac2fFmt(t) & " mm" & vbCrLf
    s = s & "   Reference face      : " & ac2fBLRefName(refFace) & vbCrLf
    s = s & "   K factor            : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_KFAC, AC2F_DEF_BL_KFAC)) & vbCrLf
    s = s & "   Flexibility ratio   : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_FLEX, AC2F_DEF_BL_FLEX)) & vbCrLf
    s = s & "   Groove depth ratio  : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_DEPTH, AC2F_DEF_BL_DEPTH)) & vbCrLf

    If draw Then
        s = s & vbCrLf & "The strips were drawn on the page." & vbCrLf & _
                "   blue = curve groove   pink = corner groove   red = cut here"
    Else
        s = s & vbCrLf & "Nothing was drawn (report only)."
    End If

    ac2fBLReport = s
End Function

Public Function ac2fBLRefName(ByVal r As Long) As String
    Select Case r
        Case 2:    ac2fBLRefName = "Vector = inner face of the strip"
        Case 3:    ac2fBLRefName = "Vector = neutral axis"
        Case Else: ac2fBLRefName = "Vector = outer face of the strip"
    End Select
End Function

'=====================================================================
' SETTINGS
'=====================================================================

Public Sub ac2fBoxLetterSettings()
Attribute ac2fBoxLetterSettings.VB_Description = "ac2f pack: Box letter material and groove settings"
    Const C As String = "Settings - Box Letter"
    Dim answer As String
    Dim v As Double
    Dim n As Long

    ' 1) Material preset
    answer = InputBox( _
        "Choose a material preset:" & vbCrLf & vbCrLf & _
        "  0  -  Set manually (leave unchanged)" & vbCrLf & _
        "  1  -  Aluminium, thin (0.5-1.5 mm)" & vbCrLf & _
        "  2  -  Galvanised, thin (0.5-1.2 mm)" & vbCrLf & _
        "  3  -  Aluminium, thick (2-4 mm)" & vbCrLf & _
        "  4  -  Stainless steel" & vbCrLf & vbCrLf & _
        "Presets are starting points, not measured values." & vbCrLf & _
        "Calibrate the flexibility ratio on your own first job.", _
        ac2fTitle(C), "0")
    If StrPtr(answer) = 0 Then Exit Sub
    ac2fBLPreset CLng(ac2fParseNum(answer, 0))

    If Not ac2fBLAsk("Material thickness (mm)", C, AC2F_K_BL_THICK, _
                     AC2F_DEF_BL_THICK, 0.01, v) Then Exit Sub
    If Not ac2fBLAsk("Strip height (mm)" & vbCrLf & "The depth of the letter.", C, _
                     AC2F_K_BL_HEIGHT, AC2F_DEF_BL_HEIGHT, 1#, v) Then Exit Sub
    If Not ac2fBLAsk("Flexibility ratio" & vbCrLf & _
                     "Higher = fewer grooves. Calibrate on your own work.", C, _
                     AC2F_K_BL_FLEX, AC2F_DEF_BL_FLEX, 0.05, v) Then Exit Sub
    If Not ac2fBLAskLng("Reference face" & vbCrLf & _
                     "1 = Vector is the outer face of the strip" & vbCrLf & _
                     "2 = Vector is the inner face of the strip" & vbCrLf & _
                     "3 = Vector is the neutral axis", C, _
                     AC2F_K_BL_REF, AC2F_DEF_BL_REF, 1, 3, n) Then Exit Sub

    If MsgBox("Edit the advanced settings as well?" & vbCrLf & _
              "(K factor, groove mouth, tolerance, spacing limits, coil length)", _
              vbQuestion + vbYesNo, ac2fTitle(C)) = vbYes Then

        If Not ac2fBLAsk("K factor (neutral axis position, 0-1)", C, _
                         AC2F_K_BL_KFAC, AC2F_DEF_BL_KFAC, 0#, v) Then GoTo Done
        If Not ac2fBLAsk("Groove depth / thickness ratio (0-0.95)", C, _
                         AC2F_K_BL_DEPTH, AC2F_DEF_BL_DEPTH, 0.05, v) Then GoTo Done
        If Not ac2fBLAsk("Maximum groove mouth (mm)" & vbCrLf & _
                         "The widest mouth that leaves no mark once closed.", C, _
                         AC2F_K_BL_MOUTH, AC2F_DEF_BL_MOUTH, 0.05, v) Then GoTo Done
        If Not ac2fBLAsk("Surface tolerance (mm)" & vbCrLf & _
                         "How far the flat between grooves may sit off the curve.", C, _
                         AC2F_K_BL_TOL, AC2F_DEF_BL_TOL, 0.01, v) Then GoTo Done
        If Not ac2fBLAsk("Minimum groove spacing (mm)", C, _
                         AC2F_K_BL_SMIN, AC2F_DEF_BL_SMIN, 0.5, v) Then GoTo Done
        If Not ac2fBLAsk("Maximum groove spacing (mm)", C, _
                         AC2F_K_BL_SMAX, AC2F_DEF_BL_SMAX, 1#, v) Then GoTo Done
        If Not ac2fBLAsk("Corner threshold (degrees)" & vbCrLf & _
                         "A turn above this counts as a corner groove.", C, _
                         AC2F_K_BL_CORNER, AC2F_DEF_BL_CORNER, 0.1, v) Then GoTo Done
        If Not ac2fBLAsk("Coil length (mm)" & vbCrLf & "0 = do not split.", C, _
                         AC2F_K_BL_COIL, AC2F_DEF_BL_COIL, 0#, v) Then GoTo Done
        If Not ac2fBLAsk("Joint allowance (mm)", C, _
                         AC2F_K_BL_JOINT, AC2F_DEF_BL_JOINT, 0#, v) Then GoTo Done
        If Not ac2fBLAsk("Gap between drawn strips (mm)", C, _
                         AC2F_K_BL_GAP, AC2F_DEF_BL_GAP, 0#, v) Then GoTo Done
    End If

Done:
    ac2fInfo "Box letter settings saved." & vbCrLf & vbCrLf & ac2fBLSettingsSummary(), C
End Sub

' Presets: starting values for flexibility and groove depth.
' These are calibration starting points, not measured shop data.
Private Sub ac2fBLPreset(ByVal p As Long)
    Select Case p
        Case 1      ' aluminium, thin
            ac2fSetNum AC2F_K_BL_FLEX, 1.2
            ac2fSetNum AC2F_K_BL_DEPTH, 0.7
            ac2fSetNum AC2F_K_BL_KFAC, 0.44
        Case 2      ' galvanised, thin
            ac2fSetNum AC2F_K_BL_FLEX, 1#
            ac2fSetNum AC2F_K_BL_DEPTH, 0.65
            ac2fSetNum AC2F_K_BL_KFAC, 0.44
        Case 3      ' aluminium, thick
            ac2fSetNum AC2F_K_BL_FLEX, 0.8
            ac2fSetNum AC2F_K_BL_DEPTH, 0.75
            ac2fSetNum AC2F_K_BL_KFAC, 0.42
        Case 4      ' stainless steel
            ac2fSetNum AC2F_K_BL_FLEX, 0.7
            ac2fSetNum AC2F_K_BL_DEPTH, 0.6
            ac2fSetNum AC2F_K_BL_KFAC, 0.45
    End Select
End Sub

Public Function ac2fBLSettingsSummary() As String
    Dim s As String
    s = "   Thickness           : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_THICK, AC2F_DEF_BL_THICK)) & " mm" & vbCrLf
    s = s & "   Strip height        : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_HEIGHT, AC2F_DEF_BL_HEIGHT), 0) & " mm" & vbCrLf
    s = s & "   Flexibility ratio   : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_FLEX, AC2F_DEF_BL_FLEX)) & vbCrLf
    s = s & "   Reference face      : " & ac2fBLRefName(ac2fGetLng(AC2F_K_BL_REF, AC2F_DEF_BL_REF)) & vbCrLf
    s = s & "   K factor            : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_KFAC, AC2F_DEF_BL_KFAC)) & vbCrLf
    s = s & "   Groove depth ratio  : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_DEPTH, AC2F_DEF_BL_DEPTH)) & vbCrLf
    s = s & "   Max groove mouth    : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_MOUTH, AC2F_DEF_BL_MOUTH)) & " mm" & vbCrLf
    s = s & "   Surface tolerance   : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_TOL, AC2F_DEF_BL_TOL)) & " mm" & vbCrLf
    s = s & "   Groove spacing      : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_SMIN, AC2F_DEF_BL_SMIN), 1) & _
            " - " & ac2fFmt(ac2fGetNum(AC2F_K_BL_SMAX, AC2F_DEF_BL_SMAX), 1) & " mm" & vbCrLf
    s = s & "   Coil / joint        : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_COIL, AC2F_DEF_BL_COIL), 0) & _
            " / " & ac2fFmt(ac2fGetNum(AC2F_K_BL_JOINT, AC2F_DEF_BL_JOINT), 0) & " mm" & vbCrLf
    ac2fBLSettingsSummary = s
End Function

Private Function ac2fBLAsk(ByVal question As String, ByVal caption As String, _
                           ByVal key As String, ByVal defValue As Double, _
                           ByVal minVal As Double, ByRef result As Double) As Boolean
    Dim current As Double, answer As String, v As Double
    current = ac2fGetNum(key, defValue)
    answer = InputBox(question & vbCrLf & vbCrLf & "(Default: " & ac2fNumStr(defValue) & ")", _
                      ac2fTitle(caption), ac2fNumStr(current))
    If StrPtr(answer) = 0 Then Exit Function
    v = ac2fParseNum(answer, current)
    If v < minVal Then v = minVal
    ac2fSetNum key, v
    result = v
    ac2fBLAsk = True
End Function

Private Function ac2fBLAskLng(ByVal question As String, ByVal caption As String, _
                              ByVal key As String, ByVal defValue As Long, _
                              ByVal minVal As Long, ByVal maxVal As Long, _
                              ByRef result As Long) As Boolean
    Dim current As Long, answer As String, v As Long
    current = ac2fGetLng(key, defValue)
    answer = InputBox(question & vbCrLf & vbCrLf & "(Default: " & ac2fNumStr(CDbl(defValue)) & ")", _
                      ac2fTitle(caption), ac2fNumStr(CDbl(current)))
    If StrPtr(answer) = 0 Then Exit Function
    v = CLng(ac2fParseNum(answer, CDbl(current)))
    If v < minVal Then v = minVal
    If v > maxVal Then v = maxVal
    ac2fSetLng key, v
    result = v
    ac2fBLAskLng = True
End Function
