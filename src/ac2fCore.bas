Attribute VB_Name = "ac2fCore"
'=====================================================================
'  ac2f pack  --  ac2fCore
'
'  Shared constants, settings storage, helper functions and the
'  geometry measuring engine.
'
'  This module is not run directly; ac2fLength, ac2fLedModule,
'  ac2fBoxLetter and ac2fMenu all build on it.
'=====================================================================
Option Explicit

'---------------------------------------------------------------------
' Package identity
'---------------------------------------------------------------------
Public Const AC2F_NAME    As String = "ac2f pack"
Public Const AC2F_VERSION As String = "1.27.0"
Public Const AC2F_REG_APP As String = "ac2fPack"
Public Const AC2F_REG_SEC As String = "Ayarlar"
Public Const AC2F_REG_PROF As String = "Profiles"

'---------------------------------------------------------------------
' Defaults (editable through ac2fLedSettings)
'---------------------------------------------------------------------
Public Const AC2F_DEF_SPACING    As Double = 100#  ' module spacing (mm)
Public Const AC2F_DEF_LEDS       As Long = 3       ' LEDs per module
Public Const AC2F_DEF_MODULE_W   As Double = 0.72  ' module power (W)
Public Const AC2F_DEF_PSU_W      As Double = 60#   ' power supply rating (W)
Public Const AC2F_DEF_SAFETY     As Double = 20#   ' safety margin (%)
Public Const AC2F_DEF_MINPERPATH As Long = 1       ' minimum modules per path
Public Const AC2F_DEF_METHOD     As Long = 1       ' 1 = perimeter, 2 = centreline
Public Const AC2F_DEF_FACTOR     As Double = 1#    ' correction factor

'---------------------------------------------------------------------
' Registry key names.
'
' These are internal identifiers that never reach the user interface.
' They are deliberately left unchanged across releases so that saved
' settings survive an upgrade -- renaming them would silently reset
' every configured value back to the defaults.
'---------------------------------------------------------------------
Public Const AC2F_K_SPACING    As String = "ModulAraligiMM"
Public Const AC2F_K_LEDS       As String = "ModulBasinaLed"
Public Const AC2F_K_MODULE_W   As String = "ModulGucuW"
Public Const AC2F_K_PSU_W      As String = "GucKaynagiW"
Public Const AC2F_K_SAFETY     As String = "GuvenlikPayiYuzde"
Public Const AC2F_K_MINPERPATH As String = "AltYolBasinaEnAzModul"
Public Const AC2F_K_METHOD     As String = "HesapYontemi"
Public Const AC2F_K_FACTOR     As String = "DuzeltmeKatsayisi"

'---------------------------------------------------------------------
' Data structures
'---------------------------------------------------------------------

' Summary of a single measured object
Public Type ac2fItem
    Name     As String
    LengthMM As Double
    SubPaths As Long
End Type

' Full measurement result for a selection
Public Type ac2fResult
    Ok          As Boolean
    Message     As String
    TotalMM     As Double
    ShapeCount  As Long
    SubCount    As Long
    OpenCount   As Long
    ItemCount   As Long
    Items()     As ac2fItem
    SubLenMM()  As Double
    SubClosed() As Boolean
End Type

' Number of temporary duplicates created during a measurement.
Private m_TempShapes As Long

'---------------------------------------------------------------------
' Temporary overrides.
'
' A run can override individual settings in memory without writing them
' to the registry, which is what "use a profile but change one value
' just this once" needs. Every read goes through ac2fGetNum, so an
' override is honoured by all modules. Overrides are cleared at the
' start of each run that does not explicitly set them.
'---------------------------------------------------------------------
Private m_ovrN As Long
Private m_ovrKey() As String
Private m_ovrVal() As Double

'---------------------------------------------------------------------
' Curve sampling state (see CURVE SAMPLING below). Declarations have to
' stand above the first procedure of the module.
'---------------------------------------------------------------------
Public ac2fSgX() As Double          ' sampled points, mm
Public ac2fSgY() As Double
Public ac2fSgN As Long              ' how many: the start included, the end left out
Public ac2fSgDev As Double          ' how far the segment strays from its own chord, mm

Private m_sgMode As Long            ' 0 untested, 1 sampling works, -1 unavailable
Private m_sgSeg As Segment
Private m_sgTol As Double
Private m_sgBad As Boolean
Private m_sgAx As Double, m_sgAy As Double
Private m_sgBx As Double, m_sgBy As Double
Private m_sgLen As Double

Private Const SG_MAX_DEPTH As Long = 12
Private Const SG_MAX_POINTS As Long = 60000

'=====================================================================
' SETTINGS  (Windows registry: VB and VBA Program Settings)
'=====================================================================

Public Function ac2fGetStr(ByVal key As String, ByVal defValue As String) As String
    Dim v As String
    On Error Resume Next
    v = GetSetting(AC2F_REG_APP, AC2F_REG_SEC, key, defValue)
    If Len(Trim$(v)) = 0 Then v = defValue
    ac2fGetStr = v
End Function

Public Sub ac2fSetStr(ByVal key As String, ByVal value As String)
    On Error Resume Next
    SaveSetting AC2F_REG_APP, AC2F_REG_SEC, key, value
End Sub

' Numbers are always stored with a dot decimal separator so that the
' stored value does not depend on the machine's locale.
'
' A temporary override, when present, wins over the stored value.
Public Function ac2fGetNum(ByVal key As String, ByVal defValue As Double) As Double
    Dim s As String
    Dim v As Double

    If ac2fGetOverride(key, v) Then
        ac2fGetNum = v
        Exit Function
    End If

    s = ac2fGetStr(key, "")
    If Len(s) = 0 Then
        ac2fGetNum = defValue
    Else
        ac2fGetNum = Val(s)
    End If
End Function

'---------------------------------------------------------------------
' Temporary override layer
'---------------------------------------------------------------------

Public Sub ac2fSetOverride(ByVal key As String, ByVal value As Double)
    Dim i As Long

    For i = 0 To m_ovrN - 1
        If StrComp(m_ovrKey(i), key, vbTextCompare) = 0 Then
            m_ovrVal(i) = value
            Exit Sub
        End If
    Next i

    If m_ovrN = 0 Then
        ReDim m_ovrKey(0 To 15)
        ReDim m_ovrVal(0 To 15)
    ElseIf m_ovrN > UBound(m_ovrKey) Then
        ReDim Preserve m_ovrKey(0 To (UBound(m_ovrKey) + 1) * 2 - 1)
        ReDim Preserve m_ovrVal(0 To UBound(m_ovrKey))
    End If

    m_ovrKey(m_ovrN) = key
    m_ovrVal(m_ovrN) = value
    m_ovrN = m_ovrN + 1
End Sub

Public Function ac2fGetOverride(ByVal key As String, ByRef value As Double) As Boolean
    Dim i As Long
    For i = 0 To m_ovrN - 1
        If StrComp(m_ovrKey(i), key, vbTextCompare) = 0 Then
            value = m_ovrVal(i)
            ac2fGetOverride = True
            Exit Function
        End If
    Next i
End Function

Public Sub ac2fClearOverrides()
    m_ovrN = 0
End Sub

Public Function ac2fOverrideCount() As Long
    ac2fOverrideCount = m_ovrN
End Function

Public Sub ac2fSetNum(ByVal key As String, ByVal value As Double)
    ac2fSetStr key, ac2fNumStr(value)
End Sub

Public Function ac2fGetLng(ByVal key As String, ByVal defValue As Long) As Long
    ac2fGetLng = CLng(ac2fGetNum(key, CDbl(defValue)))
End Function

Public Sub ac2fSetLng(ByVal key As String, ByVal value As Long)
    ac2fSetStr key, ac2fNumStr(CDbl(value))
End Sub

' Restores every setting to its factory value.
Public Sub ac2fResetSettings()
    On Error Resume Next
    DeleteSetting AC2F_REG_APP, AC2F_REG_SEC
End Sub

'=====================================================================
' HELPERS
'=====================================================================

' Converts a number to locale independent text. Str$ drops the leading
' zero, so ".72" becomes "0.72" -- used both for storage and for the
' default value shown in an InputBox.
Public Function ac2fNumStr(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(v))
    If Left$(s, 1) = "." Then
        s = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        s = "-0" & Mid$(s, 2)
    End If
    ac2fNumStr = s
End Function

' Parses text typed by the user. Both "12,5" and "12.5" are accepted.
' Thousands separators must not be used.
Public Function ac2fParseNum(ByVal s As String, ByVal defValue As Double) As Double
    Dim t As String, sep As String
    t = Trim$(s)
    If Len(t) = 0 Then
        ac2fParseNum = defValue
        Exit Function
    End If
    sep = Mid$(CStr(1.5), 2, 1)          ' local decimal separator
    t = Replace$(t, ".", sep)
    t = Replace$(t, ",", sep)
    If Not IsNumeric(t) Then
        ac2fParseNum = defValue
        Exit Function
    End If
    ac2fParseNum = CDbl(t)
End Function

' Rounds up. Positive values only.
Public Function ac2fCeil(ByVal v As Double) As Double
    If v - Int(v) > 0.000001 Then
        ac2fCeil = Int(v) + 1#
    Else
        ac2fCeil = Int(v)
    End If
End Function

Public Function ac2fFmt(ByVal v As Double, Optional ByVal dec As Long = 2) As String
    If dec > 0 Then
        ac2fFmt = Format$(v, "#,##0." & String$(dec, "0"))
    Else
        ac2fFmt = Format$(v, "#,##0")
    End If
End Function

'---------------------------------------------------------------------
' Geometry helpers shared by the drawing modules.
'
' ac2fBoxLetter and ac2fCenterline still carry their own private copies,
' left alone so as not to disturb modules that are already validated.
' New code uses these; consolidating the older two is a follow-up.
'---------------------------------------------------------------------

' Solves c / L = 2 * sin(t/2) / t for t by bisection. The right hand side
' is strictly decreasing on (0, 2*pi), so bisection always converges.
' Given a segment's chord and arc length this recovers its turn angle,
' and R = L / t its radius, without reading bezier control points.
Public Function ac2fSolveTheta(ByVal chord As Double, ByVal arc As Double) As Double
    Dim r As Double, lo As Double, hi As Double, mp As Double, f As Double
    Dim i As Long
    Const PI2 As Double = 6.28318530717959

    If arc <= 0.000000001 Then Exit Function
    r = chord / arc
    If r >= 0.999999 Then Exit Function              ' straight
    If r <= 0# Then
        ac2fSolveTheta = PI2 - 0.000001
        Exit Function
    End If
    lo = 0.000001: hi = PI2 - 0.000001
    For i = 1 To 60
        mp = (lo + hi) / 2#
        f = 2# * Sin(mp / 2#) / mp
        If f > r Then lo = mp Else hi = mp
    Next i
    ac2fSolveTheta = (lo + hi) / 2#
End Function

'=====================================================================
' CURVE SAMPLING
'
' Turns one curved segment into points that lie ON the real curve.
'
' CorelDRAW is asked for points along the segment itself
' (Segment.GetPointPositionAt), so any bezier comes out right: an S
' that turns both ways, a loop, a cusp, handles of very different
' length. Nothing is guessed from the segment's length and chord.
'
' The polyline is refined until the curve stays within tol of it:
' every span is probed at its quarter, middle and three-quarter
' points, and a span whose probes stray more than half the tolerance
' from its chord is split in two. Every segment starts with at least
' four spans, and with about one span per 8 mm, so a feature smaller
' than a span cannot hide between probes.
'
' Two independent checks guard the answer. Each probe must lie within
' reach of both ends (|AP| + |PB| <= length), which no point of a real
' curve can break. And the finished polyline must be as long as the
' length CorelDRAW reports for the segment; a loop or wiggle the
' probes skipped would leave it short, and the segment is then read
' again, four times as finely.
'
' If CorelDRAW does not answer sensibly the mode turns to -1 and the
' caller cuts the curve with ac2fSegSplit instead, which needs
' nothing but Segment.AddNodeAt and Segment.Length.
'=====================================================================
' Forget what was learned about the CorelDRAW in front of us, so each
' run tests it again.
Public Sub ac2fSegReset()
    m_sgMode = 0
End Sub

Public Function ac2fSegMode() As Long
    ac2fSegMode = m_sgMode
End Function

' A curve of length L between two points A and B can never leave the
' ellipse whose foci are A and B and whose string is L. Its farthest
' reach from the chord is therefore sqrt(L^2 - c^2) / 2. This is a
' proven upper bound, not an estimate.
Public Function ac2fSegBulge(ByVal arcL As Double, ByVal chord As Double) As Double
    Dim d As Double
    d = arcL * arcL - chord * chord
    If d > 0# Then ac2fSegBulge = Sqr(d) / 2#
End Function

' Distance from point (px, py) to the finite line a-b.
Public Function ac2fPtSegDist(ByVal px As Double, ByVal py As Double, _
                              ByVal ax As Double, ByVal ay As Double, _
                              ByVal bx As Double, ByVal by As Double) As Double
    Dim dx As Double, dy As Double, l2 As Double, t As Double
    dx = bx - ax: dy = by - ay
    l2 = dx * dx + dy * dy
    If l2 <= 0.000000000001 Then
        t = 0#
    Else
        t = ((px - ax) * dx + (py - ay) * dy) / l2
        If t < 0# Then t = 0#
        If t > 1# Then t = 1#
    End If
    dx = px - (ax + t * dx): dy = py - (ay + t * dy)
    ac2fPtSegDist = Sqr(dx * dx + dy * dy)
End Function

' Points along sg. Returns how many (ac2fSgN), the start point first and
' the end point left to the next segment, or -1 when CorelDRAW cannot
' give points along a segment (the mode is then -1 until the next reset).
Public Function ac2fSegmentPoints(ByVal sg As Segment, ByVal tol As Double) As Long
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim chord As Double, ln As Double
    Dim n0 As Long, attempt As Long
    Dim polyLen As Double, k As Long
    Dim px As Double, py As Double

    ac2fSgN = 0: ac2fSgDev = 0#
    If tol < 0.0005 Then tol = 0.0005
    If ac2fSgMax() = 0 Then ac2fSgGrow

    ax = sg.StartNode.PositionX: ay = sg.StartNode.PositionY
    bx = sg.EndNode.PositionX: by = sg.EndNode.PositionY
    chord = Sqr((bx - ax) * (bx - ax) + (by - ay) * (by - ay))
    ln = 0#
    On Error Resume Next
    ln = sg.Length
    On Error GoTo 0

    ac2fSgAdd ax, ay
    ac2fSegmentPoints = 1
    If ac2fSegBulge(ln, chord) <= tol * 0.25 Then
        ac2fSgDev = ac2fSegBulge(ln, chord)
        Exit Function
    End If
    If m_sgMode < 0 Then
        ac2fSegmentPoints = -1
        Exit Function
    End If

    Set m_sgSeg = sg
    m_sgTol = tol
    m_sgAx = ax: m_sgAy = ay: m_sgBx = bx: m_sgBy = by
    m_sgLen = ln

    If m_sgMode = 0 Then
        If Not ac2fSgEndsAgree() Then
            m_sgMode = -1
            ac2fSgN = 0
            ac2fSegmentPoints = -1
            Exit Function
        End If
        m_sgMode = 1
    End If

    n0 = CLng(ac2fCeil(ln / 8#))
    If n0 < 4 Then n0 = 4
    If n0 > 128 Then n0 = 128

    For attempt = 0 To 2
        m_sgBad = False
        ac2fSgN = 0
        ac2fSgAdd ax, ay
        ac2fSgRun n0
        If m_sgBad Then
            m_sgMode = -1
            ac2fSgN = 0
            ac2fSegmentPoints = -1
            Exit Function
        End If
        polyLen = 0#
        For k = 1 To ac2fSgN - 1
            px = ac2fSgX(k) - ac2fSgX(k - 1): py = ac2fSgY(k) - ac2fSgY(k - 1)
            polyLen = polyLen + Sqr(px * px + py * py)
        Next k
        px = bx - ac2fSgX(ac2fSgN - 1): py = by - ac2fSgY(ac2fSgN - 1)
        polyLen = polyLen + Sqr(px * px + py * py)
        If polyLen >= ln - (ln * 0.001 + 0.01) Then Exit For
        n0 = n0 * 4
        If n0 > 4096 Then Exit For
    Next attempt

    ' how far the segment really strays from its chord, as measured
    For k = 1 To ac2fSgN - 1
        px = ac2fPtSegDist(ac2fSgX(k), ac2fSgY(k), ax, ay, bx, by)
        If px > ac2fSgDev Then ac2fSgDev = px
    Next k
    ac2fSegmentPoints = ac2fSgN
End Function

' The segment must hand back its own two nodes at offsets 0 and 1.
Private Function ac2fSgEndsAgree() As Boolean
    Dim px As Double, py As Double
    On Error GoTo Fail
    m_sgSeg.GetPointPositionAt px, py, 0#, cdrRelativeSegmentOffset
    If Abs(px - m_sgAx) > 0.05 Or Abs(py - m_sgAy) > 0.05 Then Exit Function
    m_sgSeg.GetPointPositionAt px, py, 1#, cdrRelativeSegmentOffset
    If Abs(px - m_sgBx) > 0.05 Or Abs(py - m_sgBy) > 0.05 Then Exit Function
    ac2fSgEndsAgree = True
    Exit Function
Fail:
End Function

Private Sub ac2fSgRun(ByVal n0 As Long)
    Dim i As Long
    Dim t0 As Double, t1 As Double, tm As Double
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim xm As Double, ym As Double

    t0 = 0#: x0 = m_sgAx: y0 = m_sgAy
    For i = 1 To n0
        If m_sgBad Then Exit Sub
        t1 = i / n0
        If i = n0 Then
            x1 = m_sgBx: y1 = m_sgBy
        Else
            ac2fSgProbe t1, x1, y1
        End If
        tm = (t0 + t1) / 2#
        ac2fSgProbe tm, xm, ym
        ac2fSgSpan t0, x0, y0, t1, x1, y1, tm, xm, ym, 0
        If i < n0 Then ac2fSgAdd x1, y1
        t0 = t1: x0 = x1: y0 = y1
    Next i
End Sub

' One span of the polyline, from offset t0 to t1, whose middle point is
' already known. Adds the points that belong strictly between the two
' ends. Quarter and three-quarter probes decide whether the chord will do.
Private Sub ac2fSgSpan(ByVal t0 As Double, ByVal x0 As Double, ByVal y0 As Double, _
                       ByVal t1 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                       ByVal tm As Double, ByVal xm As Double, ByVal ym As Double, _
                       ByVal depth As Long)
    Dim ta As Double, xa As Double, ya As Double
    Dim tb As Double, xb As Double, yb As Double
    Dim d As Double, e As Double

    If m_sgBad Then Exit Sub
    ta = (t0 + tm) / 2#: tb = (tm + t1) / 2#
    ac2fSgProbe ta, xa, ya
    ac2fSgProbe tb, xb, yb
    If m_sgBad Then Exit Sub

    d = ac2fPtSegDist(xm, ym, x0, y0, x1, y1)
    e = ac2fPtSegDist(xa, ya, x0, y0, x1, y1)
    If e > d Then d = e
    e = ac2fPtSegDist(xb, yb, x0, y0, x1, y1)
    If e > d Then d = e

    If d <= m_sgTol * 0.5 Or depth >= SG_MAX_DEPTH Or ac2fSgN >= SG_MAX_POINTS Then Exit Sub

    ac2fSgSpan t0, x0, y0, tm, xm, ym, ta, xa, ya, depth + 1
    ac2fSgAdd xm, ym
    ac2fSgSpan tm, xm, ym, t1, x1, y1, tb, xb, yb, depth + 1
End Sub

' One point of the real curve. A point that is out of reach of either
' end cannot belong to a curve of this length: CorelDRAW has answered
' with nonsense, so the reading is flagged and thrown away.
Private Sub ac2fSgProbe(ByVal t As Double, ByRef x As Double, ByRef y As Double)
    Dim reach As Double
    On Error GoTo Bad
    m_sgSeg.GetPointPositionAt x, y, t, cdrRelativeSegmentOffset
    reach = Sqr((x - m_sgAx) * (x - m_sgAx) + (y - m_sgAy) * (y - m_sgAy)) + _
            Sqr((x - m_sgBx) * (x - m_sgBx) + (y - m_sgBy) * (y - m_sgBy))
    If reach > m_sgLen * 1.002 + 0.05 Then GoTo Bad
    Exit Sub
Bad:
    m_sgBad = True
    x = m_sgAx: y = m_sgAy
End Sub

Private Function ac2fSgMax() As Long
    On Error Resume Next
    ac2fSgMax = UBound(ac2fSgX) + 1
End Function

Private Sub ac2fSgGrow()
    ReDim ac2fSgX(0 To 1023): ReDim ac2fSgY(0 To 1023)
End Sub

Private Sub ac2fSgAdd(ByVal x As Double, ByVal y As Double)
    If ac2fSgN > UBound(ac2fSgX) Then
        ReDim Preserve ac2fSgX(0 To (UBound(ac2fSgX) + 1) * 2 - 1)
        ReDim Preserve ac2fSgY(0 To UBound(ac2fSgX))
    End If
    ac2fSgX(ac2fSgN) = x
    ac2fSgY(ac2fSgN) = y
    ac2fSgN = ac2fSgN + 1
End Sub

' The fallback for a CorelDRAW that cannot give points along a segment.
' Cuts segment k of sp in halves (Segment.AddNodeAt) until each piece
' is within tolerance of its own chord, by the same proven bound.
' Works on a COPY of the drawing: it adds nodes. Returns how many
' segments now stand where segment k stood.
Public Function ac2fSegSplit(ByVal sp As SubPath, ByVal k As Long, _
                             ByVal tol As Double) As Long
    Dim cnt As Long, pos As Long, budget As Long
    Dim sg As Segment
    Dim ln As Double, chord As Double

    If tol < 0.0005 Then tol = 0.0005
    cnt = 1: pos = k: budget = 4096
    Do While pos < k + cnt
        Set sg = sp.Segments(pos)
        ln = 0#
        On Error Resume Next
        ln = sg.Length
        On Error GoTo 0
        chord = Sqr((sg.EndNode.PositionX - sg.StartNode.PositionX) ^ 2 + _
                    (sg.EndNode.PositionY - sg.StartNode.PositionY) ^ 2)
        If ac2fSegBulge(ln, chord) > tol * 0.25 And budget > 0 Then
            sg.AddNodeAt 0.5, cdrRelativeSegmentOffset
            cnt = cnt + 1
            budget = budget - 1
        Else
            pos = pos + 1
        End If
    Loop
    ac2fSegSplit = cnt
End Function

Public Function ac2fAtan2(ByVal y As Double, ByVal x As Double) As Double
    Const PI_ As Double = 3.14159265358979
    If x > 0 Then
        ac2fAtan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then ac2fAtan2 = Atn(y / x) + PI_ Else ac2fAtan2 = Atn(y / x) - PI_
    Else
        If y > 0 Then
            ac2fAtan2 = PI_ / 2#
        ElseIf y < 0 Then
            ac2fAtan2 = -PI_ / 2#
        Else
            ac2fAtan2 = 0
        End If
    End If
End Function

Public Function ac2fWrapAngle(ByVal a As Double) As Double
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
    ac2fWrapAngle = v
End Function

' Shows a length in mm / cm / m on one line.
Public Function ac2fFmtLength(ByVal mm As Double) As String
    ac2fFmtLength = ac2fFmt(mm) & " mm   |   " & _
                    ac2fFmt(mm / 10#) & " cm   |   " & _
                    ac2fFmt(mm / 1000#, 3) & " m"
End Function

' Pads or clips text to a fixed width, for the settings sheet.
Public Function ac2fPad(ByVal s As String, ByVal w As Long) As String
    If Len(s) >= w Then
        ac2fPad = Left$(s, w)
    Else
        ac2fPad = s & Space$(w - Len(s))
    End If
End Function

' Right aligns text in a fixed width.
Public Function ac2fRPad(ByVal s As String, ByVal w As Long) As String
    If Len(s) >= w Then
        ac2fRPad = Left$(s, w)
    Else
        ac2fRPad = Space$(w - Len(s)) & s
    End If
End Function

Public Function ac2fTitle(ByVal caption As String) As String
    ac2fTitle = AC2F_NAME & " " & AC2F_VERSION & " - " & caption
End Function

Public Sub ac2fInfo(ByVal msg As String, ByVal caption As String)
    MsgBox msg, vbInformation, ac2fTitle(caption)
End Sub

Public Sub ac2fWarn(ByVal msg As String, ByVal caption As String)
    MsgBox msg, vbExclamation, ac2fTitle(caption)
End Sub

'=====================================================================
' MEASURING ENGINE
'=====================================================================

' Measures the active selection. Returns Ok = False when there is no
' document or nothing is selected.
Public Function ac2fMeasureSelection() As ac2fResult
    Dim res As ac2fResult
    Dim sr As ShapeRange

    If ActiveDocument Is Nothing Then
        res.Message = "Open a document first."
        ac2fMeasureSelection = res
        Exit Function
    End If

    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        res.Message = "No objects selected to measure."
        ac2fMeasureSelection = res
        Exit Function
    End If
    If sr.Count = 0 Then
        res.Message = "No objects selected to measure."
        ac2fMeasureSelection = res
        Exit Function
    End If

    ac2fMeasureSelection = ac2fMeasureRange(sr)
End Function

' Measures the total path length (mm) of everything in a ShapeRange.
' Groups and PowerClip contents are walked recursively.
Public Function ac2fMeasureRange(ByVal sr As ShapeRange) As ac2fResult
    Dim res As ac2fResult
    Dim oldUnit As cdrUnit
    Dim oldOpt As Boolean
    Dim unitChanged As Boolean
    Dim groupOpen As Boolean
    Dim i As Long

    ReDim res.Items(0 To 63)
    ReDim res.SubLenMM(0 To 255)
    ReDim res.SubClosed(0 To 255)

    m_TempShapes = 0

    On Error GoTo Fail

    oldOpt = Application.Optimization
    Application.Optimization = True

    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    ActiveDocument.BeginCommandGroup ac2fTitle("measure")
    groupOpen = True

    For i = 1 To sr.Count
        ac2fCollect sr(i), res
    Next i

    ActiveDocument.EndCommandGroup
    groupOpen = False

    res.Ok = True
    res.Message = ""

Cleanup:
    On Error Resume Next
    If groupOpen Then ActiveDocument.EndCommandGroup
    ' Temporary duplicates have already been deleted; undoing the command
    ' group restores the document exactly. When no temporary object was
    ' created the Undo is skipped, otherwise it would undo the user's own
    ' previous action.
    If m_TempShapes > 0 Then ActiveDocument.Undo
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    m_TempShapes = 0
    ac2fMeasureRange = res
    Exit Function

Fail:
    res.Ok = False
    res.Message = "The measurement failed: " & Err.Description
    Resume Cleanup
End Function

'---------------------------------------------------------------------
' Adds one shape to the result (recursive).
'---------------------------------------------------------------------
Private Sub ac2fCollect(ByVal s As Shape, ByRef res As ac2fResult)
    Dim i As Long
    Dim cv As Curve
    Dim dup As Shape

    If s Is Nothing Then Exit Sub

    On Error GoTo Skip

    ' --- Objects that carry no path. DisplayCurve can hand back their
    '     frame, which would inflate the total, so they are dropped. ---
    Select Case s.Type
        Case cdrBitmapShape, cdrOLEObjectShape
            Exit Sub
    End Select

    ' --- Group: walk its contents ---
    If s.Type = cdrGroupShape Then
        For i = 1 To s.Shapes.Count
            ac2fCollect s.Shapes(i), res
        Next i
        Exit Sub
    End If

    ' --- Walk PowerClip contents as well ---
    ac2fCollectPowerClip s, res

    ' --- Obtain a curve representation ---
    Set cv = Nothing
    On Error Resume Next
    Set cv = s.DisplayCurve
    On Error GoTo Skip

    If Not cv Is Nothing Then
        ac2fAddCurve cv, ac2fShapeName(s), res
        Exit Sub
    End If

    ' No DisplayCurve: convert a throwaway duplicate to curves instead.
    Set dup = s.Duplicate(0, 0)
    m_TempShapes = m_TempShapes + 1

    On Error Resume Next
    dup.ConvertToCurves
    On Error GoTo SkipDup

    If dup.Type = cdrGroupShape Then
        ac2fCollect dup, res
        dup.Delete
        Exit Sub
    End If

    Set cv = Nothing
    On Error Resume Next
    Set cv = dup.Curve
    On Error GoTo SkipDup

    If cv Is Nothing Then GoTo SkipDup

    ac2fAddCurve cv, ac2fShapeName(s), res
    dup.Delete
    Exit Sub

SkipDup:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
    Exit Sub

Skip:
    ' An object that cannot be measured is skipped silently.
End Sub

Private Sub ac2fCollectPowerClip(ByVal s As Shape, ByRef res As ac2fResult)
    Dim pc As Object
    Dim i As Long

    On Error Resume Next
    Set pc = s.PowerClip
    If pc Is Nothing Then Exit Sub

    For i = 1 To pc.Shapes.Count
        ac2fCollect pc.Shapes(i), res
    Next i
End Sub

Private Function ac2fShapeName(ByVal s As Shape) As String
    Dim n As String
    On Error Resume Next
    n = s.Name
    If Len(Trim$(n)) = 0 Then n = "Object " & CStr(s.StaticID)
    ac2fShapeName = n
End Function

'---------------------------------------------------------------------
' Writes a curve's sub-paths into the result.
'---------------------------------------------------------------------
Private Sub ac2fAddCurve(ByVal cv As Curve, ByVal shapeName As String, ByRef res As ac2fResult)
    Dim i As Long
    Dim n As Long
    Dim spLen As Double
    Dim shpLen As Double
    Dim used As Long

    On Error Resume Next
    n = cv.SubPaths.Count
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    If n = 0 Then Exit Sub

    For i = 1 To n
        spLen = 0
        On Error Resume Next
        spLen = cv.SubPaths(i).Length
        On Error GoTo 0

        If spLen > 0 Then
            shpLen = shpLen + spLen
            used = used + 1

            If res.SubCount > UBound(res.SubLenMM) Then
                ReDim Preserve res.SubLenMM(0 To (UBound(res.SubLenMM) + 1) * 2 - 1)
                ReDim Preserve res.SubClosed(0 To UBound(res.SubLenMM))
            End If

            res.SubLenMM(res.SubCount) = spLen
            res.SubClosed(res.SubCount) = ac2fIsClosed(cv, i)
            If Not res.SubClosed(res.SubCount) Then res.OpenCount = res.OpenCount + 1
            res.SubCount = res.SubCount + 1
        End If
    Next i

    If shpLen <= 0 Then Exit Sub

    If res.ItemCount > UBound(res.Items) Then
        ReDim Preserve res.Items(0 To (UBound(res.Items) + 1) * 2 - 1)
    End If

    res.Items(res.ItemCount).Name = shapeName
    res.Items(res.ItemCount).LengthMM = shpLen
    res.Items(res.ItemCount).SubPaths = used
    res.ItemCount = res.ItemCount + 1

    res.ShapeCount = res.ShapeCount + 1
    res.TotalMM = res.TotalMM + shpLen
End Sub

Private Function ac2fIsClosed(ByVal cv As Curve, ByVal idx As Long) As Boolean
    Dim b As Boolean
    On Error Resume Next
    b = cv.SubPaths(idx).Closed
    ac2fIsClosed = b
End Function
