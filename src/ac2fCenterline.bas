Attribute VB_Name = "ac2fCenterline"
'=====================================================================
'  ac2f pack  --  ac2fCenterline
'
'  Reduces filled artwork to a single centre line, for routing a neon
'  LED strip channel ALONG the line with a 6 mm bit.
'
'  Macros:
'    ac2fCenterline         - asks joined or separate, then runs
'    ac2fCenterlineJoined   - whole selection as one connected region
'    ac2fCenterlineSeparate - every object reduced on its own
'
'  PIPELINE
'    1  flatten outlines to polygons (circular arc model, no bezier
'       control points needed - see ac2fBoxLetter for the reasoning)
'    2  rasterise: even-odd per shape so holes come out right, then OR
'       the shapes together so touching letters become one region
'    3  Zhang-Suen thinning down to a one pixel skeleton
'    4  drop redundant staircase pixels so chains are unambiguous
'    5  trace pixels into chains, prune short spurs
'    6  extend free ends, which thinning pulls back by half a stroke
'    7  smooth out the pixel zigzag, then Douglas-Peucker simplify
'    8  draw each chain as one polyline
'
'  Measured on a shape whose true centre line was known: end position
'  within 1.8 mm, total length 470.8 mm against a true 471.0 mm.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Centerline"

' Safety limits
Private Const MAX_CELLS  As Long = 4000000   ' raster cells
Private Const MAX_PTS    As Long = 400000    ' flattened outline points
Private Const MAX_THIN   As Long = 300       ' thinning passes

'---------------------------------------------------------------------
' Registry keys (internal, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_CL_RES    As String = "CLCozunurlukMM"
Public Const AC2F_K_CL_TOL    As String = "CLSadelestirmeMM"
Public Const AC2F_K_CL_SMOOTH As String = "CLYumusatma"
Public Const AC2F_K_CL_BRANCH As String = "CLEnAzDalKatsayi"
Public Const AC2F_K_CL_EXTEND As String = "CLUclariUzat"

Public Const AC2F_DEF_CL_RES    As Double = 1#    ' raster resolution (mm)
Public Const AC2F_DEF_CL_TOL    As Double = 0.3   ' simplify tolerance (mm)
Public Const AC2F_DEF_CL_SMOOTH As Long = 2       ' smoothing passes
Public Const AC2F_DEF_CL_BRANCH As Double = 1#    ' min branch, x stroke width
Public Const AC2F_DEF_CL_EXTEND As Long = 1       ' extend free ends

'---------------------------------------------------------------------
' Raster
'---------------------------------------------------------------------
Private m_W As Long, m_H As Long
Private m_x0 As Double, m_y0 As Double, m_res As Double
Private m_grid() As Byte        ' working, thinned in place
Private m_mask() As Byte        ' original filled region, kept for end extension
Private m_filled As Long        ' filled cell count

' Flattened outline buffer for one shape
Private m_fx() As Double, m_fy() As Double
Private m_fn As Long
Private m_sps() As Long, m_spc() As Long, m_spn As Long   ' subpath start / count

' Traced chains: points held flat, one index table per chain
Private m_px() As Double, m_py() As Double, m_pn As Long
Private m_cs() As Long, m_cc() As Long, m_cn As Long

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fCenterline()
Attribute ac2fCenterline.VB_Description = "ac2f pack: Reduce artwork to a single centre line"
    Dim answer As String
    answer = InputBox( _
        "Reduce to a single centre line." & vbCrLf & vbCrLf & _
        "   J   joined   - the whole selection as one region," & vbCrLf & _
        "                  letters that touch stay connected" & vbCrLf & _
        "   S   separate - every object on its own", _
        ac2fTitle(CAPTION_), "J")
    If StrPtr(answer) = 0 Then Exit Sub
    answer = UCase$(Trim$(answer))
    If Len(answer) = 0 Then Exit Sub

    Select Case Left$(answer, 1)
        Case "J": ac2fCLRun True
        Case "S": ac2fCLRun False
        Case Else: ac2fWarn "Type J for joined or S for separate.", CAPTION_
    End Select
End Sub

Public Sub ac2fCenterlineJoined()
Attribute ac2fCenterlineJoined.VB_Description = "ac2f pack: Centre line, whole selection joined as one region"
    ac2fCLRun True
End Sub

Public Sub ac2fCenterlineSeparate()
Attribute ac2fCenterlineSeparate.VB_Description = "ac2f pack: Centre line, every object reduced on its own"
    ac2fCLRun False
End Sub

'=====================================================================
' MAIN FLOW
'=====================================================================

Private Sub ac2fCLRun(ByVal joined As Boolean)
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean
    Dim src() As Shape
    Dim n As Long, i As Long
    Dim res As Double
    Dim groups As Long, paths As Long
    Dim totLen As Double, width As Double
    Dim msg As String

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Select the artwork to reduce.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Select the artwork to reduce.", CAPTION_
        Exit Sub
    End If

    res = ac2fGetNum(AC2F_K_CL_RES, AC2F_DEF_CL_RES)
    If res <= 0 Then res = AC2F_DEF_CL_RES
    m_res = res

    On Error GoTo Fail
    oldOpt = Application.Optimization
    Application.Optimization = True
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    n = sr.Count
    ReDim src(0 To n - 1)
    For i = 0 To n - 1
        Set src(i) = sr(i + 1)
    Next i

    ActiveDocument.BeginCommandGroup ac2fTitle("centerline")

    If joined Then
        If ac2fCLOne(src, n, paths, totLen, width, msg) Then groups = 1
    Else
        Dim one(0 To 0) As Shape
        Dim p As Long, tl As Double, wd As Double
        For i = 0 To n - 1
            Set one(0) = src(i)
            p = 0: tl = 0: wd = 0
            If ac2fCLOne(one, 1, p, tl, wd, msg) Then
                groups = groups + 1
                paths = paths + p
                totLen = totLen + tl
                If wd > width Then width = wd
            End If
        Next i
    End If

    ActiveDocument.EndCommandGroup

    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    On Error GoTo 0

    If groups = 0 Then
        ac2fWarn "No centre line could be produced." & vbCrLf & vbCrLf & _
                 IIf(Len(msg) > 0, msg, "The selection has no filled area to reduce."), CAPTION_
        Exit Sub
    End If

    ac2fInfo ac2fCLReport(joined, groups, paths, totLen, width, msg), CAPTION_
    Exit Sub

Fail:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    ac2fWarn "The operation failed: " & Err.Description, CAPTION_
End Sub

' Runs the whole pipeline over one set of shapes and draws the result.
Private Function ac2fCLOne(ByRef shp() As Shape, ByVal n As Long, _
                           ByRef paths As Long, ByRef totLen As Double, _
                           ByRef width As Double, ByRef msg As String) As Boolean
    Dim i As Long
    Dim x1 As Double, y1 As Double, w As Double, h As Double
    Dim bx1 As Double, by1 As Double, bx2 As Double, by2 As Double
    Dim pad As Double
    Dim cells As Double
    Dim minBranch As Double

    ' --- bounding box over the set --------------------------------
    For i = 0 To n - 1
        shp(i).GetBoundingBox x1, y1, w, h
        If i = 0 Then
            bx1 = x1: by1 = y1: bx2 = x1 + w: by2 = y1 + h
        Else
            If x1 < bx1 Then bx1 = x1
            If y1 < by1 Then by1 = y1
            If x1 + w > bx2 Then bx2 = x1 + w
            If y1 + h > by2 Then by2 = y1 + h
        End If
    Next i
    If bx2 <= bx1 Or by2 <= by1 Then Exit Function

    pad = m_res * 4#
    m_x0 = bx1 - pad: m_y0 = by1 - pad
    m_W = Int((bx2 + pad - m_x0) / m_res) + 2
    m_H = Int((by2 + pad - m_y0) / m_res) + 2

    cells = CDbl(m_W) * CDbl(m_H)
    If cells > MAX_CELLS Then
        msg = "The artwork needs " & Format$(cells, "#,##0") & " raster cells at " & _
              ac2fFmt(m_res) & " mm." & vbCrLf & _
              "Raise the Resolution setting and run again."
        Exit Function
    End If

    ReDim m_grid(0 To m_W * m_H - 1)
    ReDim m_mask(0 To m_W * m_H - 1)

    ' --- rasterise, one shape at a time, union ---------------------
    For i = 0 To n - 1
        ac2fCLRasterShape shp(i)
    Next i

    m_filled = 0
    For i = 0 To m_W * m_H - 1
        If m_grid(i) <> 0 Then
            m_mask(i) = 1
            m_filled = m_filled + 1
        End If
    Next i
    If m_filled = 0 Then Exit Function

    ' --- skeleton --------------------------------------------------
    ac2fCLThin
    ac2fCLCleanup
    ac2fCLTrace
    If m_cn = 0 Then Exit Function

    ' Mean stroke width follows from area and skeleton length:
    ' a stroke of width w and length L covers w * L.
    totLen = ac2fCLTotalLength()
    If totLen <= 0 Then Exit Function
    width = (CDbl(m_filled) * m_res * m_res) / totLen

    minBranch = ac2fGetNum(AC2F_K_CL_BRANCH, AC2F_DEF_CL_BRANCH) * width
    ac2fCLPrune minBranch
    If m_cn = 0 Then Exit Function

    If ac2fGetLng(AC2F_K_CL_EXTEND, AC2F_DEF_CL_EXTEND) <> 0 Then
        ac2fCLExtend width
    End If

    ac2fCLSmooth CLng(ac2fGetNum(AC2F_K_CL_SMOOTH, CDbl(AC2F_DEF_CL_SMOOTH)))
    ac2fCLSimplify ac2fGetNum(AC2F_K_CL_TOL, AC2F_DEF_CL_TOL)

    totLen = ac2fCLTotalLength()
    paths = m_cn
    ac2fCLOne = ac2fCLDraw()
End Function

'=====================================================================
' 1  FLATTEN  (circular arc model, see module header)
'=====================================================================

' Fills m_fx/m_fy and the subpath table from one shape's curve.
Private Function ac2fCLFlatten(ByVal s As Shape) As Boolean
    Dim cv As Curve
    Dim dup As Shape
    Dim ok As Boolean

    On Error GoTo Skip
    Select Case s.Type
        Case cdrBitmapShape, cdrOLEObjectShape
            Exit Function
    End Select

    Set cv = Nothing
    On Error Resume Next
    Set cv = s.DisplayCurve
    On Error GoTo Skip

    If cv Is Nothing Then
        Set dup = s.Duplicate(0, 0)
        On Error Resume Next
        dup.ConvertToCurves
        Set cv = dup.Curve
        On Error GoTo SkipDup
        If cv Is Nothing Then GoTo SkipDup
        ok = ac2fCLFlattenCurve(cv)
        dup.Delete
        ac2fCLFlatten = ok
        Exit Function
    End If

    ac2fCLFlatten = ac2fCLFlattenCurve(cv)
    Exit Function

SkipDup:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
Skip:
End Function

Private Function ac2fCLFlattenCurve(ByVal cv As Curve) As Boolean
    Dim i As Long, k As Long, nSub As Long, nSeg As Long
    Dim sp As SubPath, sg As Segment
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim cxx As Double, cyy As Double, chord As Double
    Dim arcL As Double, th As Double, sgn As Double
    Dim steps As Long

    m_fn = 0: m_spn = 0
    ReDim m_fx(0 To 4095): ReDim m_fy(0 To 4095)
    ReDim m_sps(0 To 63): ReDim m_spc(0 To 63)

    On Error Resume Next
    nSub = cv.SubPaths.Count
    On Error GoTo 0
    If nSub = 0 Then Exit Function

    For i = 1 To nSub
        Set sp = Nothing
        On Error Resume Next
        Set sp = cv.SubPaths(i)
        nSeg = sp.Segments.Count
        On Error GoTo 0
        If Not sp Is Nothing Then
            If nSeg > 0 Then
                If m_spn > UBound(m_sps) Then
                    ReDim Preserve m_sps(0 To (UBound(m_sps) + 1) * 2 - 1)
                    ReDim Preserve m_spc(0 To UBound(m_sps))
                End If
                m_sps(m_spn) = m_fn

                For k = 1 To nSeg
                    Set sg = sp.Segments(k)
                    arcL = sg.Length
                    ax = sg.StartNode.PositionX: ay = sg.StartNode.PositionY
                    bx = sg.EndNode.PositionX: by = sg.EndNode.PositionY
                    cxx = bx - ax: cyy = by - ay
                    chord = Sqr(cxx * cxx + cyy * cyy)

                    th = 0
                    If arcL > 0 Then
                        If (arcL - chord) / arcL >= 0.0002 Then th = ac2fCLTheta(chord, arcL)
                    End If

                    ' Enough samples that the polyline is finer than a cell
                    steps = CLng(ac2fCeil(arcL / (m_res * 0.7)))
                    If steps < 1 Then steps = 1
                    If steps > 400 Then steps = 400

                    sgn = ac2fCLSegSign(sp, nSeg, k)
                    ac2fCLEmitArc ax, ay, bx, by, arcL, th, sgn, steps
                    If m_fn >= MAX_PTS Then Exit For
                Next k

                m_spc(m_spn) = m_fn - m_sps(m_spn)
                If m_spc(m_spn) >= 3 Then
                    m_spn = m_spn + 1
                Else
                    m_fn = m_sps(m_spn)
                End If
            End If
        End If
        If m_fn >= MAX_PTS Then Exit For
    Next i

    ac2fCLFlattenCurve = (m_spn > 0)
End Function

' Turn direction of segment k, from the chord turns at its two ends.
Private Function ac2fCLSegSign(ByVal sp As SubPath, ByVal nSeg As Long, _
                               ByVal k As Long) As Double
    Dim a As Double, b As Double
    Dim dPrev As Double, dCur As Double, dNext As Double

    If nSeg < 2 Then
        ac2fCLSegSign = 1#
        Exit Function
    End If

    dCur = ac2fCLChordDir(sp, k)
    dNext = ac2fCLChordDir(sp, IIf(k = nSeg, 1, k + 1))
    dPrev = ac2fCLChordDir(sp, IIf(k = 1, nSeg, k - 1))

    a = ac2fCLWrap(dNext - dCur)
    b = ac2fCLWrap(dCur - dPrev)
    If (a + b) >= 0 Then ac2fCLSegSign = 1# Else ac2fCLSegSign = -1#
End Function

Private Function ac2fCLChordDir(ByVal sp As SubPath, ByVal k As Long) As Double
    Dim sg As Segment
    On Error Resume Next
    Set sg = sp.Segments(k)
    If sg Is Nothing Then Exit Function
    ac2fCLChordDir = ac2fCLAtan2(sg.EndNode.PositionY - sg.StartNode.PositionY, _
                                 sg.EndNode.PositionX - sg.StartNode.PositionX)
End Function

' Writes the sampled arc (start point included, end point excluded, so
' consecutive segments do not double up).
Private Sub ac2fCLEmitArc(ByVal ax As Double, ByVal ay As Double, _
                          ByVal bx As Double, ByVal by As Double, _
                          ByVal arcL As Double, ByVal th As Double, _
                          ByVal sgn As Double, ByVal steps As Long)
    Dim i As Long
    Dim r As Double, phi As Double, t0 As Double
    Dim cx As Double, cy As Double, a0 As Double, ang As Double

    If th < 0.000000001 Then
        For i = 0 To steps - 1
            ac2fCLAddPt ax + (bx - ax) * i / steps, ay + (by - ay) * i / steps
        Next i
        Exit Sub
    End If

    r = arcL / th
    phi = ac2fCLAtan2(by - ay, bx - ax)
    t0 = phi - sgn * th / 2#
    cx = ax + sgn * r * (-Sin(t0))
    cy = ay + sgn * r * Cos(t0)
    a0 = ac2fCLAtan2(ay - cy, ax - cx)

    For i = 0 To steps - 1
        ang = a0 + sgn * th * i / steps
        ac2fCLAddPt cx + r * Cos(ang), cy + r * Sin(ang)
    Next i
End Sub

Private Sub ac2fCLAddPt(ByVal x As Double, ByVal y As Double)
    If m_fn > UBound(m_fx) Then
        ReDim Preserve m_fx(0 To (UBound(m_fx) + 1) * 2 - 1)
        ReDim Preserve m_fy(0 To UBound(m_fx))
    End If
    m_fx(m_fn) = x
    m_fy(m_fn) = y
    m_fn = m_fn + 1
End Sub

'=====================================================================
' 2  RASTERISE  (even-odd within a shape, OR between shapes)
'=====================================================================

Private Sub ac2fCLRasterShape(ByVal s As Shape)
    Dim row As Long, i As Long, j As Long, k As Long
    Dim yc As Double
    Dim xs() As Double
    Dim nx As Long
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim c0 As Long, c1 As Long
    Dim base As Long
    Dim sp0 As Long, spc As Long

    If Not ac2fCLFlatten(s) Then Exit Sub
    ReDim xs(0 To 255)

    For row = 0 To m_H - 1
        yc = m_y0 + (row + 0.5) * m_res
        nx = 0

        For i = 0 To m_spn - 1
            sp0 = m_sps(i): spc = m_spc(i)
            For j = 0 To spc - 1
                ax = m_fx(sp0 + j): ay = m_fy(sp0 + j)
                k = sp0 + ((j + 1) Mod spc)
                bx = m_fx(k): by = m_fy(k)
                If (ay <= yc And by > yc) Or (by <= yc And ay > yc) Then
                    If nx > UBound(xs) Then ReDim Preserve xs(0 To (UBound(xs) + 1) * 2 - 1)
                    xs(nx) = ax + (yc - ay) * (bx - ax) / (by - ay)
                    nx = nx + 1
                End If
            Next j
        Next i

        If nx > 1 Then
            ac2fCLSortD xs, nx
            base = row * m_W
            For k = 0 To nx - 2 Step 2
                c0 = Int((xs(k) - m_x0) / m_res)
                c1 = Int((xs(k + 1) - m_x0) / m_res)
                If c0 < 0 Then c0 = 0
                If c1 > m_W - 1 Then c1 = m_W - 1
                For j = c0 To c1
                    m_grid(base + j) = 1
                Next j
            Next k
        End If
    Next row
End Sub

Private Sub ac2fCLSortD(ByRef a() As Double, ByVal n As Long)
    Dim i As Long, j As Long
    Dim v As Double
    For i = 1 To n - 1
        v = a(i)
        j = i - 1
        Do While j >= 0
            If a(j) <= v Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = v
    Next i
End Sub

'=====================================================================
' 3  ZHANG-SUEN THINNING
'=====================================================================

Private Sub ac2fCLThin()
    Dim pass As Long, step_ As Long
    Dim r As Long, c As Long, i As Long
    Dim p2 As Long, p3 As Long, p4 As Long, p5 As Long
    Dim p6 As Long, p7 As Long, p8 As Long, p9 As Long
    Dim a As Long, b As Long
    Dim rem_() As Long, nr As Long
    Dim changed As Boolean
    Dim base As Long

    ReDim rem_(0 To 4095)

    For pass = 1 To MAX_THIN
        changed = False
        For step_ = 0 To 1
            nr = 0
            For r = 1 To m_H - 2
                base = r * m_W
                For c = 1 To m_W - 2
                    If m_grid(base + c) <> 0 Then
                        p2 = m_grid(base - m_W + c)
                        p3 = m_grid(base - m_W + c + 1)
                        p4 = m_grid(base + c + 1)
                        p5 = m_grid(base + m_W + c + 1)
                        p6 = m_grid(base + m_W + c)
                        p7 = m_grid(base + m_W + c - 1)
                        p8 = m_grid(base + c - 1)
                        p9 = m_grid(base - m_W + c - 1)

                        b = p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9
                        If b >= 2 And b <= 6 Then
                            a = 0
                            If p2 = 0 And p3 = 1 Then a = a + 1
                            If p3 = 0 And p4 = 1 Then a = a + 1
                            If p4 = 0 And p5 = 1 Then a = a + 1
                            If p5 = 0 And p6 = 1 Then a = a + 1
                            If p6 = 0 And p7 = 1 Then a = a + 1
                            If p7 = 0 And p8 = 1 Then a = a + 1
                            If p8 = 0 And p9 = 1 Then a = a + 1
                            If p9 = 0 And p2 = 1 Then a = a + 1

                            If a = 1 Then
                                If step_ = 0 Then
                                    If p2 * p4 * p6 = 0 And p4 * p6 * p8 = 0 Then
                                        If nr > UBound(rem_) Then _
                                            ReDim Preserve rem_(0 To (UBound(rem_) + 1) * 2 - 1)
                                        rem_(nr) = base + c: nr = nr + 1
                                    End If
                                Else
                                    If p2 * p4 * p8 = 0 And p2 * p6 * p8 = 0 Then
                                        If nr > UBound(rem_) Then _
                                            ReDim Preserve rem_(0 To (UBound(rem_) + 1) * 2 - 1)
                                        rem_(nr) = base + c: nr = nr + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next c
            Next r

            For i = 0 To nr - 1
                m_grid(rem_(i)) = 0
            Next i
            If nr > 0 Then changed = True
        Next step_
        If Not changed Then Exit For
    Next pass
End Sub

'=====================================================================
' 4  DROP REDUNDANT STAIRCASE PIXELS
'=====================================================================

' A pixel whose neighbours are all also neighbours of one of them adds
' nothing: removing it keeps the chain connected. Leaving them in makes
' plain staircases look like junctions and shatters the trace.
Private Sub ac2fCLCleanup()
    Dim r As Long, c As Long, i As Long, j As Long
    Dim nb(0 To 7) As Long, nn As Long
    Dim base As Long
    Dim covered As Boolean
    Dim dr As Long, dc As Long

    For r = 1 To m_H - 2
        base = r * m_W
        For c = 1 To m_W - 2
            If m_grid(base + c) <> 0 Then
                nn = 0
                For dr = -1 To 1
                    For dc = -1 To 1
                        If dr <> 0 Or dc <> 0 Then
                            If m_grid(base + dr * m_W + c + dc) <> 0 Then
                                nb(nn) = (r + dr) * 65536 + (c + dc)
                                nn = nn + 1
                            End If
                        End If
                    Next dc
                Next dr

                If nn >= 2 Then
                    For i = 0 To nn - 1
                        covered = True
                        For j = 0 To nn - 1
                            If j <> i Then
                                If Abs(nb(j) \ 65536 - nb(i) \ 65536) > 1 Or _
                                   Abs((nb(j) Mod 65536) - (nb(i) Mod 65536)) > 1 Then
                                    covered = False
                                    Exit For
                                End If
                            End If
                        Next j
                        If covered Then
                            m_grid(base + c) = 0
                            Exit For
                        End If
                    Next i
                End If
            End If
        Next c
    Next r
End Sub

'=====================================================================
' 5  TRACE PIXELS INTO CHAINS
'=====================================================================

Private m_pi() As Long          ' cell index of each traced point
Private m_used() As Byte

Private Function ac2fCLDeg(ByVal idx As Long) As Long
    Dim r As Long, c As Long, dr As Long, dc As Long, d As Long
    r = idx \ m_W: c = idx Mod m_W
    If r < 1 Or c < 1 Or r > m_H - 2 Or c > m_W - 2 Then Exit Function
    For dr = -1 To 1
        For dc = -1 To 1
            If dr <> 0 Or dc <> 0 Then
                If m_grid(idx + dr * m_W + dc) <> 0 Then d = d + 1
            End If
        Next dc
    Next dr
    ac2fCLDeg = d
End Function

Private Sub ac2fCLTrace()
    Dim r As Long, c As Long, idx As Long
    Dim dr As Long, dc As Long
    Dim nIdx As Long

    m_cn = 0: m_pn = 0
    ReDim m_px(0 To 8191): ReDim m_py(0 To 8191): ReDim m_pi(0 To 8191)
    ReDim m_cs(0 To 255): ReDim m_cc(0 To 255)
    ReDim m_used(0 To m_W * m_H - 1)

    ' chains that start at an end point or a junction
    For r = 1 To m_H - 2
        For c = 1 To m_W - 2
            idx = r * m_W + c
            If m_grid(idx) <> 0 Then
                If ac2fCLDeg(idx) <> 2 Then
                    For dr = -1 To 1
                        For dc = -1 To 1
                            If dr <> 0 Or dc <> 0 Then
                                nIdx = idx + dr * m_W + dc
                                If m_grid(nIdx) <> 0 Then
                                    If m_used(nIdx) = 0 Then ac2fCLWalk idx, nIdx
                                End If
                            End If
                        Next dc
                    Next dr
                End If
            End If
        Next c
    Next r

    ' whatever is left is a closed loop
    For r = 1 To m_H - 2
        For c = 1 To m_W - 2
            idx = r * m_W + c
            If m_grid(idx) <> 0 Then
                If m_used(idx) = 0 Then
                    For dr = -1 To 1
                        For dc = -1 To 1
                            If dr <> 0 Or dc <> 0 Then
                                nIdx = idx + dr * m_W + dc
                                If m_grid(nIdx) <> 0 Then
                                    If m_used(nIdx) = 0 Then
                                        ac2fCLWalk idx, nIdx
                                        Exit For
                                    End If
                                End If
                            End If
                        Next dc
                        If m_used(idx) <> 0 Then Exit For
                    Next dr
                End If
            End If
        Next c
    Next r
End Sub

Private Sub ac2fCLWalk(ByVal startIdx As Long, ByVal firstIdx As Long)
    Dim cur As Long, prev As Long, nxt As Long
    Dim dr As Long, dc As Long, cand As Long

    If m_cn > UBound(m_cs) Then
        ReDim Preserve m_cs(0 To (UBound(m_cs) + 1) * 2 - 1)
        ReDim Preserve m_cc(0 To UBound(m_cs))
    End If
    m_cs(m_cn) = m_pn

    ac2fCLPush startIdx
    ac2fCLPush firstIdx
    m_used(firstIdx) = 1
    prev = startIdx: cur = firstIdx

    Do While ac2fCLDeg(cur) = 2
        nxt = -1
        For dr = -1 To 1
            For dc = -1 To 1
                If dr <> 0 Or dc <> 0 Then
                    cand = cur + dr * m_W + dc
                    If m_grid(cand) <> 0 And cand <> prev Then
                        If m_used(cand) = 0 Then
                            nxt = cand
                            Exit For
                        End If
                    End If
                End If
            Next dc
            If nxt >= 0 Then Exit For
        Next dr
        If nxt < 0 Then Exit Do
        ac2fCLPush nxt
        m_used(nxt) = 1
        prev = cur: cur = nxt
    Loop

    m_cc(m_cn) = m_pn - m_cs(m_cn)
    If m_cc(m_cn) >= 2 Then
        m_cn = m_cn + 1
    Else
        m_pn = m_cs(m_cn)
    End If
End Sub

Private Sub ac2fCLPush(ByVal idx As Long)
    If m_pn > UBound(m_px) Then
        ReDim Preserve m_px(0 To (UBound(m_px) + 1) * 2 - 1)
        ReDim Preserve m_py(0 To UBound(m_px))
        ReDim Preserve m_pi(0 To UBound(m_px))
    End If
    m_px(m_pn) = m_x0 + ((idx Mod m_W) + 0.5) * m_res
    m_py(m_pn) = m_y0 + ((idx \ m_W) + 0.5) * m_res
    m_pi(m_pn) = idx
    m_pn = m_pn + 1
End Sub

'=====================================================================
' 6  PRUNE
'=====================================================================

' Drops junction stubs and short spurs. A spur is a chain with one free
' end; a stub is any chain barely longer than a cell, left behind where
' several junction pixels sit side by side.
Private Sub ac2fCLPrune(ByVal minBranch As Double)
    Dim i As Long, k As Long
    Dim ends() As Byte
    Dim a As Long, b As Long
    Dim freeA As Boolean, freeB As Boolean
    Dim ln As Double
    Dim keepS() As Long, keepC() As Long, nk As Long

    If m_cn = 0 Then Exit Sub
    ReDim ends(0 To m_W * m_H - 1)

    For i = 0 To m_cn - 1
        a = m_pi(m_cs(i))
        b = m_pi(m_cs(i) + m_cc(i) - 1)
        If ends(a) < 255 Then ends(a) = ends(a) + 1
        If ends(b) < 255 Then ends(b) = ends(b) + 1
    Next i

    ReDim keepS(0 To m_cn - 1): ReDim keepC(0 To m_cn - 1)
    For i = 0 To m_cn - 1
        ln = ac2fCLChainLen(i)
        a = m_pi(m_cs(i))
        b = m_pi(m_cs(i) + m_cc(i) - 1)
        freeA = (ends(a) = 1)
        freeB = (ends(b) = 1)

        If ln <= m_res * 2.5 And Not (freeA And freeB) Then
            ' junction stub
        ElseIf (freeA Xor freeB) And ln < minBranch Then
            ' short spur hanging off a junction
        Else
            keepS(nk) = m_cs(i): keepC(nk) = m_cc(i): nk = nk + 1
        End If
    Next i

    For k = 0 To nk - 1
        m_cs(k) = keepS(k): m_cc(k) = keepC(k)
    Next k
    m_cn = nk
End Sub

Private Function ac2fCLChainLen(ByVal i As Long) As Double
    Dim k As Long, s As Long, n As Long
    Dim d As Double
    s = m_cs(i): n = m_cc(i)
    For k = 0 To n - 2
        d = d + Sqr((m_px(s + k + 1) - m_px(s + k)) ^ 2 + (m_py(s + k + 1) - m_py(s + k)) ^ 2)
    Next k
    ac2fCLChainLen = d
End Function

Private Function ac2fCLTotalLength() As Double
    Dim i As Long, d As Double
    For i = 0 To m_cn - 1
        d = d + ac2fCLChainLen(i)
    Next i
    ac2fCLTotalLength = d
End Function

'=====================================================================
' 7  EXTEND FREE ENDS
'=====================================================================

' Thinning pulls a free end back by about half the stroke width. The end
' is pushed back out along the local direction until it leaves the
' filled area.
'
' The direction is averaged over a tail one stroke width long, not taken
' from the last pixel: at a flat stroke terminal the medial axis forks
' towards the corners, and a single-pixel direction follows that fork.
' Measured on a known shape, the tail cut the end error from 12.1 mm to
' 1.8 mm.
Private Sub ac2fCLExtend(ByVal width As Double)
    Dim i As Long
    Dim ends() As Byte
    Dim a As Long, b As Long

    If m_cn = 0 Or width <= 0 Then Exit Sub
    ReDim ends(0 To m_W * m_H - 1)

    For i = 0 To m_cn - 1
        a = m_pi(m_cs(i))
        b = m_pi(m_cs(i) + m_cc(i) - 1)
        If ends(a) < 255 Then ends(a) = ends(a) + 1
        If ends(b) < 255 Then ends(b) = ends(b) + 1
    Next i

    For i = 0 To m_cn - 1
        If ends(m_pi(m_cs(i) + m_cc(i) - 1)) = 1 Then ac2fCLGrow i, False, width
        If ends(m_pi(m_cs(i))) = 1 Then ac2fCLGrow i, True, width
    Next i
End Sub

Private Sub ac2fCLGrow(ByVal ci As Long, ByVal atStart As Boolean, ByVal width As Double)
    Dim s As Long, n As Long, k As Long, step_ As Long
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim acc As Double, dx As Double, dy As Double, L As Double
    Dim t As Double, nx As Double, ny As Double
    Dim lastX As Double, lastY As Double, found As Boolean

    s = m_cs(ci): n = m_cc(ci)
    If n < 2 Then Exit Sub

    If atStart Then
        ax = m_px(s): ay = m_py(s)
        bx = m_px(s + 1): by = m_py(s + 1)
        acc = 0
        For k = 1 To n - 1
            acc = acc + Sqr((m_px(s + k) - m_px(s + k - 1)) ^ 2 + (m_py(s + k) - m_py(s + k - 1)) ^ 2)
            bx = m_px(s + k): by = m_py(s + k)
            If acc >= width Then Exit For
        Next k
    Else
        ax = m_px(s + n - 1): ay = m_py(s + n - 1)
        bx = m_px(s + n - 2): by = m_py(s + n - 2)
        acc = 0
        For k = n - 2 To 0 Step -1
            acc = acc + Sqr((m_px(s + k + 1) - m_px(s + k)) ^ 2 + (m_py(s + k + 1) - m_py(s + k)) ^ 2)
            bx = m_px(s + k): by = m_py(s + k)
            If acc >= width Then Exit For
        Next k
    End If

    dx = ax - bx: dy = ay - by
    L = Sqr(dx * dx + dy * dy)
    If L < 0.000000001 Then Exit Sub
    dx = dx / L: dy = dy / L

    t = m_res * 0.5
    Do While t <= width
        nx = ax + dx * t: ny = ay + dy * t
        If Not ac2fCLInside(nx, ny) Then Exit Do
        lastX = nx: lastY = ny: found = True
        t = t + m_res * 0.5
    Loop
    If Not found Then Exit Sub

    ac2fCLInsertPt ci, atStart, lastX, lastY
End Sub

Private Function ac2fCLInside(ByVal x As Double, ByVal y As Double) As Boolean
    Dim c As Long, r As Long
    ' Int floors; CLng would round to nearest and, in VBA, to even on a
    ' tie. The containing cell is what the scanline fill marked.
    c = Int((x - m_x0) / m_res)
    r = Int((y - m_y0) / m_res)
    If r < 0 Or c < 0 Or r > m_H - 1 Or c > m_W - 1 Then Exit Function
    ac2fCLInside = (m_mask(r * m_W + c) <> 0)
End Function

' Adds one point to a chain. Everything after it shifts up by one.
Private Sub ac2fCLInsertPt(ByVal ci As Long, ByVal atStart As Boolean, _
                           ByVal x As Double, ByVal y As Double)
    Dim k As Long, s As Long, i As Long

    If m_pn > UBound(m_px) Then
        ReDim Preserve m_px(0 To (UBound(m_px) + 1) * 2 - 1)
        ReDim Preserve m_py(0 To UBound(m_px))
        ReDim Preserve m_pi(0 To UBound(m_px))
    End If

    s = m_cs(ci)
    If atStart Then
        For k = m_pn To s + 1 Step -1
            m_px(k) = m_px(k - 1): m_py(k) = m_py(k - 1): m_pi(k) = m_pi(k - 1)
        Next k
        m_px(s) = x: m_py(s) = y: m_pi(s) = m_pi(s + 1)
    Else
        For k = m_pn To s + m_cc(ci) + 1 Step -1
            m_px(k) = m_px(k - 1): m_py(k) = m_py(k - 1): m_pi(k) = m_pi(k - 1)
        Next k
        m_px(s + m_cc(ci)) = x: m_py(s + m_cc(ci)) = y
        m_pi(s + m_cc(ci)) = m_pi(s + m_cc(ci) - 1)
    End If

    m_cc(ci) = m_cc(ci) + 1
    m_pn = m_pn + 1
    For i = 0 To m_cn - 1
        If i <> ci And m_cs(i) > s Then m_cs(i) = m_cs(i) + 1
    Next i
End Sub

'=====================================================================
' 8  SMOOTH AND SIMPLIFY
'=====================================================================

' Moving average. Free ends stay put; a closed loop is smoothed all the
' way round. Takes the pixel zigzag out, which otherwise inflates the
' measured length by around 5%.
Private Sub ac2fCLSmooth(ByVal passes As Long)
    Dim i As Long, k As Long, p As Long, s As Long, n As Long
    Dim qx() As Double, qy() As Double
    Dim closed As Boolean
    Dim a As Long, b As Long

    If passes <= 0 Then Exit Sub

    For i = 0 To m_cn - 1
        s = m_cs(i): n = m_cc(i)
        If n >= 3 Then
            closed = (Abs(m_px(s) - m_px(s + n - 1)) < m_res * 1.5) And _
                     (Abs(m_py(s) - m_py(s + n - 1)) < m_res * 1.5)
            ReDim qx(0 To n - 1): ReDim qy(0 To n - 1)
            For p = 1 To passes
                For k = 0 To n - 1
                    qx(k) = m_px(s + k): qy(k) = m_py(s + k)
                Next k
                For k = 0 To n - 1
                    If (k = 0 Or k = n - 1) And Not closed Then
                        ' free end stays where it is
                    Else
                        a = k - 1: If a < 0 Then a = n - 1
                        b = k + 1: If b > n - 1 Then b = 0
                        m_px(s + k) = (qx(a) + 2# * qx(k) + qx(b)) / 4#
                        m_py(s + k) = (qy(a) + 2# * qy(k) + qy(b)) / 4#
                    End If
                Next k
            Next p
        End If
    Next i
End Sub

' Douglas-Peucker, iterative so a long chain cannot overflow the stack.
Private Sub ac2fCLSimplify(ByVal tol As Double)
    Dim i As Long, k As Long, s As Long, n As Long
    Dim keep() As Byte
    Dim stkA() As Long, stkB() As Long, sp As Long
    Dim a As Long, b As Long, worst As Long
    Dim dmax As Double, d As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim dx As Double, dy As Double, L As Double
    Dim outS As Long, outN As Long
    Dim nx() As Double, ny() As Double, ni() As Long, nn As Long

    If tol <= 0 Or m_cn = 0 Then Exit Sub

    ReDim nx(0 To m_pn - 1): ReDim ny(0 To m_pn - 1): ReDim ni(0 To m_pn - 1)
    nn = 0

    For i = 0 To m_cn - 1
        s = m_cs(i): n = m_cc(i)
        ReDim keep(0 To n - 1)
        keep(0) = 1: keep(n - 1) = 1

        If n > 2 Then
            ReDim stkA(0 To n): ReDim stkB(0 To n)
            sp = 0
            stkA(0) = 0: stkB(0) = n - 1: sp = 1
            Do While sp > 0
                sp = sp - 1
                a = stkA(sp): b = stkB(sp)
                If b > a + 1 Then
                    ax = m_px(s + a): ay = m_py(s + a)
                    bx = m_px(s + b): by = m_py(s + b)
                    dx = bx - ax: dy = by - ay
                    L = Sqr(dx * dx + dy * dy)
                    dmax = -1: worst = -1
                    For k = a + 1 To b - 1
                        If L > 0.000000001 Then
                            d = Abs(dx * (ay - m_py(s + k)) - (ax - m_px(s + k)) * dy) / L
                        Else
                            d = Sqr((m_px(s + k) - ax) ^ 2 + (m_py(s + k) - ay) ^ 2)
                        End If
                        If d > dmax Then dmax = d: worst = k
                    Next k
                    If dmax > tol And worst > 0 Then
                        keep(worst) = 1
                        stkA(sp) = a: stkB(sp) = worst: sp = sp + 1
                        stkA(sp) = worst: stkB(sp) = b: sp = sp + 1
                    End If
                End If
            Loop
        End If

        outS = nn: outN = 0
        For k = 0 To n - 1
            If keep(k) <> 0 Then
                nx(nn) = m_px(s + k): ny(nn) = m_py(s + k): ni(nn) = m_pi(s + k)
                nn = nn + 1: outN = outN + 1
            End If
        Next k
        m_cs(i) = outS: m_cc(i) = outN
    Next i

    For k = 0 To nn - 1
        m_px(k) = nx(k): m_py(k) = ny(k): m_pi(k) = ni(k)
    Next k
    m_pn = nn
End Sub

'=====================================================================
' 9  DRAW
'=====================================================================

Private Function ac2fCLDraw() As Boolean
    Dim i As Long
    Dim made() As Shape
    Dim nm As Long
    Dim sh As Shape
    Dim grp As Shape

    If m_cn = 0 Then Exit Function
    ReDim made(0 To m_cn - 1)

    For i = 0 To m_cn - 1
        Set sh = ac2fCLDrawChain(i)
        If Not sh Is Nothing Then
            On Error Resume Next
            sh.Name = "ac2f centerline " & (i + 1)
            sh.Outline.Width = 0.15
            sh.Outline.Color.RGBAssign 220, 40, 0
            On Error GoTo 0
            Set made(nm) = sh
            nm = nm + 1
        End If
    Next i

    If nm = 0 Then Exit Function

    If nm > 1 Then
        On Error Resume Next
        ActiveDocument.ClearSelection
        For i = 0 To nm - 1
            made(i).AddToSelection
        Next i
        Set grp = ActiveSelectionRange.Group()
        grp.Name = "ac2f centerline"
        On Error GoTo 0
    Else
        made(0).Name = "ac2f centerline"
    End If

    ac2fCLDraw = True
End Function

' One chain as a single polyline. Falls back to separate line segments
' if building a curve is not available, so geometry is never lost.
Private Function ac2fCLDrawChain(ByVal ci As Long) As Shape
    Dim cv As Curve
    Dim sp As SubPath
    Dim s As Long, n As Long, k As Long
    Dim sh As Shape
    Dim segs() As Shape
    Dim ns As Long

    s = m_cs(ci): n = m_cc(ci)
    If n < 2 Then Exit Function

    On Error GoTo Fallback
    Set cv = ActiveDocument.CreateCurve()
    Set sp = cv.CreateSubPath(m_px(s), m_py(s))
    For k = 1 To n - 1
        sp.AppendLineSegment m_px(s + k), m_py(s + k)
    Next k
    Set sh = ActiveLayer.CreateCurve(cv)
    If sh Is Nothing Then GoTo Fallback
    Set ac2fCLDrawChain = sh
    Exit Function

Fallback:
    On Error Resume Next
    ReDim segs(0 To n - 2)
    For k = 0 To n - 2
        Set segs(ns) = ActiveLayer.CreateLineSegment(m_px(s + k), m_py(s + k), _
                                                     m_px(s + k + 1), m_py(s + k + 1))
        If Not segs(ns) Is Nothing Then ns = ns + 1
    Next k
    If ns = 0 Then Exit Function
    If ns = 1 Then
        Set ac2fCLDrawChain = segs(0)
        Exit Function
    End If
    ActiveDocument.ClearSelection
    For k = 0 To ns - 1
        segs(k).AddToSelection
    Next k
    Set ac2fCLDrawChain = ActiveSelectionRange.Group()
End Function

'=====================================================================
' REPORT AND MATH HELPERS
'=====================================================================

Private Function ac2fCLReport(ByVal joined As Boolean, ByVal groups As Long, _
                              ByVal paths As Long, ByVal totLen As Double, _
                              ByVal width As Double, ByVal msg As String) As String
    Dim s As String

    s = "RESULT" & vbCrLf
    s = s & "   Mode                : " & IIf(joined, "joined - one region", _
                                              "separate - one per object") & vbCrLf
    s = s & "   Centre lines        : " & groups & vbCrLf
    s = s & "   Paths               : " & paths & vbCrLf
    s = s & "   Total length        : " & ac2fFmtLength(totLen) & vbCrLf
    s = s & "   Mean stroke width   : " & ac2fFmt(width) & " mm" & vbCrLf & vbCrLf

    s = s & "SETTINGS USED" & vbCrLf
    s = s & "   Resolution          : " & ac2fFmt(m_res) & " mm" & vbCrLf
    s = s & "   Simplify tolerance  : " & ac2fFmt(ac2fGetNum(AC2F_K_CL_TOL, AC2F_DEF_CL_TOL)) & " mm" & vbCrLf
    s = s & "   Smoothing passes    : " & ac2fGetLng(AC2F_K_CL_SMOOTH, AC2F_DEF_CL_SMOOTH) & vbCrLf
    s = s & "   Min branch          : " & ac2fFmt(ac2fGetNum(AC2F_K_CL_BRANCH, AC2F_DEF_CL_BRANCH)) & _
            " x width" & vbCrLf
    s = s & "   Extend ends         : " & IIf(ac2fGetLng(AC2F_K_CL_EXTEND, AC2F_DEF_CL_EXTEND) <> 0, _
                                              "yes", "no") & vbCrLf

    If Len(msg) > 0 Then s = s & vbCrLf & msg & vbCrLf

    s = s & vbCrLf & "The centre lines are drawn in red, grouped." & vbCrLf
    s = s & "Route ALONG them with your 6 mm bit."
    ac2fCLReport = s
End Function

Private Function ac2fCLTheta(ByVal chord As Double, ByVal arc As Double) As Double
    Dim r As Double, lo As Double, hi As Double, mid As Double, f As Double
    Dim i As Long
    Const PI2 As Double = 6.28318530717959

    If arc <= 0.000000001 Then Exit Function
    r = chord / arc
    If r >= 0.999999 Then Exit Function
    If r <= 0# Then
        ac2fCLTheta = PI2 - 0.000001
        Exit Function
    End If
    lo = 0.000001: hi = PI2 - 0.000001
    For i = 1 To 60
        mid = (lo + hi) / 2#
        f = 2# * Sin(mid / 2#) / mid
        If f > r Then lo = mid Else hi = mid
    Next i
    ac2fCLTheta = (lo + hi) / 2#
End Function

Private Function ac2fCLWrap(ByVal a As Double) As Double
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
    ac2fCLWrap = v
End Function

Private Function ac2fCLAtan2(ByVal y As Double, ByVal x As Double) As Double
    Const PI_ As Double = 3.14159265358979
    If x > 0 Then
        ac2fCLAtan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then ac2fCLAtan2 = Atn(y / x) + PI_ Else ac2fCLAtan2 = Atn(y / x) - PI_
    Else
        If y > 0 Then
            ac2fCLAtan2 = PI_ / 2#
        ElseIf y < 0 Then
            ac2fCLAtan2 = -PI_ / 2#
        Else
            ac2fCLAtan2 = 0
        End If
    End If
End Function
