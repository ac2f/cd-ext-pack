Attribute VB_Name = "ac2fNestCore"
'=====================================================================
'  ac2f pack  --  ac2fNestCore
'
'  The nesting engine. Pure arithmetic: nothing in here touches a
'  CorelDRAW object, which is what lets tools/vbapy.py run this very
'  file outside CorelDRAW for the tests (tools/test_nest.py). ac2fNest
'  does everything that talks to the document and the user.
'
'  MODEL
'  Every part and every sheet is a raster of square cells (m_res mm).
'  A part's raster is CONSERVATIVE: a cell is occupied when any point of
'  the part, grown by half the part gap, could fall in it. Two parts whose
'  cells do not overlap are therefore at least one gap apart, whatever
'  their outline - concave shapes interlock, small parts drop into the
'  holes of big ones, and a drawn offcut works as a sheet.
'
'    occupied = cells whose centre is inside the material
'             + cells touched by a capsule of radius m_d around every edge
'
'  Material follows cutting logic: a closed contour inside another is a
'  hole, one inside that is material again (nesting depth parity). With
'  holes switched off every closed contour is filled solid.
'
'  A part row is a list of intervals [a, b). A sheet row is stored as RUNS:
'  run(c) > 0  free, that many free cells from c on
'  run(c) < 0  occupied, at least -run(c) cells from c on (a lower
'              bound: cells only ever fill up, so it can go stale but
'              never wrong, and a jump simply lands and jumps again)
'  so testing an interval is one lookup, and a blocked test says exactly
'  how far to jump. The leftmost free x for a given y is found by sliding
'  over the part rows until a whole pass needs no jump.
'
'  PLACEMENT
'  Bottom-left fill. For every allowed orientation the lowest y with a
'  free x is searched, and the orientation whose TOP ends lowest wins
'  (strategy 1) or the one resting lowest (strategy 2). Three per-row
'  "widest free run" checks reject most rows before any sliding, and a
'  memo per (orientation, sheet) remembers where the last search ended:
'  sheets only fill up, so nothing below it can have opened. Copies of
'  one part therefore cost almost nothing after the first.
'
'  OPTIMISING
'  A pass packs every part once in a given order. ac2fNCOptStep runs a
'  few deterministic orders and strategies first, then mutates the best
'  order at random; ac2fNest keeps calling it until its time budget is
'  spent. The best pass is kept: fewest unplaced parts, then the least
'  material (every used sheet counts in full except the last one opened,
'  which counts only up to its highest part).
'=====================================================================
Option Explicit

Private Const NC_MAX_DIM As Long = 32000      ' cells per side (runs are Integers)
Private Const NC_UNLIMITED As Long = -1

' contour kinds
Public Const AC2F_NC_CLOSED As Long = 0       ' cut path: depth parity decides
Public Const AC2F_NC_OPEN   As Long = 1       ' line: only its capsule counts
Public Const AC2F_NC_SOLID  As Long = 2       ' always material (safety blob)

'---------------------------------------------------------------------
' Geometry buffer: the contours of the one thing being rasterised
'---------------------------------------------------------------------
Private m_gx() As Double, m_gy() As Double
Private m_gn As Long, m_gCap As Long
Private m_gcs() As Long, m_gcc() As Long, m_gck() As Long
Private m_gcn As Long, m_gcCap As Long
Private m_gOpen As Boolean

'---------------------------------------------------------------------
' Containment pool: contours of every selected shape, for inner/outer
'---------------------------------------------------------------------
Private m_qx() As Double, m_qy() As Double
Private m_qn As Long, m_qCap As Long
Private m_qcs() As Long, m_qcc() As Long, m_qck() As Long
Private m_qcn As Long, m_qcCap As Long
Private m_qsC0() As Long, m_qsCn() As Long          ' contour range per shape
Private m_qsX0() As Double, m_qsY0() As Double
Private m_qsX1() As Double, m_qsY1() As Double
Private m_qsPar() As Long
Private m_qsn As Long, m_qsCap As Long

'---------------------------------------------------------------------
' Raster parameters
'---------------------------------------------------------------------
Private m_res As Double          ' cell size, mm
Private m_d As Double            ' growth of every part, mm (half gap + tolerance)
Private m_holes As Boolean       ' inner contours are holes other parts may use

' scratch: crossings and row intervals of the raster being built
Private m_xr() As Long, m_xx() As Double, m_xc() As Long
Private m_xn As Long, m_xCap As Long
Private m_tr() As Long, m_ta() As Long, m_tb() As Long, m_tf() As Long
Private m_tn As Long, m_tCap As Long
Private m_dep() As Long          ' depth of each contour of the buffer
Private m_ins() As Boolean       ' sweep state per contour
Private m_dcnt() As Long         ' sweep: contours inside, per depth

'---------------------------------------------------------------------
' Variants (one orientation of one part class)
'---------------------------------------------------------------------
Private m_vn As Long, m_vCap As Long
Private m_vW() As Long, m_vH() As Long           ' cells, 0 = unusable
Private m_vOX() As Double, m_vOY() As Double     ' raster origin, geometry frame
Private m_vR0() As Long                          ' first row in m_rs
Private m_vCells() As Long                       ' occupied cells
Private m_vK1() As Long, m_vS1() As Long         ' key rows and their spans
Private m_vK2() As Long, m_vS2() As Long
Private m_vK3() As Long, m_vS3() As Long
Private m_vTag() As Long, m_vClass() As Long
Private m_rs() As Long, m_rsn As Long, m_rsCap As Long
Private m_ia() As Long, m_ib() As Long, m_ivn As Long, m_ivCap As Long

'---------------------------------------------------------------------
' Classes (identical parts share one) and part instances
'---------------------------------------------------------------------
Private m_cn As Long, m_cCap As Long
Private m_cV0() As Long, m_cVn() As Long
Private m_pn As Long, m_pCap As Long
Private m_pClass() As Long

'---------------------------------------------------------------------
' Sheet types and sheet instances
'---------------------------------------------------------------------
Private m_tyn As Long, m_tyCap As Long
Private m_tyW() As Long, m_tyH() As Long
Private m_tyOX() As Double, m_tyOY() As Double   ' cell (0,0) corner, drawn sheets
Private m_tyQty() As Long
Private m_tyRect() As Boolean
Private m_tyFree() As Long                       ' free cells of an empty sheet
Private m_tyR0() As Long                         ' drawn: first cell in m_tyRun
Private m_tyRun() As Integer, m_tyRunN As Long

Private m_kn As Long                             ' instances available
Private m_kType() As Long
Private m_kBase() As Long, m_kRow0() As Long     ' -1 = not allocated yet
Private m_run() As Integer, m_runInit() As Integer, m_runN As Long
Private m_mf() As Integer, m_mfInit() As Integer, m_mfN As Long
Private m_memoY() As Long, m_memoX() As Long

' pass state
Private m_kOpen() As Boolean
Private m_kSeq() As Long                         ' opening order, 0 = closed
Private m_kParts() As Long, m_kUsedH() As Long, m_kCells() As Long
Private m_seq As Long
Private m_ord() As Long
Private m_rK() As Long, m_rV() As Long, m_rX() As Long, m_rY() As Long
Private m_curU As Long, m_curMat As Double, m_curTie As Double

' best pass kept
Private m_bK() As Long, m_bV() As Long, m_bX() As Long, m_bY() As Long
Private m_bOrd() As Long
Private m_bU As Long, m_bMat As Double, m_bTie As Double
Private m_bStrat As Long, m_bRot As Boolean, m_bStep As Long
Private m_bKParts() As Long, m_bKUsedH() As Long, m_bKCells() As Long
Private m_bKSeq() As Long
Private m_haveBest As Boolean

' optimiser
Private m_step As Long
Private m_seed As Double
Private m_anyRot As Boolean
Private m_force As Long          ' 0 = try both strategies, 1 / 2 = only that one

'=====================================================================
' SET-UP
'=====================================================================

' Starts a fresh job. res = cell size, grow = how far every part is grown
' (half the part gap plus the outline tolerance), holes = inner contours
' are free for other parts.
Public Sub ac2fNCInit(ByVal res As Double, ByVal grow As Double, ByVal holes As Boolean)
    m_res = res
    If m_res <= 0 Then m_res = 1#
    m_d = grow
    If m_d < 0 Then m_d = 0
    m_holes = holes

    m_vn = 0: m_vCap = 0
    m_rsn = 0: m_rsCap = 0
    m_ivn = 0: m_ivCap = 0
    m_cn = 0: m_cCap = 0
    m_pn = 0: m_pCap = 0
    m_tyn = 0: m_tyCap = 0
    m_tyRunN = 0
    m_kn = 0
    m_runN = 0: m_mfN = 0
    m_haveBest = False
    m_step = 0
    m_seed = 12345
    m_force = 0
    ac2fNCGeoClear
End Sub

' 0 lets the optimiser try both placement strategies, 1 (lowest top edge)
' or 2 (bottom-left, lowest resting y) fixes one.
Public Sub ac2fNCSetStrategy(ByVal strat As Long)
    m_force = strat
    If m_force < 0 Or m_force > 2 Then m_force = 0
End Sub

Public Function ac2fNCRes() As Double
    ac2fNCRes = m_res
End Function

Public Function ac2fNCGrow() As Double
    ac2fNCGrow = m_d
End Function

'=====================================================================
' GEOMETRY BUFFER
'=====================================================================

Public Sub ac2fNCGeoClear()
    If m_gCap = 0 Then
        m_gCap = 4096
        ReDim m_gx(0 To m_gCap - 1)
        ReDim m_gy(0 To m_gCap - 1)
    End If
    If m_gcCap = 0 Then
        m_gcCap = 256
        ReDim m_gcs(0 To m_gcCap - 1)
        ReDim m_gcc(0 To m_gcCap - 1)
        ReDim m_gck(0 To m_gcCap - 1)
    End If
    m_gn = 0
    m_gcn = 0
    m_gOpen = False
End Sub

Public Sub ac2fNCGeoBegin(ByVal kind As Long)
    If m_gOpen Then ac2fNCGeoEnd
    If m_gcn >= m_gcCap Then
        m_gcCap = m_gcCap * 2
        ReDim Preserve m_gcs(0 To m_gcCap - 1)
        ReDim Preserve m_gcc(0 To m_gcCap - 1)
        ReDim Preserve m_gck(0 To m_gcCap - 1)
    End If
    m_gcs(m_gcn) = m_gn
    m_gcc(m_gcn) = 0
    m_gck(m_gcn) = kind
    m_gOpen = True
End Sub

Public Sub ac2fNCGeoPt(ByVal x As Double, ByVal y As Double)
    If Not m_gOpen Then ac2fNCGeoBegin AC2F_NC_CLOSED
    If m_gn > m_gcs(m_gcn) Then
        ' a point that repeats the previous one adds nothing
        If Abs(m_gx(m_gn - 1) - x) < 0.000001 Then
            If Abs(m_gy(m_gn - 1) - y) < 0.000001 Then Exit Sub
        End If
    End If
    If m_gn >= m_gCap Then
        m_gCap = m_gCap * 2
        ReDim Preserve m_gx(0 To m_gCap - 1)
        ReDim Preserve m_gy(0 To m_gCap - 1)
    End If
    m_gx(m_gn) = x
    m_gy(m_gn) = y
    m_gn = m_gn + 1
End Sub

Public Sub ac2fNCGeoEnd()
    Dim s As Long, n As Long

    If Not m_gOpen Then Exit Sub
    m_gOpen = False
    s = m_gcs(m_gcn)
    n = m_gn - s

    ' a closed contour that repeats its first point at the end
    If n > 1 And m_gck(m_gcn) <> AC2F_NC_OPEN Then
        If Abs(m_gx(m_gn - 1) - m_gx(s)) < 0.000001 Then
            If Abs(m_gy(m_gn - 1) - m_gy(s)) < 0.000001 Then
                m_gn = m_gn - 1
                n = n - 1
            End If
        End If
    End If

    If n = 1 Then
        ' a dot: keep it as a zero length line so its capsule still counts
        If m_gn >= m_gCap Then
            m_gCap = m_gCap * 2
            ReDim Preserve m_gx(0 To m_gCap - 1)
            ReDim Preserve m_gy(0 To m_gCap - 1)
        End If
        m_gx(m_gn) = m_gx(s)
        m_gy(m_gn) = m_gy(s)
        m_gn = m_gn + 1
        n = 2
        m_gck(m_gcn) = AC2F_NC_OPEN
    ElseIf n = 2 And m_gck(m_gcn) <> AC2F_NC_OPEN Then
        m_gck(m_gcn) = AC2F_NC_OPEN
    End If

    If n >= 2 Then
        m_gcc(m_gcn) = n
        m_gcn = m_gcn + 1
    Else
        m_gn = s
    End If
End Sub

Public Function ac2fNCGeoContours() As Long
    If m_gOpen Then ac2fNCGeoEnd
    ac2fNCGeoContours = m_gcn
End Function

Public Function ac2fNCGeoPoints() As Long
    ac2fNCGeoPoints = m_gn
End Function

' Applies x' = a*x + b*y, y' = c*x + e*y to every point of the buffer.
Public Sub ac2fNCGeoTransform(ByVal a As Double, ByVal b As Double, _
                              ByVal c As Double, ByVal e As Double)
    Dim i As Long
    Dim x As Double, y As Double
    If m_gOpen Then ac2fNCGeoEnd
    For i = 0 To m_gn - 1
        x = m_gx(i): y = m_gy(i)
        m_gx(i) = a * x + b * y
        m_gy(i) = c * x + e * y
    Next i
End Sub

Public Function ac2fNCGeoBBox(ByRef x0 As Double, ByRef y0 As Double, _
                              ByRef x1 As Double, ByRef y1 As Double) As Boolean
    Dim i As Long
    If m_gOpen Then ac2fNCGeoEnd
    If m_gn = 0 Then Exit Function
    x0 = m_gx(0): x1 = m_gx(0): y0 = m_gy(0): y1 = m_gy(0)
    For i = 1 To m_gn - 1
        If m_gx(i) < x0 Then x0 = m_gx(i)
        If m_gx(i) > x1 Then x1 = m_gx(i)
        If m_gy(i) < y0 Then y0 = m_gy(i)
        If m_gy(i) > y1 Then y1 = m_gy(i)
    Next i
    ac2fNCGeoBBox = True
End Function

' Shape fingerprint, independent of position: identical parts give the
' same numbers, a mirrored or edited part does not. Coordinates count to
' 0.01 mm relative to the bounding box.
Public Function ac2fNCGeoSig() As String
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim i As Long, k As Long
    Dim hsh As Double, v As Double
    Const MODP As Double = 2147483647#

    If Not ac2fNCGeoBBox(x0, y0, x1, y1) Then Exit Function
    hsh = 7
    For k = 0 To m_gcn - 1
        hsh = hsh * 31 + m_gcc(k) * 4 + m_gck(k)
        hsh = hsh - Int(hsh / MODP) * MODP
    Next k
    For i = 0 To m_gn - 1
        v = Int((m_gx(i) - x0) * 100# + 0.5)
        hsh = hsh * 131 + v
        hsh = hsh - Int(hsh / MODP) * MODP
        v = Int((m_gy(i) - y0) * 100# + 0.5)
        hsh = hsh * 137 + v
        hsh = hsh - Int(hsh / MODP) * MODP
    Next i
    ac2fNCGeoSig = CStr(m_gcn) & "/" & CStr(m_gn) & "/" & _
                   CStr(Int((x1 - x0) * 100# + 0.5)) & "x" & _
                   CStr(Int((y1 - y0) * 100# + 0.5)) & "/" & CStr(hsh)
End Function

' Material area of the buffer in mm2: closed contours count by nesting
' depth (even adds, odd is a hole), lines and safety blobs count nothing.
Public Function ac2fNCGeoArea() As Double
    Dim k As Long, i As Long, j As Long, s As Long, n As Long
    Dim a As Double, tot As Double

    If m_gOpen Then ac2fNCGeoEnd
    If m_gcn = 0 Then Exit Function
    ac2fNCDepths
    For k = 0 To m_gcn - 1
        If m_gck(k) = AC2F_NC_CLOSED Then
            s = m_gcs(k): n = m_gcc(k)
            a = 0
            j = n - 1
            For i = 0 To n - 1
                a = a + (m_gx(s + j) * m_gy(s + i) - m_gx(s + i) * m_gy(s + j))
                j = i
            Next i
            a = Abs(a) / 2#
            If (m_dep(k) Mod 2) = 0 Then tot = tot + a Else tot = tot - a
        End If
    Next k
    If tot < 0 Then tot = 0
    ac2fNCGeoArea = tot
End Function

'=====================================================================
' CONTAINMENT  (which selected shape lies inside which)
'=====================================================================

Public Sub ac2fNCPoolClear()
    m_qn = 0: m_qcn = 0: m_qsn = 0
    If m_qCap = 0 Then
        m_qCap = 4096
        ReDim m_qx(0 To m_qCap - 1)
        ReDim m_qy(0 To m_qCap - 1)
    End If
    If m_qcCap = 0 Then
        m_qcCap = 256
        ReDim m_qcs(0 To m_qcCap - 1)
        ReDim m_qcc(0 To m_qcCap - 1)
        ReDim m_qck(0 To m_qcCap - 1)
    End If
    If m_qsCap = 0 Then
        m_qsCap = 64
        ReDim m_qsC0(0 To m_qsCap - 1)
        ReDim m_qsCn(0 To m_qsCap - 1)
        ReDim m_qsX0(0 To m_qsCap - 1)
        ReDim m_qsY0(0 To m_qsCap - 1)
        ReDim m_qsX1(0 To m_qsCap - 1)
        ReDim m_qsY1(0 To m_qsCap - 1)
        ReDim m_qsPar(0 To m_qsCap - 1)
    End If
End Sub

' Copies the geometry buffer into the pool as shape number ac2fNCPoolCount-1.
Public Function ac2fNCPoolAdd() As Long
    Dim i As Long, k As Long, s As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double

    If m_gOpen Then ac2fNCGeoEnd
    If m_qsn >= m_qsCap Then
        m_qsCap = m_qsCap * 2
        ReDim Preserve m_qsC0(0 To m_qsCap - 1)
        ReDim Preserve m_qsCn(0 To m_qsCap - 1)
        ReDim Preserve m_qsX0(0 To m_qsCap - 1)
        ReDim Preserve m_qsY0(0 To m_qsCap - 1)
        ReDim Preserve m_qsX1(0 To m_qsCap - 1)
        ReDim Preserve m_qsY1(0 To m_qsCap - 1)
        ReDim Preserve m_qsPar(0 To m_qsCap - 1)
    End If
    Do While m_qn + m_gn > m_qCap
        m_qCap = m_qCap * 2
        ReDim Preserve m_qx(0 To m_qCap - 1)
        ReDim Preserve m_qy(0 To m_qCap - 1)
    Loop
    Do While m_qcn + m_gcn > m_qcCap
        m_qcCap = m_qcCap * 2
        ReDim Preserve m_qcs(0 To m_qcCap - 1)
        ReDim Preserve m_qcc(0 To m_qcCap - 1)
        ReDim Preserve m_qck(0 To m_qcCap - 1)
    Loop

    s = m_qsn
    m_qsC0(s) = m_qcn
    m_qsCn(s) = m_gcn
    m_qsPar(s) = -1
    If Not ac2fNCGeoBBox(x0, y0, x1, y1) Then
        x0 = 0: y0 = 0: x1 = 0: y1 = 0
    End If
    m_qsX0(s) = x0: m_qsY0(s) = y0: m_qsX1(s) = x1: m_qsY1(s) = y1

    For k = 0 To m_gcn - 1
        m_qcs(m_qcn) = m_qn + m_gcs(k)
        m_qcc(m_qcn) = m_gcc(k)
        m_qck(m_qcn) = m_gck(k)
        m_qcn = m_qcn + 1
    Next k
    For i = 0 To m_gn - 1
        m_qx(m_qn + i) = m_gx(i)
        m_qy(m_qn + i) = m_gy(i)
    Next i
    m_qn = m_qn + m_gn
    m_qsn = m_qsn + 1
    ac2fNCPoolAdd = s
End Function

Public Function ac2fNCPoolCount() As Long
    ac2fNCPoolCount = m_qsn
End Function

' Works out, for every shape, the outermost shape it sits inside (or -1).
' "Inside" means three points spread along the shape's first contour all
' fall within a closed contour of the other shape, and the other shape's
' box is strictly bigger. A shape inside one that is itself inside a third
' goes with the third, so a whole nest of frames moves as one part.
Public Sub ac2fNCPoolResolve()
    Dim i As Long, j As Long, best As Long, t As Long
    Dim bestArea As Double, ar As Double, ai As Double
    Dim par() As Long
    Dim guard As Long

    If m_qsn = 0 Then Exit Sub
    ReDim par(0 To m_qsn - 1)
    For i = 0 To m_qsn - 1
        par(i) = -1
        best = -1
        ai = (m_qsX1(i) - m_qsX0(i)) * (m_qsY1(i) - m_qsY0(i))
        If m_qsCn(i) > 0 Then
            For j = 0 To m_qsn - 1
                If j <> i Then
                    If m_qsX0(j) <= m_qsX0(i) + 0.000001 And m_qsY0(j) <= m_qsY0(i) + 0.000001 And _
                       m_qsX1(j) >= m_qsX1(i) - 0.000001 And m_qsY1(j) >= m_qsY1(i) - 0.000001 Then
                        ar = (m_qsX1(j) - m_qsX0(j)) * (m_qsY1(j) - m_qsY0(j))
                        If ar > ai + 0.000001 Then
                            If best < 0 Or ar < bestArea Then
                                If ac2fNCPoolInside(i, j) Then
                                    best = j
                                    bestArea = ar
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
        End If
        par(i) = best
    Next i

    ' follow to the outermost container
    For i = 0 To m_qsn - 1
        t = i
        guard = 0
        Do While par(t) >= 0 And guard <= m_qsn
            t = par(t)
            guard = guard + 1
        Loop
        If t = i Then
            m_qsPar(i) = -1
        Else
            m_qsPar(i) = t
        End If
    Next i
End Sub

Public Function ac2fNCPoolParent(ByVal i As Long) As Long
    If i < 0 Or i >= m_qsn Then
        ac2fNCPoolParent = -1
    Else
        ac2fNCPoolParent = m_qsPar(i)
    End If
End Function

Private Function ac2fNCPoolInside(ByVal i As Long, ByVal j As Long) As Boolean
    Dim c As Long, n As Long, s As Long, t As Long, k As Long
    Dim px As Double, py As Double
    Dim hit As Boolean

    c = m_qsC0(i)
    s = m_qcs(c): n = m_qcc(c)
    If n < 1 Then Exit Function
    For t = 0 To 2
        k = s + (n * t) \ 3
        px = m_qx(k): py = m_qy(k)
        hit = False
        For c = m_qsC0(j) To m_qsC0(j) + m_qsCn(j) - 1
            If m_qck(c) <> AC2F_NC_OPEN Then
                If ac2fNCPip(m_qx, m_qy, m_qcs(c), m_qcc(c), px, py) Then
                    hit = True
                    Exit For
                End If
            End If
        Next c
        If Not hit Then Exit Function
    Next t
    ac2fNCPoolInside = True
End Function

' Even-odd point in polygon over points s .. s+n-1 of the given arrays.
Private Function ac2fNCPip(ByRef xs() As Double, ByRef ys() As Double, _
                           ByVal s As Long, ByVal n As Long, _
                           ByVal px As Double, ByVal py As Double) As Boolean
    Dim i As Long, j As Long
    Dim inside As Boolean
    Dim xi As Double, yi As Double, xj As Double, yj As Double

    j = s + n - 1
    For i = s To s + n - 1
        xi = xs(i): yi = ys(i): xj = xs(j): yj = ys(j)
        If (yi > py) <> (yj > py) Then
            If px < xi + (py - yi) * (xj - xi) / (yj - yi) Then inside = Not inside
        End If
        j = i
    Next i
    ac2fNCPip = inside
End Function

'=====================================================================
' RASTER BUILDING
'=====================================================================

Private Sub ac2fNCScratch()
    If m_xCap = 0 Then
        m_xCap = 4096
        ReDim m_xr(0 To m_xCap - 1)
        ReDim m_xx(0 To m_xCap - 1)
        ReDim m_xc(0 To m_xCap - 1)
    End If
    If m_tCap = 0 Then
        m_tCap = 8192
        ReDim m_tr(0 To m_tCap - 1)
        ReDim m_ta(0 To m_tCap - 1)
        ReDim m_tb(0 To m_tCap - 1)
        ReDim m_tf(0 To m_tCap - 1)
    End If
    m_xn = 0
    m_tn = 0
End Sub

Private Sub ac2fNCAddCross(ByVal r As Long, ByVal x As Double, ByVal c As Long)
    If m_xn >= m_xCap Then
        m_xCap = m_xCap * 2
        ReDim Preserve m_xr(0 To m_xCap - 1)
        ReDim Preserve m_xx(0 To m_xCap - 1)
        ReDim Preserve m_xc(0 To m_xCap - 1)
    End If
    m_xr(m_xn) = r
    m_xx(m_xn) = x
    m_xc(m_xn) = c
    m_xn = m_xn + 1
End Sub

Private Sub ac2fNCAddIv(ByVal r As Long, ByVal a As Long, ByVal b As Long, ByVal f As Long)
    If b <= a Then Exit Sub
    If m_tn >= m_tCap Then
        m_tCap = m_tCap * 2
        ReDim Preserve m_tr(0 To m_tCap - 1)
        ReDim Preserve m_ta(0 To m_tCap - 1)
        ReDim Preserve m_tb(0 To m_tCap - 1)
        ReDim Preserve m_tf(0 To m_tCap - 1)
    End If
    m_tr(m_tn) = r
    m_ta(m_tn) = a
    m_tb(m_tn) = b
    m_tf(m_tn) = f
    m_tn = m_tn + 1
End Sub

' Nesting depth of every closed contour of the buffer: how many other
' closed contours hold its first edge's midpoint.
Private Function ac2fNCDepths() As Long
    Dim k As Long, j As Long
    Dim px As Double, py As Double
    Dim s As Long, maxD As Long
    Dim bx0() As Double, by0() As Double, bx1() As Double, by1() As Double
    Dim i As Long

    ReDim m_dep(0 To m_gcn)
    ReDim bx0(0 To m_gcn): ReDim by0(0 To m_gcn)
    ReDim bx1(0 To m_gcn): ReDim by1(0 To m_gcn)
    For k = 0 To m_gcn - 1
        s = m_gcs(k)
        bx0(k) = m_gx(s): bx1(k) = m_gx(s): by0(k) = m_gy(s): by1(k) = m_gy(s)
        For i = s + 1 To s + m_gcc(k) - 1
            If m_gx(i) < bx0(k) Then bx0(k) = m_gx(i)
            If m_gx(i) > bx1(k) Then bx1(k) = m_gx(i)
            If m_gy(i) < by0(k) Then by0(k) = m_gy(i)
            If m_gy(i) > by1(k) Then by1(k) = m_gy(i)
        Next i
    Next k

    For k = 0 To m_gcn - 1
        m_dep(k) = 0
        If m_gck(k) = AC2F_NC_CLOSED Then
            s = m_gcs(k)
            px = (m_gx(s) + m_gx(s + 1)) / 2#
            py = (m_gy(s) + m_gy(s + 1)) / 2#
            For j = 0 To m_gcn - 1
                If j <> k And m_gck(j) = AC2F_NC_CLOSED Then
                    If px >= bx0(j) And px <= bx1(j) And py >= by0(j) And py <= by1(j) Then
                        If ac2fNCPip(m_gx, m_gy, m_gcs(j), m_gcc(j), px, py) Then
                            m_dep(k) = m_dep(k) + 1
                        End If
                    End If
                End If
            Next j
            If m_dep(k) > maxD Then maxD = m_dep(k)
        End If
    Next k
    ac2fNCDepths = maxD
End Function

' Interior: cells whose centre is inside the material. Adds intervals
' with flag 0 for rows 0 .. h-1 of a raster whose cell (0,0) corner is
' at (ox, oy).
Private Sub ac2fNCInterior(ByVal ox As Double, ByVal oy As Double, _
                           ByVal w As Long, ByVal h As Long, ByVal useHoles As Boolean)
    Dim k As Long, i As Long, j As Long, s As Long, n As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim ylo As Double, yhi As Double, yc As Double
    Dim r As Long, r0 As Long, r1 As Long
    Dim maxD As Long
    Dim cnt() As Long, ord() As Long
    Dim q As Long, e As Long, t As Long, u As Long
    Dim tx As Double
    Dim solid As Long, closedIn As Long, md As Long
    Dim mat As Boolean, was As Boolean
    Dim xa As Double
    Dim ca As Long, cb As Long

    maxD = ac2fNCDepths()
    ReDim m_ins(0 To m_gcn)
    ReDim m_dcnt(0 To maxD + 1)

    ' crossings of every centre scanline, edge by edge
    For k = 0 To m_gcn - 1
        If m_gck(k) <> AC2F_NC_OPEN Then
            s = m_gcs(k): n = m_gcc(k)
            For i = 0 To n - 1
                j = i + 1
                If j = n Then j = 0
                x1 = m_gx(s + i): y1 = m_gy(s + i)
                x2 = m_gx(s + j): y2 = m_gy(s + j)
                If y1 <> y2 Then
                    If y1 < y2 Then
                        ylo = y1: yhi = y2
                    Else
                        ylo = y2: yhi = y1
                    End If
                    ' rows whose centre yc satisfies ylo <= yc < yhi
                    r0 = -Int(-((ylo - oy) / m_res - 0.5))
                    r1 = -Int(-((yhi - oy) / m_res - 0.5)) - 1
                    If r0 < 0 Then r0 = 0
                    If r1 > h - 1 Then r1 = h - 1
                    For r = r0 To r1
                        yc = oy + (r + 0.5) * m_res
                        ac2fNCAddCross r, x1 + (yc - y1) * (x2 - x1) / (y2 - y1), k
                    Next r
                End If
            Next i
        End If
    Next k
    If m_xn = 0 Then Exit Sub

    ' bucket by row
    ReDim cnt(0 To h)
    For q = 0 To m_xn - 1
        cnt(m_xr(q) + 1) = cnt(m_xr(q) + 1) + 1
    Next q
    For r = 1 To h
        cnt(r) = cnt(r) + cnt(r - 1)
    Next r
    ReDim ord(0 To m_xn - 1)
    For q = 0 To m_xn - 1
        r = m_xr(q)
        ord(cnt(r)) = q
        cnt(r) = cnt(r) + 1
    Next q
    ' cnt(r) is now the end of row r; its start is cnt(r - 1) (0 for row 0)

    s = 0
    For r = 0 To h - 1
        e = cnt(r)
        If e - s >= 2 Then
            ' sort the row's crossings by x (insertion; rows are short)
            For t = s + 1 To e - 1
                u = ord(t)
                tx = m_xx(u)
                q = t - 1
                Do While q >= s
                    If m_xx(ord(q)) <= tx Then Exit Do
                    ord(q + 1) = ord(q)
                    q = q - 1
                Loop
                ord(q + 1) = u
            Next t

            solid = 0: closedIn = 0: was = False
            For t = s To e - 1
                u = ord(t)
                k = m_xc(u)
                m_ins(k) = Not m_ins(k)
                If m_gck(k) = AC2F_NC_SOLID Then
                    If m_ins(k) Then solid = solid + 1 Else solid = solid - 1
                Else
                    If m_ins(k) Then
                        closedIn = closedIn + 1
                        m_dcnt(m_dep(k)) = m_dcnt(m_dep(k)) + 1
                    Else
                        closedIn = closedIn - 1
                        m_dcnt(m_dep(k)) = m_dcnt(m_dep(k)) - 1
                    End If
                End If

                If useHoles Then
                    md = -1
                    For q = maxD To 0 Step -1
                        If m_dcnt(q) > 0 Then
                            md = q
                            Exit For
                        End If
                    Next q
                    mat = (md >= 0 And (md Mod 2) = 0)
                Else
                    mat = (closedIn > 0)
                End If
                If solid > 0 Then mat = True

                If mat And Not was Then
                    xa = m_xx(u)
                ElseIf was And Not mat Then
                    ' cells whose centre lies in [xa, x)
                    ca = -Int(-((xa - ox) / m_res - 0.5))
                    cb = -Int(-((m_xx(u) - ox) / m_res - 0.5))
                    If ca < 0 Then ca = 0
                    If cb > w Then cb = w
                    ac2fNCAddIv r, ca, cb, 0
                End If
                was = mat
            Next t
        End If
        s = e
    Next r
End Sub

' Capsules: every cell within rad of an edge. Adds intervals with flag f.
Private Sub ac2fNCCapsules(ByVal ox As Double, ByVal oy As Double, _
                           ByVal w As Long, ByVal h As Long, _
                           ByVal rad As Double, ByVal f As Long)
    Dim k As Long, i As Long, j As Long, s As Long, n As Long, ne As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim ylo As Double, yhi As Double, b0 As Double, b1 As Double
    Dim t0 As Double, t1 As Double, xa As Double, xb As Double, tt As Double
    Dim r As Long, r0 As Long, r1 As Long, ca As Long, cb As Long

    For k = 0 To m_gcn - 1
        s = m_gcs(k): n = m_gcc(k)
        ne = n
        If m_gck(k) = AC2F_NC_OPEN Then ne = n - 1
        For i = 0 To ne - 1
            j = i + 1
            If j = n Then j = 0
            x1 = m_gx(s + i): y1 = m_gy(s + i)
            x2 = m_gx(s + j): y2 = m_gy(s + j)
            If y1 < y2 Then
                ylo = y1: yhi = y2
            Else
                ylo = y2: yhi = y1
            End If
            r0 = Int((ylo - rad - oy) / m_res)
            r1 = Int((yhi + rad - oy) / m_res)
            If r0 < 0 Then r0 = 0
            If r1 > h - 1 Then r1 = h - 1
            For r = r0 To r1
                ' the part of the edge with y in the row band grown by rad
                b0 = oy + r * m_res - rad
                b1 = oy + (r + 1) * m_res + rad
                If y1 = y2 Then
                    xa = x1: xb = x2
                Else
                    t0 = (b0 - y1) / (y2 - y1)
                    t1 = (b1 - y1) / (y2 - y1)
                    If t0 > t1 Then
                        tt = t0: t0 = t1: t1 = tt
                    End If
                    If t0 < 0 Then t0 = 0
                    If t1 > 1 Then t1 = 1
                    xa = x1 + t0 * (x2 - x1)
                    xb = x1 + t1 * (x2 - x1)
                End If
                If xa > xb Then
                    tt = xa: xa = xb: xb = tt
                End If
                ca = Int((xa - rad - ox) / m_res)
                cb = Int((xb + rad - ox) / m_res) + 1
                If ca < 0 Then ca = 0
                If cb > w Then cb = w
                ac2fNCAddIv r, ca, cb, f
            Next r
        Next i
    Next k
End Sub

' Sorts the scratch intervals by row, then by start, in place of m_t*.
' On return cnt(r) .. cnt(r + 1) - 1 index row r in (sa, sb, sf).
Private Sub ac2fNCSortRows(ByVal h As Long, ByRef cnt() As Long, _
                           ByRef sa() As Long, ByRef sb() As Long, ByRef sf() As Long)
    Dim q As Long, r As Long, t As Long, p As Long
    Dim pos() As Long
    Dim ta As Long, tb As Long, tf As Long

    ReDim cnt(0 To h + 1)
    For q = 0 To m_tn - 1
        cnt(m_tr(q) + 1) = cnt(m_tr(q) + 1) + 1
    Next q
    For r = 1 To h + 1
        cnt(r) = cnt(r) + cnt(r - 1)
    Next r
    If m_tn = 0 Then Exit Sub
    ReDim pos(0 To h)
    For r = 0 To h
        pos(r) = cnt(r)
    Next r
    ReDim sa(0 To m_tn - 1): ReDim sb(0 To m_tn - 1): ReDim sf(0 To m_tn - 1)
    For q = 0 To m_tn - 1
        r = m_tr(q)
        p = pos(r)
        sa(p) = m_ta(q): sb(p) = m_tb(q): sf(p) = m_tf(q)
        pos(r) = p + 1
    Next q
    For r = 0 To h - 1
        For t = cnt(r) + 1 To cnt(r + 1) - 1
            ta = sa(t): tb = sb(t): tf = sf(t)
            p = t - 1
            Do While p >= cnt(r)
                If sa(p) <= ta Then Exit Do
                sa(p + 1) = sa(p): sb(p + 1) = sb(p): sf(p + 1) = sf(p)
                p = p - 1
            Loop
            sa(p + 1) = ta: sb(p + 1) = tb: sf(p + 1) = tf
        Next t
    Next r
End Sub

'=====================================================================
' PART CLASSES AND VARIANTS
'=====================================================================

' Starts a new part class; the variants added next belong to it.
Public Function ac2fNCAddClass() As Long
    If m_cn >= m_cCap Then
        If m_cCap = 0 Then m_cCap = 64 Else m_cCap = m_cCap * 2
        ReDim Preserve m_cV0(0 To m_cCap - 1)
        ReDim Preserve m_cVn(0 To m_cCap - 1)
    End If
    m_cV0(m_cn) = m_vn
    m_cVn(m_cn) = 0
    m_cn = m_cn + 1
    ac2fNCAddClass = m_cn - 1
End Function

Public Function ac2fNCClassCount() As Long
    ac2fNCClassCount = m_cn
End Function

' Rasterises the geometry buffer as the next variant of the current class.
' tag is the caller's own number for it (which turn, mirrored or not).
' Returns the variant index; its width is 0 when it could not be built.
Public Function ac2fNCAddVariant(ByVal tag As Long) As Long
    Dim v As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim ox As Double, oy As Double
    Dim w As Long, h As Long
    Dim cnt() As Long, sa() As Long, sb() As Long, sf() As Long
    Dim r As Long, t As Long, a As Long, b As Long
    Dim cells As Long, span As Long

    If m_cn = 0 Then ac2fNCAddClass
    If m_vn >= m_vCap Then
        If m_vCap = 0 Then m_vCap = 64 Else m_vCap = m_vCap * 2
        ReDim Preserve m_vW(0 To m_vCap - 1): ReDim Preserve m_vH(0 To m_vCap - 1)
        ReDim Preserve m_vOX(0 To m_vCap - 1): ReDim Preserve m_vOY(0 To m_vCap - 1)
        ReDim Preserve m_vR0(0 To m_vCap - 1): ReDim Preserve m_vCells(0 To m_vCap - 1)
        ReDim Preserve m_vK1(0 To m_vCap - 1): ReDim Preserve m_vS1(0 To m_vCap - 1)
        ReDim Preserve m_vK2(0 To m_vCap - 1): ReDim Preserve m_vS2(0 To m_vCap - 1)
        ReDim Preserve m_vK3(0 To m_vCap - 1): ReDim Preserve m_vS3(0 To m_vCap - 1)
        ReDim Preserve m_vTag(0 To m_vCap - 1): ReDim Preserve m_vClass(0 To m_vCap - 1)
    End If
    v = m_vn
    m_vn = m_vn + 1
    m_cVn(m_cn - 1) = m_cVn(m_cn - 1) + 1
    m_vTag(v) = tag
    m_vClass(v) = m_cn - 1
    m_vW(v) = 0: m_vH(v) = 0
    m_vR0(v) = m_rsn
    m_vCells(v) = 0
    m_vK1(v) = 0: m_vS1(v) = 0: m_vK2(v) = 0: m_vS2(v) = 0: m_vK3(v) = 0: m_vS3(v) = 0
    ac2fNCAddVariant = v

    If Not ac2fNCGeoBBox(x0, y0, x1, y1) Then Exit Function
    ox = x0 - m_d
    oy = y0 - m_d
    w = Int((x1 + m_d - ox) / m_res) + 1
    h = Int((y1 + m_d - oy) / m_res) + 1
    If w > NC_MAX_DIM Or h > NC_MAX_DIM Then Exit Function
    m_vOX(v) = ox
    m_vOY(v) = oy

    ac2fNCScratch
    ac2fNCInterior ox, oy, w, h, m_holes
    ac2fNCCapsules ox, oy, w, h, m_d, 1
    ac2fNCSortRows h, cnt, sa, sb, sf

    ' merged rows into the variant store
    Do While m_rsn + h + 1 > m_rsCap
        If m_rsCap = 0 Then m_rsCap = 4096 Else m_rsCap = m_rsCap * 2
        ReDim Preserve m_rs(0 To m_rsCap - 1)
    Loop
    Do While m_ivn + m_tn > m_ivCap
        If m_ivCap = 0 Then m_ivCap = 16384 Else m_ivCap = m_ivCap * 2
        ReDim Preserve m_ia(0 To m_ivCap - 1)
        ReDim Preserve m_ib(0 To m_ivCap - 1)
    Loop

    For r = 0 To h - 1
        m_rs(m_rsn + r) = m_ivn
        span = 0
        If cnt(r + 1) > cnt(r) Then
            a = sa(cnt(r)): b = sb(cnt(r))
            For t = cnt(r) + 1 To cnt(r + 1) - 1
                If sa(t) <= b Then
                    If sb(t) > b Then b = sb(t)
                Else
                    m_ia(m_ivn) = a: m_ib(m_ivn) = b: m_ivn = m_ivn + 1
                    cells = cells + b - a
                    If b - a > span Then span = b - a
                    a = sa(t): b = sb(t)
                End If
            Next t
            m_ia(m_ivn) = a: m_ib(m_ivn) = b: m_ivn = m_ivn + 1
            cells = cells + b - a
            If b - a > span Then span = b - a
        End If
        ' keep the three rows with the widest single interval
        If span > m_vS1(v) Then
            m_vK3(v) = m_vK2(v): m_vS3(v) = m_vS2(v)
            m_vK2(v) = m_vK1(v): m_vS2(v) = m_vS1(v)
            m_vK1(v) = r: m_vS1(v) = span
        ElseIf span > m_vS2(v) Then
            m_vK3(v) = m_vK2(v): m_vS3(v) = m_vS2(v)
            m_vK2(v) = r: m_vS2(v) = span
        ElseIf span > m_vS3(v) Then
            m_vK3(v) = r: m_vS3(v) = span
        End If
    Next r
    m_rs(m_rsn + h) = m_ivn
    m_rsn = m_rsn + h + 1

    If cells = 0 Then Exit Function
    m_vW(v) = w
    m_vH(v) = h
    m_vCells(v) = cells
End Function

Public Function ac2fNCVariantCount() As Long
    ac2fNCVariantCount = m_vn
End Function

Public Sub ac2fNCVariantInfo(ByVal v As Long, ByRef w As Long, ByRef h As Long, _
                             ByRef ox As Double, ByRef oy As Double, _
                             ByRef cells As Long, ByRef tag As Long)
    w = m_vW(v): h = m_vH(v)
    ox = m_vOX(v): oy = m_vOY(v)
    cells = m_vCells(v)
    tag = m_vTag(v)
End Sub

' Is cell (c, r) of variant v occupied? For tests and previews.
Public Function ac2fNCVariantCell(ByVal v As Long, ByVal c As Long, ByVal r As Long) As Boolean
    Dim t As Long
    If r < 0 Or r >= m_vH(v) Then Exit Function
    For t = m_rs(m_vR0(v) + r) To m_rs(m_vR0(v) + r + 1) - 1
        If c >= m_ia(t) And c < m_ib(t) Then
            ac2fNCVariantCell = True
            Exit Function
        End If
    Next t
End Function

Public Function ac2fNCAddPart(ByVal cls As Long) As Long
    If m_pn >= m_pCap Then
        If m_pCap = 0 Then m_pCap = 64 Else m_pCap = m_pCap * 2
        ReDim Preserve m_pClass(0 To m_pCap - 1)
    End If
    m_pClass(m_pn) = cls
    m_pn = m_pn + 1
    ac2fNCAddPart = m_pn - 1
End Function

Public Function ac2fNCPartCount() As Long
    ac2fNCPartCount = m_pn
End Function

'=====================================================================
' SHEETS
'=====================================================================

Private Function ac2fNCNewType() As Long
    If m_tyn >= m_tyCap Then
        If m_tyCap = 0 Then m_tyCap = 16 Else m_tyCap = m_tyCap * 2
        ReDim Preserve m_tyW(0 To m_tyCap - 1): ReDim Preserve m_tyH(0 To m_tyCap - 1)
        ReDim Preserve m_tyOX(0 To m_tyCap - 1): ReDim Preserve m_tyOY(0 To m_tyCap - 1)
        ReDim Preserve m_tyQty(0 To m_tyCap - 1): ReDim Preserve m_tyRect(0 To m_tyCap - 1)
        ReDim Preserve m_tyFree(0 To m_tyCap - 1): ReDim Preserve m_tyR0(0 To m_tyCap - 1)
    End If
    m_tyn = m_tyn + 1
    ac2fNCNewType = m_tyn - 1
End Function

' A rectangular sheet: the usable area inside the margins, in mm. Parts are
' grown by m_d, so the grid is grown by m_d on every side too and a part
' may touch the margin line exactly. qty < 0 = as many as needed.
' Returns the type index, or -1 when not even one cell fits.
Public Function ac2fNCAddRectSheet(ByVal usableW As Double, ByVal usableH As Double, _
                                   ByVal qty As Long) As Long
    Dim t As Long, w As Long, h As Long
    ac2fNCAddRectSheet = -1
    w = Int((usableW + 2# * m_d) / m_res + 0.000001)
    h = Int((usableH + 2# * m_d) / m_res + 0.000001)
    If w < 1 Or h < 1 Then Exit Function
    If w > NC_MAX_DIM Or h > NC_MAX_DIM Then Exit Function
    t = ac2fNCNewType()
    m_tyW(t) = w: m_tyH(t) = h
    m_tyOX(t) = -m_d: m_tyOY(t) = -m_d
    m_tyQty(t) = qty
    m_tyRect(t) = True
    m_tyFree(t) = w * h
    m_tyR0(t) = -1
    ac2fNCAddRectSheet = t
End Function

' A drawn sheet (an offcut of any shape) from the geometry buffer: its
' closed contours bound the material, contours inside it are holes in
' the sheet. Parts keep margin mm from every edge. Always one piece.
Public Function ac2fNCAddShapeSheet(ByVal margin As Double) As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim w As Long, h As Long, t As Long
    Dim cnt() As Long, sa() As Long, sb() As Long, sf() As Long
    Dim occ() As Byte
    Dim r As Long, c As Long, q As Long, cb0 As Long
    Dim rad As Double
    Dim freeN As Long

    ac2fNCAddShapeSheet = -1
    If Not ac2fNCGeoBBox(x0, y0, x1, y1) Then Exit Function
    w = Int((x1 - x0) / m_res) + 1
    h = Int((y1 - y0) / m_res) + 1
    If w > NC_MAX_DIM Or h > NC_MAX_DIM Then Exit Function

    rad = margin - m_d
    If rad < 0 Then rad = 0
    ac2fNCScratch
    ac2fNCInterior x0, y0, w, h, True
    ac2fNCCapsules x0, y0, w, h, rad, 1
    ac2fNCSortRows h, cnt, sa, sb, sf

    t = ac2fNCNewType()
    m_tyW(t) = w: m_tyH(t) = h
    m_tyOX(t) = x0: m_tyOY(t) = y0
    m_tyQty(t) = 1
    m_tyRect(t) = False
    m_tyR0(t) = m_tyRunN
    m_tyRunN = m_tyRunN + w * h
    If m_tyRunN = w * h Then
        ReDim m_tyRun(0 To m_tyRunN - 1)
    Else
        ReDim Preserve m_tyRun(0 To m_tyRunN - 1)
    End If

    ReDim occ(0 To w - 1)
    For r = 0 To h - 1
        For c = 0 To w - 1
            occ(c) = 1
        Next c
        ' inside first, then the margin band on top
        For q = cnt(r) To cnt(r + 1) - 1
            If sf(q) = 0 Then
                For c = sa(q) To sb(q) - 1
                    occ(c) = 0
                Next c
            End If
        Next q
        For q = cnt(r) To cnt(r + 1) - 1
            If sf(q) = 1 Then
                For c = sa(q) To sb(q) - 1
                    occ(c) = 1
                Next c
            End If
        Next q
        cb0 = m_tyR0(t) + r * w
        For c = w - 1 To 0 Step -1
            If occ(c) = 0 Then
                freeN = freeN + 1
                If c < w - 1 Then
                    If occ(c + 1) = 0 Then
                        m_tyRun(cb0 + c) = m_tyRun(cb0 + c + 1) + 1
                    Else
                        m_tyRun(cb0 + c) = 1
                    End If
                Else
                    m_tyRun(cb0 + c) = 1
                End If
            Else
                If c < w - 1 Then
                    If occ(c + 1) <> 0 Then
                        m_tyRun(cb0 + c) = m_tyRun(cb0 + c + 1) - 1
                    Else
                        m_tyRun(cb0 + c) = -1
                    End If
                Else
                    m_tyRun(cb0 + c) = -1
                End If
            End If
        Next c
    Next r
    m_tyFree(t) = freeN
    ac2fNCAddShapeSheet = t
End Function

Public Function ac2fNCSheetTypeCount() As Long
    ac2fNCSheetTypeCount = m_tyn
End Function

Public Sub ac2fNCSheetTypeInfo(ByVal t As Long, ByRef w As Long, ByRef h As Long, _
                               ByRef ox As Double, ByRef oy As Double, _
                               ByRef freeCells As Long)
    w = m_tyW(t): h = m_tyH(t)
    ox = m_tyOX(t): oy = m_tyOY(t)
    freeCells = m_tyFree(t)
End Sub

' Cells one sheet of type t needs in memory.
Public Function ac2fNCSheetCells(ByVal t As Long) As Double
    ac2fNCSheetCells = CDbl(m_tyW(t)) * CDbl(m_tyH(t))
End Function

' Lays out the sheet instances: the types in order, each as many times as
' its quantity allows, unlimited ones up to maxInst in total. Nothing is
' allocated until a sheet is first used.
Public Sub ac2fNCPrepare(ByVal maxInst As Long)
    Dim t As Long, i As Long, n As Long, q As Long
    Dim unl As Long, share As Long

    m_kn = 0
    For t = 0 To m_tyn - 1
        If m_tyQty(t) > 0 Then n = n + m_tyQty(t)
        If m_tyQty(t) < 0 Then unl = unl + 1
    Next t
    If n > maxInst Then n = maxInst
    share = 0
    If unl > 0 Then
        share = (maxInst - n) \ unl
        If share < 1 Then share = 1
    End If

    ReDim m_kType(0 To n + share * unl + 1)
    For t = 0 To m_tyn - 1
        q = m_tyQty(t)
        If q < 0 Then q = share
        For i = 1 To q
            If m_kn <= UBound(m_kType) Then
                m_kType(m_kn) = t
                m_kn = m_kn + 1
            End If
        Next i
    Next t

    ReDim m_kBase(0 To m_kn): ReDim m_kRow0(0 To m_kn)
    For i = 0 To m_kn
        m_kBase(i) = -1
        m_kRow0(i) = -1
    Next i
    m_runN = 0: m_mfN = 0
    If m_vn > 0 And m_kn > 0 Then
        ReDim m_memoY(0 To m_kn * m_vn - 1)
        ReDim m_memoX(0 To m_kn * m_vn - 1)
    End If
    ReDim m_kOpen(0 To m_kn): ReDim m_kSeq(0 To m_kn)
    ReDim m_kParts(0 To m_kn): ReDim m_kUsedH(0 To m_kn): ReDim m_kCells(0 To m_kn)
    If m_pn > 0 Then
        ReDim m_rK(0 To m_pn - 1): ReDim m_rV(0 To m_pn - 1)
        ReDim m_rX(0 To m_pn - 1): ReDim m_rY(0 To m_pn - 1)
        ReDim m_bK(0 To m_pn - 1): ReDim m_bV(0 To m_pn - 1)
        ReDim m_bX(0 To m_pn - 1): ReDim m_bY(0 To m_pn - 1)
        ReDim m_ord(0 To m_pn - 1): ReDim m_bOrd(0 To m_pn - 1)
    End If
    ReDim m_bKParts(0 To m_kn): ReDim m_bKUsedH(0 To m_kn)
    ReDim m_bKCells(0 To m_kn): ReDim m_bKSeq(0 To m_kn)

    m_anyRot = False
    For i = 0 To m_cn - 1
        If m_cVn(i) > 1 Then m_anyRot = True
    Next i
    m_haveBest = False
    m_step = 0
End Sub

Public Function ac2fNCInstanceCount() As Long
    ac2fNCInstanceCount = m_kn
End Function

Public Function ac2fNCInstanceType(ByVal k As Long) As Long
    ac2fNCInstanceType = m_kType(k)
End Function

' Gives instance k its memory and its empty-sheet runs, the first time it
' is used. The layout then stays for every later pass.
Private Sub ac2fNCAlloc(ByVal k As Long)
    Dim t As Long, w As Long, h As Long
    Dim r As Long, c As Long, cb0 As Long, src As Long
    Dim mx As Long, v As Long

    If m_kBase(k) >= 0 Then Exit Sub
    t = m_kType(k)
    w = m_tyW(t): h = m_tyH(t)
    m_kBase(k) = m_runN
    m_kRow0(k) = m_mfN
    m_runN = m_runN + w * h
    m_mfN = m_mfN + h
    If m_kBase(k) = 0 Then
        ReDim m_run(0 To m_runN - 1)
        ReDim m_runInit(0 To m_runN - 1)
    Else
        ReDim Preserve m_run(0 To m_runN - 1)
        ReDim Preserve m_runInit(0 To m_runN - 1)
    End If
    If m_kRow0(k) = 0 Then
        ReDim m_mf(0 To m_mfN - 1)
        ReDim m_mfInit(0 To m_mfN - 1)
    Else
        ReDim Preserve m_mf(0 To m_mfN - 1)
        ReDim Preserve m_mfInit(0 To m_mfN - 1)
    End If

    For r = 0 To h - 1
        cb0 = m_kBase(k) + r * w
        If m_tyRect(t) Then
            For c = 0 To w - 1
                m_run(cb0 + c) = w - c
                m_runInit(cb0 + c) = w - c
            Next c
            mx = w
        Else
            src = m_tyR0(t) + r * w
            mx = 0
            For c = 0 To w - 1
                v = m_tyRun(src + c)
                m_run(cb0 + c) = v
                m_runInit(cb0 + c) = v
                If v > mx Then mx = v
            Next c
        End If
        m_mf(m_kRow0(k) + r) = mx
        m_mfInit(m_kRow0(k) + r) = mx
    Next r
End Sub

'=====================================================================
' PACKER
'=====================================================================

' Leftmost free x >= xs for variant v at row y of instance k, or -1.
Private Function ac2fNCSlide(ByVal v As Long, ByVal k As Long, _
                             ByVal xs As Long, ByVal y As Long) As Long
    Dim x As Long, maxX As Long, w As Long, ph As Long
    Dim i As Long, ok As Long, j As Long, r0 As Long
    Dim rowBase As Long, a As Long, ln As Long, c As Long, idx As Long
    Dim run As Long, nx As Long
    Dim blocked As Boolean

    w = m_tyW(m_kType(k))
    ph = m_vH(v)
    maxX = w - m_vW(v)
    x = xs
    If x > maxX Then
        ac2fNCSlide = -1
        Exit Function
    End If
    r0 = m_vR0(v)
    i = 0
    ok = 0
    Do While ok < ph
        rowBase = m_kBase(k) + (y + i) * w
        blocked = False
        For j = m_rs(r0 + i) To m_rs(r0 + i + 1) - 1
            a = m_ia(j)
            ln = m_ib(j) - a
            c = x + a
            idx = rowBase + c
            run = m_run(idx)
            If run < ln Then
                If run > 0 Then
                    ' free, but too short: jump past the occupied run after it
                    nx = c + run - m_run(idx + run) - a
                Else
                    nx = c - run - a
                End If
                x = nx
                blocked = True
                Exit For
            End If
        Next j
        If blocked Then
            If x > maxX Then
                ac2fNCSlide = -1
                Exit Function
            End If
            ok = 0
        Else
            ok = ok + 1
            i = i + 1
            If i = ph Then i = 0
        End If
    Loop
    ac2fNCSlide = x
End Function

' Lowest (then leftmost) free place for variant v on instance k with
' y <= yMax. Uses and updates the memo of where the last search ended.
Private Function ac2fNCSearch(ByVal v As Long, ByVal k As Long, ByVal yMax As Long, _
                              ByRef rx As Long, ByRef ry As Long) As Boolean
    Dim t As Long, w As Long, h As Long, ph As Long
    Dim m As Long, y As Long, y0 As Long, yEnd As Long, x As Long, xs As Long
    Dim mr As Long
    Dim k1 As Long, s1 As Long, k2 As Long, s2 As Long, k3 As Long, s3 As Long

    t = m_kType(k)
    w = m_tyW(t): h = m_tyH(t)
    ph = m_vH(v)
    If m_vW(v) = 0 Or m_vW(v) > w Or ph > h Then Exit Function
    m = k * m_vn + v
    y0 = m_memoY(m)
    If y0 < 0 Then Exit Function
    yEnd = h - ph
    If yMax < yEnd Then yEnd = yMax
    If y0 > yEnd Then Exit Function

    mr = m_kRow0(k)
    k1 = m_vK1(v): s1 = m_vS1(v)
    k2 = m_vK2(v): s2 = m_vS2(v)
    k3 = m_vK3(v): s3 = m_vS3(v)
    xs = m_memoX(m)
    For y = y0 To yEnd
        If m_mf(mr + y + k1) >= s1 Then
            If m_mf(mr + y + k2) >= s2 Then
                If m_mf(mr + y + k3) >= s3 Then
                    x = ac2fNCSlide(v, k, xs, y)
                    If x >= 0 Then
                        m_memoY(m) = y
                        m_memoX(m) = x
                        rx = x
                        ry = y
                        ac2fNCSearch = True
                        Exit Function
                    End If
                End If
            End If
        End If
        xs = 0
    Next y

    If yEnd = h - ph Then
        m_memoY(m) = -1
    Else
        m_memoY(m) = yEnd + 1
        m_memoX(m) = 0
    End If
End Function

' Marks [a, b) of one row occupied and repairs the free runs to its left.
' An occupied cell holds -k: "at least k cells from here are occupied".
' Sheets only fill up during a pass, so a stale k stays a true lower
' bound; the slide just jumps again from where it lands. Free runs on
' the other hand must be exact, and only those to the left of a change.
Private Sub ac2fNCFillRow(ByVal rowBase As Long, ByVal a As Long, ByVal b As Long, _
                          ByVal w As Long)
    Dim c As Long, k As Long, rt As Long

    k = 0
    If b < w Then
        rt = m_run(rowBase + b)
        If rt < 0 Then k = -rt
    End If
    For c = b - 1 To a Step -1
        k = k + 1
        m_run(rowBase + c) = -k
    Next c
    k = 0
    For c = a - 1 To 0 Step -1
        If m_run(rowBase + c) < 0 Then Exit For
        k = k + 1
        m_run(rowBase + c) = k
    Next c
End Sub

Private Sub ac2fNCPlace(ByVal v As Long, ByVal k As Long, ByVal x As Long, ByVal y As Long)
    Dim w As Long, ph As Long, i As Long, j As Long, r0 As Long
    Dim rowBase As Long, c As Long, run As Long, mx As Long

    w = m_tyW(m_kType(k))
    ph = m_vH(v)
    r0 = m_vR0(v)
    For i = 0 To ph - 1
        rowBase = m_kBase(k) + (y + i) * w
        For j = m_rs(r0 + i + 1) - 1 To m_rs(r0 + i) Step -1
            ac2fNCFillRow rowBase, x + m_ia(j), x + m_ib(j), w
        Next j
        ' widest free run left in the row
        mx = 0
        c = 0
        Do While c < w
            run = m_run(rowBase + c)
            If run > 0 Then
                If run > mx Then mx = run
                c = c + run
            Else
                c = c - run
            End If
        Loop
        m_mf(m_kRow0(k) + y + i) = mx
    Next i
    If y + ph > m_kUsedH(k) Then m_kUsedH(k) = y + ph
    m_kParts(k) = m_kParts(k) + 1
    m_kCells(k) = m_kCells(k) + m_vCells(v)
End Sub

' Best place for part p on instance k over its allowed variants.
Private Function ac2fNCTry(ByVal p As Long, ByVal k As Long, ByVal strat As Long, _
                           ByVal rotAll As Boolean) As Boolean
    Dim c As Long, v As Long, v0 As Long, v1 As Long
    Dim bestV As Long, bestX As Long, bestY As Long, bestKey As Long
    Dim x As Long, y As Long, key As Long, yMax As Long
    Dim h As Long

    c = m_pClass(p)
    If m_cVn(c) = 0 Then Exit Function
    v0 = m_cV0(c)
    v1 = v0
    If rotAll Then v1 = v0 + m_cVn(c) - 1
    h = m_tyH(m_kType(k))
    bestV = -1
    For v = v0 To v1
        If m_vW(v) > 0 Then
            yMax = h
            If bestV >= 0 Then
                If strat = 1 Then yMax = bestKey - m_vH(v) Else yMax = bestKey
            End If
            If yMax >= 0 Then
                If ac2fNCSearch(v, k, yMax, x, y) Then
                    If strat = 1 Then key = y + m_vH(v) Else key = y
                    If bestV < 0 Then
                        bestV = v: bestX = x: bestY = y: bestKey = key
                    ElseIf key < bestKey Or (key = bestKey And x < bestX) Then
                        bestV = v: bestX = x: bestY = y: bestKey = key
                    End If
                End If
            End If
        End If
    Next v
    If bestV < 0 Then Exit Function

    ac2fNCPlace bestV, k, bestX, bestY
    m_rK(p) = k: m_rV(p) = bestV: m_rX(p) = bestX: m_rY(p) = bestY
    ac2fNCTry = True
End Function

' Could part p fit an empty rectangular sheet of type t at all?
Private Function ac2fNCFitsEmpty(ByVal p As Long, ByVal t As Long, ByVal rotAll As Boolean) As Boolean
    Dim c As Long, v As Long, v1 As Long
    c = m_pClass(p)
    If m_cVn(c) = 0 Then Exit Function
    If Not m_tyRect(t) Then
        ac2fNCFitsEmpty = True
        Exit Function
    End If
    v1 = m_cV0(c)
    If rotAll Then v1 = m_cV0(c) + m_cVn(c) - 1
    For v = m_cV0(c) To v1
        If m_vW(v) > 0 And m_vW(v) <= m_tyW(t) And m_vH(v) <= m_tyH(t) Then
            ac2fNCFitsEmpty = True
            Exit Function
        End If
    Next v
End Function

' One complete pass over the parts in the order m_ord.
Private Sub ac2fNCPass(ByVal strat As Long, ByVal rotAll As Boolean)
    Dim i As Long, p As Long, k As Long, t As Long, last As Long
    Dim placed As Boolean
    Dim tried() As Boolean

    If m_runN > 0 Then
        m_run = m_runInit
        m_mf = m_mfInit
    End If
    If m_kn > 0 And m_vn > 0 Then
        ReDim m_memoY(0 To m_kn * m_vn - 1)
        ReDim m_memoX(0 To m_kn * m_vn - 1)
    End If
    For k = 0 To m_kn
        m_kOpen(k) = False
        m_kSeq(k) = 0
        m_kParts(k) = 0
        m_kUsedH(k) = 0
        m_kCells(k) = 0
    Next k
    m_seq = 0
    ReDim tried(0 To m_tyn)

    For i = 0 To m_pn - 1
        p = m_ord(i)
        m_rK(p) = -1: m_rV(p) = -1: m_rX(p) = 0: m_rY(p) = 0
        placed = False
        ' first fit over the sheets already open, in their order
        For k = 0 To m_kn - 1
            If m_kOpen(k) Then
                If ac2fNCTry(p, k, strat, rotAll) Then
                    placed = True
                    Exit For
                End If
            End If
        Next k
        If Not placed Then
            ' open the next sheet that can take it: one try per sheet size
            ' (a part that misses an empty sheet misses every copy of it)
            For t = 0 To m_tyn - 1
                tried(t) = False
            Next t
            For k = 0 To m_kn - 1
                If Not m_kOpen(k) Then
                    t = m_kType(k)
                    If Not (m_tyRect(t) And tried(t)) Then
                        tried(t) = True
                        If ac2fNCFitsEmpty(p, t, rotAll) Then
                            ac2fNCAlloc k
                            m_kOpen(k) = True
                            m_seq = m_seq + 1
                            m_kSeq(k) = m_seq
                            If ac2fNCTry(p, k, strat, rotAll) Then
                                placed = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next k
        End If
    Next i

    ' score
    m_curU = 0
    For p = 0 To m_pn - 1
        If m_rK(p) < 0 Then m_curU = m_curU + 1
    Next p
    last = -1
    For k = 0 To m_kn - 1
        If m_kParts(k) > 0 Then
            If last < 0 Then
                last = k
            ElseIf m_kSeq(k) > m_kSeq(last) Then
                last = k
            End If
        End If
    Next k
    m_curMat = 0: m_curTie = 0
    For k = 0 To m_kn - 1
        If m_kParts(k) > 0 Then
            t = m_kType(k)
            If k = last Then
                m_curMat = m_curMat + CDbl(m_tyW(t)) * CDbl(m_kUsedH(k))
            Else
                m_curMat = m_curMat + CDbl(m_tyW(t)) * CDbl(m_tyH(t))
            End If
            m_curTie = m_curTie + CDbl(m_kUsedH(k))
        End If
    Next k
End Sub

Private Function ac2fNCBetter() As Boolean
    If Not m_haveBest Then
        ac2fNCBetter = True
    ElseIf m_curU <> m_bU Then
        ac2fNCBetter = (m_curU < m_bU)
    ElseIf Abs(m_curMat - m_bMat) > 0.5 Then
        ac2fNCBetter = (m_curMat < m_bMat)
    Else
        ac2fNCBetter = (m_curTie < m_bTie - 0.5)
    End If
End Function

Private Sub ac2fNCKeep(ByVal strat As Long, ByVal rotAll As Boolean)
    Dim i As Long
    For i = 0 To m_pn - 1
        m_bK(i) = m_rK(i): m_bV(i) = m_rV(i)
        m_bX(i) = m_rX(i): m_bY(i) = m_rY(i)
        m_bOrd(i) = m_ord(i)
    Next i
    For i = 0 To m_kn
        m_bKParts(i) = m_kParts(i)
        m_bKUsedH(i) = m_kUsedH(i)
        m_bKCells(i) = m_kCells(i)
        m_bKSeq(i) = m_kSeq(i)
    Next i
    m_bU = m_curU: m_bMat = m_curMat: m_bTie = m_curTie
    m_bStrat = strat: m_bRot = rotAll
    m_bStep = m_step
    m_haveBest = True
End Sub

'=====================================================================
' OPTIMISER
'=====================================================================

' Fills m_ord sorted descending by key(class) of the chosen kind:
' 0 area, 1 longest side, 2 height, 3 width (all of the first variant).
Private Sub ac2fNCOrderBy(ByVal kind As Long)
    Dim i As Long, j As Long, p As Long
    Dim key() As Double
    Dim kp As Double
    Dim c As Long, v As Long

    If m_pn = 0 Then Exit Sub
    ReDim key(0 To m_pn - 1)
    For p = 0 To m_pn - 1
        c = m_pClass(p)
        key(p) = -1
        If m_cVn(c) > 0 Then
            v = m_cV0(c)
            Select Case kind
                Case 0: key(p) = m_vCells(v)
                Case 1
                    key(p) = m_vW(v)
                    If m_vH(v) > key(p) Then key(p) = m_vH(v)
                Case 2: key(p) = m_vH(v)
                Case Else: key(p) = m_vW(v)
            End Select
        End If
        ' keep the copies of one class together
        key(p) = key(p) * 100000# - c
    Next p
    For i = 0 To m_pn - 1
        m_ord(i) = i
    Next i
    For i = 1 To m_pn - 1
        p = m_ord(i)
        kp = key(p)
        j = i - 1
        Do While j >= 0
            If key(m_ord(j)) >= kp Then Exit Do
            m_ord(j + 1) = m_ord(j)
            j = j - 1
        Loop
        m_ord(j + 1) = p
    Next i
End Sub

Private Function ac2fNCRand(ByVal n As Long) As Long
    Dim r As Long
    ' Park-Miller, exact in Double
    m_seed = m_seed * 16807#
    m_seed = m_seed - Int(m_seed / 2147483647#) * 2147483647#
    If n > 1 Then
        r = Int(m_seed / 2147483647# * n)
        If r >= n Then r = n - 1
    End If
    ac2fNCRand = r
End Function

Public Sub ac2fNCOptBegin(ByVal seed As Long)
    m_step = 0
    m_seed = seed
    If m_seed <= 0 Then m_seed = 12345
    m_haveBest = False
End Sub

' How many fixed passes ac2fNCOptStep runs before it starts mutating.
Public Function ac2fNCOptFixedSteps() As Long
    If m_anyRot Then ac2fNCOptFixedSteps = 6 Else ac2fNCOptFixedSteps = 5
End Function

' Runs one pass. Returns True when it beat the best so far.
Public Function ac2fNCOptStep() As Boolean
    Dim strat As Long, rotAll As Boolean
    Dim i As Long, j As Long, t As Long, n As Long, span As Long
    Dim guard As Long

    If m_pn = 0 Or m_kn = 0 Then Exit Function
    ' with one strategy fixed, the second fixed pass would only repeat the first
    If m_step = 1 And m_force > 0 Then m_step = 2
    strat = 1
    rotAll = True
    Select Case m_step
        Case 0: ac2fNCOrderBy 0
        Case 1
            ac2fNCOrderBy 0
            strat = 2
        Case 2: ac2fNCOrderBy 1
        Case 3: ac2fNCOrderBy 2
        Case 4: ac2fNCOrderBy 3
        Case Else
            If m_step = 5 And m_anyRot Then
                ac2fNCOrderBy 0
                rotAll = False
            Else
                ' mutate the best order: swap two parts of different
                ' classes, mostly close together
                strat = m_bStrat
                rotAll = m_bRot
                For i = 0 To m_pn - 1
                    m_ord(i) = m_bOrd(i)
                Next i
                n = m_pn
                If n >= 2 Then
                    t = 1 + ac2fNCRand(3)
                    Do While t > 0
                        guard = 0
                        Do
                            i = ac2fNCRand(n)
                            span = 1 + ac2fNCRand(8)
                            If ac2fNCRand(4) = 0 Then span = 1 + ac2fNCRand(n)
                            j = i + span
                            If j >= n Then j = i - span
                            If j < 0 Then j = ac2fNCRand(n)
                            guard = guard + 1
                            If guard > 50 Then Exit Do
                        Loop While m_pClass(m_ord(i)) = m_pClass(m_ord(j))
                        If i <> j Then
                            span = m_ord(i): m_ord(i) = m_ord(j): m_ord(j) = span
                        End If
                        t = t - 1
                    Loop
                End If
                If ac2fNCRand(10) = 0 Then strat = 3 - strat
            End If
    End Select

    If m_force > 0 Then strat = m_force
    ac2fNCPass strat, rotAll
    If ac2fNCBetter() Then
        ac2fNCKeep strat, rotAll
        ac2fNCOptStep = True
    End If
    m_step = m_step + 1
End Function

Public Function ac2fNCOptStepsDone() As Long
    ac2fNCOptStepsDone = m_step
End Function

'=====================================================================
' RESULTS  (of the best pass)
'=====================================================================

Public Function ac2fNCHaveResult() As Boolean
    ac2fNCHaveResult = m_haveBest
End Function

' Where part p went: instance (-1 = did not fit), variant and cell.
Public Function ac2fNCResult(ByVal p As Long, ByRef v As Long, _
                             ByRef cx As Long, ByRef cy As Long) As Long
    If Not m_haveBest Then
        ac2fNCResult = -1
        Exit Function
    End If
    v = m_bV(p): cx = m_bX(p): cy = m_bY(p)
    ac2fNCResult = m_bK(p)
End Function

Public Sub ac2fNCInstanceUse(ByVal k As Long, ByRef parts As Long, ByRef usedRows As Long, _
                             ByRef cells As Long, ByRef seq As Long)
    parts = m_bKParts(k)
    usedRows = m_bKUsedH(k)
    cells = m_bKCells(k)
    seq = m_bKSeq(k)
End Sub

Public Function ac2fNCBestUnplaced() As Long
    ac2fNCBestUnplaced = m_bU
End Function

' Material of the best pass in cells (see the module header).
Public Function ac2fNCBestMaterial() As Double
    ac2fNCBestMaterial = m_bMat
End Function

Public Function ac2fNCBestStrategy() As Long
    ac2fNCBestStrategy = m_bStrat
End Function

Public Function ac2fNCBestRotates() As Boolean
    ac2fNCBestRotates = m_bRot
End Function

Public Function ac2fNCBestStep() As Long
    ac2fNCBestStep = m_bStep
End Function
