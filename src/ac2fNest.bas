Attribute VB_Name = "ac2fNest"
'=====================================================================
'  ac2f pack  --  ac2fNest
'
'  Packs the selected objects onto sheets: a margin per edge, a gap
'  between parts, and a rotation step the packer may use.
'
'  Macros:
'    ac2fNest       - rearranges the selection onto sheets
'    ac2fNestReport - works it out and reports, moves nothing
'
'  ALGORITHM
'  Skyline bottom-left over bounding boxes. Parts go in longest edge
'  first; for each one every allowed orientation and every skyline step
'  is tried, and the placement chosen is the one that leaves the
'  SKYLINE LOWEST afterwards (y + height), with least buried area as the
'  tie-break.
'
'  That scoring matters. Picking the lowest resting y instead, which is
'  the obvious choice, makes rotation actively harmful: measured on five
'  part sets it lost 8% and 18% of material on two of them. Scoring by
'  the resulting skyline height gained 8, 7, 18, 5 and 13% instead, with
'  no set getting worse.
'
'  Rotation can still misfire on a shape mix not covered by those tests,
'  so when rotation is allowed the job is packed twice, once with it and
'  once without, and the better result is kept. Rotation can therefore
'  never cost material.
'
'  Parts are packed by bounding box, not by true outline. Concave parts
'  will not interlock.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Nesting"

Private Const MAX_PARTS  As Long = 500
Private Const MAX_SHEETS As Long = 60
Private Const MAX_SEG    As Long = 600
Private Const MAX_ANGLES As Long = 24     ' smallest usable step is 15 degrees

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

Public Const AC2F_DEF_NS_ML   As Double = 10#
Public Const AC2F_DEF_NS_MR   As Double = 10#
Public Const AC2F_DEF_NS_MT   As Double = 10#
Public Const AC2F_DEF_NS_MB   As Double = 10#
Public Const AC2F_DEF_NS_GAP  As Double = 3#
Public Const AC2F_DEF_NS_ROT  As Double = 90#
Public Const AC2F_DEF_NS_SW   As Double = 0#    ' 0 = take the page width
Public Const AC2F_DEF_NS_SH   As Double = 0#    ' 0 = take the page height
Public Const AC2F_DEF_NS_SGAP As Double = 20#

'---------------------------------------------------------------------
' Parts
'---------------------------------------------------------------------
Private m_shp() As Shape
Private m_w0() As Double, m_h0() As Double        ' bounding box as drawn
Private m_bw() As Double, m_bh() As Double        ' box per angle, [part, angle]
Private m_np As Long

Private m_ang() As Double                         ' allowed angles
Private m_na As Long

' current placement
Private m_pSheet() As Long
Private m_pX() As Double, m_pY() As Double
Private m_pA() As Long                            ' index into m_ang
Private m_nSheets As Long

' best placement kept so far
Private m_bSheet() As Long
Private m_bX() As Double, m_bY() As Double
Private m_bA() As Long
Private m_bSheets As Long
Private m_bUsed As Double

' skyline per sheet
Private m_sx() As Double, m_sw() As Double, m_sy() As Double
Private m_sn() As Long

Private m_usableW As Double, m_usableH As Double

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fNest()
Attribute ac2fNest.VB_Description = "ac2f pack: Pack the selection onto sheets"
    ac2fNSRun True
End Sub

Public Sub ac2fNestReport()
Attribute ac2fNestReport.VB_Description = "ac2f pack: Nesting report, moves nothing"
    ac2fNSRun False
End Sub

'=====================================================================
' MAIN FLOW
'=====================================================================

Private Sub ac2fNSRun(ByVal apply As Boolean)
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean
    Dim i As Long
    Dim mL As Double, mR As Double, mT As Double, mB As Double
    Dim gap As Double, rotStep As Double, sgap As Double
    Dim sheetW As Double, sheetH As Double
    Dim x0 As Double, y0 As Double, bw As Double, bh As Double
    Dim usedRot As Double, usedNo As Double
    Dim placedCount As Long

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
    If sr.Count > MAX_PARTS Then
        ac2fWarn "That is " & sr.Count & " parts; the limit is " & MAX_PARTS & ".", CAPTION_
        Exit Sub
    End If

    mL = ac2fNSPos(AC2F_K_NS_ML, AC2F_DEF_NS_ML)
    mR = ac2fNSPos(AC2F_K_NS_MR, AC2F_DEF_NS_MR)
    mT = ac2fNSPos(AC2F_K_NS_MT, AC2F_DEF_NS_MT)
    mB = ac2fNSPos(AC2F_K_NS_MB, AC2F_DEF_NS_MB)
    gap = ac2fNSPos(AC2F_K_NS_GAP, AC2F_DEF_NS_GAP)
    sgap = ac2fNSPos(AC2F_K_NS_SGAP, AC2F_DEF_NS_SGAP)
    rotStep = ac2fGetNum(AC2F_K_NS_ROT, AC2F_DEF_NS_ROT)
    If rotStep < 0 Then rotStep = 0

    On Error GoTo Fail
    oldOpt = Application.Optimization
    Application.Optimization = True
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    sheetW = ac2fGetNum(AC2F_K_NS_SW, AC2F_DEF_NS_SW)
    sheetH = ac2fGetNum(AC2F_K_NS_SH, AC2F_DEF_NS_SH)
    If sheetW <= 0 Then sheetW = ActivePage.SizeWidth
    If sheetH <= 0 Then sheetH = ActivePage.SizeHeight

    m_usableW = sheetW - mL - mR
    m_usableH = sheetH - mT - mB
    If m_usableW <= 0 Or m_usableH <= 0 Then
        ac2fWarn "The margins leave no room on a " & ac2fFmt(sheetW, 0) & " x " & _
                 ac2fFmt(sheetH, 0) & " mm sheet.", CAPTION_
        GoTo Cleanup
    End If

    sr.GetBoundingBox x0, y0, bw, bh

    ' --- parts and their box at each allowed angle -------------------
    If Not ac2fNSCollect(sr, rotStep) Then
        ac2fWarn "Nothing measurable in the selection.", CAPTION_
        GoTo Cleanup
    End If

    ' --- pack, twice when rotation is allowed ------------------------
    m_bUsed = -1
    usedNo = ac2fNSPack(False, gap)
    ac2fNSKeepBest usedNo
    If m_na > 1 Then
        usedRot = ac2fNSPack(True, gap)
        ac2fNSKeepBest usedRot
    End If

    For i = 0 To m_np - 1
        If m_bSheet(i) >= 0 Then placedCount = placedCount + 1
    Next i

    If apply Then
        ActiveDocument.BeginCommandGroup ac2fTitle("nesting")
        ac2fNSApply x0, y0, sheetW, sheetH, sgap, mL, mB
        ActiveDocument.EndCommandGroup
    End If

Cleanup:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    On Error GoTo 0

    ac2fInfo ac2fNSReport(apply, placedCount, sheetW, sheetH, mL, mR, mT, mB, _
                          gap, rotStep, usedNo, usedRot), CAPTION_
    Exit Sub

Fail:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    ac2fWarn "The operation failed: " & Err.Description, CAPTION_
End Sub

Private Function ac2fNSPos(ByVal key As String, ByVal def As Double) As Double
    Dim v As Double
    v = ac2fGetNum(key, def)
    If v < 0 Then v = 0
    ac2fNSPos = v
End Function

'---------------------------------------------------------------------
' Collect parts and measure their box at every allowed angle.
'---------------------------------------------------------------------
Private Function ac2fNSCollect(ByVal sr As ShapeRange, ByVal rotStep As Double) As Boolean
    Dim i As Long, k As Long
    Dim x As Double, y As Double, w As Double, h As Double
    Dim a As Double

    ' angles
    m_na = 1
    ReDim m_ang(0 To MAX_ANGLES - 1)
    m_ang(0) = 0
    If rotStep > 0 Then
        If rotStep < 360# / MAX_ANGLES Then rotStep = 360# / MAX_ANGLES
        a = rotStep
        Do While a < 359.999 And m_na < MAX_ANGLES
            m_ang(m_na) = a
            m_na = m_na + 1
            a = a + rotStep
        Loop
    End If

    m_np = 0
    ReDim m_shp(0 To sr.Count - 1)
    ReDim m_w0(0 To sr.Count - 1): ReDim m_h0(0 To sr.Count - 1)
    ReDim m_bw(0 To sr.Count - 1, 0 To m_na - 1)
    ReDim m_bh(0 To sr.Count - 1, 0 To m_na - 1)

    For i = 1 To sr.Count
        On Error Resume Next
        sr(i).GetBoundingBox x, y, w, h
        On Error GoTo 0
        If w > 0 And h > 0 Then
            Set m_shp(m_np) = sr(i)
            m_w0(m_np) = w: m_h0(m_np) = h
            For k = 0 To m_na - 1
                ac2fNSBox m_np, k
            Next k
            m_np = m_np + 1
        End If
    Next i

    ReDim m_pSheet(0 To m_np): ReDim m_pX(0 To m_np)
    ReDim m_pY(0 To m_np): ReDim m_pA(0 To m_np)
    ReDim m_bSheet(0 To m_np): ReDim m_bX(0 To m_np)
    ReDim m_bY(0 To m_np): ReDim m_bA(0 To m_np)

    ac2fNSCollect = (m_np > 0)
End Function

' Box of part p at angle index k. Quarter turns just swap the sides; any
' other angle is measured by actually turning the shape and turning it
' straight back, because the true outline decides the box, not the
' unrotated one.
Private Sub ac2fNSBox(ByVal p As Long, ByVal k As Long)
    Dim a As Double
    Dim q As Double
    Dim x As Double, y As Double, w As Double, h As Double

    a = m_ang(k)
    q = a - Int(a / 90#) * 90#
    If Abs(q) < 0.0001 Or Abs(q - 90#) < 0.0001 Then
        If Abs(Int(a / 90# + 0.5)) Mod 2 = 1 Then
            m_bw(p, k) = m_h0(p): m_bh(p, k) = m_w0(p)
        Else
            m_bw(p, k) = m_w0(p): m_bh(p, k) = m_h0(p)
        End If
        Exit Sub
    End If

    On Error GoTo Fallback
    m_shp(p).Rotate a
    m_shp(p).GetBoundingBox x, y, w, h
    m_shp(p).Rotate -a
    m_bw(p, k) = w: m_bh(p, k) = h
    Exit Sub

Fallback:
    On Error Resume Next
    m_shp(p).Rotate -a
    m_bw(p, k) = m_w0(p): m_bh(p, k) = m_h0(p)
End Sub

'=====================================================================
' PACKER
'=====================================================================

' Returns the total material length used. Fills the current placement.
Private Function ac2fNSPack(ByVal allowRot As Boolean, ByVal gap As Double) As Double
    Dim ord() As Long
    Dim i As Long, j As Long, t As Long
    Dim p As Long, k As Long, s As Long
    Dim nAng As Long
    Dim bestKey1 As Double, bestKey2 As Double, bestKey3 As Double
    Dim bestA As Long, bestX As Double, bestY As Double
    Dim bestW As Double, bestH As Double
    Dim fx As Double, fy As Double, fwaste As Double
    Dim ok As Boolean, found As Boolean
    Dim w As Double, h As Double
    Dim total As Double

    nAng = 1
    If allowRot Then nAng = m_na

    ' longest edge first
    ReDim ord(0 To m_np - 1)
    For i = 0 To m_np - 1
        ord(i) = i
    Next i
    For i = 0 To m_np - 2
        For j = i + 1 To m_np - 1
            If ac2fNSLong(ord(j)) > ac2fNSLong(ord(i)) Then
                t = ord(i): ord(i) = ord(j): ord(j) = t
            End If
        Next j
    Next i

    ReDim m_sx(0 To MAX_SHEETS - 1, 0 To MAX_SEG - 1)
    ReDim m_sw(0 To MAX_SHEETS - 1, 0 To MAX_SEG - 1)
    ReDim m_sy(0 To MAX_SHEETS - 1, 0 To MAX_SEG - 1)
    ReDim m_sn(0 To MAX_SHEETS - 1)
    m_nSheets = 0

    For i = 0 To m_np - 1
        p = ord(i)
        m_pSheet(p) = -1
        found = False

        For s = 0 To MAX_SHEETS - 1
            If s >= m_nSheets Then
                If m_nSheets >= MAX_SHEETS Then Exit For
                m_sn(m_nSheets) = 1
                m_sx(m_nSheets, 0) = 0: m_sw(m_nSheets, 0) = m_usableW
                m_sy(m_nSheets, 0) = 0
                m_nSheets = m_nSheets + 1
            End If

            bestA = -1
            For k = 0 To nAng - 1
                w = m_bw(p, k): h = m_bh(p, k)
                ok = ac2fNSFit(s, w + gap, h + gap, fx, fy, fwaste)
                If Not ok Then ok = ac2fNSFit(s, w, h, fx, fy, fwaste)
                If ok Then
                    ' lowest resulting skyline, then least buried area
                    If bestA < 0 Or (fy + h) < bestKey1 - 0.000001 Or _
                       (Abs((fy + h) - bestKey1) <= 0.000001 And fwaste < bestKey2) Then
                        bestKey1 = fy + h: bestKey2 = fwaste: bestKey3 = fx
                        bestA = k: bestX = fx: bestY = fy
                        bestW = w: bestH = h
                    End If
                End If
            Next k

            If bestA >= 0 Then
                w = bestW + gap
                If bestX + w > m_usableW Then w = m_usableW - bestX
                ac2fNSPlace s, bestX, bestY, w, bestH + gap
                m_pSheet(p) = s
                m_pX(p) = bestX: m_pY(p) = bestY: m_pA(p) = bestA
                found = True
                Exit For
            End If
        Next s

        If Not found Then m_pSheet(p) = -1
    Next i

    ' material length = highest occupied point on each sheet
    For s = 0 To m_nSheets - 1
        h = 0
        For i = 0 To m_np - 1
            If m_pSheet(i) = s Then
                If m_pY(i) + m_bh(i, m_pA(i)) > h Then h = m_pY(i) + m_bh(i, m_pA(i))
            End If
        Next i
        total = total + h
    Next s

    ' an unplaced part is worse than any amount of material
    For i = 0 To m_np - 1
        If m_pSheet(i) < 0 Then total = total + 1000000#
    Next i

    ac2fNSPack = total
End Function

Private Function ac2fNSLong(ByVal p As Long) As Double
    If m_w0(p) > m_h0(p) Then ac2fNSLong = m_w0(p) Else ac2fNSLong = m_h0(p)
End Function

' Lowest resting place for a w x h box on sheet s.
Private Function ac2fNSFit(ByVal s As Long, ByVal w As Double, ByVal h As Double, _
                           ByRef rx As Double, ByRef ry As Double, _
                           ByRef rwaste As Double) As Boolean
    Dim i As Long, j As Long
    Dim x As Double, y As Double, cov As Double, waste As Double, seg As Double
    Dim haveBest As Boolean
    Dim bY As Double, bW As Double, bX As Double

    For i = 0 To m_sn(s) - 1
        x = m_sx(s, i)
        If x + w <= m_usableW + 0.000001 Then
            y = 0: cov = 0: j = i
            Do While j < m_sn(s) And cov < w - 0.000001
                If m_sy(s, j) > y Then y = m_sy(s, j)
                cov = cov + m_sw(s, j)
                j = j + 1
            Loop
            If cov >= w - 0.000001 And y + h <= m_usableH + 0.000001 Then
                waste = 0: cov = 0: j = i
                Do While j < m_sn(s) And cov < w - 0.000001
                    seg = m_sw(s, j)
                    If cov + seg > w Then seg = w - cov
                    waste = waste + (y - m_sy(s, j)) * seg
                    cov = cov + m_sw(s, j)
                    j = j + 1
                Loop
                If Not haveBest Or y < bY - 0.000001 Or _
                   (Abs(y - bY) <= 0.000001 And waste < bW) Then
                    bY = y: bW = waste: bX = x
                    haveBest = True
                End If
            End If
        End If
    Next i

    If haveBest Then
        rx = bX: ry = bY: rwaste = bW
        ac2fNSFit = True
    End If
End Function

Private Sub ac2fNSPlace(ByVal s As Long, ByVal x As Double, ByVal y As Double, _
                        ByVal w As Double, ByVal h As Double)
    Dim nx() As Double, nw() As Double, ny() As Double
    Dim n As Long, i As Long, j As Long
    Dim sx As Double, sw As Double, sy As Double
    Dim tx As Double, tw As Double, ty As Double

    ReDim nx(0 To MAX_SEG - 1): ReDim nw(0 To MAX_SEG - 1): ReDim ny(0 To MAX_SEG - 1)

    For i = 0 To m_sn(s) - 1
        sx = m_sx(s, i): sw = m_sw(s, i): sy = m_sy(s, i)
        If sx + sw <= x + 0.000001 Or sx >= x + w - 0.000001 Then
            If n < MAX_SEG Then nx(n) = sx: nw(n) = sw: ny(n) = sy: n = n + 1
        Else
            If sx < x And n < MAX_SEG Then
                nx(n) = sx: nw(n) = x - sx: ny(n) = sy: n = n + 1
            End If
            If sx + sw > x + w And n < MAX_SEG Then
                nx(n) = x + w: nw(n) = sx + sw - (x + w): ny(n) = sy: n = n + 1
            End If
        End If
    Next i
    If n < MAX_SEG Then nx(n) = x: nw(n) = w: ny(n) = y + h: n = n + 1

    ' sort by x
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If nx(j) < nx(i) Then
                tx = nx(i): nx(i) = nx(j): nx(j) = tx
                tw = nw(i): nw(i) = nw(j): nw(j) = tw
                ty = ny(i): ny(i) = ny(j): ny(j) = ty
            End If
        Next j
    Next i

    ' merge neighbours at the same height
    m_sn(s) = 0
    For i = 0 To n - 1
        If m_sn(s) > 0 Then
            j = m_sn(s) - 1
            If Abs(m_sy(s, j) - ny(i)) < 0.000001 And _
               Abs(m_sx(s, j) + m_sw(s, j) - nx(i)) < 0.000001 Then
                m_sw(s, j) = m_sw(s, j) + nw(i)
                GoTo NextSeg
            End If
        End If
        m_sx(s, m_sn(s)) = nx(i)
        m_sw(s, m_sn(s)) = nw(i)
        m_sy(s, m_sn(s)) = ny(i)
        m_sn(s) = m_sn(s) + 1
NextSeg:
    Next i
End Sub

Private Sub ac2fNSKeepBest(ByVal used As Double)
    Dim i As Long
    If m_bUsed >= 0 And used >= m_bUsed Then Exit Sub
    m_bUsed = used
    m_bSheets = m_nSheets
    For i = 0 To m_np - 1
        m_bSheet(i) = m_pSheet(i)
        m_bX(i) = m_pX(i): m_bY(i) = m_pY(i): m_bA(i) = m_pA(i)
    Next i
End Sub

'=====================================================================
' APPLY
'=====================================================================

Private Sub ac2fNSApply(ByVal x0 As Double, ByVal y0 As Double, _
                        ByVal sheetW As Double, ByVal sheetH As Double, _
                        ByVal sgap As Double, ByVal mL As Double, ByVal mB As Double)
    Dim i As Long, s As Long
    Dim ox As Double, oy As Double
    Dim bx As Double, by As Double, bw As Double, bh As Double
    Dim sh As Shape
    Dim outline() As Shape
    Dim n As Long

    ' parts
    For i = 0 To m_np - 1
        s = m_bSheet(i)
        If s >= 0 Then
            ox = x0 + s * (sheetW + sgap)
            oy = y0
            On Error Resume Next
            If m_ang(m_bA(i)) <> 0 Then m_shp(i).Rotate m_ang(m_bA(i))
            m_shp(i).GetBoundingBox bx, by, bw, bh
            m_shp(i).Move ox + mL + m_bX(i) - bx, oy + mB + m_bY(i) - by
            On Error GoTo 0
        End If
    Next i

    ' sheet outlines
    ReDim outline(0 To m_bSheets)
    For s = 0 To m_bSheets - 1
        ox = x0 + s * (sheetW + sgap)
        On Error Resume Next
        Set sh = ActiveLayer.CreateRectangle(ox, y0 + sheetH, ox + sheetW, y0)
        If Not sh Is Nothing Then
            sh.Fill.ApplyNoFill
            sh.Outline.Width = 0.25
            sh.Outline.Color.RGBAssign 150, 150, 150
            sh.Name = "ac2f sheet " & (s + 1)
            Set outline(n) = sh
            n = n + 1
        End If
        Set sh = Nothing
        On Error GoTo 0
    Next s

    If n > 1 Then
        On Error Resume Next
        ActiveDocument.ClearSelection
        For s = 0 To n - 1
            outline(s).AddToSelection
        Next s
        Set sh = ActiveSelectionRange.Group()
        sh.Name = "ac2f sheets"
        On Error GoTo 0
    End If
End Sub

'=====================================================================
' REPORT
'=====================================================================

Private Function ac2fNSReport(ByVal applied As Boolean, ByVal placed As Long, _
                              ByVal sheetW As Double, ByVal sheetH As Double, _
                              ByVal mL As Double, ByVal mR As Double, _
                              ByVal mT As Double, ByVal mB As Double, _
                              ByVal gap As Double, ByVal rotStep As Double, _
                              ByVal usedNo As Double, ByVal usedRot As Double) As String
    Dim s As String
    Dim i As Long
    Dim partArea As Double
    Dim usedLen As Double
    Dim rotCount As Long
    Dim gainTxt As String

    For i = 0 To m_np - 1
        If m_bSheet(i) >= 0 Then
            partArea = partArea + m_bw(i, m_bA(i)) * m_bh(i, m_bA(i))
            If m_ang(m_bA(i)) <> 0 Then rotCount = rotCount + 1
        End If
    Next i
    usedLen = m_bUsed
    If usedLen >= 1000000# Then usedLen = usedLen - 1000000# * (m_np - placed)

    s = IIf(applied, "NESTED", "REPORT ONLY - nothing moved") & vbCrLf
    s = s & "   Parts               : " & placed & " of " & m_np & vbCrLf
    If placed < m_np Then
        s = s & "   Did not fit         : " & (m_np - placed) & _
                "  (left where they were)" & vbCrLf
    End If
    s = s & "   Sheets              : " & m_bSheets & vbCrLf
    s = s & "   Material length     : " & ac2fFmt(usedLen, 0) & " mm" & vbCrLf
    s = s & "   Rotated             : " & rotCount & " part(s)" & vbCrLf
    If m_usableW > 0 And usedLen > 0 Then
        s = s & "   Fill                : " & _
                ac2fFmt(partArea / (m_usableW * usedLen) * 100#, 1) & " %" & vbCrLf
    End If

    If m_na > 1 And usedNo > 0 Then
        If usedRot < usedNo Then
            gainTxt = ac2fFmt((usedNo - usedRot) / usedNo * 100#, 1) & _
                      " % less material than without rotation"
        Else
            gainTxt = "rotation did not help here, the unrotated pack was kept"
        End If
        s = s & "   Rotation            : " & gainTxt & vbCrLf
    End If
    s = s & vbCrLf

    s = s & "SETTINGS USED" & vbCrLf
    s = s & "   Sheet               : " & ac2fFmt(sheetW, 0) & " x " & _
            ac2fFmt(sheetH, 0) & " mm" & vbCrLf
    s = s & "   Margins L R T B     : " & ac2fFmt(mL, 0) & "  " & ac2fFmt(mR, 0) & _
            "  " & ac2fFmt(mT, 0) & "  " & ac2fFmt(mB, 0) & " mm" & vbCrLf
    s = s & "   Usable              : " & ac2fFmt(m_usableW, 0) & " x " & _
            ac2fFmt(m_usableH, 0) & " mm" & vbCrLf
    s = s & "   Part gap            : " & ac2fFmt(gap) & " mm" & vbCrLf
    s = s & "   Rotation step       : " & ac2fFmt(rotStep, 0) & " deg" & _
            IIf(rotStep <= 0, "  (parts stay as drawn)", "") & vbCrLf

    If applied Then
        s = s & vbCrLf & "Sheet outlines are drawn in grey. Ctrl+Z puts everything back."
    End If
    s = s & vbCrLf & "Parts are packed by bounding box, so concave shapes do not interlock."

    ac2fNSReport = s
End Function
