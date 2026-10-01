Attribute VB_Name = "ac2fPlot"
'=====================================================================
'  ac2f pack  --  ac2fPlot
'
'  Sends the selection straight to a cutting plotter over the network,
'  as HPGL, with no export step and no file to fix afterwards.
'
'  Macros:
'    ac2fPlotSend    - selection -> HPGL -> plotter
'    ac2fPlotSave    - selection -> HPGL -> file (to inspect or keep)
'    ac2fPlotFixSend - normalise an existing .plt and send that
'
'  WHY THERE IS NO FIX STEP
'  The usual trouble with an exported .plt is that the geometry sits at
'  whatever coordinates the page happened to give it, often negative, so
'  it has to be parsed and shifted before the plotter will take it. Here
'  the shift is applied while the HPGL is written, from the geometry
'  itself, so there is nothing to parse and nothing to correct.
'
'  HPGL units are 40 per millimetre (1016 per inch).
'
'  ORIENTATION
'  Nothing is rotated: document X becomes HPGL X and document Y
'  becomes HPGL Y, so the job reaches the plotter exactly as it sits
'  on screen. There is no export filter in the way to turn it.
'
'  If it still comes out turned, that is the machine: on most cutters
'  the X axis runs along the media feed, so a wide job lands across
'  the roll. The Rotate setting compensates for that and defaults to
'  0, which is no rotation at all.
'
'  SENDING
'  VBA has no socket of its own, so the bytes go out through a short
'  PowerShell script using System.Net.Sockets.TcpClient. PowerShell is on
'  every supported Windows, which avoids Declare statements, 32/64 bit
'  trouble and registering any control. The script writes a log the macro
'  reads back, so a refused connection reports the real reason.
'
'  This is a raw TCP stream, which is what plotters listening on 9100 or
'  a telnet port expect. It does not negotiate the telnet protocol.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Plotter"

Public Const AC2F_UNITS_PER_MM As Double = 40#   ' HPGL resolution

Private Const MAX_PTS      As Long = 500000
Private Const CHUNK_BYTES  As Long = 1024        ' used when a send delay is set
Private Const CONNECT_MS   As Long = 5000

'---------------------------------------------------------------------
' Registry keys (internal, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_PL_MARGIN As String = "PLKenarPayiMM"
Public Const AC2F_K_PL_TOL    As String = "PLEgriToleransiMM"
Public Const AC2F_K_PL_ROT    As String = "PLDondurmeDerece"
Public Const AC2F_K_PL_MIR    As String = "PLAynalama"
Public Const AC2F_K_PL_PAIRS  As String = "PLPDCiftSayisi"
Public Const AC2F_K_PL_PRE    As String = "PLOnsoz"
Public Const AC2F_K_PL_DELAY  As String = "PLGonderimGecikmesiMS"
Public Const AC2F_K_PL_TARGET As String = "PLPlotterAdresi"   ' string, not on the sheet

Public Const AC2F_DEF_PL_MARGIN As Double = 5#     ' margin from the origin (mm)
Public Const AC2F_DEF_PL_TOL    As Double = 0.05   ' curve flattening tolerance (mm)
Public Const AC2F_DEF_PL_ROT    As Long = 0        ' 0 = exactly as you see it
Public Const AC2F_DEF_PL_MIR    As Long = 0        ' 0 none, 1 flip X, 2 flip Y, 3 both
Public Const AC2F_DEF_PL_PAIRS  As Long = 1        ' coordinate pairs per PD
Public Const AC2F_DEF_PL_PRE    As Long = 1        ' 0 none, 1 IN;SP1;PA;, 2 SP1;PA;
Public Const AC2F_DEF_PL_DELAY  As Long = 0        ' ms per 1 KB, 0 = send in one go
Public Const AC2F_DEF_PL_TARGET As String = "192.168.1.100:9100"

'---------------------------------------------------------------------
' Flattened geometry
'---------------------------------------------------------------------
Private m_x() As Double, m_y() As Double, m_n As Long
Private m_ss() As Long, m_sc() As Long, m_scl() As Boolean, m_sn As Long

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fPlotSend()
Attribute ac2fPlotSend.VB_Description = "ac2f pack: Send the selection to the plotter as HPGL"
    ac2fPLRun True
End Sub

Public Sub ac2fPlotSave()
Attribute ac2fPlotSave.VB_Description = "ac2f pack: Write the selection to an HPGL .plt file"
    ac2fPLRun False
End Sub

' Sends a 40 x 80 mm letter L through exactly the same writer, sender
' and orientation step a real job uses.
'
' The shape is deliberately lopsided in both axes. A square cannot show
' you anything about orientation: it looks identical whichever way the
' machine turns it, and mirrored it is still a square. An L tells you in
' one send whether the job is turned, mirrored, scaled or simply not
' being cut.
Public Sub ac2fPlotTest()
Attribute ac2fPlotTest.VB_Description = "ac2f pack: Send a test L to the plotter to check cutting and orientation"
    Dim host As String, port As Long
    Dim path As String, errText As String
    Dim bytesOut As Long
    Dim margin As Double
    Dim minX As Double, minY As Double
    Dim i As Long

    If Not ac2fPLAskTarget(host, port) Then Exit Sub

    margin = ac2fGetNum(AC2F_K_PL_MARGIN, AC2F_DEF_PL_MARGIN)
    If margin < 0 Then margin = 0

    ' tall leg on the left, foot to the right
    m_n = 0: m_sn = 0
    ReDim m_x(0 To 15): ReDim m_y(0 To 15)
    ReDim m_ss(0 To 0): ReDim m_sc(0 To 0): ReDim m_scl(0 To 0)
    ac2fPLAddPt 0, 0
    ac2fPLAddPt 40, 0
    ac2fPLAddPt 40, 20
    ac2fPLAddPt 15, 20
    ac2fPLAddPt 15, 80
    ac2fPLAddPt 0, 80
    m_ss(0) = 0: m_sc(0) = 6: m_scl(0) = True
    m_sn = 1

    ac2fPLOrient

    minX = m_x(0): minY = m_y(0)
    For i = 1 To m_n - 1
        If m_x(i) < minX Then minX = m_x(i)
        If m_y(i) < minY Then minY = m_y(i)
    Next i

    path = ac2fPLTempPath("ac2f_test.plt")
    bytesOut = ac2fPLWrite(path, margin * AC2F_UNITS_PER_MM - minX * AC2F_UNITS_PER_MM, _
                                 margin * AC2F_UNITS_PER_MM - minY * AC2F_UNITS_PER_MM)
    If bytesOut = 0 Then
        ac2fWarn "Could not write " & path, CAPTION_
        Exit Sub
    End If

    If ac2fPLSendFile(path, host, port, errText) Then
        ac2fInfo _
            "A test L went to " & host & ":" & port & vbCrLf & _
            "Rotate " & ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT) & _
            ", Mirror " & ac2fGetLng(AC2F_K_PL_MIR, AC2F_DEF_PL_MIR) & vbCrLf & vbCrLf & _
            "It should cut this, 40 mm wide and 80 mm tall:" & vbCrLf & vbCrLf & _
            "      |" & vbCrLf & _
            "      |" & vbCrLf & _
            "      |___" & vbCrLf & vbCrLf & _
            "Tall leg on the LEFT, foot pointing RIGHT, 80 mm the tall way." & _
            vbCrLf & vbCrLf & _
            "LYING ON ITS SIDE" & vbCrLf & _
            "   The machine swaps the axes. Set Rotate to 270 and send " & _
            "again; if that turns it the wrong way, use 90." & vbCrLf & vbCrLf & _
            "A MIRROR IMAGE, foot pointing LEFT" & vbCrLf & _
            "   Set Mirror to 1. Upside down instead, set Mirror to 2." & vbCrLf & vbCrLf & _
            "RIGHT SHAPE BUT ONLY TRACED" & vbCrLf & _
            "   The machine is not cutting. Knife force, blade depth, " & _
            "tool, then try Preamble = 2." & vbCrLf & vbCrLf & _
            "NOTHING MOVED" & vbCrLf & _
            "   It never arrived. Check address, port and cable.", CAPTION_
    Else
        ac2fWarn "The test could not be sent." & vbCrLf & vbCrLf & _
                 errText & vbCrLf & vbCrLf & "File: " & path, CAPTION_
    End If
End Sub

'=====================================================================
' SELECTION -> HPGL
'=====================================================================

Private Sub ac2fPLRun(ByVal send As Boolean)
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim i As Long
    Dim tol As Double, margin As Double
    Dim dx As Double, dy As Double
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    Dim path As String, target As String
    Dim host As String
    Dim port As Long
    Dim errText As String
    Dim bytesOut As Long
    Dim rotDeg As Long

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Select what you want to plot.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Select what you want to plot.", CAPTION_
        Exit Sub
    End If

    tol = ac2fGetNum(AC2F_K_PL_TOL, AC2F_DEF_PL_TOL)
    If tol <= 0 Then tol = AC2F_DEF_PL_TOL
    margin = ac2fGetNum(AC2F_K_PL_MARGIN, AC2F_DEF_PL_MARGIN)
    If margin < 0 Then margin = 0

    ' Ask for the destination before anything is built, so a cancel here
    ' leaves the document untouched.
    If send Then
        If Not ac2fPLAskTarget(host, port) Then Exit Sub
        path = ac2fPLTempPath("ac2f_plot.plt")
    Else
        path = InputBox("Write the HPGL to:", ac2fTitle(CAPTION_), _
                        ac2fPLTempPath("ac2f_plot.plt"))
        If StrPtr(path) = 0 Then Exit Sub
        path = Trim$(path)
        If Len(path) = 0 Then Exit Sub
    End If

    On Error GoTo Fail
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    m_n = 0: m_sn = 0
    ReDim m_x(0 To 8191): ReDim m_y(0 To 8191)
    ReDim m_ss(0 To 255): ReDim m_sc(0 To 255): ReDim m_scl(0 To 255)

    For i = 1 To sr.Count
        ac2fPLShape sr(i), tol
    Next i

    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    On Error GoTo Fail

    If m_sn = 0 Then
        ac2fWarn "Nothing plottable in the selection." & vbCrLf & vbCrLf & _
                 "Bitmaps and objects with no path are skipped.", CAPTION_
        Exit Sub
    End If

    ' --- orientation --------------------------------------------------
    rotDeg = ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT)
    ac2fPLOrient

    ' --- normalisation, from the geometry, not from parsed text -------
    minX = m_x(0): maxX = m_x(0): minY = m_y(0): maxY = m_y(0)
    For i = 1 To m_n - 1
        If m_x(i) < minX Then minX = m_x(i)
        If m_x(i) > maxX Then maxX = m_x(i)
        If m_y(i) < minY Then minY = m_y(i)
        If m_y(i) > maxY Then maxY = m_y(i)
    Next i
    dx = margin * AC2F_UNITS_PER_MM - minX * AC2F_UNITS_PER_MM
    dy = margin * AC2F_UNITS_PER_MM - minY * AC2F_UNITS_PER_MM

    bytesOut = ac2fPLWrite(path, dx, dy)
    If bytesOut = 0 Then
        ac2fWarn "Could not write " & path, CAPTION_
        Exit Sub
    End If

    If send Then
        If ac2fPLSendFile(path, host, port, errText) Then
            ' The file was only a handoff to the sender; nothing is kept.
            On Error Resume Next
            Kill path
            On Error GoTo Fail
            ac2fInfo ac2fPLReport(True, host, port, "", bytesOut, margin, tol, _
                                  minX, minY, maxX, maxY, ""), CAPTION_
        Else
            ac2fWarn ac2fPLReport(False, host, port, path, bytesOut, margin, tol, _
                                  minX, minY, maxX, maxY, errText), CAPTION_
        End If
    Else
        ac2fInfo ac2fPLReport(True, "", 0, path, bytesOut, margin, tol, _
                              minX, minY, maxX, maxY, ""), CAPTION_
    End If
    Exit Sub

Fail:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    ac2fWarn "The operation failed: " & Err.Description, CAPTION_
End Sub

'---------------------------------------------------------------------
' Flatten one shape into the point buffer.
'---------------------------------------------------------------------
Private Sub ac2fPLShape(ByVal s As Shape, ByVal tol As Double)
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
            ac2fPLShape s.Shapes(i), tol
        Next i
        Exit Sub
    End If

    Set cv = Nothing
    On Error Resume Next
    Set cv = s.DisplayCurve
    On Error GoTo Skip

    If Not cv Is Nothing Then
        ac2fPLCurve cv, tol
        Exit Sub
    End If

    Set dup = s.Duplicate(0, 0)
    On Error Resume Next
    dup.ConvertToCurves
    Set cv = dup.Curve
    On Error GoTo SkipDup
    If cv Is Nothing Then GoTo SkipDup
    ac2fPLCurve cv, tol
    dup.Delete
    Exit Sub

SkipDup:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
    Exit Sub
Skip:
End Sub

Private Sub ac2fPLCurve(ByVal cv As Curve, ByVal tol As Double)
    Dim i As Long, k As Long, nSub As Long, nSeg As Long
    Dim sp As SubPath, sg As Segment
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim cxx As Double, cyy As Double, chord As Double
    Dim arcL As Double, th As Double, turnDir As Double
    Dim steps As Long, r As Double, stepLen As Double

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
                If m_sn > UBound(m_ss) Then
                    ReDim Preserve m_ss(0 To (UBound(m_ss) + 1) * 2 - 1)
                    ReDim Preserve m_sc(0 To UBound(m_ss))
                    ReDim Preserve m_scl(0 To UBound(m_ss))
                End If
                m_ss(m_sn) = m_n
                m_scl(m_sn) = False
                On Error Resume Next
                m_scl(m_sn) = sp.Closed
                On Error GoTo 0

                For k = 1 To nSeg
                    Set sg = sp.Segments(k)
                    arcL = sg.Length
                    ax = sg.StartNode.PositionX: ay = sg.StartNode.PositionY
                    bx = sg.EndNode.PositionX: by = sg.EndNode.PositionY
                    cxx = bx - ax: cyy = by - ay
                    chord = Sqr(cxx * cxx + cyy * cyy)

                    th = 0
                    If arcL > 0 Then
                        If (arcL - chord) / arcL >= 0.0002 Then th = ac2fSolveTheta(chord, arcL)
                    End If

                    ' Sagitta h = s^2 / (8R), so a chord of sqrt(8*R*tol)
                    ' stays within tol of the true arc.
                    steps = 1
                    If th > 0.000000001 Then
                        r = arcL / th
                        stepLen = Sqr(8# * r * tol)
                        If stepLen > 0 Then steps = CLng(ac2fCeil(arcL / stepLen))
                        If steps < 1 Then steps = 1
                        If steps > 2000 Then steps = 2000
                    End If

                    turnDir = ac2fPLSegSign(sp, nSeg, k)
                    ac2fPLEmitArc ax, ay, bx, by, arcL, th, turnDir, steps
                    If m_n >= MAX_PTS Then Exit For
                Next k

                ' last node of the subpath
                Set sg = sp.Segments(nSeg)
                ac2fPLAddPt sg.EndNode.PositionX, sg.EndNode.PositionY

                m_sc(m_sn) = m_n - m_ss(m_sn)
                If m_sc(m_sn) >= 2 Then
                    m_sn = m_sn + 1
                Else
                    m_n = m_ss(m_sn)
                End If
            End If
        End If
        If m_n >= MAX_PTS Then Exit For
    Next i
End Sub

Private Function ac2fPLSegSign(ByVal sp As SubPath, ByVal nSeg As Long, _
                               ByVal k As Long) As Double
    Dim a As Double, b As Double
    Dim dPrev As Double, dCur As Double, dNext As Double

    If nSeg < 2 Then
        ac2fPLSegSign = 1#
        Exit Function
    End If
    dCur = ac2fPLChordDir(sp, k)
    dNext = ac2fPLChordDir(sp, IIf(k = nSeg, 1, k + 1))
    dPrev = ac2fPLChordDir(sp, IIf(k = 1, nSeg, k - 1))
    a = ac2fWrapAngle(dNext - dCur)
    b = ac2fWrapAngle(dCur - dPrev)
    If (a + b) >= 0 Then ac2fPLSegSign = 1# Else ac2fPLSegSign = -1#
End Function

Private Function ac2fPLChordDir(ByVal sp As SubPath, ByVal k As Long) As Double
    Dim sg As Segment
    On Error Resume Next
    Set sg = sp.Segments(k)
    If sg Is Nothing Then Exit Function
    ac2fPLChordDir = ac2fAtan2(sg.EndNode.PositionY - sg.StartNode.PositionY, _
                               sg.EndNode.PositionX - sg.StartNode.PositionX)
End Function

' Start point included, end point left to the next segment.
Private Sub ac2fPLEmitArc(ByVal ax As Double, ByVal ay As Double, _
                          ByVal bx As Double, ByVal by As Double, _
                          ByVal arcL As Double, ByVal th As Double, _
                          ByVal turnDir As Double, ByVal steps As Long)
    Dim i As Long
    Dim r As Double, phi As Double, t0 As Double
    Dim cx As Double, cy As Double, a0 As Double, ang As Double

    If th < 0.000000001 Then
        For i = 0 To steps - 1
            ac2fPLAddPt ax + (bx - ax) * i / steps, ay + (by - ay) * i / steps
        Next i
        Exit Sub
    End If

    r = arcL / th
    phi = ac2fAtan2(by - ay, bx - ax)
    t0 = phi - turnDir * th / 2#
    cx = ax + turnDir * r * (-Sin(t0))
    cy = ay + turnDir * r * Cos(t0)
    a0 = ac2fAtan2(ay - cy, ax - cx)

    For i = 0 To steps - 1
        ang = a0 + turnDir * th * i / steps
        ac2fPLAddPt cx + r * Cos(ang), cy + r * Sin(ang)
    Next i
End Sub

Private Sub ac2fPLAddPt(ByVal x As Double, ByVal y As Double)
    If m_n > UBound(m_x) Then
        ReDim Preserve m_x(0 To (UBound(m_x) + 1) * 2 - 1)
        ReDim Preserve m_y(0 To UBound(m_x))
    End If
    m_x(m_n) = x
    m_y(m_n) = y
    m_n = m_n + 1
End Sub

' Turns the flattened points a whole number of quarter turns. Applied
' before the extent is taken, so the margin still lands correctly
' whichever way the job ends up facing.
' Rotation then mirroring, in that order, on the flattened points.
' Both the real job and the test shape go through here, so whatever the
' test tells you to set is exactly what a real job will do.
Private Sub ac2fPLOrient()
    ac2fPLRotate ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT)
    ac2fPLMirror ac2fGetLng(AC2F_K_PL_MIR, AC2F_DEF_PL_MIR)
End Sub

' Rotation cannot undo a mirrored axis, and some machines do mirror one.
Private Sub ac2fPLMirror(ByVal m As Long)
    Dim i As Long
    If m <= 0 Then Exit Sub
    For i = 0 To m_n - 1
        If m = 1 Or m = 3 Then m_x(i) = -m_x(i)
        If m = 2 Or m = 3 Then m_y(i) = -m_y(i)
    Next i
End Sub

Private Sub ac2fPLRotate(ByVal deg As Long)
    Dim i As Long, q As Long
    Dim t As Double

    q = ((deg Mod 360) + 360) Mod 360
    q = ((q + 45) \ 90) Mod 4              ' snap to the nearest quarter
    If q = 0 Then Exit Sub                 ' 0 = exactly as on screen

    For i = 0 To m_n - 1
        Select Case q
            Case 1                          ' 90 counter-clockwise
                t = m_x(i): m_x(i) = -m_y(i): m_y(i) = t
            Case 2
                m_x(i) = -m_x(i): m_y(i) = -m_y(i)
            Case 3                          ' 270, i.e. 90 clockwise
                t = m_x(i): m_x(i) = m_y(i): m_y(i) = -t
        End Select
    Next i
End Sub

'=====================================================================
' HPGL OUTPUT
'=====================================================================

' Writes the file and returns its size in bytes, 0 on failure.
Private Function ac2fPLWrite(ByVal path As String, ByVal dx As Double, _
                             ByVal dy As Double) As Long
    Dim f As Integer
    Dim i As Long, k As Long, s As Long, n As Long
    Dim line_ As String
    Dim cnt As Long
    Dim total As Long
    Dim perPD As Long
    Dim pre As Long

    perPD = ac2fGetLng(AC2F_K_PL_PAIRS, AC2F_DEF_PL_PAIRS)
    If perPD < 1 Then perPD = 1
    pre = ac2fGetLng(AC2F_K_PL_PRE, AC2F_DEF_PL_PRE)

    On Error GoTo Fail
    f = FreeFile
    Open path For Output As #f

    ' IN resets the device to its power-on defaults. On a good many
    ' cutters that also throws away the knife force and the tool set on
    ' the panel, and the machine then traces the job without cutting.
    ' Preamble 2 keeps the panel settings; 0 sends no preamble at all.
    If pre = 1 Then Print #f, "IN;"
    If pre >= 1 Then
        Print #f, "SP1;"
        Print #f, "PA;"
    End If

    For i = 0 To m_sn - 1
        s = m_ss(i): n = m_sc(i)
        Print #f, "PU" & ac2fPLU(m_x(s), dx) & "," & ac2fPLU(m_y(s), dy) & ";"

        line_ = "": cnt = 0
        For k = 1 To n - 1
            If cnt > 0 Then line_ = line_ & ","
            line_ = line_ & ac2fPLU(m_x(s + k), dx) & "," & ac2fPLU(m_y(s + k), dy)
            cnt = cnt + 1
            If cnt >= perPD Then
                Print #f, "PD" & line_ & ";"
                line_ = "": cnt = 0
            End If
        Next k

        ' close the outline by returning to its first point
        If m_scl(i) Then
            If cnt > 0 Then line_ = line_ & ","
            line_ = line_ & ac2fPLU(m_x(s), dx) & "," & ac2fPLU(m_y(s), dy)
            cnt = cnt + 1
        End If
        If cnt > 0 Then Print #f, "PD" & line_ & ";"
    Next i

    If pre >= 1 Then
        Print #f, "PU0,0;"
        Print #f, "SP0;"
    End If
    Close #f

    total = ac2fPLFileSize(path)
    ac2fPLWrite = total
    Exit Function

Fail:
    On Error Resume Next
    Close #f
End Function

' Millimetres to plotter units, with the normalising shift applied.
Private Function ac2fPLU(ByVal mm As Double, ByVal d As Double) As String
    ac2fPLU = CStr(CLng(mm * AC2F_UNITS_PER_MM + d))
End Function

'=====================================================================
' SENDING
'=====================================================================

' Asks for host:port, remembering what was typed.
Private Function ac2fPLAskTarget(ByRef host As String, ByRef port As Long) As Boolean
    Dim answer As String
    Dim stored As String
    Dim p As Long

    stored = ac2fGetStr(AC2F_K_PL_TARGET, AC2F_DEF_PL_TARGET)

    answer = InputBox( _
        "Plotter address as host:port" & vbCrLf & vbCrLf & _
        "   192.168.1.100:9100    raw socket, the usual case" & vbCrLf & _
        "   192.168.1.100:23      a telnet port" & vbCrLf & vbCrLf & _
        "The bytes go out as a raw TCP stream. Port 9100 if you are " & _
        "not sure.", ac2fTitle(CAPTION_), stored)

    If StrPtr(answer) = 0 Then Exit Function
    answer = Trim$(answer)
    If Len(answer) = 0 Then Exit Function

    p = InStrRev(answer, ":")
    If p <= 1 Then
        host = answer
        port = 9100
    Else
        host = Trim$(Left$(answer, p - 1))
        port = CLng(Val(Mid$(answer, p + 1)))
        If port <= 0 Or port > 65535 Then port = 9100
    End If
    If Len(host) = 0 Then
        ac2fWarn "No host in """ & answer & """.", CAPTION_
        Exit Function
    End If

    ac2fSetStr AC2F_K_PL_TARGET, host & ":" & port
    ac2fPLAskTarget = True
End Function

' Streams the file to host:port through a small PowerShell script.
Private Function ac2fPLSendFile(ByVal path As String, ByVal host As String, _
                                ByVal port As Long, ByRef errText As String) As Boolean
    Dim ps As String, logf As String, cmd As String
    Dim sh As Object
    Dim rc As Long

    ps = ac2fPLTempPath("ac2f_send.ps1")
    logf = ac2fPLTempPath("ac2f_send.log")

    On Error Resume Next
    Kill logf
    On Error GoTo Fail

    If Not ac2fPLWriteSender(ps) Then
        errText = "Could not write the sender script to " & ps
        Exit Function
    End If

    Set sh = CreateObject("WScript.Shell")
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ps & """" & _
          " -Path """ & path & """" & _
          " -Target """ & host & """" & _
          " -Port " & port & _
          " -TimeoutMs " & CONNECT_MS & _
          " -ChunkBytes " & CHUNK_BYTES & _
          " -ChunkDelayMs " & ac2fGetLng(AC2F_K_PL_DELAY, AC2F_DEF_PL_DELAY) & _
          " -LogPath """ & logf & """"

    rc = sh.Run(cmd, 0, True)
    If rc = 0 Then
        ac2fPLSendFile = True
    Else
        errText = ac2fPLReadAll(logf)
        If Len(Trim$(errText)) = 0 Then _
            errText = "PowerShell exited with code " & rc & "."
    End If
    Exit Function

Fail:
    errText = "Could not start PowerShell: " & Err.Description
End Function

' The sender. Written out each run so it always matches this module.
' The parameter is called Target, not Host: $Host is a PowerShell
' automatic variable and binding to it fails.
Private Function ac2fPLWriteSender(ByVal ps As String) As Boolean
    Dim f As Integer
    Dim q As String

    q = Chr$(34)

    On Error GoTo Fail
    f = FreeFile
    Open ps For Output As #f
    Print #f, "param("
    Print #f, "  [string]$Path,"
    Print #f, "  [string]$Target,"
    Print #f, "  [int]$Port,"
    Print #f, "  [string]$LogPath,"
    Print #f, "  [int]$TimeoutMs = 5000,"
    Print #f, "  [int]$ChunkBytes = 0,"
    Print #f, "  [int]$ChunkDelayMs = 0"
    Print #f, ")"
    Print #f, "$ErrorActionPreference = 'Stop'"
    Print #f, "try {"
    Print #f, "  $bytes = [System.IO.File]::ReadAllBytes($Path)"
    Print #f, "  $client = New-Object System.Net.Sockets.TcpClient"
    Print #f, "  $iar = $client.BeginConnect($Target, $Port, $null, $null)"
    Print #f, "  if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {"
    Print #f, "    $client.Close()"
    Print #f, "    throw 'No answer from ' + $Target + ':' + $Port + " & _
              "' within ' + $TimeoutMs + ' ms'"
    Print #f, "  }"
    Print #f, "  $client.EndConnect($iar)"
    Print #f, "  $stream = $client.GetStream()"
    Print #f, "  if ($ChunkBytes -gt 0 -and $ChunkDelayMs -gt 0) {"
    Print #f, "    $off = 0"
    Print #f, "    while ($off -lt $bytes.Length) {"
    Print #f, "      $n = [Math]::Min($ChunkBytes, $bytes.Length - $off)"
    Print #f, "      $stream.Write($bytes, $off, $n)"
    Print #f, "      $stream.Flush()"
    Print #f, "      $off += $n"
    Print #f, "      Start-Sleep -Milliseconds $ChunkDelayMs"
    Print #f, "    }"
    Print #f, "  } else {"
    Print #f, "    $stream.Write($bytes, 0, $bytes.Length)"
    Print #f, "    $stream.Flush()"
    Print #f, "  }"
    Print #f, "  Start-Sleep -Milliseconds 400"
    Print #f, "  $stream.Close()"
    Print #f, "  $client.Close()"
    Print #f, "  ('OK ' + $bytes.Length + ' bytes to ' + $Target + ':' + $Port) | " & _
              "Out-File -FilePath $LogPath -Encoding ascii"
    Print #f, "  exit 0"
    Print #f, "} catch {"
    Print #f, "  $_.Exception.Message | Out-File -FilePath $LogPath -Encoding ascii"
    Print #f, "  exit 1"
    Print #f, "}"
    Close #f
    ac2fPLWriteSender = True
    Exit Function

Fail:
    On Error Resume Next
    Close #f
End Function

'=====================================================================
' NORMALISE AN EXISTING .plt
'=====================================================================

Public Sub ac2fPlotFixSend()
Attribute ac2fPlotFixSend.VB_Description = "ac2f pack: Normalise an existing .plt file and send it"
    Dim inp As String, outp As String
    Dim txt As String
    Dim minX As Long, minY As Long, maxX As Long, maxY As Long
    Dim dx As Long, dy As Long
    Dim margin As Double
    Dim host As String, port As Long
    Dim errText As String
    Dim bytesOut As Long
    Dim s As String

    inp = InputBox("Path of the .plt file to normalise:", ac2fTitle(CAPTION_), "")
    If StrPtr(inp) = 0 Then Exit Sub
    inp = Trim$(inp)
    If Len(inp) = 0 Then Exit Sub

    txt = ac2fPLReadAll(inp)
    If Len(txt) = 0 Then
        ac2fWarn "Could not read " & inp, CAPTION_
        Exit Sub
    End If

    If Not ac2fPLExtent(txt, minX, minY, maxX, maxY) Then
        ac2fWarn "No PU/PD coordinates found in that file.", CAPTION_
        Exit Sub
    End If

    margin = ac2fGetNum(AC2F_K_PL_MARGIN, AC2F_DEF_PL_MARGIN)
    If margin < 0 Then margin = 0
    dx = CLng(margin * AC2F_UNITS_PER_MM) - minX
    dy = CLng(margin * AC2F_UNITS_PER_MM) - minY

    outp = Left$(inp, InStrRev(inp, ".") - 1) & "_fixed.plt"
    If InStrRev(inp, ".") = 0 Then outp = inp & "_fixed.plt"

    bytesOut = ac2fPLShiftFile(inp, outp, dx, dy)
    If bytesOut = 0 Then
        ac2fWarn "Could not write " & outp, CAPTION_
        Exit Sub
    End If

    s = "NORMALISED" & vbCrLf
    s = s & "   In                  : " & inp & vbCrLf
    s = s & "   Out                 : " & outp & vbCrLf
    s = s & "   Shift               : X " & dx & "   Y " & dy & " units" & vbCrLf
    s = s & "   Was                 : X " & minX & ".." & maxX & "   Y " & minY & ".." & maxY & vbCrLf
    s = s & "   Now                 : X " & (minX + dx) & ".." & (maxX + dx) & _
            "   Y " & (minY + dy) & ".." & (maxY + dy) & vbCrLf
    s = s & "   Size                : " & ac2fFmt((maxX - minX) / AC2F_UNITS_PER_MM, 0) & _
            " x " & ac2fFmt((maxY - minY) / AC2F_UNITS_PER_MM, 0) & " mm" & vbCrLf
    s = s & "   Bytes               : " & Format$(bytesOut, "#,##0") & vbCrLf

    If MsgBox(s & vbCrLf & "Send it to the plotter now?", _
              vbQuestion + vbYesNo, ac2fTitle(CAPTION_)) <> vbYes Then
        ac2fInfo s, CAPTION_
        Exit Sub
    End If

    If Not ac2fPLAskTarget(host, port) Then Exit Sub
    If ac2fPLSendFile(outp, host, port, errText) Then
        ac2fInfo s & vbCrLf & "Sent to " & host & ":" & port & ".", CAPTION_
    Else
        ac2fWarn s & vbCrLf & "Send failed: " & errText, CAPTION_
    End If
End Sub

' Extent of the DRAWING only: every PD, plus a PU that is followed by a
' PD. A trailing "PU0,0;" parks the pen and is not artwork; counting it
' puts the minimum at 0,0, and then a drawing that already sits in
' positive coordinates gets shifted by a flat margin instead of being
' normalised, wasting material.
Private Function ac2fPLExtent(ByVal txt As String, ByRef minX As Long, _
                              ByRef minY As Long, ByRef maxX As Long, _
                              ByRef maxY As Long) As Boolean
    Dim i As Long, j As Long
    Dim cmd As String
    Dim body As String
    Dim nums() As Double
    Dim cnt As Long
    Dim nextIsPD As Boolean
    Dim first As Boolean
    Dim k As Long

    first = True
    i = 1
    Do
        i = ac2fPLNextCmd(txt, i, cmd, body, j)
        If i = 0 Then Exit Do

        If cmd = "PD" Or cmd = "PU" Then
            If cmd = "PD" Then
                nextIsPD = True
            Else
                nextIsPD = ac2fPLFollowedByPD(txt, j)
            End If

            If nextIsPD Then
                cnt = ac2fPLNumbers(body, nums)
                For k = 0 To cnt - 2 Step 2
                    If first Then
                        minX = CLng(nums(k)): maxX = minX
                        minY = CLng(nums(k + 1)): maxY = minY
                        first = False
                    Else
                        If nums(k) < minX Then minX = CLng(nums(k))
                        If nums(k) > maxX Then maxX = CLng(nums(k))
                        If nums(k + 1) < minY Then minY = CLng(nums(k + 1))
                        If nums(k + 1) > maxY Then maxY = CLng(nums(k + 1))
                    End If
                Next k
            End If
        End If
        i = j
    Loop

    ac2fPLExtent = Not first
End Function

Private Function ac2fPLFollowedByPD(ByVal txt As String, ByVal from As Long) As Boolean
    Dim cmd As String, body As String, j As Long
    If ac2fPLNextCmd(txt, from, cmd, body, j) = 0 Then Exit Function
    ac2fPLFollowedByPD = (cmd = "PD")
End Function

' Finds the next PU or PD command at or after 'from'. Returns its start
' position, 0 when there is none, and sets j past the ';'.
Private Function ac2fPLNextCmd(ByVal txt As String, ByVal from As Long, _
                               ByRef cmd As String, ByRef body As String, _
                               ByRef j As Long) As Long
    Dim p As Long, q As Long, e As Long
    Dim two As String

    p = from
    Do While p > 0 And p <= Len(txt) - 1
        two = UCase$(Mid$(txt, p, 2))
        If two = "PU" Or two = "PD" Then
            e = InStr(p, txt, ";")
            If e = 0 Then Exit Function
            cmd = two
            body = Mid$(txt, p + 2, e - p - 2)
            j = e + 1
            ac2fPLNextCmd = p
            Exit Function
        End If
        p = p + 1
    Loop
End Function

' Pulls signed integers out of a command body.
Private Function ac2fPLNumbers(ByVal body As String, ByRef nums() As Double) As Long
    Dim i As Long, n As Long
    Dim c As String
    Dim cur As String
    Dim have As Boolean

    ReDim nums(0 To 63)
    For i = 1 To Len(body) + 1
        If i <= Len(body) Then c = Mid$(body, i, 1) Else c = " "
        If (c >= "0" And c <= "9") Or (c = "-" And Not have) Then
            cur = cur & c
            have = True
        Else
            If have Then
                If n > UBound(nums) Then ReDim Preserve nums(0 To (UBound(nums) + 1) * 2 - 1)
                nums(n) = Val(cur)
                n = n + 1
                cur = "": have = False
            End If
        End If
    Next i
    ac2fPLNumbers = n
End Function

' Copies the file, shifting every PU/PD coordinate pair.
Private Function ac2fPLShiftFile(ByVal inp As String, ByVal outp As String, _
                                 ByVal dx As Long, ByVal dy As Long) As Long
    Dim txt As String
    Dim fo As Integer
    Dim i As Long, j As Long, k As Long
    Dim cmd As String, body As String
    Dim nums() As Double
    Dim cnt As Long
    Dim out As String
    Dim pos As Long
    Dim piece As String

    txt = ac2fPLReadAll(inp)
    If Len(txt) = 0 Then Exit Function

    On Error GoTo Fail
    fo = FreeFile
    Open outp For Output As #fo

    pos = 1
    Do
        i = ac2fPLNextCmd(txt, pos, cmd, body, j)
        If i = 0 Then Exit Do

        ' everything before this command goes out unchanged
        If i > pos Then Print #fo, Mid$(txt, pos, i - pos);

        cnt = ac2fPLNumbers(body, nums)
        If cnt >= 2 Then
            piece = cmd
            For k = 0 To cnt - 2 Step 2
                If k > 0 Then piece = piece & ","
                piece = piece & CStr(CLng(nums(k)) + dx) & "," & CStr(CLng(nums(k + 1)) + dy)
            Next k
            Print #fo, piece & ";";
        Else
            Print #fo, Mid$(txt, i, j - i);
        End If
        pos = j
    Loop
    If pos <= Len(txt) Then Print #fo, Mid$(txt, pos);
    Close #fo

    ac2fPLShiftFile = ac2fPLFileSize(outp)
    Exit Function

Fail:
    On Error Resume Next
    Close #fo
End Function

'=====================================================================
' FILE AND REPORT HELPERS
'=====================================================================

Public Function ac2fPLMirName(ByVal m As Long) As String
    Select Case m
        Case 1:    ac2fPLMirName = "flipped left to right"
        Case 2:    ac2fPLMirName = "flipped top to bottom"
        Case 3:    ac2fPLMirName = "flipped both ways"
        Case Else: ac2fPLMirName = "none"
    End Select
End Function

Public Function ac2fPLPreName(ByVal p As Long) As String
    Select Case p
        Case 0:    ac2fPLPreName = "none"
        Case 2:    ac2fPLPreName = "SP1;PA;  (no IN, panel settings kept)"
        Case Else: ac2fPLPreName = "IN;SP1;PA;"
    End Select
End Function

Private Function ac2fPLTempPath(ByVal nm As String) As String
    Dim t As String
    On Error Resume Next
    t = Environ$("TEMP")
    If Len(t) = 0 Then t = Environ$("TMP")
    If Len(t) = 0 Then t = "C:\"
    If Right$(t, 1) <> "\" Then t = t & "\"
    ac2fPLTempPath = t & nm
End Function

Private Function ac2fPLReadAll(ByVal path As String) As String
    Dim f As Integer
    Dim s As String
    On Error GoTo Fail
    f = FreeFile
    Open path For Input As #f
    s = Input$(LOF(f), f)
    Close #f
    ac2fPLReadAll = s
    Exit Function
Fail:
    On Error Resume Next
    Close #f
End Function

Private Function ac2fPLFileSize(ByVal path As String) As Long
    On Error Resume Next
    ac2fPLFileSize = FileLen(path)
End Function

Private Function ac2fPLReport(ByVal ok As Boolean, ByVal host As String, _
                              ByVal port As Long, ByVal path As String, _
                              ByVal bytesOut As Long, ByVal margin As Double, _
                              ByVal tol As Double, ByVal minX As Double, _
                              ByVal minY As Double, ByVal maxX As Double, _
                              ByVal maxY As Double, ByVal errText As String) As String
    Dim s As String
    Dim u As Double
    u = AC2F_UNITS_PER_MM

    If Len(host) = 0 Then
        s = "WRITTEN" & vbCrLf
    ElseIf ok Then
        s = "SENT" & vbCrLf
    Else
        s = "SEND FAILED" & vbCrLf
    End If

    If Len(host) > 0 Then
        s = s & "   Plotter             : " & host & ":" & port & vbCrLf
    End If
    If Len(path) > 0 Then
        s = s & "   File                : " & path & vbCrLf
    Else
        s = s & "   File                : none kept (streamed)" & vbCrLf
    End If
    s = s & "   Bytes               : " & Format$(bytesOut, "#,##0") & vbCrLf
    s = s & "   Paths               : " & m_sn & vbCrLf
    s = s & "   Points              : " & Format$(m_n, "#,##0") & vbCrLf & vbCrLf

    s = s & "PLACEMENT" & vbCrLf
    s = s & "   Size                : " & ac2fFmt(maxX - minX) & " x " & _
            ac2fFmt(maxY - minY) & " mm" & vbCrLf
    s = s & "   Margin              : " & ac2fFmt(margin) & " mm  (" & _
            CLng(margin * u) & " units)" & vbCrLf
    s = s & "   X range             : " & CLng(margin * u) & " .. " & _
            CLng(margin * u + (maxX - minX) * u) & vbCrLf
    s = s & "   Y range             : " & CLng(margin * u) & " .. " & _
            CLng(margin * u + (maxY - minY) * u) & vbCrLf
    s = s & "   Curve tolerance     : " & ac2fFmt(tol) & " mm" & vbCrLf
    s = s & "   Mirror              : " & ac2fPLMirName(ac2fGetLng(AC2F_K_PL_MIR, AC2F_DEF_PL_MIR)) & vbCrLf
    s = s & "   PD pairs            : " & ac2fGetLng(AC2F_K_PL_PAIRS, AC2F_DEF_PL_PAIRS) & vbCrLf
    s = s & "   Preamble            : " & ac2fPLPreName(ac2fGetLng(AC2F_K_PL_PRE, AC2F_DEF_PL_PRE)) & vbCrLf
    s = s & "   Rotate              : " & ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT) & _
            " deg" & IIf(ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT) = 0, _
                         "  (as on screen)", "") & vbCrLf

    If Len(errText) > 0 Then
        s = s & vbCrLf & "REASON" & vbCrLf & "   " & errText & vbCrLf
        s = s & vbCrLf & "The file is written, so you can send it by hand."
    ElseIf Len(host) > 0 And ok Then
        s = s & vbCrLf & "Written straight from the geometry and streamed out." & vbCrLf
        s = s & "No CorelDRAW export, no file left behind, nothing to fix."
    End If

    ac2fPLReport = s
End Function
