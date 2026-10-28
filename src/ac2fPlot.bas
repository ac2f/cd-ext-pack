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
'  CUTTING ORDER
'  Objects go out in the order the Object Manager lists them, top
'  first, and nothing reorders them on the way. To cut in bending
'  order, run Bending order first: it rearranges the Object Manager,
'  and this module then sends exactly that. Send order 2 asks for the
'  bending order directly instead.
'
'  END OF JOB
'  When the last outline is done the head is lifted and parked at the
'  top of the job, so the work can be taken off. Park sets where.
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
Public Const AC2F_K_PL_SEQ    As String = "PLGonderimSirasi"
Public Const AC2F_K_PL_PARK   As String = "PLBitisKonumu"
Public Const AC2F_K_PL_TARGET As String = "PLPlotterAdresi"   ' string, not on the sheet
Public Const AC2F_K_PL_CALIB  As String = "PLYonKontrolEdildi" ' 1 once the wizard has run, not on the sheet

Public Const AC2F_DEF_PL_MARGIN As Double = 5#     ' margin from the origin (mm)
Public Const AC2F_DEF_PL_TOL    As Double = 0.05   ' curve flattening tolerance (mm)
Public Const AC2F_DEF_PL_ROT    As Long = 0        ' 0 = exactly as you see it
Public Const AC2F_DEF_PL_MIR    As Long = 0        ' 0 none, 1 flip X, 2 flip Y, 3 both
Public Const AC2F_DEF_PL_PAIRS  As Long = 1        ' coordinate pairs per PD
Public Const AC2F_DEF_PL_PRE    As Long = 1        ' 0 none, 1 IN;SP1;PA;, 2 SP1;PA;
Public Const AC2F_DEF_PL_SEQ    As Long = 0        ' 0 Object Manager top first, 1 bottom first, 2 bending order
Public Const AC2F_DEF_PL_PARK   As Long = 1        ' 0 origin, 1 top of the job, 2 stay put
Public Const AC2F_DEF_PL_DELAY  As Long = 0        ' ms per 1 KB, 0 = send in one go
Public Const AC2F_DEF_PL_TARGET As String = "192.168.1.100:9100"

Private m_seqMissing As Long        ' objects not found in the page stack
Private m_desOK As Boolean          ' the design's own size was read from CorelDRAW
Private m_desW As Double, m_desH As Double

'---------------------------------------------------------------------
' The flattened geometry lives in ac2fBending, which also puts it in
' bending order. This module turns it into HPGL and sends it.
'---------------------------------------------------------------------

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

    If Not ac2fPLAskTarget(host, port) Then Exit Sub

    If ac2fPLSendTestL(host, port, path, errText) Then
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
            "NOT RIGHT? Run the orientation wizard (main menu 27): you say what " & _
            "came out and it sets Rotate and Mirror for you, then checks." & _
            vbCrLf & vbCrLf & _
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

' Builds the test L, puts it through Rotate and Mirror as set now, writes
' and sends it. Returns False with errText when it could not be sent.
Private Function ac2fPLSendTestL(ByVal host As String, ByVal port As Long, _
                                 ByRef path As String, ByRef errText As String) As Boolean
    Dim margin As Double
    Dim minX As Double, minY As Double
    Dim i As Long
    Dim parkIdx As Long
    Dim bytesOut As Long

    margin = ac2fGetNum(AC2F_K_PL_MARGIN, AC2F_DEF_PL_MARGIN)
    If margin < 0 Then margin = 0

    ' tall leg on the left, foot to the right
    ac2fBNReset
    ac2fBNAddPt 0, 0
    ac2fBNAddPt 40, 0
    ac2fBNAddPt 40, 20
    ac2fBNAddPt 15, 20
    ac2fBNAddPt 15, 80
    ac2fBNAddPt 0, 80
    ac2fGS(0) = 0: ac2fGC(0) = 6: ac2fGCl(0) = True
    ac2fGSub = 1
    parkIdx = ac2fPLAddPark()

    ac2fPLOrient

    minX = ac2fGX(0): minY = ac2fGY(0)
    For i = 1 To ac2fGN - 1
        If ac2fGX(i) < minX Then minX = ac2fGX(i)
        If ac2fGY(i) < minY Then minY = ac2fGY(i)
    Next i

    path = ac2fPLTempPath("ac2f_test.plt")
    bytesOut = ac2fPLWrite(path, margin * AC2F_UNITS_PER_MM - minX * AC2F_UNITS_PER_MM, _
                                 margin * AC2F_UNITS_PER_MM - minY * AC2F_UNITS_PER_MM, parkIdx)
    If bytesOut = 0 Then
        errText = "Could not write " & path
        Exit Function
    End If
    ac2fPLSendTestL = ac2fPLSendFile(path, host, port, errText)
End Function

'=====================================================================
' ORIENTATION WIZARD
'
' Finding Rotate and Mirror by trial is slow: there are eight ways a
' machine can show a job, the first guess is wrong half the time, and
' every try costs material. The wizard needs one cut.
'
' 1  A test L goes out with NO turning and NO flipping, so it shows
'    what the machine does by itself.
' 2  You pick the picture that matches the piece. The eight pictures
'    are the eight turns and mirrors of the L; each is worked out from
'    the geometry and written in words (which corner, which way the
'    long leg runs).
' 3  Rotate and Mirror are set to the opposite turn.
' 4  The L goes out again with those settings, and you say whether it
'    is right. If it is not, what you now see is used to correct the
'    settings once more, until it is right.
'
' Turns and mirrors are kept as 2 x 2 matrices (entries -1, 0, 1), so
' undoing one, or following one by another, is exact and has no cases.
'=====================================================================
Public Sub ac2fPlotCalibrate()
Attribute ac2fPlotCalibrate.VB_Description = "ac2f pack: Find the plotter's Rotate and Mirror from one test cut"
    Dim host As String, port As Long
    Dim path As String, errText As String
    Dim sa As Long, sb As Long, sc As Long, sd As Long        ' what is stored now (S)
    Dim oa As Long, ob As Long, oc As Long, od As Long        ' what was seen (O)
    Dim na As Long, nb As Long, nc As Long, nd As Long
    Dim pick As Long, rounds As Long
    Dim q As Long, m As Long
    Dim oldRot As Long, oldMir As Long

    If Not ac2fPLAskTarget(host, port) Then Exit Sub

    If MsgBox("ORIENTATION WIZARD" & vbCrLf & vbCrLf & _
              "A small test L (40 x 80 mm) will be cut with NO turning and NO " & _
              "flipping, so it shows what the machine does on its own." & vbCrLf & vbCrLf & _
              "Put material in and press OK. Afterwards look at the piece the way " & _
              "it came off the machine, the edge nearest you at the bottom." & vbCrLf & vbCrLf & _
              "Machine: " & host & ":" & port, _
              vbOKCancel + vbInformation, ac2fTitle(CAPTION_)) <> vbOK Then Exit Sub

    ' What was set before comes back if the wizard is not finished.
    oldRot = ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT)
    oldMir = ac2fGetLng(AC2F_K_PL_MIR, AC2F_DEF_PL_MIR)

    ' S starts as "nothing": identity
    sa = 1: sb = 0: sc = 0: sd = 1
    ac2fPLStoreOrient 0, 0

    For rounds = 1 To 4
        If Not ac2fPLSendTestL(host, port, path, errText) Then
            ac2fPLStoreOrient oldRot, oldMir
            ac2fWarn "The test could not be sent." & vbCrLf & vbCrLf & _
                     errText & vbCrLf & vbCrLf & "File: " & path & vbCrLf & vbCrLf & _
                     "Rotate and Mirror are as they were.", CAPTION_
            Exit Sub
        End If

        pick = ac2fPLAskPicture(rounds)
        If pick = 0 Then
            ac2fPLStoreOrient oldRot, oldMir
            Exit Sub                                ' cancelled
        End If
        ac2fPLMatFromPick pick - 1, oa, ob, oc, od

        If pick = 1 Then
            ' The piece is right with the settings as they stand.
            ac2fSetLng AC2F_K_PL_CALIB, 1
            ac2fInfo "ORIENTATION SET" & vbCrLf & vbCrLf & _
                     "Rotate " & ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT) & _
                     ", Mirror " & ac2fGetLng(AC2F_K_PL_MIR, AC2F_DEF_PL_MIR) & vbCrLf & vbCrLf & _
                     "Remembered for every job from now on. Change it only if " & _
                     "the machine or its settings change; run this wizard again then.", CAPTION_
            Exit Sub
        End If

        ' what is stored now, followed by the opposite of what was seen
        ' (the opposite of a turn is the turn back; a mirror is its own
        ' opposite): the new setting S' = S * O^-1
        ac2fPLMatMul sa, sb, sc, sd, oa, oc, ob, od, na, nb, nc, nd
        sa = na: sb = nb: sc = nc: sd = nd
        If Not ac2fPLMatToQM(sa, sb, sc, sd, q, m) Then
            ac2fPLStoreOrient oldRot, oldMir
            Exit Sub
        End If
        ac2fPLStoreOrient q * 90, m
    Next rounds

    ac2fPLStoreOrient oldRot, oldMir
    ac2fWarn "The L still does not come out right after four tries." & vbCrLf & vbCrLf & _
             "Rotate and Mirror are as they were (" & oldRot & " and " & oldMir & "). " & _
             "If the shape itself is wrong (stretched, not 40 x 80), the plotter's own " & _
             "unit or scale setting is the cause, not the orientation.", CAPTION_
End Sub

Private Sub ac2fPLStoreOrient(ByVal rot As Long, ByVal mir As Long)
    ac2fSetLng AC2F_K_PL_ROT, rot
    ac2fSetLng AC2F_K_PL_MIR, mir
End Sub

' The eight pictures: index k = 0..7 is Rotate (k mod 4) quarter turns
' counter-clockwise, then a left-right flip when k >= 4. Index 0 is the
' L as drawn. Returns the matrix of that turn.
Private Sub ac2fPLMatFromPick(ByVal k As Long, ByRef a As Long, ByRef b As Long, _
                              ByRef c As Long, ByRef d As Long)
    ac2fPLMatFromQM k Mod 4, k \ 4, a, b, c, d
End Sub

' Rotation by q quarter turns (counter-clockwise), then mirror flags m
' (1 flip X, 2 flip Y, 3 both), as ac2fPLOrient applies them.
Private Sub ac2fPLMatFromQM(ByVal q As Long, ByVal m As Long, ByRef a As Long, _
                            ByRef b As Long, ByRef c As Long, ByRef d As Long)
    Select Case q
        Case 1:    a = 0:  b = -1: c = 1:  d = 0
        Case 2:    a = -1: b = 0:  c = 0:  d = -1
        Case 3:    a = 0:  b = 1:  c = -1: d = 0
        Case Else: a = 1:  b = 0:  c = 0:  d = 1
    End Select
    If m = 1 Or m = 3 Then a = -a: b = -b
    If m = 2 Or m = 3 Then c = -c: d = -d
End Sub

' Finds Rotate quarter turns q and a flip m that make the matrix.
Private Function ac2fPLMatToQM(ByVal a As Long, ByVal b As Long, ByVal c As Long, _
                               ByVal d As Long, ByRef q As Long, ByRef m As Long) As Boolean
    Dim qq As Long, mm As Long
    Dim ta As Long, tb As Long, tc As Long, td As Long
    For mm = 0 To 1
        For qq = 0 To 3
            ac2fPLMatFromQM qq, mm, ta, tb, tc, td
            If ta = a And tb = b And tc = c And td = d Then
                q = qq: m = mm
                ac2fPLMatToQM = True
                Exit Function
            End If
        Next qq
    Next mm
End Function

' R = A * B for 2 x 2 matrices.
Private Sub ac2fPLMatMul(ByVal a1 As Long, ByVal b1 As Long, ByVal c1 As Long, ByVal d1 As Long, _
                         ByVal a2 As Long, ByVal b2 As Long, ByVal c2 As Long, ByVal d2 As Long, _
                         ByRef ra As Long, ByRef rb As Long, ByRef rc As Long, ByRef rd As Long)
    ra = a1 * a2 + b1 * c2
    rb = a1 * b2 + b1 * d2
    rc = c1 * a2 + d1 * c2
    rd = c1 * b2 + d1 * d2
End Sub

' Where the corner and the legs of the L land when it is put through the
' matrix, in words. The L's long leg points up (0,1) and its short leg
' right (1,0); the corner is where they meet.
Private Function ac2fPLPictureText(ByVal a As Long, ByVal b As Long, _
                                   ByVal c As Long, ByVal d As Long) As String
    Dim ux As Long, uy As Long, vx As Long, vy As Long
    Dim corner As String

    ux = b: uy = d          ' long leg
    vx = a: vy = c          ' short leg
    If uy + vy > 0 Then corner = "bottom" Else corner = "top"
    If ux + vx > 0 Then corner = corner & "-left" Else corner = corner & "-right"
    ac2fPLPictureText = "corner " & corner & ", long leg " & ac2fPLDirName(ux, uy) & _
                        ", short leg " & ac2fPLDirName(vx, vy)
End Function

Private Function ac2fPLDirName(ByVal x As Long, ByVal y As Long) As String
    If y > 0 Then
        ac2fPLDirName = "UP"
    ElseIf y < 0 Then
        ac2fPLDirName = "DOWN"
    ElseIf x > 0 Then
        ac2fPLDirName = "RIGHT"
    Else
        ac2fPLDirName = "LEFT"
    End If
End Function

' Asks which of the eight pictures the piece looks like. Returns 1..8,
' 0 when cancelled. Picture 1 is the L as it should be.
Private Function ac2fPLAskPicture(ByVal rounds As Long) As Long
    Dim k As Long, a As Long, b As Long, c As Long, d As Long
    Dim t As String, answer As String, v As Long

    t = "Which picture is the piece you cut?" & vbCrLf & _
        "(edge nearest you at the bottom; do not flip the sheet over)" & vbCrLf & vbCrLf
    For k = 0 To 7
        ac2fPLMatFromPick k, a, b, c, d
        t = t & "  " & (k + 1) & "  -  " & ac2fPLPictureText(a, b, c, d)
        If k = 0 Then t = t & "   (RIGHT)"
        t = t & vbCrLf
    Next k
    If rounds > 1 Then
        t = t & vbCrLf & "Try " & rounds & ": the settings were corrected, this is what " & _
            "the new cut shows."
    End If

    answer = InputBox(t, ac2fTitle(CAPTION_), "1")
    If StrPtr(answer) = 0 Then Exit Function
    v = CLng(Val(Trim$(answer)))
    If v < 1 Or v > 8 Then
        ac2fWarn "Type a number from 1 to 8.", CAPTION_
        Exit Function
    End If
    ac2fPLAskPicture = v
End Function

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
    Dim orderAlg As Long
    Dim seqMode As Long
    Dim parkIdx As Long
    Dim bbx As Double, bby As Double, bbw As Double, bbh As Double

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

    seqMode = ac2fGetLng(AC2F_K_PL_SEQ, AC2F_DEF_PL_SEQ)
    ac2fPLCollectInOrder sr, tol, seqMode

    ' CorelDRAW's own size of what is selected, to set against what is
    ' about to be written. It is an independent measurement: it does not
    ' come from the points this module flattened.
    m_desOK = False
    On Error Resume Next
    Err.Clear
    sr.GetBoundingBox bbx, bby, bbw, bbh
    If Err.Number = 0 Then
        m_desW = bbw: m_desH = bbh: m_desOK = True
    End If
    Err.Clear
    If unitChanged Then ActiveDocument.Unit = oldUnit
    On Error GoTo Fail

    If ac2fGSub = 0 Then
        ac2fWarn "Nothing plottable in the selection." & vbCrLf & vbCrLf & _
                 "Bitmaps and objects with no path are skipped.", CAPTION_
        Exit Sub
    End If

    ' --- cutting order ------------------------------------------------
    ' Normally the Object Manager order, already set by the collection.
    ' Only Send order 2 reorders here, in document coordinates, before
    ' anything is turned.
    If seqMode = 2 Then orderAlg = ac2fBNOrderNow()

    ' --- where the head waits afterwards --------------------------------
    parkIdx = ac2fPLAddPark()

    ' --- orientation --------------------------------------------------
    ac2fPLOrient

    ' --- normalisation, from the geometry, not from parsed text -------
    minX = ac2fGX(0): maxX = ac2fGX(0): minY = ac2fGY(0): maxY = ac2fGY(0)
    For i = 1 To ac2fGN - 1
        If ac2fGX(i) < minX Then minX = ac2fGX(i)
        If ac2fGX(i) > maxX Then maxX = ac2fGX(i)
        If ac2fGY(i) < minY Then minY = ac2fGY(i)
        If ac2fGY(i) > maxY Then maxY = ac2fGY(i)
    Next i
    dx = margin * AC2F_UNITS_PER_MM - minX * AC2F_UNITS_PER_MM
    dy = margin * AC2F_UNITS_PER_MM - minY * AC2F_UNITS_PER_MM

    bytesOut = ac2fPLWrite(path, dx, dy, parkIdx)
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
                                  minX, minY, maxX, maxY, "", orderAlg, seqMode), CAPTION_
        Else
            ac2fWarn ac2fPLReport(False, host, port, path, bytesOut, margin, tol, _
                                  minX, minY, maxX, maxY, errText, orderAlg, seqMode), CAPTION_
        End If
    Else
        ac2fInfo ac2fPLReport(True, "", 0, path, bytesOut, margin, tol, _
                              minX, minY, maxX, maxY, "", orderAlg, seqMode), CAPTION_
    End If
    Exit Sub

Fail:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    ac2fWarn "The operation failed: " & Err.Description, CAPTION_
End Sub

' Flattens the selection in the order it should be cut.
'   0  the Object Manager order, top first
'   1  the Object Manager order, bottom first
'   2  selection as given; the caller then applies the bending order
' The Object Manager lists the page front to back, which is the order
' of ActivePage.Shapes. ActiveSelectionRange does not promise that
' order, so each selected object is looked up in the page stack by its
' StaticID. One that is not at the top of the page stack, an object
' picked inside a group for instance, keeps its selection order after
' the rest and is counted in m_seqMissing for the report.
Private Sub ac2fPLCollectInOrder(ByVal sr As ShapeRange, ByVal tol As Double, _
                                 ByVal mode As Long)
    Dim keyed As Collection
    Dim taken() As Boolean, ordList() As Long
    Dim n As Long, i As Long, j As Long, k As Long, cnt As Long, t As Long
    Dim ps As Shapes

    ac2fBNReset
    m_seqMissing = 0
    n = sr.Count

    If mode = 2 Then
        For i = 1 To n
            ac2fBNAddObject sr(i), tol
        Next i
        Exit Sub
    End If

    ReDim taken(1 To n)
    ReDim ordList(1 To n)
    Set keyed = New Collection

    On Error Resume Next
    For i = 1 To n
        keyed.Add i, "k" & CStr(sr(i).StaticID)
    Next i
    Set ps = ActivePage.Shapes
    For j = 1 To ps.Count
        k = 0
        k = keyed("k" & CStr(ps(j).StaticID))
        If k > 0 Then
            If Not taken(k) Then
                taken(k) = True
                cnt = cnt + 1
                ordList(cnt) = k
            End If
        End If
    Next j
    On Error GoTo 0

    For i = 1 To n
        If Not taken(i) Then
            cnt = cnt + 1
            ordList(cnt) = i
            m_seqMissing = m_seqMissing + 1
        End If
    Next i

    If mode = 1 Then
        For i = 1 To n \ 2
            t = ordList(i): ordList(i) = ordList(n + 1 - i): ordList(n + 1 - i) = t
        Next i
        ac2fBNBackToFront = True
    End If

    For i = 1 To n
        ac2fBNAddObject sr(ordList(i)), tol
    Next i
End Sub

' Adds the point the head parks on after the job: the top left corner
' of the job as drawn. It belongs to no contour, so nothing is cut to
' it, and it lies on the job's own box, so the margin is unaffected.
' Returns its index, or -1 when Park is not 1.
Private Function ac2fPLAddPark() As Long
    Dim i As Long
    Dim minX As Double, maxY As Double

    ac2fPLAddPark = -1
    If ac2fGetLng(AC2F_K_PL_PARK, AC2F_DEF_PL_PARK) <> 1 Then Exit Function
    If ac2fGN = 0 Then Exit Function

    minX = ac2fGX(0): maxY = ac2fGY(0)
    For i = 1 To ac2fGN - 1
        If ac2fGX(i) < minX Then minX = ac2fGX(i)
        If ac2fGY(i) > maxY Then maxY = ac2fGY(i)
    Next i
    ac2fBNAddPt minX, maxY
    ac2fPLAddPark = ac2fGN - 1
End Function

Public Function ac2fPLSeqName(ByVal m As Long) As String
    Select Case m
        Case 1:    ac2fPLSeqName = "Object Manager, bottom first"
        Case 2:    ac2fPLSeqName = "bending order"
        Case Else: ac2fPLSeqName = "Object Manager, top first"
    End Select
End Function

Public Function ac2fPLParkName(ByVal m As Long) As String
    Select Case m
        Case 0:    ac2fPLParkName = "origin (0,0)"
        Case 2:    ac2fPLParkName = "where the last cut ended"
        Case Else: ac2fPLParkName = "top of the job"
    End Select
End Function

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
    For i = 0 To ac2fGN - 1
        If m = 1 Or m = 3 Then ac2fGX(i) = -ac2fGX(i)
        If m = 2 Or m = 3 Then ac2fGY(i) = -ac2fGY(i)
    Next i
End Sub

Private Sub ac2fPLRotate(ByVal deg As Long)
    Dim i As Long, q As Long
    Dim t As Double

    q = ((deg Mod 360) + 360) Mod 360
    q = ((q + 45) \ 90) Mod 4              ' snap to the nearest quarter
    If q = 0 Then Exit Sub                 ' 0 = exactly as on screen

    For i = 0 To ac2fGN - 1
        Select Case q
            Case 1                          ' 90 counter-clockwise
                t = ac2fGX(i): ac2fGX(i) = -ac2fGY(i): ac2fGY(i) = t
            Case 2
                ac2fGX(i) = -ac2fGX(i): ac2fGY(i) = -ac2fGY(i)
            Case 3                          ' 270, i.e. 90 clockwise
                t = ac2fGX(i): ac2fGX(i) = ac2fGY(i): ac2fGY(i) = -t
        End Select
    Next i
End Sub

'=====================================================================
' HPGL OUTPUT
'=====================================================================

' Writes the file and returns its size in bytes, 0 on failure.
Private Function ac2fPLWrite(ByVal path As String, ByVal dx As Double, _
                             ByVal dy As Double, ByVal parkIdx As Long) As Long
    Dim f As Integer
    Dim i As Long, k As Long, s As Long, n As Long
    Dim line_ As String
    Dim cnt As Long
    Dim total As Long
    Dim perPD As Long
    Dim pre As Long
    Dim x0 As String, y0 As String, px As String, py As String
    Dim lastX As String, lastY As String

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

    For i = 0 To ac2fGSub - 1
        s = ac2fGS(i): n = ac2fGC(i)
        x0 = ac2fPLU(ac2fGX(s), dx): y0 = ac2fPLU(ac2fGY(s), dy)
        Print #f, "PU" & x0 & "," & y0 & ";"

        ' A point that rounds to the one before it adds nothing but a
        ' zero length move, which some cutters answer with a pause or a
        ' dent in the material. Drop it.
        lastX = x0: lastY = y0
        line_ = "": cnt = 0
        For k = 1 To n - 1
            px = ac2fPLU(ac2fGX(s + k), dx): py = ac2fPLU(ac2fGY(s + k), dy)
            If px <> lastX Or py <> lastY Then
                If cnt > 0 Then line_ = line_ & ","
                line_ = line_ & px & "," & py
                cnt = cnt + 1
                lastX = px: lastY = py
                If cnt >= perPD Then
                    Print #f, "PD" & line_ & ";"
                    line_ = "": cnt = 0
                End If
            End If
        Next k

        ' close the outline by returning to its first point, unless the
        ' last point already is that point
        If ac2fGCl(i) Then
            If lastX <> x0 Or lastY <> y0 Then
                If cnt > 0 Then line_ = line_ & ","
                line_ = line_ & x0 & "," & y0
                cnt = cnt + 1
            End If
        End If
        If cnt > 0 Then Print #f, "PD" & line_ & ";"
    Next i

    ' Lift the knife and park. The park point went through the same
    ' rotation, mirror and shift as the job, so "top" is the top of the
    ' job as it sits on screen.
    Select Case ac2fGetLng(AC2F_K_PL_PARK, AC2F_DEF_PL_PARK)
        Case 0
            Print #f, "PU0,0;"
        Case 1
            If parkIdx >= 0 Then
                Print #f, "PU" & ac2fPLU(ac2fGX(parkIdx), dx) & "," & _
                          ac2fPLU(ac2fGY(parkIdx), dy) & ";"
            Else
                Print #f, "PU;"
            End If
        Case Else
            Print #f, "PU;"
    End Select
    If pre >= 1 Then Print #f, "SP0;"
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
        "ORIENTATION NOW:  Rotate " & ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT) & _
        ", Mirror " & ac2fGetLng(AC2F_K_PL_MIR, AC2F_DEF_PL_MIR) & vbCrLf & _
        "Cut comes out turned or mirrored? Run the orientation wizard " & _
        "(main menu 27); it finds the setting from one test cut." & vbCrLf & _
        "By hand: add r270 (or r90) after the address, m1 / m2 to flip, " & _
        "r0 m0 to undo. Remembered.", _
        ac2fTitle(CAPTION_), stored)

    If StrPtr(answer) = 0 Then Exit Function
    answer = Trim$(answer)
    If Len(answer) = 0 Then Exit Function
    answer = ac2fPLTakeOrient(answer)
    ' only an orientation was typed: keep the address it already had
    If Len(answer) = 0 Then answer = stored

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

' Pulls "r270" / "m1" style words off the address line, stores them as
' the Rotate and Mirror settings, and returns what is left.
Private Function ac2fPLTakeOrient(ByVal answer As String) As String
    Dim parts() As String
    Dim i As Long
    Dim w As String, rest As String
    Dim v As Long

    parts = Split(answer, " ")
    For i = LBound(parts) To UBound(parts)
        w = Trim$(parts(i))
        If Len(w) >= 2 And (LCase$(Left$(w, 1)) = "r" Or LCase$(Left$(w, 1)) = "m") _
           And IsNumeric(Mid$(w, 2)) Then
            v = CLng(Val(Mid$(w, 2)))
            If LCase$(Left$(w, 1)) = "r" Then
                v = (((v Mod 360) + 360) Mod 360)
                v = (((v + 45) \ 90) Mod 4) * 90
                ac2fSetLng AC2F_K_PL_ROT, v
            ElseIf v >= 0 And v <= 3 Then
                ac2fSetLng AC2F_K_PL_MIR, v
            End If
        ElseIf Len(w) > 0 Then
            If Len(rest) > 0 Then rest = rest & " "
            rest = rest & w
        End If
    Next i
    ac2fPLTakeOrient = rest
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

' Sets the size of the file against the size CorelDRAW reports for the
' same selection. The file is turned by Rotate, so the design's width and
' height swap for a quarter turn. Outline width and effects can make
' CorelDRAW's box a little larger than the path, so the wording says
' where to look rather than calling it an error.
Private Function ac2fPLDesignCheck(ByVal fileW As Double, ByVal fileH As Double, _
                                   ByVal tol As Double) As String
    Dim q As Long, w As Double, h As Double, lim As Double, d As Double
    Dim s As String

    q = ((ac2fGetLng(AC2F_K_PL_ROT, AC2F_DEF_PL_ROT) Mod 360) + 360) Mod 360
    q = ((q + 45) \ 90) Mod 4
    If q Mod 2 = 1 Then
        w = m_desH: h = m_desW
    Else
        w = m_desW: h = m_desH
    End If
    lim = tol * 3#
    If lim < 0.3 Then lim = 0.3
    d = Abs(fileW - w)
    If Abs(fileH - h) > d Then d = Abs(fileH - h)

    s = "   Design (CorelDRAW)  : " & ac2fFmt(w) & " x " & ac2fFmt(h) & " mm" & vbCrLf
    If d <= lim Then
        s = s & "   Check               : the file is the size of the design" & vbCrLf
    Else
        s = s & "   Check               : DIFFERS by " & ac2fFmt(d) & " mm. Outline width, " & _
                "an effect or a part outside the page can do this." & vbCrLf
    End If
    ac2fPLDesignCheck = s
End Function

Private Function ac2fPLReport(ByVal ok As Boolean, ByVal host As String, _
                              ByVal port As Long, ByVal path As String, _
                              ByVal bytesOut As Long, ByVal margin As Double, _
                              ByVal tol As Double, ByVal minX As Double, _
                              ByVal minY As Double, ByVal maxX As Double, _
                              ByVal maxY As Double, ByVal errText As String, _
                              ByVal orderAlg As Long, ByVal seqMode As Long) As String
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
    s = s & "   Paths               : " & ac2fGSub & vbCrLf
    s = s & "   Points              : " & Format$(ac2fGN, "#,##0") & vbCrLf & vbCrLf

    s = s & "ORDER" & vbCrLf
    s = s & "   Sent in             : " & ac2fPLSeqName(seqMode) & vbCrLf
    If seqMode = 2 And orderAlg > 0 Then
        s = s & "   Rule                : " & ac2fBNOrderName(orderAlg) & vbCrLf
        s = s & "   Letters             : " & ac2fBNLetters() & ", " & _
                ac2fBNSetups() & " setups" & vbCrLf
    ElseIf m_seqMissing > 0 Then
        s = s & "   Not in page stack   : " & m_seqMissing & _
                " (picked inside a group; sent last, as selected)" & vbCrLf
    End If
    s = s & "   Head afterwards     : " & _
            ac2fPLParkName(ac2fGetLng(AC2F_K_PL_PARK, AC2F_DEF_PL_PARK)) & vbCrLf
    s = s & vbCrLf

    s = s & "PLACEMENT" & vbCrLf
    s = s & "   Size                : " & ac2fFmt(maxX - minX) & " x " & _
            ac2fFmt(maxY - minY) & " mm" & vbCrLf
    s = s & "   Margin              : " & ac2fFmt(margin) & " mm  (" & _
            CLng(margin * u) & " units)" & vbCrLf
    s = s & "   X range             : " & CLng(margin * u) & " .. " & _
            CLng(margin * u + (maxX - minX) * u) & vbCrLf
    s = s & "   Y range             : " & CLng(margin * u) & " .. " & _
            CLng(margin * u + (maxY - minY) * u) & vbCrLf
    If m_desOK Then s = s & ac2fPLDesignCheck(maxX - minX, maxY - minY, tol)
    s = s & "   Curve tolerance     : " & ac2fFmt(tol) & " mm  (the file never strays further)" & vbCrLf
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
        If ac2fGetLng(AC2F_K_PL_CALIB, 0) = 0 Then
            s = s & vbCrLf & vbCrLf & "ORIENTATION NOT CHECKED YET" & vbCrLf & _
                "   If the job comes out turned or mirrored, run the orientation" & vbCrLf & _
                "   wizard (main menu 27): one test cut and it sets Rotate and" & vbCrLf & _
                "   Mirror for this machine."
        End If
    End If

    ac2fPLReport = s
End Function
