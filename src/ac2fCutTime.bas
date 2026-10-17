Attribute VB_Name = "ac2fCutTime"
'=====================================================================
'  ac2f pack  --  ac2fCutTime
'
'  How long the selection takes to cut, and what that costs at the
'  hourly rate.
'
'  Macro:
'    ac2fCutTime - time and price of the selection
'
'  TIME
'    cutting  = outline length x passes / feed
'    travel   = moves between outlines / rapid
'    plunges  = outlines x passes x plunge time
'
'  Outline lengths, and where every outline starts and ends, come from
'  the flattening the plotter uses (ac2fBending). Travel is the nearest
'  neighbour tour: from the bottom left corner of the job to the nearest
'  outline, round it, on to the next nearest, and back to the corner.
'  Open outlines may be entered from either end. A CAM program's path
'  is usually a little shorter; the figure is meant to be rough, and on
'  the safe side. Past MAX_TOUR outlines the tour is estimated from the
'  area of the job instead (0.71 x sqrt(outlines x area), the usual
'  figure for a tour through scattered points).
'
'  Passes are the thickness of the job's sheet over the pass depth,
'  rounded up; 1 when either is unknown or the pass depth is 0.
'
'  PRICE
'    billed minutes = the larger of the minutes and the minimum
'    price          = billed minutes / 60 x hourly rate
'
'  Rate, minimum and speeds come from StockTrack (Settings > Cut
'  pricing), the same values a quote priced by time uses. The last
'  values read are kept, so this works offline. Any of them can be
'  changed for one run: f=30 r=150 p=1 n=2 h=1800 m=20.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Cut Time"
Private Const MAX_TOUR As Long = 2000

'=====================================================================
' MACRO
'=====================================================================

Public Sub ac2fCutTime()
Attribute ac2fCutTime.VB_Description = "ac2f pack: Cutting time of the selection and its price at the hourly rate"
    Dim sr As ShapeRange
    Dim v() As Double
    Dim cur As String, src As String, m As String, answer As String
    Dim errText As String, passNote As String, kind As String
    Dim cutLen As Double, travel As Double, nCont As Long
    Dim feed As Double, rapid As Double, plunge As Double, depth As Double
    Dim rate As Double, minMin As Double, passes As Double
    Dim sw As Double, sh As Double, thick As Double

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Select what is to be cut.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Select what is to be cut.", CAPTION_
        Exit Sub
    End If

    If Not ac2fCTMeasure(sr, cutLen, travel, nCont, errText) Then
        ac2fWarn errText, CAPTION_
        Exit Sub
    End If

    ac2fJobPricing v, cur, src
    rate = v(1): minMin = v(2)
    feed = v(5): rapid = v(6): plunge = v(7): depth = v(8)

    passes = 1
    If ac2fJobSheetInfo(sw, sh, thick, kind) Then
        If depth > 0 And thick > 0 Then
            passes = ac2fCeil(thick / depth)
            passNote = kind & " " & ac2fNumStr(thick) & " mm at " & _
                       ac2fNumStr(depth) & " mm a pass"
        End If
    End If

    Do
        m = ac2fCTReport(cutLen, travel, nCont, feed, rapid, plunge, passes, _
                         rate, minMin, cur, src, passNote)
        answer = InputBox(m, ac2fTitle(CAPTION_), "")
        If StrPtr(answer) = 0 Then Exit Sub
        answer = Trim$(answer)
        If Len(answer) = 0 Then Exit Sub
        If Not ac2fCTTake(answer, feed, rapid, plunge, passes, rate, minMin, errText) Then
            ac2fWarn errText, CAPTION_
        Else
            passNote = "changed for this run"
        End If
    Loop
End Sub

'=====================================================================
' MEASURING
'=====================================================================

' Total outline length, travel between outlines and how many outlines,
' all in mm, from the flattened selection.
Private Function ac2fCTMeasure(ByVal sr As ShapeRange, ByRef cutLen As Double, _
                               ByRef travel As Double, ByRef nCont As Long, _
                               ByRef errText As String) As Boolean
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim sx() As Double, sy() As Double, ex() As Double, ey() As Double
    Dim cl() As Boolean
    Dim i As Long, k As Long, s As Long, n As Long
    Dim ox As Double, oy As Double, mx As Double, my As Double
    Dim seg As Double

    On Error GoTo Fail
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    ac2fBNCollect sr, ac2fGetNum(AC2F_K_PL_TOL, AC2F_DEF_PL_TOL)
    nCont = ac2fGSub
    If nCont = 0 Or ac2fGN = 0 Then
        errText = "Nothing in the selection has an outline to cut."
        GoTo Done
    End If

    ReDim sx(0 To nCont - 1): ReDim sy(0 To nCont - 1)
    ReDim ex(0 To nCont - 1): ReDim ey(0 To nCont - 1)
    ReDim cl(0 To nCont - 1)
    ox = ac2fGX(0): oy = ac2fGY(0): mx = ox: my = oy
    For i = 0 To nCont - 1
        s = ac2fGS(i): n = ac2fGC(i)
        For k = 0 To n - 1
            If ac2fGX(s + k) < ox Then ox = ac2fGX(s + k)
            If ac2fGY(s + k) < oy Then oy = ac2fGY(s + k)
            If ac2fGX(s + k) > mx Then mx = ac2fGX(s + k)
            If ac2fGY(s + k) > my Then my = ac2fGY(s + k)
            If k < n - 1 Then
                seg = ac2fCTDist(ac2fGX(s + k), ac2fGY(s + k), ac2fGX(s + k + 1), ac2fGY(s + k + 1))
                cutLen = cutLen + seg
            End If
        Next k
        cl(i) = ac2fGCl(i)
        sx(i) = ac2fGX(s): sy(i) = ac2fGY(s)
        ex(i) = ac2fGX(s + n - 1): ey(i) = ac2fGY(s + n - 1)
        If cl(i) Then cutLen = cutLen + ac2fCTDist(ex(i), ey(i), sx(i), sy(i))
    Next i

    If nCont > MAX_TOUR Then
        travel = 0.71 * Sqr(nCont * (mx - ox) * (my - oy))
    Else
        travel = ac2fCTTour(sx, sy, ex, ey, cl, nCont, ox, oy)
    End If
    ac2fCTMeasure = True

Done:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Exit Function

Fail:
    errText = "Could not measure the selection: " & Err.Description
    Resume Done
End Function

' Nearest neighbour tour from (ox, oy) through every outline and back.
Private Function ac2fCTTour(ByRef sx() As Double, ByRef sy() As Double, _
                            ByRef ex() As Double, ByRef ey() As Double, _
                            ByRef cl() As Boolean, ByVal n As Long, _
                            ByVal ox As Double, ByVal oy As Double) As Double
    Dim done() As Boolean
    Dim i As Long, stp As Long, best As Long
    Dim cx As Double, cy As Double, d As Double, bestD As Double, total As Double
    Dim rev As Boolean, bestRev As Boolean

    ReDim done(0 To n - 1)
    cx = ox: cy = oy
    For stp = 1 To n
        best = -1
        bestD = 1E+300
        For i = 0 To n - 1
            If Not done(i) Then
                d = (sx(i) - cx) * (sx(i) - cx) + (sy(i) - cy) * (sy(i) - cy)
                rev = False
                If Not cl(i) Then
                    If (ex(i) - cx) * (ex(i) - cx) + (ey(i) - cy) * (ey(i) - cy) < d Then
                        d = (ex(i) - cx) * (ex(i) - cx) + (ey(i) - cy) * (ey(i) - cy)
                        rev = True
                    End If
                End If
                If d < bestD Then
                    bestD = d
                    best = i
                    bestRev = rev
                End If
            End If
        Next i
        If best < 0 Then Exit For
        total = total + Sqr(bestD)
        done(best) = True
        ' a closed outline ends where it began; an open one at its far end
        If cl(best) Then
            cx = sx(best): cy = sy(best)
        ElseIf bestRev Then
            cx = sx(best): cy = sy(best)
        Else
            cx = ex(best): cy = ey(best)
        End If
    Next stp
    ac2fCTTour = total + ac2fCTDist(cx, cy, ox, oy)
End Function

Private Function ac2fCTDist(ByVal x1 As Double, ByVal y1 As Double, _
                            ByVal x2 As Double, ByVal y2 As Double) As Double
    ac2fCTDist = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
End Function

'=====================================================================
' REPORT
'=====================================================================

Private Function ac2fCTReport(ByVal cutLen As Double, ByVal travel As Double, _
                              ByVal nCont As Long, ByVal feed As Double, _
                              ByVal rapid As Double, ByVal plunge As Double, _
                              ByVal passes As Double, ByVal rate As Double, _
                              ByVal minMin As Double, ByVal cur As String, _
                              ByVal src As String, ByVal passNote As String) As String
    Dim tCut As Double, tTravel As Double, tPlunge As Double, total As Double
    Dim minutes As Double, billed As Double
    Dim m As String, who As String

    If feed > 0 Then tCut = cutLen * passes / feed
    If rapid > 0 Then tTravel = travel / rapid
    tPlunge = nCont * passes * plunge
    total = tCut + tTravel + tPlunge
    minutes = total / 60#
    billed = minutes
    If billed < minMin Then billed = minMin

    who = ac2fJobCustomerName()
    m = "CUT TIME"
    If Len(who) > 0 Then m = m & "  -  " & who
    m = m & vbCrLf & nCont & " outline(s), " & ac2fNumStr(passes) & " pass(es)"
    If Len(passNote) > 0 Then m = m & "  (" & passNote & ")"
    m = m & vbCrLf & vbCrLf

    m = m & ac2fPad("Cutting", 10) & ac2fPad(ac2fFmt(cutLen, 0) & " mm x" & ac2fNumStr(passes) & _
            " at " & ac2fNumStr(feed) & " mm/s", 32) & ac2fCTDur(tCut) & vbCrLf
    m = m & ac2fPad("Travel", 10) & ac2fPad(ac2fFmt(travel, 0) & " mm at " & _
            ac2fNumStr(rapid) & " mm/s", 32) & ac2fCTDur(tTravel) & vbCrLf
    m = m & ac2fPad("Plunges", 10) & ac2fPad(nCont & " x" & ac2fNumStr(passes) & " x " & _
            ac2fNumStr(plunge) & " s", 32) & ac2fCTDur(tPlunge) & vbCrLf
    m = m & ac2fPad("Total", 42) & ac2fCTDur(total) & vbCrLf & vbCrLf

    If rate > 0 Then
        m = m & "Billed    " & ac2fFmt(billed, 0) & " min"
        If billed > minutes Then m = m & "  (minimum " & ac2fFmt(minMin, 0) & ")"
        m = m & vbCrLf
        m = m & "Price     " & ac2fFmt(billed, 0) & " min x " & ac2fFmt(rate, 2) & " " & cur & _
            "/h = " & ac2fFmt(billed / 60# * rate, 2) & " " & cur & vbCrLf
    Else
        m = m & "No hourly rate yet: StockTrack, Settings > Cut pricing." & vbCrLf
    End If
    m = m & vbCrLf & "Values: " & src & vbCrLf
    m = m & "For this run: f=feed r=rapid p=plunge s n=passes" & vbCrLf
    m = m & "h=hourly rate m=minimum min, e.g.  f=30 n=1" & vbCrLf
    m = m & "Enter closes."
    ac2fCTReport = m
End Function

' "16 min 36 s", "52 s", "1 h 05 min".
Private Function ac2fCTDur(ByVal sec As Double) As String
    Dim t As Long, h As Long, mi As Long, s As Long
    t = CLng(sec + 0.5)
    h = t \ 3600
    mi = (t Mod 3600) \ 60
    s = t Mod 60
    If h > 0 Then
        ac2fCTDur = h & " h " & Format$(mi, "00") & " min"
    ElseIf mi > 0 Then
        ac2fCTDur = mi & " min " & Format$(s, "00") & " s"
    Else
        ac2fCTDur = s & " s"
    End If
End Function

' Reads "f=30 r=150 p=1 n=2 h=1800 m=20" into the values.
Private Function ac2fCTTake(ByVal txt As String, ByRef feed As Double, ByRef rapid As Double, _
                            ByRef plunge As Double, ByRef passes As Double, _
                            ByRef rate As Double, ByRef minMin As Double, _
                            ByRef errText As String) As Boolean
    Dim words() As String
    Dim i As Long, p As Long
    Dim key As String, num As Double

    words = Split(Replace$(txt, ",", " "), " ")
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then
            p = InStr(words(i), "=")
            If p < 2 Then
                errText = "Write key=value, e.g. f=30 (got """ & words(i) & """)."
                Exit Function
            End If
            key = LCase$(Left$(words(i), p - 1))
            num = ac2fParseNum(Mid$(words(i), p + 1), -1)
            If num < 0 Then
                errText = "Not a number: """ & words(i) & """."
                Exit Function
            End If
            Select Case key
                Case "f": If num > 0 Then feed = num
                Case "r": If num > 0 Then rapid = num
                Case "p": plunge = num
                Case "n": If num >= 1 Then passes = Int(num)
                Case "h": rate = num
                Case "m": minMin = num
                Case Else
                    errText = "No value called """ & key & """: f r p n h m."
                    Exit Function
            End Select
        End If
    Next i
    ac2fCTTake = True
End Function
