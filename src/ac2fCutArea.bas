Attribute VB_Name = "ac2fCutArea"
'=====================================================================
'  ac2f pack  --  ac2fCutArea
'
'  The price of the selection by area, at the m2 rate. The companion of
'  ac2fCutTime, with the same kind of settings: a rate and a minimum.
'
'  Macro:
'    ac2fCutArea - area and price of the selection
'
'  THREE WAYS TO TAKE THE AREA
'    1  selection box     the box round everything selected: one panel
'    2  each object's box  the boxes of the selected objects added up:
'                          separate parts, letters, each priced alone
'    3  nest sheets        Nesting's sheet outlines on the page x their
'                          size: the sheets the job really takes
'  The first is used unless the page has nest sheets, then the third.
'  1, 2 or 3 switches between them.
'
'  PRICE
'    billed area = the larger of area x copies and the minimum
'    price       = billed area x m2 rate
'
'  Rate and minimum come from StockTrack (Settings > Cut pricing); the
'  last values read are kept for offline use. For one run: a=450 m=0.5
'  x=2 (copies).
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Area Price"

Public Sub ac2fCutArea()
Attribute ac2fCutArea.VB_Description = "ac2f pack: Area of the selection in m2 and its price at the m2 rate"
    Dim sr As ShapeRange
    Dim v() As Double
    Dim cur As String, src As String, m As String, answer As String, errText As String
    Dim areas(1 To 3) As Double, notes(1 To 3) As String
    Dim mode As Long, copies As Double, rate As Double, minArea As Double

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Select what is to be priced.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Select what is to be priced.", CAPTION_
        Exit Sub
    End If

    If Not ac2fCAMeasure(sr, areas, notes, errText) Then
        ac2fWarn errText, CAPTION_
        Exit Sub
    End If

    ac2fJobPricing v, cur, src
    rate = v(3): minArea = v(4)
    copies = 1
    mode = 1
    If areas(3) > 0 Then mode = 3

    Do
        m = ac2fCAReport(areas, notes, mode, copies, rate, minArea, cur, src)
        answer = InputBox(m, ac2fTitle(CAPTION_), "")
        If StrPtr(answer) = 0 Then Exit Sub
        answer = Trim$(answer)
        If Len(answer) = 0 Then Exit Sub
        If answer = "1" Or answer = "2" Or answer = "3" Then
            If areas(CLng(answer)) > 0 Then
                mode = CLng(answer)
            Else
                ac2fWarn "There are no nest sheet outlines on the page.", CAPTION_
            End If
        ElseIf Not ac2fCATake(answer, rate, minArea, copies, errText) Then
            ac2fWarn errText, CAPTION_
        End If
    Loop
End Sub

' areas(1) selection box, (2) each object's box added up, (3) nest
' sheets; m2. notes() say what each was made of.
Private Function ac2fCAMeasure(ByVal sr As ShapeRange, ByRef areas() As Double, _
                               ByRef notes() As String, ByRef errText As String) As Boolean
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim x As Double, y As Double, w As Double, h As Double
    Dim i As Long, ns As Long
    Dim sw As Double, sh As Double, sum As Double
    Dim ok As Boolean

    On Error GoTo Fail
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    sr.GetBoundingBox x, y, w, h
    areas(1) = w * h / 1000000#
    notes(1) = Format$(w, "0") & " x " & Format$(h, "0") & " mm"

    For i = 1 To sr.Count
        sr(i).GetBoundingBox x, y, w, h
        sum = sum + w * h
    Next i
    areas(2) = sum / 1000000#
    notes(2) = sr.Count & " object(s)"

    ac2fCASheets ActivePage.Shapes, ns, sw, sh
    If ns > 0 Then
        areas(3) = ns * sw * sh / 1000000#
        notes(3) = ns & " x " & Format$(sw, "0") & "x" & Format$(sh, "0")
    End If
    ok = areas(1) > 0
    If Not ok Then errText = "The selection has no area."
    ac2fCAMeasure = ok
Done:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Exit Function

Fail:
    errText = "Could not measure the selection: " & Err.Description
    Resume Done
End Function

' Counts Nesting's sheet outlines; sw, sh return the size of the first.
Private Sub ac2fCASheets(ByVal coll As Object, ByRef ns As Long, _
                         ByRef sw As Double, ByRef sh As Double)
    Dim s As Shape
    Dim i As Long
    Dim nm As String
    Dim x As Double, y As Double, w As Double, h As Double

    On Error Resume Next
    For i = 1 To coll.Count
        Set s = Nothing
        Set s = coll(i)
        nm = ""
        nm = LCase$(s.Name)
        If nm = "ac2f sheets" Then
            ac2fCASheets s.Shapes, ns, sw, sh
        ElseIf nm Like "ac2f sheet #*" Then
            s.GetBoundingBox x, y, w, h
            If w > 0 And h > 0 Then
                ns = ns + 1
                If ns = 1 Then
                    sw = w
                    sh = h
                End If
            End If
        End If
    Next i
End Sub

Private Function ac2fCAReport(ByRef areas() As Double, ByRef notes() As String, _
                              ByVal mode As Long, ByVal copies As Double, _
                              ByVal rate As Double, ByVal minArea As Double, _
                              ByVal cur As String, ByVal src As String) As String
    Dim labels(1 To 3) As String
    Dim m As String, who As String, mark As String
    Dim i As Long
    Dim area As Double, billed As Double

    labels(1) = "Selection box"
    labels(2) = "Each object's box"
    labels(3) = "Nest sheets"

    who = ac2fJobCustomerName()
    m = "AREA PRICE"
    If Len(who) > 0 Then m = m & "  -  " & who
    m = m & vbCrLf & vbCrLf
    For i = 1 To 3
        If areas(i) > 0 Then
            If i = mode Then mark = "  <- used" Else mark = ""
            m = m & i & "  " & ac2fPad(labels(i), 18) & ac2fPad(notes(i), 20) & _
                ac2fRPad(ac2fFmt(areas(i), 3), 9) & " m2" & mark & vbCrLf
        End If
    Next i

    area = areas(mode) * copies
    billed = area
    If billed < minArea Then billed = minArea
    m = m & vbCrLf
    If copies <> 1 Then m = m & "Copies    " & ac2fNumStr(copies) & "  ->  " & ac2fFmt(area, 3) & " m2" & vbCrLf
    If rate > 0 Then
        m = m & "Billed    " & ac2fFmt(billed, 3) & " m2"
        If billed > area Then m = m & "  (minimum " & ac2fFmt(minArea, 3) & ")"
        m = m & vbCrLf
        m = m & "Price     " & ac2fFmt(billed, 3) & " m2 x " & ac2fFmt(rate, 2) & " " & cur & _
            " = " & ac2fFmt(billed * rate, 2) & " " & cur & vbCrLf
    Else
        m = m & "No m2 rate yet: StockTrack, Settings > Cut pricing." & vbCrLf
    End If
    m = m & vbCrLf & "Values: " & src & vbCrLf
    m = m & "1 / 2 / 3 picks the area. For this run:" & vbCrLf
    m = m & "a=m2 rate m=minimum m2 x=copies, e.g.  x=4" & vbCrLf
    m = m & "Enter closes."
    ac2fCAReport = m
End Function

' Reads "a=450 m=0.5 x=2" into the values.
Private Function ac2fCATake(ByVal txt As String, ByRef rate As Double, _
                            ByRef minArea As Double, ByRef copies As Double, _
                            ByRef errText As String) As Boolean
    Dim words() As String
    Dim i As Long, p As Long
    Dim key As String, num As Double

    words = Split(txt, " ")
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then
            p = InStr(words(i), "=")
            If p < 2 Then
                errText = "Write key=value, e.g. x=2 (got """ & words(i) & """)."
                Exit Function
            End If
            key = LCase$(Left$(words(i), p - 1))
            num = ac2fParseNum(Mid$(words(i), p + 1), -1)
            If num < 0 Then
                errText = "Not a number: """ & words(i) & """."
                Exit Function
            End If
            Select Case key
                Case "a": rate = num
                Case "m": minArea = num
                Case "x": If num > 0 Then copies = num
                Case Else
                    errText = "No value called """ & key & """: a m x."
                    Exit Function
            End Select
        End If
    Next i
    ac2fCATake = True
End Function
