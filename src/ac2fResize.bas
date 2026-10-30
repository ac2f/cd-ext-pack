Attribute VB_Name = "ac2fResize"
'=====================================================================
'  ac2f pack  --  ac2fResize
'
'  Resize each: gives EVERY selected object a size you type, one by
'  one. CorelDRAW's own size boxes treat a multiple selection as a
'  single block, so typing a width scales all objects by the same
'  factor around the biggest box.
'
'  Macros:
'    ac2fResizeEach       - asks for the size, then resizes the selection
'    ac2fResizeEachAgain  - same with the last answer, no question
'
'  THE ANSWER  (sizes in mm; add cm, m or in)
'    50x30       exactly 50 wide and 30 high (stretches the object)
'    w50 / h30   one side given, the other follows (shape kept)
'    fit 50x30   shape kept, fits inside the box; "fit 50" = 50 x 50
'    @7          the point of each object that stays where it was,
'                keypad style (7 8 9 / 4 5 6 / 1 2 3); default @5
'    deep        objects inside groups are resized one by one
'
'  HOW
'  The size is read and set on the object's bounding box, the same box
'  every other macro of the pack measures. CorelDRAW may leave a macro
'  resize anchored somewhere else or count the outline in, so the size
'  is read back and corrected, then the object is moved so the chosen
'  point sits exactly where it was.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Resize Each"
Private Const FLAT_ As Double = 0.001   ' a side this short cannot be scaled (mm)
Private Const TOL_ As Double = 0.01     ' a size this close counts as reached (mm)

'---------------------------------------------------------------------
' Registry key (internal, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_RZ_LAST   As String = "RZSon"
Public Const AC2F_DEF_RZ_LAST As String = "50x50"

' the answer of this run
Private m_w As Double, m_h As Double    ' mm, 0 = not given
Private m_fit As Boolean, m_deep As Boolean
Private m_pos As Long                   ' keypad position that stays put

' counters for the report
Private m_done As Long, m_same As Long, m_flat As Long
Private m_off As Long, m_failed As Long
Private m_why As String

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fResizeEach()
Attribute ac2fResizeEach.VB_Description = "ac2f pack: Resize every selected object to a size you type"
    Dim ans As String
    Dim why As String

    If Not ac2fRZCheck() Then Exit Sub
    ans = ac2fGetStr(AC2F_K_RZ_LAST, AC2F_DEF_RZ_LAST)
    Do
        ans = InputBox(ac2fRZPrompt(ActiveSelectionRange.Count), ac2fTitle(CAPTION_), ans)
        If StrPtr(ans) = 0 Then Exit Sub
        ans = Trim$(ans)
        If Len(ans) = 0 Then Exit Sub
        If ac2fRZParse(ans, why) Then Exit Do
        ac2fWarn why, CAPTION_
    Loop
    ac2fSetStr AC2F_K_RZ_LAST, ans
    ac2fRZRun
End Sub

Public Sub ac2fResizeEachAgain()
Attribute ac2fResizeEachAgain.VB_Description = "ac2f pack: Resize every selected object with the last size"
    Dim last As String
    Dim why As String

    If Not ac2fRZCheck() Then Exit Sub
    last = ac2fGetStr(AC2F_K_RZ_LAST, "")
    If Len(last) = 0 Then
        ac2fWarn "There is no earlier size yet." & vbCrLf & "Run ac2fResizeEach first.", CAPTION_
        Exit Sub
    End If
    If Not ac2fRZParse(last, why) Then
        ac2fWarn "The saved size cannot be read:" & vbCrLf & why, CAPTION_
        Exit Sub
    End If
    ac2fRZRun
End Sub

'=====================================================================
' FLOW
'=====================================================================

Private Function ac2fRZCheck() As Boolean
    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Function
    End If
    If ActiveSelectionRange Is Nothing Then
        ac2fWarn "Select the objects first.", CAPTION_
        Exit Function
    End If
    If ActiveSelectionRange.Count = 0 Then
        ac2fWarn "Select the objects first.", CAPTION_
        Exit Function
    End If
    ac2fRZCheck = True
End Function

Private Sub ac2fRZRun()
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean
    Dim i As Long
    Dim s As String

    Set sr = ActiveSelectionRange
    m_done = 0: m_same = 0: m_flat = 0: m_off = 0: m_failed = 0: m_why = ""
    On Error GoTo Fail
    oldOpt = Application.Optimization
    Application.Optimization = True
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If
    ActiveDocument.BeginCommandGroup ac2fTitle("resize each")

    For i = 1 To sr.Count
        ac2fRZShape sr(i)
    Next i

    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh

    s = "Resized            : " & m_done & " object(s)" & vbCrLf
    If m_same > 0 Then s = s & "Already that size  : " & m_same & vbCrLf
    If m_flat > 0 Then
        s = s & "Left alone         : " & m_flat & " object(s) with no width or height" & vbCrLf & _
                "  to scale (a straight line cannot be given the other side)." & vbCrLf
    End If
    If m_off > 0 Then
        s = s & "Not exactly reached: " & m_off & " object(s), more than " & TOL_ & " mm off" & vbCrLf & _
                "  (rotated or skewed objects are sized by their bounding box)." & vbCrLf
    End If
    If m_failed > 0 Then
        s = s & "Could not resize   : " & m_failed & " object(s): " & m_why & vbCrLf
    End If
    s = s & vbCrLf & ac2fRZDescribe() & vbCrLf & vbCrLf & "Ctrl+Z undoes it."
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

' One selected object; with "deep" a group is opened and its members
' are the objects.
Private Sub ac2fRZShape(ByVal s As Shape)
    Dim i As Long

    If m_deep Then
        If s.Type = cdrGroupShape Then
            For i = 1 To s.Shapes.Count
                ac2fRZShape s.Shapes(i)
            Next i
            Exit Sub
        End If
    End If
    ac2fRZOne s
End Sub

'=====================================================================
' ONE OBJECT
'=====================================================================

Private Sub ac2fRZOne(ByVal s As Shape)
    Dim x As Double, y As Double, w As Double, h As Double
    Dim nx As Double, ny As Double, aw As Double, ah As Double
    Dim nw As Double, nh As Double, tw As Double, th As Double
    Dim px As Double, py As Double, qx As Double, qy As Double
    Dim k As Long

    On Error GoTo Fail
    s.GetBoundingBox x, y, w, h
    If ac2fRZPlan(w, h, nw, nh) <> 0 Then
        m_flat = m_flat + 1
        Exit Sub
    End If
    If Abs(nw - w) <= TOL_ And Abs(nh - h) <= TOL_ Then
        m_same = m_same + 1
        Exit Sub
    End If

    ac2fRZAnchor x, y, w, h, px, py
    tw = nw: th = nh
    For k = 1 To 3
        ac2fRZSetSize s, tw, th
        s.GetBoundingBox nx, ny, aw, ah
        If Abs(aw - nw) <= TOL_ And Abs(ah - nh) <= TOL_ Then Exit For
        ' ask for what is missing (an outline counted in, a different anchor)
        tw = tw + (nw - aw)
        th = th + (nh - ah)
    Next k
    If Abs(aw - nw) > TOL_ Or Abs(ah - nh) > TOL_ Then m_off = m_off + 1

    ac2fRZAnchor nx, ny, aw, ah, qx, qy
    If Abs(px - qx) > 0.0001 Or Abs(py - qy) > 0.0001 Then s.Move px - qx, py - qy
    m_done = m_done + 1
    Exit Sub

Fail:
    m_failed = m_failed + 1
    If Len(m_why) = 0 Then m_why = Err.Description
End Sub

' The size one object should get. Returns 0, or 1 when the object has no
' width or height that the answer needs to scale.
Private Function ac2fRZPlan(ByVal cw As Double, ByVal ch As Double, _
                            ByRef nw As Double, ByRef nh As Double) As Long
    Dim flatW As Boolean, flatH As Boolean
    Dim k As Double

    flatW = (cw < FLAT_)
    flatH = (ch < FLAT_)
    If m_w > 0 And m_h > 0 And Not m_fit Then
        If flatW Or flatH Then
            ac2fRZPlan = 1
            Exit Function
        End If
        nw = m_w
        nh = m_h
        Exit Function
    End If

    ' shape kept: the tightest of the limits that apply decides
    k = 0
    If m_w > 0 And Not flatW Then k = m_w / cw
    If m_h > 0 And Not flatH Then
        If k = 0 Then
            k = m_h / ch
        ElseIf m_h / ch < k Then
            k = m_h / ch
        End If
    End If
    If k <= 0 Then
        ac2fRZPlan = 1
        Exit Function
    End If
    nw = cw * k
    nh = ch * k
End Function

' Sets the bounding box size. A straight line only has one side to set.
Private Sub ac2fRZSetSize(ByVal s As Shape, ByVal tw As Double, ByVal th As Double)
    If th < FLAT_ Then
        s.SizeWidth = tw
    ElseIf tw < FLAT_ Then
        s.SizeHeight = th
    Else
        s.SetSize tw, th
    End If
End Sub

' The point of a box that stays in place (keypad position m_pos).
Private Sub ac2fRZAnchor(ByVal x As Double, ByVal y As Double, ByVal w As Double, _
                         ByVal h As Double, ByRef ax As Double, ByRef ay As Double)
    ax = x + w * ((m_pos - 1) Mod 3) / 2#
    ay = y + h * ((m_pos - 1) \ 3) / 2#
End Sub

'=====================================================================
' THE ANSWER
'=====================================================================

' Reads the answer into m_w, m_h, m_fit, m_deep and m_pos. When it cannot
' be read, says why and returns False.
Private Function ac2fRZParse(ByVal ans As String, ByRef why As String) As Boolean
    Dim tok() As String
    Dim n As Long, i As Long, p As Long
    Dim t As String, a As String, b As String
    Dim v As Double, bare As Double

    m_w = 0: m_h = 0: m_fit = False: m_deep = False: m_pos = 5
    n = ac2fRZTokens(ans, tok)
    For i = 1 To n
        t = tok(i)
        If t = "fit" Then
            m_fit = True
        ElseIf t = "deep" Then
            m_deep = True
        ElseIf Left$(t, 1) = "@" Then
            a = Mid$(t, 2)
            If Len(a) <> 1 Then
                why = "'" & t & "': after @ type one digit from 1 to 9."
                Exit Function
            End If
            If InStr("123456789", a) = 0 Then
                why = "'" & t & "': after @ type one digit from 1 to 9."
                Exit Function
            End If
            m_pos = CLng(a)
        ElseIf Left$(t, 1) = "w" Or Left$(t, 1) = "h" Then
            v = ac2fRZLen(Mid$(t, 2))
            If v <= 0 Then
                why = "'" & t & "': cannot read a size above zero."
                Exit Function
            End If
            If Left$(t, 1) = "w" Then m_w = v Else m_h = v
        ElseIf InStr(t, "x") > 0 Then
            p = InStr(t, "x")
            a = Left$(t, p - 1)
            b = Mid$(t, p + 1)
            If Len(a) = 0 And Len(b) = 0 Then
                why = "'" & t & "': type a width, a height or both."
                Exit Function
            End If
            If Len(a) > 0 Then
                v = ac2fRZLen(a)
                If v <= 0 Then
                    why = "'" & a & "': cannot read a width above zero."
                    Exit Function
                End If
                m_w = v
            End If
            If Len(b) > 0 Then
                v = ac2fRZLen(b)
                If v <= 0 Then
                    why = "'" & b & "': cannot read a height above zero."
                    Exit Function
                End If
                m_h = v
            End If
        Else
            bare = ac2fRZLen(t)
            If bare <= 0 Then
                why = "'" & t & "': cannot read this." & vbCrLf & _
                      "Use 50x30, w50, h30, fit 50x30, @7 or deep."
                Exit Function
            End If
        End If
    Next i

    If bare > 0 Then
        If m_fit And m_w = 0 And m_h = 0 Then
            m_w = bare
            m_h = bare
        Else
            why = "A lone number does not say which side it is." & vbCrLf & _
                  "Use w" & ac2fNumStr(bare) & " (width), h" & ac2fNumStr(bare) & _
                  " (height), " & ac2fNumStr(bare) & "x" & ac2fNumStr(bare) & _
                  " or fit " & ac2fNumStr(bare) & " (longest side)."
            Exit Function
        End If
    End If
    If m_w = 0 And m_h = 0 Then
        why = "Type a size, for example 50x30, w50, h30 or fit 50x30."
        Exit Function
    End If
    ac2fRZParse = True
End Function

' Lower case words. A unit written apart ("50 mm") and an x written apart
' ("50 x 30") are joined to the number they belong to.
Private Function ac2fRZTokens(ByVal ans As String, ByRef tok() As String) As Long
    Dim raw() As String
    Dim i As Long, n As Long
    Dim t As String
    Dim joined As Boolean

    ans = LCase$(ans)
    ans = Replace$(ans, "*", "x")
    ans = Replace$(ans, ChrW$(215), "x")
    ans = Replace$(ans, ";", " ")
    ans = Replace$(ans, vbTab, " ")
    raw = Split(ans, " ")
    ReDim tok(1 To UBound(raw) + 2)
    n = 0
    For i = 0 To UBound(raw)
        t = raw(i)
        If Len(t) > 0 Then
            joined = False
            If n > 0 Then joined = ac2fRZJoins(tok(n), t)    ' And would read tok(0)
            If joined Then
                tok(n) = tok(n) & t
            Else
                n = n + 1
                tok(n) = t
            End If
        End If
    Next i
    ac2fRZTokens = n
End Function

' Does word t continue the word before it?  "50" + "mm", "50" + "x",
' "50x" + "30", "50x30" + "cm", "w" + "50", "@" + "7".
Private Function ac2fRZJoins(ByVal prev As String, ByVal t As String) As Boolean
    Dim last As String, first As String

    last = Right$(prev, 1)
    first = Left$(t, 1)
    If t = "mm" Or t = "cm" Or t = "m" Or t = "in" Then
        ac2fRZJoins = (InStr("0123456789.", last) > 0)
    ElseIf prev = "w" Or prev = "h" Or prev = "@" Then
        ac2fRZJoins = (InStr("0123456789.", first) > 0)
    ElseIf last = "x" Then
        ac2fRZJoins = (InStr("0123456789.", first) > 0)
    ElseIf first = "x" Then
        ac2fRZJoins = (InStr("0123456789.mn", last) > 0)
    End If
End Function

' "50", "5cm", "12,5mm", "2m", "3in" -> mm; 0 when it cannot be read.
Private Function ac2fRZLen(ByVal s As String) As Double
    Dim t As String
    Dim mult As Double
    Dim v As Double

    t = Trim$(s)
    mult = 1#
    If Right$(t, 2) = "mm" Then
        t = Left$(t, Len(t) - 2)
    ElseIf Right$(t, 2) = "cm" Then
        mult = 10#
        t = Left$(t, Len(t) - 2)
    ElseIf Right$(t, 2) = "in" Then
        mult = 25.4
        t = Left$(t, Len(t) - 2)
    ElseIf Right$(t, 1) = "m" Then
        mult = 1000#
        t = Left$(t, Len(t) - 1)
    End If
    If Len(t) = 0 Then Exit Function
    v = ac2fParseNum(t, -1#)
    If v <= 0 Then Exit Function
    ac2fRZLen = v * mult
End Function

Private Function ac2fRZPrompt(ByVal n As Long) As String
    Dim s As String

    s = n & " object(s) selected. Each one gets its own size" & vbCrLf & _
        "(CorelDRAW's size boxes treat the selection as one block)." & vbCrLf & vbCrLf
    s = s & "  50x30       exactly 50 wide and 30 high" & vbCrLf
    s = s & "  w50         width 50, height follows" & vbCrLf
    s = s & "  h30         height 30, width follows" & vbCrLf
    s = s & "  fit 50x30   keep the shape, fit inside 50 x 30" & vbCrLf
    s = s & "  fit 50      keep the shape, longest side 50" & vbCrLf & vbCrLf
    s = s & "Sizes are mm; add cm, m or in:  12cm x 4cm" & vbCrLf & vbCrLf
    s = s & "After the size, if you want:" & vbCrLf
    s = s & "  @7          this point stays put: 7 8 9 / 4 5 6 / 1 2 3" & vbCrLf
    s = s & "              (default @5, the centre of each object)" & vbCrLf
    s = s & "  deep        objects inside groups, one by one"
    ac2fRZPrompt = s
End Function

Private Function ac2fRZDescribe() As String
    Dim s As String

    If m_w > 0 And m_h > 0 Then
        If m_fit Then
            s = "fit inside " & ac2fFmt(m_w) & " x " & ac2fFmt(m_h) & " mm, shape kept"
        Else
            s = "exactly " & ac2fFmt(m_w) & " x " & ac2fFmt(m_h) & " mm"
        End If
    ElseIf m_w > 0 Then
        s = "width " & ac2fFmt(m_w) & " mm, height follows"
    Else
        s = "height " & ac2fFmt(m_h) & " mm, width follows"
    End If
    s = "Size: " & s & vbCrLf & "Point kept in place: " & m_pos & " (" & ac2fRZPosText(m_pos) & ")"
    If m_deep Then s = s & vbCrLf & "Objects inside groups were sized one by one."
    ac2fRZDescribe = s
End Function

Private Function ac2fRZPosText(ByVal p As Long) As String
    ac2fRZPosText = Choose(p, "bottom left", "bottom centre", "bottom right", _
                           "left centre", "centre", "right centre", _
                           "top left", "top centre", "top right")
End Function
