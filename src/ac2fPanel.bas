Attribute VB_Name = "ac2fPanel"
'=====================================================================
'  ac2f pack  --  ac2fPanel
'
'  V-groove layout for aluminium composite panel (ACP) boxes.
'
'  Takes a rectangle, works out the flat sheet and draws the four
'  V-groove lines a CAM program can follow, in clockwise order with
'  the start point of each line placed so the tool enters where the
'  operator expects.
'
'  Macros:
'    ac2fPanelGroove - draw the sheet and its four V-grooves
'
'  LAYOUT
'  With a 100x200 rectangle and a 5 (cm) fold, "out" gives a 110x210
'  sheet whose groove lines intersect exactly on the original 100x200.
'  Lines run clockwise and span the full sheet:
'
'      1  left    x = X0+d      bottom -> top     length H
'      2  top     y = Y0+H-d    left   -> right   length W
'      3  right   x = X0+W-d    top    -> bottom  length H
'      4  bottom  y = Y0+d      right  -> left    length W
'
'  Each line is created in that order and named, so the object order
'  in the Object Manager is the cutting order.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "ACP Panel Grooves"

Private Const MAX_PANELS As Long = 100

'---------------------------------------------------------------------
' Registry keys (internal identifiers, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_ACP_FOLD As String = "ACPDerzMM"
Public Const AC2F_K_ACP_DIR  As String = "ACPYon"
Public Const AC2F_K_ACP_KEEP As String = "ACPKaynagiKoru"

Public Const AC2F_DEF_ACP_FOLD As Double = 50#   ' fold size (mm) = 5 cm
Public Const AC2F_DEF_ACP_DIR  As Long = 1       ' 1 = out, 2 = in
Public Const AC2F_DEF_ACP_KEEP As Long = 1       ' 1 = keep source rectangle

Public Const AC2F_ACP_OUT As Long = 1
Public Const AC2F_ACP_IN  As Long = 2

'=====================================================================
' MACRO
'=====================================================================

Public Sub ac2fPanelGroove()
Attribute ac2fPanelGroove.VB_Description = "ac2f pack: Draw ACP panel V-groove lines around a rectangle"
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean
    Dim n As Long, i As Long
    Dim src() As Shape
    Dim bx() As Double, by() As Double, bw() As Double, bh() As Double
    Dim fold As Double
    Dim dir_ As Long
    Dim keepSrc As Boolean
    Dim made As Long, skipped As Long, ungrouped As Long
    Dim rep As String

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Select the rectangle to groove.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Select the rectangle to groove.", CAPTION_
        Exit Sub
    End If

    If Not ac2fAskFold(fold, dir_) Then Exit Sub
    keepSrc = (ac2fGetLng(AC2F_K_ACP_KEEP, AC2F_DEF_ACP_KEEP) <> 0)

    On Error GoTo Fail
    oldOpt = Application.Optimization
    Application.Optimization = True
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    ' Read every bounding box up front. Creating shapes and grouping
    ' them changes the selection, so the source range must not be
    ' walked once drawing has started.
    n = sr.Count
    If n > MAX_PANELS Then n = MAX_PANELS
    ReDim src(0 To n - 1)
    ReDim bx(0 To n - 1): ReDim by(0 To n - 1)
    ReDim bw(0 To n - 1): ReDim bh(0 To n - 1)
    For i = 0 To n - 1
        Set src(i) = sr(i + 1)
        src(i).GetBoundingBox bx(i), by(i), bw(i), bh(i)
    Next i

    ActiveDocument.BeginCommandGroup ac2fTitle("ACP grooves")

    For i = 0 To n - 1
        Select Case ac2fPanelOne(bx(i), by(i), bw(i), bh(i), fold, dir_, rep)
            Case 1: made = made + 1
            Case 2: made = made + 1: ungrouped = ungrouped + 1
            Case Else: skipped = skipped + 1
        End Select
    Next i

    If Not keepSrc Then
        For i = 0 To n - 1
            On Error Resume Next
            src(i).Delete
            On Error GoTo Fail
        Next i
    End If

    ActiveDocument.EndCommandGroup

    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    On Error GoTo 0

    ac2fInfo ac2fPanelReport(made, skipped, ungrouped, fold, dir_, keepSrc, rep), CAPTION_
    Exit Sub

Fail:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    ac2fWarn "The operation failed: " & Err.Description, CAPTION_
End Sub

'=====================================================================
' INPUT
'=====================================================================

' Asks for the fold size and direction in one line, for example
' "5cm out" or "50mm in" or just "5cm". A bare number is millimetres.
Private Function ac2fAskFold(ByRef fold As Double, ByRef dir_ As Long) As Boolean
    Dim answer As String
    Dim sizeTok As String, dirTok As String
    Dim p As Long
    Dim d As Double
    Dim curFold As Double
    Dim curDir As Long

    curFold = ac2fGetNum(AC2F_K_ACP_FOLD, AC2F_DEF_ACP_FOLD)
    curDir = ac2fGetLng(AC2F_K_ACP_DIR, AC2F_DEF_ACP_DIR)

    answer = InputBox( _
        "Fold size and direction." & vbCrLf & vbCrLf & _
        "   5cm out    5 cm fold, sheet grows outwards" & vbCrLf & _
        "   50mm out   the same, written in millimetres" & vbCrLf & _
        "   5cm in     grooves set inwards, sheet size kept" & vbCrLf & vbCrLf & _
        "A bare number is read as millimetres." & vbCrLf & _
        "out = the selected rectangle is the FINISHED size" & vbCrLf & _
        "in  = the selected rectangle is the SHEET size", _
        ac2fTitle(CAPTION_), _
        ac2fNumStr(curFold) & "mm " & IIf(curDir = AC2F_ACP_IN, "in", "out"))

    If StrPtr(answer) = 0 Then Exit Function
    answer = Trim$(answer)
    If Len(answer) = 0 Then Exit Function

    p = InStr(answer, " ")
    If p > 0 Then
        sizeTok = Left$(answer, p - 1)
        dirTok = LCase$(Trim$(Mid$(answer, p + 1)))
    Else
        sizeTok = answer
        dirTok = ""
    End If

    d = ac2fParseLen(sizeTok, -1)
    If d <= 0 Then
        ac2fWarn "Could not read a fold size from """ & sizeTok & """." & vbCrLf & _
                 "Write it as 5cm, 50mm or 50.", CAPTION_
        Exit Function
    End If

    Select Case dirTok
        Case "": dir_ = curDir
        Case "in", "inner", "i", "ic": dir_ = AC2F_ACP_IN
        Case "out", "outer", "o", "dis": dir_ = AC2F_ACP_OUT
        Case Else
            ac2fWarn "Do not understand the direction """ & dirTok & """." & vbCrLf & _
                     "Use out or in.", CAPTION_
            Exit Function
    End Select

    fold = d
    ac2fSetNum AC2F_K_ACP_FOLD, d
    ac2fSetLng AC2F_K_ACP_DIR, dir_
    ac2fAskFold = True
End Function

' Reads a length written as "5cm", "50mm" or "50" and returns mm.
Public Function ac2fParseLen(ByVal s As String, ByVal defMM As Double) As Double
    Dim t As String
    Dim mult As Double

    t = LCase$(Trim$(s))
    mult = 1#
    If Len(t) > 2 Then
        If Right$(t, 2) = "cm" Then
            mult = 10#
            t = Trim$(Left$(t, Len(t) - 2))
        ElseIf Right$(t, 2) = "mm" Then
            t = Trim$(Left$(t, Len(t) - 2))
        End If
    End If

    If Len(t) = 0 Then
        ac2fParseLen = defMM
        Exit Function
    End If
    ac2fParseLen = ac2fParseNum(t, defMM) * mult
End Function

'=====================================================================
' ONE PANEL
'=====================================================================

' Returns 1 when drawn and grouped, 2 when drawn but grouping failed,
' 0 when the rectangle was skipped.
Private Function ac2fPanelOne(ByVal bx As Double, ByVal by As Double, _
                              ByVal bw As Double, ByVal bh As Double, _
                              ByVal fold As Double, ByVal dir_ As Long, _
                              ByRef rep As String) As Long
    Dim x0 As Double, y0 As Double, w As Double, h As Double
    Dim g(0 To 3) As Shape
    Dim parts(0 To 1) As Shape
    Dim sheet As Shape
    Dim grp As Shape, outer As Shape
    Dim i As Long

    If bw <= 0 Or bh <= 0 Then Exit Function

    If dir_ = AC2F_ACP_IN Then
        x0 = bx: y0 = by: w = bw: h = bh
        If 2# * fold >= w Or 2# * fold >= h Then
            rep = rep & "   skipped " & ac2fFmt(bw, 0) & "x" & ac2fFmt(bh, 0) & _
                  " - fold too large for an inward layout" & vbCrLf
            Exit Function
        End If
    Else
        x0 = bx - fold: y0 = by - fold
        w = bw + 2# * fold: h = bh + 2# * fold
    End If

    ' --- sheet outline -------------------------------------------------
    Set sheet = ac2fMakeRect(x0, y0, w, h)
    If sheet Is Nothing Then Exit Function
    sheet.Name = "ac2f ACP sheet " & ac2fFmt(w, 0) & "x" & ac2fFmt(h, 0)

    ' --- four groove lines, clockwise, in cutting order ----------------
    ' 1 left: starts at the bottom so the tool enters low
    Set g(0) = ac2fMakeLine(x0 + fold, y0, x0 + fold, y0 + h, "1 left")
    ' 2 top: starts at the left
    Set g(1) = ac2fMakeLine(x0, y0 + h - fold, x0 + w, y0 + h - fold, "2 top")
    ' 3 right: starts at the top
    Set g(2) = ac2fMakeLine(x0 + w - fold, y0 + h, x0 + w - fold, y0, "3 right")
    ' 4 bottom: starts at the right
    Set g(3) = ac2fMakeLine(x0 + w, y0 + fold, x0, y0 + fold, "4 bottom")

    For i = 0 To 3
        If g(i) Is Nothing Then
            ac2fPanelOne = 2
            Exit Function
        End If
    Next i

    rep = rep & "   sheet " & ac2fFmt(w, 0) & "x" & ac2fFmt(h, 0) & " mm" & _
          "   folded " & ac2fFmt(w - 2# * fold, 0) & "x" & ac2fFmt(h - 2# * fold, 0) & _
          " mm" & vbCrLf

    ' --- grouping ------------------------------------------------------
    Set grp = ac2fGroupShapes(g, 4, "ac2f ACP grooves")
    If grp Is Nothing Then
        ac2fPanelOne = 2
        Exit Function
    End If

    Set parts(0) = sheet
    Set parts(1) = grp
    Set outer = ac2fGroupShapes(parts, 2, _
                "ac2f ACP panel " & ac2fFmt(w, 0) & "x" & ac2fFmt(h, 0))
    If outer Is Nothing Then
        ac2fPanelOne = 2
        Exit Function
    End If

    ac2fPanelOne = 1
End Function

'---------------------------------------------------------------------
' Drawing helpers. Geometry must succeed; styling is best effort.
'---------------------------------------------------------------------

Private Function ac2fMakeRect(ByVal x0 As Double, ByVal y0 As Double, _
                              ByVal w As Double, ByVal h As Double) As Shape
    Dim s As Shape
    On Error GoTo Fail
    Set s = ActiveLayer.CreateRectangle(x0, y0 + h, x0 + w, y0)
    On Error Resume Next
    s.Fill.ApplyNoFill
    s.Outline.Width = 0.2
    s.Outline.Color.RGBAssign 0, 0, 0            ' sheet / cut contour
    On Error GoTo 0
    Set ac2fMakeRect = s
    Exit Function
Fail:
End Function

Private Function ac2fMakeLine(ByVal x1 As Double, ByVal y1 As Double, _
                              ByVal x2 As Double, ByVal y2 As Double, _
                              ByVal nm As String) As Shape
    Dim s As Shape
    On Error GoTo Fail
    ' The first pair is the start node, which is the point the CAM
    ' program enters the cut at.
    Set s = ActiveLayer.CreateLineSegment(x1, y1, x2, y2)
    On Error Resume Next
    s.Name = "ac2f groove " & nm
    s.Outline.Width = 0.15
    s.Outline.Color.RGBAssign 0, 160, 220        ' V-groove
    On Error GoTo 0
    Set ac2fMakeLine = s
    Exit Function
Fail:
End Function

' Groups shapes in the order given. Returns Nothing when grouping is
' not possible, in which case the shapes are still on the page.
Private Function ac2fGroupShapes(ByRef arr() As Shape, ByVal n As Long, _
                                 ByVal nm As String) As Shape
    Dim i As Long
    Dim g As Shape

    If n <= 0 Then Exit Function

    On Error GoTo Fail
    ActiveDocument.ClearSelection
    For i = 0 To n - 1
        If arr(i) Is Nothing Then Exit Function
        arr(i).AddToSelection
    Next i

    Set g = ActiveSelectionRange.Group()
    On Error Resume Next
    g.Name = nm
    On Error GoTo 0
    Set ac2fGroupShapes = g
    Exit Function
Fail:
End Function

'=====================================================================
' REPORT
'=====================================================================

Private Function ac2fPanelReport(ByVal made As Long, ByVal skipped As Long, _
                                 ByVal ungrouped As Long, ByVal fold As Double, _
                                 ByVal dir_ As Long, ByVal keepSrc As Boolean, _
                                 ByVal rep As String) As String
    Dim s As String

    s = "RESULT" & vbCrLf
    s = s & "   Panels drawn        : " & made & vbCrLf
    If skipped > 0 Then s = s & "   Skipped             : " & skipped & vbCrLf
    s = s & "   Fold size           : " & ac2fFmt(fold) & " mm  (" & _
            ac2fFmt(fold / 10#) & " cm)" & vbCrLf
    s = s & "   Direction           : " & ac2fPanelDirName(dir_) & vbCrLf
    s = s & "   Source rectangle    : " & IIf(keepSrc, "kept", "deleted") & vbCrLf

    If Len(rep) > 0 Then
        s = s & vbCrLf & "PANELS" & vbCrLf & rep
    End If

    s = s & vbCrLf & "GROOVE ORDER  (clockwise, object order = cut order)" & vbCrLf
    s = s & "   1 left    starts at the bottom, full height" & vbCrLf
    s = s & "   2 top     starts at the left,   full width" & vbCrLf
    s = s & "   3 right   starts at the top,    full height" & vbCrLf
    s = s & "   4 bottom  starts at the right,  full width" & vbCrLf
    s = s & vbCrLf & "   black = sheet outline    blue = V-groove" & vbCrLf
    s = s & "   The four grooves intersect on the folded size." & vbCrLf

    If ungrouped > 0 Then
        s = s & vbCrLf & "WARNING: " & ungrouped & " panel(s) could not be grouped." & _
                vbCrLf & "The lines are on the page but left loose."
    End If

    ac2fPanelReport = s
End Function

Public Function ac2fPanelDirName(ByVal d As Long) As String
    If d = AC2F_ACP_IN Then
        ac2fPanelDirName = "in - selection is the sheet"
    Else
        ac2fPanelDirName = "out - selection is the folded size"
    End If
End Function
