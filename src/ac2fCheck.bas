Attribute VB_Name = "ac2fCheck"
'=====================================================================
'  ac2f pack  --  ac2fCheck
'
'  Looks the drawing over before it goes out as DXF, so a fault turns
'  up in CorelDRAW instead of at the machine.
'
'  Macro:
'    ac2fCheck - checks the selection (or the page) and reports
'
'  Export DXF (ac2fJobExport) runs the same check first when Check first
'  is on, and stops only for the faults that spoil a cut.
'
'  WHAT IS LOOKED FOR
'
'    Spoils the cut (Export asks before going on):
'      text left as text   DXF carries it as text; CAM cannot cut it
'      bitmap or OLE       nothing to cut
'      double line         two objects on top of each other: cut twice
'      off the sheet       a part crossing a nest sheet's edge, or a job
'                          bigger than the job's sheet
'      sheet outline       Nesting's sheet rectangle in the export: the
'                          machine would cut round the sheet
'
'    Worth a look (reported, never stops an export):
'      open outline        right for engraving, wrong for a cut-out
'      speck               smaller than 0.5 mm: a stray node or dot
'      PowerClip           contents may not export the way they look
'
'  Double lines are found by bounding box and length, both within
'  0.05 mm. Groups and PowerClip contents are looked inside.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Check"

Private Const SPECK_MM As Double = 0.5
Private Const DUP_MM   As Double = 0.05
Private Const DUP_LEN  As Double = 0.1
Private Const EDGE_MM  As Double = 0.1
Private Const MAX_DUP  As Long = 3000      ' pairs are compared, n squared

' leaf kinds
Private Const K_CURVE As Long = 0
Private Const K_TEXT  As Long = 1
Private Const K_IMAGE As Long = 2

' problem kinds, in report order; the first five spoil a cut
Private Const P_TEXT  As Long = 1
Private Const P_IMAGE As Long = 2
Private Const P_DUP   As Long = 3
Private Const P_EDGE  As Long = 4
Private Const P_SHEET As Long = 5
Private Const P_OPEN  As Long = 6
Private Const P_SPECK As Long = 7
Private Const P_CLIP  As Long = 8
Private Const P_LAST  As Long = 8
Private Const P_WORST As Long = 5          ' kinds up to here stop an export

' leaves
Private m_n As Long
Private m_shp() As Shape
Private m_kind() As Long
Private m_x0() As Double, m_y0() As Double, m_x1() As Double, m_y1() As Double
Private m_len() As Double
Private m_open() As Long
Private m_prob() As Long                   ' worst problem found on the leaf

' extra shapes that are problems but not leaves
Private m_extN As Long
Private m_ext() As Shape
Private m_extKind() As Long

Private m_count(1 To 8) As Long
Private m_bigNote As String                ' job bigger than its sheet

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fCheck()
Attribute ac2fCheck.VB_Description = "ac2f pack: Check the drawing before export - text, bitmaps, double lines, parts off the sheet"
    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If
    ac2fCKRun ActiveSelectionRange, False
End Sub

' Called by Export DXF. True when the export may go ahead: nothing that
' spoils the cut, or the user chose to export anyway.
Public Function ac2fCheckBeforeExport(ByVal sr As ShapeRange) As Boolean
    ac2fCheckBeforeExport = ac2fCKRun(sr, True)
End Function

'=====================================================================
' THE CHECK
'=====================================================================

Private Function ac2fCKRun(ByVal sr As ShapeRange, ByVal forExport As Boolean) As Boolean
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim wholePage As Boolean
    Dim i As Long, worst As Long, answer As VbMsgBoxResult
    Dim m As String

    ac2fCKReset
    On Error GoTo Fail
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    wholePage = sr Is Nothing
    If Not wholePage Then wholePage = (sr.Count = 0)
    If wholePage Then
        For i = 1 To ActivePage.Shapes.Count
            ac2fCKWalk ActivePage.Shapes(i)
        Next i
    Else
        For i = 1 To sr.Count
            ac2fCKWalk sr(i)
        Next i
    End If

    ac2fCKDuplicates
    ac2fCKSheets
    On Error GoTo 0

    For i = 1 To P_LAST
        If m_count(i) > 0 And worst = 0 Then worst = i
    Next i
    If unitChanged Then ActiveDocument.Unit = oldUnit
    unitChanged = False

    ' --- nothing wrong ---
    If worst = 0 Then
        If Not forExport Then
            ac2fInfo "All clear: " & m_n & " object(s), " & _
                     ac2fFmt(ac2fCKTotalLen() / 1000#, 2) & " m of outline.", CAPTION_
        End If
        ac2fCKRun = True
        Exit Function
    End If

    ' --- report ---
    m = ac2fCKReport(wholePage)
    If forExport Then
        If worst > P_WORST Then
            ac2fCKRun = True                ' only worth-a-look items
            Exit Function
        End If
        m = m & vbCrLf & "Yes       export anyway" & vbCrLf & _
                         "No        stop and select them" & vbCrLf & _
                         "Cancel    stop"
        answer = MsgBox(m, vbExclamation + vbYesNoCancel, ac2fTitle(CAPTION_))
        If answer = vbYes Then
            ac2fCKRun = True
        ElseIf answer = vbNo Then
            ac2fCKSelect
        End If
    Else
        If MsgBox(m & vbCrLf & "Select them?", vbQuestion + vbYesNo, _
                  ac2fTitle(CAPTION_)) = vbYes Then ac2fCKSelect
        ac2fCKRun = (worst > P_WORST)
    End If
    Exit Function

Fail:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    ac2fWarn "The check could not finish: " & Err.Description & vbCrLf & _
             "Nothing was changed.", CAPTION_
    ' a check that cannot run must not block the export
    ac2fCKRun = forExport
End Function

Private Sub ac2fCKReset()
    Dim i As Long
    m_n = 0
    m_extN = 0
    m_bigNote = ""
    ReDim m_shp(1 To 64): ReDim m_kind(1 To 64)
    ReDim m_x0(1 To 64): ReDim m_y0(1 To 64): ReDim m_x1(1 To 64): ReDim m_y1(1 To 64)
    ReDim m_len(1 To 64): ReDim m_open(1 To 64): ReDim m_prob(1 To 64)
    ReDim m_ext(1 To 16): ReDim m_extKind(1 To 16)
    For i = 1 To P_LAST
        m_count(i) = 0
    Next i
End Sub

' Groups are opened, PowerClip contents looked into, Nesting's sheet
' outlines noted, and everything else becomes a leaf.
Private Sub ac2fCKWalk(ByVal s As Shape)
    Dim i As Long
    Dim nm As String
    Dim pc As Object

    If s Is Nothing Then Exit Sub
    On Error Resume Next
    nm = LCase$(s.Name)
    On Error GoTo 0

    If nm = "ac2f sheets" Or nm Like "ac2f sheet #*" Then
        ac2fCKExtra s, P_SHEET
        Exit Sub
    End If
    If s.Type = cdrGroupShape Then
        For i = 1 To s.Shapes.Count
            ac2fCKWalk s.Shapes(i)
        Next i
        Exit Sub
    End If

    Set pc = Nothing
    On Error Resume Next
    Set pc = s.PowerClip
    On Error GoTo 0
    If Not pc Is Nothing Then
        If pc.Shapes.Count > 0 Then
            ac2fCKExtra s, P_CLIP
            For i = 1 To pc.Shapes.Count
                ac2fCKWalk pc.Shapes(i)
            Next i
        End If
    End If

    ac2fCKLeaf s
End Sub

Private Sub ac2fCKLeaf(ByVal s As Shape)
    Dim cv As Curve
    Dim x As Double, y As Double, w As Double, h As Double
    Dim i As Long, n As Long, nOpen As Long
    Dim total As Double, part As Double
    Dim closedFlag As Boolean

    m_n = m_n + 1
    If m_n > UBound(m_shp) Then
        ReDim Preserve m_shp(1 To m_n * 2): ReDim Preserve m_kind(1 To m_n * 2)
        ReDim Preserve m_x0(1 To m_n * 2): ReDim Preserve m_y0(1 To m_n * 2)
        ReDim Preserve m_x1(1 To m_n * 2): ReDim Preserve m_y1(1 To m_n * 2)
        ReDim Preserve m_len(1 To m_n * 2): ReDim Preserve m_open(1 To m_n * 2)
        ReDim Preserve m_prob(1 To m_n * 2)
    End If
    Set m_shp(m_n) = s
    m_prob(m_n) = 0
    m_len(m_n) = -1
    m_open(m_n) = 0

    On Error Resume Next
    s.GetBoundingBox x, y, w, h
    m_x0(m_n) = x: m_y0(m_n) = y: m_x1(m_n) = x + w: m_y1(m_n) = y + h

    Select Case s.Type
        Case cdrTextShape
            m_kind(m_n) = K_TEXT
            ac2fCKMark m_n, P_TEXT
            Exit Sub
        Case cdrBitmapShape, cdrOLEObjectShape
            m_kind(m_n) = K_IMAGE
            ac2fCKMark m_n, P_IMAGE
            Exit Sub
    End Select
    m_kind(m_n) = K_CURVE

    ' outline length and open sub-paths, without touching the document
    Set cv = Nothing
    Set cv = s.DisplayCurve
    If Not cv Is Nothing Then
        n = cv.SubPaths.Count
        For i = 1 To n
            part = 0
            part = cv.SubPaths(i).Length
            closedFlag = True
            closedFlag = cv.SubPaths(i).Closed
            total = total + part
            If Not closedFlag And part > 0 Then nOpen = nOpen + 1
        Next i
        m_len(m_n) = total
        m_open(m_n) = nOpen
    End If
    On Error GoTo 0

    If w < SPECK_MM And h < SPECK_MM Then
        ac2fCKMark m_n, P_SPECK
    ElseIf m_len(m_n) >= 0 And m_len(m_n) < SPECK_MM * 2 Then
        ac2fCKMark m_n, P_SPECK
    End If
    If m_open(m_n) > 0 Then ac2fCKMark m_n, P_OPEN
End Sub

Private Sub ac2fCKMark(ByVal i As Long, ByVal kind As Long)
    m_count(kind) = m_count(kind) + 1
    If m_prob(i) = 0 Or kind < m_prob(i) Then m_prob(i) = kind
End Sub

Private Sub ac2fCKExtra(ByVal s As Shape, ByVal kind As Long)
    m_extN = m_extN + 1
    If m_extN > UBound(m_ext) Then
        ReDim Preserve m_ext(1 To m_extN * 2)
        ReDim Preserve m_extKind(1 To m_extN * 2)
    End If
    Set m_ext(m_extN) = s
    m_extKind(m_extN) = kind
    m_count(kind) = m_count(kind) + 1
End Sub

' Two outlines with the same box and length lie on top of each other:
' the machine would cut the line twice. The later one is reported.
Private Sub ac2fCKDuplicates()
    Dim i As Long, j As Long

    If m_n > MAX_DUP Then Exit Sub
    For i = 1 To m_n - 1
        If m_kind(i) = K_CURVE And m_len(i) > 0 Then
            For j = i + 1 To m_n
                If m_kind(j) = K_CURVE And m_prob(j) <> P_DUP Then
                    If Abs(m_x0(i) - m_x0(j)) <= DUP_MM And Abs(m_y0(i) - m_y0(j)) <= DUP_MM _
                       And Abs(m_x1(i) - m_x1(j)) <= DUP_MM And Abs(m_y1(i) - m_y1(j)) <= DUP_MM _
                       And Abs(m_len(i) - m_len(j)) <= DUP_LEN Then
                        ac2fCKMark j, P_DUP
                    End If
                End If
            Next j
        End If
    Next i
End Sub

' With Nesting's sheet outlines on the page: a part must not cross an
' edge. Without them: the whole job must fit the job's sheet.
Private Sub ac2fCKSheets()
    Dim sx0() As Double, sy0() As Double, sx1() As Double, sy1() As Double
    Dim ns As Long, i As Long, j As Long
    Dim inside As Boolean, touches As Boolean
    Dim bx0 As Double, by0 As Double, bx1 As Double, by1 As Double
    Dim sw As Double, sh As Double, thick As Double, bw As Double, bh As Double
    Dim kind As String

    ReDim sx0(1 To 8): ReDim sy0(1 To 8): ReDim sx1(1 To 8): ReDim sy1(1 To 8)
    ac2fCKSheetList ActivePage.Shapes, sx0, sy0, sx1, sy1, ns

    If ns > 0 Then
        For i = 1 To m_n
            touches = False
            inside = False
            For j = 1 To ns
                If m_x1(i) > sx0(j) And m_x0(i) < sx1(j) And _
                   m_y1(i) > sy0(j) And m_y0(i) < sy1(j) Then
                    touches = True
                    If m_x0(i) >= sx0(j) - EDGE_MM And m_x1(i) <= sx1(j) + EDGE_MM And _
                       m_y0(i) >= sy0(j) - EDGE_MM And m_y1(i) <= sy1(j) + EDGE_MM Then
                        inside = True
                    End If
                End If
            Next j
            If touches And Not inside Then ac2fCKMark i, P_EDGE
        Next i
        Exit Sub
    End If

    If m_n = 0 Then Exit Sub
    If Not ac2fJobSheetInfo(sw, sh, thick, kind) Then Exit Sub
    If sw <= 0 Or sh <= 0 Then Exit Sub
    bx0 = m_x0(1): by0 = m_y0(1): bx1 = m_x1(1): by1 = m_y1(1)
    For i = 2 To m_n
        If m_x0(i) < bx0 Then bx0 = m_x0(i)
        If m_y0(i) < by0 Then by0 = m_y0(i)
        If m_x1(i) > bx1 Then bx1 = m_x1(i)
        If m_y1(i) > by1 Then by1 = m_y1(i)
    Next i
    bw = bx1 - bx0: bh = by1 - by0
    If (bw <= sw + EDGE_MM And bh <= sh + EDGE_MM) Or _
       (bh <= sw + EDGE_MM And bw <= sh + EDGE_MM) Then Exit Sub
    m_count(P_EDGE) = m_count(P_EDGE) + 1
    m_bigNote = Format$(bw, "0") & " x " & Format$(bh, "0") & " mm does not fit the " & _
                Format$(sw, "0") & " x " & Format$(sh, "0") & " sheet"
End Sub

Private Sub ac2fCKSheetList(ByVal coll As Object, _
                            ByRef sx0() As Double, ByRef sy0() As Double, _
                            ByRef sx1() As Double, ByRef sy1() As Double, ByRef ns As Long)
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
            ac2fCKSheetList s.Shapes, sx0, sy0, sx1, sy1, ns
        ElseIf nm Like "ac2f sheet #*" Then
            s.GetBoundingBox x, y, w, h
            If w > 0 And h > 0 Then
                ns = ns + 1
                If ns > UBound(sx0) Then
                    ReDim Preserve sx0(1 To ns * 2): ReDim Preserve sy0(1 To ns * 2)
                    ReDim Preserve sx1(1 To ns * 2): ReDim Preserve sy1(1 To ns * 2)
                End If
                sx0(ns) = x: sy0(ns) = y: sx1(ns) = x + w: sy1(ns) = y + h
            End If
        End If
    Next i
End Sub

'=====================================================================
' REPORT
'=====================================================================

Private Function ac2fCKReport(ByVal wholePage As Boolean) As String
    Dim m As String, bad As String, look As String

    If wholePage Then
        m = "CHECK  -  the whole page, " & m_n & " object(s)" & vbCrLf & vbCrLf
    Else
        m = "CHECK  -  " & m_n & " selected object(s)" & vbCrLf & vbCrLf
    End If

    bad = bad & ac2fCKLine(P_TEXT, "text object(s) still text  (convert: Ctrl+Q)")
    bad = bad & ac2fCKLine(P_IMAGE, "bitmap / OLE object(s): nothing to cut")
    bad = bad & ac2fCKLine(P_DUP, "double line(s): one on top of another, cut twice")
    If Len(m_bigNote) > 0 Then
        bad = bad & "   " & m_bigNote & vbCrLf
    Else
        bad = bad & ac2fCKLine(P_EDGE, "part(s) crossing a nest sheet edge")
    End If
    bad = bad & ac2fCKLine(P_SHEET, "Nesting sheet outline(s) in the export: would be cut")

    look = look & ac2fCKLine(P_OPEN, "object(s) with open outlines")
    look = look & ac2fCKLine(P_SPECK, "speck(s) under 0.5 mm")
    look = look & ac2fCKLine(P_CLIP, "PowerClip(s): contents may export differently")

    If Len(bad) > 0 Then m = m & "Spoils the cut:" & vbCrLf & bad & vbCrLf
    If Len(look) > 0 Then m = m & "Worth a look:" & vbCrLf & look & vbCrLf
    If m_n > MAX_DUP Then
        m = m & "(over " & MAX_DUP & " objects: double lines were not checked)" & vbCrLf
    End If
    ac2fCKReport = m
End Function

Private Function ac2fCKLine(ByVal kind As Long, ByVal what As String) As String
    If m_count(kind) > 0 Then ac2fCKLine = ac2fRPad(CStr(m_count(kind)), 5) & "  " & what & vbCrLf
End Function

' Selects every object behind a problem so it can be fixed in place.
' ClearSelection + AddToSelection, the way Nesting selects its sheets,
' rather than building a ShapeRange.
Private Sub ac2fCKSelect()
    Dim i As Long, n As Long

    On Error Resume Next
    ActiveDocument.ClearSelection
    For i = 1 To m_n
        If m_prob(i) > 0 Then
            m_shp(i).AddToSelection
            If Err.Number = 0 Then n = n + 1
            Err.Clear
        End If
    Next i
    For i = 1 To m_extN
        m_ext(i).AddToSelection
        If Err.Number = 0 Then n = n + 1
        Err.Clear
    Next i
    If n = 0 Then ac2fWarn "Could not select them.", CAPTION_
End Sub

Private Function ac2fCKTotalLen() As Double
    Dim i As Long, t As Double
    For i = 1 To m_n
        If m_len(i) > 0 Then t = t + m_len(i)
    Next i
    ac2fCKTotalLen = t
End Function
