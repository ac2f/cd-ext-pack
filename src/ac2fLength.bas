Attribute VB_Name = "ac2fLength"
'=====================================================================
'  ac2f pack  --  ac2fLength
'
'  Measures the total length of the outlines around vector objects.
'
'  Macros:
'    ac2fMeasureLength - measures the selection and reports
'    ac2fLabelLength   - measures and drops the result on the page
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Length Measurement"
Private Const MAX_DETAIL As Long = 15

'---------------------------------------------------------------------
' Measures the total outline length of the selected objects.
'---------------------------------------------------------------------
Public Sub ac2fMeasureLength()
Attribute ac2fMeasureLength.VB_Description = "ac2f pack: Measure total outline length of the selection"
    Dim res As ac2fResult

    res = ac2fMeasureSelection()
    If Not res.Ok Then
        ac2fWarn res.Message, CAPTION_
        Exit Sub
    End If

    If res.SubCount = 0 Then
        ac2fWarn "No measurable path found in the selection." & vbCrLf & vbCrLf & _
                 "Bitmaps, empty text and zero length objects cannot be measured.", CAPTION_
        Exit Sub
    End If

    ac2fInfo ac2fLengthReport(res), CAPTION_
End Sub

'---------------------------------------------------------------------
' Places the measurement on the page as artistic text.
'---------------------------------------------------------------------
Public Sub ac2fLabelLength()
Attribute ac2fLabelLength.VB_Description = "ac2f pack: Add the measurement to the page as text"
    Dim res As ac2fResult
    Dim sr As ShapeRange
    Dim txt As String

    If ActiveDocument Is Nothing Then
        ac2fWarn "Open a document first.", CAPTION_
        Exit Sub
    End If

    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "No objects selected to measure.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "No objects selected to measure.", CAPTION_
        Exit Sub
    End If

    res = ac2fMeasureRange(sr)
    If Not res.Ok Then
        ac2fWarn res.Message, CAPTION_
        Exit Sub
    End If
    If res.SubCount = 0 Then
        ac2fWarn "No measurable path found in the selection.", CAPTION_
        Exit Sub
    End If

    txt = "Total length: " & ac2fFmt(res.TotalMM) & " mm" & _
          "  (" & ac2fFmt(res.TotalMM / 1000#, 2) & " m)"

    If ac2fPlaceLabel(sr, txt) Then
        ac2fInfo "Label added to the page." & vbCrLf & vbCrLf & txt, CAPTION_
    Else
        ac2fWarn "The label could not be added, the result is shown here only." & _
                 vbCrLf & vbCrLf & txt, CAPTION_
    End If
End Sub

'=====================================================================
' Internals
'=====================================================================

' Builds a readable report from a measurement result.
Public Function ac2fLengthReport(ByRef res As ac2fResult) As String
    Dim s As String
    Dim i As Long
    Dim longest As Double, shortest As Double

    shortest = -1
    For i = 0 To res.SubCount - 1
        If res.SubLenMM(i) > longest Then longest = res.SubLenMM(i)
        If shortest < 0 Or res.SubLenMM(i) < shortest Then shortest = res.SubLenMM(i)
    Next i
    If shortest < 0 Then shortest = 0

    s = "TOTAL LENGTH" & vbCrLf
    s = s & "   " & ac2fFmtLength(res.TotalMM) & vbCrLf & vbCrLf

    s = s & "SELECTION" & vbCrLf
    s = s & "   Objects             : " & res.ShapeCount & vbCrLf
    s = s & "   Outlines (sub-paths): " & res.SubCount & vbCrLf
    s = s & "      closed           : " & (res.SubCount - res.OpenCount) & vbCrLf
    s = s & "      open             : " & res.OpenCount & vbCrLf & vbCrLf

    s = s & "OUTLINE STATISTICS" & vbCrLf
    s = s & "   Longest             : " & ac2fFmt(longest) & " mm" & vbCrLf
    s = s & "   Shortest            : " & ac2fFmt(shortest) & " mm" & vbCrLf
    s = s & "   Average             : " & ac2fFmt(res.TotalMM / res.SubCount) & " mm" & vbCrLf

    If res.ItemCount > 0 Then
        s = s & vbCrLf & "PER OBJECT"
        If res.ItemCount > MAX_DETAIL Then
            s = s & " (first " & MAX_DETAIL & " of " & res.ItemCount & ")"
        End If
        s = s & vbCrLf

        For i = 0 To res.ItemCount - 1
            If i >= MAX_DETAIL Then Exit For
            s = s & "   " & (i + 1) & ". " & res.Items(i).Name & " - " & _
                    ac2fFmt(res.Items(i).LengthMM) & " mm" & _
                    " (" & res.Items(i).SubPaths & " outlines)" & vbCrLf
        Next i
    End If

    ac2fLengthReport = s
End Function

' Adds artistic text below the selection. Returns True on success.
Private Function ac2fPlaceLabel(ByVal sr As ShapeRange, ByVal txt As String) As Boolean
    Dim oldUnit As cdrUnit
    Dim unitChanged As Boolean
    Dim x As Double, y As Double, w As Double, h As Double
    Dim t As Shape

    On Error GoTo Fail

    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    sr.GetBoundingBox x, y, w, h

    ActiveDocument.BeginCommandGroup ac2fTitle("label")
    Set t = ActiveLayer.CreateArtisticText(x, y - 8#, txt)
    On Error Resume Next
    t.Text.Story.Size = 10
    On Error GoTo Fail
    ActiveDocument.EndCommandGroup

    ac2fPlaceLabel = True

Cleanup:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Exit Function

Fail:
    ac2fPlaceLabel = False
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    Resume Cleanup
End Function
