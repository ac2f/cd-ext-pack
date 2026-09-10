Attribute VB_Name = "ac2fLedModule"
'=====================================================================
'  ac2f pack  --  ac2fLedModule
'
'  Quantity calculation for 3-LED module layouts.
'
'  Divides the measured outline length by the module spacing to get the
'  module count, LED count, total power and power supply requirement.
'
'  Macros:
'    ac2fLedModuleCount - calculates using the stored settings
'    ac2fLedQuickCount  - asks for the spacing and calculates once
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "3-LED Module Count"

' Calculation methods
Public Const AC2F_METHOD_PERIMETER  As Long = 1   ' full length of each outline
Public Const AC2F_METHOD_CENTRELINE As Long = 2   ' half the outline length

' Calculation result
Public Type ac2fLedResult
    Ok           As Boolean
    Message      As String
    LengthMM     As Double   ' effective length used by the calculation
    RawLengthMM  As Double   ' measured outline length
    PathCount    As Long
    ModuleCount  As Long
    LedCount     As Long
    TotalWatt    As Double
    NeededWatt   As Double
    PsuCount     As Long
    SpacingMM    As Double
    Method       As Long
    Factor       As Double
End Type

'---------------------------------------------------------------------
' Calculates using the stored settings.
'---------------------------------------------------------------------
Public Sub ac2fLedModuleCount()
Attribute ac2fLedModuleCount.VB_Description = "ac2f pack: Calculate the 3-LED module count"
    ac2fLedRun ac2fGetNum(AC2F_K_SPACING, AC2F_DEF_SPACING)
End Sub

'---------------------------------------------------------------------
' Asks for the spacing and calculates once, leaving settings untouched.
'---------------------------------------------------------------------
Public Sub ac2fLedQuickCount()
Attribute ac2fLedQuickCount.VB_Description = "ac2f pack: Calculate the LED module count, asking for the spacing"
    Dim stored As Double
    Dim answer As String
    Dim spacing As Double

    stored = ac2fGetNum(AC2F_K_SPACING, AC2F_DEF_SPACING)

    answer = InputBox("Module spacing (mm):", ac2fTitle(CAPTION_), ac2fNumStr(stored))
    If StrPtr(answer) = 0 Then Exit Sub      ' Cancel

    spacing = ac2fParseNum(answer, stored)
    If spacing <= 0 Then
        ac2fWarn "The module spacing must be greater than zero.", CAPTION_
        Exit Sub
    End If

    ac2fLedRun spacing
End Sub

'=====================================================================
' Internals
'=====================================================================

Private Sub ac2fLedRun(ByVal spacingMM As Double)
    Dim res As ac2fResult
    Dim led As ac2fLedResult

    If spacingMM <= 0 Then
        ac2fWarn "The module spacing must be greater than zero." & vbCrLf & _
                 "You can correct it under Settings.", CAPTION_
        Exit Sub
    End If

    res = ac2fMeasureSelection()
    If Not res.Ok Then
        ac2fWarn res.Message, CAPTION_
        Exit Sub
    End If
    If res.SubCount = 0 Then
        ac2fWarn "No measurable path found in the selection.", CAPTION_
        Exit Sub
    End If

    led = ac2fLedCalc(res, spacingMM)
    If Not led.Ok Then
        ac2fWarn led.Message, CAPTION_
        Exit Sub
    End If

    ac2fInfo ac2fLedReport(led), CAPTION_
End Sub

'---------------------------------------------------------------------
' Works out modules, LEDs and power from a measurement result.
'---------------------------------------------------------------------
Public Function ac2fLedCalc(ByRef res As ac2fResult, _
                            ByVal spacingMM As Double) As ac2fLedResult
    Dim out As ac2fLedResult
    Dim i As Long
    Dim segLen As Double
    Dim count As Double
    Dim total As Double
    Dim ledsPer As Long
    Dim moduleW As Double
    Dim psuW As Double
    Dim margin As Double
    Dim minPer As Long
    Dim method As Long
    Dim factor As Double

    If spacingMM <= 0 Then
        out.Message = "The module spacing must be greater than zero."
        ac2fLedCalc = out
        Exit Function
    End If

    ledsPer = ac2fGetLng(AC2F_K_LEDS, AC2F_DEF_LEDS)
    moduleW = ac2fGetNum(AC2F_K_MODULE_W, AC2F_DEF_MODULE_W)
    psuW = ac2fGetNum(AC2F_K_PSU_W, AC2F_DEF_PSU_W)
    margin = ac2fGetNum(AC2F_K_SAFETY, AC2F_DEF_SAFETY)
    minPer = ac2fGetLng(AC2F_K_MINPERPATH, AC2F_DEF_MINPERPATH)
    method = ac2fGetLng(AC2F_K_METHOD, AC2F_DEF_METHOD)
    factor = ac2fGetNum(AC2F_K_FACTOR, AC2F_DEF_FACTOR)

    If ledsPer < 1 Then ledsPer = AC2F_DEF_LEDS
    If minPer < 0 Then minPer = 0
    If factor <= 0 Then factor = 1#
    If method <> AC2F_METHOD_CENTRELINE Then method = AC2F_METHOD_PERIMETER

    For i = 0 To res.SubCount - 1
        segLen = res.SubLenMM(i)
        If method = AC2F_METHOD_CENTRELINE Then segLen = segLen / 2#

        If res.SubClosed(i) Then
            ' On a closed loop the modules are spread evenly around it.
            count = ac2fCeil(segLen / spacingMM)
        Else
            ' On an open path both ends carry a module.
            count = Int(segLen / spacingMM) + 1#
        End If

        If count < minPer Then count = minPer

        total = total + count
        out.LengthMM = out.LengthMM + segLen
    Next i

    total = ac2fCeil(total * factor)

    out.Ok = True
    out.RawLengthMM = res.TotalMM
    out.PathCount = res.SubCount
    out.SpacingMM = spacingMM
    out.Method = method
    out.Factor = factor
    out.ModuleCount = CLng(total)
    out.LedCount = out.ModuleCount * ledsPer
    out.TotalWatt = out.ModuleCount * moduleW
    out.NeededWatt = out.TotalWatt * (1# + margin / 100#)

    If psuW > 0 Then
        out.PsuCount = CLng(ac2fCeil(out.NeededWatt / psuW))
    Else
        out.PsuCount = 0
    End If

    ac2fLedCalc = out
End Function

'---------------------------------------------------------------------
' Turns the calculation into a readable report.
'---------------------------------------------------------------------
Public Function ac2fLedReport(ByRef led As ac2fLedResult) As String
    Dim s As String
    Dim ledsPer As Long
    Dim moduleW As Double
    Dim psuW As Double
    Dim margin As Double

    ledsPer = ac2fGetLng(AC2F_K_LEDS, AC2F_DEF_LEDS)
    moduleW = ac2fGetNum(AC2F_K_MODULE_W, AC2F_DEF_MODULE_W)
    psuW = ac2fGetNum(AC2F_K_PSU_W, AC2F_DEF_PSU_W)
    margin = ac2fGetNum(AC2F_K_SAFETY, AC2F_DEF_SAFETY)

    s = "RESULT" & vbCrLf
    s = s & "   Modules             : " & led.ModuleCount & vbCrLf
    s = s & "   LEDs                : " & led.LedCount & _
            "  (" & ledsPer & " LEDs per module)" & vbCrLf & vbCrLf

    s = s & "POWER" & vbCrLf
    s = s & "   Total power         : " & ac2fFmt(led.TotalWatt) & " W" & vbCrLf
    s = s & "   With " & ac2fFmt(margin, 0) & "% margin     : " & _
            ac2fFmt(led.NeededWatt) & " W" & vbCrLf
    If led.PsuCount > 0 Then
        s = s & "   Power supplies      : " & led.PsuCount & " x " & _
                ac2fFmt(psuW, 0) & " W" & vbCrLf
    End If
    s = s & vbCrLf

    s = s & "MEASUREMENT" & vbCrLf
    s = s & "   Raw outline length  : " & ac2fFmt(led.RawLengthMM) & " mm" & vbCrLf
    s = s & "   Length used         : " & ac2fFmt(led.LengthMM) & " mm" & vbCrLf
    s = s & "   Outlines            : " & led.PathCount & vbCrLf & vbCrLf

    s = s & "SETTINGS USED" & vbCrLf
    s = s & "   Module spacing      : " & ac2fFmt(led.SpacingMM) & " mm" & vbCrLf
    s = s & "   Method              : " & ac2fMethodName(led.Method) & vbCrLf
    s = s & "   Correction factor   : " & ac2fFmt(led.Factor) & vbCrLf
    s = s & "   Module power        : " & ac2fFmt(moduleW) & " W" & vbCrLf

    ac2fLedReport = s
End Function

Public Function ac2fMethodName(ByVal method As Long) As String
    If method = AC2F_METHOD_CENTRELINE Then
        ac2fMethodName = "Centreline estimate (outline / 2)"
    Else
        ac2fMethodName = "Perimeter based (full outline)"
    End If
End Function
