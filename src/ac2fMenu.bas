Attribute VB_Name = "ac2fMenu"
'=====================================================================
'  ac2f pack  --  ac2fMenu
'
'  Entry point, settings and about box.
'
'  Macros:
'    ac2fPack             - main menu
'    ac2fLedSettings      - LED module settings
'    ac2fResetAllSettings - restore every setting to its default
'    ac2fAbout            - version and contents
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Main Menu"

'---------------------------------------------------------------------
' Main menu of the package.
'---------------------------------------------------------------------
Public Sub ac2fPack()
Attribute ac2fPack.VB_Description = "ac2f pack: Main menu"
    Dim m As String
    Dim choice As String

    m = "Enter the number of the action you want:" & vbCrLf & vbCrLf & _
        "  1  -  Length measurement" & vbCrLf & _
        "  2  -  Length measurement + label on page" & vbCrLf & _
        "  3  -  3-LED module count" & vbCrLf & _
        "  4  -  3-LED module count (ask for spacing)" & vbCrLf & _
        "  5  -  Box letter strip (calculate and draw)" & vbCrLf & _
        "  6  -  Box letter report (no drawing)" & vbCrLf & _
        "  7  -  Settings - LED module" & vbCrLf & _
        "  8  -  Settings - Box letter" & vbCrLf & _
        "  9  -  About" & vbCrLf

    choice = InputBox(m, ac2fTitle(CAPTION_), "1")
    If StrPtr(choice) = 0 Then Exit Sub          ' Cancel
    choice = Trim$(choice)
    If Len(choice) = 0 Then Exit Sub

    Select Case choice
        Case "1": ac2fMeasureLength
        Case "2": ac2fLabelLength
        Case "3": ac2fLedModuleCount
        Case "4": ac2fLedQuickCount
        Case "5": ac2fBoxLetterStrip
        Case "6": ac2fBoxLetterReport
        Case "7": ac2fLedSettings
        Case "8": ac2fBoxLetterSettings
        Case "9": ac2fAbout
        Case Else
            ac2fWarn "Invalid choice: " & choice, CAPTION_
    End Select
End Sub

'---------------------------------------------------------------------
' Asks for the LED settings one by one and stores them.
'---------------------------------------------------------------------
Public Sub ac2fLedSettings()
Attribute ac2fLedSettings.VB_Description = "ac2f pack: Edit the LED module settings"
    Const C As String = "Settings - LED Module"
    Dim v As Double
    Dim n As Long

    If Not ac2fAskNum("Module spacing (mm)" & vbCrLf & _
                      "Distance between two LED modules.", C, _
                      AC2F_K_SPACING, AC2F_DEF_SPACING, 0.001, v) Then Exit Sub

    If Not ac2fAskLng("LEDs per module" & vbCrLf & _
                      "3 for 3-LED modules.", C, _
                      AC2F_K_LEDS, AC2F_DEF_LEDS, 1, n) Then Exit Sub

    If Not ac2fAskNum("Module power (W)" & vbCrLf & _
                      "Power drawn by a single module.", C, _
                      AC2F_K_MODULE_W, AC2F_DEF_MODULE_W, 0#, v) Then Exit Sub

    If Not ac2fAskNum("Power supply rating (W)" & vbCrLf & _
                      "Enter 0 to skip the power supply count.", C, _
                      AC2F_K_PSU_W, AC2F_DEF_PSU_W, 0#, v) Then Exit Sub

    If Not ac2fAskNum("Safety margin (%)" & vbCrLf & _
                      "Added when sizing the power supply.", C, _
                      AC2F_K_SAFETY, AC2F_DEF_SAFETY, 0#, v) Then Exit Sub

    If Not ac2fAskLng("Minimum modules per outline" & vbCrLf & _
                      "So that short pieces are not left empty.", C, _
                      AC2F_K_MINPERPATH, AC2F_DEF_MINPERPATH, 0, n) Then Exit Sub

    If Not ac2fAskLng("Calculation method" & vbCrLf & _
                      "1 = Perimeter based (full outline)" & vbCrLf & _
                      "2 = Centreline estimate (outline / 2, for hollow letters)", C, _
                      AC2F_K_METHOD, AC2F_DEF_METHOD, 1, n, 2) Then Exit Sub

    If Not ac2fAskNum("Correction factor" & vbCrLf & _
                      "Scales the result. 1 = no change.", C, _
                      AC2F_K_FACTOR, AC2F_DEF_FACTOR, 0.01, v) Then Exit Sub

    ac2fInfo "Settings saved." & vbCrLf & vbCrLf & ac2fLedSettingsSummary(), C
End Sub

'---------------------------------------------------------------------
' Version and contents.
'---------------------------------------------------------------------
Public Sub ac2fAbout()
Attribute ac2fAbout.VB_Description = "ac2f pack: Version and contents"
    Dim s As String

    s = AC2F_NAME & "  " & AC2F_VERSION & vbCrLf & vbCrLf
    s = s & "Add-on package for CorelDRAW." & vbCrLf & vbCrLf
    s = s & "CONTENTS" & vbCrLf
    s = s & "   - Length measurement" & vbCrLf
    s = s & "     Measures the total length of the outlines" & vbCrLf
    s = s & "     around the selected vectors." & vbCrLf & vbCrLf
    s = s & "   - 3-LED module count" & vbCrLf
    s = s & "     Works out the modules, LEDs, power and power" & vbCrLf
    s = s & "     supplies needed for the measured length." & vbCrLf & vbCrLf
    s = s & "   - Box letter strip" & vbCrLf
    s = s & "     Develops the return, lays out the grooves and" & vbCrLf
    s = s & "     draws a flat strip ready to cut." & vbCrLf & vbCrLf
    s = s & "LED SETTINGS" & vbCrLf & ac2fLedSettingsSummary() & vbCrLf
    s = s & "BOX LETTER SETTINGS" & vbCrLf & ac2fBLSettingsSummary()

    ac2fInfo s, "About"
End Sub

'---------------------------------------------------------------------
' Restores every setting to its factory value.
'---------------------------------------------------------------------
Public Sub ac2fResetAllSettings()
Attribute ac2fResetAllSettings.VB_Description = "ac2f pack: Restore every setting to its default"
    If MsgBox("Every ac2f pack setting will go back to its default value." & vbCrLf & _
              "Continue?", vbQuestion + vbYesNo, _
              ac2fTitle("Reset Settings")) <> vbYes Then Exit Sub

    ac2fResetSettings
    ac2fInfo "Settings reset." & vbCrLf & vbCrLf & ac2fLedSettingsSummary(), "Reset Settings"
End Sub

'=====================================================================
' Internals
'=====================================================================

Public Function ac2fLedSettingsSummary() As String
    Dim s As String

    s = "   Module spacing       : " & ac2fFmt(ac2fGetNum(AC2F_K_SPACING, AC2F_DEF_SPACING)) & " mm" & vbCrLf
    s = s & "   LEDs per module      : " & ac2fGetLng(AC2F_K_LEDS, AC2F_DEF_LEDS) & vbCrLf
    s = s & "   Module power         : " & ac2fFmt(ac2fGetNum(AC2F_K_MODULE_W, AC2F_DEF_MODULE_W)) & " W" & vbCrLf
    s = s & "   Power supply         : " & ac2fFmt(ac2fGetNum(AC2F_K_PSU_W, AC2F_DEF_PSU_W), 0) & " W" & vbCrLf
    s = s & "   Safety margin        : " & ac2fFmt(ac2fGetNum(AC2F_K_SAFETY, AC2F_DEF_SAFETY), 0) & "%" & vbCrLf
    s = s & "   Min per outline      : " & ac2fGetLng(AC2F_K_MINPERPATH, AC2F_DEF_MINPERPATH) & " modules" & vbCrLf
    s = s & "   Method               : " & ac2fMethodName(ac2fGetLng(AC2F_K_METHOD, AC2F_DEF_METHOD)) & vbCrLf
    s = s & "   Correction factor    : " & ac2fFmt(ac2fGetNum(AC2F_K_FACTOR, AC2F_DEF_FACTOR)) & vbCrLf

    ac2fLedSettingsSummary = s
End Function

' Asks for a decimal setting. Returns False when cancelled.
Private Function ac2fAskNum(ByVal question As String, ByVal caption As String, _
                            ByVal key As String, ByVal defValue As Double, _
                            ByVal minVal As Double, ByRef result As Double) As Boolean
    Dim current As Double
    Dim answer As String
    Dim v As Double

    current = ac2fGetNum(key, defValue)
    answer = InputBox(question & vbCrLf & vbCrLf & _
                      "(Default: " & ac2fNumStr(defValue) & ")", _
                      ac2fTitle(caption), ac2fNumStr(current))
    If StrPtr(answer) = 0 Then Exit Function

    v = ac2fParseNum(answer, current)
    If v < minVal Then v = minVal

    ac2fSetNum key, v
    result = v
    ac2fAskNum = True
End Function

' Asks for an integer setting. Returns False when cancelled.
Private Function ac2fAskLng(ByVal question As String, ByVal caption As String, _
                            ByVal key As String, ByVal defValue As Long, _
                            ByVal minVal As Long, ByRef result As Long, _
                            Optional ByVal maxVal As Long = 0) As Boolean
    Dim current As Long
    Dim answer As String
    Dim v As Long

    current = ac2fGetLng(key, defValue)
    answer = InputBox(question & vbCrLf & vbCrLf & _
                      "(Default: " & ac2fNumStr(CDbl(defValue)) & ")", _
                      ac2fTitle(caption), ac2fNumStr(CDbl(current)))
    If StrPtr(answer) = 0 Then Exit Function

    v = CLng(ac2fParseNum(answer, CDbl(current)))
    If v < minVal Then v = minVal
    If maxVal > 0 And v > maxVal Then v = maxVal

    ac2fSetLng key, v
    result = v
    ac2fAskLng = True
End Function
