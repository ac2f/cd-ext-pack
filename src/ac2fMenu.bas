Attribute VB_Name = "ac2fMenu"
'=====================================================================
'  ac2f pack  --  ac2fMenu
'
'  Entry point, settings and about box.
'
'  Macros:
'    ac2fPack             - main menu
'    ac2fResetAllSettings - restore every setting to its default
'    ac2fAbout            - version and contents
'
'  All settings live in one sheet, ac2fSettings. There is deliberately
'  no second settings menu.
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
        "  5  -  Box letter strip" & vbCrLf & _
        "  6  -  Box letter strip (pick profile, tweak, draw)" & vbCrLf & _
        "  7  -  Box letter report (no drawing)" & vbCrLf & _
        "  8  -  Settings and profiles" & vbCrLf & _
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
        Case "6": ac2fBoxLetterStripProfile
        Case "7": ac2fBoxLetterReport
        Case "8": ac2fSettings
        Case "9": ac2fAbout
        Case Else
            ac2fWarn "Invalid choice: " & choice, CAPTION_
    End Select
End Sub

'---------------------------------------------------------------------
' Version and contents.
'---------------------------------------------------------------------
Public Sub ac2fAbout()
Attribute ac2fAbout.VB_Description = "ac2f pack: Version and contents"
    Dim s As String
    Dim prof As String

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
    prof = ac2fActiveProfile()
    If Len(prof) = 0 Then prof = "<none>"

    s = s & "CURRENT SETTINGS  (profile: " & prof & ")" & vbCrLf
    s = s & ac2fSettingsBrief() & vbCrLf
    s = s & "Change any of them from the main menu, option 8." & vbCrLf
    s = s & "There, ?N explains what setting N does."

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
    ac2fSetActiveProfile ""
    ac2fInfo "Settings reset." & vbCrLf & vbCrLf & ac2fSettingsBrief(), "Reset Settings"
End Sub
