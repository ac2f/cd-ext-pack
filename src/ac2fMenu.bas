Attribute VB_Name = "ac2fMenu"
'=====================================================================
'  ac2f pack  --  ac2fMenu
'
'  Entry point, settings and about box.
'
'  Macros:
'    ac2fPack             - main menu (J or 21 opens the customer job menu)
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
        "  2  -  Length + label on page" & vbCrLf & _
        "  3  -  3-LED module count" & vbCrLf & _
        "  4  -  3-LED module count (ask spacing)" & vbCrLf & _
        "  5  -  Box letter strip" & vbCrLf & _
        "  6  -  Box letter strip (pick a profile)" & vbCrLf & _
        "  7  -  Box letter report (no drawing)" & vbCrLf & _
        "  8  -  ACP panel V-grooves" & vbCrLf & _
        "  9  -  Centerline (reduce to a single line)" & vbCrLf & _
        " 10  -  Nesting: settings, preview, apply" & vbCrLf & _
        " 11  -  Nesting report (nothing moves)" & vbCrLf
    m = m & _
        " 12  -  Bending order: apply it to the objects" & vbCrLf & _
        " 13  -  Bending order: number them on the page" & vbCrLf & _
        " 14  -  Bending order report (nothing changes)" & vbCrLf & _
        " 15  -  Send to plotter" & vbCrLf & _
        " 16  -  Plotter test shape (40 x 80 L)" & vbCrLf & _
        " 17  -  Write HPGL to a file" & vbCrLf & _
        " 18  -  Normalise an existing .plt and send" & vbCrLf & _
        " 19  -  Settings and profiles" & vbCrLf & _
        " 20  -  About" & vbCrLf & _
        " 21  -  Customer job: export, folders, price  (J)" & vbCrLf & _
        " 22  -  Nesting again, last settings" & vbCrLf & _
        " 23  -  Start points   24 - again   25 - markers" & vbCrLf & _
        " 26  -  Mark selection as offcut sheets" & vbCrLf & _
        " 27  -  Plotter orientation wizard" & vbCrLf

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
        Case "8": ac2fPanelGroove
        Case "9": ac2fCenterline.ac2fCenterline
        Case "10": ac2fNest.ac2fNest
        Case "11": ac2fNestReport
        Case "12": ac2fBendOrder
        Case "13": ac2fBendPreview
        Case "14": ac2fBendReport
        Case "15": ac2fPlotSend
        Case "16": ac2fPlotTest
        Case "17": ac2fPlotSave
        Case "18": ac2fPlotFixSend
        Case "19": ac2fSettings.ac2fSettings
        Case "20": ac2fAbout
        Case "21", "J", "j": ac2fJob.ac2fJob
        Case "22": ac2fNestAgain
        Case "23": ac2fStartPoints
        Case "24": ac2fStartPointsAgain
        Case "25": ac2fStartPointsShow
        Case "26": ac2fNestMarkSheets
        Case "27": ac2fPlotCalibrate
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

    ' A message box shows about 1024 characters and cuts the rest off, so
    ' every feature gets one line.
    s = AC2F_NAME & "  " & AC2F_VERSION & "  -  add-on package for CorelDRAW" & vbCrLf & vbCrLf
    s = s & "CONTENTS" & vbCrLf
    s = s & "  Length measurement   total length of the outlines" & vbCrLf
    s = s & "  3-LED module count   modules, LEDs, power, supplies" & vbCrLf
    s = s & "  Box letter strip     developed return, grooves, strip" & vbCrLf
    s = s & "  ACP panel V-grooves  flat sheet and groove lines" & vbCrLf
    s = s & "  Centerline           filled artwork to a single line" & vbCrLf
    s = s & "  Nesting              true shape, offcuts, preview window" & vbCrLf
    s = s & "  Start points         corner, side or centre; direction" & vbCrLf
    s = s & "  Send to plotter      HPGL straight from the selection" & vbCrLf
    s = s & "  Bending order        letters, setups, counters last" & vbCrLf
    s = s & "  Customer job         StockTrack, DXF name, folders" & vbCrLf
    s = s & "  Check, cut time, m2  look over and price the job" & vbCrLf & vbCrLf
    prof = ac2fActiveProfile()
    If Len(prof) = 0 Then prof = "<none>"
    s = s & "SETTINGS: " & ac2fSettingCount() & " in " & ac2fGroupCount() & _
            " groups, profile " & prof & vbCrLf
    s = s & "Change them from the main menu, option 19;" & vbCrLf
    s = s & "there ?N explains what setting N does."

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
