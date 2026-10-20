Attribute VB_Name = "ac2fJob"
'=====================================================================
'  ac2f pack  --  ac2fJob
'
'  Customer jobs. Ties the open file to a customer from StockTrack,
'  takes the sheet from stock, and names every export and every
'  toolpath folder from those two, so nothing is typed by hand.
'
'  Macros:
'    ac2fJob         - job menu
'    ac2fJobLink     - link the document to a customer
'    ac2fJobMaterial - pick the sheet from stock
'    ac2fJobExport   - export DXF with an automatic name
'    ac2fJobFolders  - create, or stack onto, the toolpath folders
'    ac2fJobUsage    - after the quote: mark the part of a sheet it used
'    ac2fJobOpen     - open the customer folder in Explorer
'    ac2fJobSetup    - server, key, day folder, separator
'    ac2fJobHistory  - this customer's earlier jobs
'    ac2fJobDay      - which day folder the job goes into (change it)
'
'  The job menu also reaches Check before export (ac2fCheck) and the two
'  price modules, Cut time (ac2fCutTime) and Area price (ac2fCutArea),
'  which read their rates from StockTrack through this module.
'
'  WHERE THINGS GO
'
'    <day>\<customer>\exp01--pleksi.dxf
'    <day>\<customer>\001-pleksi-1220x2440--2pcs--2800um-akpolimer-dokme-beyaz\001
'                                                                           \002
'                                                                           \u1220x900
'
'  The day folder is the one CDR Auto Renamer makes every morning. Its
'  config.json is read for the base folder, the date format and the
'  separator, so the two agree without a second setting, and a missing
'  day folder is created the same way.
'
'  The customer folder is the company name when there is one, else the
'  person's name, lower case with dashes for spaces (ozturk-reklam-ltd-sti);
'  a folder made under the older spaced name is still found. A job that already has a folder on an earlier day (a
'  quote given on the 22nd, approved on the 25th) asks whether to carry
'  on there or to start again under today; the answer stays with the
'  job for the rest of the day. Every name is plain ASCII: Turkish letters are folded
'  (s for s-cedilla, i for dotless i), since CAM programs and the ANSI
'  file calls VBA uses both trip over them.
'
'  SIZES
'  A folder always carries the size of the sheet as stocked, whole. How
'  much of it a job really used is known only once the quote is made, so
'  Used parts reads the quote afterwards and drops a file named after
'  the part into the sheet's folder: u1350x1000 is 1350 x 1000 mm used
'  of a 1350 x 2000 sheet, the sheet's width first. A whole sheet gets
'  no file.
'
'  STACKING
'  Sheets that agree on kind, size, thickness, brand, grade and colour
'  share one folder whose count grows: --2pcs-- becomes --3pcs-- and a
'  003 folder is added. Any difference opens the next numbered folder.
'  Sheets longer than Long sheet (4000 mm) get p001, p002 inside every
'  sheet folder, because the bed takes them in two goes.
'
'  STOCKTRACK
'  Customers and stock come from StockTrack over HTTP (MSXML, nothing to
'  install). Replies are tab separated ASCII, so no JSON parser is
'  needed. The key is made in StockTrack under Settings > Corel
'  integration and pasted into Job setup once. Without a server the
'  module still works: names and sheets are typed instead.
'
'  Customers, sheets, quotes and old jobs are picked in a list window
'  (PowerShell, part of Windows) that narrows as you type; every hit is
'  listed, even a single one. Without PowerShell the numbered prompts
'  come back.
'
'  The link is kept in the document (a document property), so reopening
'  the file brings the customer back. A copy goes to the registry
'  against the file path, for files that are never saved.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Customer Job"

'---------------------------------------------------------------------
' Registry keys (internal, kept stable across releases)
'---------------------------------------------------------------------
Public Const AC2F_K_JB_SERVER As String = "JBSunucuAdresi"      ' string
Public Const AC2F_K_JB_KEY    As String = "JBAnahtar"           ' string
Public Const AC2F_K_JB_BASE   As String = "JBGunlukKlasorKoku"  ' string, empty = CDR Auto Renamer
Public Const AC2F_K_JB_SEP    As String = "JBAyrac"             ' string
Public Const AC2F_K_JB_RCUST  As String = "JBSonMusteriler"     ' string
Public Const AC2F_K_JB_RMAT   As String = "JBSonMalzemeler"     ' string
Public Const AC2F_K_JB_LONG   As String = "JBUzunPlakaMM"
Public Const AC2F_K_JB_DIALOG As String = "JBExportPenceresi"
Public Const AC2F_K_JB_OPEN   As String = "JBKlasoruAc"
Public Const AC2F_K_JB_NEST   As String = "JBNestPlakasi"
Public Const AC2F_K_JB_BACK   As String = "JBGeriBakmaGun"
Public Const AC2F_K_JB_CHECK  As String = "JBExportOncesiKontrol"
Public Const AC2F_K_JB_PRICE  As String = "JBKesimFiyatlari"   ' string, last pricing row
Public Const AC2F_K_JB_LIST   As String = "JBListePenceresi"

Public Const AC2F_DEF_JB_LONG   As Double = 4000#   ' longer sheets get p001 / p002
Public Const AC2F_DEF_JB_DIALOG As Long = 1         ' show the DXF export dialog
Public Const AC2F_DEF_JB_OPEN   As Long = 0         ' open Explorer afterwards
Public Const AC2F_DEF_JB_NEST   As Long = 1         ' picking a sheet sets the nest sheet
Public Const AC2F_DEF_JB_BACK   As Long = 30        ' days to look back for the job's folder
Public Const AC2F_DEF_JB_CHECK  As Long = 1         ' check the drawing before every export
Public Const AC2F_DEF_JB_LIST   As Long = 1         ' pick customers and sheets in a list window
Public Const AC2F_DEF_JB_SEP    As String = "--"

Private Const JOB_SEC     As String = "Jobs"        ' per document links
Private Const PROP_NAME   As String = "ac2fJob"
Private Const PROP_ID     As Long = 1
Private Const REC_VER     As String = "j1"
Private Const MAX_RECENT  As Long = 9
Private Const HISTORY_DAYS As Long = 365
Private Const HTTP_HEAD   As String = "#ac2f-corel "
Private Const PROMPT_MAX  As Long = 880             ' InputBox stops near 1024
Private Const CACHE_LINES As Long = 1000            ' customers / sheets kept for offline search

' What the document is linked to. The sheet fields hold StockTrack's
' ready made name pieces: lower case ASCII, thickness in microns.
Private Type ac2fJobRec
    CustId     As String
    CustFolder As String
    CustLabel  As String
    MatId      As String
    MatOwner   As String     ' C customer's own, B business stock, M typed
    MatKind    As String
    MatThick   As String
    MatBrand   As String
    MatGrade   As String     ' variant: dokme, cekme ...
    MatColor   As String
    MatW       As Double
    MatH       As Double
    MatLabel   As String
    JobDay     As String     ' day folder the job's customer folder is in
    JobAsked   As String     ' date (serial) the day was last chosen
End Type

' When the last StockTrack reply came from the offline copy: the copy's
' date, else empty. Lists say so in their heading.
Private mJBSaved As String

'=====================================================================
' MACROS
'=====================================================================

Public Sub ac2fJob()
Attribute ac2fJob.VB_Description = "ac2f pack: Customer job menu - link, sheet, export, folders"
    Dim jr As ac2fJobRec
    Dim m As String, choice As String, dflt As String

    If Not ActiveDocument Is Nothing Then ac2fJBLoad jr

    m = "CUSTOMER JOB" & vbCrLf & vbCrLf & ac2fJBSummary(jr) & vbCrLf
    m = m & "  1  -  Link customer" & vbCrLf & _
            "  2  -  Pick the sheet from stock" & vbCrLf & _
            "  3  -  Export DXF (automatic name)" & vbCrLf & _
            "  4  -  Toolpath folders" & vbCrLf & _
            "  5  -  Used parts from the quote (u1350x1000 files)" & vbCrLf & _
            "  6  -  Open the customer folder" & vbCrLf & _
            "  7  -  Job setup (server, key, day folder)" & vbCrLf & _
            "  8  -  Old jobs of this customer" & vbCrLf & _
            "  9  -  Check the drawing before export" & vbCrLf & _
            " 10  -  Cut time and price" & vbCrLf & _
            " 11  -  Area (m2) price" & vbCrLf & _
            " 12  -  Day folder of this job (today / an earlier day)"

    If Len(jr.CustFolder) = 0 Then
        dflt = "1"
    ElseIf Len(jr.MatKind) = 0 Then
        dflt = "2"
    Else
        dflt = "3"
    End If

    choice = InputBox(m, ac2fTitle(CAPTION_), dflt)
    If StrPtr(choice) = 0 Then Exit Sub
    Select Case Trim$(choice)
        Case "": Exit Sub
        Case "1": ac2fJobLink
        Case "2": ac2fJobMaterial
        Case "3": ac2fJobExport
        Case "4": ac2fJobFolders
        Case "5": ac2fJobUsage
        Case "6": ac2fJobOpen
        Case "7": ac2fJobSetup
        Case "8": ac2fJobHistory
        Case "9": ac2fCheck
        Case "10": ac2fCutTime
        Case "11": ac2fCutArea
        Case "12": ac2fJobDay
        Case Else
            ac2fWarn "Invalid choice: " & choice, CAPTION_
    End Select
End Sub

Public Sub ac2fJobLink()
Attribute ac2fJobLink.VB_Description = "ac2f pack: Link the document to a customer from StockTrack"
    Dim jr As ac2fJobRec
    Dim where As String, errText As String

    If Not ac2fJBNeedDoc() Then Exit Sub
    ac2fJBLoad jr
    If Not ac2fJBPickCustomer(jr) Then Exit Sub
    ac2fJBSave jr

    where = ac2fJBCustomerDir(jr, False, errText)
    If Len(where) = 0 Then where = errText
    ac2fInfo "Linked to:  " & jr.CustLabel & vbCrLf & vbCrLf & _
             "Folder:" & vbCrLf & "   " & where & vbCrLf & vbCrLf & _
             "Export (3) asks for the sheet if none is picked yet.", CAPTION_
End Sub

Public Sub ac2fJobMaterial()
Attribute ac2fJobMaterial.VB_Description = "ac2f pack: Pick the sheet for this job from StockTrack stock"
    Dim jr As ac2fJobRec
    Dim s As String

    If Not ac2fJBNeedDoc() Then Exit Sub
    ac2fJBLoad jr
    If Not ac2fJBPickMaterial(jr) Then Exit Sub
    ac2fJBSave jr

    s = "Sheet:  " & ac2fJBMatText(jr)
    If jr.MatOwner = "C" Then s = s & vbCrLf & "(the customer's own material)"
    If ac2fGetLng(AC2F_K_JB_NEST, AC2F_DEF_JB_NEST) <> 0 And jr.MatW > 0 And jr.MatH > 0 Then
        s = s & vbCrLf & vbCrLf & "Nesting now packs onto " & jr.MatW & " x " & jr.MatH & " mm."
    End If
    ac2fInfo s, CAPTION_
End Sub

' Exports the selection (or the page) as DXF into the customer folder,
' named exp01--pleksi.dxf, exp02-- and so on. Asks for whatever the job
' is still missing first, so on a fresh file this one macro does it all.
Public Sub ac2fJobExport()
Attribute ac2fJobExport.VB_Description = "ac2f pack: Export DXF into the customer folder with an automatic name"
    Dim jr As ac2fJobRec
    Dim sr As ShapeRange
    Dim rng As Long
    Dim custDir As String, fname As String, errText As String
    Dim flt As Object

    If Not ac2fJBNeedDoc() Then Exit Sub

    rng = cdrSelection
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        rng = cdrCurrentPage
    ElseIf sr.Count = 0 Then
        rng = cdrCurrentPage
    End If
    If rng = cdrCurrentPage Then
        If MsgBox("Nothing is selected. Export the whole page?", _
                  vbQuestion + vbYesNo, ac2fTitle(CAPTION_)) <> vbYes Then Exit Sub
    End If

    ac2fJBLoad jr
    If Not ac2fJBEnsure(jr) Then Exit Sub

    ' text left as text, bitmaps, double lines, parts off the sheet
    If ac2fGetLng(AC2F_K_JB_CHECK, AC2F_DEF_JB_CHECK) <> 0 Then
        If Not ac2fCheckBeforeExport(sr) Then Exit Sub
    End If

    custDir = ac2fJBCustomerDir(jr, True, errText, True)
    If Len(custDir) = 0 Then
        If Len(errText) > 0 Then ac2fWarn errText, CAPTION_
        Exit Sub
    End If
    ac2fJBSave jr
    fname = "exp" & ac2fJBNextExport(custDir) & ac2fJBSep() & jr.MatKind & ".dxf"

    On Error GoTo Fail
    Set flt = ActiveDocument.ExportEx(custDir & "\" & fname, cdrDXF, rng)
    If ac2fGetLng(AC2F_K_JB_DIALOG, AC2F_DEF_JB_DIALOG) <> 0 Then
        If flt.HasDialog Then
            If Not flt.ShowDialog() Then
                ac2fInfo "Export cancelled. Nothing was written.", CAPTION_
                Exit Sub
            End If
        End If
    End If
    flt.Finish
    On Error GoTo 0

    If Not ac2fJBFileExists(custDir & "\" & fname) Then
        ac2fWarn "CorelDRAW reported no error, but this file is not there:" & vbCrLf & _
                 custDir & "\" & fname, CAPTION_
        Exit Sub
    End If

    If ac2fGetLng(AC2F_K_JB_OPEN, AC2F_DEF_JB_OPEN) <> 0 Then ac2fJBExplore custDir
    ac2fInfo "Exported:" & vbCrLf & "   " & fname & vbCrLf & vbCrLf & _
             "into:" & vbCrLf & "   " & custDir & vbCrLf & vbCrLf & _
             "Sheet: " & ac2fJBMatText(jr) & vbCrLf & vbCrLf & _
             "Wrong day? Customer job menu > 12 (Day folder).", CAPTION_
    Exit Sub

Fail:
    ac2fWarn "The export failed: " & Err.Description & vbCrLf & vbCrLf & _
             custDir & "\" & fname, CAPTION_
End Sub

' Creates the toolpath folders for the sheets this job uses, or stacks
' them onto an existing folder for the same sheet. The number of sheets
' is read off the drawing: Nesting's sheet outlines when there are any,
' otherwise one sheet for the selection.
Public Sub ac2fJobFolders()
Attribute ac2fJobFolders.VB_Description = "ac2f pack: Create or stack the toolpath folders for this job"
    Dim jr As ac2fJobRec
    Dim custDir As String, errText As String, why As String
    Dim proposal As String, answer As String, m As String
    Dim toks() As String, cnts() As Long, n As Long

    If Not ac2fJBNeedDoc() Then Exit Sub
    ac2fJBLoad jr
    If Not ac2fJBEnsure(jr) Then Exit Sub

    custDir = ac2fJBCustomerDir(jr, False, errText, True)
    If Len(custDir) = 0 Then
        If Len(errText) > 0 Then ac2fWarn errText, CAPTION_
        Exit Sub
    End If
    ac2fJBSave jr

    proposal = ac2fJBPropose(jr, why)
    Do
        m = "TOOLPATH FOLDERS  -  " & jr.CustFolder & vbCrLf
        m = m & "Sheet: " & ac2fJBMatText(jr) & vbCrLf & vbCrLf
        m = m & why & vbCrLf & vbCrLf
        m = m & "Enter            accept what is shown below" & vbCrLf
        m = m & "3                three of the job's sheet" & vbCrLf
        m = m & "1 2050x3050      count and size of another sheet" & vbCrLf
        m = m & "2, 1 1220x1300   several sizes at once" & vbCrLf
        m = m & "M                pick another sheet first" & vbCrLf & vbCrLf
        m = m & "Every sheet counts whole here; the part a job used" & vbCrLf
        m = m & "is marked from the quote later (job menu 5)."
        answer = InputBox(m, ac2fTitle(CAPTION_), proposal)
        If StrPtr(answer) = 0 Then Exit Sub
        answer = Trim$(answer)
        If Len(answer) = 0 Then Exit Sub

        If UCase$(answer) = "M" Then
            If ac2fJBPickMaterial(jr) Then
                ac2fJBSave jr
                proposal = ac2fJBPropose(jr, why)
            End If
        Else
            n = ac2fJBParsePlan(jr, answer, toks, cnts, errText)
            If n > 0 Then Exit Do
            If n = 0 Then errText = "Nothing to create."
            ac2fWarn errText, CAPTION_
            proposal = answer
        End If
    Loop

    ac2fJBMakeFolders jr, custDir, toks, cnts, n
End Sub

' Run after the quote is made. Reads the quote's sheet lines from
' StockTrack and, for every sheet only part of which was used, puts a
' file named after that part (u1350x1000) into the sheet's toolpath
' folder. A whole sheet gets nothing; a file already there is kept.
Public Sub ac2fJobUsage()
Attribute ac2fJobUsage.VB_Description = "ac2f pack: After the quote, mark the part of each sheet it used (u1350x1000 file in the folder)"
    Dim jr As ac2fJobRec
    Dim qrows() As String, irows() As String
    Dim nq As Long, ni As Long, i As Long, k As Long
    Dim m As String, answer As String, errText As String
    Dim labels() As String, act As String, typed As String

    If Not ac2fJBNeedDoc() Then Exit Sub
    ac2fJBLoad jr
    If Len(jr.CustFolder) = 0 Then
        If Not ac2fJBPickCustomer(jr) Then Exit Sub
        ac2fJBSave jr
    End If
    If Len(jr.CustId) = 0 Then
        ac2fWarn jr.CustFolder & " was typed in, not taken from StockTrack," & vbCrLf & _
                 "so there is no quote to read. Link the customer from" & vbCrLf & _
                 "StockTrack (job menu 1) and run this again.", CAPTION_
        Exit Sub
    End If

    nq = ac2fJBGetRows("/corel/quotes?limit=6&customerId=" & jr.CustId, qrows, errText)
    If nq < 0 Then
        ac2fWarn errText, CAPTION_
        Exit Sub
    End If
    If nq = 0 Then
        ac2fWarn "StockTrack has no quote for " & jr.CustLabel & " yet." & vbCrLf & _
                 "Make the quote, then run this again.", CAPTION_
        Exit Sub
    End If

    k = 1
    If nq > 1 Then
        ReDim labels(1 To nq)
        For i = 1 To nq
            labels(i) = ac2fJBTab(qrows(i), 1) & "   " & ac2fJBTab(qrows(i), 2) & "   " & _
                        ac2fJBTab(qrows(i), 3) & "   " & ac2fJBTab(qrows(i), 4) & " sheet line(s)"
        Next i
        k = ac2fJBListWin("USED PARTS  -  " & jr.CustLabel & "~nWhich quote? Newest first.", _
                          labels, nq, "Use this quote", "", "", _
                          "Type to narrow the list; Enter takes the highlighted quote.", act, typed)
        If k = -1 Or k = 0 Then Exit Sub
    End If
    If k = -2 Then
        m = "USED PARTS  -  " & jr.CustFolder & vbCrLf & vbCrLf & _
            "Which quote? Newest first." & vbCrLf & vbCrLf
        For i = 1 To nq
            m = m & ac2fRPad(CStr(i), 3) & "  " & labels(i) & vbCrLf
        Next i
        Do
            answer = InputBox(m, ac2fTitle(CAPTION_), "1")
            If StrPtr(answer) = 0 Then Exit Sub
            answer = Trim$(answer)
            If Len(answer) = 0 Then Exit Sub
            If ac2fJBIsCount(answer) Then
                k = CLng(answer)
                If k >= 1 And k <= nq Then Exit Do
            End If
            ac2fWarn "Type a number from the list.", CAPTION_
        Loop
    End If

    ni = ac2fJBGetRows("/corel/quotes/" & ac2fJBTab(qrows(k), 0) & "/usage", irows, errText)
    If ni < 0 Then
        ac2fWarn errText, CAPTION_
        Exit Sub
    End If
    If ni = 0 Then
        ac2fInfo "Quote " & ac2fJBTab(qrows(k), 1) & " has no sheet from stock on it.", CAPTION_
        Exit Sub
    End If
    ac2fJBMarkUsage jr, qrows(k), irows, ni
End Sub

Public Sub ac2fJobOpen()
Attribute ac2fJobOpen.VB_Description = "ac2f pack: Open the customer folder of this job in Explorer"
    Dim jr As ac2fJobRec
    Dim custDir As String, errText As String

    If Not ac2fJBNeedDoc() Then Exit Sub
    ac2fJBLoad jr
    If Len(jr.CustFolder) = 0 Then
        If Not ac2fJBPickCustomer(jr) Then Exit Sub
        ac2fJBSave jr
    End If
    custDir = ac2fJBCustomerDir(jr, True, errText, True)
    If Len(custDir) = 0 Then
        If Len(errText) > 0 Then ac2fWarn errText, CAPTION_
        Exit Sub
    End If
    ac2fJBSave jr
    ac2fJBExplore custDir
End Sub

Public Sub ac2fJobSetup()
Attribute ac2fJobSetup.VB_Description = "ac2f pack: StockTrack server, key, day folder and separator"
    Dim m As String, answer As String, cmd As String, arg As String
    Dim rows() As String, n As Long, p As Long
    Dim errText As String, today As String, ovr As String

    Do
        ovr = ac2fGetStr(AC2F_K_JB_BASE, "")
        today = ac2fJBDayDir(False, errText)
        If Len(today) = 0 Then today = "<not found: see 3>"

        m = "JOB SETUP" & vbCrLf & vbCrLf
        m = m & "1  Server     " & ac2fJBOr(ac2fJBServer(), "<not set>") & vbCrLf
        m = m & "2  Key        " & ac2fJBMask(ac2fGetStr(AC2F_K_JB_KEY, "")) & vbCrLf
        m = m & "3  Day base   " & ac2fJBOr(ovr, "<from CDR Auto Renamer>") & vbCrLf
        m = m & "   today  ->  " & ac2fJBClip(today, 64) & vbCrLf
        m = m & "4  Separator  " & ac2fJBSep() & vbCrLf & vbCrLf
        m = m & "Server and key are on StockTrack's Settings page," & vbCrLf
        m = m & "under Corel integration. Paste them as:" & vbCrLf
        m = m & "   1=http://192.168.1.50/api/v1" & vbCrLf
        m = m & "   2=<key>" & vbCrLf
        m = m & "3=  with nothing after it goes back to CDR Auto" & vbCrLf
        m = m & "Renamer's daily folder." & vbCrLf & vbCrLf
        m = m & "T  test the connection        Enter  close"

        answer = InputBox(m, ac2fTitle("Job Setup"), "")
        If StrPtr(answer) = 0 Then Exit Sub
        answer = Trim$(answer)
        If Len(answer) = 0 Then Exit Sub

        If UCase$(answer) = "T" Then
            n = ac2fJBGetRows("/corel/ping", rows, errText)
            If n < 0 Then
                ac2fWarn errText, "Job Setup"
            ElseIf n = 0 Then
                ac2fWarn "The server answered with an empty list.", "Job Setup"
            Else
                ac2fInfo "Connected to StockTrack:  " & ac2fJBTab(rows(1), 1), "Job Setup"
            End If
        Else
            p = InStr(answer, "=")
            If p < 2 Then
                ac2fWarn "Type the number, = and the value, e.g. 4=--", "Job Setup"
            Else
                cmd = Trim$(Left$(answer, p - 1))
                arg = Trim$(Mid$(answer, p + 1))
                Select Case cmd
                    Case "1": ac2fSetStr AC2F_K_JB_SERVER, ac2fJBNormServer(arg)
                    Case "2": ac2fSetStr AC2F_K_JB_KEY, arg
                    Case "3": ac2fSetStr AC2F_K_JB_BASE, ac2fJBNoSlash(arg)
                    Case "4"
                        If Len(arg) = 0 Then arg = AC2F_DEF_JB_SEP
                        If Not ac2fJBSepOk(arg) Then
                            ac2fWarn "A separator is plain ASCII, without spaces, and cannot hold" & vbCrLf & _
                                     "\ / : * ? "" < > |", "Job Setup"
                        Else
                            ac2fSetStr AC2F_K_JB_SEP, arg
                        End If
                    Case Else
                        ac2fWarn "No setting " & cmd & ".", "Job Setup"
                End Select
            End If
        End If
    Loop
End Sub

' Earlier work for this customer: every day folder in the last year that
' holds a folder for them, newest first. Pick one to open it, open its
' CorelDRAW file, or carry this job on in it.
' Which day folder this job's customer folder is in: today, or an earlier
' day that already has a folder for this customer. Export asks this once
' and keeps the answer with the job; this is where it is changed, also on
' the same day and also back to today.
Public Sub ac2fJobDay()
Attribute ac2fJobDay.VB_Description = "ac2f pack: Choose or change the day folder this job goes into"
    Dim jr As ac2fJobRec
    Dim root As String, fmt As String, dsep As String, errText As String
    Dim todayDir As String, todayName As String, cur As String, cdir As String
    Dim days() As String, cand() As String, info() As String, labels() As String
    Dim n As Long, nd As Long, back As Long, lim As Long, i As Long, k As Long
    Dim act As String, typed As String, m As String, answer As String

    If Not ac2fJBNeedDoc() Then Exit Sub
    ac2fJBLoad jr
    If Len(jr.CustFolder) = 0 Then
        If Not ac2fJBPickCustomer(jr) Then Exit Sub
        ac2fJBSave jr
    End If
    todayDir = ac2fJBDayDir(False, errText)
    If Len(todayDir) = 0 Or Not ac2fJBDayRule(root, fmt, dsep) Then
        ac2fWarn ac2fJBOr(errText, "The daily folder is not known (Job setup, 3)."), CAPTION_
        Exit Sub
    End If
    root = Left$(todayDir, InStrRev(todayDir, "\") - 1)
    todayName = Mid$(todayDir, InStrRev(todayDir, "\") + 1)

    ' where the job goes now (only looking, nothing is asked)
    cdir = ac2fJBCustomerDir(jr, False, errText)
    If Len(cdir) > 0 Then
        cur = Left$(cdir, InStrRev(cdir, "\") - 1)
        cur = Mid$(cur, InStrRev(cur, "\") + 1)
    End If

    ' today first, then earlier days that have a folder for this customer
    lim = ac2fGetLng(AC2F_K_JB_BACK, AC2F_DEF_JB_BACK)
    If lim < 7 Then lim = 7
    ReDim cand(1 To 16): ReDim info(1 To 16)
    n = 1
    cand(1) = todayName
    info(1) = "today, " & ac2fJBDayInfo(root & "\" & todayName, jr.CustFolder)
    For back = 1 To lim
        nd = ac2fJBListNames(root, ac2fJBStrftime(fmt, Date - back) & "*", True, days)
        For i = 1 To nd
            If StrComp(days(i), todayName, vbTextCompare) <> 0 Then
                If Len(ac2fJBCustSub(root & "\" & days(i), jr.CustFolder)) > 0 Then
                    n = n + 1
                    If n > UBound(cand) Then
                        ReDim Preserve cand(1 To n * 2)
                        ReDim Preserve info(1 To n * 2)
                    End If
                    cand(n) = days(i)
                    info(n) = ac2fJBDayInfo(root & "\" & days(i), jr.CustFolder)
                End If
            End If
        Next i
    Next back
    ' the job's own day, if it is older than Look back
    If Len(jr.JobDay) > 0 Then
        k = 0
        For i = 1 To n
            If StrComp(cand(i), jr.JobDay, vbTextCompare) = 0 Then k = i
        Next i
        If k = 0 And Len(ac2fJBCustSub(root & "\" & jr.JobDay, jr.CustFolder)) > 0 Then
            n = n + 1
            If n > UBound(cand) Then
                ReDim Preserve cand(1 To n)
                ReDim Preserve info(1 To n)
            End If
            cand(n) = jr.JobDay
            info(n) = ac2fJBDayInfo(root & "\" & jr.JobDay, jr.CustFolder)
        End If
    End If

    ReDim labels(1 To n)
    For i = 1 To n
        labels(i) = ac2fPad(cand(i), 26) & info(i)
        If StrComp(cand(i), cur, vbTextCompare) = 0 Then labels(i) = labels(i) & "   <- now"
    Next i
    k = ac2fJBListWin("DAY FOLDER  -  " & jr.CustLabel & _
                      "~nExports and toolpath folders of this job go into the chosen day.", _
                      labels, n, "Use this day", "", "", _
                      "Today is " & todayName & ". Earlier days listed have a folder for this customer.", _
                      act, typed)
    If k = -1 Or k = 0 Then Exit Sub
    If k = -2 Then
        ' list window off or failed: numbered prompt
        m = "DAY FOLDER  -  " & jr.CustFolder & vbCrLf & vbCrLf
        For i = 1 To n
            If Len(m) > PROMPT_MAX - 200 Then
                m = m & "     ... " & (n - i + 1) & " more" & vbCrLf
                Exit For
            End If
            m = m & ac2fRPad(CStr(i), 3) & "  " & labels(i) & vbCrLf
        Next i
        m = m & vbCrLf & "Type the number of the day this job goes into."
        Do
            answer = InputBox(m, ac2fTitle(CAPTION_), "1")
            If StrPtr(answer) = 0 Then Exit Sub
            answer = Trim$(answer)
            If Len(answer) = 0 Then Exit Sub
            k = CLng(Val(answer))
            If k >= 1 And k <= n Then Exit Do
            ac2fWarn "Type a number from the list.", CAPTION_
        Loop
    End If

    jr.JobDay = cand(k)
    jr.JobAsked = CStr(CLng(Date))
    ac2fJBSave jr
    ac2fInfo "This job now goes into" & vbCrLf & "   " & _
             ac2fJBCustPath(root & "\" & cand(k), jr.CustFolder) & vbCrLf & vbCrLf & _
             "Save the document to keep this. Change it again with" & vbCrLf & _
             "Customer job menu > 12 (Day folder).", CAPTION_
End Sub

' "2 export, 3 folder" for a customer's folder in a day, or "no folder yet".
Private Function ac2fJBDayInfo(ByVal dayPath As String, ByVal custFolder As String) As String
    Dim names() As String
    Dim cdir As String, h As String, t As String
    Dim nExp As Long, nf As Long, nGrp As Long, i As Long, grpNo As Long, c As Long

    If Len(ac2fJBCustSub(dayPath, custFolder)) = 0 Then
        ac2fJBDayInfo = "no folder yet (it is made on export)"
        Exit Function
    End If
    cdir = ac2fJBCustPath(dayPath, custFolder)
    nExp = ac2fJBListNames(cdir, "exp*", False, names)
    nf = ac2fJBListNames(cdir, "*", True, names)
    For i = 1 To nf
        If ac2fJBSplitGroup(names(i), grpNo, h, c, t) Then nGrp = nGrp + 1
    Next i
    ac2fJBDayInfo = nExp & " export, " & nGrp & " toolpath folder"
End Function

Public Sub ac2fJobHistory()
Attribute ac2fJobHistory.VB_Description = "ac2f pack: Find this customer's earlier jobs and open one"
    Dim jr As ac2fJobRec
    Dim root As String, fmt As String, dsep As String
    Dim days() As String, names() As String
    Dim hitDay() As String, hitInfo() As String, hitCdr() As String
    Dim nHit As Long, nd As Long, nf As Long, back As Long, i As Long, k As Long
    Dim cdir As String, answer As String
    Dim nExp As Long, nGrp As Long, grpNo As Long, c As Long
    Dim h As String, t As String
    Dim labels() As String, act As String, typed As String

    If Not ac2fJBNeedDoc() Then Exit Sub
    ac2fJBLoad jr
    If Len(jr.CustFolder) = 0 Then
        If Not ac2fJBPickCustomer(jr) Then Exit Sub
        ac2fJBSave jr
    End If
    If Not ac2fJBDayRule(root, fmt, dsep) Then
        ac2fWarn "The daily folder is not known (Job setup, 3).", CAPTION_
        Exit Sub
    End If

    ReDim hitDay(1 To 32): ReDim hitInfo(1 To 32): ReDim hitCdr(1 To 32)
    For back = 0 To HISTORY_DAYS
        nd = ac2fJBListNames(root, ac2fJBStrftime(fmt, Date - back) & "*", True, days)
        For i = 1 To nd
            cdir = ac2fJBCustPath(root & "\" & days(i), jr.CustFolder)
            If ac2fJBDirExists(cdir) Then
                nExp = ac2fJBListNames(cdir, "exp*", False, names)
                nf = ac2fJBListNames(cdir, "*", True, names)
                nGrp = 0
                For k = 1 To nf
                    If ac2fJBSplitGroup(names(k), grpNo, h, c, t) Then nGrp = nGrp + 1
                Next k
                nf = ac2fJBListNames(cdir, "*.cdr", False, names)
                nHit = nHit + 1
                If nHit > UBound(hitDay) Then
                    ReDim Preserve hitDay(1 To nHit * 2)
                    ReDim Preserve hitInfo(1 To nHit * 2)
                    ReDim Preserve hitCdr(1 To nHit * 2)
                End If
                hitDay(nHit) = days(i)
                hitInfo(nHit) = nExp & " export, " & nGrp & " folder"
                If nf > 0 Then
                    hitInfo(nHit) = hitInfo(nHit) & ", " & nf & " cdr"
                    hitCdr(nHit) = cdir & "\" & names(nf)
                End If
            End If
        Next i
    Next back

    If nHit = 0 Then
        ac2fInfo "No folder for " & jr.CustFolder & " in the last " & HISTORY_DAYS & _
                 " days.", CAPTION_
        Exit Sub
    End If

    ReDim labels(1 To nHit)
    For i = 1 To nHit
        labels(i) = ac2fPad(hitDay(i), 26) & hitInfo(i)
    Next i
    k = ac2fJBListWin("OLD JOBS  -  " & jr.CustLabel & "~nNewest first.", labels, nHit, _
                      "Open folder", "O=Open its CorelDRAW file|C=Carry this job on here", "", _
                      "Type a date or month to narrow the list (10-2026).", act, typed)
    If k = -1 Then Exit Sub
    If k > 0 Then
        answer = act
    Else
        answer = ac2fJBHistoryPrompt(jr, hitDay, hitInfo, nHit, k)
        If k = 0 Then Exit Sub
    End If

    cdir = ac2fJBCustPath(root & "\" & hitDay(k), jr.CustFolder)
    If Right$(answer, 1) = "C" Then
        jr.JobDay = hitDay(k)
        jr.JobAsked = CStr(CLng(Date))
        ac2fJBSave jr
        ac2fInfo "This job now carries on in" & vbCrLf & "   " & cdir & vbCrLf & vbCrLf & _
                 "Exports and toolpath folders go there today.", CAPTION_
    ElseIf Right$(answer, 1) = "O" Then
        If Len(hitCdr(k)) = 0 Then
            ac2fWarn "There is no .cdr file in that folder; opening the folder.", CAPTION_
            ac2fJBExplore cdir
        Else
            On Error Resume Next
            Application.OpenDocument hitCdr(k)
            If Err.Number <> 0 Then ac2fWarn "Could not open " & hitCdr(k) & vbCrLf & _
                                             Err.Description, CAPTION_
            On Error GoTo 0
        End If
    Else
        ac2fJBExplore cdir
    End If
End Sub

' Old jobs as a numbered prompt, when the list window is not available.
' Returns "", "O" or "C" and the day picked in k (0 on cancel).
Private Function ac2fJBHistoryPrompt(ByRef jr As ac2fJobRec, ByRef hitDay() As String, _
                                     ByRef hitInfo() As String, ByVal nHit As Long, _
                                     ByRef k As Long) As String
    Dim m As String, ln As String, answer As String
    Dim i As Long

    k = 0
    m = "OLD JOBS  -  " & jr.CustFolder & vbCrLf & "Newest first." & vbCrLf & vbCrLf
    For i = 1 To nHit
        ln = ac2fRPad(CStr(i), 3) & "  " & ac2fPad(hitDay(i), 24) & hitInfo(i) & vbCrLf
        If Len(m) + Len(ln) > PROMPT_MAX - 160 Then
            m = m & "     ... " & (nHit - i + 1) & " older" & vbCrLf
            Exit For
        End If
        m = m & ln
    Next i
    m = m & vbCrLf & "A number opens that folder. Add O to open its CorelDRAW" & vbCrLf & _
        "file too (3 O), or C to carry this job on in it (3 C)."

    Do
        answer = InputBox(m, ac2fTitle(CAPTION_), "1")
        If StrPtr(answer) = 0 Then Exit Function
        answer = UCase$(Trim$(answer))
        If Len(answer) = 0 Then Exit Function
        k = CLng(Val(answer))
        If k >= 1 And k <= nHit Then Exit Do
        k = 0
        ac2fWarn "Type a number from the list.", CAPTION_
    Loop
    ac2fJBHistoryPrompt = Right$(answer, 1)
End Function

'=====================================================================
' FOR THE PRICE AND CHECK MODULES
'=====================================================================

' The linked job's sheet: size and thickness (mm) and kind. False when
' the document has no sheet picked.
Public Function ac2fJobSheetInfo(ByRef w As Double, ByRef h As Double, _
                                 ByRef thickMm As Double, ByRef kind As String) As Boolean
    Dim jr As ac2fJobRec
    If ActiveDocument Is Nothing Then Exit Function
    ac2fJBLoad jr
    If Len(jr.MatKind) = 0 Then Exit Function
    w = jr.MatW
    h = jr.MatH
    thickMm = Val(jr.MatThick) / 1000#
    kind = jr.MatKind
    ac2fJobSheetInfo = True
End Function

' The linked customer's name, for headings ("" when none).
Public Function ac2fJobCustomerName() As String
    Dim jr As ac2fJobRec
    If ActiveDocument Is Nothing Then Exit Function
    ac2fJBLoad jr
    ac2fJobCustomerName = jr.CustLabel
End Function

' Cut pricing from StockTrack into v(1 To 8): hourly rate, minimum
' minutes, area rate, minimum area, feed mm/s, rapid mm/s, plunge s,
' pass depth mm; cur is the currency. The last answer is kept, so the
' price modules still work offline; src says where the values came
' from. False only when there is nothing at all.
Public Function ac2fJobPricing(ByRef v() As Double, ByRef cur As String, _
                               ByRef src As String) As Boolean
    Dim rows() As String, f() As String
    Dim n As Long, i As Long
    Dim row As String, errText As String

    ReDim v(1 To 8)
    If Len(ac2fJBServer()) > 0 Then
        n = ac2fJBGetRows("/corel/pricing", rows, errText)
        If n > 0 Then
            row = rows(1)
            ac2fSetStr AC2F_K_JB_PRICE, row
            src = "StockTrack"
        Else
            src = "saved copy (StockTrack did not answer)"
        End If
    Else
        src = "saved copy (StockTrack is not set up)"
    End If
    If Len(row) = 0 Then row = ac2fGetStr(AC2F_K_JB_PRICE, "")
    If Len(row) = 0 Then
        ' Nothing ever read: speeds that make sense, no prices.
        v(5) = 20#: v(6) = 100#: v(7) = 1#
        cur = ""
        src = "defaults: set the prices in StockTrack, Settings"
        Exit Function
    End If

    f = Split(row & String$(9, vbTab), vbTab)
    For i = 1 To 8
        v(i) = Val(f(i - 1))
    Next i
    cur = f(8)
    ac2fJobPricing = True
End Function

'=====================================================================
' PICKING
'=====================================================================

' Fills in whatever the job is still missing. False when cancelled.
Private Function ac2fJBEnsure(ByRef jr As ac2fJobRec) As Boolean
    If Len(jr.CustFolder) = 0 Then
        If Not ac2fJBPickCustomer(jr) Then Exit Function
        ac2fJBSave jr
    End If
    If Len(jr.MatKind) = 0 Then
        If Not ac2fJBPickMaterial(jr) Then Exit Function
        ac2fJBSave jr
    End If
    ac2fJBEnsure = True
End Function

' Recent customers by number, anything else searches StockTrack. A file
' already sitting in a <day>\<customer> folder offers that name.
Private Function ac2fJBPickCustomer(ByRef jr As ac2fJobRec) As Boolean
    Dim rec() As String, nRec As Long
    Dim rows() As String, n As Long
    Dim m As String, answer As String, dflt As String, pending As String
    Dim errText As String, nm As String
    Dim i As Long, k As Long
    Dim online As Boolean
    Dim handled As Boolean

    ac2fJBPickCustomer = ac2fJBPickCustomerWin(jr, handled)
    If handled Then Exit Function

    nRec = ac2fJBRecentLoad(AC2F_K_JB_RCUST, rec)
    online = Len(ac2fJBServer()) > 0
    dflt = ac2fJBFolderHint()
    If Len(dflt) = 0 And nRec > 0 Then dflt = "1"

    Do
        If Len(pending) > 0 Then
            answer = pending
            pending = ""
        Else
            m = "LINK CUSTOMER" & vbCrLf & vbCrLf
            If Len(jr.CustFolder) > 0 Then m = m & "Now:  " & jr.CustLabel & vbCrLf & vbCrLf
            If nRec > 0 Then
                m = m & "Recent:" & vbCrLf
                For i = 1 To nRec
                    m = m & ac2fRPad(CStr(i), 3) & "  " & _
                        ac2fJBClip(ac2fJBTab(rec(i), 2), 56) & vbCrLf
                Next i
                m = m & vbCrLf
            End If
            If online Then
                m = m & "A number, or part of a name or phone to search" & vbCrLf & _
                        "StockTrack. =name takes the name as typed."
            Else
                m = m & "A number, or the customer's name." & vbCrLf & _
                        "(StockTrack is not set up - see Job setup.)"
            End If
            answer = InputBox(m, ac2fTitle(CAPTION_), dflt)
            If StrPtr(answer) = 0 Then Exit Function
            answer = Trim$(answer)
            If Len(answer) = 0 Then Exit Function
        End If

        If ac2fJBIsCount(answer) And nRec > 0 Then
            k = CLng(answer)
            If k >= 1 And k <= nRec Then
                ac2fJBSetCust jr, rec(k)
                ac2fJBPickCustomer = True
                Exit Function
            End If
            ac2fWarn "There is no " & k & " on the list.", CAPTION_
        ElseIf Left$(answer, 1) = "=" Or Not online Then
            If Left$(answer, 1) = "=" Then answer = Trim$(Mid$(answer, 2))
            nm = ac2fJBSafeName(answer)
            If Len(nm) = 0 Then
                ac2fWarn "That name leaves nothing to use as a folder name.", CAPTION_
            Else
                ac2fJBSetCust jr, "" & vbTab & nm & vbTab & nm
                ac2fJBPickCustomer = True
                Exit Function
            End If
        Else
            n = ac2fJBGetRows("/corel/customers?limit=12&q=" & ac2fJBUrl(answer), rows, errText, answer)
            If n < 0 Then
                nm = ac2fJBSafeName(answer)
                If Len(nm) = 0 Then
                    ac2fWarn errText, CAPTION_
                    Exit Function
                End If
                If MsgBox(errText & vbCrLf & vbCrLf & "Use """ & nm & _
                          """ as the customer folder anyway?", _
                          vbQuestion + vbYesNo, ac2fTitle(CAPTION_)) <> vbYes Then Exit Function
                ac2fJBSetCust jr, "" & vbTab & nm & vbTab & nm
                ac2fJBPickCustomer = True
                Exit Function
            ElseIf n = 0 Then
                ac2fWarn "No customer matches """ & answer & """." & vbCrLf & _
                         "Try fewer letters, or =" & answer & " to use it as typed.", CAPTION_
                dflt = answer
            Else
                ' even a single hit is shown, so the wrong customer is never linked unseen
                k = ac2fJBChoose("Customers matching """ & answer & """" & ac2fJBSavedNote() & ":", _
                                 rows, n, 2, pending)
                If k > 0 Then
                    ac2fJBSetCust jr, rows(k)
                    ac2fJBPickCustomer = True
                    Exit Function
                End If
                If k < 0 Then Exit Function
                dflt = answer
            End If
        End If
    Loop
End Function

' The customer's own sheets first, then recent picks, then the rest of
' the stock. Words search StockTrack; =kind;thick;... types one in.
Private Function ac2fJBPickMaterial(ByRef jr As ac2fJobRec) As Boolean
    Dim rec() As String, nRec As Long
    Dim srv() As String, nSrv As Long
    Dim opts() As String, nOpt As Long
    Dim m As String, answer As String, dflt As String, row As String
    Dim errText As String, cq As String, head As String
    Dim k As Long
    Dim online As Boolean
    Dim handled As Boolean

    nRec = ac2fJBRecentLoad(AC2F_K_JB_RMAT, rec)
    online = Len(ac2fJBServer()) > 0
    If Len(jr.CustId) > 0 Then cq = "&customerId=" & jr.CustId

    If online Then
        nSrv = ac2fJBGetRows("/corel/materials?limit=30" & cq, srv, errText)
        If nSrv < 0 Then
            ac2fWarn errText & vbCrLf & vbCrLf & _
                     "Only recent sheets are offered; others can be typed.", CAPTION_
            nSrv = 0
            online = False
        End If
    End If
    nOpt = ac2fJBMergeMat(srv, nSrv, rec, nRec, opts)
    If nOpt > 0 Then dflt = "1"
    head = "PICK THE SHEET"
    If Len(jr.CustFolder) > 0 Then head = head & "  -  " & jr.CustFolder
    head = head & ac2fJBSavedNote()

    ac2fJBPickMaterial = ac2fJBPickMaterialWin(jr, opts, nOpt, online, cq, head, handled)
    If handled Then Exit Function

    Do
        m = ac2fJBMatPrompt(head, opts, nOpt, online)
        answer = InputBox(m, ac2fTitle(CAPTION_), dflt)
        If StrPtr(answer) = 0 Then Exit Function
        answer = Trim$(answer)
        If Len(answer) = 0 Then Exit Function

        If ac2fJBIsCount(answer) Then
            k = CLng(answer)
            If k >= 1 And k <= nOpt Then
                ac2fJBSetMat jr, opts(k)
                ac2fJBPickMaterial = True
                Exit Function
            End If
            ac2fWarn "There is no " & k & " on the list.", CAPTION_
        ElseIf Left$(answer, 1) = "=" Then
            row = ac2fJBTypedMat(Mid$(answer, 2))
            If Len(row) = 0 Then
                ac2fWarn "Type it as kind;thickness;brand;variant;colour;size," & vbCrLf & _
                         "e.g. =pleksi;2.8;akpolimer;dokme;beyaz;1220x2440" & vbCrLf & _
                         "Only the kind is required.", CAPTION_
                dflt = answer
            Else
                ac2fJBSetMat jr, row
                ac2fJBPickMaterial = True
                Exit Function
            End If
        ElseIf online Then
            nSrv = ac2fJBGetRows("/corel/materials?limit=15" & cq & "&q=" & ac2fJBUrl(answer), _
                                 srv, errText, answer)
            If nSrv < 0 Then
                ac2fWarn errText, CAPTION_
            ElseIf nSrv = 0 Then
                ac2fWarn "Nothing in stock matches """ & answer & """.", CAPTION_
                dflt = answer
            Else
                opts = srv
                nOpt = nSrv
                head = "SHEETS MATCHING """ & answer & """" & ac2fJBSavedNote()
                dflt = "1"
            End If
        Else
            ac2fWarn "StockTrack is not available, so type the sheet:" & vbCrLf & _
                     "=pleksi;2.8;akpolimer;dokme;beyaz;1220x2440", CAPTION_
            dflt = "=" & ac2fJBToken(answer)
        End If
    Loop
End Function

' The customer from a list window: recent picks and StockTrack's latest
' customers to start with. Typing narrows the list; Search StockTrack
' (or Enter when nothing on the list matches) asks the server, and every
' hit is listed, even a single one, to be picked by hand. Use as typed
' links a name StockTrack does not have. handled is False when the window
' cannot be shown.
Private Function ac2fJBPickCustomerWin(ByRef jr As ac2fJobRec, ByRef handled As Boolean) As Boolean
    Dim rows() As String, srv() As String, labels() As String
    Dim n As Long, nSrv As Long, k As Long, i As Long
    Dim head As String, hint As String, textBtns As String
    Dim act As String, typed As String, errText As String, nm As String
    Dim online As Boolean

    online = Len(ac2fJBServer()) > 0
    n = ac2fJBRecentLoad(AC2F_K_JB_RCUST, rows)
    head = "LINK CUSTOMER"
    If Len(jr.CustFolder) > 0 Then head = head & "   (now: " & jr.CustLabel & ")"
    If online Then
        nSrv = ac2fJBGetRows("/corel/customers?limit=40", srv, errText)
        If nSrv > 0 Then
            n = ac2fJBAddRows(rows, n, srv, nSrv)
            head = head & ac2fJBSavedNote()
        End If
        textBtns = "S=Search StockTrack|T=Use as typed"
        hint = "Type part of a name or phone. Enter links the highlighted one; " & _
               "with nothing left on the list it searches StockTrack (F3)."
    Else
        textBtns = "T=Use as typed"
        hint = "StockTrack is not set up (Job setup), so type the name and press Enter."
    End If

    Do
        ReDim labels(1 To n + 1)
        For i = 1 To n
            labels(i) = ac2fJBOr(ac2fJBTab(rows(i), 2), ac2fJBTab(rows(i), 1))
        Next i
        k = ac2fJBListWin(head, labels, n, "Link", "", textBtns, hint, act, typed)
        If k = -2 Then Exit Function
        handled = True
        If k = -1 Then Exit Function
        If k > 0 Then
            ac2fJBSetCust jr, rows(k)
            ac2fJBPickCustomerWin = True
            Exit Function
        End If

        If Left$(typed, 1) = "=" Then
            typed = Trim$(Mid$(typed, 2))
            act = "T"
        End If
        If act = "T" Or Not online Then
            nm = ac2fJBSafeName(typed)
            If Len(nm) = 0 Then
                ac2fWarn "That name leaves nothing to use as a folder name.", CAPTION_
            Else
                ac2fJBSetCust jr, "" & vbTab & nm & vbTab & typed
                ac2fJBPickCustomerWin = True
                Exit Function
            End If
        Else
            nSrv = ac2fJBGetRows("/corel/customers?limit=40&q=" & ac2fJBUrl(typed), srv, errText, typed)
            If nSrv < 0 Then
                ac2fWarn errText & vbCrLf & vbCrLf & _
                         "Use as typed links a name without StockTrack.", CAPTION_
            ElseIf nSrv = 0 Then
                ac2fWarn "No customer matches """ & typed & """." & vbCrLf & _
                         "Try fewer letters, or Use as typed.", CAPTION_
            Else
                rows = srv
                n = nSrv
                head = "CUSTOMERS MATCHING """ & typed & """" & ac2fJBSavedNote()
            End If
        End If
    Loop
End Function

' The sheet from a list window: the customer's own sheets, recent picks
' and the stock to start with, narrowed as you type. Search stock asks
' StockTrack and lists every hit; Use as typed takes
' =kind;thickness;brand;variant;colour;size. handled is False when the
' window cannot be shown.
Private Function ac2fJBPickMaterialWin(ByRef jr As ac2fJobRec, ByRef opts() As String, _
                                       ByVal nOpt As Long, ByVal online As Boolean, _
                                       ByVal cq As String, ByVal head As String, _
                                       ByRef handled As Boolean) As Boolean
    Dim labels() As String, srv() As String
    Dim nSrv As Long, k As Long, i As Long
    Dim hint As String, textBtns As String, act As String, typed As String
    Dim errText As String, row As String

    If online Then
        textBtns = "S=Search stock|T=Use as typed"
        hint = "Type to narrow the list (pleksi beyaz 3). Enter picks the highlighted " & _
               "sheet; with nothing left it searches the stock (F3).~n" & _
               "A sheet not in stock: =pleksi;2.8;akpolimer;dokme;beyaz;1220x2440, then Use as typed."
    Else
        textBtns = "T=Use as typed"
        hint = "Type to narrow the list. A sheet not on it: " & _
               "=pleksi;2.8;akpolimer;dokme;beyaz;1220x2440, then Enter."
    End If

    Do
        ReDim labels(1 To nOpt + 1)
        For i = 1 To nOpt
            labels(i) = ac2fJBTab(opts(i), 10)
            If Len(labels(i)) = 0 Or ac2fJBTab(opts(i), 1) = "M" Then labels(i) = ac2fJBRowText(opts(i))
        Next i
        k = ac2fJBListWin(head, labels, nOpt, "Use this sheet", "", textBtns, hint, act, typed)
        If k = -2 Then Exit Function
        handled = True
        If k = -1 Then Exit Function
        If k > 0 Then
            ac2fJBSetMat jr, opts(k)
            ac2fJBPickMaterialWin = True
            Exit Function
        End If

        If Left$(typed, 1) = "=" Or act = "T" Or Not online Then
            If Left$(typed, 1) = "=" Then typed = Mid$(typed, 2)
            row = ac2fJBTypedMat(typed)
            If Len(row) = 0 Then
                ac2fWarn "Type it as kind;thickness;brand;variant;colour;size," & vbCrLf & _
                         "e.g. =pleksi;2.8;akpolimer;dokme;beyaz;1220x2440" & vbCrLf & _
                         "Only the kind is required.", CAPTION_
            Else
                ac2fJBSetMat jr, row
                ac2fJBPickMaterialWin = True
                Exit Function
            End If
        Else
            nSrv = ac2fJBGetRows("/corel/materials?limit=40" & cq & "&q=" & ac2fJBUrl(typed), _
                                 srv, errText, typed)
            If nSrv < 0 Then
                ac2fWarn errText, CAPTION_
            ElseIf nSrv = 0 Then
                ac2fWarn "Nothing in stock matches """ & typed & """.", CAPTION_
            Else
                opts = srv
                nOpt = nSrv
                head = "SHEETS MATCHING """ & typed & """" & ac2fJBSavedNote()
            End If
        End If
    Loop
End Function

' rows plus the rows of more that are not on it yet (same StockTrack id,
' or the same folder for typed ones). Returns the new count.
Private Function ac2fJBAddRows(ByRef rows() As String, ByVal n As Long, _
                               ByRef more() As String, ByVal nMore As Long) As Long
    Dim i As Long, j As Long
    Dim dup As Boolean

    ReDim Preserve rows(1 To n + nMore + 1)
    For i = 1 To nMore
        dup = False
        For j = 1 To n
            If Len(ac2fJBTab(more(i), 0)) > 0 Then
                dup = ac2fJBTab(rows(j), 0) = ac2fJBTab(more(i), 0)
            Else
                dup = ac2fJBSafeName(ac2fJBTab(rows(j), 1)) = ac2fJBSafeName(ac2fJBTab(more(i), 1))
            End If
            If dup Then Exit For
        Next j
        If Not dup Then
            n = n + 1
            rows(n) = more(i)
        End If
    Next i
    ac2fJBAddRows = n
End Function

' Numbered list of rows showing field col. Returns the number picked,
' 0 when something else was typed (handed back in typed), -1 on cancel.
Private Function ac2fJBChoose(ByVal head As String, ByRef rows() As String, _
                              ByVal n As Long, ByVal col As Long, _
                              ByRef typed As String) As Long
    Dim m As String, ln As String, answer As String
    Dim i As Long, k As Long

    m = head & vbCrLf & vbCrLf
    For i = 1 To n
        ln = ac2fRPad(CStr(i), 3) & "  " & ac2fJBClip(ac2fJBTab(rows(i), col), 60) & vbCrLf
        If Len(m) + Len(ln) > PROMPT_MAX Then Exit For
        m = m & ln
    Next i
    m = m & vbCrLf & "A number, or new words to search again."

    Do
        answer = InputBox(m, ac2fTitle(CAPTION_), "1")
        If StrPtr(answer) = 0 Then
            ac2fJBChoose = -1
            Exit Function
        End If
        answer = Trim$(answer)
        If Len(answer) = 0 Then
            ac2fJBChoose = -1
            Exit Function
        End If
        If Not ac2fJBIsCount(answer) Then Exit Do
        k = CLng(answer)
        If k >= 1 And k <= n Then
            ac2fJBChoose = k
            Exit Function
        End If
        ac2fWarn "There is no " & k & " on the list.", CAPTION_
    Loop
    typed = answer
End Function

' Customer's own sheets, then recent picks (a fresh copy from the server
' when it has one), then the rest of the stock.
Private Function ac2fJBMergeMat(ByRef srv() As String, ByVal nSrv As Long, _
                                ByRef rec() As String, ByVal nRec As Long, _
                                ByRef opts() As String) As Long
    Dim used() As Boolean
    Dim i As Long, j As Long, n As Long
    Dim id As String

    ReDim opts(1 To nSrv + nRec + 1)
    ReDim used(0 To nSrv)

    For i = 1 To nSrv
        If ac2fJBTab(srv(i), 1) = "C" Then
            n = n + 1
            opts(n) = srv(i)
            used(i) = True
        End If
    Next i
    For j = 1 To nRec
        id = ac2fJBTab(rec(j), 0)
        For i = 1 To nSrv
            If Not used(i) And Len(id) > 0 And ac2fJBTab(srv(i), 0) = id _
               And ac2fJBTab(srv(i), 1) = "B" Then Exit For
        Next i
        n = n + 1
        If i <= nSrv Then
            opts(n) = srv(i)
            used(i) = True
        Else
            opts(n) = rec(j)
        End If
    Next j
    For i = 1 To nSrv
        If Not used(i) Then
            n = n + 1
            opts(n) = srv(i)
        End If
    Next i
    ac2fJBMergeMat = n
End Function

Private Function ac2fJBMatPrompt(ByVal head As String, ByRef opts() As String, _
                                 ByVal n As Long, ByVal online As Boolean) As String
    Dim m As String, ln As String, tail As String, lbl As String
    Dim i As Long, shown As Long

    If online Then
        tail = "A number, or words to search the stock (pleksi beyaz 3)." & vbCrLf
    Else
        tail = "A number, or type the sheet:" & vbCrLf
    End If
    tail = tail & "=pleksi;2.8;akpolimer;dokme;beyaz;1220x2440 types one in."

    m = head & vbCrLf & vbCrLf
    For i = 1 To n
        lbl = ac2fJBTab(opts(i), 10)
        If Len(lbl) = 0 Or ac2fJBTab(opts(i), 1) = "M" Then lbl = ac2fJBRowText(opts(i))
        ln = ac2fRPad(CStr(i), 3) & "  " & ac2fJBClip(lbl, 62) & vbCrLf
        If Len(m) + Len(ln) + Len(tail) > PROMPT_MAX Then Exit For
        m = m & ln
        shown = shown + 1
    Next i
    If n = 0 Then m = m & "(nothing to list)" & vbCrLf
    If shown < n Then m = m & "     ... " & (n - shown) & " more: search to narrow it" & vbCrLf
    ac2fJBMatPrompt = m & vbCrLf & tail
End Function

' "pleksi;2.8;akpolimer;dokme;beyaz;1220x2440" -> a material row, the
' same shape StockTrack sends. Empty when there is no kind.
Private Function ac2fJBTypedMat(ByVal s As String) As String
    Dim f() As String
    Dim kind As String, thick As String
    Dim w As Double, h As Double
    Dim r As String

    f = Split(s & ";;;;;", ";")
    kind = ac2fJBToken(f(0))
    If Len(kind) = 0 Then Exit Function
    If ac2fParseNum(f(1), 0) > 0 Then thick = CStr(CLng(ac2fParseNum(f(1), 0) * 1000#))
    If Len(Trim$(f(5))) > 0 Then
        If Not ac2fJBParseSize(f(5), w, h) Then Exit Function
    End If

    r = "" & vbTab & "M" & vbTab & kind & vbTab & thick & vbTab & _
        ac2fJBToken(f(2)) & vbTab & ac2fJBToken(f(3)) & vbTab & ac2fJBToken(f(4)) & vbTab
    If w > 0 Then
        r = r & CStr(CLng(w)) & vbTab & CStr(CLng(h))
    Else
        r = r & vbTab
    End If
    r = r & vbTab & "" & vbTab
    ac2fJBTypedMat = r & ac2fJBRowText(r)
End Function

Private Sub ac2fJBSetCust(ByRef jr As ac2fJobRec, ByVal row As String)
    Dim changed As Boolean

    ' rows saved before folder names went lower case with dashes, and
    ' typed names, all come out the same way
    changed = ac2fJBTab(row, 0) <> jr.CustId Or ac2fJBSafeName(ac2fJBTab(row, 1)) <> jr.CustFolder
    jr.CustId = ac2fJBTab(row, 0)
    jr.CustFolder = ac2fJBSafeName(ac2fJBTab(row, 1))
    jr.CustLabel = ac2fJBOr(ac2fJBTab(row, 2), jr.CustFolder)
    ' the previous customer's own material and day are not this one's
    If changed And jr.MatOwner = "C" Then ac2fJBSetMat jr, ""
    If changed Then
        jr.JobDay = ""
        jr.JobAsked = ""
    End If
    ac2fJBRecentPush AC2F_K_JB_RCUST, row
End Sub

Private Sub ac2fJBSetMat(ByRef jr As ac2fJobRec, ByVal row As String)
    jr.MatId = ac2fJBTab(row, 0)
    jr.MatOwner = ac2fJBTab(row, 1)
    jr.MatKind = ac2fJBTab(row, 2)
    jr.MatThick = ac2fJBTab(row, 3)
    jr.MatBrand = ac2fJBTab(row, 4)
    jr.MatGrade = ac2fJBTab(row, 5)
    jr.MatColor = ac2fJBTab(row, 6)
    jr.MatW = Val(ac2fJBTab(row, 7))
    jr.MatH = Val(ac2fJBTab(row, 8))
    jr.MatLabel = ac2fJBTab(row, 10)
    If Len(row) = 0 Then Exit Sub

    ' a customer's own sheet belongs to that customer, not to the list
    If jr.MatOwner <> "C" Then ac2fJBRecentPush AC2F_K_JB_RMAT, row
    If ac2fGetLng(AC2F_K_JB_NEST, AC2F_DEF_JB_NEST) <> 0 And jr.MatW > 0 And jr.MatH > 0 Then
        ac2fSetNum AC2F_K_NS_SW, jr.MatW
        ac2fSetNum AC2F_K_NS_SH, jr.MatH
    End If
End Sub

' A file opened from <day>\<customer>\ offers that customer's name.
Private Function ac2fJBFolderHint() As String
    Dim fp As String
    Dim parts() As String
    Dim n As Long

    On Error Resume Next
    fp = ActiveDocument.FilePath
    If Len(fp) = 0 Then Exit Function
    fp = ac2fJBNoSlash(fp)
    parts = Split(fp, "\")
    n = UBound(parts)
    If n < 1 Then Exit Function
    If parts(n - 1) Like "##[-.]##[-.]####*" Then ac2fJBFolderHint = parts(n)
End Function

'=====================================================================
' TOOLPATH FOLDERS
'=====================================================================

' Works out a plan from the drawing, written the way the user would
' type it ("2", or "2, 1 2050x3050"), and says where it came from. Every
' sheet counts whole: the part a job used is marked from the quote.
Private Function ac2fJBPropose(ByRef jr As ac2fJobRec, ByRef why As String) As String
    Dim sx0() As Double, sy0() As Double, sx1() As Double, sy1() As Double
    Dim px0() As Double, py0() As Double, pw() As Double, ph() As Double
    Dim ns As Long, np As Long
    Dim toks() As String, cnts() As Long, nt As Long
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim i As Long, j As Long, k As Long
    Dim x As Double, y As Double, w As Double, h As Double
    Dim cx As Double, cy As Double, sw As Double, sh As Double
    Dim filled As Boolean
    Dim s As String, own As String

    why = "Type how many sheets, and the size if it is not the job's sheet."
    ReDim sx0(1 To 8): ReDim sy0(1 To 8): ReDim sx1(1 To 8): ReDim sy1(1 To 8)
    ReDim toks(1 To 8): ReDim cnts(1 To 8)

    On Error GoTo Fail
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    ' --- Nesting's sheet outlines: in the selection, else on the page ---
    Set sr = ActiveSelectionRange
    If Not sr Is Nothing Then ac2fJBScanSheets sr, sx0, sy0, sx1, sy1, ns
    If ns = 0 Then ac2fJBScanSheets ActivePage.Shapes, sx0, sy0, sx1, sy1, ns

    If ns > 0 Then
        ' an outline counts when something sits on it
        np = ac2fJBScanParts(px0, py0, pw, ph)
        For j = 1 To ns
            sw = sx1(j) - sx0(j)
            sh = sy1(j) - sy0(j)
            filled = False
            For i = 1 To np
                cx = px0(i) + pw(i) / 2#
                cy = py0(i) + ph(i) / 2#
                If cx >= sx0(j) And cx <= sx1(j) And cy >= sy0(j) And cy <= sy1(j) _
                   And pw(i) <= sw + 1# And ph(i) <= sh + 1# Then
                    filled = True
                    Exit For
                End If
            Next i
            If filled Then
                ' the sheet as stocked when the outline matches it
                If Abs(sw - jr.MatW) <= 2# And Abs(sh - jr.MatH) <= 2# Then
                    sw = jr.MatW
                    sh = jr.MatH
                End If
                k = ac2fJBTally(toks, cnts, nt, ac2fJBSizeTok(sw, sh), 1)
            End If
        Next j
        If nt > 0 Then why = "From the " & ns & " nest sheet outline(s) on the page:"

    ' --- otherwise one sheet for the selection ---
    ElseIf jr.MatW > 0 And jr.MatH > 0 And Not sr Is Nothing Then
        If sr.Count > 0 Then
            sr.GetBoundingBox x, y, w, h
            s = "Selection " & Format$(w, "0") & " x " & Format$(h, "0") & " mm"
            If (w <= jr.MatW + 1# And h <= jr.MatH + 1#) Or _
               (h <= jr.MatW + 1# And w <= jr.MatH + 1#) Then
                why = s & ", fits on one sheet:"
            Else
                why = s & " is bigger than one sheet; how many?"
            End If
            k = ac2fJBTally(toks, cnts, nt, ac2fJBSizeTok(jr.MatW, jr.MatH), 1)
        End If
    End If

Done:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    If jr.MatW > 0 And jr.MatH > 0 Then own = ac2fJBSizeTok(jr.MatW, jr.MatH)
    If nt = 0 And Len(own) > 0 Then k = ac2fJBTally(toks, cnts, nt, own, 1)
    s = ""
    For i = 1 To nt
        If i > 1 Then s = s & ", "
        ' the job's own sheet goes without its size
        If toks(i) = own Then
            s = s & cnts(i)
        Else
            s = s & cnts(i) & " " & toks(i)
        End If
    Next i
    ac2fJBPropose = s
    Exit Function

Fail:
    Resume Done
End Function

' Collects Nesting's sheet outlines ("ac2f sheet N", or the group of
' them, "ac2f sheets") from a ShapeRange or a Shapes collection.
Private Sub ac2fJBScanSheets(ByVal coll As Object, _
                             ByRef sx0() As Double, ByRef sy0() As Double, _
                             ByRef sx1() As Double, ByRef sy1() As Double, _
                             ByRef ns As Long)
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
            ac2fJBScanSheets s.Shapes, sx0, sy0, sx1, sy1, ns
        ElseIf nm Like "ac2f sheet #*" Then
            s.GetBoundingBox x, y, w, h
            If w > 0 And h > 0 Then
                ns = ns + 1
                If ns > UBound(sx0) Then
                    ReDim Preserve sx0(1 To ns * 2): ReDim Preserve sy0(1 To ns * 2)
                    ReDim Preserve sx1(1 To ns * 2): ReDim Preserve sy1(1 To ns * 2)
                End If
                sx0(ns) = x: sy0(ns) = y
                sx1(ns) = x + w: sy1(ns) = y + h
            End If
        End If
    Next i
End Sub

' Bounding boxes of everything on the page except the sheet outlines.
Private Function ac2fJBScanParts(ByRef px0() As Double, ByRef py0() As Double, _
                                 ByRef pw() As Double, ByRef ph() As Double) As Long
    Dim s As Shape
    Dim i As Long, n As Long
    Dim nm As String
    Dim x As Double, y As Double, w As Double, h As Double

    ReDim px0(1 To 64): ReDim py0(1 To 64): ReDim pw(1 To 64): ReDim ph(1 To 64)
    ac2fJBScanPartsIn ActivePage.Shapes, px0, py0, pw, ph, n
    ac2fJBScanParts = n
End Function

' Nesting's "everything" group spans several sheets, so its parts are
' looked at one by one.
Private Sub ac2fJBScanPartsIn(ByVal coll As Object, ByRef px0() As Double, ByRef py0() As Double, _
                              ByRef pw() As Double, ByRef ph() As Double, ByRef n As Long)
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
        If nm = "ac2f nesting" Then
            ac2fJBScanPartsIn s.Shapes, px0, py0, pw, ph, n
        ElseIf nm <> "ac2f sheets" And Not (nm Like "ac2f sheet #*") And Not s Is Nothing Then
            w = 0: h = 0
            s.GetBoundingBox x, y, w, h
            If w > 0 Or h > 0 Then
                n = n + 1
                If n > UBound(px0) Then
                    ReDim Preserve px0(1 To n * 2): ReDim Preserve py0(1 To n * 2)
                    ReDim Preserve pw(1 To n * 2): ReDim Preserve ph(1 To n * 2)
                End If
                px0(n) = x: py0(n) = y: pw(n) = w: ph(n) = h
            End If
        End If
    Next i
End Sub

Private Function ac2fJBSizeTok(ByVal w As Double, ByVal h As Double) As String
    ac2fJBSizeTok = CStr(CLng(w)) & "x" & CStr(CLng(h))
End Function

' "1220x2440" (or 1220*2440) -> width and height.
Private Function ac2fJBParseSize(ByVal t As String, ByRef w As Double, _
                                 ByRef h As Double) As Boolean
    Dim p As Long

    t = Replace$(LCase$(Trim$(t)), "*", "x")
    p = InStr(t, "x")
    If p <= 1 Or p >= Len(t) Then Exit Function
    w = ac2fParseNum(Left$(t, p - 1), 0)
    h = ac2fParseNum(Mid$(t, p + 1), 0)
    If w <= 0 Or h <= 0 Then Exit Function
    ac2fJBParseSize = True
End Function

' A size typed against the job's sheet, turned round if it only fits
' that way ("2440x1220" for a 1220x2440 sheet is the same sheet).
Private Function ac2fJBNormTok(ByRef jr As ac2fJobRec, ByVal w As Double, _
                               ByVal h As Double) As String
    Dim t As Double

    If jr.MatW > 0 And jr.MatH > 0 Then
        If (w > jr.MatW + 1# Or h > jr.MatH + 1#) And _
           h <= jr.MatW + 1# And w <= jr.MatH + 1# Then
            t = w: w = h: h = t
        End If
    End If
    ac2fJBNormTok = ac2fJBSizeTok(w, h)
End Function

' Adds cnt of tok to the plan, merging identical sizes. Returns its row.
Private Function ac2fJBTally(ByRef toks() As String, ByRef cnts() As Long, _
                             ByRef nt As Long, ByVal tok As String, _
                             ByVal cnt As Long) As Long
    Dim i As Long

    For i = 1 To nt
        If toks(i) = tok Then
            cnts(i) = cnts(i) + cnt
            ac2fJBTally = i
            Exit Function
        End If
    Next i
    nt = nt + 1
    If nt > UBound(toks) Then
        ReDim Preserve toks(1 To nt * 2)
        ReDim Preserve cnts(1 To nt * 2)
    End If
    toks(nt) = tok
    cnts(nt) = cnt
    ac2fJBTally = nt
End Function

' Reads "2, 1 2050x3050". A count alone means the job's own sheet.
' Returns the number of sizes, or -1 with errText.
Private Function ac2fJBParsePlan(ByRef jr As ac2fJobRec, ByVal txt As String, _
                                 ByRef toks() As String, ByRef cnts() As Long, _
                                 ByRef errText As String) As Long
    Dim entries() As String, words() As String
    Dim i As Long, j As Long, k As Long, n As Long
    Dim cnt As Long, tok As String
    Dim w As Double, h As Double

    ReDim toks(1 To 8): ReDim cnts(1 To 8)
    entries = Split(Replace$(txt, "+", ","), ",")
    For i = LBound(entries) To UBound(entries)
        If Len(Trim$(entries(i))) > 0 Then
            cnt = 1
            tok = ""
            words = Split(Trim$(entries(i)), " ")
            For j = LBound(words) To UBound(words)
                If Len(words(j)) = 0 Then
                    ' double space
                ElseIf ac2fJBIsCount(words(j)) Then
                    cnt = CLng(words(j))
                ElseIf ac2fJBParseSize(words(j), w, h) Then
                    tok = ac2fJBNormTok(jr, w, h)
                Else
                    errText = "Cannot read """ & words(j) & """." & vbCrLf & _
                              "Write a count, a size, or both: 2 2050x3050"
                    ac2fJBParsePlan = -1
                    Exit Function
                End If
            Next j
            If cnt < 1 Or cnt > 99 Then
                errText = "A count runs from 1 to 99."
                ac2fJBParsePlan = -1
                Exit Function
            End If
            If Len(tok) = 0 Then
                If jr.MatW <= 0 Or jr.MatH <= 0 Then
                    errText = "This sheet has no size on record; type it, e.g. 1 1220x2440"
                    ac2fJBParsePlan = -1
                    Exit Function
                End If
                tok = ac2fJBSizeTok(jr.MatW, jr.MatH)
            End If
            k = ac2fJBTally(toks, cnts, n, tok, cnt)
        End If
    Next i
    ac2fJBParsePlan = n
End Function

' Shows what will be created or stacked, asks, then does it.
Private Sub ac2fJBMakeFolders(ByRef jr As ac2fJobRec, ByVal custDir As String, _
                              ByRef toks() As String, ByRef cnts() As Long, _
                              ByVal n As Long)
    Dim found() As String, oldCnt() As Long, grp() As Long, splits() As Long
    Dim tail As String, head As String, nm As String, leaf As String
    Dim m As String, rep As String, errText As String
    Dim i As Long, k As Long, q As Long, first As Long
    Dim maxNo As Long, nextNo As Long, c As Long
    Dim w As Double, h As Double, lim As Double

    ReDim found(1 To n): ReDim oldCnt(1 To n): ReDim grp(1 To n): ReDim splits(1 To n)
    tail = ac2fJBSpec(jr)
    lim = ac2fGetNum(AC2F_K_JB_LONG, AC2F_DEF_JB_LONG)

    ' --- dry run: what is there already ---
    For i = 1 To n
        head = jr.MatKind & "-" & toks(i)
        found(i) = ac2fJBFindGroup(custDir, head, tail, c, maxNo)
        oldCnt(i) = c
        If ac2fJBParseSize(toks(i), w, h) Then
            If lim > 0 And h > lim Then
                splits(i) = CLng(ac2fCeil(h / lim))
                If splits(i) < 2 Then splits(i) = 2
            End If
        End If
    Next i
    nextNo = maxNo
    For i = 1 To n
        If Len(found(i)) > 0 Then
            grp(i) = CLng(Left$(found(i), 3))
        Else
            nextNo = nextNo + 1
            grp(i) = nextNo
        End If
    Next i

    m = "In  " & custDir & vbCrLf & vbCrLf
    For i = 1 To n
        head = jr.MatKind & "-" & toks(i)
        If Len(found(i)) > 0 Then
            m = m & "STACK  " & found(i) & vbCrLf & "    -> " & (oldCnt(i) + cnts(i)) & _
                "pcs, adds " & ac2fJBRange(oldCnt(i) + 1, oldCnt(i) + cnts(i)) & vbCrLf
        Else
            m = m & "NEW    " & ac2fJBGroupName(grp(i), head, cnts(i), tail) & vbCrLf & _
                "    -> " & ac2fJBRange(1, cnts(i)) & vbCrLf
        End If
        If splits(i) > 0 Then
            m = m & "    with " & ac2fJBRange(1, splits(i), "p") & " in each (over " & _
                Format$(lim, "0") & " mm)" & vbCrLf
        End If
    Next i
    rep = ac2fJBRunnerNote(custDir)
    If Len(rep) > 0 Then
        m = m & vbCrLf & "StockTrack Runner already made the folders here when these" & vbCrLf & _
            "quotes became jobs: " & rep & vbCrLf & _
            "If this drawing is one of them, Yes stacks its sheets a second time." & vbCrLf
        rep = ""
    End If
    If Len(m) > PROMPT_MAX Then m = Left$(m, PROMPT_MAX) & " ..."
    If MsgBox(m & vbCrLf & "Create them?", vbQuestion + vbYesNo, _
              ac2fTitle(CAPTION_)) <> vbYes Then Exit Sub

    ' --- do it ---
    If Not ac2fJBMakeDirs(custDir) Then
        ac2fWarn "Could not create " & custDir, CAPTION_
        Exit Sub
    End If
    For i = 1 To n
        head = jr.MatKind & "-" & toks(i)
        If Len(found(i)) > 0 Then
            nm = ac2fJBGroupName(grp(i), head, oldCnt(i) + cnts(i), tail)
            first = oldCnt(i) + 1
            If StrComp(nm, found(i), vbBinaryCompare) <> 0 Then
                On Error Resume Next
                Err.Clear
                Name custDir & "\" & found(i) As custDir & "\" & nm
                If Err.Number <> 0 Then
                    errText = errText & found(i) & vbCrLf & "   could not be renamed: " & _
                              Err.Description & " (open in Explorer or the CAM program?)" & vbCrLf
                    nm = ""
                End If
                On Error GoTo 0
            End If
        Else
            nm = ac2fJBGroupName(grp(i), head, cnts(i), tail)
            first = 1
            If Not ac2fJBMakeDirs(custDir & "\" & nm) Then
                errText = errText & "Could not create " & nm & vbCrLf
                nm = ""
            End If
        End If

        If Len(nm) > 0 Then
            For k = first To first + cnts(i) - 1
                leaf = custDir & "\" & nm & "\" & Format$(k, "000")
                If Not ac2fJBMakeDirs(leaf) Then errText = errText & "Could not create " & leaf & vbCrLf
                For q = 1 To splits(i)
                    If Not ac2fJBMakeDirs(leaf & "\p" & Format$(q, "000")) Then
                        errText = errText & "Could not create " & leaf & "\p" & Format$(q, "000") & vbCrLf
                    End If
                Next q
            Next k
            rep = rep & nm & vbCrLf
        End If
    Next i

    If ac2fGetLng(AC2F_K_JB_OPEN, AC2F_DEF_JB_OPEN) <> 0 Then ac2fJBExplore custDir
    If Len(errText) > 0 Then
        ac2fWarn "Done, with problems:" & vbCrLf & vbCrLf & errText & vbCrLf & _
                 "Ready:" & vbCrLf & rep, CAPTION_
    Else
        ac2fInfo "Ready in " & custDir & ":" & vbCrLf & vbCrLf & rep, CAPTION_
    End If
End Sub

' Quotes StockTrack Runner has opened folders for in this customer folder
' (it leaves a .stocktrack\<quote>.done file, first line "Teklif TKF-...,"),
' as "TKF-2026-0003, TKF-2026-0004". Empty when there are none.
Private Function ac2fJBRunnerNote(ByVal custDir As String) As String
    Dim names() As String
    Dim i As Long, n As Long, p As Long
    Dim txt As String, q As String, out As String

    n = ac2fJBListNames(custDir & "\.stocktrack", "*.done", False, names)
    For i = 1 To n
        txt = ac2fJBReadUtf8(custDir & "\.stocktrack\" & names(i))
        q = ""
        If Left$(txt, 7) = "Teklif " Then
            p = InStr(8, txt, ",")
            If p > 8 Then q = Mid$(txt, 8, p - 8)
        End If
        If Len(q) = 0 Then q = "?"
        out = ac2fJBJoin(out, q, ", ")
    Next i
    ac2fJBRunnerNote = out
End Function

' "001", or "001-003"; prefix "p" gives "p001-p002".
Private Function ac2fJBRange(ByVal a As Long, ByVal b As Long, _
                             Optional ByVal pre As String = "") As String
    If b <= a Then
        ac2fJBRange = pre & Format$(a, "000")
    Else
        ac2fJBRange = pre & Format$(a, "000") & "-" & pre & Format$(b, "000")
    End If
End Function

' 001-pleksi-1220x2440--2pcs--2800um-akpolimer-dokme-beyaz
Private Function ac2fJBGroupName(ByVal grpNo As Long, ByVal head As String, _
                                 ByVal cnt As Long, ByVal tail As String) As String
    Dim s As String
    s = Format$(grpNo, "000") & "-" & head & ac2fJBSep() & CStr(cnt) & "pcs"
    If Len(tail) > 0 Then s = s & ac2fJBSep() & tail
    ac2fJBGroupName = s
End Function

' Looks for a group folder with this head and tail. Returns its name
' (empty when there is none) and its count; maxNo is the highest group
' number already in use.
Private Function ac2fJBFindGroup(ByVal custDir As String, ByVal head As String, _
                                 ByVal tail As String, ByRef cnt As Long, _
                                 ByRef maxNo As Long) As String
    Dim names() As String
    Dim i As Long, n As Long, grpNo As Long, c As Long
    Dim h As String, t As String, hit As String

    cnt = 0
    maxNo = 0
    n = ac2fJBListNames(custDir, "*", True, names)
    For i = 1 To n
        If ac2fJBSplitGroup(names(i), grpNo, h, c, t) Then
            If grpNo > maxNo Then maxNo = grpNo
            If Len(hit) = 0 And StrComp(h, head, vbTextCompare) = 0 _
               And StrComp(t, tail, vbTextCompare) = 0 Then
                hit = names(i)
                cnt = c
            End If
        End If
    Next i
    ac2fJBFindGroup = hit
End Function

' "001-pleksi-1220x2440--2pcs--2800um-..." -> 1, "pleksi-1220x2440", 2,
' "2800um-...". False for anything that is not a group folder.
Private Function ac2fJBSplitGroup(ByVal nm As String, ByRef grpNo As Long, _
                                  ByRef head As String, ByRef cnt As Long, _
                                  ByRef tail As String) As Boolean
    Dim rest As String, sep As String
    Dim p As Long, q As Long, d As Long, after As Long

    If Len(nm) < 6 Then Exit Function
    If Not (Left$(nm, 3) Like "###") Or Mid$(nm, 4, 1) <> "-" Then Exit Function
    rest = Mid$(nm, 5)
    sep = ac2fJBSep()

    p = InStr(1, rest, sep)
    Do While p > 1
        q = p + Len(sep)
        d = q
        Do While Mid$(rest, d, 1) Like "#"
            d = d + 1
        Loop
        If d > q And LCase$(Mid$(rest, d, 3)) = "pcs" Then
            after = d + 3
            If after > Len(rest) Then
                tail = ""
            ElseIf Mid$(rest, after, Len(sep)) = sep Then
                tail = Mid$(rest, after + Len(sep))
            Else
                after = 0
            End If
            If after > 0 Then
                grpNo = CLng(Left$(nm, 3))
                head = Left$(rest, p - 1)
                cnt = CLng(Mid$(rest, q, d - q))
                ac2fJBSplitGroup = True
                Exit Function
            End If
        End If
        p = InStr(p + 1, rest, sep)
    Loop
End Function

' Next free export number in the folder: "01", "02" ...
Private Function ac2fJBNextExport(ByVal custDir As String) As String
    Dim names() As String
    Dim i As Long, n As Long, d As Long, v As Long, best As Long

    n = ac2fJBListNames(custDir, "exp*", False, names)
    For i = 1 To n
        d = 4
        Do While Mid$(names(i), d, 1) Like "#"
            d = d + 1
        Loop
        If d > 4 And d - 4 <= 6 Then
            v = CLng(Mid$(names(i), 4, d - 4))
            If v > best Then best = v
        End If
    Next i
    ac2fJBNextExport = Format$(best + 1, "00")
End Function

'=====================================================================
' USED PARTS
'=====================================================================

' Finds each quote line's toolpath folder, shows what will be written,
' asks, then writes the u-files.
Private Sub ac2fJBMarkUsage(ByRef jr As ac2fJobRec, ByVal qrow As String, _
                            ByRef irows() As String, ByVal ni As Long)
    Dim it As ac2fJobRec
    Dim paths() As String, bodies() As String
    Dim nf As Long, i As Long, j As Long
    Dim quoteNo As String, qdate As String, lbl As String
    Dim grp As String, target As String, body As String
    Dim todo As String, whole As String, missing As String
    Dim made As String, errText As String, m As String
    Dim uw As Double, uh As Double, gw As Double, gh As Double
    Dim d0 As Date

    quoteNo = ac2fJBTab(qrow, 1)
    qdate = ac2fJBTab(qrow, 2)
    d0 = ac2fJBIsoDate(qdate)
    ReDim paths(1 To ni): ReDim bodies(1 To ni)

    For i = 1 To ni
        ac2fJBSetFields it, irows(i)
        uw = it.MatW
        uh = it.MatH
        lbl = ac2fJBClip(ac2fJBTab(irows(i), 10), 60)
        gw = 0: gh = 0
        grp = ac2fJBFindSheetGroup(jr.CustFolder, jr.JobDay, d0, it.MatKind, ac2fJBSpec(it), _
                                   uw, uh, gw, gh)
        If Len(grp) = 0 Then
            missing = missing & "   " & lbl & vbCrLf
        ElseIf uw <= 0 Or uh <= 0 Or ac2fJBSameSize(uw, uh, gw, gh) Then
            whole = whole & "   " & lbl & vbCrLf
        Else
            target = grp & "\" & ac2fJBUseName(uw, uh, gw, gh)
            body = "Quote " & quoteNo & "  " & qdate & vbCrLf & _
                   ac2fJBTab(irows(i), 10) & vbCrLf & _
                   "Used " & Format$(uw, "0") & " x " & Format$(uh, "0") & " mm of the " & _
                   Format$(gw, "0") & " x " & Format$(gh, "0") & " mm sheet, qty " & _
                   ac2fJBTab(irows(i), 9) & vbCrLf
            ' two lines landing on the same file share it
            For j = 1 To nf
                If StrComp(paths(j), target, vbTextCompare) = 0 Then Exit For
            Next j
            If j <= nf Then
                bodies(j) = bodies(j) & body
            Else
                nf = nf + 1
                paths(nf) = target
                bodies(nf) = body
                todo = todo & "   " & ac2fJBTail2(target) & vbCrLf
            End If
        End If
    Next i

    m = "QUOTE " & quoteNo & "  -  " & jr.CustFolder & vbCrLf & vbCrLf
    If nf > 0 Then m = m & "Will write:" & vbCrLf & todo & vbCrLf
    If Len(whole) > 0 Then m = m & "Whole sheet, nothing to mark:" & vbCrLf & whole & vbCrLf
    If Len(missing) > 0 Then
        m = m & "No toolpath folder found for:" & vbCrLf & missing & _
            "(searched the customer's folders up to two weeks back)" & vbCrLf & vbCrLf
    End If
    If Len(m) > PROMPT_MAX Then m = Left$(m, PROMPT_MAX) & " ..." & vbCrLf & vbCrLf
    If nf = 0 Then
        ac2fInfo m & "Nothing to write.", CAPTION_
        Exit Sub
    End If
    If MsgBox(m & "Write them?", vbQuestion + vbYesNo, ac2fTitle(CAPTION_)) <> vbYes Then Exit Sub

    For i = 1 To nf
        If ac2fJBFileExists(paths(i)) Then
            made = made & "   " & ac2fJBTail2(paths(i)) & "  (already there)" & vbCrLf
        ElseIf ac2fJBWriteText(paths(i), bodies(i)) Then
            made = made & "   " & ac2fJBTail2(paths(i)) & vbCrLf
        Else
            errText = errText & "   " & paths(i) & vbCrLf
        End If
    Next i

    If Len(errText) > 0 Then
        ac2fWarn "Could not write:" & vbCrLf & errText & vbCrLf & "Written:" & vbCrLf & made, CAPTION_
    Else
        ac2fInfo "Written:" & vbCrLf & made, CAPTION_
    End If
End Sub

' The toolpath folder a quote line's sheet went into: same kind and the
' same thickness, brand, variant and colour, in this customer's folder.
' The job's own day folder is tried first. Otherwise, a quote being made
' after the cutting, the day folders are searched from the day after the
' quote back two weeks, then today. Among several sizes, the smallest
' sheet the used part fits on. gw, gh return its size.
Private Function ac2fJBFindSheetGroup(ByVal custFolder As String, ByVal jobDay As String, _
                                      ByVal d0 As Date, ByVal kind As String, _
                                      ByVal tail As String, ByVal uw As Double, _
                                      ByVal uh As Double, ByRef gw As Double, _
                                      ByRef gh As Double) As String
    Dim root As String, fmt As String, dsep As String, best As String
    Dim days() As String
    Dim nd As Long, i As Long, back As Long
    Dim bestArea As Double
    Dim d As Date

    If Len(kind) = 0 Then Exit Function
    If Not ac2fJBDayRule(root, fmt, dsep) Then Exit Function

    If Len(jobDay) > 0 Then
        ac2fJBBestGroup ac2fJBCustPath(root & "\" & jobDay, custFolder), kind, tail, uw, uh, _
                        best, bestArea, gw, gh
    End If
    If Len(best) > 0 Then
        ac2fJBFindSheetGroup = best
        Exit Function
    End If

    For back = -1 To 15
        If back = 15 Then d = Date Else d = d0 - back
        nd = ac2fJBListNames(root, ac2fJBStrftime(fmt, d) & "*", True, days)
        For i = 1 To nd
            ac2fJBBestGroup ac2fJBCustPath(root & "\" & days(i), custFolder), kind, tail, uw, uh, _
                            best, bestArea, gw, gh
        Next i
        If Len(best) > 0 Then Exit For
    Next back
    ac2fJBFindSheetGroup = best
End Function

' Looks in one customer folder for a group of this kind and tail that the
' used part fits on, keeping the smallest sheet found so far in best.
Private Sub ac2fJBBestGroup(ByVal cdir As String, ByVal kind As String, _
                            ByVal tail As String, ByVal uw As Double, ByVal uh As Double, _
                            ByRef best As String, ByRef bestArea As Double, _
                            ByRef gw As Double, ByRef gh As Double)
    Dim grps() As String
    Dim ng As Long, j As Long, grpNo As Long, c As Long
    Dim h As String, t As String
    Dim w As Double, ht As Double

    If Not ac2fJBDirExists(cdir) Then Exit Sub
    ng = ac2fJBListNames(cdir, "*", True, grps)
    For j = 1 To ng
        If ac2fJBSplitGroup(grps(j), grpNo, h, c, t) Then
            If StrComp(t, tail, vbTextCompare) = 0 And _
               StrComp(Left$(h, Len(kind) + 1), kind & "-", vbTextCompare) = 0 Then
                If ac2fJBParseSize(Mid$(h, Len(kind) + 2), w, ht) Then
                    If (uw <= w + 1# And uh <= ht + 1#) Or _
                       (uh <= w + 1# And uw <= ht + 1#) Then
                        If Len(best) = 0 Or w * ht < bestArea Then
                            best = cdir & "\" & grps(j)
                            bestArea = w * ht
                            gw = w
                            gh = ht
                        End If
                    End If
                End If
            End If
        End If
    Next j
End Sub

' True when the part used is the whole sheet, either way round.
Private Function ac2fJBSameSize(ByVal uw As Double, ByVal uh As Double, _
                                ByVal gw As Double, ByVal gh As Double) As Boolean
    ac2fJBSameSize = (Abs(uw - gw) <= 1# And Abs(uh - gh) <= 1#) Or _
                     (Abs(uw - gh) <= 1# And Abs(uh - gw) <= 1#)
End Function

' u + the part used, the side along the sheet's width first: 1000 x 1350
' used of a 1350 x 2000 sheet is u1350x1000.
Private Function ac2fJBUseName(ByVal uw As Double, ByVal uh As Double, _
                               ByVal gw As Double, ByVal gh As Double) As String
    Dim a As Double, b As Double

    If Abs(uw - gw) <= 1# Then
        a = uw: b = uh
    ElseIf Abs(uh - gw) <= 1# Then
        a = uh: b = uw
    ElseIf uw <= gw + 1# And uh <= gh + 1# Then
        a = uw: b = uh
    Else
        a = uh: b = uw
    End If
    ac2fJBUseName = "u" & CStr(CLng(a)) & "x" & CStr(CLng(b))
End Function

' "2026-10-17" -> that date; anything else is today.
Private Function ac2fJBIsoDate(ByVal s As String) As Date
    On Error GoTo Bad
    If Len(s) >= 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        ac2fJBIsoDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
        Exit Function
    End If
Bad:
    ac2fJBIsoDate = Date
End Function

' The last two parts of a path: "001-pleksi-...\u1350x1000".
Private Function ac2fJBTail2(ByVal p As String) As String
    Dim a As Long, b As Long
    a = InStrRev(p, "\")
    If a > 1 Then b = InStrRev(p, "\", a - 1)
    ac2fJBTail2 = Mid$(p, b + 1)
End Function

Private Function ac2fJBWriteText(ByVal path As String, ByVal txt As String) As Boolean
    Dim f As Integer
    On Error GoTo Fail
    f = FreeFile
    Open path For Output As #f
    Print #f, txt;
    Close #f
    ac2fJBWriteText = True
    Exit Function
Fail:
    On Error Resume Next
    Close #f
End Function

'=====================================================================
' NAMES
'=====================================================================

' "2800um-akpolimer-dokme-beyaz": everything after the count.
Private Function ac2fJBSpec(ByRef jr As ac2fJobRec) As String
    Dim s As String
    If Len(jr.MatThick) > 0 Then s = jr.MatThick & "um"
    s = ac2fJBJoin(s, jr.MatBrand, "-")
    s = ac2fJBJoin(s, jr.MatGrade, "-")
    s = ac2fJBJoin(s, jr.MatColor, "-")
    ac2fJBSpec = s
End Function

' "pleksi 2.8mm akpolimer dokme beyaz 1220x2440", for messages.
Private Function ac2fJBMatText(ByRef jr As ac2fJobRec) As String
    Dim s As String
    If Len(jr.MatKind) = 0 Then
        ac2fJBMatText = "<none>"
        Exit Function
    End If
    s = jr.MatKind
    If Val(jr.MatThick) > 0 Then s = s & " " & ac2fNumStr(Val(jr.MatThick) / 1000#) & "mm"
    s = ac2fJBJoin(s, jr.MatBrand, " ")
    s = ac2fJBJoin(s, jr.MatGrade, " ")
    s = ac2fJBJoin(s, jr.MatColor, " ")
    If jr.MatW > 0 And jr.MatH > 0 Then s = s & " " & ac2fJBSizeTok(jr.MatW, jr.MatH)
    ac2fJBMatText = s
End Function

' Same text for a material row (typed or remembered).
Private Function ac2fJBRowText(ByVal row As String) As String
    Dim jr As ac2fJobRec
    ac2fJBSetFields jr, row
    ac2fJBRowText = ac2fJBMatText(jr)
End Function

Private Sub ac2fJBSetFields(ByRef jr As ac2fJobRec, ByVal row As String)
    jr.MatKind = ac2fJBTab(row, 2)
    jr.MatThick = ac2fJBTab(row, 3)
    jr.MatBrand = ac2fJBTab(row, 4)
    jr.MatGrade = ac2fJBTab(row, 5)
    jr.MatColor = ac2fJBTab(row, 6)
    jr.MatW = Val(ac2fJBTab(row, 7))
    jr.MatH = Val(ac2fJBTab(row, 8))
End Sub

' Folds Turkish and accented letters to ASCII and drops the rest.
' Character codes are used so that this source stays pure ASCII.
Private Function ac2fJBFold(ByVal s As String) As String
    Dim i As Long, w As Long
    Dim c As String, out As String

    For i = 1 To Len(s)
        w = AscW(Mid$(s, i, 1))
        If w < 0 Then w = w + 65536
        Select Case w
            Case &HE7: c = "c"                          ' c cedilla
            Case &HC7: c = "C"
            Case &H11F: c = "g"                         ' g breve
            Case &H11E: c = "G"
            Case &H131: c = "i"                         ' dotless i
            Case &H130: c = "I"                         ' I with dot
            Case &HF6, &HF2, &HF3, &HF4: c = "o"
            Case &HD6, &HD2, &HD3, &HD4: c = "O"
            Case &H15F: c = "s"                         ' s cedilla
            Case &H15E: c = "S"
            Case &HFC, &HF9, &HFA, &HFB: c = "u"
            Case &HDC, &HD9, &HDA, &HDB: c = "U"
            Case &HE0, &HE1, &HE2, &HE4: c = "a"
            Case &HC0, &HC1, &HC2, &HC4: c = "A"
            Case &HE8, &HE9, &HEA, &HEB: c = "e"
            Case &HC8, &HC9, &HCA, &HCB: c = "E"
            Case &HEC, &HED, &HEE, &HEF: c = "i"
            Case &HCC, &HCD, &HCE, &HCF: c = "I"
            Case 9, 10, 13, &HA0: c = " "
            Case 32 To 126: c = ChrW$(w)
            Case Else: c = ""
        End Select
        out = out & c
    Next i
    ac2fJBFold = out
End Function

' Folder name, the same rule StockTrack applies: lower case ASCII, every
' run of other characters one "-", so no space or capital reaches a path
' ("Ozturk Reklam Ltd. Sti." -> "ozturk-reklam-ltd-sti").
Private Function ac2fJBSafeName(ByVal s As String) As String
    Dim out As String

    out = ac2fJBToken(s)
    If Len(out) > 60 Then out = Left$(out, 60)
    Do While Right$(out, 1) = "-"
        out = Left$(out, Len(out) - 1)
    Loop
    Select Case out
        Case "con", "prn", "aux", "nul", "com1", "com2", "com3", "com4", "com5", _
             "com6", "com7", "com8", "com9", "lpt1", "lpt2", "lpt3", "lpt4", _
             "lpt5", "lpt6", "lpt7", "lpt8", "lpt9"
            out = out & "_"
    End Select
    ac2fJBSafeName = out
End Function

' A separator: printable ASCII, no space, nothing Windows refuses in a name.
Private Function ac2fJBSepOk(ByVal s As String) As Boolean
    Dim i As Long
    Dim c As String

    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If AscW(c) < 33 Or AscW(c) > 126 Then Exit Function
        If InStr("\/:*?""<>|", c) > 0 Then Exit Function
    Next i
    ac2fJBSepOk = True
End Function

' The folder this customer really has in dayPath: the name as it is now
' or, for folders made before names went lower case with dashes, the one
' that reads the same ("Ozturk Reklam" for "ozturk-reklam"). Empty when
' there is none.
Private Function ac2fJBCustSub(ByVal dayPath As String, ByVal custFolder As String) As String
    Dim names() As String
    Dim i As Long, n As Long
    Dim want As String

    If Len(custFolder) = 0 Then Exit Function
    If ac2fJBDirExists(dayPath & "\" & custFolder) Then
        ac2fJBCustSub = custFolder
        Exit Function
    End If
    want = ac2fJBToken(custFolder)
    If Len(want) = 0 Then Exit Function
    n = ac2fJBListNames(dayPath, "*", True, names)
    For i = 1 To n
        If ac2fJBToken(names(i)) = want Then
            ac2fJBCustSub = names(i)
            Exit Function
        End If
    Next i
End Function

' dayPath\<customer folder>, the existing one when there is one.
Private Function ac2fJBCustPath(ByVal dayPath As String, ByVal custFolder As String) As String
    ac2fJBCustPath = dayPath & "\" & ac2fJBOr(ac2fJBCustSub(dayPath, custFolder), custFolder)
End Function

' One name piece: lower case ASCII, every run of other characters one
' "-", so a piece never holds the "--" separator.
Private Function ac2fJBToken(ByVal s As String) As String
    Dim i As Long
    Dim c As String, out As String
    Dim gap As Boolean

    s = LCase$(ac2fJBFold(s))
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Then
            If gap And Len(out) > 0 Then out = out & "-"
            gap = False
            out = out & c
        Else
            gap = True
        End If
    Next i
    ac2fJBToken = out
End Function

Private Function ac2fJBSep() As String
    ac2fJBSep = ac2fGetStr(AC2F_K_JB_SEP, AC2F_DEF_JB_SEP)
End Function

Private Function ac2fJBJoin(ByVal a As String, ByVal b As String, ByVal sep As String) As String
    If Len(a) = 0 Then
        ac2fJBJoin = b
    ElseIf Len(b) = 0 Then
        ac2fJBJoin = a
    Else
        ac2fJBJoin = a & sep & b
    End If
End Function

Private Function ac2fJBOr(ByVal a As String, ByVal b As String) As String
    If Len(a) > 0 Then ac2fJBOr = a Else ac2fJBOr = b
End Function

Private Function ac2fJBClip(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then ac2fJBClip = Left$(s, n - 3) & "..." Else ac2fJBClip = s
End Function

Private Function ac2fJBMask(ByVal k As String) As String
    If Len(k) = 0 Then
        ac2fJBMask = "<not set>"
    ElseIf Len(k) <= 8 Then
        ac2fJBMask = "****"
    Else
        ac2fJBMask = "****" & Right$(k, 4)
    End If
End Function

Private Function ac2fJBIsCount(ByVal s As String) As Boolean
    Dim i As Long
    s = Trim$(s)
    If Len(s) = 0 Or Len(s) > 3 Then Exit Function
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Function
    Next i
    ac2fJBIsCount = True
End Function

' Field idx (0 based) of a tab separated row.
Private Function ac2fJBTab(ByVal row As String, ByVal idx As Long) As String
    Dim f() As String
    f = Split(row, vbTab)
    If idx >= LBound(f) And idx <= UBound(f) Then ac2fJBTab = f(idx)
End Function

Private Function ac2fJBSummary(ByRef jr As ac2fJobRec) As String
    Dim s As String, where As String, errText As String

    If ActiveDocument Is Nothing Then
        ac2fJBSummary = "(no document open)" & vbCrLf
        Exit Function
    End If
    s = "Customer :  " & ac2fJBOr(jr.CustLabel, "<not linked>") & vbCrLf
    s = s & "Sheet    :  " & ac2fJBMatText(jr) & vbCrLf
    If Len(jr.CustFolder) > 0 Then
        where = ac2fJBCustomerDir(jr, False, errText)
        If Len(where) = 0 Then where = "<day folder not found: see 7>"
        s = s & "Folder   :  " & ac2fJBClip(where, 70) & vbCrLf
        s = s & "           (another day: 12)" & vbCrLf
    End If
    ac2fJBSummary = s
End Function

'=====================================================================
' THE LINK, KEPT IN THE DOCUMENT
'=====================================================================

Private Function ac2fJBLoad(ByRef jr As ac2fJobRec) As Boolean
    Dim s As String
    Dim f() As String

    s = ac2fJBPropGet()
    If Left$(s, Len(REC_VER)) <> REC_VER Then
        On Error Resume Next
        s = GetSetting(AC2F_REG_APP, JOB_SEC, ac2fJBDocKey(), "")
        On Error GoTo 0
    End If
    If Left$(s, Len(REC_VER)) <> REC_VER Then Exit Function

    f = Split(s & String$(16, Chr$(31)), Chr$(31))
    jr.CustId = f(1)
    jr.CustFolder = ac2fJBSafeName(f(2))  ' an older, spaced name finds its folder by ac2fJBCustSub
    jr.CustLabel = f(3)
    jr.MatId = f(4)
    jr.MatOwner = f(5)
    jr.MatKind = f(6)
    jr.MatThick = f(7)
    jr.MatBrand = f(8)
    jr.MatGrade = f(9)
    jr.MatColor = f(10)
    jr.MatW = Val(f(11))
    jr.MatH = Val(f(12))
    jr.MatLabel = f(13)
    jr.JobDay = f(14)
    jr.JobAsked = f(15)
    ac2fJBLoad = True
End Function

Private Sub ac2fJBSave(ByRef jr As ac2fJobRec)
    Dim s As String
    Dim us As String
    Dim d As Object

    us = Chr$(31)
    s = REC_VER & us & jr.CustId & us & jr.CustFolder & us & jr.CustLabel & us & _
        jr.MatId & us & jr.MatOwner & us & jr.MatKind & us & jr.MatThick & us & _
        jr.MatBrand & us & jr.MatGrade & us & jr.MatColor & us & _
        ac2fNumStr(jr.MatW) & us & ac2fNumStr(jr.MatH) & us & jr.MatLabel & us & _
        jr.JobDay & us & jr.JobAsked

    ' Late bound: a CorelDRAW without document properties must still
    ' compile, and then only the registry copy is kept.
    On Error Resume Next
    Set d = ActiveDocument
    d.Properties.Item(PROP_NAME, PROP_ID) = s
    Err.Clear
    SaveSetting AC2F_REG_APP, JOB_SEC, ac2fJBDocKey(), s
End Sub

Private Function ac2fJBPropGet() As String
    Dim v As Variant
    Dim d As Object
    On Error Resume Next
    Set d = ActiveDocument
    v = Empty
    v = d.Properties.Item(PROP_NAME, PROP_ID)
    If Err.Number <> 0 Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    ac2fJBPropGet = CStr(v)
End Function

' The registry copy is keyed by a hash of the file path. A file never
' saved is keyed by its title and the date, so Graphic1 tomorrow is not
' mistaken for Graphic1 today.
Private Function ac2fJBDocKey() As String
    Dim f As String
    On Error Resume Next
    f = ActiveDocument.FullFileName
    If Len(f) = 0 Then f = "unsaved|" & ActiveDocument.Title & "|" & CStr(CLng(Date))
    ac2fJBDocKey = "D" & ac2fJBHash(LCase$(f))
End Function

' Two polynomial hashes kept in Doubles, which stay exact well past the
' modulus and never overflow the way Long arithmetic would.
Private Function ac2fJBHash(ByVal s As String) As String
    Const HMOD As Double = 2147483629#
    Dim a As Double, b As Double
    Dim i As Long, w As Long

    For i = 1 To Len(s)
        w = AscW(Mid$(s, i, 1))
        If w < 0 Then w = w + 65536
        a = a * 31# + w
        a = a - Int(a / HMOD) * HMOD
        b = b * 131# + w
        b = b - Int(b / HMOD) * HMOD
    Next i
    ac2fJBHash = Right$("0000000" & Hex$(CLng(a)), 8) & Right$("0000000" & Hex$(CLng(b)), 8)
End Function

'---------------------------------------------------------------------
' Recent lists: rows joined by Chr(30), newest first.
'---------------------------------------------------------------------

Private Function ac2fJBRecentLoad(ByVal key As String, ByRef rows() As String) As Long
    Dim s As String
    Dim parts() As String
    Dim i As Long, n As Long

    ReDim rows(1 To MAX_RECENT)
    s = ac2fGetStr(key, "")
    If Len(s) = 0 Then Exit Function
    parts = Split(s, Chr$(30))
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 And n < MAX_RECENT Then
            n = n + 1
            rows(n) = parts(i)
        End If
    Next i
    ac2fJBRecentLoad = n
End Function

Private Sub ac2fJBRecentPush(ByVal key As String, ByVal row As String)
    Dim rows() As String
    Dim i As Long, n As Long, kept As Long
    Dim s As String, id As String

    n = ac2fJBRecentLoad(key, rows)
    id = ac2fJBRowId(row)
    s = row
    kept = 1
    For i = 1 To n
        If ac2fJBRowId(rows(i)) <> id And kept < MAX_RECENT Then
            s = s & Chr$(30) & rows(i)
            kept = kept + 1
        End If
    Next i
    ac2fSetStr key, s
End Sub

' A row's identity: its StockTrack id, or the row itself when typed.
Private Function ac2fJBRowId(ByVal row As String) As String
    Dim id As String
    id = ac2fJBTab(row, 0)
    If Len(id) = 0 Then id = LCase$(row)
    ac2fJBRowId = id
End Function

'=====================================================================
' LIST WINDOW
'=====================================================================

' A list to click in, with a box above it that narrows the list as you
' type (Turkish letters match either way: ozturk finds the dotted and
' dotless spellings). It is a small Windows Forms window run through
' PowerShell, which comes with Windows and which Plot already uses.
'   okLabel   the main button, acting on the item picked
'   itemBtns  more buttons on the item picked: "O=Open the .cdr|C=Carry on"
'   textBtns  buttons that send the box's text back: "S=Search|T=Use as typed"
' Returns the item number (act = its button's code, "" for the main one),
' 0 when text came back (act = the button's code, typed = the text), -1
' on cancel, and -2 when the window could not be shown (List window is 0,
' or PowerShell is blocked): the caller then asks the old way.
Private Function ac2fJBListWin(ByVal head As String, ByRef labels() As String, _
                               ByVal n As Long, ByVal okLabel As String, _
                               ByVal itemBtns As String, ByVal textBtns As String, _
                               ByVal hint As String, ByRef act As String, _
                               ByRef typed As String) As Long
    Dim ps As String, cfg As String, outf As String, body As String, cmd As String
    Dim f() As String
    Dim i As Long, rc As Long
    Dim sh As Object

    ac2fJBListWin = -2
    act = ""
    typed = ""
    If ac2fGetLng(AC2F_K_JB_LIST, AC2F_DEF_JB_LIST) = 0 Then Exit Function
    ps = ac2fJBTempPath("ac2f_list.ps1")
    cfg = ac2fJBTempPath("ac2f_list_in.txt")
    outf = ac2fJBTempPath("ac2f_list_out.txt")
    On Error Resume Next
    Kill outf
    On Error GoTo 0

    body = ac2fJBLine(ac2fTitle(CAPTION_)) & vbCrLf & ac2fJBLine(head) & vbCrLf & _
           ac2fJBLine(hint) & vbCrLf & ac2fJBLine(okLabel) & vbCrLf & _
           ac2fJBLine(itemBtns) & vbCrLf & ac2fJBLine(textBtns)
    For i = 1 To n
        body = body & vbCrLf & ac2fJBOr(ac2fJBLine(labels(i)), "-")
    Next i
    If Not ac2fJBWriteListScript(ps) Then Exit Function
    If Not ac2fJBWriteUtf8(cfg, body) Then Exit Function

    On Error GoTo Fail
    Set sh = CreateObject("WScript.Shell")
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden" & _
          " -File """ & ps & """ -In """ & cfg & """ -Out """ & outf & """"
    rc = sh.Run(cmd, 0, True)
    On Error GoTo 0
    If rc <> 0 Or Not ac2fJBFileExists(outf) Then Exit Function

    body = ac2fJBReadUtf8(outf)
    If Len(body) > 0 Then
        If AscW(Left$(body, 1)) = &HFEFF Then body = Mid$(body, 2)
    End If
    f = Split(body & vbTab & vbTab & vbTab, vbTab)
    Select Case f(0)
        Case "PICK"
            i = CLng(Val(f(1)))
            If i >= 1 And i <= n Then
                act = f(2)
                ac2fJBListWin = i
            Else
                ac2fJBListWin = -1
            End If
        Case "TEXT"
            act = f(1)
            typed = Trim$(f(2))
            ac2fJBListWin = 0
        Case Else
            ac2fJBListWin = -1
    End Select
    Exit Function

Fail:
    ac2fJBListWin = -2
End Function

' One line of the list file: no line breaks (a heading's become ~n) or tabs.
Private Function ac2fJBLine(ByVal s As String) As String
    s = Replace$(s, vbCrLf, "~n")
    s = Replace$(s, vbCr, "~n")
    s = Replace$(s, vbLf, "~n")
    ac2fJBLine = Replace$(s, vbTab, " ")
End Function

Private Function ac2fJBTempPath(ByVal nm As String) As String
    Dim t As String
    t = Environ$("TEMP")
    If Len(t) = 0 Then t = Environ$("TMP")
    If Len(t) = 0 Then t = "C:\Windows\Temp"
    ac2fJBTempPath = ac2fJBNoSlash(t) & "\" & nm
End Function

' UTF-8, so customer names keep their Turkish letters in the window.
Private Function ac2fJBWriteUtf8(ByVal path As String, ByVal txt As String) As Boolean
    Dim st As Object

    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' text
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2       ' overwrite
    st.Close
    ac2fJBWriteUtf8 = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
End Function

' The window. Written out each run so it always matches this module.
Private Function ac2fJBWriteListScript(ByVal ps As String) As Boolean
    Dim f As Integer

    On Error GoTo Fail
    f = FreeFile
    Open ps For Output As #f
    Print #f, "param([string]$In, [string]$Out)"
    Print #f, "$ErrorActionPreference = 'Stop'"
    Print #f, "Add-Type -AssemblyName System.Windows.Forms"
    Print #f, "Add-Type -AssemblyName System.Drawing"
    Print #f, "[System.Windows.Forms.Application]::EnableVisualStyles()"
    Print #f, "$all = [System.IO.File]::ReadAllLines($In, [System.Text.Encoding]::UTF8)"
    Print #f, "function T([string]$s) { return $s.Replace('~n', [Environment]::NewLine) }"
    Print #f, "function Btns([string]$spec) {"
    Print #f, "  $r = @()"
    Print #f, "  foreach ($p in $spec.Split('|')) {"
    Print #f, "    $i = $p.IndexOf('=')"
    Print #f, "    if ($i -gt 0) { $r += ,@($p.Substring(0, $i), $p.Substring($i + 1)) }"
    Print #f, "  }"
    Print #f, "  return ,$r"
    Print #f, "}"
    Print #f, "function Fold([string]$s) {"
    Print #f, "  $t = $s.Replace([string][char]0x130, 'i').Replace([string][char]0x131, 'i').ToLowerInvariant()"
    Print #f, "  $t = $t.Normalize([System.Text.NormalizationForm]::FormD)"
    Print #f, "  $b = New-Object System.Text.StringBuilder"
    Print #f, "  foreach ($c in $t.ToCharArray()) {"
    Print #f, "    if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) { [void]$b.Append($c) }"
    Print #f, "  }"
    Print #f, "  return $b.ToString()"
    Print #f, "}"
    Print #f, "$items = @()"
    Print #f, "if ($all.Length -gt 6) { $items = @($all[6..($all.Length - 1)]) }"
    Print #f, "$folded = @()"
    Print #f, "foreach ($x in $items) { $folded += (Fold $x) }"
    Print #f, "$itemBtns = Btns $all[4]"
    Print #f, "$textBtns = Btns $all[5]"
    Print #f, "$script:res = 'CANCEL'"
    Print #f, "$script:map = New-Object System.Collections.Generic.List[int]"
    Print #f, ""
    Print #f, "$f = New-Object System.Windows.Forms.Form"
    Print #f, "$f.Text = $all[0]"
    Print #f, "$f.StartPosition = 'CenterScreen'"
    Print #f, "$f.Size = New-Object System.Drawing.Size(760, 580)"
    Print #f, "$f.MinimumSize = New-Object System.Drawing.Size(520, 380)"
    Print #f, "$f.Font = New-Object System.Drawing.Font('Segoe UI', 10)"
    Print #f, "$f.TopMost = $true"
    Print #f, "$f.KeyPreview = $true"
    Print #f, "$f.ShowInTaskbar = $true"
    Print #f, ""
    Print #f, "$tl = New-Object System.Windows.Forms.TableLayoutPanel"
    Print #f, "$tl.Dock = 'Fill'"
    Print #f, "$tl.ColumnCount = 1"
    Print #f, "$tl.RowCount = 5"
    Print #f, "$tl.Padding = New-Object System.Windows.Forms.Padding(12)"
    Print #f, "[void]$tl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))"
    Print #f, "for ($i = 0; $i -lt 5; $i++) {"
    Print #f, "  if ($i -eq 2) { $st = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100) }"
    Print #f, "  else { $st = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize) }"
    Print #f, "  [void]$tl.RowStyles.Add($st)"
    Print #f, "}"
    Print #f, ""
    Print #f, "$head = New-Object System.Windows.Forms.Label"
    Print #f, "$head.Text = T $all[1]"
    Print #f, "$head.AutoSize = $true"
    Print #f, "$head.MaximumSize = New-Object System.Drawing.Size(720, 0)"
    Print #f, "$head.Font = New-Object System.Drawing.Font('Segoe UI', 10.5, [System.Drawing.FontStyle]::Bold)"
    Print #f, "$head.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)"
    Print #f, ""
    Print #f, "$box = New-Object System.Windows.Forms.TextBox"
    Print #f, "$box.Dock = 'Fill'"
    Print #f, "$box.Font = New-Object System.Drawing.Font('Segoe UI', 11)"
    Print #f, "$box.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)"
    Print #f, ""
    Print #f, "$list = New-Object System.Windows.Forms.ListBox"
    Print #f, "$list.Dock = 'Fill'"
    Print #f, "$list.IntegralHeight = $false"
    Print #f, "$list.Font = New-Object System.Drawing.Font('Segoe UI', 10.5)"
    Print #f, "$list.HorizontalScrollbar = $true"
    Print #f, ""
    Print #f, "$hint = New-Object System.Windows.Forms.Label"
    Print #f, "$hint.AutoSize = $true"
    Print #f, "$hint.MaximumSize = New-Object System.Drawing.Size(720, 0)"
    Print #f, "$hint.ForeColor = [System.Drawing.Color]::DimGray"
    Print #f, "$hint.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 6)"
    Print #f, ""
    Print #f, "$bar = New-Object System.Windows.Forms.FlowLayoutPanel"
    Print #f, "$bar.Dock = 'Fill'"
    Print #f, "$bar.AutoSize = $true"
    Print #f, "$bar.FlowDirection = 'RightToLeft'"
    Print #f, "$bar.WrapContents = $true"
    Print #f, ""
    Print #f, "function Finish([string]$r) { $script:res = $r; $f.Close() }"
    Print #f, "function Picked([string]$code) {"
    Print #f, "  if ($list.SelectedIndex -ge 0) { Finish ('PICK' + ""`t"" + ($script:map[$list.SelectedIndex] + 1) + ""`t"" + $code) }"
    Print #f, "}"
    Print #f, "function SendText([string]$code) {"
    Print #f, "  $t = $box.Text.Trim()"
    Print #f, "  if ($t.Length -gt 0) { Finish ('TEXT' + ""`t"" + $code + ""`t"" + $t) } else { $box.Focus() }"
    Print #f, "}"
    Print #f, "function Fill {"
    Print #f, "  $tokens = @((Fold $box.Text).Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries))"
    Print #f, "  $list.BeginUpdate()"
    Print #f, "  $list.Items.Clear()"
    Print #f, "  $script:map.Clear()"
    Print #f, "  for ($i = 0; $i -lt $items.Count; $i++) {"
    Print #f, "    $ok = $true"
    Print #f, "    foreach ($t in $tokens) { if ($folded[$i].IndexOf($t) -lt 0) { $ok = $false; break } }"
    Print #f, "    if ($ok) { [void]$list.Items.Add($items[$i]); $script:map.Add($i) }"
    Print #f, "  }"
    Print #f, "  $list.EndUpdate()"
    Print #f, "  if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }"
    Print #f, "  $okBtn.Enabled = ($list.Items.Count -gt 0)"
    Print #f, "  foreach ($b in $script:itemButtons) { $b.Enabled = $okBtn.Enabled }"
    Print #f, "  $hint.Text = '' + $list.Items.Count + ' / ' + $items.Count + '    ' + (T $all[2])"
    Print #f, "}"
    Print #f, "function OnEnter {"
    Print #f, "  $t = $box.Text.Trim()"
    Print #f, "  $byText = $null"
    Print #f, "  foreach ($b in $textBtns) {"
    Print #f, "    if ($t.StartsWith('=') -and $b[0] -eq 'T') { $byText = 'T' }"
    Print #f, "  }"
    Print #f, "  if (-not $byText -and $list.SelectedIndex -ge 0) { Picked ''; return }"
    Print #f, "  if (-not $byText -and $textBtns.Count -gt 0) { $byText = $textBtns[0][0] }"
    Print #f, "  if ($byText -and $t.Length -gt 0) { SendText $byText }"
    Print #f, "}"
    Print #f, ""
    Print #f, "$cancel = New-Object System.Windows.Forms.Button"
    Print #f, "$cancel.Text = 'Cancel'"
    Print #f, "$cancel.AutoSize = $true"
    Print #f, "$cancel.Add_Click({ Finish 'CANCEL' })"
    Print #f, "[void]$bar.Controls.Add($cancel)"
    Print #f, ""
    Print #f, "$okBtn = New-Object System.Windows.Forms.Button"
    Print #f, "$okBtn.Text = $all[3]"
    Print #f, "$okBtn.AutoSize = $true"
    Print #f, "$okBtn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)"
    Print #f, "$okBtn.Add_Click({ Picked '' })"
    Print #f, "[void]$bar.Controls.Add($okBtn)"
    Print #f, ""
    Print #f, "$script:itemButtons = @()"
    Print #f, "foreach ($spec in $itemBtns) {"
    Print #f, "  $b = New-Object System.Windows.Forms.Button"
    Print #f, "  $b.Text = $spec[1]"
    Print #f, "  $b.AutoSize = $true"
    Print #f, "  $b.Tag = $spec[0]"
    Print #f, "  $b.Add_Click({ param($s, $e) Picked ([string]$s.Tag) })"
    Print #f, "  [void]$bar.Controls.Add($b)"
    Print #f, "  $script:itemButtons += $b"
    Print #f, "}"
    Print #f, "foreach ($spec in $textBtns) {"
    Print #f, "  $b = New-Object System.Windows.Forms.Button"
    Print #f, "  $b.Text = $spec[1]"
    Print #f, "  $b.AutoSize = $true"
    Print #f, "  $b.Tag = $spec[0]"
    Print #f, "  $b.Add_Click({ param($s, $e) SendText ([string]$s.Tag) })"
    Print #f, "  [void]$bar.Controls.Add($b)"
    Print #f, "}"
    Print #f, ""
    Print #f, "$box.Add_TextChanged({ Fill })"
    Print #f, "$box.Add_KeyDown({ param($s, $e)"
    Print #f, "  if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Down) {"
    Print #f, "    if ($list.SelectedIndex -lt $list.Items.Count - 1) { $list.SelectedIndex = $list.SelectedIndex + 1 }"
    Print #f, "    $e.Handled = $true; $e.SuppressKeyPress = $true"
    Print #f, "  } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Up) {"
    Print #f, "    if ($list.SelectedIndex -gt 0) { $list.SelectedIndex = $list.SelectedIndex - 1 }"
    Print #f, "    $e.Handled = $true; $e.SuppressKeyPress = $true"
    Print #f, "  } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {"
    Print #f, "    $e.Handled = $true; $e.SuppressKeyPress = $true"
    Print #f, "    OnEnter"
    Print #f, "  }"
    Print #f, "})"
    Print #f, "$list.Add_DoubleClick({ Picked '' })"
    Print #f, "$list.Add_KeyDown({ param($s, $e)"
    Print #f, "  if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $e.Handled = $true; $e.SuppressKeyPress = $true; Picked '' }"
    Print #f, "})"
    Print #f, "$f.Add_KeyDown({ param($s, $e)"
    Print #f, "  if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { Finish 'CANCEL' }"
    Print #f, "  elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::F3) {"
    Print #f, "    foreach ($b in $textBtns) { if ($b[0] -eq 'S') { SendText 'S' } }"
    Print #f, "  }"
    Print #f, "})"
    Print #f, "$f.Add_Shown({ $f.Activate(); $box.Focus() })"
    Print #f, ""
    Print #f, "$tl.Controls.Add($head, 0, 0)"
    Print #f, "$tl.Controls.Add($box, 0, 1)"
    Print #f, "$tl.Controls.Add($list, 0, 2)"
    Print #f, "$tl.Controls.Add($hint, 0, 3)"
    Print #f, "$tl.Controls.Add($bar, 0, 4)"
    Print #f, "$f.Controls.Add($tl)"
    Print #f, "Fill"
    Print #f, "[void]$f.ShowDialog()"
    Print #f, "[System.IO.File]::WriteAllText($Out, $script:res, (New-Object System.Text.UTF8Encoding($false)))"
    Close #f
    ac2fJBWriteListScript = True
    Exit Function
Fail:
    On Error Resume Next
    Close #f
End Function

'=====================================================================
' STOCKTRACK
'=====================================================================

' Server address with the API prefix, no trailing slash.
Private Function ac2fJBServer() As String
    ac2fJBServer = ac2fGetStr(AC2F_K_JB_SERVER, "")
End Function

' "192.168.1.50" -> "http://192.168.1.50/api/v1"
Private Function ac2fJBNormServer(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    Do While Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
    Loop
    If LCase$(Left$(s, 7)) <> "http://" And LCase$(Left$(s, 8)) <> "https://" Then
        s = "http://" & s
    End If
    If LCase$(Right$(s, 4)) = "/api" Then
        s = s & "/v1"
    ElseIf InStr(LCase$(s), "/api/") = 0 Then
        s = s & "/api/v1"
    End If
    ac2fJBNormServer = s
End Function

' Data lines of a StockTrack reply, or -1 with errText when the server
' could not be used. When the server is away (no answer, or the gateway
' in front of it says it is down) the last good reply to the same request
' is used instead; customer and sheet searches are then run on every
' customer / stock sheet seen so far. `search` is that search text.
Private Function ac2fJBGetRows(ByVal path As String, ByRef rows() As String, _
                               ByRef errText As String, _
                               Optional ByVal search As String = "") As Long
    Dim srv As String, body As String
    Dim code As Long, n As Long
    Dim answered As Boolean

    ac2fJBGetRows = -1
    mJBSaved = ""
    srv = ac2fJBServer()
    If Len(srv) = 0 Then
        errText = "StockTrack is not set up. Enter its address and key in Job setup."
        Exit Function
    End If
    answered = ac2fJBHttpGet(srv & path, body, code, errText)
    If answered And code = 200 And Left$(body, Len(HTTP_HEAD)) = HTTP_HEAD Then
        ac2fJBCacheSave path, body
        ac2fJBGetRows = ac2fJBBodyRows(body, rows)
        Exit Function
    End If

    If Not answered Or code = 502 Or code = 503 Or code = 504 Then
        n = ac2fJBCacheRows(path, search, rows)
        If n >= 0 Then
            ac2fJBGetRows = n
            Exit Function
        End If
    End If

    If Not answered Then
        errText = "StockTrack did not answer at" & vbCrLf & "   " & srv & vbCrLf & vbCrLf & errText
    ElseIf code <> 200 Then
        errText = "StockTrack answered " & code & ":" & vbCrLf & ac2fJBJsonMessage(body)
    Else
        errText = srv & vbCrLf & "answers, but not as StockTrack's Corel interface." & vbCrLf & _
                  "Check the address (it ends in /api/v1) and update StockTrack."
    End If
End Function

' The data lines of a reply body (the header and # lines dropped).
Private Function ac2fJBBodyRows(ByVal body As String, ByRef rows() As String) As Long
    Dim lines() As String
    Dim i As Long, n As Long

    lines = Split(Replace$(body, vbCr, ""), vbLf)
    ReDim rows(1 To UBound(lines) - LBound(lines) + 2)
    For i = LBound(lines) To UBound(lines)
        If Len(lines(i)) > 0 And Left$(lines(i), 1) <> "#" Then
            n = n + 1
            rows(n) = lines(i)
        End If
    Next i
    ac2fJBBodyRows = n
End Function

'---------------------------------------------------------------------
' Offline copy: %LOCALAPPDATA%\ac2f\stocktrack. One file per request,
' plus every customer and every business stock sheet seen (a customer's
' own sheets are left out: offline they could be offered to another
' customer). Kept apart per server and key, so a computer that works
' for two shops never mixes their lists.
'---------------------------------------------------------------------

' "customers" or "materials" for the lists searched offline, else empty.
Private Function ac2fJBCacheList(ByVal path As String) As String
    If Left$(path, 16) = "/corel/customers" And Mid$(path, 17, 1) <> "/" Then
        ac2fJBCacheList = "customers"
    ElseIf Left$(path, 16) = "/corel/materials" Then
        ac2fJBCacheList = "materials"
    End If
End Function

Private Function ac2fJBCacheFile(ByVal what As String) As String
    Dim d As String
    d = Environ$("LOCALAPPDATA")
    If Len(d) = 0 Then d = Environ$("TEMP")
    ac2fJBCacheFile = d & "\ac2f\stocktrack\" & _
                      ac2fJBHash(ac2fJBServer() & "|" & ac2fGetStr(AC2F_K_JB_KEY, "") & "|" & what) & ".tsv"
End Function

Private Sub ac2fJBCacheSave(ByVal path As String, ByVal body As String)
    Dim f As String, old As String, list As String
    Dim rows() As String, n As Long, i As Long, kept As Long

    On Error Resume Next
    f = ac2fJBCacheFile(path)
    If Not ac2fJBMakeDirs(Left$(f, InStrRev(f, "\") - 1)) Then Exit Sub
    ac2fJBWriteText f, body
    list = ac2fJBCacheList(path)
    If Len(list) = 0 Then Exit Sub
    n = ac2fJBBodyRows(body, rows)
    If list = "materials" Then
        For i = 1 To n
            If ac2fJBTab(rows(i), 1) = "B" Then
                kept = kept + 1
                rows(kept) = rows(i)
            End If
        Next i
        n = kept
    End If
    f = ac2fJBCacheFile(list)
    If ac2fJBFileExists(f) Then old = ac2fJBReadUtf8(f)
    ac2fJBWriteText f, ac2fJBMergeLines(rows, n, old, CACHE_LINES)
End Sub

' New lines first, then the old ones whose id (first field) is not among
' them; at most `most` lines.
Private Function ac2fJBMergeLines(ByRef rows() As String, ByVal n As Long, _
                                  ByVal old As String, ByVal most As Long) As String
    Dim seen As String, out As String, id As String
    Dim prev() As String
    Dim i As Long, kept As Long

    seen = vbLf
    For i = 1 To n
        id = ac2fJBTab(rows(i), 0)
        If Len(id) > 0 And InStr(seen, vbLf & id & vbLf) = 0 And kept < most Then
            seen = seen & id & vbLf
            out = out & rows(i) & vbLf
            kept = kept + 1
        End If
    Next i
    prev = Split(Replace$(old, vbCr, ""), vbLf)
    For i = LBound(prev) To UBound(prev)
        id = ac2fJBTab(prev(i), 0)
        If Len(id) > 0 And InStr(seen, vbLf & id & vbLf) = 0 And kept < most Then
            seen = seen & id & vbLf
            out = out & prev(i) & vbLf
            kept = kept + 1
        End If
    Next i
    ac2fJBMergeLines = out
End Function

' Lines of the offline copy, or -1 when there is none. A customer or sheet
' search that was never asked before is answered from every one seen: each
' word must appear in the line (folder name, label, phone, colour ...).
Private Function ac2fJBCacheRows(ByVal path As String, ByVal search As String, _
                                 ByRef rows() As String) As Long
    Dim f As String, words() As String, all() As String, hay As String
    Dim i As Long, j As Long, n As Long, nAll As Long
    Dim hit As Boolean

    ac2fJBCacheRows = -1
    On Error GoTo Done
    f = ac2fJBCacheFile(path)
    If ac2fJBFileExists(f) Then
        mJBSaved = Format$(FileDateTime(f), "dd.mm.yyyy hh:nn")
        ac2fJBCacheRows = ac2fJBBodyRows(ac2fJBReadUtf8(f), rows)
        Exit Function
    End If
    If Len(ac2fJBCacheList(path)) = 0 Or Len(Trim$(search)) = 0 Then Exit Function
    f = ac2fJBCacheFile(ac2fJBCacheList(path))
    If Not ac2fJBFileExists(f) Then Exit Function

    nAll = ac2fJBBodyRows(ac2fJBReadUtf8(f), all)
    words = Split(LCase$(ac2fJBFold(Trim$(search))), " ")
    ReDim rows(1 To nAll + 1)
    For i = 1 To nAll
        hay = LCase$(all(i))
        hit = True
        For j = LBound(words) To UBound(words)
            If Len(words(j)) > 0 Then
                If InStr(hay, words(j)) = 0 Then hit = False
            End If
        Next j
        If hit Then
            n = n + 1
            rows(n) = all(i)
        End If
    Next i
    mJBSaved = Format$(FileDateTime(f), "dd.mm.yyyy hh:nn")
    ac2fJBCacheRows = n
Done:
End Function

' "  (offline, saved 17.10.2026 09:30)" when the last list came from the
' offline copy, else empty.
Private Function ac2fJBSavedNote() As String
    If Len(mJBSaved) > 0 Then ac2fJBSavedNote = "  (offline, saved " & mJBSaved & ")"
End Function

' One GET. MSXML's server object is used rather than the browser one: it
' takes time-outs and does not answer from the Internet Explorer cache.
Private Function ac2fJBHttpGet(ByVal url As String, ByRef body As String, _
                               ByRef code As Long, ByRef errText As String) As Boolean
    Dim x As Object

    On Error Resume Next
    Set x = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If x Is Nothing Then Set x = CreateObject("MSXML2.ServerXMLHTTP")
    If x Is Nothing Then Set x = CreateObject("WinHttp.WinHttpRequest.5.1")
    Err.Clear
    On Error GoTo Fail
    If x Is Nothing Then
        errText = "No HTTP component (MSXML or WinHttp) on this computer."
        Exit Function
    End If

    x.setTimeouts 3000, 3000, 5000, 10000
    x.Open "GET", url, False
    x.setRequestHeader "X-Api-Key", ac2fGetStr(AC2F_K_JB_KEY, "")
    x.setRequestHeader "Accept", "text/plain"
    x.send
    code = x.Status
    body = x.responseText
    ac2fJBHttpGet = True
    Exit Function

Fail:
    errText = Err.Description
End Function

' The message of a JSON error reply: {"message":"..."} or ["..."].
Private Function ac2fJBJsonMessage(ByVal body As String) As String
    Dim p As Long
    p = InStr(body, """message""")
    If p = 0 Then
        ac2fJBJsonMessage = ac2fJBClip(body, 200)
    Else
        ac2fJBJsonMessage = ac2fJBJsonStrAt(body, p + 9)
    End If
End Function

' Value of "key" in a JSON text, when it is a string.
Private Function ac2fJBJsonStr(ByVal js As String, ByVal key As String) As String
    Dim p As Long
    p = InStr(js, """" & key & """")
    If p > 0 Then ac2fJBJsonStr = ac2fJBJsonStrAt(js, p + Len(key) + 2)
End Function

' Reads the JSON string that follows position p, past the colon (and an
' opening bracket, for message lists). Empty when it is not a string.
Private Function ac2fJBJsonStrAt(ByVal js As String, ByVal p As Long) As String
    Dim c As String, out As String

    On Error Resume Next
    Do While p <= Len(js)
        c = Mid$(js, p, 1)
        If c = """" Then Exit Do
        If InStr(" :[" & vbTab & vbCr & vbLf, c) = 0 Then Exit Function
        p = p + 1
    Loop
    p = p + 1
    Do While p <= Len(js)
        c = Mid$(js, p, 1)
        If c = """" Then Exit Do
        If c = "\" Then
            p = p + 1
            c = Mid$(js, p, 1)
            Select Case c
                Case "n": c = vbLf
                Case "t": c = vbTab
                Case "r", "b", "f": c = ""
                Case "u"
                    c = ChrW$(CLng("&H" & Mid$(js, p + 1, 4)))
                    p = p + 4
            End Select
        End If
        out = out & c
        p = p + 1
    Loop
    ac2fJBJsonStrAt = out
End Function

' Search text for a URL: folded to ASCII, then percent encoded.
Private Function ac2fJBUrl(ByVal s As String) As String
    Dim i As Long, a As Long
    Dim c As String, out As String

    s = ac2fJBFold(s)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        a = Asc(c)
        If (a >= 48 And a <= 57) Or (a >= 65 And a <= 90) Or (a >= 97 And a <= 122) _
           Or c = "-" Or c = "." Or c = "_" Or c = "~" Then
            out = out & c
        Else
            out = out & "%" & Right$("0" & Hex$(a), 2)
        End If
    Next i
    ac2fJBUrl = out
End Function

'=====================================================================
' FOLDERS
'=====================================================================

' Today's folder: <base>\<date><separator><day name>, the way CDR Auto
' Renamer names it. An existing folder that starts with today's date is
' used as it is. create = False only works out where it is.
Private Function ac2fJBDayDir(ByVal create As Boolean, ByRef errText As String) As String
    Dim root As String, fmt As String, dsep As String
    Dim stamp As String, want As String
    Dim names() As String

    If Not ac2fJBDayRule(root, fmt, dsep) Then
        errText = "The daily folder is not known." & vbCrLf & _
                  "Turn on daily folders in CDR Auto Renamer, or set" & vbCrLf & _
                  "the day folder base in Job setup."
        Exit Function
    End If
    stamp = ac2fJBStrftime(fmt, Date)
    want = root & "\" & stamp & dsep & ac2fJBDayName(Date)

    If ac2fJBDirExists(want) Then
        ac2fJBDayDir = want
    ElseIf ac2fJBListNames(root, stamp & "*", True, names) > 0 Then
        ac2fJBDayDir = root & "\" & names(1)
    ElseIf Not create Then
        ac2fJBDayDir = want
    ElseIf ac2fJBMakeDirs(want) Then
        ac2fJBDayDir = want
    Else
        errText = "Could not create the day folder" & vbCrLf & "   " & want
    End If
End Function

' The job's customer folder. It goes under today's day folder, unless
' this job (or, for a job not placed yet, this customer) already has a
' folder on an earlier day: a quote made on the 22nd and approved on the
' 25th. Then, when ask is True, the user chooses between carrying on in
' the earlier folder and starting again under today. The answer is kept
' with the job (save it afterwards), so it is asked once a day.
' An empty result with an empty errText means the user cancelled.
Private Function ac2fJBCustomerDir(ByRef jr As ac2fJobRec, ByVal create As Boolean, _
                                   ByRef errText As String, _
                                   Optional ByVal ask As Boolean = False) As String
    Dim todayDir As String, root As String, todayName As String
    Dim earlier As String, dayName As String, stamp As String
    Dim cdir As String
    Dim p As Long

    If Len(jr.CustFolder) = 0 Then
        errText = "No customer is linked."
        Exit Function
    End If
    todayDir = ac2fJBDayDir(False, errText)
    If Len(todayDir) = 0 Then Exit Function
    p = InStrRev(todayDir, "\")
    root = Left$(todayDir, p - 1)
    todayName = Mid$(todayDir, p + 1)
    stamp = CStr(CLng(Date))

    ' --- an earlier day this job could carry on in ---
    If StrComp(jr.JobDay, todayName, vbTextCompare) = 0 Then
        earlier = ""
    ElseIf Len(jr.JobDay) > 0 And Len(ac2fJBCustSub(root & "\" & jr.JobDay, jr.CustFolder)) > 0 Then
        earlier = jr.JobDay
    Else
        earlier = ac2fJBEarlierDay(root, jr.CustFolder, todayName)
    End If

    dayName = todayName
    If Len(earlier) > 0 Then
        If jr.JobAsked = stamp And StrComp(jr.JobDay, earlier, vbTextCompare) = 0 Then
            dayName = earlier                       ' answered already today
        ElseIf ask Then
            Select Case ac2fJBAskDay(root, earlier, todayName, jr.CustFolder)
                Case vbYes: dayName = earlier
                Case vbNo: dayName = todayName
                Case Else
                    errText = ""
                    Exit Function
            End Select
            jr.JobAsked = stamp
        ElseIf StrComp(jr.JobDay, earlier, vbTextCompare) = 0 Then
            dayName = earlier                       ' only looking: where it was left
        End If
    End If
    If ask Then jr.JobDay = dayName

    cdir = ac2fJBCustPath(root & "\" & dayName, jr.CustFolder)
    If create Then
        If Not ac2fJBMakeDirs(cdir) Then
            errText = "Could not create" & vbCrLf & "   " & cdir
            Exit Function
        End If
    End If
    ac2fJBCustomerDir = cdir
End Function

' Most recent earlier day folder, within Look back days, holding a folder
' for this customer. Empty when there is none or Look back is 0.
Private Function ac2fJBEarlierDay(ByVal root As String, ByVal custFolder As String, _
                                  ByVal todayName As String) As String
    Dim r As String, fmt As String, dsep As String
    Dim names() As String
    Dim lim As Long, back As Long, n As Long, i As Long

    lim = ac2fGetLng(AC2F_K_JB_BACK, AC2F_DEF_JB_BACK)
    If lim <= 0 Then Exit Function
    If Not ac2fJBDayRule(r, fmt, dsep) Then Exit Function
    For back = 1 To lim
        n = ac2fJBListNames(root, ac2fJBStrftime(fmt, Date - back) & "*", True, names)
        For i = 1 To n
            If StrComp(names(i), todayName, vbTextCompare) <> 0 Then
                If Len(ac2fJBCustSub(root & "\" & names(i), custFolder)) > 0 Then
                    ac2fJBEarlierDay = names(i)
                    Exit Function
                End If
            End If
        Next i
    Next back
End Function

' Carry on in the earlier day's folder (Yes), start under today (No),
' or stop (Cancel). Says what is already in the earlier folder.
Private Function ac2fJBAskDay(ByVal root As String, ByVal earlier As String, _
                              ByVal todayName As String, ByVal custFolder As String) As VbMsgBoxResult
    Dim names() As String
    Dim cdir As String, m As String, lastExp As String
    Dim nExp As Long, nDir As Long, nGrp As Long
    Dim i As Long, grpNo As Long, c As Long
    Dim h As String, t As String

    cdir = ac2fJBCustPath(root & "\" & earlier, custFolder)
    nExp = ac2fJBListNames(cdir, "exp*", False, names)
    If nExp > 0 Then lastExp = names(nExp)
    nDir = ac2fJBListNames(cdir, "*", True, names)
    For i = 1 To nDir
        If ac2fJBSplitGroup(names(i), grpNo, h, c, t) Then nGrp = nGrp + 1
    Next i

    m = "CUSTOMER FOLDER  -  " & custFolder & vbCrLf & vbCrLf
    m = m & "This job already has a folder from an earlier day:" & vbCrLf
    m = m & "   " & earlier & "\" & Mid$(cdir, InStrRev(cdir, "\") + 1) & vbCrLf
    m = m & "   " & nExp & " export(s), " & nGrp & " toolpath folder(s)"
    If Len(lastExp) > 0 Then m = m & ", last " & lastExp
    m = m & vbCrLf & vbCrLf
    m = m & "Today is " & todayName & "." & vbCrLf & vbCrLf
    m = m & "Yes       carry on in " & earlier & vbCrLf
    m = m & "No        start a new folder under today" & vbCrLf
    m = m & "Cancel    stop" & vbCrLf & vbCrLf
    m = m & "The answer is kept with the job. To change it later:" & vbCrLf
    m = m & "Customer job menu > 12 (Day folder)."
    ac2fJBAskDay = MsgBox(m, vbQuestion + vbYesNoCancel, ac2fTitle(CAPTION_))
End Function

' Base folder, date format and separator: Job setup's base when one is
' set, otherwise CDR Auto Renamer's; format and separator come from its
' config.json when it is there.
Private Function ac2fJBDayRule(ByRef root As String, ByRef fmt As String, _
                               ByRef dsep As String) As Boolean
    Dim js As String, t As String

    root = ac2fGetStr(AC2F_K_JB_BASE, "")
    fmt = "%d-%m-%Y"
    dsep = "--"

    js = ac2fJBReadUtf8(Environ$("LOCALAPPDATA") & "\CdrAutoRenamer\config.json")
    If Len(js) > 0 Then
        If Len(root) = 0 Then root = ac2fJBJsonStr(js, "base_dir")
        t = ac2fJBJsonStr(js, "date_format")
        If Len(t) > 0 Then fmt = t
        t = ac2fJBJsonStr(js, "separator")
        If Len(t) > 0 Then dsep = t
    End If
    root = ac2fJBNoSlash(Trim$(root))
    ac2fJBDayRule = Len(root) > 0
End Function

' The Python strftime codes CDR Auto Renamer's date_format can use.
Private Function ac2fJBStrftime(ByVal fmt As String, ByVal d As Date) As String
    Dim i As Long
    Dim c As String, out As String

    i = 1
    Do While i <= Len(fmt)
        c = Mid$(fmt, i, 1)
        If c = "%" And i < Len(fmt) Then
            i = i + 1
            Select Case Mid$(fmt, i, 1)
                Case "d": out = out & Right$("0" & CStr(Day(d)), 2)
                Case "m": out = out & Right$("0" & CStr(Month(d)), 2)
                Case "Y": out = out & CStr(Year(d))
                Case "y": out = out & Right$(CStr(Year(d)), 2)
                Case "A": out = out & ac2fJBDayName(d)
                Case "%": out = out & "%"
                Case Else: out = out & "%" & Mid$(fmt, i, 1)
            End Select
        Else
            out = out & c
        End If
        i = i + 1
    Loop
    ac2fJBStrftime = out
End Function

' Turkish day names in ASCII, as CDR Auto Renamer writes them.
Private Function ac2fJBDayName(ByVal d As Date) As String
    ac2fJBDayName = Choose(Weekday(d, vbMonday), "pazartesi", "sali", "carsamba", _
                           "persembe", "cuma", "cumartesi", "pazar")
End Function

' config.json is UTF-8 (a base folder may be called Masaustu with a
' u-umlaut); ADODB.Stream decodes it, a plain read is the fallback.
Private Function ac2fJBReadUtf8(ByVal path As String) As String
    Dim st As Object
    Dim f As Integer

    If Not ac2fJBFileExists(path) Then Exit Function
    On Error Resume Next
    Set st = CreateObject("ADODB.Stream")
    If Not st Is Nothing Then
        st.Type = 2                 ' text
        st.Charset = "utf-8"
        st.Open
        st.LoadFromFile path
        If Err.Number = 0 Then
            ac2fJBReadUtf8 = st.ReadText(-1)
            st.Close
            Exit Function
        End If
        st.Close
        Err.Clear
    End If

    f = FreeFile
    Open path For Input As #f
    If Err.Number = 0 Then ac2fJBReadUtf8 = Input$(LOF(f), f)
    Close #f
End Function

' Names in parent matching pattern: folders when wantDirs, else files.
' Collected in one go, because Dir cannot be nested.
Private Function ac2fJBListNames(ByVal parent As String, ByVal pattern As String, _
                                 ByVal wantDirs As Boolean, ByRef names() As String) As Long
    Dim nm As String
    Dim n As Long
    Dim isDir As Boolean

    ReDim names(1 To 16)
    On Error GoTo Done
    If wantDirs Then
        nm = Dir$(parent & "\" & pattern, vbDirectory)
    Else
        nm = Dir$(parent & "\" & pattern, vbNormal)
    End If
    Do While Len(nm) > 0
        If nm <> "." And nm <> ".." Then
            isDir = (GetAttr(parent & "\" & nm) And vbDirectory) = vbDirectory
            If isDir = wantDirs Then
                n = n + 1
                If n > UBound(names) Then ReDim Preserve names(1 To n * 2)
                names(n) = nm
            End If
        End If
        nm = Dir$()
    Loop
Done:
    ac2fJBListNames = n
End Function

' Uses GetAttr, not Dir, so it is safe inside a Dir listing.
Private Function ac2fJBDirExists(ByVal p As String) As Boolean
    Dim a As Long
    On Error Resume Next
    a = GetAttr(p)
    If Err.Number = 0 Then ac2fJBDirExists = (a And vbDirectory) = vbDirectory
End Function

Private Function ac2fJBFileExists(ByVal p As String) As Boolean
    Dim a As Long
    On Error Resume Next
    a = GetAttr(p)
    If Err.Number = 0 Then ac2fJBFileExists = (a And vbDirectory) = 0
End Function

' MkDir for every missing level. Drive and \\server\share are kept.
Private Function ac2fJBMakeDirs(ByVal p As String) As Boolean
    Dim parts() As String
    Dim i As Long, start As Long
    Dim cur As String

    p = ac2fJBNoSlash(p)
    If ac2fJBDirExists(p) Then
        ac2fJBMakeDirs = True
        Exit Function
    End If
    parts = Split(p, "\")
    If Left$(p, 2) = "\\" Then
        If UBound(parts) < 4 Then Exit Function
        cur = "\\" & parts(2) & "\" & parts(3)
        start = 4
    Else
        cur = parts(0)
        start = 1
    End If
    On Error Resume Next
    For i = start To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & "\" & parts(i)
            If Not ac2fJBDirExists(cur) Then MkDir cur
        End If
    Next i
    ac2fJBMakeDirs = ac2fJBDirExists(p)
End Function

Private Function ac2fJBNoSlash(ByVal p As String) As String
    Do While Len(p) > 3 And (Right$(p, 1) = "\" Or Right$(p, 1) = "/")
        p = Left$(p, Len(p) - 1)
    Loop
    ac2fJBNoSlash = p
End Function

Private Sub ac2fJBExplore(ByVal p As String)
    On Error Resume Next
    Shell "explorer.exe """ & p & """", vbNormalFocus
End Sub

Private Function ac2fJBNeedDoc() As Boolean
    If ActiveDocument Is Nothing Then
        ac2fWarn "Open the customer's file first.", CAPTION_
        Exit Function
    End If
    ac2fJBNeedDoc = True
End Function
