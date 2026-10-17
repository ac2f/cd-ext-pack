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
'    ac2fJobOpen     - open the customer folder in Explorer
'    ac2fJobSetup    - server, key, day folder, separator
'
'  WHERE THINGS GO
'
'    <day>\<customer>\exp01--pleksi.dxf
'    <day>\<customer>\001-pleksi-1220x2440--2pcs--2800um-akpolimer-dokme-beyaz\001
'                                                                           \002
'
'  The day folder is the one CDR Auto Renamer makes every morning. Its
'  config.json is read for the base folder, the date format and the
'  separator, so the two agree without a second setting, and a missing
'  day folder is created the same way.
'
'  The customer folder is the company name when there is one, else the
'  person's name. Every name is plain ASCII: Turkish letters are folded
'  (s for s-cedilla, i for dotless i), since CAM programs and the ANSI
'  file calls VBA uses both trip over them.
'
'  SIZES
'  1220x2440 is a whole sheet. A leading X means only part of it is
'  used: X1220x900 is a 900 mm strip across a 1220 wide sheet. A strip
'  that would leave less than MIN_REMNANT behind counts as whole.
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
Public Const AC2F_K_JB_ROUND  As String = "JBYuvarlamaMM"

Public Const AC2F_DEF_JB_LONG   As Double = 4000#   ' longer sheets get p001 / p002
Public Const AC2F_DEF_JB_DIALOG As Long = 1         ' show the DXF export dialog
Public Const AC2F_DEF_JB_OPEN   As Long = 0         ' open Explorer afterwards
Public Const AC2F_DEF_JB_NEST   As Long = 1         ' picking a sheet sets the nest sheet
Public Const AC2F_DEF_JB_ROUND  As Double = 10#     ' partial sizes round up to this
Public Const AC2F_DEF_JB_SEP    As String = "--"

Private Const JOB_SEC     As String = "Jobs"        ' per document links
Private Const PROP_NAME   As String = "ac2fJob"
Private Const PROP_ID     As Long = 1
Private Const REC_VER     As String = "j1"
Private Const MAX_RECENT  As Long = 9
Private Const MIN_REMNANT As Double = 100#          ' mm; less left over = whole sheet
Private Const HTTP_HEAD   As String = "#ac2f-corel "
Private Const PROMPT_MAX  As Long = 880             ' InputBox stops near 1024

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
End Type

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
            "  5  -  Open the customer folder" & vbCrLf & _
            "  6  -  Job setup (server, key, day folder)"

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
        Case "5": ac2fJobOpen
        Case "6": ac2fJobSetup
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

    custDir = ac2fJBCustomerDir(jr, True, errText)
    If Len(custDir) = 0 Then
        ac2fWarn errText, CAPTION_
        Exit Sub
    End If
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
             "Sheet: " & ac2fJBMatText(jr), CAPTION_
    Exit Sub

Fail:
    ac2fWarn "The export failed: " & Err.Description & vbCrLf & vbCrLf & _
             custDir & "\" & fname, CAPTION_
End Sub

' Creates the toolpath folders for the sheets this job uses, or stacks
' them onto an existing folder for the same sheet. The plan is read off
' the drawing: Nesting's sheet outlines when there are any, otherwise
' the selection measured against the sheet.
Public Sub ac2fJobFolders()
Attribute ac2fJobFolders.VB_Description = "ac2f pack: Create or stack the toolpath folders for this job"
    Dim jr As ac2fJobRec
    Dim custDir As String, errText As String, why As String
    Dim proposal As String, answer As String, m As String
    Dim toks() As String, cnts() As Long, n As Long

    If Not ac2fJBNeedDoc() Then Exit Sub
    ac2fJBLoad jr
    If Not ac2fJBEnsure(jr) Then Exit Sub

    custDir = ac2fJBCustomerDir(jr, False, errText)
    If Len(custDir) = 0 Then
        ac2fWarn errText, CAPTION_
        Exit Sub
    End If

    proposal = ac2fJBPropose(jr, why)
    Do
        m = "TOOLPATH FOLDERS  -  " & jr.CustFolder & vbCrLf
        m = m & "Sheet: " & ac2fJBMatText(jr) & vbCrLf & vbCrLf
        m = m & why & vbCrLf & vbCrLf
        m = m & "Enter          accept what is shown below" & vbCrLf
        m = m & "3              three whole sheets" & vbCrLf
        m = m & "1 X1220x900    count and size (X = part of the sheet)" & vbCrLf
        m = m & "2, 1 X1220x900 several sizes at once" & vbCrLf
        m = m & "M              pick another sheet first"
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
    custDir = ac2fJBCustomerDir(jr, True, errText)
    If Len(custDir) = 0 Then
        ac2fWarn errText, CAPTION_
        Exit Sub
    End If
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
                        If ac2fJBSafeName("a" & arg & "a") <> "a" & arg & "a" Then
                            ac2fWarn "A separator is plain ASCII and cannot hold" & vbCrLf & _
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
            n = ac2fJBGetRows("/corel/customers?limit=12&q=" & ac2fJBUrl(answer), rows, errText)
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
            ElseIf n = 1 Then
                ac2fJBSetCust jr, rows(1)
                ac2fJBPickCustomer = True
                Exit Function
            Else
                k = ac2fJBChoose("Customers matching """ & answer & """:", rows, n, 2, pending)
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
                                 srv, errText)
            If nSrv < 0 Then
                ac2fWarn errText, CAPTION_
            ElseIf nSrv = 0 Then
                ac2fWarn "Nothing in stock matches """ & answer & """.", CAPTION_
                dflt = answer
            ElseIf nSrv = 1 Then
                ac2fJBSetMat jr, srv(1)
                ac2fJBPickMaterial = True
                Exit Function
            Else
                opts = srv
                nOpt = nSrv
                head = "SHEETS MATCHING """ & answer & """"
                dflt = "1"
            End If
        Else
            ac2fWarn "StockTrack is not available, so type the sheet:" & vbCrLf & _
                     "=pleksi;2.8;akpolimer;dokme;beyaz;1220x2440", CAPTION_
            dflt = "=" & ac2fJBToken(answer)
        End If
    Loop
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
    Dim w As Double, h As Double, part As Boolean
    Dim r As String

    f = Split(s & ";;;;;", ";")
    kind = ac2fJBToken(f(0))
    If Len(kind) = 0 Then Exit Function
    If ac2fParseNum(f(1), 0) > 0 Then thick = CStr(CLng(ac2fParseNum(f(1), 0) * 1000#))
    If Len(Trim$(f(5))) > 0 Then
        If Not ac2fJBParseSize(f(5), w, h, part) Then Exit Function
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

    changed = ac2fJBTab(row, 0) <> jr.CustId Or ac2fJBTab(row, 1) <> jr.CustFolder
    jr.CustId = ac2fJBTab(row, 0)
    jr.CustFolder = ac2fJBTab(row, 1)
    jr.CustLabel = ac2fJBOr(ac2fJBTab(row, 2), jr.CustFolder)
    ' the previous customer's own material is not this one's
    If changed And jr.MatOwner = "C" Then ac2fJBSetMat jr, ""
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
' type it ("2 1220x2440, 1 X1220x900"), and says where it came from.
Private Function ac2fJBPropose(ByRef jr As ac2fJobRec, ByRef why As String) As String
    Dim sx0() As Double, sy0() As Double, sx1() As Double, sy1() As Double
    Dim px0() As Double, py0() As Double, pw() As Double, ph() As Double
    Dim ns As Long, np As Long
    Dim toks() As String, cnts() As Long, nt As Long
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim i As Long, j As Long, k As Long
    Dim x As Double, y As Double, w As Double, h As Double
    Dim cx As Double, cy As Double, topY As Double, used As Double
    Dim sw As Double, sh As Double, mT As Double, mB As Double
    Dim tok As String, s As String

    why = "Type how many sheets, and the size if it is not the whole sheet."
    mT = ac2fGetNum(AC2F_K_NS_MT, AC2F_DEF_NS_MT)
    mB = ac2fGetNum(AC2F_K_NS_MB, AC2F_DEF_NS_MB)
    If mT < 0 Then mT = 0
    If mB < 0 Then mB = 0
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
        np = ac2fJBScanParts(px0, py0, pw, ph)
        For j = 1 To ns
            sw = sx1(j) - sx0(j)
            sh = sy1(j) - sy0(j)
            topY = -1E+30
            For i = 1 To np
                cx = px0(i) + pw(i) / 2#
                cy = py0(i) + ph(i) / 2#
                If cx >= sx0(j) And cx <= sx1(j) And cy >= sy0(j) And cy <= sy1(j) _
                   And pw(i) <= sw + 1# And ph(i) <= sh + 1# Then
                    If py0(i) + ph(i) > topY Then topY = py0(i) + ph(i)
                End If
            Next i
            If topY > -1E+29 Then
                ' the sheet as stocked when the outline matches it
                If Abs(sw - jr.MatW) <= 2# And Abs(sh - jr.MatH) <= 2# Then
                    sw = jr.MatW
                    sh = jr.MatH
                End If
                used = topY - sy0(j) + mT
                tok = ac2fJBUsedTok(sw, sh, used)
                k = ac2fJBTally(toks, cnts, nt, tok, 1)
            End If
        Next j
        If nt > 0 Then why = "From the " & ns & " nest sheet outline(s) on the page:"

    ' --- otherwise the selection, against the sheet ---
    ElseIf jr.MatW > 0 And jr.MatH > 0 And Not sr Is Nothing Then
        If sr.Count > 0 Then
            sr.GetBoundingBox x, y, w, h
            s = "Selection " & Format$(w, "0") & " x " & Format$(h, "0") & " mm"
            If w <= jr.MatW + 1# And h <= jr.MatH + 1# Then
                tok = ac2fJBUsedTok(jr.MatW, jr.MatH, h + mT + mB)
                why = s & ", along the sheet:"
            ElseIf h <= jr.MatW + 1# And w <= jr.MatH + 1# Then
                tok = ac2fJBUsedTok(jr.MatW, jr.MatH, w + mT + mB)
                why = s & ", turned on the sheet:"
            Else
                tok = ac2fJBSizeTok(jr.MatW, jr.MatH, False)
                why = s & " is bigger than one sheet; how many?"
            End If
            k = ac2fJBTally(toks, cnts, nt, tok, 1)
        End If
    End If

Done:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    If nt = 0 And jr.MatW > 0 And jr.MatH > 0 Then
        k = ac2fJBTally(toks, cnts, nt, ac2fJBSizeTok(jr.MatW, jr.MatH, False), 1)
    End If
    For i = 1 To nt
        If i > 1 Then s = s & ", " Else s = ""
        s = s & cnts(i) & " " & toks(i)
    Next i
    If nt = 0 Then s = ""
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
    On Error Resume Next
    For i = 1 To ActivePage.Shapes.Count
        Set s = Nothing
        Set s = ActivePage.Shapes(i)
        nm = ""
        nm = LCase$(s.Name)
        If nm <> "ac2f sheets" And Not (nm Like "ac2f sheet #*") And Not s Is Nothing Then
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
    ac2fJBScanParts = n
End Function

' Size token for a sheet sw x sh of which a strip "used" mm long is cut.
Private Function ac2fJBUsedTok(ByVal sw As Double, ByVal sh As Double, _
                               ByVal used As Double) As String
    Dim stp As Double, u As Double

    stp = ac2fGetNum(AC2F_K_JB_ROUND, AC2F_DEF_JB_ROUND)
    If stp < 1# Then stp = 1#
    u = ac2fCeil(used / stp) * stp
    If u >= sh - MIN_REMNANT Then
        ac2fJBUsedTok = ac2fJBSizeTok(sw, sh, False)
    Else
        ac2fJBUsedTok = ac2fJBSizeTok(sw, u, True)
    End If
End Function

Private Function ac2fJBSizeTok(ByVal w As Double, ByVal h As Double, _
                               ByVal part As Boolean) As String
    Dim s As String
    If part Then s = "X"
    ac2fJBSizeTok = s & CStr(CLng(w)) & "x" & CStr(CLng(h))
End Function

' "1220x2440" / "X1220x900" -> width, height and whether it is a part.
Private Function ac2fJBParseSize(ByVal t As String, ByRef w As Double, _
                                 ByRef h As Double, ByRef part As Boolean) As Boolean
    Dim p As Long

    t = LCase$(Trim$(t))
    part = False
    If Left$(t, 1) = "x" Then
        part = True
        t = Mid$(t, 2)
    End If
    t = Replace$(t, "*", "x")
    p = InStr(t, "x")
    If p <= 1 Or p >= Len(t) Then Exit Function
    w = ac2fParseNum(Left$(t, p - 1), 0)
    h = ac2fParseNum(Mid$(t, p + 1), 0)
    If w <= 0 Or h <= 0 Then Exit Function
    ac2fJBParseSize = True
End Function

' A size typed against the job's sheet: turned round if it only fits
' that way, and marked as a part when it is smaller than the sheet.
Private Function ac2fJBNormTok(ByRef jr As ac2fJobRec, ByVal w As Double, _
                               ByVal h As Double, ByVal part As Boolean) As String
    Dim t As Double

    If jr.MatW > 0 And jr.MatH > 0 Then
        If (w > jr.MatW + 1# Or h > jr.MatH + 1#) And _
           h <= jr.MatW + 1# And w <= jr.MatH + 1# Then
            t = w: w = h: h = t
        End If
        If w < jr.MatW - 1# Or h < jr.MatH - 1# Then part = True
    End If
    ac2fJBNormTok = ac2fJBSizeTok(w, h, part)
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

' Reads "2 1220x2440, 1 X1220x900". A count alone means whole sheets.
' Returns the number of sizes, or -1 with errText.
Private Function ac2fJBParsePlan(ByRef jr As ac2fJobRec, ByVal txt As String, _
                                 ByRef toks() As String, ByRef cnts() As Long, _
                                 ByRef errText As String) As Long
    Dim entries() As String, words() As String
    Dim i As Long, j As Long, k As Long, n As Long
    Dim cnt As Long, tok As String
    Dim w As Double, h As Double, part As Boolean

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
                ElseIf ac2fJBParseSize(words(j), w, h, part) Then
                    tok = ac2fJBNormTok(jr, w, h, part)
                Else
                    errText = "Cannot read """ & words(j) & """." & vbCrLf & _
                              "Write a count, a size, or both: 2 X1220x900"
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
                tok = ac2fJBSizeTok(jr.MatW, jr.MatH, False)
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
    Dim w As Double, h As Double, part As Boolean, lim As Double

    ReDim found(1 To n): ReDim oldCnt(1 To n): ReDim grp(1 To n): ReDim splits(1 To n)
    tail = ac2fJBSpec(jr)
    lim = ac2fGetNum(AC2F_K_JB_LONG, AC2F_DEF_JB_LONG)

    ' --- dry run: what is there already ---
    For i = 1 To n
        head = jr.MatKind & "-" & toks(i)
        found(i) = ac2fJBFindGroup(custDir, head, tail, c, maxNo)
        oldCnt(i) = c
        If ac2fJBParseSize(toks(i), w, h, part) Then
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
    If jr.MatW > 0 And jr.MatH > 0 Then s = s & " " & ac2fJBSizeTok(jr.MatW, jr.MatH, False)
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

' Folder name, the same rule StockTrack applies: ASCII, case and spaces
' kept, characters Windows refuses become spaces, no trailing dot.
Private Function ac2fJBSafeName(ByVal s As String) As String
    Dim i As Long
    Dim c As String, out As String

    s = ac2fJBFold(s)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If InStr("\/:*?""<>|", c) > 0 Then c = " "
        out = out & c
    Next i
    Do While InStr(out, "  ") > 0
        out = Replace$(out, "  ", " ")
    Loop
    out = Trim$(out)
    If Len(out) > 60 Then out = Left$(out, 60)
    Do While Right$(out, 1) = "." Or Right$(out, 1) = " "
        out = Left$(out, Len(out) - 1)
    Loop
    Select Case LCase$(out)
        Case "con", "prn", "aux", "nul", "com1", "com2", "com3", "com4", "com5", _
             "com6", "com7", "com8", "com9", "lpt1", "lpt2", "lpt3", "lpt4", _
             "lpt5", "lpt6", "lpt7", "lpt8", "lpt9"
            out = out & "_"
    End Select
    ac2fJBSafeName = out
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
        If Len(where) = 0 Then where = "<day folder not found: see 6>"
        s = s & "Folder   :  " & ac2fJBClip(where, 70) & vbCrLf
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

    f = Split(s & String$(14, Chr$(31)), Chr$(31))
    jr.CustId = f(1)
    jr.CustFolder = f(2)
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
        ac2fNumStr(jr.MatW) & us & ac2fNumStr(jr.MatH) & us & jr.MatLabel

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
' could not be used.
Private Function ac2fJBGetRows(ByVal path As String, ByRef rows() As String, _
                               ByRef errText As String) As Long
    Dim srv As String, body As String
    Dim code As Long, i As Long, n As Long
    Dim lines() As String

    ac2fJBGetRows = -1
    srv = ac2fJBServer()
    If Len(srv) = 0 Then
        errText = "StockTrack is not set up. Enter its address and key in Job setup."
        Exit Function
    End If
    If Not ac2fJBHttpGet(srv & path, body, code, errText) Then
        errText = "StockTrack did not answer at" & vbCrLf & "   " & srv & vbCrLf & vbCrLf & errText
        Exit Function
    End If
    If code <> 200 Then
        errText = "StockTrack answered " & code & ":" & vbCrLf & ac2fJBJsonMessage(body)
        Exit Function
    End If
    If Left$(body, Len(HTTP_HEAD)) <> HTTP_HEAD Then
        errText = srv & vbCrLf & "answers, but not as StockTrack's Corel interface." & vbCrLf & _
                  "Check the address (it ends in /api/v1) and update StockTrack."
        Exit Function
    End If

    lines = Split(Replace$(body, vbCr, ""), vbLf)
    ReDim rows(1 To UBound(lines) - LBound(lines) + 1)
    For i = LBound(lines) To UBound(lines)
        If Len(lines(i)) > 0 And Left$(lines(i), 1) <> "#" Then
            n = n + 1
            rows(n) = lines(i)
        End If
    Next i
    ac2fJBGetRows = n
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

Private Function ac2fJBCustomerDir(ByRef jr As ac2fJobRec, ByVal create As Boolean, _
                                   ByRef errText As String) As String
    Dim dayDir As String

    If Len(jr.CustFolder) = 0 Then
        errText = "No customer is linked."
        Exit Function
    End If
    dayDir = ac2fJBDayDir(create, errText)
    If Len(dayDir) = 0 Then Exit Function
    If create Then
        If Not ac2fJBMakeDirs(dayDir & "\" & jr.CustFolder) Then
            errText = "Could not create" & vbCrLf & "   " & dayDir & "\" & jr.CustFolder
            Exit Function
        End If
    End If
    ac2fJBCustomerDir = dayDir & "\" & jr.CustFolder
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
