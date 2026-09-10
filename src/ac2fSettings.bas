Attribute VB_Name = "ac2fSettings"
'=====================================================================
'  ac2f pack  --  ac2fSettings
'
'  One settings sheet for the whole package, plus named profiles.
'
'  Every setting lives in one table (ac2fSettingsTable) that carries its
'  key, label, limits and help text. The sheet, the profile store and
'  the temporary override parser all read that single table, so adding
'  a setting means adding one row and nothing else.
'
'  Macros:
'    ac2fSettings        - the settings sheet (all settings, profiles)
'    ac2fProfiles        - jumps straight to profile management
'
'  PROFILES
'  A profile is the whole settings set saved under a name, stored as one
'  registry string "key=value|key=value|...". Loading writes the values
'  back; running with a profile can also apply one-off overrides that
'  are never written to the registry.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Settings"

' Value kinds
Public Const AC2F_KIND_NUM  As Long = 0
Public Const AC2F_KIND_INT  As Long = 1

Private Const LAB_W As Long = 18
Private Const VAL_W As Long = 6

Public Type ac2fSetting
    Key    As String
    Group  As String
    Label  As String
    Unit   As String
    Kind   As Long
    DefVal As Double
    MinVal As Double
    MaxVal As Double      ' 0 = no upper limit
    Help   As String
End Type

Private m_set() As ac2fSetting
Private m_setN  As Long

'=====================================================================
' SETTINGS TABLE
'=====================================================================

' Builds the table once. Adding a setting means adding one row here.
Private Sub ac2fBuildTable()
    If m_setN > 0 Then Exit Sub

    ReDim m_set(0 To 31)
    m_setN = 0

    '--- LED module ---------------------------------------------------
    ac2fAddSetting AC2F_K_SPACING, "LED MODULE", "Module spacing", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_SPACING, 0.001, 0, _
        "Distance between two LED modules, measured along the outline." & vbCrLf & _
        "This is the single value that drives the module count: the " & _
        "length of every outline is divided by it." & vbCrLf & vbCrLf & _
        "Smaller spacing means more modules, brighter and more even " & _
        "light, higher cost and higher power draw. Typical channel " & _
        "letter work sits between 80 and 150 mm; go tighter on shallow " & _
        "letters where the light has less room to spread."

    ac2fAddSetting AC2F_K_LEDS, "LED MODULE", "LEDs per module", "", _
        AC2F_KIND_INT, CDbl(AC2F_DEF_LEDS), 1, 0, _
        "How many LED chips sit on one module. 3 for the common " & _
        "3-LED modules this package is built around." & vbCrLf & vbCrLf & _
        "It does not change the module count. It only multiplies the " & _
        "module count into the LED count reported for ordering."

    ac2fAddSetting AC2F_K_MODULE_W, "LED MODULE", "Module power", "W", _
        AC2F_KIND_NUM, AC2F_DEF_MODULE_W, 0, 0, _
        "Power drawn by a single module, from its datasheet." & vbCrLf & vbCrLf & _
        "Total power = modules x this value. Everything downstream " & _
        "(safety margin, power supply count) is built on it, so an " & _
        "optimistic number here undersizes the whole supply chain. " & _
        "0.72 W is typical for a 3-LED 12 V module."

    ac2fAddSetting AC2F_K_PSU_W, "LED MODULE", "Power supply", "W", _
        AC2F_KIND_NUM, AC2F_DEF_PSU_W, 0, 0, _
        "Rated output of one power supply." & vbCrLf & vbCrLf & _
        "The report divides the required power by this and rounds up " & _
        "to get the number of supplies." & vbCrLf & vbCrLf & _
        "Set it to 0 if you size supplies yourself; the power supply " & _
        "line is then left out of the report."

    ac2fAddSetting AC2F_K_SAFETY, "LED MODULE", "Safety margin", "%", _
        AC2F_KIND_NUM, AC2F_DEF_SAFETY, 0, 0, _
        "Headroom added on top of the calculated load before the " & _
        "power supplies are counted." & vbCrLf & vbCrLf & _
        "Running a supply at its full rating shortens its life and " & _
        "leaves nothing for inrush or a hot sign box. 20% is a common " & _
        "starting point; 30% is safer for enclosed or sunlit signs."

    ac2fAddSetting AC2F_K_MINPERPATH, "LED MODULE", "Min per outline", "", _
        AC2F_KIND_INT, CDbl(AC2F_DEF_MINPERPATH), 0, 0, _
        "The fewest modules any single outline may receive." & vbCrLf & vbCrLf & _
        "Without it, an outline shorter than the module spacing would " & _
        "round down to zero and a small counter, a dot on an i or an " & _
        "accent would be left unlit. Keep it at 1 unless you " & _
        "deliberately want to leave small shapes dark."

    ac2fAddSetting AC2F_K_METHOD, "LED MODULE", "Count method", "", _
        AC2F_KIND_INT, CDbl(AC2F_DEF_METHOD), 1, 2, _
        "How the measured length is turned into a module count." & vbCrLf & vbCrLf & _
        "1 = Perimeter based. Uses the full length of every outline. " & _
        "Correct when modules follow the line itself, such as modules " & _
        "on the return wall of a channel letter." & vbCrLf & vbCrLf & _
        "2 = Centreline estimate. Uses half of each outline. A letter " & _
        "drawn as a closed band has an inner and an outer contour, so " & _
        "its perimeter is roughly twice the real run; this halves it." & _
        vbCrLf & vbCrLf & _
        "Method 2 is an estimate. It is good when the band width is " & _
        "even and drifts when it varies. Compare against a real job " & _
        "and put the difference into the correction factor."

    ac2fAddSetting AC2F_K_FACTOR, "LED MODULE", "Correction factor", "", _
        AC2F_KIND_NUM, AC2F_DEF_FACTOR, 0.01, 0, _
        "Scales the final module count. 1 leaves it unchanged." & vbCrLf & vbCrLf & _
        "This is the calibration handle. After a real job, divide the " & _
        "modules you actually used by the number this reported and " & _
        "enter the result. From then on the estimate matches your own " & _
        "practice instead of the generic model."

    '--- Box letter ---------------------------------------------------
    ac2fAddSetting AC2F_K_BL_THICK, "BOX LETTER", "Thickness", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_THICK, 0.01, 0, _
        "Thickness of the sheet the return is made from." & vbCrLf & vbCrLf & _
        "It drives two separate things. First the developed length: " & _
        "the neutral axis sits inside the material, so a thicker sheet " & _
        "shifts the flat length further from the drawn outline. The " & _
        "correction is 2*pi*g and depends only on thickness, not on " & _
        "how long the outline is: about 3.5 mm on 1 mm sheet and " & _
        "10.6 mm on 3 mm." & vbCrLf & vbCrLf & _
        "Second the groove spacing: a thicker sheet means a deeper " & _
        "groove, which closes over a smaller angle, so grooves must " & _
        "sit closer together."

    ac2fAddSetting AC2F_K_BL_HEIGHT, "BOX LETTER", "Strip height", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_HEIGHT, 1, 0, _
        "Height of the drawn strip, which is the depth of the letter." & _
        vbCrLf & vbCrLf & _
        "Purely a drawing dimension: it sets how tall the rectangle " & _
        "is and how far the groove lines run. It has no effect on the " & _
        "developed length or on the groove positions." & vbCrLf & vbCrLf & _
        "Add your own flange or return allowance on top if the " & _
        "fabricated part needs one."

    ac2fAddSetting AC2F_K_BL_FLEX, "BOX LETTER", "Flexibility", "", _
        AC2F_KIND_NUM, AC2F_DEF_BL_FLEX, 0.05, 0, _
        "How much turn one groove is trusted to absorb, relative to " & _
        "the geometric estimate. This is the main calibration handle " & _
        "for the groove layout." & vbCrLf & vbCrLf & _
        "Higher = fewer grooves, further apart. Lower = more grooves." & _
        vbCrLf & vbCrLf & _
        "Calibrate it on real work:" & vbCrLf & _
        "  grooves will not close, material binds -> lower it" & vbCrLf & _
        "  far more grooves than the bend needs   -> raise it" & vbCrLf & _
        "  faceted marks on the visible surface   -> leave this alone" & _
        vbCrLf & "     and lower the surface tolerance instead" & vbCrLf & vbCrLf & _
        "The material presets set a starting value. They are starting " & _
        "points, not measured shop data; the number you find on your " & _
        "own first job is the one worth keeping."

    ac2fAddSetting AC2F_K_BL_REF, "BOX LETTER", "Reference face", "", _
        AC2F_KIND_INT, CDbl(AC2F_DEF_BL_REF), 1, 3, _
        "Which face of the strip the drawn vector represents. This " & _
        "decides which way the neutral axis is offset, so getting it " & _
        "wrong shifts the developed length by twice the correction." & _
        vbCrLf & vbCrLf & _
        "1 = Outer face. The usual case: the letter outline is drawn " & _
        "flush with the outside of the return, and the material lies " & _
        "inside it. Offset = (1 - K) x thickness inwards." & vbCrLf & vbCrLf & _
        "2 = Inner face. The outline marks the inside of the return. " & _
        "Offset = K x thickness outwards." & vbCrLf & vbCrLf & _
        "3 = Neutral axis. The outline is already the neutral line, " & _
        "so no correction is applied." & vbCrLf & vbCrLf & _
        "Hole (counter) outlines are detected automatically and the " & _
        "offset direction is flipped for them."

    ac2fAddSetting AC2F_K_BL_KFAC, "BOX LETTER", "K factor", "", _
        AC2F_KIND_NUM, AC2F_DEF_BL_KFAC, 0, 1, _
        "Where the neutral axis sits across the thickness, as a " & _
        "fraction measured from the inner surface." & vbCrLf & vbCrLf & _
        "When sheet is bent the outside stretches and the inside " & _
        "compresses; one line in between keeps its length, and that " & _
        "is the line the flat pattern is measured along." & vbCrLf & vbCrLf & _
        "0.44 suits most sheet metal. Tighter bend radii push it " & _
        "lower, generous radii push it towards 0.5. Only worth " & _
        "touching if your fabricated parts come out consistently " & _
        "long or short by a few millimetres."

    ac2fAddSetting AC2F_K_BL_DEPTH, "BOX LETTER", "Groove depth ratio", "", _
        AC2F_KIND_NUM, AC2F_DEF_BL_DEPTH, 0.05, 0.95, _
        "Depth of the groove as a fraction of the material thickness." & _
        vbCrLf & vbCrLf & _
        "0.7 means the cutter removes 70% of the thickness and leaves " & _
        "30% as the hinge." & vbCrLf & vbCrLf & _
        "Deeper grooves bend more easily and let grooves sit further " & _
        "apart, but the remaining hinge gets thin and can crack or " & _
        "tear. Shallower is stronger and needs more grooves." & vbCrLf & vbCrLf & _
        "This must match what your machine or cutter actually does. " & _
        "It is not a target; it is a description of your tooling."

    ac2fAddSetting AC2F_K_BL_MOUTH, "BOX LETTER", "Max groove mouth", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_MOUTH, 0.05, 0, _
        "The widest groove opening that still closes without leaving " & _
        "a visible mark on the outside of the bend." & vbCrLf & vbCrLf & _
        "A groove of depth d closing through angle a opens by about " & _
        "d x a at the surface. Turning that around gives the largest " & _
        "turn one groove may take, and therefore how far apart the " & _
        "grooves can sit on a given radius." & vbCrLf & vbCrLf & _
        "Lower it if closed grooves show as dents or gaps on finished " & _
        "letters."

    ac2fAddSetting AC2F_K_BL_TOL, "BOX LETTER", "Surface tolerance", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_TOL, 0.01, 0, _
        "How far the flat between two grooves may sit off the true " & _
        "curve." & vbCrLf & vbCrLf & _
        "Grooving turns a curve into a chain of short straight " & _
        "facets. The gap at the middle of a facet is about " & _
        "s squared / (8 x R), so demanding a smaller gap forces the " & _
        "grooves closer together." & vbCrLf & vbCrLf & _
        "This is the setting that controls how round the finished " & _
        "letter looks. Lower it when you can see flats on the " & _
        "curves; 0.1 mm or less for close viewing, 0.3 mm is fine " & _
        "for signs read from across a street."

    ac2fAddSetting AC2F_K_BL_SMIN, "BOX LETTER", "Min groove spacing", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_SMIN, 0.5, 0, _
        "Grooves are never placed closer together than this, whatever " & _
        "the curve asks for." & vbCrLf & vbCrLf & _
        "It reflects a physical limit: the cutter has a width, and " & _
        "overlapping grooves would cut the strip through. On a very " & _
        "tight radius this clamp is what stops the layout from " & _
        "shredding the material." & vbCrLf & vbCrLf & _
        "It is also used as the step when a sharp corner needs " & _
        "several grooves side by side."

    ac2fAddSetting AC2F_K_BL_SMAX, "BOX LETTER", "Max groove spacing", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_SMAX, 1, 0, _
        "Grooves are never placed further apart than this, even on a " & _
        "very gentle curve." & vbCrLf & vbCrLf & _
        "A ceiling that keeps long sweeping curves from being bent " & _
        "over just two or three grooves, which would show as visible " & _
        "kinks rather than a smooth arc." & vbCrLf & vbCrLf & _
        "Straight segments never receive grooves, so this does not " & _
        "add grooves to flat sections."

    ac2fAddSetting AC2F_K_BL_CORNER, "BOX LETTER", "Corner threshold", "deg", _
        AC2F_KIND_NUM, AC2F_DEF_BL_CORNER, 0.1, 0, _
        "A direction change larger than this counts as a corner and " & _
        "gets its own groove, drawn in pink." & vbCrLf & vbCrLf & _
        "Below the threshold the turn is treated as part of the " & _
        "curve and absorbed by the ordinary grooves." & vbCrLf & vbCrLf & _
        "Raise it if smooth joins in your artwork are being marked " & _
        "as corners. Lower it if genuine sharp corners are being " & _
        "missed. If a corner turns more than one groove can absorb, " & _
        "several grooves are placed side by side automatically."

    ac2fAddSetting AC2F_K_BL_COIL, "BOX LETTER", "Coil length", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_COIL, 0, 0, _
        "Usable length of one coil or sheet of return material." & vbCrLf & vbCrLf & _
        "When a strip comes out longer than this, red marks are drawn " & _
        "across it showing where to cut and join. The report also " & _
        "estimates how many pieces the job needs." & vbCrLf & vbCrLf & _
        "Set it to 0 if you do not want the strip split at all." & _
        vbCrLf & vbCrLf & _
        "The strip itself is always drawn in one piece; only the cut " & _
        "positions are marked."

    ac2fAddSetting AC2F_K_BL_JOINT, "BOX LETTER", "Joint allowance", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_JOINT, 0, 0, _
        "Extra length added at each joint for the overlap or backing " & _
        "strip." & vbCrLf & vbCrLf & _
        "Each piece after the first effectively contributes the coil " & _
        "length minus this allowance, which is why the piece count " & _
        "rises faster than a plain division would suggest." & vbCrLf & vbCrLf & _
        "It must be smaller than the coil length. If it is not, the " & _
        "cut marks and the piece estimate are skipped rather than " & _
        "producing a meaningless answer."

    ac2fAddSetting AC2F_K_BL_GAP, "BOX LETTER", "Strip gap", "mm", _
        AC2F_KIND_NUM, AC2F_DEF_BL_GAP, 0, 0, _
        "Vertical space left between strips when several are drawn " & _
        "below one another." & vbCrLf & vbCrLf & _
        "Purely a layout convenience with no effect on any " & _
        "calculation. Raise it if the labels above each strip run " & _
        "into the strip above."
End Sub

Private Sub ac2fAddSetting(ByVal key As String, ByVal grp As String, _
                           ByVal label As String, ByVal unit As String, _
                           ByVal kind As Long, ByVal defVal As Double, _
                           ByVal minVal As Double, ByVal maxVal As Double, _
                           ByVal help As String)
    If m_setN > UBound(m_set) Then
        ReDim Preserve m_set(0 To (UBound(m_set) + 1) * 2 - 1)
    End If
    m_set(m_setN).Key = key
    m_set(m_setN).Group = grp
    m_set(m_setN).Label = label
    m_set(m_setN).Unit = unit
    m_set(m_setN).Kind = kind
    m_set(m_setN).DefVal = defVal
    m_set(m_setN).MinVal = minVal
    m_set(m_setN).MaxVal = maxVal
    m_set(m_setN).Help = help
    m_setN = m_setN + 1
End Sub

'=====================================================================
' PUBLIC ACCESS TO THE TABLE
'=====================================================================

Public Function ac2fSettingCount() As Long
    ac2fBuildTable
    ac2fSettingCount = m_setN
End Function

' Index is 1 based, matching the numbers shown on the sheet.
Public Function ac2fSettingAt(ByVal idx As Long) As ac2fSetting
    ac2fBuildTable
    If idx >= 1 And idx <= m_setN Then ac2fSettingAt = m_set(idx - 1)
End Function

' Finds a setting by number, key or a unique piece of its label.
Public Function ac2fFindSetting(ByVal token As String) As Long
    Dim i As Long, n As Long, hits As Long, hit As Long
    Dim t As String

    ac2fBuildTable
    t = Trim$(token)
    If Len(t) = 0 Then Exit Function

    If IsNumeric(t) Then
        n = CLng(Val(t))
        If n >= 1 And n <= m_setN Then ac2fFindSetting = n
        Exit Function
    End If

    For i = 0 To m_setN - 1
        If StrComp(m_set(i).Key, t, vbTextCompare) = 0 Then
            ac2fFindSetting = i + 1
            Exit Function
        End If
    Next i

    For i = 0 To m_setN - 1
        If InStr(1, m_set(i).Label, t, vbTextCompare) > 0 Then
            hits = hits + 1
            hit = i + 1
        End If
    Next i
    If hits = 1 Then ac2fFindSetting = hit
End Function

Public Function ac2fSettingValue(ByVal idx As Long) As Double
    Dim st As ac2fSetting
    st = ac2fSettingAt(idx)
    If Len(st.Key) = 0 Then Exit Function
    ac2fSettingValue = ac2fGetNum(st.Key, st.DefVal)
End Function

' Clamps and stores a value. Returns the value actually stored.
Public Function ac2fSettingStore(ByVal idx As Long, ByVal v As Double, _
                                 Optional ByVal tempOnly As Boolean = False) As Double
    Dim st As ac2fSetting
    Dim x As Double

    st = ac2fSettingAt(idx)
    If Len(st.Key) = 0 Then Exit Function

    x = v
    If x < st.MinVal Then x = st.MinVal
    If st.MaxVal > 0 And x > st.MaxVal Then x = st.MaxVal
    If st.Kind = AC2F_KIND_INT Then x = CDbl(CLng(x))

    If tempOnly Then
        ac2fSetOverride st.Key, x
    Else
        ac2fSetNum st.Key, x
    End If
    ac2fSettingStore = x
End Function

Private Function ac2fSettingText(ByVal idx As Long) As String
    Dim st As ac2fSetting
    Dim v As Double

    st = ac2fSettingAt(idx)
    v = ac2fSettingValue(idx)
    If st.Kind = AC2F_KIND_INT Then
        ac2fSettingText = ac2fFmt(v, 0)
    Else
        ac2fSettingText = ac2fFmt(v)
    End If
End Function

'=====================================================================
' THE SETTINGS SHEET  (single menu)
'=====================================================================

Public Sub ac2fSettings()
Attribute ac2fSettings.VB_Description = "ac2f pack: All settings and profiles in one sheet"
    Dim cmd As String

    ac2fBuildTable
    Do
        cmd = InputBox(ac2fSheet(False), ac2fTitle(CAPTION_), "")
        If StrPtr(cmd) = 0 Then Exit Do          ' Cancel
        If Len(Trim$(cmd)) = 0 Then Exit Do      ' Enter closes
    Loop While ac2fHandle(cmd, False)
End Sub

' Renders the whole settings sheet. tempMode marks it as a one-off run.
Private Function ac2fSheet(ByVal tempMode As Boolean) As String
    Dim s As String
    Dim i As Long
    Dim st As ac2fSetting
    Dim grp As String
    Dim prof As String

    prof = ac2fActiveProfile()
    If Len(prof) = 0 Then prof = "<none>"

    If tempMode Then
        s = "RUN SETTINGS - changes apply to this run only" & vbCrLf
    Else
        s = "SETTINGS  (profile: " & prof & ")" & vbCrLf
    End If
    s = s & vbCrLf

    For i = 1 To m_setN
        st = ac2fSettingAt(i)
        If st.Group <> grp Then
            grp = st.Group
            s = s & "-- " & grp & " --" & vbCrLf
        End If
        s = s & ac2fRPad(CStr(i), 2) & " " & ac2fPad(st.Label, LAB_W) & " " & _
                ac2fRPad(ac2fSettingText(i), VAL_W) & " " & st.Unit
        If ac2fIsOverridden(st.Key) Then s = s & " *"
        s = s & vbCrLf
    Next i

    s = s & vbCrLf
    s = s & "N=value   change        ?N   explain" & vbCrLf
    If tempMode Then
        s = s & "Enter     run" & vbCrLf
        s = s & "* = changed for this run only"
    Else
        s = s & "P         profiles      R    reset" & vbCrLf
        s = s & "Enter     close"
    End If

    ac2fSheet = s
End Function

Private Function ac2fIsOverridden(ByVal key As String) As Boolean
    Dim v As Double
    ac2fIsOverridden = ac2fGetOverride(key, v)
End Function

' Executes one sheet command. Returns True to keep the sheet open.
Private Function ac2fHandle(ByVal cmd As String, ByVal tempMode As Boolean) As Boolean
    Dim t As String
    Dim p As Long
    Dim idx As Long
    Dim v As Double
    Dim lhs As String, rhs As String

    ac2fHandle = True
    t = Trim$(cmd)
    If Len(t) = 0 Then Exit Function

    ' ?N  -> explain
    If Left$(t, 1) = "?" Then
        idx = ac2fFindSetting(Mid$(t, 2))
        If idx = 0 Then
            ac2fWarn "No setting matches """ & Mid$(t, 2) & """." & vbCrLf & _
                     "Use its number, for example ?11", CAPTION_
        Else
            ac2fExplain idx
        End If
        Exit Function
    End If

    If Not tempMode Then
        If StrComp(t, "P", vbTextCompare) = 0 Then
            ac2fProfiles
            Exit Function
        End If
        If StrComp(t, "R", vbTextCompare) = 0 Then
            ac2fResetToDefaults
            Exit Function
        End If
    End If

    ' N=value, possibly several separated by spaces or semicolons
    t = Replace$(t, ";", " ")
    Do
        p = InStr(t, " ")
        If p > 0 Then
            lhs = Left$(t, p - 1)
            t = Trim$(Mid$(t, p + 1))
        Else
            lhs = t
            t = ""
        End If

        If Len(lhs) > 0 Then
            p = InStr(lhs, "=")
            If p = 0 Then
                ac2fWarn "Do not understand """ & lhs & """." & vbCrLf & vbCrLf & _
                         "Use  N=value  to change a setting" & vbCrLf & _
                         "     ?N       to read what it does", CAPTION_
            Else
                rhs = Mid$(lhs, p + 1)
                idx = ac2fFindSetting(Left$(lhs, p - 1))
                If idx = 0 Then
                    ac2fWarn "No setting matches """ & Left$(lhs, p - 1) & """.", CAPTION_
                Else
                    v = ac2fSettingStore(idx, ac2fParseNum(rhs, ac2fSettingValue(idx)), tempMode)
                End If
            End If
        End If
    Loop While Len(t) > 0
End Function

Private Sub ac2fExplain(ByVal idx As Long)
    Dim st As ac2fSetting
    Dim s As String

    st = ac2fSettingAt(idx)

    s = idx & ".  " & st.Label
    If Len(st.Unit) > 0 Then s = s & "  (" & st.Unit & ")"
    s = s & vbCrLf & String$(46, "-") & vbCrLf & vbCrLf
    s = s & st.Help & vbCrLf & vbCrLf
    s = s & String$(46, "-") & vbCrLf
    s = s & "Now      : " & ac2fSettingText(idx) & " " & st.Unit
    If ac2fIsOverridden(st.Key) Then s = s & "   (this run only)"
    s = s & vbCrLf
    s = s & "Default  : " & ac2fNumStr(st.DefVal) & " " & st.Unit & vbCrLf
    s = s & "Range    : " & ac2fNumStr(st.MinVal)
    If st.MaxVal > 0 Then
        s = s & " to " & ac2fNumStr(st.MaxVal)
    Else
        s = s & " and up"
    End If

    ac2fInfo s, "Setting " & idx
End Sub

Private Sub ac2fResetToDefaults()
    If MsgBox("Every setting goes back to its default value." & vbCrLf & _
              "Saved profiles are kept." & vbCrLf & vbCrLf & "Continue?", _
              vbQuestion + vbYesNo, ac2fTitle(CAPTION_)) <> vbYes Then Exit Sub
    ac2fResetSettings
    ac2fSetActiveProfile ""
End Sub

'=====================================================================
' PROFILES
'=====================================================================

Public Sub ac2fProfiles()
Attribute ac2fProfiles.VB_Description = "ac2f pack: Save, load and delete settings profiles"
    Dim cmd As String
    Dim nm As String
    Dim verb As String
    Dim arg As String

    ac2fBuildTable
    Do
        cmd = InputBox(ac2fProfileSheet(), ac2fTitle("Profiles"), "")
        If StrPtr(cmd) = 0 Then Exit Sub
        cmd = Trim$(cmd)
        If Len(cmd) = 0 Then Exit Sub

        verb = ac2fSplitCmd(cmd, arg)
        Select Case verb
            Case "S"
                nm = arg
                If Len(nm) = 0 Then nm = InputBox("Save the current settings as:", _
                                                  ac2fTitle("Profiles"), "")
                If StrPtr(nm) <> 0 Then
                    If Len(Trim$(nm)) > 0 Then ac2fProfileSave Trim$(nm)
                End If
            Case "L"
                nm = ac2fPickProfile(arg)
                If Len(nm) > 0 Then
                    If ac2fProfileLoad(nm) Then
                        ac2fInfo "Profile """ & nm & """ loaded.", "Profiles"
                    End If
                End If
            Case "D"
                nm = ac2fPickProfile(arg)
                If Len(nm) > 0 Then
                    If MsgBox("Delete profile """ & nm & """?", _
                              vbQuestion + vbYesNo, ac2fTitle("Profiles")) = vbYes Then
                        ac2fProfileDelete nm
                    End If
                End If
            Case "M"
                ac2fMaterialPreset
            Case Else
                ac2fWarn "Use S, L, D or M.", "Profiles"
        End Select
    Loop
End Sub

Private Function ac2fProfileSheet() As String
    Dim s As String
    Dim names() As String
    Dim n As Long, i As Long
    Dim act As String

    n = ac2fProfileNames(names)
    act = ac2fActiveProfile()

    s = "PROFILES" & vbCrLf & vbCrLf
    If n = 0 Then
        s = s & "   (none saved yet)" & vbCrLf
    Else
        For i = 0 To n - 1
            s = s & "   " & names(i)
            If StrComp(names(i), act, vbTextCompare) = 0 Then s = s & "   <- loaded"
            s = s & vbCrLf
        Next i
    End If

    s = s & vbCrLf
    s = s & "S name    save the current settings under that name" & vbCrLf
    s = s & "L name    load a profile" & vbCrLf
    s = s & "D name    delete a profile" & vbCrLf
    s = s & "M         apply a material preset" & vbCrLf
    s = s & "Enter     back"
    ac2fProfileSheet = s
End Function

' Returns the saved profile names. Function result is the count.
' Splits a profile command into its verb and its argument.
'
' The verb must stand alone or be followed by a space, otherwise a
' profile named ""Letters3mm"" would be read as verb L plus the name
' ""etters3mm"". Returns an empty verb when the command is not one.
Private Function ac2fSplitCmd(ByVal cmd As String, ByRef arg As String) As String
    Dim t As String
    t = Trim$(cmd)
    arg = ""
    If Len(t) = 0 Then Exit Function
    If Len(t) = 1 Then
        ac2fSplitCmd = UCase$(t)
        Exit Function
    End If
    If Mid$(t, 2, 1) = " " Then
        ac2fSplitCmd = UCase$(Left$(t, 1))
        arg = Trim$(Mid$(t, 3))
    End If
End Function

Public Function ac2fProfileNames(ByRef names() As String) As Long
    Dim raw As Variant
    Dim i As Long, n As Long

    On Error Resume Next
    raw = GetAllSettings(AC2F_REG_APP, AC2F_REG_PROF)
    On Error GoTo 0

    If IsEmpty(raw) Then
        ReDim names(0 To 0)
        Exit Function
    End If

    n = UBound(raw, 1) - LBound(raw, 1) + 1
    ReDim names(0 To n - 1)
    For i = 0 To n - 1
        names(i) = raw(LBound(raw, 1) + i, 0)
    Next i
    ac2fProfileNames = n
End Function

' Resolves a name, asking the user when it is missing or ambiguous.
Private Function ac2fPickProfile(ByVal wanted As String) As String
    Dim names() As String
    Dim n As Long, i As Long
    Dim answer As String

    n = ac2fProfileNames(names)
    If n = 0 Then
        ac2fWarn "No profiles saved yet.", "Profiles"
        Exit Function
    End If

    If Len(wanted) > 0 Then
        For i = 0 To n - 1
            If StrComp(names(i), wanted, vbTextCompare) = 0 Then
                ac2fPickProfile = names(i)
                Exit Function
            End If
        Next i
        ac2fWarn "No profile named """ & wanted & """.", "Profiles"
        Exit Function
    End If

    answer = InputBox("Which profile?" & vbCrLf & vbCrLf & ac2fProfileSheet(), _
                      ac2fTitle("Profiles"), "")
    If StrPtr(answer) = 0 Then Exit Function
    ' An empty answer means cancel. Recursing on it would keep reopening
    ' the box and eventually overflow the stack.
    If Len(Trim$(answer)) = 0 Then Exit Function
    ac2fPickProfile = ac2fPickProfile(Trim$(answer))
End Function

Public Sub ac2fProfileSave(ByVal nm As String)
    Dim i As Long
    Dim st As ac2fSetting
    Dim blob As String

    ac2fBuildTable
    For i = 1 To m_setN
        st = ac2fSettingAt(i)
        If Len(blob) > 0 Then blob = blob & "|"
        blob = blob & st.Key & "=" & ac2fNumStr(ac2fGetNum(st.Key, st.DefVal))
    Next i

    On Error Resume Next
    SaveSetting AC2F_REG_APP, AC2F_REG_PROF, nm, blob
    On Error GoTo 0

    ac2fSetActiveProfile nm
    ac2fInfo "Profile """ & nm & """ saved." & vbCrLf & _
             m_setN & " settings stored.", "Profiles"
End Sub

Public Function ac2fProfileLoad(ByVal nm As String) As Boolean
    Dim blob As String
    Dim parts() As String
    Dim i As Long, p As Long, idx As Long
    Dim key As String, val As String

    ac2fBuildTable

    On Error Resume Next
    blob = GetSetting(AC2F_REG_APP, AC2F_REG_PROF, nm, "")
    On Error GoTo 0

    If Len(blob) = 0 Then
        ac2fWarn "Profile """ & nm & """ is empty or missing.", "Profiles"
        Exit Function
    End If

    parts = Split(blob, "|")
    For i = LBound(parts) To UBound(parts)
        p = InStr(parts(i), "=")
        If p > 0 Then
            key = Left$(parts(i), p - 1)
            val = Mid$(parts(i), p + 1)
            idx = ac2fFindSetting(key)
            ' A key that is no longer in the table is ignored, so an old
            ' profile still loads after a setting is removed.
            If idx > 0 Then ac2fSettingStore idx, Val(val), False
        End If
    Next i

    ac2fSetActiveProfile nm
    ac2fProfileLoad = True
End Function

Public Sub ac2fProfileDelete(ByVal nm As String)
    On Error Resume Next
    DeleteSetting AC2F_REG_APP, AC2F_REG_PROF, nm
    On Error GoTo 0
    If StrComp(ac2fActiveProfile(), nm, vbTextCompare) = 0 Then ac2fSetActiveProfile ""
End Sub

' Material presets: starting values for the box letter groove layout.
' These are calibration starting points, not measured shop data. Save one
' as a profile once you have tuned it on real work.
Private Sub ac2fMaterialPreset()
    Dim answer As String
    Dim flex As Double, depth As Double, kfac As Double
    Dim nm As String

    answer = InputBox( _
        "MATERIAL PRESETS" & vbCrLf & vbCrLf & _
        "  1  Aluminium, thin   (0.5-1.5 mm)" & vbCrLf & _
        "  2  Galvanised, thin  (0.5-1.2 mm)" & vbCrLf & _
        "  3  Aluminium, thick  (2-4 mm)" & vbCrLf & _
        "  4  Stainless steel" & vbCrLf & vbCrLf & _
        "Sets flexibility, groove depth ratio and K factor only." & vbCrLf & _
        "Thickness and strip height are left as they are." & vbCrLf & vbCrLf & _
        "These are starting points, not measured values. Tune the" & vbCrLf & _
        "flexibility on your first job, then save it as a profile.", _
        ac2fTitle("Profiles"), "1")
    If StrPtr(answer) = 0 Then Exit Sub

    Select Case CLng(ac2fParseNum(answer, 0))
        Case 1: flex = 1.2:  depth = 0.7:  kfac = 0.44: nm = "Aluminium thin"
        Case 2: flex = 1#:   depth = 0.65: kfac = 0.44: nm = "Galvanised thin"
        Case 3: flex = 0.8:  depth = 0.75: kfac = 0.42: nm = "Aluminium thick"
        Case 4: flex = 0.7:  depth = 0.6:  kfac = 0.45: nm = "Stainless steel"
        Case Else
            ac2fWarn "Choose 1 to 4.", "Profiles"
            Exit Sub
    End Select

    ac2fSetNum AC2F_K_BL_FLEX, flex
    ac2fSetNum AC2F_K_BL_DEPTH, depth
    ac2fSetNum AC2F_K_BL_KFAC, kfac

    ac2fInfo nm & " preset applied." & vbCrLf & vbCrLf & _
             "   Flexibility        : " & ac2fNumStr(flex) & vbCrLf & _
             "   Groove depth ratio : " & ac2fNumStr(depth) & vbCrLf & _
             "   K factor           : " & ac2fNumStr(kfac), "Profiles"
End Sub

' The settings sheet without its command footer, for the About box.
Public Function ac2fSettingsBrief() As String
    Dim s As String
    Dim i As Long
    Dim st As ac2fSetting
    Dim grp As String

    ac2fBuildTable
    For i = 1 To m_setN
        st = ac2fSettingAt(i)
        If st.Group <> grp Then
            grp = st.Group
            s = s & "-- " & grp & " --" & vbCrLf
        End If
        s = s & "   " & ac2fPad(st.Label, LAB_W) & " " & _
                ac2fRPad(ac2fSettingText(i), VAL_W) & " " & st.Unit & vbCrLf
    Next i
    ac2fSettingsBrief = s
End Function

Public Function ac2fActiveProfile() As String
    ac2fActiveProfile = ac2fGetStr("ActiveProfile", "")
End Function

Public Sub ac2fSetActiveProfile(ByVal nm As String)
    ac2fSetStr "ActiveProfile", nm
End Sub

'=====================================================================
' RUN WITH A PROFILE, WITH OPTIONAL ONE-OFF CHANGES
'=====================================================================

' Loads a profile if one is chosen, then lets the user change any value
' for this run only. Returns False when the user cancels.
Public Function ac2fPrepareRun() As Boolean
    Dim nm As String
    Dim cmd As String
    Dim arg As String
    Dim names() As String
    Dim n As Long

    ac2fBuildTable
    ac2fClearOverrides

    n = ac2fProfileNames(names)
    If n > 0 Then
        cmd = InputBox(ac2fProfileSheet() & vbCrLf & vbCrLf & _
                       "Enter alone keeps the settings as they are.", _
                       ac2fTitle("Run"), "")
        If StrPtr(cmd) = 0 Then Exit Function
        cmd = Trim$(cmd)
        If Len(cmd) > 0 Then
            If ac2fSplitCmd(cmd, arg) = "L" Then
                nm = arg
            Else
                nm = cmd
            End If
            nm = ac2fPickProfile(nm)
            If Len(nm) > 0 Then ac2fProfileLoad nm
        End If
    End If

    ' One-off changes
    Do
        cmd = InputBox(ac2fSheet(True), ac2fTitle("Run"), "")
        If StrPtr(cmd) = 0 Then
            ac2fClearOverrides
            Exit Function
        End If
        If Len(Trim$(cmd)) = 0 Then Exit Do
    Loop While ac2fHandle(cmd, True)

    ac2fPrepareRun = True
End Function
