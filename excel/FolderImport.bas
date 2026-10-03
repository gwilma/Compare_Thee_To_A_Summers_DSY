Attribute VB_Name = "FolderImport"
Option Explicit

' Compare thee to a summer's DSY: folder import.
'
' ImportWeatherFolder walks a folder and all its sub-folders. Each .epw or .csv weather file is
' loaded into the Hourly sheet in turn, the workbook is recalculated, and the Summary metrics
' are written to the Library (one row per file, replacing any row with the same location,
' file type, period, emissions and percentile). Progress is written to the Import log sheet.
' The Hourly data and Settings you had before the import are put back afterwards.
'
' Location, file type (TRY/DSY1-3), period, emissions and percentile come from the file name,
' e.g. Z1_DSY1_2050s_HIGH50_CIBSE_v1.1.epw or London_LHR_DSY2_2080High90.epw. If the name has
' no recognisable location, the name of the folder holding the file is used.
'
' Thresholds: with Settings > "Thresholds for imported files" set to "Each location's TRY",
' the SWCDH threshold and TWCDH offset for every location are derived from that location's
' earliest-period TRY (else DSY1), so all files of a location share the same values.

Private Const COL_METRICS As Long = 8      ' Library column H: the metrics start here, in Summary order.
' The metric count comes from the summary_values range, so adding a metric to the workbook needs no
' change here. The stored thresholds and source path follow the metrics.
Private Const MAX_HOURS As Long = 10000
Private Const MIN_HOURS As Long = 4000     ' fewer than this cannot cover a season plus spin-up

Private mSilent As Boolean
Private mLogRow As Long

' ------------------------------------------------------------------ entry points

Public Sub ImportWeatherFolder()
    mSilent = False
    RunImport ""
End Sub

Public Sub ImportWeatherFolderFrom(ByVal folder As String)
    ' Same as ImportWeatherFolder, without dialogs, for a given folder (used for testing and automation).
    mSilent = True
    RunImport folder
End Sub

Public Sub UseLibraryThresholds()
    ' Copy the thresholds stored with the Library files of the Compare location into Settings,
    ' so the observed year is analysed on the same basis as the files it is compared with.
    mSilent = False
    ApplyLibraryThresholds
End Sub

Public Sub UseLibraryThresholdsSilent()
    mSilent = True
    ApplyLibraryThresholds
End Sub

Private Sub ApplyLibraryThresholds()
    Dim tbl As Variant, r As Long, loc As String
    loc = Trim$(CStr(NR("compare_loc").Value))
    tbl = NR("lib_table").Value
    For r = 1 To UBound(tbl, 1)
        If CStr(tbl(r, 1)) = "1" And (loc = "" Or LCase$(CStr(tbl(r, 2))) = LCase$(loc)) Then
            If IsNumeric(tbl(r, ColSwcdh())) And Not IsEmpty(tbl(r, ColSwcdh())) And Not IsEmpty(tbl(r, ColTwcdh())) Then
                NR("swcdh_override").Value = tbl(r, ColSwcdh())
                NR("twcdh_override").Value = tbl(r, ColTwcdh())
                Application.Calculate
                Say "Settings now use the thresholds stored for " & CStr(tbl(r, 2)) & ": SWCDH " & _
                    Format$(tbl(r, ColSwcdh()), "0.00") & " °C, TWCDH offset " & Format$(tbl(r, ColTwcdh()), "0.00") & " K."
                Exit Sub
            End If
        End If
    Next r
    Say "No included Library row" & IIf(loc = "", "", " for " & loc) & " has stored thresholds."
End Sub

' ------------------------------------------------------------------ main loop

Private Sub RunImport(ByVal folder As String)
    Dim files As New Collection, locs As New Collection
    Dim infos() As Variant, i As Long, j As Long, li As Long
    Dim savedInput As Variant, savedCyclic As Variant, savedS As Variant, savedT As Variant
    Dim calcMode As Long, thrMode As String, loc As String, basePath As String, msg As String
    Dim nOk As Long, nFail As Long, t0 As Double

    If Len(folder) = 0 Then folder = ChooseFolder()
    If Len(folder) = 0 Then Exit Sub
    folder = StripSep(folder)

    On Error Resume Next
    CollectFiles folder, files
    If Err.Number <> 0 Then
        Say "Could not read the folder " & folder & ": " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0
    If files.Count = 0 Then
        Say "No .epw or .csv files were found in " & folder & " or its sub-folders."
        Exit Sub
    End If

    thrMode = CStr(NR("import_thr_mode").Value)
    If LCase$(Left$(thrMode, 8)) = "settings" Then
        If IsBlank(NR("swcdh_override").Value) Or IsBlank(NR("twcdh_override").Value) Then
            Say "Set the SWCDH threshold and TWCDH offset overrides in Settings first, or choose " & _
                """Each location's TRY"" for the imported files' thresholds."
            Exit Sub
        End If
    End If

    savedInput = NR("hourly_input").Value
    savedCyclic = NR("cyclic").Value
    savedS = NR("swcdh_override").Value
    savedT = NR("twcdh_override").Value
    calcMode = Application.Calculation
    t0 = Timer

    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.Calculation = -4135   ' xlCalculationManual
    StartLog folder, files.Count

    ReDim infos(1 To files.Count)
    For i = 1 To files.Count
        infos(i) = FileMeta(CStr(files(i)), folder)
        AddUnique locs, CStr(infos(i)(1))
    Next i

    For li = 1 To locs.Count
        loc = CStr(locs(li))
        If LCase$(Left$(thrMode, 8)) <> "settings" Then
            basePath = BaselineFile(infos, loc)
            NR("swcdh_override").Value = savedS
            NR("twcdh_override").Value = savedT
            If Len(basePath) > 0 Then
                msg = ""
                If LoadFile(basePath, True, msg) Then
                    NR("swcdh_override").ClearContents
                    NR("twcdh_override").ClearContents
                    Application.Calculate
                    NR("swcdh_override").Value = Round(NR("swcdh_auto").Value, 3)
                    NR("twcdh_override").Value = Round(NR("twcdh_auto").Value, 3)
                    LogLine basePath, "Thresholds", loc, "", "", "", "", "", _
                        "SWCDH " & Format$(NR("swcdh_override").Value, "0.000") & " °C, TWCDH offset " & _
                        Format$(NR("twcdh_override").Value, "0.000") & " K, from this file"
                Else
                    LogLine basePath, "Thresholds failed", loc, "", "", "", "", "", msg & "; using Settings values"
                End If
            Else
                LogLine "", "Thresholds", loc, "", "", "", "", "", "No TRY or DSY1 found; using Settings values"
            End If
        End If
        For j = 1 To files.Count
            If CStr(infos(j)(1)) = loc Then
                If ImportOne(infos(j)) Then nOk = nOk + 1 Else nFail = nFail + 1
            End If
        Next j
    Next li

Finish:
    On Error Resume Next
    NR("hourly_input").Value = savedInput
    NR("cyclic").Value = savedCyclic
    NR("swcdh_override").Value = savedS
    NR("twcdh_override").Value = savedT
    Application.Calculation = calcMode
    Application.Calculate
    Application.ScreenUpdating = True
    LogLine "", "Finished", "", "", "", "", "", "", nOk & " imported, " & nFail & " not imported, " & _
        Format$(Timer - t0, "0") & " s"
    Say nOk & " file(s) added to the Library, " & nFail & " not imported. See the Import log sheet."
    Exit Sub

Failed:
    LogLine "", "Error", "", "", "", "", "", "", Err.Description
    nFail = nFail + 1
    Resume Finish
End Sub

Private Function ImportOne(info As Variant) As Boolean
    Dim msg As String, vals As Variant, r As Long, k As Long, row As Long, libRange As Range
    Dim outRow() As Variant

    If Not LoadFile(CStr(info(0)), CBool(info(6)), msg) Then
        LogLine CStr(info(0)), "Skipped", CStr(info(1)), CStr(info(2)), CStr(info(3)), CStr(info(4)), CStr(info(5)), "", msg
        Exit Function
    End If
    Application.Calculate
    vals = NR("summary_values").Value
    For k = 1 To NMetrics()
        If Not IsNumeric(vals(k, 1)) Or IsBlank(vals(k, 1)) Then
            LogLine CStr(info(0)), "Skipped", CStr(info(1)), CStr(info(2)), CStr(info(3)), CStr(info(4)), CStr(info(5)), "", _
                "Metrics could not be calculated (is a full season of data present?)"
            Exit Function
        End If
    Next k

    row = LibraryRow(info)
    If row = 0 Then
        LogLine CStr(info(0)), "Skipped", CStr(info(1)), CStr(info(2)), CStr(info(3)), CStr(info(4)), CStr(info(5)), "", _
            "The Library is full"
        Exit Function
    End If

    Set libRange = NR("lib_table")
    ReDim outRow(1 To 1, 1 To ColSource())
    outRow(1, 1) = 1
    outRow(1, 2) = info(1)
    outRow(1, 3) = info(2)
    outRow(1, 4) = info(3)
    outRow(1, 5) = BlankIfEmpty(info(4))
    If CStr(info(5)) = "" Then outRow(1, 6) = Empty Else outRow(1, 6) = CLng(info(5))
    For k = 1 To NMetrics()
        outRow(1, COL_METRICS + k - 1) = vals(k, 1)
    Next k
    outRow(1, ColSwcdh()) = NR("swcdh_thr").Value
    outRow(1, ColTwcdh()) = NR("twcdh_off").Value
    outRow(1, ColSource()) = CStr(info(0))
    libRange.Cells(row, 1).Resize(1, 6).Value = SliceRow(outRow, 1, 6)
    libRange.Cells(row, COL_METRICS).Resize(1, ColSource() - COL_METRICS + 1).Value = SliceRow(outRow, COL_METRICS, ColSource())
    LogLine CStr(info(0)), "Imported", CStr(info(1)), CStr(info(2)), CStr(info(3)), CStr(info(4)), CStr(info(5)), _
        CStr(libRange.Cells(row, 1).Row), ""
    ImportOne = True
End Function

' ------------------------------------------------------------------ reading weather files

Private Function LoadFile(ByVal path As String, ByVal typical As Boolean, ByRef msg As String) As Boolean
    Dim txt As String, lines As Variant, data() As Variant, n As Long
    On Error GoTo Bad
    txt = ReadAll(path)
    lines = Split(Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ReDim data(1 To MAX_HOURS, 1 To 5)
    If UCase$(Left$(Trim$(CStr(lines(0))), 8)) = "LOCATION" Then
        n = ParseEpw(lines, data, typical)
    Else
        n = ParseCsv(lines, data, typical, msg)
        If n < 0 Then Exit Function
    End If
    If n < MIN_HOURS Then
        msg = "Only " & n & " hours of data (at least " & MIN_HOURS & " needed)"
        Exit Function
    End If
    NR("hourly_input").Value = data
    NR("cyclic").Value = IIf(typical, "Yes", "No")
    LoadFile = True
    Exit Function
Bad:
    msg = "Could not read the file: " & Err.Description
End Function

Private Function ReadAll(ByVal path As String) As String
    ' Line Input copes with CRLF, CR and (as embedded characters) LF line endings; lines are
    ' collected in an array and joined once, which keeps large files fast.
    Dim f As Integer, buf() As String, n As Long, lineText As String
    ReDim buf(0 To 9999)
    f = FreeFile
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If n > UBound(buf) Then
            ReDim Preserve buf(0 To 2 * UBound(buf) + 1)
        End If
        buf(n) = lineText
        n = n + 1
    Loop
    Close #f
    If n = 0 Then Exit Function
    ReDim Preserve buf(0 To n - 1)
    ReadAll = Join(buf, vbLf)
End Function

Private Function ParseEpw(lines As Variant, data() As Variant, ByVal typical As Boolean) As Long
    Dim i As Long, n As Long, p As Variant, m As Long, d As Long, t As Double
    For i = 8 To UBound(lines)
        If Len(lines(i)) > 0 Then
            p = Split(lines(i), ",")
            If UBound(p) >= 6 Then
                m = CLng(Val(p(1)))
                d = CLng(Val(p(2)))
                t = Val(p(6))
                If t < 99.9 And Not (typical And m = 2 And d = 29) Then
                    n = n + 1
                    If n > MAX_HOURS Then
                        n = MAX_HOURS
                        Exit For
                    End If
                    If Not typical Then data(n, 1) = CLng(Val(p(0)))
                    data(n, 2) = m
                    data(n, 3) = d
                    data(n, 4) = CLng(Val(p(3)))
                    data(n, 5) = t
                End If
            End If
        End If
    Next i
    ParseEpw = n
End Function

Private Function ParseCsv(lines As Variant, data() As Variant, ByVal typical As Boolean, ByRef msg As String) As Long
    ' CSV with a header row containing Month, Day, Hour, an optional Year and a dry-bulb column.
    Dim i As Long, hdr As Long, sep As String, cells As Variant, c As Long, nm As String
    Dim cY As Long, cM As Long, cD As Long, cH As Long, cT As Long, n As Long, minH As Long, v As Variant

    hdr = -1
    For i = 0 To IIfLng(UBound(lines) < 80, UBound(lines), 80)
        sep = BestSep(CStr(lines(i)))
        cells = Split(LCase$(CStr(lines(i))), sep)
        cY = -1
        cM = -1
        cD = -1
        cH = -1
        cT = -1
        For c = 0 To UBound(cells)
            nm = CleanName(CStr(cells(c)))
            Select Case nm
                Case "year"
                    cY = c
                Case "month"
                    cM = c
                Case "day"
                    cD = c
                Case "hour"
                    cH = c
            End Select
            If cT < 0 And IsTempName(nm) Then cT = c
        Next c
        If cM >= 0 And cD >= 0 And cH >= 0 And cT >= 0 Then
            hdr = i
            Exit For
        End If
    Next i
    If hdr < 0 Then
        msg = "CSV needs a header row with Month, Day, Hour and a dry-bulb temperature column"
        ParseCsv = -1
        Exit Function
    End If
    If cY < 0 Then typical = True

    minH = 99
    For i = hdr + 1 To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) > 0 Then
            cells = Split(CStr(lines(i)), sep)
            If UBound(cells) >= MaxOf(cM, cD, cH, cT) Then
                v = Trim$(Replace(CStr(cells(cT)), """", ""))
                If LooksNumeric(CStr(v)) Then
                    If Not (typical And Val(cells(cM)) = 2 And Val(cells(cD)) = 29) Then
                        n = n + 1
                        If n > MAX_HOURS Then
                            n = MAX_HOURS
                            Exit For
                        End If
                        If Not typical Then data(n, 1) = CLng(Val(cells(cY)))
                        data(n, 2) = CLng(Val(cells(cM)))
                        data(n, 3) = CLng(Val(cells(cD)))
                        data(n, 4) = CLng(Val(cells(cH)))
                        data(n, 5) = Val(v)
                        If data(n, 4) < minH Then minH = data(n, 4)
                    End If
                End If
            End If
        End If
    Next i
    If minH = 0 Then   ' hours 0-23: convert to the 1-24 hour-ending convention used by the workbook
        For i = 1 To n
            data(i, 4) = data(i, 4) + 1
        Next i
    End If
    ParseCsv = n
End Function

Private Function LooksNumeric(ByVal s As String) As Boolean
    ' Locale-independent check for a plain decimal number such as -3.25 (Val is used to convert).
    Dim i As Long, ch As String, digits As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            digits = digits + 1
        ElseIf Not (ch = "." Or ((ch = "-" Or ch = "+") And i = 1)) Then
            Exit Function
        End If
    Next i
    LooksNumeric = digits > 0
End Function

Private Function BestSep(ByVal s As String) As String
    Dim nC As Long, nS As Long, nT As Long
    nC = UBound(Split(s, ","))
    nS = UBound(Split(s, ";"))
    nT = UBound(Split(s, vbTab))
    BestSep = ","
    If nS > nC And nS >= nT Then BestSep = ";"
    If nT > nC And nT > nS Then BestSep = vbTab
End Function

Private Function CleanName(ByVal s As String) As String
    Dim p As Long
    s = LCase$(Trim$(Replace(s, """", "")))
    p = InStr(s, "(")
    If p > 0 Then s = Trim$(Left$(s, p - 1))
    p = InStr(s, "[")
    If p > 0 Then s = Trim$(Left$(s, p - 1))
    CleanName = s
End Function

Private Function IsTempName(ByVal nm As String) As Boolean
    If InStr(nm, "dew") > 0 Or InStr(nm, "wet") > 0 Or InStr(nm, "soil") > 0 Then Exit Function
    IsTempName = (InStr(nm, "dry") > 0) Or nm = "dbt" Or nm = "db" Or nm = "tdb" Or nm = "temp" Or _
        nm = "temperature" Or nm = "air_temperature" Or nm = "air temperature" Or nm = "temperature_2m"
End Function

' ------------------------------------------------------------------ file names and folders

Private Function FileMeta(ByVal path As String, ByVal root As String) As Variant
    ' Array(path, location, kind, period, emissions, percentile, typical)
    Dim fname As String, parent As String, u As String, toks As Variant, i As Long, t As String
    Dim loc As String, kind As String, period As String, em As String, pct As String, pos As Long

    fname = FileName(path)
    parent = FileName(StripSep(Left$(path, Len(path) - Len(fname))))
    u = UCase$(fname)
    If InStr(u, ".") > 0 Then u = Left$(u, InStrRev(u, ".") - 1)
    u = Replace(Replace(Replace(u, "-", "_"), " ", "_"), ".", "_")
    toks = Split(u, "_")

    ' File type
    For i = 0 To UBound(toks)
        t = CStr(toks(i))
        If Left$(t, 3) = "DSY" Then
            If Mid$(t, 4, 1) Like "[123]" And Not (Mid$(t, 5, 1) Like "#") Then kind = "DSY" & Mid$(t, 4, 1) Else kind = "DSY1"
            Exit For
        ElseIf t = "TRY" Then
            kind = "TRY"
        End If
    Next i

    ' Period
    For i = 2 To 8
        If InStr(u, "20" & i & "0") > 0 Then
            period = "20" & i & "0s"
            Exit For
        End If
    Next i
    If period = "" And kind <> "" Then period = "Baseline"

    ' Emissions and percentile (only for future periods)
    If period <> "" And period <> "Baseline" Then
        For i = 0 To UBound(toks)
            t = CStr(toks(i))
            em = WordIn(t, "MEDIUM", pos)
            If em = "" Then em = WordIn(t, "HIGH", pos)
            If em = "" Then em = WordIn(t, "MED", pos)
            If em = "" Then em = WordIn(t, "LOW", pos)
            If em <> "" Then
                pct = LeadingDigits(Mid$(t, pos))
                If pct = "" And i < UBound(toks) Then pct = LeadingDigits(CStr(toks(i + 1)))
                Exit For
            End If
        Next i
        If em = "MED" Then em = "MEDIUM"
        If em <> "" Then em = Left$(em, 1) & LCase$(Mid$(em, 2))
        If pct = "" Then
            For i = 0 To UBound(toks)
                t = CStr(toks(i))
                If t Like "10P*" Or t Like "10TH*" Or t Like "50P*" Or t Like "50TH*" Or t Like "90P*" Or t Like "90TH*" Then
                    pct = Left$(t, 2)
                    Exit For
                End If
            Next i
        End If
        If pct <> "10" And pct <> "50" And pct <> "90" Then pct = ""
    End If

    ' Location: 2025 zone, 2016 CIBSE location, else the folder name
    For i = 0 To UBound(toks)
        t = CStr(toks(i))
        If (t Like "Z#" Or t Like "Z##") Then
            loc = "Zone " & CLng(Mid$(t, 2))
            Exit For
        End If
        If (t Like "ZONE#" Or t Like "ZONE##") Then
            loc = "Zone " & CLng(Mid$(t, 5))
            Exit For
        End If
        If t = "ZONE" And i < UBound(toks) Then
            If IsNumeric(toks(i + 1)) Then
                loc = "Zone " & CLng(toks(i + 1))
                Exit For
            End If
        End If
    Next i
    If loc = "" Then loc = CibseLocation(Replace(u, "_", ""))
    If loc = "" Then loc = parent
    If loc = "" Then loc = "Unknown"
    If kind = "" Then kind = "Other"
    If period = "" Then period = "Baseline"

    FileMeta = Array(path, loc, kind, period, em, pct, (kind <> "Other"))
End Function

Private Function WordIn(ByVal tok As String, ByVal word As String, ByRef afterPos As Long) As String
    ' word at the start of tok or right after a digit, followed by the end or a digit
    Dim p As Long, before As String, after As String
    p = InStr(tok, word)
    If p = 0 Then Exit Function
    If p > 1 Then before = Mid$(tok, p - 1, 1)
    after = Mid$(tok, p + Len(word), 1)
    If (before = "" Or before Like "#") And (after = "" Or after Like "#") Then
        WordIn = word
        afterPos = p + Len(word)
    End If
End Function

Private Function LeadingDigits(ByVal s As String) As String
    Dim i As Long
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit For
        LeadingDigits = LeadingDigits & Mid$(s, i, 1)
    Next i
End Function

Private Function CibseLocation(ByVal flat As String) As String
    Dim names As Variant, keys As Variant, i As Long
    keys = Array("BELFAST", "BIRMINGHAM", "CARDIFF", "EDINBURGH", "GLASGOW", "LEEDS", "LHR", "HEATHROW", "LWC", _
                 "WEATHERCENTRE", "LGW", "GATWICK", "MANCHESTER", "NEWCASTLE", "NORWICH", "NOTTINGHAM", "PLYMOUTH", _
                 "SOUTHAMPTON", "SWINDON")
    names = Array("Belfast", "Birmingham", "Cardiff", "Edinburgh", "Glasgow", "Leeds", "London (Heathrow)", _
                  "London (Heathrow)", "London (Weather Centre)", "London (Weather Centre)", "London (Gatwick)", _
                  "London (Gatwick)", "Manchester", "Newcastle", "Norwich", "Nottingham", "Plymouth", "Southampton", "Swindon")
    For i = 0 To UBound(keys)
        If InStr(flat, keys(i)) > 0 Then
            CibseLocation = names(i)
            Exit Function
        End If
    Next i
    If InStr(flat, "LONDON") > 0 Then CibseLocation = "London (Heathrow)"
End Function

Private Function BaselineFile(infos() As Variant, ByVal loc As String) As String
    ' The location's TRY with the earliest period (50th percentile preferred), else its DSY1.
    Dim k As Long, i As Long, best As Long, bestScore As Long, score As Long, kinds As Variant
    kinds = Array("TRY", "DSY1")
    For k = 0 To 1
        best = 0
        bestScore = 1000000
        For i = LBound(infos) To UBound(infos)
            If CStr(infos(i)(1)) = loc And CStr(infos(i)(2)) = kinds(k) Then
                score = PeriodRank(CStr(infos(i)(3))) * 100 + IIfLng(CStr(infos(i)(5)) = "50" Or CStr(infos(i)(5)) = "", 0, 1)
                If score < bestScore Then
                    bestScore = score
                    best = i
                End If
            End If
        Next i
        If best > 0 Then
            BaselineFile = CStr(infos(best)(0))
            Exit Function
        End If
    Next k
End Function

Private Function PeriodRank(ByVal period As String) As Long
    Select Case period
        Case "Baseline"
            PeriodRank = 0
        Case "2020s"
            PeriodRank = 1
        Case "2030s"
            PeriodRank = 2
        Case "2050s"
            PeriodRank = 3
        Case "2080s"
            PeriodRank = 4
        Case Else
            PeriodRank = 9
    End Select
End Function

Private Sub CollectFiles(ByVal folder As String, files As Collection)
    ' Recursive listing with Dir (works in Excel for Windows and Mac).
    Dim sep As String, nm As String, subs As New Collection, item As Variant, ext As String, isDir As Boolean
    sep = PathSep(folder)
    folder = StripSep(folder) & sep
    nm = Dir$(folder & "*", vbDirectory)
    Do While Len(nm) > 0
        If nm <> "." And nm <> ".." And Left$(nm, 1) <> "." And Left$(nm, 2) <> "~$" Then
            isDir = False
            On Error Resume Next
            isDir = ((GetAttr(folder & nm) And vbDirectory) = vbDirectory)
            On Error GoTo 0
            If isDir Then
                subs.Add folder & nm
            Else
                ext = LCase$(Mid$(nm, InStrRev(nm, ".") + 1))
                If ext = "epw" Or ext = "csv" Then files.Add folder & nm
            End If
        End If
        nm = Dir$()
    Loop
    For Each item In subs
        CollectFiles CStr(item), files
    Next item
End Sub

Private Function ChooseFolder() As String
    Dim p As String
    p = Trim$(CStr(NR("import_folder").Value))
    If Len(p) > 0 Then
        ChooseFolder = p
        Exit Function
    End If
    On Error GoTo NoDialog
    With Application.FileDialog(4)   ' msoFileDialogFolderPicker
        .Title = "Choose the folder that contains the weather files"
        If .Show = -1 Then ChooseFolder = .SelectedItems(1)
    End With
    Exit Function
NoDialog:
    ChooseFolder = InputBox("Folder that contains the weather files (sub-folders are included):", "Import weather files")
End Function

' ------------------------------------------------------------------ Library and log

Private Function LibraryRow(info As Variant) As Long
    ' Row (within lib_table) of the entry with the same identity, else the first empty row.
    Dim tbl As Variant, r As Long, firstEmpty As Long
    tbl = NR("lib_table").Value
    For r = 1 To UBound(tbl, 1)
        If IsBlank(tbl(r, 2)) And IsBlank(tbl(r, 3)) And IsBlank(tbl(r, COL_METRICS)) Then
            If firstEmpty = 0 Then firstEmpty = r
        ElseIf LCase$(CStr(tbl(r, 2))) = LCase$(CStr(info(1))) And LCase$(CStr(tbl(r, 3))) = LCase$(CStr(info(2))) And _
               LCase$(CStr(tbl(r, 4))) = LCase$(CStr(info(3))) And LCase$(CStr(tbl(r, 5))) = LCase$(CStr(info(4))) And _
               CStr(tbl(r, 6)) = CStr(info(5)) Then
            LibraryRow = r
            Exit Function
        End If
    Next r
    LibraryRow = firstEmpty
End Function

Private Sub StartLog(ByVal folder As String, ByVal nFiles As Long)
    Dim ws As Worksheet
    Set ws = NR("log_start").Worksheet
    ws.Range(ws.Cells(NR("log_start").Row, 1), ws.Cells(ws.Rows.Count, 10)).ClearContents
    ws.Cells(2, 1).Value = "Last import: " & folder & " (" & nFiles & " weather files found) at " & Format$(Now, "yyyy-mm-dd hh:nn")
    mLogRow = NR("log_start").Row
End Sub

Private Sub LogLine(ByVal path As String, ByVal result As String, ByVal loc As String, ByVal kind As String, _
                    ByVal period As String, ByVal em As String, ByVal pct As String, ByVal libRow As String, ByVal note As String)
    Dim ws As Worksheet
    If mLogRow = 0 Then Exit Sub
    Set ws = NR("log_start").Worksheet
    ws.Range(ws.Cells(mLogRow, 1), ws.Cells(mLogRow, 9)).Value = Array(path, result, loc, kind, period, em, pct, libRow, note)
    mLogRow = mLogRow + 1
End Sub

' ------------------------------------------------------------------ helpers

Private Function NMetrics() As Long
    NMetrics = NR("summary_values").Rows.Count
End Function

Private Function ColSwcdh() As Long
    ColSwcdh = COL_METRICS + NMetrics()
End Function

Private Function ColTwcdh() As Long
    ColTwcdh = ColSwcdh() + 1
End Function

Private Function ColSource() As Long
    ColSource = ColSwcdh() + 2
End Function

Private Function NR(ByVal nm As String) As Range
    Set NR = ThisWorkbook.Names(nm).RefersToRange
End Function

Private Sub Say(ByVal msg As String)
    If mSilent Then Exit Sub
    MsgBox msg, vbInformation, "Compare thee to a summer's DSY"
End Sub

Private Function IsBlank(v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlank = True
        Exit Function
    End If
    If IsError(v) Then Exit Function
    IsBlank = (Len(Trim$(CStr(v))) = 0)
End Function

Private Function BlankIfEmpty(v As Variant) As Variant
    If Len(CStr(v)) = 0 Then BlankIfEmpty = Empty Else BlankIfEmpty = v
End Function

Private Function SliceRow(arr() As Variant, ByVal c1 As Long, ByVal c2 As Long) As Variant
    Dim out() As Variant, c As Long
    ReDim out(1 To 1, 1 To c2 - c1 + 1)
    For c = c1 To c2
        out(1, c - c1 + 1) = arr(1, c)
    Next c
    SliceRow = out
End Function

Private Sub AddUnique(col As Collection, ByVal s As String)
    Dim item As Variant
    For Each item In col
        If CStr(item) = s Then Exit Sub
    Next item
    col.Add s
End Sub

Private Function PathSep(ByVal folder As String) As String
    If InStr(folder, "\") > 0 Then PathSep = "\" Else PathSep = "/"
End Function

Private Function StripSep(ByVal folder As String) As String
    Do While Len(folder) > 1 And (Right$(folder, 1) = "\" Or Right$(folder, 1) = "/")
        folder = Left$(folder, Len(folder) - 1)
    Loop
    StripSep = folder
End Function

Private Function FileName(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(Replace(path, "\", "/"), "/")
    FileName = Mid$(path, p + 1)
End Function

Private Function IIfLng(ByVal cond As Boolean, ByVal a As Long, ByVal b As Long) As Long
    If cond Then IIfLng = a Else IIfLng = b
End Function

Private Function MaxOf(ParamArray v() As Variant) As Long
    Dim i As Long
    MaxOf = v(0)
    For i = 1 To UBound(v)
        If v(i) > MaxOf Then MaxOf = v(i)
    Next i
End Function
