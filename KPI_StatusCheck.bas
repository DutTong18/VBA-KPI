Option Explicit
' Config constants and SheetOrNothing/CleanStr/ReadColumn live in KPI_Common.

' ================== MACRO: DAILY STATUS CHECK ==================
Public Sub RunStatusCheck()
    Dim prevSU As Boolean, prevEv As Boolean
    prevSU = Application.ScreenUpdating: prevEv = Application.EnableEvents
    On Error GoTo CleanFail
    Application.ScreenUpdating = False: Application.EnableEvents = False

    Dim wsTgt As Worksheet: Set wsTgt = SheetOrNothing(TGT_SHEET)
    If wsTgt Is Nothing Then MsgBox "Target sheet '" & TGT_SHEET & "' not found.", vbExclamation: GoTo CleanExit

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsTgt.ListObjects(TBL_NAME)
    On Error GoTo CleanFail
    If lo Is Nothing Then MsgBox "KPI table not found. Run BuildKPITable first.", vbExclamation: GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then MsgBox "KPI table has no rows. Run BuildKPITable first.", vbExclamation: GoTo CleanExit

    Dim baseHeader As String: baseHeader = Format(Date, "dd-mm-yyyy")

    ' ---- DOUBLE-RUN GUARD ----
    If TodayColumnExists(lo, baseHeader) Then
        MsgBox "Status for today (" & baseHeader & ") is already recorded." & vbCrLf & _
               "Aborting to protect the baseline. Delete today's column manually to re-run.", vbExclamation
        GoTo CleanExit
    End If

    Application.Calculate   ' resolve lookup formulas before reading

    ' ---- Load cache ----
    Dim stateWs As Worksheet: Set stateWs = GetStateSheet()
    Dim savedState As Object: Set savedState = ReadSavedState(stateWs)
    Dim order As Collection: Set order = ReadStageOrder(stateWs)
    Dim stageIdx As Object: Set stageIdx = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To order.Count: stageIdx(CStr(order(i))) = i - 1: Next i

    ' ---- Read KPI data ----
    Dim kpi As Variant: kpi = lo.DataBodyRange.Value
    Dim n As Long: n = UBound(kpi, 1)

    ' Register any new stage keys
    Dim orderChanged As Boolean, key As String
    For i = 1 To n
        key = StageKey(kpi(i, 3), kpi(i, 4))
        If Not stageIdx.Exists(key) Then
            stageIdx(key) = order.Count
            order.Add key
            orderChanged = True
        End If
    Next i

    ' ---- Evaluate each stope ----
    Dim results() As String: ReDim results(1 To n)
    Dim stateArr() As Variant: ReDim stateArr(1 To n, 1 To 3)
    Dim userStats As Object: Set userStats = CreateObject("Scripting.Dictionary")
    Dim passCount As Long, nCount As Long, cBlack As Long, cRed As Long

    For i = 1 To n
        Dim id As String, usr As String, stg As String, sub_ As String, zone As String
        id = CleanStr(kpi(i, 1))
        usr = CleanStr(kpi(i, 2)): If Len(usr) = 0 Then usr = "(unassigned)"
        stg = CleanStr(kpi(i, 3)): sub_ = CleanStr(kpi(i, 4))
        zone = UCase(CleanStr(kpi(i, 5)))
        If zone = "BLACK" Then cBlack = cBlack + 1 ElseIf zone = "RED" Then cRed = cRed + 1

        stateArr(i, 1) = id: stateArr(i, 2) = stg: stateArr(i, 3) = sub_

        results(i) = GradeRow(savedState, stageIdx, id, stg, sub_, zone)
        Select Case results(i)
            Case "Y": passCount = passCount + 1: BumpUser userStats, usr, 0
            Case "N": nCount = nCount + 1: BumpUser userStats, usr, 1
        End Select
    Next i

    ' ---- Append dated status column (Y/N/new/?) ----
    Dim outArr() As Variant: ReDim outArr(1 To n, 1 To 1)
    For i = 1 To n: outArr(i, 1) = results(i): Next i
    Dim dc As ListColumn: Set dc = AppendColumn(lo, baseHeader, outArr)
    For i = 1 To n: FormatStatusCell dc.DataBodyRange.Cells(i, 1), results(i): Next i

    ' ---- Append dated comments snapshot (Stope Cadence col X) ----
    Dim cmtMap As Object: Set cmtMap = ReadCommentsMap(SheetOrNothing(SRC_SHEET))
    Dim cArr() As Variant: ReDim cArr(1 To n, 1 To 1)
    For i = 1 To n
        Dim rid As String: rid = CleanStr(kpi(i, 1))
        If cmtMap.Exists(rid) Then cArr(i, 1) = cmtMap(rid) Else cArr(i, 1) = ""
    Next i
    AppendColumn lo, baseHeader, cArr

    ' ---- Summary + user breakdown + persist ----
    WriteSummaryBlock wsTgt, SUMMARY_ANCHOR, n, cBlack, cRed, Now
    WriteUserBreakdown wsTgt, SUMMARY_ANCHOR, userStats, passCount, nCount, lo.Range.Row
    WriteSavedState stateWs, stateArr, n
    If orderChanged Then WriteStageOrder stateWs, order

    MsgBox "Done - '" & baseHeader & "' added." & vbCrLf & _
           passCount & " passed, " & nCount & " non-progressions across " & n & " stopes." & vbCrLf & _
           "BLACK=" & cBlack & "  RED=" & cRed, vbInformation

CleanExit:
    Application.ScreenUpdating = prevSU: Application.EnableEvents = prevEv
    Exit Sub
CleanFail:
    MsgBox "RunStatusCheck error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' ================== FUNCTIONS ==================
Private Function GradeRow(savedState As Object, stageIdx As Object, _
                          id As String, stg As String, sub_ As String, zone As String) As String
    If Not savedState.Exists(id) Then GradeRow = "new": Exit Function
    Dim prev As Variant: prev = savedState(id)
    Dim currI As Long: currI = StagePos(stageIdx, StageKey(stg, sub_))
    Dim prevI As Long: prevI = StagePos(stageIdx, StageKey(prev(0), prev(1)))
    If currI = -1 Or prevI = -1 Then GradeRow = "?": Exit Function
    Dim thr As Long: thr = IIf(zone = "RED", STEPS_RED, STEPS_BLACK)
    GradeRow = IIf((currI - prevI) >= thr, "Y", "N")
End Function

Private Function StagePos(stageIdx As Object, key As String) As Long
    If stageIdx.Exists(key) Then StagePos = stageIdx(key) Else StagePos = -1
End Function

Private Function AppendColumn(lo As ListObject, headerName As String, values As Variant) As ListColumn
    Dim col As ListColumn: Set col = lo.ListColumns.Add
    col.Name = UniqueColName(lo, headerName)
    col.DataBodyRange.Value = values
    Set AppendColumn = col
End Function

Private Function TodayColumnExists(lo As ListObject, baseHeader As String) As Boolean
    Dim col As ListColumn
    For Each col In lo.ListColumns
        If col.Name = baseHeader Then TodayColumnExists = True: Exit Function
        If col.Name Like baseHeader & " (*)" Then TodayColumnExists = True: Exit Function
    Next col
End Function

Private Function UniqueColName(lo As ListObject, desired As String) As String
    Dim nm As String, k As Long
    nm = desired: k = 2
    Do While ColumnNameExists(lo, nm)
        nm = desired & " (" & k & ")": k = k + 1
    Loop
    UniqueColName = nm
End Function

Private Function ColumnNameExists(lo As ListObject, nm As String) As Boolean
    Dim col As ListColumn
    For Each col In lo.ListColumns
        If col.Name = nm Then ColumnNameExists = True: Exit Function
    Next col
End Function

Private Function ReadCommentsMap(wsSrc As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    If wsSrc Is Nothing Then Set ReadCommentsMap = d: Exit Function
    Dim lastRow As Long: lastRow = wsSrc.Cells(wsSrc.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < 2 Then Set ReadCommentsMap = d: Exit Function
    Dim idv As Variant, cv As Variant, r As Long, id As String
    idv = ReadColumn(wsSrc, COL_ID, 2, lastRow)
    cv = ReadColumn(wsSrc, COL_COMMENTS, 2, lastRow)
    For r = 1 To UBound(idv, 1)
        id = CleanStr(idv(r, 1))
        If Len(id) > 0 Then
            If Not d.Exists(id) Then d(id) = CStr(cv(r, 1))   ' first occurrence, mirrors MATCH
        End If
    Next r
    Set ReadCommentsMap = d
End Function

Private Sub FormatStatusCell(cell As Range, v As String)
    Select Case v
        Case "Y":   cell.Interior.Color = RGB(198, 239, 206): cell.Font.Color = RGB(39, 98, 33)
        Case "N":   cell.Interior.Color = RGB(255, 199, 206): cell.Font.Color = RGB(156, 0, 6)
        Case "new": cell.Interior.Color = RGB(221, 235, 247): cell.Font.Color = RGB(31, 78, 121): cell.Font.Italic = True
        Case Else:  cell.Interior.Color = RGB(255, 235, 156): cell.Font.Color = RGB(156, 87, 0)
    End Select
    cell.HorizontalAlignment = xlCenter
End Sub

Private Function StageKey(ds As Variant, sp As Variant) As String
    StageKey = CleanStr(ds) & "::" & CleanStr(sp)
End Function

' ---- Cache sheet ----
Private Function GetStateSheet() As Worksheet
    Dim ws As Worksheet: Set ws = SheetOrNothing(STATE_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = STATE_SHEET
        ws.Range("A1:C1").Value = Array("StopeID", "DesignStage", "SubProcess")
        ws.Range("E1").Value = "StageOrder"
        ws.Visible = xlSheetVeryHidden
    End If
    Set GetStateSheet = ws
End Function

Private Function ReadSavedState(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim last As Long: last = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long, id As String
    For i = 2 To last
        id = CleanStr(ws.Cells(i, 1).Value)
        If Len(id) > 0 And id <> "undefined" Then
            d(id) = Array(CleanStr(ws.Cells(i, 2).Value), CleanStr(ws.Cells(i, 3).Value))
        End If
    Next i
    Set ReadSavedState = d
End Function

Private Sub WriteSavedState(ws As Worksheet, arr As Variant, n As Long)
    ws.Range("A2:C" & ws.Rows.Count).ClearContents
    If n > 0 Then ws.Range("A2").Resize(n, 3).Value = arr
End Sub

Private Function ReadStageOrder(ws As Worksheet) As Collection
    Dim c As New Collection
    Dim last As Long: last = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    Dim i As Long, k As String
    For i = 2 To last
        k = CleanStr(ws.Cells(i, 5).Value)
        If Len(k) > 0 Then c.Add k
    Next i
    If c.Count = 0 Then Set c = SeedOrder()
    Set ReadStageOrder = c
End Function

Private Sub WriteStageOrder(ws As Worksheet, order As Collection)
    ws.Range("E2:E" & ws.Rows.Count).ClearContents
    Dim i As Long
    For i = 1 To order.Count: ws.Cells(1 + i, 5).Value = order(i): Next i
End Sub

Private Function SeedOrder() As Collection
    Dim c As New Collection, a As Variant, x As Variant
    a = Array( _
        "Re-Sequenced::", "Not_Started::", "Geology_Review::", _
        "Draft_Design::", "Draft_Design::0%", "Draft_Design::25%", "Draft_Design::50%", "Draft_Design::75%", _
        "External_Review::", "External_Review::0%", "External_Review::25%", "External_Review::50%", "External_Review::75%", _
        "Concept::", _
        "Shape_Review::", "Shape_Review::Wait on Meeting", "Shape_Review::Updates Post ISR", _
        "Final_Design::", "Final_Design::0%", "Final_Design::25%", "Final_Design::50%", "Final_Design::75%", _
        "Peer_Review::", "Peer_Review::Wait on Meeting", "Peer_Review::Updates Post PR", _
        "Direction_Meeting::", "Direction_Meeting::Wait on Meeting", "Direction_Meeting::Updates Post meeting", _
        "IFR::", "02_Technical::", "03_Operations::", "04_Superintendent::", "05_Manager::", "06_Upload::", "COMPLETE::")
    For Each x In a: c.Add CStr(x): Next x
    Set SeedOrder = c
End Function

' ---- Summary blocks ----
Private Sub WriteSummaryBlock(ws As Worksheet, anchor As String, total As Long, black As Long, red As Long, runDate As Date)
    Dim a As Range: Set a = ws.Range(anchor)
    Dim r0 As Long, c0 As Long: r0 = a.Row: c0 = a.Column
    Dim labels As Variant, lf As Variant, lfont As Variant, counts As Variant, vf As Variant, vfont As Variant, c As Long, x As Range
    labels = Array("Total Stopes", "BLACK", "RED")
    lf = Array(RGB(217, 225, 242), RGB(0, 0, 0), RGB(255, 0, 0))
    lfont = Array(RGB(0, 0, 0), RGB(255, 255, 255), RGB(255, 255, 255))
    For c = 0 To 2
        Set x = ws.Cells(r0, c0 + c)
        x.Value = labels(c): x.Interior.Color = lf(c): x.Font.Color = lfont(c)
        x.Font.Bold = True: x.HorizontalAlignment = xlCenter
    Next c
    counts = Array(total, black, red)
    vf = Array(RGB(235, 240, 250), RGB(89, 89, 89), RGB(255, 224, 224))
    vfont = Array(RGB(0, 0, 0), RGB(255, 255, 255), RGB(156, 0, 6))
    For c = 0 To 2
        Set x = ws.Cells(r0 + 1, c0 + c)
        x.Value = counts(c): x.Interior.Color = vf(c): x.Font.Color = vfont(c)
        x.Font.Bold = True: x.HorizontalAlignment = xlCenter
    Next c
    Set x = ws.Cells(r0 + 2, c0)
    x.Value = "Last updated: " & Format(runDate, "dd/mm/yyyy hh:nn:ss")
    x.Font.Italic = True: x.Font.Color = RGB(89, 89, 89)
End Sub

Private Sub BumpUser(d As Object, user As String, idx As Long)
    Dim a As Variant
    If d.Exists(user) Then a = d(user) Else a = Array(0, 0)   ' (0)=Progressions, (1)=Non-Progressions
    a(idx) = a(idx) + 1
    d(user) = a
End Sub

Private Sub WriteUserBreakdown(ws As Worksheet, anchor As String, stats As Object, totalY As Long, totalN As Long, tableTopRow As Long)
    Dim a As Range: Set a = ws.Range(anchor)
    Dim r0 As Long, c0 As Long: r0 = a.Row + 4: c0 = a.Column
    Dim lastFree As Long: lastFree = tableTopRow - 1   ' never touch the KPI table below

    Dim clearRows As Long: clearRows = lastFree - r0 + 1
    If clearRows > 0 Then
        With ws.Cells(r0, c0).Resize(clearRows, 3)
            .ClearContents: .Interior.ColorIndex = xlNone
        End With
    End If

    With ws.Cells(r0, c0)
        .Value = "Progress by User (this run: " & totalY & " Y / " & totalN & " N)"
        .Font.Bold = True: .Interior.Color = RGB(217, 225, 242)
    End With

    Dim hdr As Variant: hdr = Array("User", "Progressions", "Non-Progressions")
    Dim c As Long
    For c = 0 To 2
        With ws.Cells(r0 + 1, c0 + c)
            .Value = hdr(c): .Font.Bold = True
            .Interior.Color = RGB(217, 225, 242): .HorizontalAlignment = xlCenter
        End With
    Next c

    If stats.Count = 0 Then
        With ws.Cells(r0 + 2, c0)
            .Value = "(none)": .Font.Italic = True: .Font.Color = RGB(89, 89, 89)
        End With
        Exit Sub
    End If

    Dim keys() As String: keys = SortUsersByN(stats)
    Dim i As Long, rowAt As Long, v As Variant
    For i = 0 To UBound(keys)
        rowAt = r0 + 2 + i
        If rowAt > lastFree Then
            ws.Cells(lastFree, c0).Value = "... (+" & (UBound(keys) - i + 1) & " more users)"
            ws.Cells(lastFree, c0).Font.Italic = True
            Exit For
        End If
        v = stats(keys(i))
        ws.Cells(rowAt, c0).Value = keys(i)
        ws.Cells(rowAt, c0 + 1).Resize(1, 2).Value = Array(v(0), v(1))
        ws.Cells(rowAt, c0 + 1).Resize(1, 2).HorizontalAlignment = xlCenter
    Next i
End Sub

Private Function SortUsersByN(d As Object) As String()
    Dim keys() As String, n As Long, i As Long, j As Long, idx As Long, kk As Variant
    n = d.Count: ReDim keys(0 To n - 1)
    idx = 0
    For Each kk In d.Keys: keys(idx) = CStr(kk): idx = idx + 1: Next kk
    ' Sort by Non-Progressions desc, then Progressions desc
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If UserRank(d, keys(j + 1)) > UserRank(d, keys(j)) Then
                Dim tmp As String: tmp = keys(j): keys(j) = keys(j + 1): keys(j + 1) = tmp
            End If
        Next j
    Next i
    SortUsersByN = keys
End Function

Private Function UserRank(d As Object, user As String) As Double
    Dim v As Variant: v = d(user)
    UserRank = v(1) * 1000000# + v(0)   ' Non-Progressions dominant, Progressions as tiebreak
End Function
