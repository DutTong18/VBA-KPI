Option Explicit
' Config constants and all helper functions live in KPI_Common.

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

    ' Register any new stage keys (skip GREEN/YELLOW)
    Dim orderChanged As Boolean, key As String
    For i = 1 To n
        If Not IsSkippableZone(CStr(kpi(i, 5))) Then
            key = StageKey(kpi(i, 3), kpi(i, 4))
            If Not stageIdx.Exists(key) Then
                stageIdx(key) = order.Count
                order.Add key
                orderChanged = True
            End If
        End If
    Next i

    ' ---- Evaluate each stope ----
    Dim results() As String: ReDim results(1 To n)
    Dim stateArr() As Variant: ReDim stateArr(1 To n, 1 To 3)
    Dim userStats As Object: Set userStats = CreateObject("Scripting.Dictionary")
    Dim passCount As Long, nCount As Long, cBlack As Long, cRed As Long, skipCount As Long

    For i = 1 To n
        Dim id As String, usr As String, stg As String, sub_ As String, zone As String
        id = CleanStr(kpi(i, 1))
        usr = CleanStr(kpi(i, 2)): If Len(usr) = 0 Then usr = "(unassigned)"
        stg = CleanStr(kpi(i, 3)): sub_ = CleanStr(kpi(i, 4))
        zone = UCase(CleanStr(kpi(i, 5)))

        If IsSkippableZone(zone) Then
            results(i) = ""          ' skipped: no grade, no state, no tally
            skipCount = skipCount + 1
        Else
            If zone = "BLACK" Then
                cBlack = cBlack + 1
            ElseIf zone = "RED" Then
                cRed = cRed + 1
            End If

            stateArr(i, 1) = id: stateArr(i, 2) = stg: stateArr(i, 3) = sub_

            results(i) = GradeRow(savedState, stageIdx, id, stg, sub_, zone)
            Select Case results(i)
                Case "Y": passCount = passCount + 1: BumpUser userStats, usr, 0
                Case "N": nCount = nCount + 1: BumpUser userStats, usr, 1
            End Select
        End If
    Next i

    ' ---- Append dated status column (Y/N/new/?) ----
    Dim outArr() As Variant: ReDim outArr(1 To n, 1 To 1)
    For i = 1 To n: outArr(i, 1) = results(i): Next i
    Dim dc As ListColumn: Set dc = AppendColumn(lo, baseHeader, outArr)
    For i = 1 To n
        If Len(results(i)) > 0 Then FormatStatusCell dc.DataBodyRange.Cells(i, 1), results(i)
    Next i

    ' ---- Append dated comments snapshot (Stope Cadence col X) ----
    Dim cmtMap As Object: Set cmtMap = ReadCommentsMap(SheetOrNothing(SRC_SHEET))
    Dim cArr() As Variant: ReDim cArr(1 To n, 1 To 1)
    For i = 1 To n
        Dim rid As String: rid = CleanStr(kpi(i, 1))
        If cmtMap.Exists(rid) Then cArr(i, 1) = cmtMap(rid) Else cArr(i, 1) = ""
    Next i
    Dim cc As ListColumn: Set cc = AppendColumn(lo, baseHeader, cArr)
    For i = 1 To n
        If Len(results(i)) > 0 Then
            With cc.DataBodyRange.Cells(i, 1)
                .Interior.Color = dc.DataBodyRange.Cells(i, 1).Interior.Color
                .Font.Color = dc.DataBodyRange.Cells(i, 1).Font.Color
            End With
        End If
    Next i

    ' ---- Summary + user breakdown + persist ----
    WriteSummaryBlock wsTgt, SUMMARY_ANCHOR, n, cBlack, cRed, Now
    WriteUserBreakdown wsTgt, SUMMARY_ANCHOR, userStats, passCount, nCount, lo.Range.Row
    WriteSavedState stateWs, stateArr, n
    If orderChanged Then WriteStageOrder stateWs, order

    MsgBox "Done - '" & baseHeader & "' added." & vbCrLf & _
           passCount & " passed, " & nCount & " non-progressions across " & _
           (n - skipCount) & " graded stopes (" & skipCount & " GREEN/YELLOW skipped)." & vbCrLf & _
           "BLACK=" & cBlack & "  RED=" & cRed, vbInformation

CleanExit:
    Application.ScreenUpdating = prevSU: Application.EnableEvents = prevEv
    Exit Sub
CleanFail:
    MsgBox "RunStatusCheck error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
