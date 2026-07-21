Option Explicit
' Config constants and all helper functions live in KPI_Common.

' ================== MACRO: BUILD / REFRESH TABLE ==================
Public Sub BuildKPITable()
    Dim prevSU As Boolean, prevEv As Boolean
    prevSU = Application.ScreenUpdating: prevEv = Application.EnableEvents
    On Error GoTo CleanFail
    Application.ScreenUpdating = False: Application.EnableEvents = False

    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Set wsSrc = SheetOrNothing(SRC_SHEET)
    Set wsTgt = SheetOrNothing(TGT_SHEET)
    If wsSrc Is Nothing Then MsgBox "Source sheet '" & SRC_SHEET & "' not found.", vbExclamation: GoTo CleanExit
    If wsTgt Is Nothing Then MsgBox "Target sheet '" & TGT_SHEET & "' not found.", vbExclamation: GoTo CleanExit

    Dim ids As Collection
    Set ids = CollectAllIds(wsSrc)
    If ids.Count = 0 Then MsgBox "No stopes found in source.", vbInformation: GoTo CleanExit

    Dim lo As ListObject
    Set lo = GetOrCreateTable(wsTgt, ids)   ' creates seeded with all ids if new

    ' Append any ids not already present
    Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        Dim idBody As Variant, r As Long
        idBody = lo.ListColumns(H_ID).DataBodyRange.Value
        If IsArray(idBody) Then
            For r = 1 To UBound(idBody, 1)
                Dim v As String: v = CleanStr(idBody(r, 1))
                If Len(v) > 0 Then existing(v) = True
            Next r
        ElseIf Len(CleanStr(idBody)) > 0 Then
            existing(CleanStr(idBody)) = True
        End If
    End If

    Dim added As Long, k As Long
    For k = 1 To ids.Count
        If Not existing.Exists(CStr(ids(k))) Then
            lo.ListRows.Add.Range.Cells(1, 1).Value = ids(k)
            added = added + 1
        End If
    Next k

    ApplyLookupFormulas lo
    Application.Calculate
    FilterOutSkipped lo   ' hide GREEN/YELLOW and IFR-stage rows; other BLACK/RED stay visible
    MsgBox "KPI table ready. " & added & " new stope(s) added. Total rows: " & _
           IIf(lo.DataBodyRange Is Nothing, 0, lo.ListRows.Count) & _
           " (GREEN/YELLOW and IFR hidden by filter).", vbInformation

CleanExit:
    Application.ScreenUpdating = prevSU: Application.EnableEvents = prevEv
    Exit Sub
CleanFail:
    MsgBox "BuildKPITable error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
