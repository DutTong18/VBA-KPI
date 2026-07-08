Attribute VB_Name = "KPI_Build"
Option Explicit

' ================== CONFIG ==================
Private Const SRC_SHEET     As String = "Stope Cadence"
Private Const TGT_SHEET     As String = "SchedulerData"
Private Const TBL_NAME      As String = "KPI"
Private Const TABLE_ANCHOR  As String = "A25"

' Source column numbers (1 based)
Private Const COL_ID    As Long = 3    ' C
Private Const COL_USER  As Long = 8    ' H
Private Const COL_STAGE As Long = 22   ' V
Private Const COL_SUB   As Long = 23   ' W
Private Const COL_ZONE  As Long = 35   ' AI

' Base header names
Private Const H_ID    As String = "ID Stope"
Private Const H_USER  As String = "User"
Private Const H_STAGE As String = "Design Stage"
Private Const H_SUB   As String = "Sub Process"
Private Const H_ZONE  As String = "Zone Status"

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
    Set ids = CollectBlackRedIds(wsSrc)
    If ids.Count = 0 Then MsgBox "No BLACK or RED stopes found in source.", vbInformation: GoTo CleanExit

    Dim lo As ListObject
    Set lo = GetOrCreateTable(wsTgt, ids)   ' creates seeded with all ids if new

    ' Append any BLACK/RED ids not already present
    Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        Dim idBody As Variant, r As Long
        idBody = lo.ListColumns(H_ID).DataBodyRange.Value
        If IsArray(idBody) Then
            For r = 1 To UBound(idBody, 1)
                Dim v As String: v = Trim(CStr(idBody(r, 1)))
                If Len(v) > 0 Then existing(v) = True
            Next r
        ElseIf Len(Trim(CStr(idBody))) > 0 Then
            existing(Trim(CStr(idBody))) = True
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
    MsgBox "KPI table ready. " & added & " new stope(s) added. Total rows: " & _
           IIf(lo.DataBodyRange Is Nothing, 0, lo.ListRows.Count) & ".", vbInformation

CleanExit:
    Application.ScreenUpdating = prevSU: Application.EnableEvents = prevEv
    Exit Sub
CleanFail:
    MsgBox "BuildKPITable error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' ================== HELPERS ==================
Private Function SheetOrNothing(nm As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Private Function CollectBlackRedIds(wsSrc As Worksheet) As Collection
    Dim c As New Collection
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long: lastRow = wsSrc.Cells(wsSrc.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < 2 Then Set CollectBlackRedIds = c: Exit Function

    Dim idv As Variant, zv As Variant
    idv = ReadColumn(wsSrc, COL_ID, 2, lastRow)
    zv = ReadColumn(wsSrc, COL_ZONE, 2, lastRow)

    Dim r As Long, id As String, zone As String
    For r = 1 To UBound(idv, 1)
        id = Trim(CStr(idv(r, 1)))
        zone = UCase(Trim(CStr(zv(r, 1))))
        If Len(id) > 0 Then
            If (zone = "BLACK" Or zone = "RED") And Not seen.Exists(id) Then
                seen(id) = True
                c.Add id
            End If
        End If
    Next r
    Set CollectBlackRedIds = c
End Function

Private Function ReadColumn(ws As Worksheet, col As Long, r1 As Long, r2 As Long) As Variant
    Dim v As Variant
    v = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Value
    If Not IsArray(v) Then
        Dim a(1 To 1, 1 To 1) As Variant: a(1, 1) = v: v = a
    End If
    ReadColumn = v
End Function

Private Function GetOrCreateTable(wsTgt As Worksheet, ids As Collection) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsTgt.ListObjects(TBL_NAME)
    On Error GoTo 0
    If Not lo Is Nothing Then Set GetOrCreateTable = lo: Exit Function

    Dim rng As Range
    Set rng = wsTgt.Range(TABLE_ANCHOR).Resize(2, 5)
    Set lo = wsTgt.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = TBL_NAME
    lo.HeaderRowRange.Value = Array(H_ID, H_USER, H_STAGE, H_SUB, H_ZONE)
    lo.ListRows(1).Range.Cells(1, 1).Value = ids(1)
    Dim k As Long
    For k = 2 To ids.Count
        lo.ListRows.Add.Range.Cells(1, 1).Value = ids(k)
    Next k
    Set GetOrCreateTable = lo
End Function

Private Sub ApplyLookupFormulas(lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    SetLookup lo, H_USER, COL_USER
    SetLookup lo, H_STAGE, COL_STAGE
    SetLookup lo, H_SUB, COL_SUB
    SetLookup lo, H_ZONE, COL_ZONE
End Sub

Private Sub SetLookup(lo As ListObject, kpiHeader As String, srcCol As Long)
    Dim lc As String, ic As String
    lc = ColLetter(srcCol): ic = ColLetter(COL_ID)
    lo.ListColumns(kpiHeader).DataBodyRange.Formula = _
        "=IFERROR(INDEX('" & SRC_SHEET & "'!" & lc & ":" & lc & _
        ",MATCH([@[" & H_ID & "]],'" & SRC_SHEET & "'!" & ic & ":" & ic & ",0)),"""")"
End Sub

Private Function ColLetter(ByVal n As Long) As String
    Dim s As String, r As Long
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function
