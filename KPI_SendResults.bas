Option Explicit
' Config constants and all helper functions live in KPI_Common.

' Recipient for the results email - edit as needed (semicolon-separate multiples).
Public Const RESULTS_TO As String = "xyz@company.com"

' MACRO: Email KPI results
' Copies the cache sheet (cumulative engineer breakdown + baseline) and the
' target sheet (KPI table + summary) into a temporary .xlsx, emails it via
' Outlook as an attachment, then deletes the temp file.
Public Sub emailResults()
    Dim prevSU As Boolean, prevEv As Boolean, prevDA As Boolean
    prevSU = Application.ScreenUpdating: prevEv = Application.EnableEvents
    prevDA = Application.DisplayAlerts
    On Error GoTo CleanFail
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim wsState As Worksheet, wsTgt As Worksheet
    Set wsState = SheetOrNothing(STATE_SHEET)
    Set wsTgt = SheetOrNothing(TGT_SHEET)
    If wsState Is Nothing Then MsgBox "Sheet '" & STATE_SHEET & "' not found. Run RunStatusCheck first.", vbExclamation: GoTo CleanExit
    If wsTgt Is Nothing Then MsgBox "Sheet '" & TGT_SHEET & "' not found. Run BuildKPITable first.", vbExclamation: GoTo CleanExit

    ' A sheet must be visible to be copied to a new workbook; remember the
    ' cache sheet's state and restore it afterwards.
    Dim prevStateVis As Long: prevStateVis = wsState.Visible
    wsState.Visible = xlSheetVisible

    ' Copy both sheets into a fresh workbook and save it in the temp folder.
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\KPI_Results_" & Format(Now, "dd-mm-yyyy_hh-mm-ss") & ".xlsx"
    ThisWorkbook.Worksheets(Array(STATE_SHEET, TGT_SHEET)).Copy
    Dim tempWb As Workbook: Set tempWb = ActiveWorkbook
    tempWb.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook   ' .xlsx, no macros
    tempWb.Close SaveChanges:=False

    wsState.Visible = prevStateVis

    ' Build and send the email.
    Dim outlookApp As Object, mail As Object
    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)   ' olMailItem
    With mail
        .To = RESULTS_TO
        .Subject = "KPI Results - " & Format(Date, "dd-mm-yyyy")
        .body = "Please find the attached KPI results."
        .Attachments.Add tempPath
        .Send
    End With

    ' Attachment is embedded in the mail item, safe to delete the temp file.
    Kill tempPath

    MsgBox "KPI results sent to " & RESULTS_TO & ".", vbInformation

CleanExit:
    Application.ScreenUpdating = prevSU: Application.EnableEvents = prevEv
    Application.DisplayAlerts = prevDA
    Exit Sub
CleanFail:
    MsgBox "emailResults error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
