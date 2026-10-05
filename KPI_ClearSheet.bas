Option Explicit
' Config constants and all helper functions live in KPI_Common.
' Config constants and all helper functions live in KPI_Common.

' ================== MACRO: First Send KPI Results via Email then clear sheets ==================
' Recipient for the results email.
Public Const RESULTS_TO As String = "xyz@company.com
Public Const RESULTS_CC As String = "xyz@company.com
Public Const RESULTS_BCC As String = "xyz@company.com

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
    If wsState Is Nothing Then MsgBox "Sheet '" & STATE_SHEET & "' not found. Run RunStatusCheck first.", vbExclamation: GoTo CleanStateFail
    If wsTgt Is Nothing Then MsgBox "Sheet '" & TGT_SHEET & "' not found. Run BuildKPITable first.", vbExclamation: GoTo CleanFail

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
        .cc = RESULTS_CC
        .bcc = RESULTS_BCC
        .Subject = "KPI Results - " & Format(Date, "dd/mm/yyyy")
        .body = "Please find the attached KPI results for the prior forecast"
        .Attachments.Add tempPath
        .Send
    End With

    'delete the temp file after sending the email.
    Kill tempPath

    MsgBox "KPI results sent to " & RESULTS_TO & ".", vbInformation
    Call ClearSheets

CleanExit:
    Application.ScreenUpdating = prevSU: Application.EnableEvents = prevEv
    Application.DisplayAlerts = prevDA
    Exit Sub
CleanFail:
    MsgBox "emailResults error: " & Err.Description, vbCritical
    'send error email to user
    Dim errorMail As Object, errorOutlookApp As Object
    Set errorOutlookApp = CreateObject("Outlook.Application")
    Set errorMail = errorOutlookApp.CreateItem(0)   ' olMailItem
    With errorMail
        .To = RESULTS_TO
        .cc = RESULTS_CC
        .bcc = RESULTS_BCC
        .Subject = "KPI Results - Error Occurred"
        .body = "An error occurred while sending KPI results: " & Err.Description
        .Send
    End With
    Resume CleanExit
    
'============================================================================================================ Statesheet not found
CleanStateFail:
    MsgBox "emailResults error: " & Err.Description, vbCritical
    'send error email to user
    Dim error1Mail As Object, error1OutlookApp As Object
    Set error1OutlookApp = CreateObject("Outlook.Application")
    Set error1Mail = error1OutlookApp.CreateItem(0)   ' olMailItem
    With error1Mail
        .To = RESULTS_TO
        .cc = RESULTS_CC
        .bcc = RESULTS_BCC
        .Subject = "KPI Results - Error Occurred"
        .body = "An error occurred while sending KPI results: " & Err.Description & "State Sheet Could Not Be Found"
        .Send
    End With
    Resume CleanExit

'============================================================================================================ Target Sheet Not Found
CleanTGTFail:
    MsgBox "emailResults error: " & Err.Description, vbCritical
    'send error email to user
    Dim error2Mail As Object, error2OutlookApp As Object
    Set error2OutlookApp = CreateObject("Outlook.Application")
    Set error2Mail = error2OutlookApp.CreateItem(0)   ' olMailItem
    With error2Mail
        .To = RESULTS_TO
        .cc = RESULTS_CC
        .bcc = RESULTS_BCC
        .Subject = "KPI Results - Error Occurred"
        .body = "An error occurred while sending KPI results: " & Err.Description & "Target Sheet Could Not Be Found"
        .Send
    End With
    Resume CleanExit
    
End Sub
' ================== MACRO: ClearSheets / Clearing sheets at forecast change ==================
Public Sub ClearSheets()
'Stop screen updating and set calculation to manual to speed up the process
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
'clear the contents of the source and target sheets
    'Sheets(SRC_SHEET).UsedRange.ClearContents
    Sheets(TGT_SHEET).UsedRange.ClearContents
    Sheets(TGT_SHEET).Delete
    Sheets.Add(After:=Sheets("Stope DCB Graph Data")).Name = TGT_SHEET 'Enable once in final sheet
    'Sheets.Add.Name = TGT_SHEET 'remove once in live sheet
    Sheets(STATE_SHEET).Visible = True 'Not redundant needed to clear Method 'Delete' of worksheet failing
    Sheets(STATE_SHEET).Delete 'Currently deleting the stage cache sheet, to allow the KPI to be re-run and the stage cache as hidden rather than very hidden.  This is a temporary fix until I'm satisfied with testing and count placements
'start screen updating and set calculation back to automatic
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
'possibly add the creration of the hidden sheet to further simplify KPI_build.bas
End Sub



