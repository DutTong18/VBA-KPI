Sub emailResults()
'function that copies the two worksheets to a temporary file and emails them as an attachment
    dim xOutlookObj as Object
    dim xEmailObj as Object
    dim tempFile as Workbook
    dim tempFileName as String
    dim body as String
    dim wb as Workbook
    dim ws1 as Worksheet
    dim ws2 as Worksheet

    Set wb = ActiveWorkbook
    set ws1 = wb.Worksheets("_StageStateCache") 'first worksheet to email
    set ws2 = wb.Worksheets("SchedulerData") 'second worksheet to email

    worksheets("_StageStateCache").visible = True   
    worksheets("SchedulerData").visible = True

    application.DisplayAlerts = False
    application.ScreenUpdating = False
    application.EnableEvents = False

    tempFile = Environ$("temp") & "\" & "KPI_Results_" & Format(Now, "dd-mm-yyyy_hh-mm-ss") & ".xlsx"
    'tempFileName = "KPI_Results_" & Format(Now, "dd-mm-yyyy_hh-mm-ss") & ".xlsx"
    'creates a temporary file to save the worksheets to be emailed
    wb.Worksheets(Array("_StageStateCache", "SchedulerData")).Copy
    ActiveWorkbook.SaveAs tempFile, FileFormat:=xlOpenXMLWorkbook
    wb.Worksheets(Array("_StageStateCache", "SchedulerData")).pasteSpecial Paste:=xlPasteall
    activeWorkbook.Close False
    
    Set xOutlookObj = CreateObject("Outlook.Application")
    Set xEmailObj = xOutlookObj.CreateItem(0)
    on error resume next
    With xEmailObj
        .To = "dut.tong@bhp.com"
        .Subject = "KPI Results"
        .Body = "Please find the attached KPI results."
        .Attachments.Add tempFile
        .Send
    End With
    tempfile.changeFileAccess Mode:=xlReadOnly
    kill tempFile
    tempfile.close savechanges:=False
    
    application.DisplayAlerts = True
    set xEmailObj = Nothing
    set xOutlookObj = Nothing
    application.ScreenUpdating = True
    application.EnableEvents = True
    
End Sub