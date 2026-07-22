Sub emailResults()
'function that copies the two worksheets to a temporary file and emails them as an attachment
    Dim xOutlookObj As Object
    Dim xEmailObj As Object
    Dim tempFileName As String
    Dim body As String
    Dim wb As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim tempFile As Workbook

    Set wb = ActiveWorkbook 'name of the workbook that contains the worksheets to be emailed
    Set ws1 = wb.Worksheets("_StageStateCache") 'first worksheet to email
    Set ws2 = wb.Worksheets("SchedulerData") 'second worksheet to email
    Worksheets("_StageStateCache").Visible = True
    Worksheets("SchedulerData").Visible = True

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    
   
    'tempFileName = "KPI_Results_" & Format(Now, "dd-mm-yyyy_hh-mm-ss") & ".xlsx"
    'creates a temporary file to save the worksheets to be emailed
    wb.Worksheets(Array("_StageStateCache", "SchedulerData")).Copy
    ActiveWorkbook.SaveAs tempFile, FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close False

    Set tempFile = Workbooks.Add 'create a new temporary workbook
    tempFile = Environ$("temp") & "\" & "KPI_Results_" & Format(Now, "dd-mm-yyyy_hh-mm-ss") & ".xlsx"
    tempFile.Worksheets(Array("_StageStateCache", "SchedulerData")).PasteSpecial Paste:=xlPasteAll

    
    Set xOutlookObj = CreateObject("Outlook.Application")
    Set xEmailObj = xOutlookObj.CreateItem(0)
    On Error Resume Next
    With xEmailObj
        .To = "dut.tong@bhp.com"
        .Subject = "KPI Results"
        .body = "Please find the attached KPI results."
        .Attachments.Add tempFile
        .Send
    End With
    tempFile.ChangeFileAccess Mode:=xlReadOnly
    Kill tempFile
    tempFile.Close savechanges:=False
    
    Application.DisplayAlerts = True
    Set xEmailObj = Nothing
    Set xOutlookObj = Nothing
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub
