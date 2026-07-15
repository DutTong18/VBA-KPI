Option Explicit
' Config constants and all helper functions live in KPI_Common.

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
    'Sheets.Add(After:=Sheets("Stope DCB Graph Data")).Name = TGT_SHEET 'Enable once in final sheet
    Sheets.Add.Name = TGT_SHEET 'remove once in live sheet
    Sheets(STATE_SHEET).Visible = True 'Not redundant needed to clear "Method 'Delete' of worksheet failing
    Sheets(STATE_SHEET).Delete 'Currently deleting the stage cache sheet, to allow the KPI to be re-run and the stage cache as hidden rather than very hidden.  This is a temporary fix until Im satisfied with testing and count placements
'start screen updating and set calculation back to automatic
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
'possibly add the creration of the hidden sheet to further simplify KPI_build.bas
End Sub
