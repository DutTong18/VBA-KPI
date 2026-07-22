Option Explicit
Sub openKpiQrg()
    
    Dim webURL As String
    webURL = "https://spo.bhpbilliton.com/:w:/r/sites/ResEngMDOD/_layouts/15/Doc.aspx?action=edit&sourcedoc=%7B5aeee561-e367-4c15-bdf4-4dc156760f2d%7D&wdExp=TEAMS-TREATMENT&web=1&TeamsCID=e2634c22-99d3-4b1d-a56b-9e4050d95697&wdLOR=cB6EFC9DC-5F3D-4E42-B5FA-593D227C67CD"

    If URLcheck(webURL) = True Then
        ThisWorkbook.FollowHyperlink Address:=webURL, NewWindow:=True
    Else
        MsgBox "QRG not found, link broken :(", vbExclamation
    End If

End Sub

'errHandler:
    'MsgBox "QRG not found, link broken :(", vbExclamation
'End Sub


