Attribute VB_Name = "Protect_test"


'Sub AddUserEditRange()
'    Dim aer As AllowEditRange
'    Set ws = ThisWorkbook.Sheets("Protection")
'    ws.Unprotect "Excel2003"
'    Set aer = ws.Protection.AllowEditRanges.Add("User Range", ws.Range("A1:D4"))
'    aer.Users.Add "Power Users", True
'    ws.Protect "Excel2003"
'End Sub

'Sub RemoveUserEditRange()
'    Dim rng As Range, aer As AllowEditRange
'    Set ws = ThisWorkbook.Sheets("Protection")
'    ws.Unprotect
'    For Each aer In ws.Protection.AllowEditRanges
'        aer.Delete
'    Next
'End Sub



'Sub UCase()
''Upadateby20140701
'Dim rng As Range
'Dim WorkRng As Range
'On Error Resume Next
'xTitleId = "KutoolsforExcel"
'Set WorkRng = Application.Selection
'Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
'For Each rng In WorkRng
'    rng.Value = VBA.UCase(rng.Value)
'Next
'End Sub


'Sub ProperCase()
''Updateby20131127
'Dim rng As Range
'Dim WorkRng As Range
'On Error Resume Next
'xTitleId = "KutoolsforExcel"
'Set WorkRng = Application.Selection
'Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
'For Each rng In WorkRng
'    rng.Value = Application.WorksheetFunction.Proper(rng.Value)
'Next
'End Sub

