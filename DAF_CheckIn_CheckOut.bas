Attribute VB_Name = "DAF_CheckIn_CheckOut"

Sub CheckIn(Cancel As Boolean)

Dim Msg As String
Dim Ans As Integer  'Answer

'    If Not Me.Saved = True Then
'        Msg = "Wilt u de wijzigingen in "
'        Msg = Msg & Me.Name & " opslaan?"
'        Ans = MsgBox(Msg, vbQuestion + vbYesNoCancel)
'        Select Case Ans
'            Case vbYes
'                Me.Save
'            Case vbNo
'                Me.Saved = True
'            Case vbCancel
'                Cancel = True
'                Exit Sub
'          End Select
'    End If

    On Error Resume Next
''===============================||====================================
Dim strFile As String
    strFile = ThisWorkbook.Path & "/" & ThisWorkbook.Name

'    ThisWorkbook.Sheets("Databestand").EnableCalculation = False
    ThisWorkbook.Sheets(ThisWorkbook.Name).EnableCalculation = False

If Workbooks(ThisWorkbook.Name).CanCheckIn = True Then
   ThisWorkbook.CheckIn strFile
   ThisWorkbook.Close (True)
ElseIf Workbooks(ThisWorkbook.Name).CanCheckIn = False Then
   ThisWorkbook.Close (False)
'Else
''    MsgBox "Inchecken van het bestand is niet mogelijk."
'    ThisWorkbook.Close (False)
End If

End Sub

Sub CheckOut()
 
   ThisWorkbook.ActiveSheet.EnableCalculation = True
  
   Application.Run "'Lijsten_new.xlsm'!ProtectOff"
   Application.Run "'Lijsten_new.xlsm'!Generate_Ranges_All"
   
   Application.Run "'Lijsten_new.xlsm'!Affix_Case"
   Application.Run "'Lijsten_new.xlsm'!ProtectOnRows"

End Sub

Sub CheckIn_CheckOut()
'CheckIn
'    If Right(DefPath, 1) <> "\" Then
'        DefPath = DefPath & "\"
'    End If
'CheckOut
End Sub

