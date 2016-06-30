Attribute VB_Name = "DAF_File_Open_Close"
Option Explicit

Sub Workbook_Open_File()
        Dim WB As Workbook
        Set WB = Workbooks(ThisWorkbook.Name)
        WB.Activate
        WB.Worksheets("Databestand").Visible = xlSheetVisible
''===============================||====================================''
        Dim Lijsten_status As Boolean
        Lijsten_status = False

On Error Resume Next

Dim filespec As String
    filespec = ThisWorkbook.Path & "/" & "Lijsten_new.xlsm"

For Each WB In Workbooks
    If WB.FullName = filespec Then
       Lijsten_status = True
    End If
Next
''===============================||====================================''
''===============================||====================================''
   Application.ScreenUpdating = False
If Lijsten_status = False Then
   Application.Workbooks.Open (filespec), ReadOnly:=True
   ActiveWindow.Visible = False
   ThisWorkbook.Application.Visible = False
End If
   Application.ScreenUpdating = True
''===============================||====================================''
   
   ThisWorkbook.Application.Visible = True
   ThisWorkbook.Activate
   ThisWorkbook.Application.Visible = True
   Sheets("Databestand").EnableCalculation = True
   
   Application.Run "'Lijsten_new.xlsm'!ProtectOff"
   Application.Run "'Lijsten_new.xlsm'!Generate_Ranges_All"
   
   Application.Run "'Lijsten_new.xlsm'!Prefix_Case"
   Application.Run "'Lijsten_new.xlsm'!ProtectOnRows"

End Sub

''===============================||====================================''
Sub Workbook_BeforeClose_File(Cancel As Boolean)
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
End Sub


''===============================||====================================
Sub Workbook_Close_File()
Dim strFile As String
    strFile = ThisWorkbook.Path & "/" & ThisWorkbook.Name

    ThisWorkbook.Sheets("Databestand").EnableCalculation = False

If Workbooks(ThisWorkbook.Name).CanCheckIn = True Then
   ThisWorkbook.CheckIn strFile
   ThisWorkbook.Close (True)
Else
    MsgBox "Inchecken van het bestand is niet mogelijk."
    ThisWorkbook.Close (False)
End If

End Sub

''===============================||====================================
'Sub Workbook_Activate_File()
'
'            If ActiveWorkbook.ReadOnly Then
'                ''Call SheetsVeryHidden
'                MsgBox "File was opened as read-only"
'            Else
'              ''MsgBox "File was not opened as read-only"
'            End If
'
''        Run "DoubleClickDisable"
'End Sub

'Private Sub Workbook_Deactivate()
'
'End Sub





