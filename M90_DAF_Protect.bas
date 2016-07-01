Attribute VB_Name = "M90_DAF_Protect"
Option Explicit

'Private Sub ProtectOn()
Public Sub ProtectOn()
                '       Application.Run "RetrieveUser"

       If Niveau = 1 And ActiveSheet.Name <> "" Then
          Call ProtectOff
       Else
          Call ProtectOnALL
       End If
End Sub

Public Sub ProtectOnALL()
       Call Protection                                   ''AllowDeletingRows:=True          RISICO
       ActiveSheet.Protect Password:="aaaaaa", DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                           AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                           AllowInsertingHyperlinks:=True, AllowDeletingRows:=True, _
                           AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
                                                                    'Userinterfaceonly:=False    'POPUP PASSWORD INVOEREN
End Sub

Public Sub ProtectOnRows()
       Call ProtectOff
       ActiveSheet.Rows.Locked = False
'       ActiveSheet.Rows("1:5").Locked = True             ''AllowDeletingRows:=True          RISICO
       ActiveSheet.Rows("1:2").Locked = True             ''AllowDeletingRows:=True          RISICO
       ActiveSheet.Rows("4:5").Locked = True             ''AllowDeletingRows:=True          RISICO
       ActiveSheet.Protect Password:="aaaaaa", DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                           AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                           AllowInsertingHyperlinks:=True, AllowDeletingRows:=True, _
                           AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
                                                                    'Userinterfaceonly:=False    'POPUP PASSWORD INVOEREN
End Sub

'Private Sub ProtectOff()
Public Sub ProtectOff()
'       ActiveSheet.Rows.Locked = False
'       ActiveSheet.Columns.Locked = False
       ActiveSheet.Unprotect Password:="aaaaaa"
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Protection()

If ActiveWorkbook.Name <> "Artikelbeheer.xlsm" Then
           If ActiveSheet.Name = "Werkbestand" Then
              Call ProtectOnRows
       ElseIf ActiveSheet.Name = "Container" Then
              Call ProtectOnRows
       ElseIf ActiveSheet.Name = "Databestand" Then
              Call ProtectOnRows
       End If
Else
           If ActiveSheet.Name = "IN" Then
              Range("IN_Aanvraag.code").Locked = False
       ElseIf ActiveSheet.Name = "Accordering" Then
'              ActiveSheet.Rows("1:5").Locked = True
              ActiveSheet.Rows("1:2").Locked = True
              ActiveSheet.Rows("4:5").Locked = True
              Range("ACC_Aanvraag.code").Locked = False
       ElseIf ActiveSheet.Name = "OUT" Then
              Range("OUT_Aanvraag.code").Locked = False
'              ActiveSheet.Rows("1:5").Locked = True
              ActiveSheet.Rows("1:2").Locked = True
              ActiveSheet.Rows("4:5").Locked = True
       End If
End If

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

