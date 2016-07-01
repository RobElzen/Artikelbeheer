Attribute VB_Name = "DAF_Case_MUTATIE"
Option Explicit

Sub Aanvraag_code_Case()

Call Affix_Case
Call ProtectOff
'Application.Run "'Lijsten_new.xlsm'!ProtectOnRows"
On Error Resume Next
If Not Application.Intersect(ActiveCell, Range(Affix & "Aanvraag.code")) Is Nothing Then
''========================================================
Dim tmpIndex_new As Variant
Dim tmpMatch_new As Variant
Dim tmpIndex_old As Variant
Dim tmpMatch_old As Variant
Dim Aanvraagcode_new As String
Dim Aanvraagcode_old As String

Set WB = Workbooks("Lijsten_new.xlsm")
    tmpIndex_new = 0
    tmpMatch_new = 0
    tmpIndex_old = 0
    tmpMatch_old = 0
''***********************
    Aanvraagcode_old = ActiveCell.Value
    Application.EnableEvents = False
    Aanvraagcode_new = ActiveCell.Value
    Application.Undo
    Aanvraagcode_old = ActiveCell.Value
    ActiveCell.Value = Aanvraagcode_new
    Application.EnableEvents = True
'    MsgBox "Oude Aanvraag.code " & old_Aanvraagcode & vbLf & vbLf & "Nieuwe Aanvraag.code " & new_Aanvraagcode
''***********************
tmpMatch_new = Application.Match(Aanvraagcode_new, WB.Worksheets("Aanvraag_code").Range("Lst_Aanvraag.code"), 0)
        If IsError(tmpMatch_new) Then
           MsgBox "Status niet in de Lijsten_new.xlsm gevonden!"
           Call ProtectOn
           Exit Sub
        Else
        End If
tmpIndex_new = Application.WorksheetFunction.Index(WB.Worksheets("Aanvraag_code").Range("Lst_Aanvraag.Level"), tmpMatch_new)
''***********************
tmpMatch_old = Application.Match(Aanvraagcode_old, WB.Worksheets("Aanvraag_code").Range("Lst_Aanvraag.code"), 0)
        If IsError(tmpMatch_old) Then
           MsgBox "Status niet in de Lijsten_new.xlsm gevonden!"
           Call ProtectOn
           Exit Sub
        Else
        End If
tmpIndex_old = Application.WorksheetFunction.Index(WB.Worksheets("Aanvraag_code").Range("Lst_Aanvraag.Level"), tmpMatch_old)
''***********************
Dim Lower_limit As Integer
Dim Upper_limit As Integer
Dim Warn_limit As Integer
Dim Worksheet_proces As String

Worksheet_proces = ActiveWorkbook.ActiveSheet.Name
Call SpeedOn
If ActiveWorkbook.Name <> "Artikelbeheer.xlsm" Then
    If ActiveSheet.Name = "Werkbestand" Then
       ActiveSheet.Range("B2") = "ME"
       Lower_limit = 10
       Warn_limit = 15
       Upper_limit = 19
ElseIf ActiveSheet.Name = "Container" Then
       ActiveSheet.Range("B2").Value = "CNT"
       Lower_limit = 20
       Warn_limit = 25
       Upper_limit = 29
ElseIf ActiveSheet.Name = "Databestand" Then
       ActiveSheet.Range("B2") = "DB"
       Lower_limit = 40
       Warn_limit = 45
       Upper_limit = 49
   End If
Else
    If ActiveSheet.Name = "IN" Then
       ActiveSheet.Range("B2") = "IN"
       Lower_limit = 30
       Warn_limit = 35
       Upper_limit = 39
ElseIf ActiveSheet.Name = "Accordering" Then
       ActiveSheet.Range("B2") = "ACC"
       Lower_limit = 60
       Warn_limit = 65
       Upper_limit = 69
ElseIf ActiveSheet.Name = "OUT" Then
       ActiveSheet.Range("B2") = "OUT"
       Lower_limit = 70
       Warn_limit = 75
       Upper_limit = 79
   End If
End If
Call SpeedOff
''''''''''''''''''''''''''''''''''''''''''''''''''''
'     MsgBox "Deze stap is gewardeerd op: " & "   | " & Lower_limit & " =<     " & tmpIndex_new & "     =< " & Upper_limit & " |", vbOKOnly + vbInformation
''''''''''''''''
'Afgehandelde status niet meer herstelen bijv. ingediend niet opnieuw kunnen indienen
    If (tmpIndex_new < Lower_limit Or tmpIndex_new > Upper_limit) Then
       MsgBox "Geen toegestane actie op dit werkblad: " & Worksheet_proces & vbNewLine & vbNewLine & _
              "Aanvraag.code:   " & Aanvraagcode_new, vbOKOnly + vbCritical
       Application.EnableEvents = False
       ActiveCell = Aanvraagcode_old
       GoTo ENDE

ElseIf Mid(tmpIndex_old, 2, 1) = 9 And _
       tmpIndex_old = Upper_limit Then
       MsgBox "Reeds afgehandeld!", vbOKOnly + vbExclamation, "Kies een andere Aaanvraag.code!"
       Application.EnableEvents = False
       ActiveCell = Aanvraagcode_old
       GoTo ENDE
        
ElseIf (Mid(tmpIndex_new, 2, 1) = 0 And _
       (tmpIndex_old > Lower_limit And _
        tmpIndex_old < Upper_limit)) Then
        MsgBox "Reeds in behandeling genomen!", vbOKOnly + vbExclamation, "Kies een andere Aaanvraag.code!"
        Application.EnableEvents = False
        ActiveCell = Aanvraagcode_old
        GoTo ENDE
        
'inbouwen controle dat handmatig op ingediend zetten kan voorkomen worden.
ElseIf (Mid(tmpIndex_new, 2, 1) >= 5 And _
       (tmpIndex_old >= Warn_limit And _
        tmpIndex_old < Upper_limit)) Then
        MsgBox "Melding STATUS !", vbOKOnly + vbExclamation, "Kies een geldige Aaanvraag.code!"
        Application.EnableEvents = False
        ActiveCell = Aanvraagcode_old
        GoTo ENDE

ElseIf tmpIndex_new > Warn_limit And _
        tmpIndex_new < Upper_limit Then
        MsgBox "Melding STATUS !", vbOKOnly + vbExclamation, "Kies een geldige Aaanvraag.code!"
        Application.EnableEvents = False
        ActiveCell = Aanvraagcode_old
       GoTo ENDE
       
ElseIf tmpIndex_new > Lower_limit And tmpIndex_new < Warn_limit Then
'       MsgBox "Approved: " & Worksheet_proces & "Aanvraag.code:   " & Aanvraagcode_new, vbOKOnly + vbExclamation, "Succes verder!"
'    If tmpIndex = Lower_limit Then Exit Sub      'Reeds ingedient naar volgende behandeling
'    If tmpIndex = Upper_limit - 1 Then Exit Sub  'Reeds aangenomen van vorige behandeling
       Application.EnableEvents = False
       ActiveCell = Aanvraagcode_new
       GoTo ENDE
''ElseIf tmpIndex_new >= Warn_limit And tmpIndex_new <= Upper_limit Then

ElseIf tmpIndex_new = Upper_limit Then
       MsgBox "Geen toegestane actie op dit werkblad: " & Worksheet_proces & vbNewLine & vbNewLine & _
              "Aanvraag.code:   " & Aanvraagcode_new, vbOKOnly + vbCritical
'    If tmpIndex = Lower_limit Then Exit Sub      'Reeds ingedient naar volgende behandeling
'    If tmpIndex = Upper_limit - 1 Then Exit Sub  'Reeds aangenomen van vorige behandeling
       Application.EnableEvents = False
       GoTo ENDE

ElseIf tmpIndex_new = Lower_limit And Aanvraagcode_old = "---NIEUW---" Then
'       MsgBox "Approved: " & Worksheet_proces & "Aanvraag.code:   " & Aanvraagcode_new, vbOKOnly + vbExclamation, "Succes verder!"
'    If tmpIndex = Lower_limit Then Exit Sub      'Reeds ingedient naar volgende behandeling
'    If tmpIndex = Upper_limit - 1 Then Exit Sub  'Reeds aangenomen van vorige behandeling
       Application.EnableEvents = False
       ActiveCell = Aanvraagcode_new
       GoTo ENDE
Else: MsgBox "Onbekende keuze!"
End If

ENDE:
''========================================================
End If
     Application.EnableEvents = True

If ActiveSheet.Name = "Databestand" Or ActiveSheet.Name = "Werkbestand" Then
   Call ProtectOnRows
Else
   Call ProtectOnALL
End If

End Sub
