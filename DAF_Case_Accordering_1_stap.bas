Attribute VB_Name = "DAF_Case_Accordering_1_stap"

'Çode wort gebruikt alleen bij werkblad "Accordering"
'Alleen door Accorderdeers te gebruiken
'Rechten toekenning van toepassing

Sub Accordering_stap() ''(ByVal Target As Range)
'Date & Time Stamp voor de kolommen "Screening" t/m "Contract"
    On Error Resume Next

    Set Network = CreateObject("wscript.network")
'----- 1. Declaratie Variabelen -----

Set wbA = Workbooks("Artikelbeheer.xlsm")
Set ws = Worksheets("Accordering")

Dim Accoordcode_new As String
Dim Acc_Range As Range
''***********************
    Accoordcode_new = ActiveCell.Value
''***********************
On Error Resume Next
Set Acc_Range = Range("ACC_" & Cells(1, ActiveCell.Column))
Set Network = CreateObject("wscript.network")

'----- 3. Code Activiteiten -----
''Application.ScreenUpdating = False
ActiveCell.Select
If Not Intersect(ActiveCell, Acc_Range) Is Nothing And _
       Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) <> Aanvraag_level_69 Then
'      Range("ACC_Aanvraag.code").Cells(ActiveCell.Row, 1) <> Aanvraag_level_69 Then
         If Accoordcode_new = "NEE" Then
            ActiveCell.Offset(0, 1).Value = Naam
            ActiveCell.Offset(0, 2).Value = Format(Now, "dd-mm-yyyy h:mm")
        ElseIf Accoordcode_new = "JA" Then
            ActiveCell.Offset(0, 1).Value = Naam
            ActiveCell.Offset(0, 2).Value = Format(Now, "dd-mm-yyyy h:mm")
        Else
        End If
Else
      Dim Aanvraagcode_new As String
     Dim Aanvraagcode_old As String
    
     Application.EnableEvents = False
     Aanvraagcode_new = ActiveCell.Value
     Application.Undo
     Aanvraagcode_old = ActiveCell.Value
     Application.EnableEvents = True
End If

'Door Inkoop contract gekoppeld
'If Not Intersect(Target, Range(KolTO_Cntrct & ":" & KolTO_Cntrct)) Is Nothing Then
'    For Each C In Target
'        If UCase(C.Value) = "NEE" Then
'            C.Offset(0, 14).Value = Now
'        ElseIf UCase(C.Value) = "JA" Then
'            C.Offset(0, 15).Value = Now
'        Else
'        C.Offset(0, 14).ClearContents
'        C.Offset(0, 15).ClearContents
'        End If
'    Next C
'End If

''Application.ScreenUpdating = True
'----- ------------------- -----
'Call Generate_Ranges_ALL
End Sub

