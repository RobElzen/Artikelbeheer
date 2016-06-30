Attribute VB_Name = "M82_DAF_Lijsten_Interface"
Option Explicit
Public Selected_Cell As Variant
Public ROWnr As Variant
Public WSI As Worksheet
Public WSS As Worksheet
'Public Counter As Integer

Sub RECORD_Nieuw()
    Set WSI = Worksheets("Interface")
''===============================||====================================
''You can select cells on a hidden sheet... if you activate the sheet first... (leave it hidden)
    WSI.Activate
'Make all Collumns / Rows Visible
    Dim Sheet_Cells_All
    
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    Range("A:EZ").EntireColumn.Hidden = False
    Range("1:65000").EntireRow.Hidden = False
''===============================||====================================
    'CONTROLE OP DUBBELING: Afbreken indien dubbel anders Opslaan als nieuwe record
If Not Intersect(ActiveCell, Range("B2:B65000")) Is Nothing And _
       ActiveCell.Value <> "" Then

    Dim Target_new As Variant, Target_Range_new As Variant
        Target_new = ActiveCell.Value
        Target_Range_new = ActiveCell.Offset(0, -1).Value
                
                If Application.CountIf(Range("Lst_" & Target_Range_new), Target_new) >= 1 Then
                       MsgBox "Deze waarde bestaat al:" & vbTab & "'' " & Target_new & " ''" & vbLf & vbLf & _
                              "Aub. een unieke waarde toevoegen!"
                       WSI.Select
                       Exit Sub

                ElseIf Application.CountIf(Range("Lst_" & Target_Range_new), Target_new) = 0 Then
                       MsgBox "Deze waarde wordt bij een uniek gegevensreeks toegevoegd!" & vbLf & vbLf & vbTab & _
                              "'' " & Target_new & " ''"
                
                
                       Dim newEntry As Variant, rws As Long, adres As String
                       newEntry = ActiveCell.Value
                       rws = Range("Lst_" & Target_Range_new).Rows.Count
                       
                       If Target_Range_new = "Statistieknummer" Or _
                          Target_Range_new = "Leveranciersnummer" Then
                          
                                If Not IsNumeric(ActiveCell) Or _
                                   Len(ActiveCell) > ActiveCell.Offset(0, 2) Then
                                       ActiveCell.NumberFormat = "###"
                                   ActiveCell.Interior.Color = vbRed
                                   MsgBox "error"
                                   Exit Sub
                                Else
                                   ActiveCell.Interior.Pattern = xlNone
                                   ActiveCell.Interior.TintAndShade = 0
                                   ActiveCell.Interior.PatternTintAndShade = 0
                                End If
                           
                       End If
                       ''''''''''''''''''''''''''''''''
                       
                       Range("Lst_" & Target_Range_new).Cells(1, 1).Offset(rws, 0).Value = newEntry
                       With Range("Lst_" & Target_Range_new).Name
                            adres = .RefersTo
                            .Delete
                       End With
                       Range(adres).Resize(rws + 1, 1).Name = "Lst_" & Target_Range_new
                End If

'If Range("Lst_" & Target_Range_new) = "Leveranciersnummer" Then
'
'End If
'
End If

If Target_Range_new = "Statistieknummer" Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  Range("Lst_" & Target_Range_new).NumberFormat = "###"
'SORT
    Range(adres).Sort Key1:=Range(adres), Order1:=xlAscending, Header:=xlNo, OrderCustom:=1, _
    MatchCase:=True, Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers


'If Target_Range_new = "Producent" Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                  Range("Lst_" & Target_Range_new).NumberFormat = "###"
''SORT
'    Range(adres).Sort Key1:=Range(adres), Order1:=xlAscending, Header:=xlNo, OrderCustom:=1, _
'    MatchCase:=True, Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers

''===========================COLUMNWIDTH===============================
'ElseIf Target_Range_new = "Leveranciersnummer" Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                  Range("Lst_" & Target_Range_new).NumberFormat = "###"
''SORT
'    Range(adres).Sort Key1:=Range(adres), Order1:=xlAscending, Key2:=Range(adres).Offset(0, 1), Header:=xlNo, OrderCustom:=1, _
'    MatchCase:=True, Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers '', DataOption1:=xlSortNormaal
''===========================COLUMNWIDTH===============================
''Sorteren Leveranciersnummer icm. Leveranciersnaam moet uitgewerkt worden
''capital letters Leveranciersnaam toepassen
End If

End Sub


Sub RECORD_Wijzigen()
MsgBox "In progress!"
Exit Sub
End Sub
'
'    Set WSI = Worksheets("MASTER")
'    Set WSS = Worksheets("START")
'    Application.ScreenUpdating = False
'''===============================||====================================
'''You can select cells on a hidden sheet... if you activate the sheet first... (leave it hidden)
'    WSI.Activate
''Make all Collumns / Rows Visible
'    Dim Sheet_Cells_All
'
'    If ActiveSheet.AutoFilterMode = True Then
'       ActiveSheet.AutoFilterMode = False
'    End If
'
'    Range("A:EZ").EntireColumn.Hidden = False
'    Range("1:65000").EntireRow.Hidden = False
''===============================||====================================
''Copy L Column [Range Names] to Worksheet "MASTER"
'    WSS.Select
'    WSS.Range("B2:B112").COPY
''===============================||====================================
'''Define Filter Values into ARRAY/RANGE defined in sheet SETTINGS
'    Dim RowsCounter11 As Variant, RowsCounter21 As Variant, RowsCounter30 As Variant
'    Dim Filter_crit11 As Variant, Filter_crit21 As Variant, Filter_crit30 As Variant
'
'    Filter_crit11 = WSS.Cells(11, "B").Value    ''11, "C"
'    Filter_crit21 = WSS.Cells(21, "B").Value    ''21, "C"
'    Filter_crit30 = WSS.Cells(30, "B").Value    ''30, "C"
'    RowsCounter11 = Application.Match(Filter_crit11, Range("MASTER.TRUCK.Nr"), 0)
'    RowsCounter21 = Application.Match(Filter_crit21, Range("MASTER.TRUCK.Serienr"), 0)
'    RowsCounter30 = Application.Match(Filter_crit30, Range("MASTER.SAP.ID"), 0)
'
'    Dim Target As Variant, Target_new As Variant
'    Dim Serienr As Variant, Serienr_new As Variant
'    Dim SAP_ID As Variant, SAP_ID_new As Variant
'    Target = WSS.Cells(11, "B").Value
'    Target_new = WSS.Cells(11, "C").Value
'    Serienr = WSS.Cells(21, "B").Value
'    Serienr_new = WSS.Cells(21, "C").Value
'    SAP_ID = WSS.Cells(30, "B").Value
'    SAP_ID_new = WSS.Cells(30, "C").Value
'''////////////////////////////////////////////////////////////
'    'CONTROLE OP DUBBELINGEN | WIJZIGEN VAN UNIEKE WAARDEN | TRUCK.Nr & Serienr |
'    If Target_new <> "" And Target <> Target_new Then
'          ''WIJZIGEN VAN TRUCK.Nr IS MOGELIJK DOOR BEIDE WAARDEN TE WIJZIGEN "Target = Target_new"
'          MsgBox "Unique TRUCK.Nr van een BESTAAND Object wijzigen is niet mogelijk! " & vbLf & vbLf & _
'                 "" & Target_new & vbTab & vbTab & "anders dan" & vbTab & Target & vbLf & vbLf & vbLf & _
'                 "Wilt u dit object als een nieuw object met een uniek TRUCK.Nr & een uniek Serienr opvoeren?" & vbLf & vbLf & _
'                 "JA: Gebruik knop NIEUW" & vbTab & vbTab & "NEE: Pas " & Target_new & vbTab & vbTab & Serienr_new & vbTab & " aan!"
'          Exit Sub
'    ElseIf Serienr_new <> "" And Serienr <> Serienr_new Then
'           ''WIJZIGEN VAN Serienr IS MOGELIJK DOOR BEIDE WAARDEN TE WIJZIGEN "Serienr = Serienr_new"
'           MsgBox "Unique Serienr van een BESTAAND Object wijzigen is niet mogelijk! " & vbLf & vbLf & _
'                  "" & Serienr_new & vbTab & vbTab & "anders dan" & vbTab & Serienr & vbLf & vbLf & vbLf & _
'                  "Wilt u dit object als een nieuw object met een uniek Serienr & een uniek Serienr opvoeren?" & vbLf & vbLf & _
'                  "JA: Gebruik knop NIEUW" & vbTab & vbTab & "NEE: Pas " & Target_new & vbTab & vbTab & Serienr_new & vbTab & " aan!"
'          Exit Sub
'    ElseIf SAP_ID_new <> "" And SAP_ID <> SAP_ID_new Then
'           ''WIJZIGEN VAN Serienr IS MOGELIJK DOOR BEIDE WAARDEN TE WIJZIGEN "Serienr = Serienr_new"
'           MsgBox "Unique SAP_ID van een BESTAAND Object wijzigen is niet mogelijk! " & vbLf & vbLf & _
'                  "" & SAP_ID_new & vbTab & vbTab & "anders dan" & vbTab & SAP_ID & vbLf & vbLf & vbLf & _
'                  "Wilt u dit object als een nieuw object met een uniek SAP_ID & een uniek Serienr opvoeren?" & vbLf & vbLf & _
'                  "JA: Gebruik knop NIEUW" & vbTab & vbTab & "NEE: Pas " & Target_new & vbTab & vbTab & Serienr_new & vbTab & SAP_ID_new & vbTab & " aan!"
'          Exit Sub
'    ElseIf Target <> "" And Application.CountIf(Range("MASTER.TRUCK.Nr"), Target) > 1 Then
'          MsgBox "Er bestaat al een object met TRUCK.Nr: " & vbTab & Target_new & vbLf & vbLf & _
'                 "Aub. een uniek TRUCK.Nr toekennen!"
'          Exit Sub
'    ElseIf Serienr <> "" And Application.CountIf(Range("MASTER.TRUCK.Serienr"), Serienr) > 1 Then
'          MsgBox "Er bestaat al een object met TRUCK.Serienr: " & vbTab & Serienr_new & vbLf & vbLf & _
'                 "Aub. een uniek TRUCK.Serienr toekennen!"
'          Exit Sub
'    ElseIf SAP_ID_new <> "" And Application.CountIf(Range("MASTER.SAP.ID"), SAP_ID) > 5 Then
'          MsgBox "Er bestaat al een object met SAP.ID: " & vbTab & SAP_ID_new & vbLf & vbLf & _
'                 "Aub. een uniek SAP.ID toekennen!"
'          Exit Sub
'    Else
'    End If
''===============================||====================================
'       WSI.Activate
'       WSI.Range("A" & RowsCounter11 + 1).PasteSpecial Paste:=xlPasteValues, _
'                            Operation:=xlNone, SkipBlanks:=False, Transpose:=True
'
'      WSI.Range("J:J").NumberFormat = "####0000"
'      WSI.Range("T:T").NumberFormat = "###"
'      WSI.Range("AC:AC").NumberFormat = "###"
'      'WSI.Range("T:T").NumberFormat = "@"       ''Als Tekst
'      'WSI.Range("AC:AC").NumberFormat = "#####00000"
'
'       Application.CutCopyMode = False
''===============================||====================================
'    WSI.Visible = xlVeryHidden
'    WSS.Select
'End Sub

