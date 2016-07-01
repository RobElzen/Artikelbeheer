Attribute VB_Name = "DAF_Case_Intersect"

Sub Aanvraag_code_Intersect() ''(ByVal Target As Range)

Call Affix_Case

Dim strAddress As String
    strAddress = ActiveCell.Address
''========================================================
If Not Application.Intersect(ActiveCell, Range(Affix & "Leveranciersnummer")) Is Nothing Then

Dim Match As Variant
Dim tmpIndex As Variant
Dim tmpMatch As Variant

tmpIndex = 0
tmpMatch = 0
Match = ActiveCell

Set WB = Workbooks("Lijsten_new.xlsm")
Set ws = WB.Worksheets("Leverancier")
ws.Range("A:B").Sort Key1:=ws.Range("Lst_Leveranciersnummer"), Order1:=xlAscending, Header:=xlYes, Orientation:=xlSortColumns
ActiveSheet.Range(strAddress).Select    '1-5-2016 laatst toegevoegd
''''''''

tmpMatch = Application.Match(Match, ws.Range("Lst_Leveranciersnummer"), 0)
        If IsError(tmpMatch) Then
'           MsgBox "Status niet in de Lijsten_new.xlsm gevonden!"
           Exit Sub
        Else
        End If

tmpIndex = Application.WorksheetFunction.Index(ws.Range("Lst_Leveranciersnaam"), tmpMatch)
''''''''
    Call SpeedOn
    ActiveCell.Offset(0, 1) = tmpIndex
    Call SpeedOff
End If
''========================================================
If Not Application.Intersect(ActiveCell, Range(Affix & "Leveranciersnaam")) Is Nothing Then

tmpIndex = 0
tmpMatch = 0
Match = ActiveCell

Set WB = Workbooks("Lijsten_new.xlsm")
Set ws = WB.Worksheets("Leverancier")
ws.Range("A:B").Sort Key1:=ws.Range("Lst_Leveranciersnaam"), Order1:=xlAscending, Header:=xlYes, Orientation:=xlSortColumns
ActiveSheet.Range(strAddress).Select
''''''''
tmpMatch = Application.Match(Match, ws.Range("Lst_Leveranciersnaam"), 0)
        If IsError(tmpMatch) Then
'           MsgBox "Status niet in de Lijsten_new.xlsm gevonden!"
           Exit Sub
        Else
        End If

tmpIndex = Application.WorksheetFunction.Index(ws.Range("Lst_Leveranciersnummer"), tmpMatch)
''''''''
    Call SpeedOn
'     ActiveCells = Range(strAddress)      '1-5-2016 laatst gemarkeerd
    ActiveCell.Offset(0, -1) = tmpIndex
    Call SpeedOff
End If
''========================================================
If Not Application.Intersect(ActiveCell, Range(Affix & "Screening.DB")) Is Nothing Then
''''''''
Set Network = CreateObject("wscript.network")
''''''''
    Call SpeedOn
    If ActiveCell = "NEE" Then
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "NEE"
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbRed
       Range(Affix & "Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_44

       ActiveCell.Offset(0, 1) = Naam             ''Naam.DB
       ActiveCell.Offset(0, 2) = Now()            ''Datum.DB
    ElseIf ActiveCell = "JA" Then
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = ""
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone
       Range(Affix & "Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_44
       
       ActiveCell.Offset(0, 1) = Naam             ''Naam.DB
       ActiveCell.Offset(0, 2) = Now()            ''Datum.DB
    ElseIf ActiveCell = "" Then
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = ""
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone
       Range(Affix & "Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_41
       
       ActiveCell.Offset(0, 1) = ""
       ActiveCell.Offset(0, 2) = ""
    End If
    Call SpeedOff
''''''''
End If
''========================================================
End Sub
''========================================================
''========================================================
''========================================================
Sub Aanvraag_code_Intersect_Change() ''(ByVal Target As Range)

Call Affix_Case
     
If Not Application.Intersect(ActiveCell, Range(Affix & "Leveranciersnummer")) Is Nothing And _
       ActiveCell.Value <> 0 Then

Dim Match As Variant
Dim tmpIndex As Variant
Dim tmpMatch As Variant

tmpIndex = 0
tmpMatch = 0
Match = ActiveCell

Set WB = Workbooks("Lijsten_new.xlsm")
Set ws = WB.Worksheets("Leverancier")
''''''''
       Call SpeedOn
    If Match <> "" Then
    tmpMatch = Application.Match(Match, ws.Range("Lst_Leveranciersnummer"), 0)
            If IsError(tmpMatch) Then
'               MsgBox "Status niet in de Lijsten_new.xlsm gevonden!"
               Exit Sub
            Else
            End If

    tmpIndex = Application.WorksheetFunction.Index(ws.Range("Lst_Leveranciersnaam"), tmpMatch)
    ''''''''
        ActiveCell.Offset(0, 1) = tmpIndex
    End If
        Call SpeedOff
''''''''
End If
''========================================================
If Not Application.Intersect(ActiveCell, Range(Affix & "Leveranciersnaam")) Is Nothing And _
       ActiveCell.Value <> 0 Then

tmpIndex = 0
tmpMatch = 0
Match = ActiveCell

Set WB = Workbooks("Lijsten_new.xlsm")
Set ws = WB.Worksheets("Leverancier")
''''''''
        Call SpeedOn
    tmpMatch = Application.Match(Match, ws.Range("Lst_Leveranciersnaam"), 0)
            If IsError(tmpMatch) Then
'               MsgBox "Status niet in de Lijsten_new.xlsm gevonden!"
               Exit Sub
            Else
            End If

    tmpIndex = Application.WorksheetFunction.Index(ws.Range("Lst_Leveranciersnummer"), tmpMatch)
    ''''''''
        Call SpeedOn
        ActiveCell.Offset(0, -1) = tmpIndex
        Call SpeedOff
'''''''
End If
''========================================================
If Not Application.Intersect(ActiveCell, Range(Affix & "Screening.DB")) Is Nothing Then
''''''''
Set Network = CreateObject("wscript.network")
''''''''
    Call SpeedOn
    If ActiveCell = "NEE" Then
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "NEE"
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbRed    '1-5-2016 laatst toegevoegd
       ActiveCell.Offset(0, 1) = Naam             ''Naam.DB
       ActiveCell.Offset(0, 2) = Now()            ''Datum.DB
    ElseIf ActiveCell = "JA" Then
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "JA"
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbGreen  '1-5-2016 laatst toegevoegd
       ActiveCell.Offset(0, 1) = Naam             ''Naam.DB
       ActiveCell.Offset(0, 2) = Now()            ''Datum.DB
    ElseIf ActiveCell = "" Then
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = ""
       Range(Affix & "Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone   '1-5-2016 laatst toegevoegd
       ActiveCell.Offset(0, 1) = ""
       ActiveCell.Offset(0, 2) = ""
    End If
    Call SpeedOff
''''''''
End If
''========================================================
End Sub



