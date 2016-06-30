Attribute VB_Name = "DAF_Case_Rows_Delete"
Sub RetrieveStatus()
Dim CellStatusReversed As String
Dim CellStatus As String
Dim answer As Integer

If ActiveWorkbook.Name = "Lijsten_new.xlsm" Then Exit Sub
answer = MsgBox("Are you sure you want to delete completed rows?", vbYesNo + vbQuestion, "Empty Sheet")
If answer = vbYes Then

    For Each Worksheet In ActiveWorkbook.Worksheets
       Application.Run ("'Lijsten_new.xlsm'!Affix_Case")
       Dim Pcell As Variant
       For Each Pcell In Range(Affix & "Aanvraag.code")
       
           CellStatusReversed = Left(StrReverse(Pcell), 3)
           CellStatus = Left(StrReverse(CellStatusReversed), 3)
              
           If CellStatus = "OUT" Then
              Pcell.Select
              Pcell.EntireRow.Select
              Pcell.EntireRow.Delete
           End If
       Next Pcell
       
    Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL") ''TESTEN
    Next Worksheet

ElseIf Antwoord = vbNo Then Exit Sub
End If

End Sub






''===============================||====================================
''===============================||====================================
''===============================||====================================
''Controle invoeren of bestand open of gesloten is  om risico's uitgechekt JA/Nee uit te sluiten
' Dim wbName(25) As String ''(20)
' Dim wbCheck As Boolean
' wbCheck = False
' wbCount = Workbooks.Count
' For x = 1 To wbCount
'        wbName(x) = Workbooks(x).FullName
'     If InStr(1, wbName(x), "Container", vbTextCompare) = 0 Then ''> 0 Then
'  Application.Workbooks.Open ("http://dafshare-org.eu.paccar.com/organization/ops-mtc/Artikelaanvraag/Container.xlsm")
'        wbCheck = True
'     Else: GoTo wbCheck
'     End If
' Next x
''                       'Check of het bestand als read only is geopend <<<<<<<<<<<<<<<<<<<<<<<<<<<<
''                    '    If wbCheck = True Then
''                    '       MsgBox "Bestand   Container.xlsm   is gesloten of geopend als READ only" & vbNewLine & vbNewLine & _
''                    '              "Open het bestand", vbOKOnly + vbExclamation
''                    '      ''Exit Sub
''                    '    End If
'wbCheck:




''========================================================



'Sub Delete_Worksheets()
''Delete Specific Worksheet
'        Dim wkb As Worksheets, wks As Worksheets
'        Dim Current As Worksheet
'
'        Application.DisplayAlerts = False
'For Each Current In Worksheets
'        If Left(Current.Name, 8) = "RAPPORT_" Then
'        Application.DisplayAlerts = False
'        Current.Delete
'        Application.DisplayAlerts = True
'        End If
'Next Current
'End Sub
''========================================================




'Sub DeleteUnused()
'
'Dim myLastRow As Double ''Long
'Dim myLastCol As Double ''Long
'Dim wks As Worksheet
'Dim dummyRng As Range
'
'
'For Each wks In ActiveWorkbook.Worksheets
'  With wks
'If wks.Name = "Formulier" Then GoTo Eind
'If wks.Name = "Print" Then GoTo Eind
'
'    myLastRow = 0
'    myLastCol = 0
'    Set dummyRng = .UsedRange
'    On Error Resume Next
'    myLastRow = _
'      .Cells.Find("*", after:=.Cells(1), _
'        LookIn:=xlFormulas, lookat:=xlWhole, _
'        searchdirection:=xlPrevious, _
'        searchorder:=xlByRows).Row
'    myLastCol = _
'      .Cells.Find("*", after:=.Cells(1), _
'        LookIn:=xlFormulas, lookat:=xlWhole, _
'        searchdirection:=xlPrevious, _
'        searchorder:=xlByColumns).Column
'    On Error GoTo 0
'
'    If myLastRow * myLastCol = 0 Then
'        .Columns.Delete
'    Else
'        .Range(.Cells(myLastRow + 1, 1), _
'          .Cells(.Rows.Count, 1)).EntireRow.Delete
'        .Range(.Cells(1, myLastCol + 1), _
'          .Cells(1, .Columns.Count)).EntireColumn.Delete
'    End If
'  End With
'Eind:
'Next wks
'
'End Sub

