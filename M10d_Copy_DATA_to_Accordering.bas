Attribute VB_Name = "M10d_Copy_DATA_to_Accordering"

Sub copy_DATA_to_Accordering()
   Set Network = CreateObject("wscript.network")
''===============================||====================================
Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
''===============================||====================================
    Application.EnableEvents = False
'    Call Apply_UserNames
    Application.EnableEvents = True
''===============================||====================================
''Define Filter Values into ARRAY/RANGE defined in sheet SETTINGS
    Dim Databestand_naam As String
'    Dim Path As String
    Dim statusRange As Range
    Dim Column_Name_Werkbestand As String
    Dim Column_Name_OUT As Range                  ''Doe ik iets hiermee?
    Dim rCell As Variant
    Dim I As Integer
    Dim j As Integer
    I = 0
    j = 0
    HeaderNameColumn = 0
    mylastRow_Accordering = 0
''===============================||====================================
    Databestand_naam = ActiveWorkbook.Name
''===============================||====================================
'Open Artikelbeheer.xlsm bestand om data te verwerken
   'Controle invoeren of bestand open of gesloten is  om risico's uitgechekt JA/Nee uit te sluiten
    Dim wbName(25) As String ''(20)
    Dim wbCheck As Boolean
    wbCheck = False
    wbCount = Workbooks.Count
    
    For X = 1 To wbCount
           wbName(X) = Workbooks(X).Name
        If InStr(1, wbName(X), "Artikelbeheer.xlsm", vbTextCompare) <> 0 Then
            wbCheck = True
            GoTo wbCheck
        End If
    Next
    
    If wbCheck = False Then
       Dim xlApp As Excel.Application
       Dim strFile As String
           strFile = Path & "/" & "Artikelbeheer.xlsm"
        
       If Workbooks.CanCheckOut(strFile) = True Then
          Application.Workbooks.Open strFile
          Workbooks.CheckOut strFile
          wbCheck = True
       Else
           MsgBox "Uitchecken van het bestand is niet mogelijk." & vbNewLine & vbNewLine & _
                  "Probeer later nog een keer.", vbOKOnly + vbExclamation
           Exit Sub
       End If
    End If

wbCheck:
''===============================||====================================
    Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
   'Define OUT last row   '3 regels zijn leeg en worden niet opgeteld dus *** LET OP ".UsedRange.Rows.Count" telt alleen NonEempty regels
    HeaderName = 0
    Set WB = Workbooks("Artikelbeheer.xlsm")
    HeaderNameColumn = WB.Worksheets("Accordering").Cells(1, 1).SpecialCells(xlLastCell).Column
''''''''''''''
'   'zie init UserName
'    UserName = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
'    Werkbestand_IN_DB = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
''===============================||====================================
''controle invoeren als bestand al geopend  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
   'Define OUT last row   '3 regels zijn leeg en worden niet opgeteld dus *** LET OP ".UsedRange.Rows.Count" telt alleen NonEempty regels
    Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
    mylastRow_Accordering = Workbooks("Artikelbeheer").Worksheets("Accordering").UsedRange.Rows.Count
''===============================||====================================
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    Range("A:DZ").EntireColumn.Hidden = False
    Range("1:65000").EntireRow.Hidden = False
''===============================||====================================
'SETTINGS CALL
    Workbooks(Databestand_naam).Worksheets("Databestand").Activate
    Workbooks(Databestand_naam).Worksheets("Databestand").Select
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
    Set statusRange = Range("DB_Aanvraag.code")
'rCell is positie in RangeName excl. HeadingRows
CHECK_START:
SpeedOn
        Workbooks(Databestand_naam).Worksheets("Databestand").Rows(3).ClearContents
'        Selection.Cells.ClearContents
    

    For Each rCell In statusRange.Cells
    Workbooks(Databestand_naam).Worksheets("Databestand").Activate
    Workbooks(Databestand_naam).Worksheets("Databestand").Select
    statusRange(rCell.Row - HeadingRows, 1).Select
        If statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_44 Then                   ''"DB_inleveren"
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_60                        ''"AC_IN"
           Range("DB_Databeheerder").Cells(rCell.Row - HeadingRows, 1).Value = Naam
           Range("DB_Datum_OUT_DB").Cells(rCell.Row - HeadingRows, 1).Value = Now()                 ''DB_Datum_OUT_DB
           Range("DB_Accordeerder").Cells(rCell.Row - HeadingRows, 1).Value = Naam
           Range("DB_Datum_IN_ACC").Cells(rCell.Row - HeadingRows, 1).Value = Now()                 ''DB_Datum_IN_ACC
''========================================================================''
''========================================================================''
'Hier komt afhankelijkheid tussen de kolomen
'overnemen uit    M10a_Copy_WERK_to_Container

'WB_   en DB_   samenvoegen voorwaarden
If Cells(rCell.Row, Range("DB_Veiligheidsvoorraad.Bestelpunt").Column) = 0 And _
   Cells(rCell.Row, Range("DB_Voorraad.locatie").Column) = "" Then
   Cells(rCell.Row, Range("DB_Voorraad.locatie").Column) = "B55"
'  MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("DB_Voorraad.locatie").Column).Value & " : " & "  B55  "
ElseIf Cells(rCell.Row, Range("DB_Veiligheidsvoorraad.Bestelpunt").Column) > 0 And _
   Cells(rCell.Row, Range("DB_Voorraad.locatie").Column) = "" Then
   Cells(rCell.Row, Range("DB_Voorraad.locatie").Column) = "B00055"
'  MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("DB_Voorraad.locatie").Column).Value & " : " & "  B00055  "
End If

'Altijd laten berekenen?
If Cells(rCell.Row, Range("DB_Brutto_Inkoopprijs").Column) <> "" Then
   Cells(rCell.Row, Range("DB_Netto_Inkoopprijs").Column) = Cells(rCell.Row, Range("DB_Brutto_Inkoopprijs").Column) * _
                                                           (100 - Cells(rCell.Row, Range("DB_Korting").Column)) / 100
'  MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("DB_Netto_Inkoopprijs").Column).Value & " : " & Cells(rCell.Row, Range("DB_Netto_Inkoopprijs").Column).value
End If

'Altijd laten berekenen?
If Cells(rCell.Row, Range("DB_Netto_Inkoopprijs").Column) <> "" Then
   Cells(rCell.Row, Range("DB_Aanvraagbedrag").Column) = Cells(rCell.Row, Range("DB_Netto_Inkoopprijs").Column) * _
                                                         Cells(rCell.Row, Range("DB_Veiligheidsvoorraad.Bestelpunt").Column)
'  MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("DB_Aanvraagbedrag").Column).Value & " : " & Cells(rCell.Row, Range("DB_Aanvraagbedrag").Column).value
End If

''Aanvragernaam
 Cells(rCell.Row, Range("DB_Aanvrager.Vestiging").Column) = Vestiging
 Cells(rCell.Row, Range("DB_Aanvrager.Afdeling").Column) = Afdeling



''========================================================================''
CHECK_CHAR_START:   ''check aantal karakters als check aanwezig in kolom 3
'    Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Activate
    j = 0
'    Dim CHECK_CHAR_FLAG As Boolean
'    Dim CHECK_EMPTY_FLAG As Boolean
'    CHECK_CHAR_FLAG = False
'    CHECK_EMPTY_FLAG = False
    
    Dim CHECK_REQU_FLAG As Boolean
    Dim CHECK_FORM_FLAG As Boolean
    Dim CHECK_CHAR_FLAG As Boolean
    CHECK_REQU_FLAG = False
    CHECK_FORM_FLAG = False
    CHECK_CHAR_FLAG = False
''========================================================================''
''========================================================================''
''========================================================================''
''''''''''''''
''Check op lege cellen (+1   betekent incl. Heading) || Workbooks(Werkbestand_naam).Worksheets("Werkbestand").
    Dim requRange As Range       ''Column COL_REQUIRED
    Dim formRange As Range       ''Column COL_FORMAT
    Dim charRange As Range       ''Page COL_CHAR
    Dim ALL_Range As Range       ''Page RANGE_ALL
''===============================||====================================
'SETTINGS CALL
    Set requRange = Workbooks("Lijsten_new").Worksheets("SETTINGS").Range("SET.COL_REQUIRED_DB")
    Set formRange = Workbooks("Lijsten_new").Worksheets("SETTINGS").Range("SET.COL_FORMAT")
    Set charRange = Workbooks("Lijsten_new").Worksheets("SETTINGS").Range("SET.COL_CHAR")
    Set ALL_Range = Workbooks("Lijsten_new").Worksheets("SETTINGS").Range("SET.RANGE_ALL")
''===============================||====================================
    For j = 1 To Workbooks(Databestand_naam).Worksheets("Databestand").Cells(1, 1).SpecialCells(xlLastCell).Column

Dim HEAD_zoekwaarde As String
HEAD_zoekwaarde = Cells(1, j)
Dim hCell As Variant
Dim Match_Check As Boolean
Match_Check = False
''Workbooks(Werkbestand_naam).Worksheets ("Werkbestand")

''wbL.Activate
For Each hCell In ALL_Range ''.Cells
If hCell.Value <> "" Then
      If hCell.Value = HEAD_zoekwaarde Then
'         Workbooks(Werkbestand_naam).Activate
         Workbooks(Databestand_naam).Worksheets("Databestand").Cells(rCell.Row, j).Select
         Match_Check = True
''''''''''''
COL_REQUIRED:
         If requRange(hCell.Row - 1, 1).Value = "X" Then    ''Required field
         
'               Cells(rCell.Row, j).Select
            If Selection = "" Then
               'MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("WB_Type").Column).Value & " : " & "  Offerte door DataBeheerder laten aanvragen"
               Selection.Interior.Color = vbYellow
'              CHECK_EMPTY_FLAG = True
               CHECK_REQU_FLAG = True
               GoTo COL_CHAR
            Else
               Selection.Interior.Color = xlNone
            End If
         End If
''''''''''''
COL_FORMAT:
         If formRange(hCell.Row - 1, 1).Value = "T" Then    ''Text
            Selection.NumberFormat = "General"
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Cells(rCell.Row, j).NumberFormat = "General"
         GoTo COL_CHAR
         End If
         
         If formRange(hCell.Row - 1, 1).Value = "N" Then    ''Numeric
         
'               Cells(rCell.Row, j).Select
            If Not IsNumeric(Selection) Then
               'MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("WB_Type").Column).Value & " : " & "  Offerte door DataBeheerder laten aanvragen"
               Selection.Interior.Color = vbRed
               CHECK_FORM_FLAG = True
            Else
               If Selection.Interior.Color = vbYellow Then
               CHECK_FORM_FLAG = True
               Else
                  Selection.Interior.Color = xlNone
               End If
            End If
            
            Selection.NumberFormat = "#0_ ;-#0 "
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Selection.NumberFormat = "#,#0.0_ ;-#,#0.0 "
         
         End If
         
         If formRange(hCell.Row - 1, 1).Value = "N1" Then    ''Numeric
         
'               Cells(rCell.Row, j).Select
            If Not IsNumeric(Selection) Then
               'MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("WB_Type").Column).Value & " : " & "  Offerte door DataBeheerder laten aanvragen"
               Selection.Interior.Color = vbRed
               CHECK_FORM_FLAG = True
            Else
               If Selection.Interior.Color = vbYellow Then
               CHECK_FORM_FLAG = True
               Else
                  Selection.Interior.Color = xlNone
               End If
            End If
            
            Selection.NumberFormat = "#,#0.0_ ;-#,#0.0 "
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Selection.NumberFormat = "#,#0.0_ ;-#,#0.0 "
         
         End If
         
         If formRange(hCell.Row - 1, 1).Value = "N2" Then    ''Numeric
         
'               Cells(rCell.Row, j).Select
            If Not IsNumeric(Selection) Then
               'MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("WB_Type").Column).Value & " : " & "  Offerte door DataBeheerder laten aanvragen"
               Selection.Interior.Color = vbRed
               CHECK_FORM_FLAG = True
            Else
               If Selection.Interior.Color = vbYellow Then
               CHECK_FORM_FLAG = True
               Else
                  Selection.Interior.Color = xlNone
               End If
            End If
            
            Selection.NumberFormat = "#,##0.00_ ;-#,##0.00 "
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Selection.NumberFormat = "#,#0.0_ ;-#,#0.0 "
         
         End If
         
         If formRange(hCell.Row - 1, 1).Value = "N3" Then    ''Numeric
         
'               Cells(rCell.Row, j).Select
            If Not IsNumeric(Selection) Then
               'MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("WB_Type").Column).Value & " : " & "  Offerte door DataBeheerder laten aanvragen"
               Selection.Interior.Color = vbRed
               CHECK_FORM_FLAG = True
            Else
               If Selection.Interior.Color = vbYellow Then
               CHECK_FORM_FLAG = True
               Else
                  Selection.Interior.Color = xlNone
               End If
            End If
            
            Selection.NumberFormat = "#,###0.000_ ;-#,###0.000 "
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Selection.NumberFormat = "#,#0.0_ ;-#,#0.0 "
         
         End If
         
         If formRange(hCell.Row - 1, 1).Value = "NE" Then    ''Numeric Exact
         End If
         
         If formRange(hCell.Row - 1, 1).Value = "D" Then    ''Date time
            Selection.NumberFormat = "[$-13]dd-mm-yyyy h:mm;@" ''Format(Date, "dd-mm-yyyy h:mm")
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Selection.NumberFormat = "[$-13]dd-mm-yyyy;@" ''Format(Date, "dd-mm-yyyy")
         GoTo COL_CHAR
         End If
        
         If formRange(hCell.Row - 1, 1).Value = "V" Then    ''Valuta
            Selection.NumberFormat = "€* #,##0.00"
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Selection.NumberFormat = "#,##0.00"
         GoTo COL_CHAR
         End If
''''''''''''
COL_CHAR:
         If charRange(hCell.Row - 1, 1).Value > 0 Then
         
            If Len(Cells(rCell.Row, j)) > charRange(hCell.Row - 1, 1).Value Then
               Cells(rCell.Row, j).Select
               Selection.Interior.Color = vbRed
               CHECK_CHAR_FLAG = True
'               MsgBox "Er zijn te lange teksten aanwezig. Svp. aanpassen."
            Else
'               Cells(rCell.Row, j).Select
               If Selection.Interior.Color = vbYellow Then
                  CHECK_CHAR_FLAG = True
               Cells(3, j) = Len(Cells(rCell.Row, j))
               Else
                  Selection.Interior.Color = xlNone
               End If
            End If
            
'            Cells(3, j) = Cells(3, j) & vbNewLine & "{RIJ " & rCell.Row & " } " & "Geteld: " & Len(Cells(rCell.Row, j)) & " (Toegestaan: " & charRange(hCell.Row - 1, 1).Value & " ) "
'            Cells(3, j) = Cells(3, j) & vbNewLine & "{RIJ " & rCell.Row & " } " & "Char." & Len(Cells(rCell.Row, j)) & " van " & charRange(hCell.Row - 1, 1).Value & ") "
            Test_Col_Char = Cells(3, j) & "{RIJ " & rCell.Row & " } " & "Char." & Len(Cells(rCell.Row, j)) & " van " & charRange(hCell.Row - 1, 1).Value & ") "
            Cells(3, j) = Test_Col_Char & vbNewLine ''& "{RIJ " & rCell.Row & " } " & "Char." & Len(Cells(rCell.Row, j)) & " van " & charRange(hCell.Row - 1, 1).Value & ") "
         End If
''''''''''''
'''''''''
      Else
      End If
End If
     If Match_Check = True Then GoTo Match_Check_OK
Next hCell
''===============================||====================================
'               Cells(rCell.Row, j).Select
'            If Cells(rCell.Row, j) = "" Then
'               Cells(rCell.Row, j).Select
'               Selection.Interior.Color = vbYellow
'               CHECK_EMPTY_FLAG = True
''               MsgBox "Er zijn lege velden aanwezig. Svp. invullen."
'            Else
'               Cells(rCell.Row, j).Select
'               Selection.Interior.Color = xlNone
'            End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Het controlegetal is te vinden in de 4de regel : LOOP door elke cell in 4de regel
        'Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Activate
'         Column_Name_Werkbestand_CHECK = Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Cells(4, j)
''''''''''''''
''''''''''''''

Match_Check_OK:
Match_Check = False
''''''''''''''
    Next j
''''''''''''''
If CHECK_REQU_FLAG = True Or _
   CHECK_FORM_FLAG = True Or _
   CHECK_CHAR_FLAG = True Then
   
   CHECK_REQU_FLAG = False
   CHECK_FORM_FLAG = False
   CHECK_CHAR_FLAG = False
   
   statusRange(rCell.Row - HeadingRows, 1).Select
   statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_47
   GoTo CHECK_OVER
Else
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHECK_CHAR_GESLAAGD:
'CHECK_OVER:
'End If
'     rCell = (rCell.Row + 1)
'Next rCell



''Controle hoeveel regels zijn er om over te zetten naar Accordering
        If Application.WorksheetFunction.CountIf(statusRange, Aanvraag_level_60) > 0 Then
           GoTo COPY
        Else
           MsgBox "Er zijn geen regels om over te zetten."
           Exit Sub
        End If
COPY:
'''''''
SpeedOff
SpeedOn
'''''''
''========================================================================''
''========================================================================''
''========================================================================''
''========================================================================''
''========================================================================''
''========================================================================''
''========================================================================''
''========================================================================''
COPY_START:
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).Select
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).COPY Destination:=Workbooks("Artikelbeheer").Worksheets("Accordering").Cells(mylastRow_Accordering + 1, 1)
''========================================================================''
''========================================================================''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CHECK_CHAR_GESLAAGD:
'          Controle van velden is al gedaan
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_49 ''"DB_OUT"    overschrijven waarde  Aanvraag_level_60
          'Define OUT laast row   '3 regels zijn leeg dus let op ".UsedRange.Rows.Count" telt alleen NonEempty regels
           mylastRow_Accordering = mylastRow_Accordering + 1
'COPY:
'''''''''
'SpeedOff
'SpeedOn
'''''''''
        Else
        End If
''========================================================================''
CHECK_OVER:
    rCell = (rCell.Row + 1)
'SpeedOff
Next rCell
''===============================||====================================
''===============================||====================================
''===============================||====================================
   Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
   Workbooks("Artikelbeheer").Worksheets("Accordering").Select
   strFile = Path & "/" & ActiveWorkbook.Name
'''
   Application.Run ("'Lijsten_new.xlsm'!ProtectOnALL")
   Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
   SpeedOff
COPY_END:
'''
If ActiveWorkbook.CanCheckIn = True Then
   ActiveWorkbook.CheckIn (strFile)
   wbCheck = False
'  MsgBox "The file has been checked in."
End If
''===============================||====================================
''===============================||====================================
''===============================||====================================
SpeedOff
COPY_END1:
    Workbooks(Databestand_naam).Worksheets("Databestand").Activate
'''
   Application.Run ("'Lijsten_new.xlsm'!ProtectOnRows")
   Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
   SpeedOff
'''
   If wbCheck = False Then
      Workbooks(Databestand_naam).Save
   Else
      MsgBox "Overzetten van aanvragen is mislukt. Probeer het opnieuw."
   End If
''===============================||====================================
End Sub



