Attribute VB_Name = "M10c_Copy_IN_to_DATA"

Sub copy_IN_to_DATA()
    
    Application.EnableEvents = False
'    Call Apply_UserNames
    Application.EnableEvents = True
''===============================||====================================
''Define Filter Values into ARRAY/RANGE defined in sheet SETTINGS
    Dim Databestand_naam As String
    Dim statusRange As Range
    Dim Column_Name_Werkbestand As String
    Dim Column_Name_Container As Range
    Dim rCell As Variant
    Dim I As Integer
    Dim j As Integer
    I = 0
    j = 0
    HeaderNameColumn = 0
    mylastRow_IN = 0
    mylastRow_Databestand = 0
''===============================||====================================
    Databestand_naam = ActiveWorkbook.Name
''===============================||====================================
   'Open Artikelbeheer.xlsm bestand om data te verwerken
   'Controle invoeren of bestand open of gesloten is  om risico's uitgechekt JA/Nee uit te sluiten
    Dim wbName(25) As String ''(20)
    Dim wbCheck As Boolean
    wbCheck = False
    wbCount = Workbooks.Count
    
    For X = 1 To wbCount          '.FullName 'incl.Path
           wbName(X) = Workbooks(X).Name
        If InStr(1, wbName(X), "Artikelbeheer.xlsm", vbTextCompare) <> 0 Then
            wbCheck = True
        End If
    Next
    
    If wbCheck = False Then
       MsgBox "Bestand   Artikelbeheer.xlsm   is gesloten" & vbNewLine & vbNewLine & _
              "Open het bestand en selecteer Aanvraag.ID dmv. Aanvraag.code", vbOKOnly + vbExclamation
       Exit Sub
    End If
wbCheck:
''===============================||====================================
    Workbooks("Artikelbeheer.xlsm").Worksheets("IN").Activate
    Call Generate_Ranges_ALL                           ''ERGENS ANDERS PLAATSEN
   'Define IN last row   '3 regels zijn leeg en worden niet opgeteld dus *** LET OP ".UsedRange.Rows.Count" telt alleen NonEempty regels
    HeaderName = 0
    Set WB = Workbooks("Artikelbeheer.xlsm")
    HeaderNameColumn = WB.Worksheets("IN").Cells(1, 1).SpecialCells(xlLastCell).Column
''===============================||====================================
   'Define Container last row   '3 regels zijn leeg en worden niet opgeteld dus *** LET OP ".UsedRange.Rows.Count" telt alleen NonEempty regels
    WB.Worksheets("IN").Activate
    Workbooks("Artikelbeheer").Worksheets("IN").Select
    Set statusRange = Range("IN_Aanvraag.code")
    mylastRow_IN = WB.Worksheets("IN").UsedRange.Rows.Count
''''''''''''''
   'Databestand rows aantal
    mylastRow_Databestand = Workbooks(Databestand_naam).Worksheets("Databestand").UsedRange.Rows.Count '' - HeadingRows
''===============================||====================================
''===============================||====================================
'rCell is positie in RangeName excl. HeadingRows
COPY_START:
SpeedOn
    Set Network = CreateObject("wscript.network")

    For Each rCell In statusRange.Cells
    Workbooks("Artikelbeheer").Worksheets("IN").Activate
    Workbooks("Artikelbeheer").Worksheets("IN").Select
'    statusRange(rCell.Row - HeadingRows, 1).Select
        If statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_34 Then                 ''"IN_inleveren"
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_40                      ''"DB_IN"
           Range("IN_Datum_OUT_AB").Cells(rCell.Row - HeadingRows, 1).Value = Now()               ''Date ''Datum_OUT
           
           Set Network = CreateObject("wscript.network")
           Range("IN_Databeheerder").Cells(rCell.Row - HeadingRows, 1).Value = Naam
           Range("IN_Datum_IN_DB").Cells(rCell.Row - HeadingRows, 1).Value = Now()                ''Date ''Datum_OUT
          'TESTEN om tekst in code te vermijden
          'dit moet zoals Range(Aanvraag_level_34).cells(rcell.row,1)=Now()
''========================================================================''
''========================================================================''
           
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).Select
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).COPY Destination:=Workbooks(Databestand_naam).Worksheets("Databestand").Cells(mylastRow_Databestand + 1, 1)

''========================================================================''
''========================================================================''
'CHECK_CHAR_START:   ''check aantal karakters als check aanwezig in kolom 3
'''''''''''''''
'    Workbooks("Artikelbeheer").Worksheets("Container").Activate
'    j = 0
'    Dim CHECK_CHAR_FLAG As Boolean
'    Dim CHECK_EMPTY_FLAG As Boolean
'    CHECK_CHAR_FLAG = False
'    CHECK_EMPTY_FLAG = False
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHECK_CHAR_GESLAAGD:
'          Controle van velden is al gedaan
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_39 ''"IN_OUT"
          'Define Container laast row   '3 regels zijn leeg dus let op ".UsedRange.Rows.Count" telt alleen NonEempty regels
           mylastRow_Databestand = mylastRow_Databestand + 1
COPY:
''''''''SpeedOff
''''''''SpeedOn
        Else
        End If
''========================================================================''
CHECK_OVER:
    rCell = (rCell.Row + 1)
''SpeedOff
Next rCell
''===============================||====================================
''===============================||====================================
''===============================||====================================
   Workbooks("Artikelbeheer").Worksheets("IN").Activate
   Workbooks("Artikelbeheer").Worksheets("IN").Select
   strFile = Path & "/" & ActiveWorkbook.Name
'''
   Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
   SpeedOff
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
COPY_END:
    Workbooks(Databestand_naam).Worksheets("Databestand").Activate
'''
   Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
   SpeedOff
'''
   If wbCheck = False Then
      Workbooks(Databestand_naam).Save
   Else
      MsgBox "Overzetten van aanvragen is mislukt. Probeer het opnieuw."
   End If
''===============================||====================================
COPY_END1:
End Sub
