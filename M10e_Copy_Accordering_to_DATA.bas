Attribute VB_Name = "M10e_Copy_Accordering_to_DATA"

Sub copy_Accordering_to_DATA()
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
    Dim Column_Name_OUT As Range
    Dim rCell As Variant
    Dim I As Integer
    Dim j As Integer
    I = 0
    j = 0
    HeaderNameColumn = 0
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
'    Application.Workbooks.Open ("http://dafshare-org.eu.paccar.com/organization/ops-mtc/Artikelaanvraag/Artikelbeheer.xlsm")
    Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
   'Define OUT last row   '3 regels zijn leeg en worden niet opgeteld dus *** LET OP ".UsedRange.Rows.Count" telt alleen NonEempty regels
    HeaderName = 0
    Set WB = Workbooks("Artikelbeheer.xlsm")
    HeaderNameColumn = WB.Worksheets("Accordering").Cells(1, 1).SpecialCells(xlLastCell).Column
''===============================||====================================
''''''''''''''
    Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
''controle invoeren als bestand al geopend
   'Define OUT last row   '3 regels zijn leeg en worden niet opgeteld dus *** LET OP ".UsedRange.Rows.Count" telt alleen NonEempty regels
    Workbooks(Databestand_naam).Worksheets("Databestand").Activate
    Workbooks(Databestand_naam).Worksheets("Databestand").Select
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
    mylastRow_Databestand = Workbooks(Databestand_naam).Worksheets("Databestand").UsedRange.Rows.Count
''''''''''''''
'SETTINGS CALL
    Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
    Workbooks("Artikelbeheer").Worksheets("Accordering").Select
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
    Set statusRange = Range("ACC_Aanvraag.code")
''===============================||====================================
''===============================||====================================
'rCell is positie in RangeName excl. HeadingRows
COPY_START:
SpeedOn
    For Each rCell In statusRange.Cells
    Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
    Workbooks("Artikelbeheer").Worksheets("Accordering").Select
    statusRange(rCell.Row - HeadingRows, 1).Select
        If statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_63 Then                    ''"ACC_retour_DB"
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_48                         ''DB_retour_zie_Opmerkingen
           Range("ACC_Databeheerder").Cells(rCell.Row - HeadingRows, 1).Value = Naam
           Range("ACC_Datum_IN_DB").Cells(rCell.Row - HeadingRows, 1).Value = Now()                  ''DB_Datum_OUT_DB
           Range("ACC_Accordeerder").Cells(rCell.Row - HeadingRows, 1).Value = Naam
           Range("ACC_Datum_OUT_ACC").Cells(rCell.Row - HeadingRows, 1).Value = Now()                ''DB_Datum_IN_ACC
''========================================================================''
''========================================================================''
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).Select
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).COPY Destination:=Workbooks(Databestand_naam).Worksheets("Databestand").Cells(mylastRow_Databestand + 1, 1)
''========================================================================''
''========================================================================''
CHECK_CHAR_START:   ''check aantal karakters als check aanwezig in kolom 3
''''''''''''''
    Workbooks(Databestand_naam).Worksheets("Databestand").Activate
    j = 0
    Dim CHECK_CHAR_FLAG As Boolean
    Dim CHECK_EMPTY_FLAG As Boolean
    CHECK_CHAR_FLAG = False
    CHECK_EMPTY_FLAG = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CHECK_CHAR_GESLAAGD:
'          Controle van velden is al gedaan
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_69                           ''"ACC_OUT
          'Define OUT laast row   '3 regels zijn leeg dus let op ".UsedRange.Rows.Count" telt alleen NonEempty regels
           mylastRow_Databestand = mylastRow_Databestand + 1
COPY:
''''''''
SpeedOff
SpeedOn
''''''''
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
   ActiveWorkbook.CheckIn strFile
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



