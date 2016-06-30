Attribute VB_Name = "M10f_Copy_Accordering_to_OUT"

Sub copy_Accordering_to_OUT()
   Set Network = CreateObject("wscript.network")
''===============================||====================================
Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
''===============================||====================================
    Application.EnableEvents = False
'    Call Apply_UserNames
    Application.EnableEvents = True
''===============================||====================================
''Define Filter Values into ARRAY/RANGE defined in sheet SETTINGS
    Dim Werkbestand_naam As String
'    Dim Path As String
    Dim Databestand_naam As String
    Dim statusRange As Range
    Dim Column_Name_Werkbestand As String
    Dim Column_Name_OUT As Range
    Dim rCell As Variant
    Dim I As Integer
    Dim j As Integer
    I = 0
    j = 0
    HeaderNameColumn = 0    ''OUT
    mylastRow_OUT = 0
''===============================||====================================
    Werkbestand_naam = ActiveWorkbook.Name
''===============================||====================================
    Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
''''
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    Range("A:DZ").EntireColumn.Hidden = False
    Range("1:65000").EntireRow.Hidden = False
''''
   'Define OUT last row   '3 regels zijn leeg en worden niet opgeteld dus *** LET OP ".UsedRange.Rows.Count" telt alleen NonEempty regels
    HeaderName = 0
    Set WB = Workbooks("Artikelbeheer.xlsm")
    HeaderNameColumn = WB.Worksheets("Accordering").Cells(1, 1).SpecialCells(xlLastCell).Column
''===============================||====================================
    WB.Worksheets("OUT").Activate
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
''===============================||====================================
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    Range("A:DZ").EntireColumn.Hidden = False
    Range("1:65000").EntireRow.Hidden = False
''===============================||====================================
   'Define OUT last row  ".UsedRange.Rows.Count" telt alleen NonEempty regels
    mylastRow_OUT = WB.Worksheets("OUT").UsedRange.Rows.Count
'SETTINGS CALL
    WB.Worksheets("Accordering").Activate
    WB.Worksheets("Accordering").Select
    Set statusRange = Range("ACC_Aanvraag.code")
''===============================||====================================
'rCell is positie in RangeName excl. HeadingRows
COPY_START:
SpeedOn
    For Each rCell In statusRange.Cells
    WB.Worksheets("Accordering").Activate
    WB.Worksheets("Accordering").Select
    statusRange(rCell.Row - HeadingRows, 1).Select
        If Range("ACC_Gereed_voor_Upload.SAP").Cells(rCell.Row - HeadingRows, 1).Value <> "" And _
          (statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_64 Or _
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_67) Then                  ''"ACC_inleveren" of "ACC_afgewezen"
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_70                        ''"OUT_IN"
           Range("ACC_Datum_OUT_ACC").Cells(rCell.Row - HeadingRows, 1).Value = Now()               ''ACC_Datum_IN_ACC
           Range("ACC_Generator").Cells(rCell.Row - HeadingRows, 1).Value = Naam
           Range("ACC_Datum_IN_OUT").Cells(rCell.Row - HeadingRows, 1).Value = Now()                ''ACC_Datum_IN_ACC
''========================================================================''
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).Select
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).COPY Destination:=Workbooks("Artikelbeheer").Worksheets("OUT").Cells(mylastRow_OUT + 1, 1)
''========================================================================''
CHECK_CHAR_START:   ''check aantal karakters als check aanwezig in kolom 3
''''''''''''''
    WB.Worksheets("Accordering").Activate
    j = 0
    Dim CHECK_CHAR_FLAG As Boolean
    Dim CHECK_EMPTY_FLAG As Boolean
    CHECK_CHAR_FLAG = False
    CHECK_EMPTY_FLAG = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CHECK_CHAR_GESLAAGD:
'          Controle van velden is al gedaan
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_69        ''"ACC_OUT"
          'Define OUT laast row
           mylastRow_OUT = mylastRow_OUT + 1
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
    WB.Worksheets("OUT").Activate
    Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
    Application.Run ("'Lijsten_new.xlsm'!ProtectOn")
''===============================||====================================
    Workbooks("Artikelbeheer").Worksheets("Accordering").Activate
'    Workbooks("Artikelbeheer").Close savechanges:=True
SpeedOff
COPY_END:
    WB.Save
''===============================||====================================
COPY_END1:
'    Workbooks("Artikelbeheer").Worksheets("OUT").Activate
'    Workbooks("Artikelbeheer").Worksheets("OUT").Select
'    ActiveWorkbook.Close savechanges:=True ''False
'''''
'    Dim Path
'    Path = "http://dafshare-org.eu.paccar.com/organization/ops-mtc/Artikelaanvraag/"
'   Workbooks("Artikelbeheer").Select
'   check of die open is want dan kan je die sluiten
'   ActiveWorkbook.Close ("http://dafshare-org.eu.paccar.com/organization/ops-mtc/Artikelaanvraag/Artikelbeheer.xlsm") ''savechanges:=True   ''False
''''

End Sub

