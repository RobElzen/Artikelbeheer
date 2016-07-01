Attribute VB_Name = "M55_Init_Generator_CSV"
Option Explicit

Public Sub CSV_RAPPORT()
    On Error Resume Next
    
    Application.StatusBar = ""
    Set Network = CreateObject("wscript.network")
'Sheet SETTINGS
    Dim clmInvisible As Range   ''C_olumn_zichtbaar
    Dim clmRange As Range       ''C_olumn_User definition RAPPORT_OUT "UserGroup"
    Dim clhRange As Range       ''C_olumn LAY_out (Hide)
    Dim clwRange As Range       ''C_oLumn W_idth
    Dim clmRangeMail As Range   ''C_olumn_User_Mail
    
    Dim pbrRange As Range       ''Column Page_Break
    Dim fltRange As Range       ''Column Filtering
    Dim floRange As Range       ''Column Filtering_OUT
    Dim srtRange As Range       ''Column Sorting
    Dim pstRange As Range       ''Page Setup
    Dim atrRange As Range       ''Attributes
    Dim molRange As Range       ''Column MailOnly
    Dim mailRange As Range      ''Column Mailing
    Dim abgRange As Range       ''Column Accoordbedrag
    Dim uplRange As Range       ''Column Upload
    Dim upwRange As Range       ''Column Upload_waarde
    
    Dim r           As Long
    Dim c           As Byte
    Dim Antwoord As Integer
    
'Initialise Invisible Data Collection
    Dim rRange As Range
    Dim fCell As Variant
    Dim rCell As Variant
    Dim rCelly As Variant
    Dim mCell As Variant
    Dim rapCell As Variant
    Dim clmCell As Range
    Dim Invisible As Range
    Dim ColumnCounter As Byte
    Dim fssFILT As Range
    Dim EnkelRapport As Boolean
    Dim rngUserGroup As Range

'Define PageBreak hor-vert
    Dim rngMyRange_data As Range
    Dim rngMyRange_data_rapport As Range
    Dim rngCell_data As Variant
    Dim Sheet_Cells_All
    
'Define Layout Rapport Page
    Dim vOrientation As String
    Dim vPaperSize As String
''===============================||====================================
    Set wbL = Workbooks("Lijsten_new.xlsm")
    Set wbS = wbL.Worksheets("SETTINGS")
    Set wbA = Workbooks("Artikelbeheer.xlsm")
    Set wbAS = wbA.Worksheets(ActiveSheet.Name)

    Set clhRange = wbS.Range("SET.ColumnHide")
    Set pbrRange = wbS.Range("SET.PageBreak")
    Set clwRange = wbS.Range("SET.ColumnWidth")
   'Set fltRange = wbS.Range("SET.Filtering")
    Set floRange = wbS.Range("SET.Filtering_OUT")
    Set srtRange = wbS.Range("SET.Sorting")
    Set pstRange = wbS.Range("SET.PageSetup")
   'Set atrRange = wbS.Range("SET.Attributes")
    Set abgRange = wbS.Range("SET.Accoordbedrag")
    Set uplRange = wbS.Range("SET.Upload")
    Set upwRange = wbS.Range("SET.Upload_waarde")
''===============================||====================================
    Call SpeedOn
    Call Affix_Case
    Call Apply_UserNames
    Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''Define File Name to save a sheets accoording to Values in RANGE("Bestandsnaam") in sheet SETTINGS
''============================CREATE FILE================================
    'AS: ArtikelBeheer
     wbAS.Activate
     wbAS.Select
Call SpeedOn    ''new test
For Each fCell In wbS.Range("SET.Bestandsnaam").Cells
      If fCell.Value <> "" Then
      
            'Makes a copy of the active sheet and save it to a temporary file
             Dim wbUpload As Worksheet ''Workbook
             Dim wsUpload As Worksheet

             filename = fCell.Value & ".xlsx"  ''" " & Format(Now, "d mmm yyyy hhnn").Value & ".xlsx"                ''= fCell.Value & ".xlsx"
             '==============================================================
             'Replace existing files
             Application.SheetsInNewWorkbook = 1
             Workbooks.Add          '(After:=Worksheets(Worksheets.Count)).Name = "Master"
             ActiveSheet.Name = "Master"
             ActiveWorkbook.SaveAs "C:\Temp\" & filename, FileFormat:=51

             Dim WinShuttle_Name As Integer
             WinShuttle_Name = fCell.Row

     Else: GoTo Next_filename
     End If
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
     wbAS.Activate
     wbAS.Select

    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

Call SpeedOn    ''new test
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
'Define which Rapport Set is active (via rapCell)
UserGroup = ""
Application.StatusBar = ""
For Each rapCell In wbS.Range("SET.RANGE_ALL").Cells
      If rapCell.Value = ActiveSheet.Name Then
         GoTo Rapport_sheet
      End If
     rapCell = (rapCell.Row + 1)
Next rapCell
Rapport_sheet:
''===============================||====================================
      'Define First User Column
       For UserColumn = 1 To wbS.Cells(1, 1).SpecialCells(xlLastCell).Column
           If wbS.Cells(1, UserColumn).Value = "Master" Then
              FirstUserColumn = UserColumn + 1
           ElseIf wbS.Cells(1, UserColumn).Value = "Bestandsnaam" Then
              LastUserColumn = UserColumn - 1
           End If
       Next UserColumn
''FIRST RUN TO CREATE NEW RANGE, TO FILTER A NEW RANGES A COPY TO NEW SHEET CSV_x
''======================== CREATE RANGES VOOR CSV_x ============================
    Dim I
    Dim j
    Dim RapportSelection As String
    I = 0
    j = 0
     
     For I = FirstUserColumn To LastUserColumn
           RapportSelection = wbS.Cells(1, I).Value
           UserGroup = RapportSelection     'CSV_group
           Set clmRange = wbS.Range("SET." & RapportSelection) ''UserGroup)
''===============================||====================================
Application.StatusBar = ". . . . . . . User Rapport: " & " : " & UserGroup
''===============================||====================================
'Fill new created range with values specifiek for CSV_"x" (only columns to be added)

Call SpeedOff

'    If ActiveSheet.AutoFilterMode = True Then
'       ActiveSheet.AutoFilterMode = False
'    End If
     wbAS.Activate
     wbAS.Select
Call ProtectOff
For Each rCell In uplRange.Cells            ''SET.Upload
      If rCell.Value = "Y" Then ''And
        'wbS.Range("SET." & UserGroup).Cells(rCell.Row - 1, 1).Value <> "" Then  ''CSV_x <> ""
        'Define Column Heading (Name)
         Cells(1, rCell.Row - 1).Select
         Cells(1, rCell.Row - 1).Value = wbS.Range("SET.VariableName").Cells(rCell.Row - 1, 1).Value
'STAP 1 schrijf alle waarde van kolom SET.Upload
'STAP 2 "COPY"  waarden van genoemde kolom in SET.Upload
'STAP 3 "NL-BE" waarde zetten afhankelijk van kolom in SET.Upload
'STAP 4 "NL-BE" waarde zetten afhankelijk van kolom in SET.Upload

Call SpeedOn
          Dim COPY_Range As String
              COPY_Range = wbS.Range("SET.VariableName").Cells(rCell.Row - 1, 1).Value
              
             'Range creeren
              r = 0
              c = 0
              r = Cells(Rows.Count, 1).End(xlUp).Row
              c = Cells(1, rCell.Row - 1).End(xlUp).Column
              Range(Cells(6, c), Cells(r, c)).Name = Affix & Cells(1, c).Value
              Range(Affix & COPY_Range).EntireColumn.NumberFormat = "@"
          
          Dim IdemDito As String
              IdemDito = wbS.Range("SET.Range_ALL").Cells(rCell.Row - 1, 1).Value
''''''''''''''''''''''''''''''''
          If wbS.Range("SET.Upload_waarde").Cells(rCell.Row - 1, 1).Value <> "" Then

             For Each rCelly In Range(Affix & COPY_Range).Cells
                      rCelly.Select
                      rCelly.Value = ""
                      rCelly.Value = wbS.Range("SET.Upload_waarde").Cells(rCell.Row - 1, 1).Value
                 Next rCelly
             
           End If
''''''''''''''''''''''''''''''''
'COPY'COPY'COPY'COPY'COPY'COPY
''''''''''''''''''''''''''''''''
        If wbS.Range("SET.Upload_waarde").Cells(rCell.Row - 1, 1).Value = "COPY" Then

             For Each rCelly In Range(Affix & COPY_Range).Cells

                      rCelly.Select
                      rCelly.Value = ""
''''''''''''''''
                 If COPY_Range = "VERPR" And _
                    clmRange.Cells(fCell.Row - 1, 1).Value = "Repdelen" Then
                    
                              If Range("OUT_Reparatiedeel").Cells(rCelly.Row - HeadingRows, 1).Value = "Ja" And _
                              Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "NL" Then
                                 rCelly.Value = "0,01"
                              Else
                                 rCelly.Value = ""
                              End If
                 Else
                              rCelly.Value = Range(Affix & IdemDito).Cells(rCelly.Row - HeadingRows, 1).Value
                             'rCelly.NumberFormat = "General"
                              rCelly.NumberFormat = "#,##0.00"
                 End If
''''''''''''''''
                 If COPY_Range = "MINBE" Then
                    
                              If Range("OUT_Reparatiedeel").Cells(rCelly.Row - HeadingRows, 1).Value = "Ja" And _
                              Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "NL" Then
                              rCelly.Value = Range(Affix & IdemDito).Cells(rCelly.Row - HeadingRows, 1).Value
                              Else
                                 rCelly.Value = ""
                              End If
                 Else
                              rCelly.Value = Range(Affix & IdemDito).Cells(rCelly.Row - HeadingRows, 1).Value
                             'rCelly.NumberFormat = "General"
                              rCelly.NumberFormat = "#,##0.00"
                 End If
''''''''''''''''
                 If COPY_Range = "EISBE" Then
                    
                              If Range("OUT_Reparatiedeel").Cells(rCelly.Row - HeadingRows, 1).Value = "Ja" And _
                                 Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "NL" Then
                                 rCelly.Value = ""
                              Else
                                 rCelly.Value = Range(Affix & IdemDito).Cells(rCelly.Row - HeadingRows, 1).Value
                              End If
                 Else
                              rCelly.Value = Range(Affix & IdemDito).Cells(rCelly.Row - HeadingRows, 1).Value
                             'rCelly.NumberFormat = "General"
                              rCelly.NumberFormat = "#,##0.00"
                 End If
''''''''''''''''
             Next rCelly
        
        End If
''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''
'NL-BE
''''''''''''''''''''''''''''''''
    If wbS.Range("SET.Upload_waarde").Cells(rCell.Row - 1, 1).Value = "NL-BE" Then
        
        If COPY_Range = "WERKS" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      
                      If Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "NL" Then
                         rCelly.Value = "NL01"
                      ElseIf Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "BE" Then
                         rCelly.Value = "BE01"
                      End If
                 Next rCelly
         End If

        If COPY_Range = "EKGRP" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      If Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "NL" Then
                         rCelly.Value = "E01"
                      ElseIf Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "BE" Then
                         rCelly.Value = "W01"
                      End If
                 Next rCelly
         End If

        If COPY_Range = "BKLAS" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      If Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "NL" Then
                         rCelly.Value = "3040"
                      ElseIf Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "BE" Then
                         rCelly.Value = "2855"
                      End If
                 Next rCelly
         End If

        If COPY_Range = "BUKRS" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      If Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "NL" Then
                         rCelly.Value = "7002"
                      ElseIf Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "BE" Then
                         rCelly.Value = "7019"
                      End If
                 Next rCelly
         End If

        If COPY_Range = "WBKLA" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      If Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "NL" Then
                         rCelly.Value = "3040"
                      ElseIf Range("OUT_Vestiging").Cells(rCelly.Row - HeadingRows, 1).Value = "BE" Then
                         rCelly.Value = "2855"
                      End If
                 Next rCelly
         End If
    
'Reparatiedelen
        If COPY_Range = "BWTTY" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      If Range("OUT_Reparatiedeel").Cells(rCelly.Row - HeadingRows, 1).Value = "Ja" Then
                         rCelly.Value = "C"
                      Else
                      End If
                 Next rCelly
         End If

        If COPY_Range = "SPART" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      If Range("OUT_Reparatiedeel").Cells(rCelly.Row - HeadingRows, 1).Value = "Ja" Then
                         rCelly.Value = "RD"
                      Else
                      End If
                 Next rCelly
         End If
                
        If COPY_Range = "BWTAR_NEW" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      If Range("OUT_Reparatiedeel").Cells(rCelly.Row - HeadingRows, 1).Value = "Ja" Then
                         rCelly.Value = "NIEUW"
                      Else
                         rCelly.Value = ""
                     End If
                 Next rCelly
         End If
                      
    Range(Affix & COPY_Range).EntireColumn.NumberFormat = "@"
    
    End If
''''''''''''''''''''''''''''''''
      End If
     rCell = (rCell.Row + 1)
Next rCell
'''''''''''''''''''''''''''''''''
'6de rij leegmaken van alle waardes
Worksheets("OUT").Rows(6).Select
Worksheets("OUT").Rows(6).ClearContents
'totdat ik iets beters bedenk / even geen tijd
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
'Is het "OUT" werkblad en staat er "x" bij  CSV_"x"
'Is het CSV_name geldig (Naamgeving bijv. "Stam")
'Is de destination file geldig (Bestandsnaam gelijk aan aangemaakt Bestand *.xlsx)
           If clmRange.Cells(rapCell.Row - 1, 1).Value = "x" And _
              clmRange.Cells(fCell.Row - 1, 1).Value <> "" And _
              wbS.Range("SET.Bestandsnaam").Cells(fCell.Row - 1, 1).Value = fCell.Value Then
           Else
               GoTo ENDstory
           End If
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
    Call ProtectOff
''''
'Define Variables op basis van Role such | Vestiging | UserName | Email | Afdeling |
    Dim RoleCell As Variant
''''''''''''''''''''''''''''
''''''''''''''''''''''''''''
''''''''''''''''''''''''''''
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    Range("A:DZ").EntireColumn.Hidden = False
    Range("1:65000").EntireRow.Hidden = False
    
    Call Generate_Ranges_ALL
    Set Sheet_Cells_All = Range("A5", ActiveCell.SpecialCells(xlLastCell))
''===============================||====================================
''Define Filter Values into ARRAY/RANGE defined in sheet SETTINGS
    Set fssFILT = Nothing
''============================FILTERING================================   moet hier ook (rCell.Row -1, 1)
Call SpeedOn    ''new test
For Each rCell In wbS.Range("SET.RANGE_ALL").Cells
      If rCell.Value <> "" Then
        If floRange(rCell.Row, 1).Value = "Y" Then                 ''SET.Filtering_OUT   (start bij rij 187)
           If clmRange(rCell.Row, 1).Value <> "" Then              ''SET.CSV_x
              'Each Column Filter Value defined in sheet SETTINGS
               Set clmCell = clmRange(rCell.Row, 1)
               
'               If uplRange.Range(rCell.Row, 1) = "X" Then
                  Set fssFILT = wbS.Range("SET.RANGE_ALL").Range("A" & rCell.Row)
'               ElseIf uplRange.Range(rCell.Row, 1) = "Y" Then
'                  Set fssFILT = wbS.Range("SET.ValidateName").Range("A" & rCell.Row)
'               Else
'                  MsgBox "Oeps foutje bij het filteren"
'               End If

                   ColumnCounter = Range(Affix & fssFILT).Column
                   
                    'Functie om te checken of ze teken(s) bevat "<" of ">"     <GreaTer    Lower<
                    'Filterwaarde splitsen dmv. komma teken ","
                  If ContainsGTLT(clmCell) = False Then
                     Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Split(clmCell, ","), Operator:=xlFilterValues
                  Else
                     Dim Crt() As Variant
                     ReDim Crt(0) '(1 to 5)
                     Crt(0) = Split(clmCell, ",")
                    'Crt(1) = Split(clmCell, ",")
                     Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Crt(0)(0), Operator:=xlAnd, Criteria2:=Crt(0)(1)
                  End If
             Else
             End If
        End If
     End If
    rCell = (rCell.Row + 1)
Next rCell

FILTERING_END:
''''''''''''''''''''''''''''''''''''''''''''''''''''
               
Call SpeedOn
                            
Application.EnableEvents = True
Application.ScreenUpdating = True
''============================CREATE WORKSHEET=========================
'Copy all rows from Worksheet "DATA" to Worksheet RAPPORT_UserGroup
    Dim RAPPORT_OUT As String
    Dim Sheet As Worksheet
    Dim bestaat As Boolean
    
   'Naam Rapport_OUT is te vinden in CSV_x rapport, in dezelfde rij als filename
    RAPPORT_OUT = wsLS.Range("SET." & UserGroup).Cells(fCell.Row - 1, 1).Value
            ''WbL wsLS
    'Controle op aanwezigheid op tijdelijke RAPPORT_OUT sheet in Artikelbeheer
    For Each Sheet In Workbooks("ARTIKELBEHEER").Sheets
        If Sheet.Name = RAPPORT_OUT Then bestaat = True: Exit For
    Next Sheet

    If bestaat = True Then
        'Clear the contents of "RAPPORT_OUT" sheet
        Worksheets(RAPPORT_OUT).Select
        If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
        End If
'       Make all Columns / Rows Visible
        Range("A:DZ").EntireColumn.Hidden = False
        Range("1:65000").EntireRow.Hidden = False
        Worksheets(RAPPORT_OUT).Cells.Clear
'       MsgBox = "Het tabblad " & RAPPORT_OUT & " bestaat al."
    
    ElseIf bestaat = False Then
        Worksheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = RAPPORT_OUT
    End If



'Copy the contents of "DATA" sheet  || ArtikelBeheer activate / Select
     wbAS.Activate
     wbAS.Select
Call SpeedOn    ''new test

'Select columns en copier naar RAPPORT_OUT sheet
''===================COLUMNSORT BY NUMBER==============================
For Each rCell In wbS.Range("SET.Upload").Cells
      If rCell.Value <> "" Then
        
         Dim ColumnNr As Integer
         ColumnNr = wsLS.Range("SET." & UserGroup).Cells(rCell.Row - 1, 1).Value
         Dim RangeName As String
         RangeName = wsLS.Range("SET.Range_ALL").Cells(rCell.Row - 1, 1).Value
         Dim SAPName As String
         SAPName = wsLS.Range("SET.VariableName").Cells(rCell.Row - 1, 1).Value
        
        'Optioneel on iets mee te doen
         If rCell.Value = "X" Then
           'Optioneel keuze vaak tussen NL/BE
           
            Range(Affix & RangeName).COPY Destination:=Sheets(RAPPORT_OUT).Cells(2, ColumnNr) ''Range("G:G")     Cells(2, ColumnNr)
            Sheets(RAPPORT_OUT).Cells(1, ColumnNr) = SAPName
             
           'Validation weghalen indien aanwezig (source columns)
            Sheets(RAPPORT_OUT).Cells(2, ColumnNr).SpecialCells(xlCellTypeSameValidation).Validation.Delete
           
         ElseIf rCell.Value = "Y" And _
          True Then
           'Optioneel keuze om een kolom genereren niet slecht een keer
           'Een keer voor een rapport met vaste waarde          (vermelding in CSV_x aanwezig verplicht)
           'Andere keer voor een rapport met gecopieerde range  (anders wordt de waarde niet meegenomen)
           'Voorbeeld Reparatiedelen
           'NETPR (Nettoprijs) in CSV_1 (STAM) = STPRS (Standaardprijs)
           'NETPR (Nettoprijs) in CSV_5 (Repdelen) = 0,01 (Vaste waarde)
           
            Range(Affix & SAPName).COPY Destination:=Sheets(RAPPORT_OUT).Cells(1, ColumnNr) ''Range("G:G")     Cells(2, ColumnNr)
            Sheets(RAPPORT_OUT).Cells(1, ColumnNr) = SAPName
         End If
      
      End If
     
     rCell = (rCell.Row + 1)
Next rCell

     wbAS.Activate
     wbAS.Select
    
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If
''===============================||====================================
'DRAGON filtercolom: This has impact on a new sheet and NOT source sheet
    Worksheets(RAPPORT_OUT).Select
''===============================||====================================
                            'Hier moet nog komen toevoeging van SAP teksten
                            'Dus bestande + nieuwe regel SAp tekst
                            '
Dim numero As Integer
numero = Workbooks(filename).Worksheets.Count
Worksheets(RAPPORT_OUT).Move After:=Workbooks(filename).Sheets(numero) ''Sheets(1)
'                             ActiveSheet.Move Before:=Workbooks("Test.xls").Sheets(1)
Workbooks(filename).Worksheets(RAPPORT_OUT).UsedRange.EntireColumn.AutoFit
''===============================||====================================
     wbAS.Activate
     wbAS.Select
Call SpeedOn    ''new test
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
ENDstory:
Next I
''======================START CREATE FILE END============================
                            Call SpeedOff

'dit onderdeel als laatste van alle stappen toevoegen om volgende file te creeren
     Workbooks(filename).Save
     Workbooks(filename).Close
Next_filename:
     fCell = (fCell.Row + 1)
Next fCell
''========================END CREATE FILE END============================
LaatstRapport:

Call SpeedOff
EnkelRapport = False

       Call SpeedOff
       Application.DisplayAlerts = True
       Application.EnableEvents = True
       Application.ScreenUpdating = True
''===============================||====================================
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    Range("A:DZ").EntireColumn.Hidden = False
    Range("1:65000").EntireRow.Hidden = False

     wbAS.Activate
     wbAS.Select

End Sub


