Attribute VB_Name = "M55_Init_Generator_CSV_oud"
Option Explicit

Public Sub CSV_RAPPORT_oud()
    On Error Resume Next
    
    Application.StatusBar = ""
    Set Network = CreateObject("wscript.network")
'Sheet SETTINGS
    Dim clmInvisible As Range   ''C_olumn_zichtbaar
    Dim clmRange As Range       ''C_olumn_User
    Dim clhRange As Range       ''C_olumn LAY_out
    Dim clwRange As Range       ''C_oLumn W_idth
    Dim clmRangeMail As Range   ''C_olumn_User_Mail
    
    Dim pbrRange As Range       ''Column Page_Break
    Dim fltRange As Range       ''Column Filtering
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
    Call SpeedOn
    Call Affix_Case
''===============================||====================================
    Set wbL = Workbooks("Lijsten_new.xlsm")
    Set wbS = wbL.Worksheets("SETTINGS")
    Set clhRange = wbS.Range("SET.ColumnHide")
    Set pbrRange = wbS.Range("SET.PageBreak")
    Set clwRange = wbS.Range("SET.ColumnWidth")
   'Set fltRange = wbS.Range("SET.Filtering")
    Set srtRange = wbS.Range("SET.Sorting")
    Set pstRange = wbS.Range("SET.PageSetup")
   'Set atrRange = wbS.Range("SET.Attributes")
    Set abgRange = wbS.Range("SET.Accoordbedrag")
    Set uplRange = wbS.Range("SET.Upload")
    Set upwRange = wbS.Range("SET.Upload_waarde")
''===============================||====================================
'    Workbooks("Artikelbeheer.xlsm").Sheets("OUT").Visible = xlSheetVisible
'    Workbooks("Artikelbeheer.xlsm").Sheets("OUT").Activate
'    Workbooks("Artikelbeheer.xlsm").Sheets("OUT").Select
''===============================||====================================
Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================





''===============================||====================================
''===============================||====================================
''===============================||====================================
'Fill new created range with values

For Each rCell In uplRange.Cells            ''SET.Upload
      If rCell.Value = "Y" Then
        'Define Column Heading (Name)
         Cells(1, rCell.Row - 1).Select
         Cells(1, rCell.Row - 1).Value = wbS.Range("SET.VariableName").Cells(rCell.Row - 1, 1).Value
'STAP 1 schrijf alle waarde van kolom SET.Upload
'STAP 2 "COPY"  waarden van genoemde kolom in SET.Upload
'STAP 3 "NL-BE" waarde zetten afhankelijk van kolom in SET.Upload
'STAP 4 "NL-BE" waarde zetten afhankelijk van kolom in SET.Upload

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
'COPY
''''''''''''''''''''''''''''''''
          If wbS.Range("SET.Upload_waarde").Cells(rCell.Row - 1, 1).Value = "COPY" Then
            ''format meecopieren
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      rCelly.Value = Range(Affix & IdemDito).Cells(rCelly.Row - HeadingRows, 1).Value
                 Next rCelly
         End If
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
           
'        If COPY_Range = "VPRSV" Then
'             For Each rCelly In Range(Affix & COPY_Range).Cells
'
'                      rCelly.Select
'                      rCelly.Value = ""
'                      If Range("OUT_Reparatiedeel").Cells(rCelly.Row - HeadingRows, 1).Value = "Ja" Then
'                         rCelly.Value = "V"
'                      Else
'                         rCelly.Value = ""
'                     End If
'                 Next rCelly
'         End If
           
     
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
'Correctie:
        If COPY_Range = "VERPR" Then
             For Each rCelly In Range(Affix & COPY_Range).Cells
                      
                      rCelly.Select
                      rCelly.Value = ""
                      If Range("OUT_Reparatiedeel").Cells(rCelly.Row - HeadingRows, 1).Value = "Ja" Then
                         rCelly.Value = "0,01"
                      Else
                         rCelly.Value = ""
                     End If
                 Next rCelly
         End If
      
'''''''''''''''''''''''''''''''''
'For Each rCelly In Range(Affix  & Cells(1, rCell.Row - 1)).Cells
'                      rCell.Value = "test"
'                 Next rCelly
'''''             Cells(rCell.Row - 1, 1).Cells = Range(Affix  & COPY_Range).Cells(rCell.Row - 1, 1)
            ''Range(Affix  & Cells(1, rCell.Row - 1)) = wbS.Range("SET.Range_ALL").Cells(rCell.Row - 1, 1).Value
'         End If
        
'         If Range(Affix  & Cells(1, rCell.Row - 1)) <> "" Then
'           ''Range(Affix  & Range_ALL(rCell.Row - 1, 1)).Cells = Range("SET.Upload_waarde"(rCell.Row - 1, 1))
'
'         End If
         
         
         
      End If
     rCell = (rCell.Row + 1)
'If COPY_Range <> "" Then
'   Range(Affix & COPY_Range).EntireColumn.NumberFormat = "@"
'   Range(Affix & COPY_Range).EntireColumn.AutoFit
'
'   Range(Affix & COPY_Range).EntireColumn.Validation.Delete
'   Range(Affix & COPY_Range).SpecialCells(xlCellTypeSameValidation).Validation.Delete
'End If
'    ActiveCell.SpecialCells(xlCellTypeSameValidation).Select
'    With Selection.Validation
'        .Delete
'        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
'        :=xlBetween
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .ShowInput = True
'        .ShowError = True
'    End With
       
       'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
       'SkipBlanks:=False, Transpose:=False
'        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'            SkipBlanks:=False, Transpose:=False


Next rCell
'''''''''''''''''''''''''''''''''
'6de rij leegmaken van alle waardes
Worksheets("OUT").Rows(6).Select
Worksheets("OUT").Rows(6).ClearContents
'totdat ik iets beters bedenk / even geen tijd
''===============================||====================================
''ALLE KOLOMEN ZIJN GECREERD
''ALLE KOLOMEN ZIJN GECREERD
''ALLE KOLOMEN ZIJN GECREERD
''ALLE KOLOMEN ZIJN GECREERD
''ALLE KOLOMEN ZIJN GECREERD
''ALLE KOLOMEN ZIJN GECREERD
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
'Define Rapport Set via rapCell
UserGroup = ""
Application.StatusBar = ""
For Each rapCell In wbS.Range("SET.RANGE_ALL").Cells
      If rapCell.Value = ActiveSheet.Name Then
      Set wbR = Workbooks(ActiveWorkbook.Name)
      Set wbRS = wbR.Worksheets(ActiveSheet.Name)
         GoTo Rapport_sheet
      End If
     rapCell = (rapCell.Row + 1)
Next rapCell
Rapport_sheet:
''===============================||====================================
    Call Apply_UserNames
''===============================||====================================
'   workbook RapportSheet          Make all Collumns / Rows Visible
    wbRS.Activate
    Call Affix_Case
''===============================||====================================
'       Set mailRange = wbS.Range("SET.Mailing_OUT")
       Set fltRange = wbS.Range("SET.Filtering_OUT")
      'ColumnCounter = Range(Affix & "Aanvrager").Column
      'Define First User Column
       For UserColumn = 1 To wbS.Cells(1, 1).SpecialCells(xlLastCell).Column
           If wbS.Cells(1, UserColumn).Value = "Master" Then
              FirstUserColumn = UserColumn + 1
           ElseIf wbS.Cells(1, UserColumn).Value = "Bestandsnaam" Then
              LastUserColumn = UserColumn - 1
           End If
       Next
''===============================||====================================
''Define File Name to save a sheets accoording to Values in RANGE("Bestandsnaam") in sheet SETTINGS
''============================CREATE FILE================================
Call SpeedOn    ''new test
For Each fCell In wbS.Range("SET.Bestandsnaam").Cells
      If fCell.Value <> "" Then
      
            'Makes a copy of the active sheet and save it to a temporary file
             Dim wbUpload As Worksheet ''Workbook
             Dim wsUpload As Worksheet

             filename = fCell.Value & ".xlsx"  ''" " & Format(Now, "d mmm yyyy hhnn").Value & ".xlsx"                ''= fCell.Value & ".xlsx"
             '==============================================================
             'Replace existing files
              Application.DisplayAlerts = False   'replacing
             Application.SheetsInNewWorkbook = 1
             Workbooks.Add          ''(After:=Worksheets(Worksheets.Count)).Name = "Master"
             ActiveSheet.Name = "Master"
                 ActiveWorkbook.SaveAs "C:\Temp\" & filename, FileFormat:=51
'                 ActiveWorkbook.Close
Dim WinShuttle_Name As Integer
WinShuttle_Name = fCell.Row

     Else: GoTo Next_filename
     End If
      
     
     
     
     
     
     
     
     
     
     
''===============================||====================================
''Define Filter Values into ARRAY/RANGE defined in sheet SETTINGS
'    Set fssFILT = Nothing
''============================FILTERING================================   moet hier ook (rCell.Row -1, 1)
'Akkoord voor MAILING
Dim AccMAILING As Boolean
Dim OutMAILING As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim I
    Dim j
    Dim RapportSelection As String
    I = 0
    j = 0
     
     For I = FirstUserColumn To LastUserColumn
           RapportSelection = wbS.Cells(1, I).Value
           UserGroup = RapportSelection     'CSV_group
           Set clmRange = wbS.Range("SET." & RapportSelection) ''UserGroup)
'''''''''''
           If clmRange.Cells(rapCell.Row - 1, 1).Value = "x" And _
              clmRange.Cells(fCell.Row - 1, 1).Value <> "" And _
              wbS.Range("SET.Bestandsnaam").Cells(fCell.Row - 1, 1).Value = fCell.Value Then
           Else
               GoTo ENDstory
           End If
'''''''''''
Application.StatusBar = ". . . . . . . User Rapport: " & " : " & UserGroup
''===============================||====================================
If Niveau = 1 Then  ''UserGroup = "DB" Or UserGroup = "MMP" Then    ''
   Dim Antwoord As Integer
   Dim AantalRapporten As Integer
   
   Dim Bijlage_toevoegen As Integer
   Dim Bijlage_bestandsnaam As String
  'Dim AntwoordVestiging As Integer
  
'   AantalRapporten = 1
   
'   Antwoord = MsgBox("Wil je Rapport voor " & UserGroup & " creeren?", vbYesNo + vbQuestion, "RAPORT filteren")
'                 If Antwoord = vbYes Then
'                 ElseIf Antwoord = vbNo Then
'                        GoTo ENDstory
'                 End If
Else
  'Je mag alleen voor eigen profiel Rapport maken    UserGroup = Role
End If
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
wbRS.Activate
''Worksheets("OUT").Select
    Call ProtectOff
''''
'Define Variables op basis van Role such | Vestiging | UserName | Email | Afdeling |
    Dim RoleCell As Variant
    
''''''''''''''''''''''''''''
''''''''''''''''''''''''''''
''''''''''''''''''''''''''''

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
        If fltRange(rCell.Row, 1).Value = "Y" Then                 ''SET.Filtering_OUT   (start bij rij 187)
           If clmRange(rCell.Row, 1).Value <> "" Then              ''SET.CSV_x
              'Each Column Filter Value defined in sheet SETTINGS
               Set clmCell = clmRange(rCell.Row, 1)
               Set fssFILT = wbS.Range("SET.RANGE_ALL").Range("A" & rCell.Row)
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

''''''''
''HIER ZOU NOG SCHEIDING TUSSEN NL EN BE ARTIKELEN FILTER KUNNEN KOMEN
''IF Aanvraagbestand_BE    filter op BE-artikelen
''IF Aanvraagbestand_NL    filter op NL-artikelen
''''''''
'    If ActiveSheet.Name = "Accordering" Then
'                        ColumnCounter = Range(Affix & "Aanvrager.Vestiging").Column
'                        Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Vestiging_var, Operator:=xlFilterValues
'
'                        If UserGroup = "ME" Or UserGroup = "MMR" Then
'                           Afdeling_var = wsLU.Range("USER.Afdeling").Cells(RoleCell.Row - 1, 1)
'                           ColumnCounter = Range(Affix & "Aanvrager.Afdeling").Column
'                           Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Afdeling_var, Operator:=xlFilterValues
'                        End If
'
'ElseIf ActiveSheet.Name = "OUT" Then
'                        ColumnCounter = Range(Affix & "Aanvrager.Naam").Column
'                        Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Aanvrager_var, Operator:=xlFilterValues
'
'                        If UserGroup = "ME" Then
'                           ColumnCounter = Range(Affix & "Aanvrager.Afdeling").Column
'                           Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Afdeling_var, Operator:=xlFilterValues
'
'                           ColumnCounter = Range(Affix & "Aanvrager.Vestiging").Column
'                           Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Vestiging_var, Operator:=xlFilterValues
'                        End If
'
'End If

FILTERING_END:
'resultaat filtering >0 doorgaan anders stoppen
'Set rngUserGroup = ActiveSheet.AutoFilter.Range
'If (rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) = 0 Then GoTo ENDstory     ''Andere_Vestiging
''===========================COLUMNHIDE================================
Call SpeedOn
'    Set clhRange = wbS.Range("SET.Upload")
'    Set clhRange = wbS.Range("SET." & UserGroup)

For Each rCell In wbS.Range("SET.RANGE_ALL").Cells
      If rCell.Value <> "" Then
           If clhRange(rCell.Row - 1, 1).Value <> "" Then   ''="x" bestaande   || ="y"  nieuwe
           
                  If clmRange(rCell.Row - 1, 1).Value = "" Then                             ''HIDDEN LOCKED
                     Range(Affix & rCell).Locked = False
                     Range(Affix & rCell).EntireColumn.Hidden = True
              ElseIf clmRange(rCell.Row - 1, 1).Value <> "" Then                            ''HIDDEN LOCKED
                     Range(Affix & rCell).Locked = False
                     Range(Affix & rCell).EntireColumn.Hidden = False
              End If
                 
           End If
    End If

     rCell = (rCell.Row + 1)
Next rCell
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim rngUserGroup As Range
'Set rngUserGroup = ActiveSheet.AutoFilter.Range
'
'If (rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) = 0 Then GoTo Andere_Vestiging ''ENDstory
'
'If rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1 > 0 Then
'
'
'    If AccMAILING = False Then
'      'Geen mailing
'       MsgBox "Rapport gefilterd voor:  " & vbNewLine & vbNewLine & _
'              "Gebruikersrol: " & UserGroup & vbNewLine & vbNewLine & _
'              "Afdeling:      " & Afdeling_var & vbNewLine & vbNewLine & _
'              "Vestiging:     " & Vestiging_var & vbNewLine & vbNewLine & _
'              "Regels:        " & rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
'
'    Else
'       Antwoord = MsgBox("Wil je Rapportlink voor " & UserGroup & " " & Afdeling_var & "  -  " & Vestiging_var & _
'                        "   ( " & (rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) & " regels)   " & " emailen?", _
'                        vbYesNo + vbQuestion, "RAPORT link emailen")
'       If Antwoord = vbYes Then
''===============================||====================================
''===============================||====================================
''===============================||====================================
'                        Bijlage_toevoegen = MsgBox("Wil je Rapport als bijlage " & _
'                                            UserGroup & " creeren?", _
'                                            vbYesNo + vbQuestion, "RAPPORT bijlage creeren")
'                        If Bijlage_toevoegen = vbYes Then
                             
                           Call SpeedOn
                            
Application.EnableEvents = True
Application.ScreenUpdating = True
''============================CREATE WORKSHEET=========================
'Copy all rows from Worksheet "DATA" to Worksheet RAPPORT_UserGroup
    Dim RAPPORT_OUT As String
    Dim Sheet As Worksheet
    Dim bestaat As Boolean
    
  ''  set WRSA
    
    RAPPORT_OUT = wsLS.Range("SET." & UserGroup).Cells(WinShuttle_Name - 1, 1).Value
            ''WbL wsLS
    For Each Sheet In Workbooks("ARTIKELBEHEER").Sheets ''ThisWorkbook.Sheets
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
'   if error aantal tekens tellen max 31 tekens
    End If
'Copy the contents of "DATA" sheet
wbRS.Activate
wbRS.Select





''''''''''''''''Range("A1", ActiveCell.SpecialCells(xlLastCell)).Select
''''''''''''''''Selection.COPY
''''''''''''''''
''''''''''''''''
''''''''''''''''Worksheets(RAPPORT_OUT).Select
''''''''''''''''Range("A1").Select
''''''''''''''''
''''''''''''''''Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
''''''''''''''''SkipBlanks:=False, Transpose:=False
''''''''''''''''Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
''    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
''        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False

'    ActiveWorkbook.Worksheets("Accordering").UsedRange.COPY Destination:=ActiveWorkbook.Worksheets(RAPPORT_OUT).Range("A1")
'Worksheets("DATA").Visible = xlSheetHidden
''===============================||====================================
'    Set clhRange = wbS.Range("SET.Upload")
'    Set clhRange = wbS.Range("SET." & UserGroup)

'Select columns en copier naar RAPPORT_OUT sheet
''===================COLUMNSORT BY NUMBER==============================
For Each rCell In wbS.Range("SET.Upload").Cells
      If rCell.Value <> "" Then

         If rCell.Value = "x" Then

     ElseIf rCell.Value = "Y" Then

         End If

           If wsLS.Range("SET." & UserGroup).Cells(rCell.Row - 1, 1).Value <> "" Then   ''="x" bestaande   || ="y"  nieuwe
              Dim ColumnNr As Integer
              ColumnNr = wsLS.Range("SET." & UserGroup).Cells(rCell.Row - 1, 1).Value
              Dim RangeName As String
              RangeName = wsLS.Range("SET.Range_ALL").Cells(rCell.Row - 1, 1).Value
              Dim SAPName As String
              SAPName = wsLS.Range("SET.VariableName").Cells(rCell.Row - 1, 1).Value
              
              Range(Affix & RangeName).COPY Destination:=Sheets(RAPPORT_OUT).Cells(2, ColumnNr) ''Range("G:G")
              Sheets(RAPPORT_OUT).Cells(1, ColumnNr) = SAPName

'Validation weghalen indien aanwezig (source columns)
'Sheets(RAPPORT_OUT).Cells(2, ColumnNr).EntireColumn.Validation.Delete
Sheets(RAPPORT_OUT).Cells(2, ColumnNr).SpecialCells(xlCellTypeSameValidation).Validation.Delete
'    ActiveCell.SpecialCells(xlCellTypeSameValidation).Select
'    With Selection.Validation
'        .Delete
'        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
'        :=xlBetween
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .ShowInput = True
'        .ShowError = True
'    End With


'              Collumn(ColumnNr).COPY
              
'                  If clmRange(rCell.Row - 1, 1).Value = "" Then                             ''HIDDEN LOCKED
'                     Range(Affix & rCell).Locked = True
'                     Range(Affix & rCell).EntireColumn.Hidden = True
'              ElseIf clmRange(rCell.Row - 1, 1).Value <> "" Then                            ''HIDDEN LOCKED
'                     Range(Affix & rCell).Locked = False
'                     Range(Affix & rCell).EntireColumn.Hidden = False
'              End If

           End If
    End If

     rCell = (rCell.Row + 1)
Next rCell


''===============================||====================================
'DRAGON filtercolom: This has impact on a new sheet and NOT source sheet
    Worksheets(RAPPORT_OUT).Select
'    Range("A1:DZ65000").Select
'    Selection.AutoFilter Field:=1, Criteria1:="<>"
'    ActiveSheet.Rows("2:5").Delete

    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If
''===============================||====================================
'SpeedUp the macro
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
                            
                            
                            'Hier moet nog komen toevoeging van SAP teksten
                            'Dus bestande + nieuwe regel SAp tekst
                            '
Dim numero As Integer
numero = Workbooks(filename).Worksheets.Count
Worksheets(RAPPORT_OUT).Move After:=Workbooks(filename).Sheets(numero) ''Sheets(1)
'                             ActiveSheet.Move Before:=Workbooks("Test.xls").Sheets(1)
Workbooks(filename).Worksheets(RAPPORT_OUT).UsedRange.EntireColumn.AutoFit

'If COPY_Range <> "" Then
'   Range(Affix & COPY_Range).EntireColumn.NumberFormat = "@"
'   Range(Affix & COPY_Range).EntireColumn.AutoFit
'
'   Range(Affix & COPY_Range).EntireColumn.Validation.Delete
'   Range(Affix & COPY_Range).SpecialCells(xlCellTypeSameValidation).Validation.Delete
'End If
'    End If
                            
                            
                            
                            
                            
                            
''OKOKOKOKOKOK
'                        Else
'                        End If
''===============================||====================================
''rngHTMLshort
''Worksheets("RAPPORT_" & UserGroup).activate
          wbRS.Activate
''===============================||====================================
''===========================COLUMNHIDE================================
Call SpeedOff
'    If ActiveSheet.Name = "Accordering" Then
'       Set clmRangeMail = wbS.Range("SET.Mail_ACC")
'ElseIf ActiveSheet.Name = "OUT" Then
'       Set clmRangeMail = wbS.Range("SET.Mail_OUT")
'End If
'
'For Each rCell In wbS.Range("SET.Range_ALL").Cells
'      If rCell.Value <> "" Then
'           If clhRange(rCell.Row - 1, 1).Value <> "" Then
'
'                   If clmRangeMail(rCell.Row - 1, 1).Value = "" Then
'                      Range(Affix & rCell).EntireColumn.Hidden = True
'               ElseIf clmRangeMail(rCell.Row - 1, 1).Value = "x" Then
'                      Range(Affix & rCell).EntireColumn.Hidden = False
'               End If
'
'           End If
'    End If
'
'     rCell = (rCell.Row + 1)
'Next rCell
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
          wbRS.Activate
'    Worksheets(RAPPORT_OUT).Select
'    Range("A1:DZ65000").Select
'    Selection.AutoFilter Field:=1, Criteria1:="<>"
'    ActiveSheet.Rows("2:5").EntireRow.Hidden = True
'    ActiveSheet.SpecialCells(xlCellTypeVisible).Select
'    Set rngHTML = Selection.SpecialCells(xlCellTypeVisible)
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''===============================||====================================
'                            Application.DisplayAlerts = False    'replacing
'          wbRS.Activate
'               ActiveSheet.Rows("2:5").EntireRow.Hidden = False
'          Worksheets("RAPPORT_" & UserGroup).Delete
     
                            Application.DisplayAlerts = True    'replacing
                            Application.EnableEvents = True
                            Application.ScreenUpdating = True
         
Application.StatusBar = ". . . . . . . User Rapport: " & " : " & UserGroup
''===============================||====================================
''===============================||====================================
''===============================||====================================
''''''''''''''''''''''''''''''''''
'    EnkelRapport = True








If EnkelRapport = True Then
   GoTo LaatstRapport

   'Stoppen om rapport te bekijken
    Antwoord = MsgBox("Wil je verder met Rapport creeren?", vbYesNo + vbQuestion, "RAPPORTEREN STOPPEN")
        If Antwoord = vbYes Then
    ElseIf Antwoord = vbNo Then
           GoTo LaatstRapport
    End If
''''''''''''''''''''''''''''''''''
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Else
'End If

'Andere_Vestiging:
'Next RoleCell
''===============================||====================================
ENDstory:
                            Call SpeedOff
                            Application.DisplayAlerts = True    'replacing
                            Application.EnableEvents = True
                            Application.ScreenUpdating = True
Next I

''======================START CREATE FILE END============================

'dit onderdeel als laatste van alle stappen toevoegen om volgende file te creeren
Next_filename:
     Workbooks(filename).Save
     Workbooks(filename).Close
     fCell = (fCell.Row + 1)
Next fCell
''========================END CREATE FILE END============================









LaatstRapport:
Call ProtectOn
Call SpeedOff
EnkelRapport = False

                            Call SpeedOff
                            Application.DisplayAlerts = True    'replacing
                            Application.EnableEvents = True
                            Application.ScreenUpdating = True
''===============================||====================================
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    Range("A:DZ").EntireColumn.Hidden = False
    Range("1:65000").EntireRow.Hidden = False
End Sub
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================

