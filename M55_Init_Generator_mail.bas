Attribute VB_Name = "M55_Init_Generator_mail"
Option Explicit

Public Sub Mailing_RAPPORT()
    On Error Resume Next
    
    Application.StatusBar = ""
    Set Network = CreateObject("wscript.network")
'Sheet SETTINGS
    Dim clmInvisible As Range   ''C_olumn_zichtbaar
    Dim clmRange As Range       ''C_olumn_User
    Dim clhRange As Range       ''C_olumn LAY_out
    Dim clwRange As Range       ''C_oLumn W_idth
    Dim clmRangeMail As Range       ''C_olumn_User_Mail
    
    Dim pbrRange As Range       ''Column Page_Break
    Dim fltRange As Range       ''Column Filtering
    Dim srtRange As Range       ''Column Sorting
    Dim pstRange As Range       ''Page Setup
    Dim atrRange As Range       ''Attributes
    Dim molRange As Range       ''Column MailOnly
    Dim mailRange As Range      ''Column Mailing
    Dim abgRange As Range       ''Column Accoordbedrag
    
'Initialise Invisible Data Collection
    Dim rRange As Range
    Dim rCell As Variant
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
''===============================||====================================
'    Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Visible = xlSheetVisible
'    Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Activate
'    Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Select
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
''===============================||====================================
    Call Apply_UserNames
''===============================||====================================
'   workbook RapportSheet          Make all Collumns / Rows Visible
    wbRS.Activate
    Call Affix_Case
''===============================||====================================
'    If Workbooks("Artikelbeheer.xlsm").ActiveSheet.Name <> "Accordering" And _
'       Workbooks("Artikelbeheer.xlsm").ActiveSheet.Name <> "OUT" Then
'
'
'
'
'
'
'                        If Role = "ME" Then
'                           Workbooks("Artikelbeheer.xlsm").Activate
'                           Worksheets("OUT").Select
'                           Set mailRange = wbS.Range("SET.Mailing_OUT")
'                           Set fltRange = wbS.Range("SET.Filtering_OUT")
'                          'ColumnCounter = Range(Affix & "Aanvrager").Column
'                        End If
'                        If Role <> "ME" Then
'                           Workbooks("Artikelbeheer.xlsm").Activate
'                           Worksheets("Accordering").Select
'                           Set mailRange = wbS.Range("SET.Mailing_ACC")
'                           Set fltRange = wbS.Range("SET.Filtering_ACC")
'                          'ColumnCounter = Range(Affix & "Vestiging").Column
'                        End If
'    End If
'
'
'
'
    If ActiveSheet.Name = "Accordering" Then
       Set mailRange = wbS.Range("SET.Mailing_ACC")
       Set fltRange = wbS.Range("SET.Filtering_ACC")
      'ColumnCounter = Range(Affix & "Vestiging").Column
      'Define First User Column
       For UserColumn = 1 To wbS.Cells(1, 1).SpecialCells(xlLastCell).Column
           If wbS.Cells(1, UserColumn).Value = "FSB" Then
              FirstUserColumn = UserColumn
           ElseIf wbS.Cells(1, UserColumn).Value = "CSV_1" Then
              LastUserColumn = UserColumn - 1
           End If
       Next
       
ElseIf ActiveSheet.Name = "OUT" Then
       Set mailRange = wbS.Range("SET.Mailing_OUT")
       Set fltRange = wbS.Range("SET.Filtering_OUT")
      'ColumnCounter = Range(Affix & "Aanvrager").Column
      'Define First User Column
       For UserColumn = 1 To wbS.Cells(1, 1).SpecialCells(xlLastCell).Column
           If wbS.Cells(1, UserColumn).Value = "CSV_1" Then
              FirstUserColumn = UserColumn
           ElseIf wbS.Cells(1, UserColumn).Value = "CSV_END" Then
              LastUserColumn = UserColumn - 1
           End If
       Next
       
End If
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Akkoord voor MAILING
Dim AccMAILING As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim I
    Dim j
    Dim RapportSelection As String
'    Dim clmRangeSelection As String
    I = 0
    j = 0
     
     For I = FirstUserColumn To LastUserColumn
           RapportSelection = wbS.Cells(1, I).Value
'           clmRangeSelection = "SET." & RapportSelection
            
    If ActiveSheet.Name = "Accordering" Then
            If Niveau = 1 Then               ''UserGroup = "DB" Or UserGroup = "MMP" Then
               AccMAILING = True    ''rapMAILING
               UserGroup = RapportSelection
               Set clmRange = wbS.Range("SET." & UserGroup)
            Else
               AccMAILING = False
               UserGroup = Role
               Set clmRange = wbS.Range("SET." & UserGroup)
               EnkelRapport = True
            End If
ElseIf ActiveSheet.Name = "OUT" Then
              'If Niveau = 1 Then
               AccMAILING = False
               UserGroup = RapportSelection
               Set clmRange = wbS.Range("SET." & UserGroup)
End If
'''''''''''
'          If wbS.Range(clmRangeSelection).Cells(rapCell.Row - 1, 1).Value = "" Then
           If clmRange.Cells(rapCell.Row - 1, 1).Value = "" Then
              GoTo ENDstory
           Else
           End If
'''''''''''
Application.StatusBar = ". . . . . . . User Rapport: " & " : " & UserGroup
''===============================||====================================
If Niveau = 1 Then  ''UserGroup = "DB" Or UserGroup = "MMP" Then    ''
  'Je mag voor alle profielen, per profiel Rapport maken
   Dim Antwoord As Integer
   Dim AantalRapporten As Integer
   
   Dim Bijlage_toevoegen As Integer
   Dim Bijlage_bestandsnaam As String
  'Dim AntwoordVestiging As Integer
  
   If ActiveSheet.Name = "Accordering" Then
      AantalRapporten = Application.CountIf(wsLU.Range("USER.Role"), UserGroup)
   ElseIf ActiveSheet.Name = "OUT" Then
      AantalRapporten = 1
   End If
   
   Antwoord = MsgBox("Wil je " & AantalRapporten & "x Rapport voor " & UserGroup & " creeren?", vbYesNo + vbQuestion, "RAPORT filteren")
                 If Antwoord = vbYes Then
                 ElseIf Antwoord = vbNo Then
                        GoTo ENDstory
                 End If
Else
  'Je mag alleen voor eigen profiel Rapport maken    UserGroup = Role
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
    
For Each RoleCell In wsLU.Range("USER.Role").Cells
''''''''''''''''''''''''''''
''''''''''''''''''''''''''''
''''''''''''''''''''''''''''
If Niveau = 1 Then      ''UserGroup = "DB" Or UserGroup = "MMP" Then    ''
       Dim Vestiging_var As Variant
       Dim Aanvrager_var As Variant
       Dim UserEmail_var As Variant
       Dim Afdeling_var As Variant
    If RoleCell.Value = UserGroup Then
       Vestiging_var = wsLU.Range("USER.Vestiging").Cells(RoleCell.Row - 1, 1)
       Aanvrager_var = wsLU.Range("USER.Naam").Cells(RoleCell.Row - 1, 1) ''new
       UserEmail_var = wsLU.Range("USER.Email").Cells(RoleCell.Row - 1, 1)
       Afdeling_var = wsLU.Range("USER.Afdeling").Cells(RoleCell.Row - 1, 1) ''new
    Else
       GoTo Andere_Vestiging
    End If
Else
       Aanvrager_var = wsLU.Range("USER.Naam").Cells(RoleCell.Row - 1, 1) ''new
       If Aanvrager_var = Naam Then
          Vestiging_var = wsLU.Range("USER.Vestiging").Cells(RoleCell.Row - 1, 1)
          UserEmail_var = wsLU.Range("USER.Email").Cells(RoleCell.Row - 1, 1)
          Afdeling_var = wsLU.Range("USER.Afdeling").Cells(RoleCell.Row - 1, 1) ''new
       Else
          GoTo Andere_Vestiging
       End If
End If
''''''''''''''''''''''''''''
''''''''''''''''''''''''''''
''''''''''''''''''''''''''''
USERMAIL = UserEmail_var
''''
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
        If fltRange(rCell.Row, 1).Value = "Y" Then
           If clmRange(rCell.Row, 1).Value <> "" Then
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
                    
    If ActiveSheet.Name = "Accordering" Then
                        ColumnCounter = Range(Affix & "Aanvrager.Vestiging").Column
                        Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Vestiging_var, Operator:=xlFilterValues
                        
                        If UserGroup = "ME" Or UserGroup = "MMR" Then
                           Afdeling_var = wsLU.Range("USER.Afdeling").Cells(RoleCell.Row - 1, 1)
                           ColumnCounter = Range(Affix & "Aanvrager.Afdeling").Column
                           Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Afdeling_var, Operator:=xlFilterValues
                        End If

ElseIf ActiveSheet.Name = "OUT" Then
                        ColumnCounter = Range(Affix & "Aanvrager.Naam").Column
                        Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Aanvrager_var, Operator:=xlFilterValues
                        
                        If UserGroup = "ME" Then
                           ColumnCounter = Range(Affix & "Aanvrager.Afdeling").Column
                           Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Afdeling_var, Operator:=xlFilterValues

                           ColumnCounter = Range(Affix & "Aanvrager.Vestiging").Column
                           Sheet_Cells_All.AutoFilter Field:=ColumnCounter, Criteria1:=Vestiging_var, Operator:=xlFilterValues
                        End If

End If

FILTERING_END:
'resultaat filtering >0 doorgaan anders stoppen
'Set rngUserGroup = ActiveSheet.AutoFilter.Range
'If (rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) = 0 Then GoTo ENDstory     ''Andere_Vestiging
''===========================COLUMNHIDE================================
Call SpeedOn
For Each rCell In wbS.Range("SET.RANGE_ALL").Cells
      If rCell.Value <> "" Then
           If clhRange(rCell.Row - 1, 1).Value <> "" Then
           
                  If clmRange(rCell.Row - 1, 1).Value = "" Then                             ''HIDDEN LOCKED
                     Range(Affix & rCell).Locked = True
                     Range(Affix & rCell).EntireColumn.Hidden = True
              ElseIf clmRange(rCell.Row - 1, 1).Value = "H" Then                            ''HIDDEN LOCKED
                     Range(Affix & rCell).Locked = False
                     Range(Affix & rCell).EntireColumn.Hidden = True
              ElseIf clmRange(rCell.Row - 1, 1).Value = "R" Then
                     Range(Affix & rCell).Locked = True
                     Range(Affix & rCell).EntireColumn.Hidden = False
              ElseIf clmRange(rCell.Row - 1, 1).Value = "W" Then
                     Range(Affix & rCell).Locked = False
                     Range(Affix & rCell).EntireColumn.Hidden = False
              End If
                 
           End If
    End If

     rCell = (rCell.Row + 1)
Next rCell
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim rngUserGroup As Range
Set rngUserGroup = ActiveSheet.AutoFilter.Range

If (rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) = 0 Then GoTo Andere_Vestiging ''ENDstory

If rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1 > 0 Then


    If AccMAILING = False Then
      'Geen mailing
       MsgBox "Rapport gefilterd voor:  " & vbNewLine & vbNewLine & _
              "Gebruikersrol: " & UserGroup & vbNewLine & vbNewLine & _
              "Afdeling:      " & Afdeling_var & vbNewLine & vbNewLine & _
              "Vestiging:     " & Vestiging_var & vbNewLine & vbNewLine & _
              "Regels:        " & rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    
    Else
       Antwoord = MsgBox("Wil je Rapportlink voor " & UserGroup & " " & Afdeling_var & "  -  " & Vestiging_var & _
                        "   ( " & (rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) & " regels)   " & " emailen?", _
                        vbYesNo + vbQuestion, "RAPORT link emailen")
       If Antwoord = vbYes Then
''===============================||====================================
''===============================||====================================
''===============================||====================================
                        Bijlage_toevoegen = MsgBox("Wil je Rapport als bijlage voor " & _
                                            UserGroup & " " & Afdeling_var & "  -  " & Vestiging_var & _
                                            "   ( " & (rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) & _
                                            " regels)   " & " in de mail toevoegen?", _
                                            vbYesNo + vbQuestion, "RAPPORT bijlage emailen")
                        If Antwoord = vbYes Then
                             
                            Call SpeedOn
                            
                           Call Bijlage_bestand
                           'Makes a copy of the active sheet and save it to a temporary file
                            Dim wbRapport As Worksheet
'                            Dim Path_C As String
                            Worksheets("RAPPORT_" & UserGroup).COPY
                            Set wbRapport = Worksheets("RAPPORT_" & UserGroup)
                            ''filename = "RAPPORT_" & UserGroup & ".xlsx"
                            filename = "RAPPORT_" & UserGroup & " " & Aanvrager_var & " " & Format(Now, "d mmm yyyy   hh.nn") & ".xlsx"
'                            Path_C = "C:\Temp\"
                            '==============================================================
                            'Replace existing files
                             Application.DisplayAlerts = False   'replacing
                                wbRapport.SaveAs Path_C & filename, FileFormat:=51
                                ActiveWorkbook.Close
                              'MsgBox "Look in <" & Path & "> ... for the files!" & vbNewLine & "Look if map <" & Path & "> is present!"
                               'ActiveWorkbook.SaveAs "C:\ron.xlsm", FileFormat:=52
                               '50 = xlExcel12 (Excel Binary Workbook in 2007-2013 with or without macro's, xlsb)
                               '51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
                               '52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
                               '56 = xlExcel8 (97-2003 format in Excel 2007-2013, xls)
                               'FileExtStr = ".csv": FileFormatNum = 6
                               'FileExtStr = ".txt": FileFormatNum = -4158
                               'FileExtStr = ".prn": FileFormatNum = 36
                            '==============================================================
'                            Set rngHTML = Nothing
'                            On Error Resume Next
'                                           Worksheets("RAPPORT_" & UserGroup).SpecialCells(xlCellTypeVisible).Select
'                             Set rngHTML = Selection.SpecialCells(xlCellTypeVisible)
'                            If rngHTML Is Nothing Then
'                               MsgBox "The selection is not a range or the sheet is protected" & _
'                                       vbNewLine & "please correct and try again.", vbOKOnly
'                            Exit Sub
'                            End If
                            '==============================================================
                        Else
                        End If
''===============================||====================================
''rngHTMLshort
''Worksheets("RAPPORT_" & UserGroup).activate
          wbRS.Activate
''===============================||====================================
''===========================COLUMNHIDE================================
Call SpeedOff
    If ActiveSheet.Name = "Accordering" Then
       Set clmRangeMail = wbS.Range("SET.Mail_ACC")
ElseIf ActiveSheet.Name = "OUT" Then
       Set clmRangeMail = wbS.Range("SET.Mail_OUT")
End If

For Each rCell In wbS.Range("SET.Range_ALL").Cells
      If rCell.Value <> "" Then
           If clhRange(rCell.Row - 1, 1).Value <> "" Then
           
                   If clmRangeMail(rCell.Row - 1, 1).Value = "" Then
                      Range(Affix & rCell).EntireColumn.Hidden = True
               ElseIf clmRangeMail(rCell.Row - 1, 1).Value = "x" Then
                      Range(Affix & rCell).EntireColumn.Hidden = False
               End If
            
           End If
    End If

     rCell = (rCell.Row + 1)
Next rCell
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''rngHTMLshort
''Worksheets("RAPPORT_" & UserGroup).activate
          wbRS.Activate
'    Worksheets(RAPPORT_OUT).Select
'    Range("A1:DZ65000").Select
'    Selection.AutoFilter Field:=1, Criteria1:="<>"
    ActiveSheet.Rows("2:5").EntireRow.Hidden = True
    ActiveSheet.SpecialCells(xlCellTypeVisible).Select
    Set rngHTML = Selection.SpecialCells(xlCellTypeVisible)
    
    If rngHTML Is Nothing Then
       MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
    Exit Sub
    End If
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
          Call Mailing_ACC_link
                    
                    
                            Application.DisplayAlerts = False    'replacing
          wbRS.Activate
               ActiveSheet.Rows("2:5").EntireRow.Hidden = False
          Worksheets("RAPPORT_" & UserGroup).Delete
     
                            Application.DisplayAlerts = True    'replacing
                            Application.EnableEvents = True
                            Application.ScreenUpdating = True
         
         ''''''''''''''''''''''
         'Status toekennen door Mailing datum vast te leggen naar "Mailing.UserGroup"    TODO   Add Column
    
    If ActiveSheet.Name = "Accordering" Then
       For Each mCell In Range(Affix & "Screening." & UserGroup).SpecialCells(xlCellTypeVisible)
              Range(Affix & "Mailing." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Select
           If Range(Affix & "Mailing." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Value = "" And _
              Range(Affix & "Screening." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Value = "" Then
              Range(Affix & "Mailing." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Select
              Range(Affix & "Mailing." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Value = Format(Now, "dd-mm-yyyy h:mm")
              Range(Affix & "Screening." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Interior.Color = vbYellow
           End If
       Next mCell
ElseIf ActiveSheet.Name = "OUT" Then
       For Each mCell In Range(Affix & "Mailing." & UserGroup).SpecialCells(xlCellTypeVisible)
              Range(Affix & "Mailing." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Select
          If Range(Affix & "Mailing." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Value = "" And _
              Range(Affix & "Opgevoerd.in.SAP").Cells(mCell.Row - HeadingRows, 1).Value = "" Then  ''<>
             Range(Affix & "Mailing." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Select
             Range(Affix & "Mailing." & UserGroup).Cells(mCell.Row - HeadingRows, 1).Value = Format(Now, "dd-mm-yyyy h:mm")
          End If
          Next mCell
End If
         ''''''''''''''''''''''
       ElseIf Antwoord = vbNo Then
       End If
    End If
'''''''''''
'''''''''''
'''''''''''
Application.StatusBar = ". . . . . . . User Rapport: " & " : " & Aanvrager_var & "  -  " & UserGroup & "  -  " & Afdeling_var & "  -  " & Vestiging_var & _
                        "   ( " & (rngUserGroup.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) & " regels )"
''===============================||====================================
''===============================||====================================
''===============================||====================================
''''''''''''''''''''''''''''''''''
'    EnkelRapport = True

If EnkelRapport = True Then GoTo LaatstRapport

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

Andere_Vestiging:
Next RoleCell
''===============================||====================================
ENDstory:
                            Call SpeedOff
                            Application.DisplayAlerts = True    'replacing
                            Application.EnableEvents = True
                            Application.ScreenUpdating = True
Next I

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
''===============================||====================================
End Sub
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
'Sub Mailing_All()
''SpeedUp the macro
'Application.EnableEvents = True
'Application.ScreenUpdating = True
'''===============================||====================================
''Copy all rows from Worksheet "DATA" to Worksheet RAPPORT_UserGroup
'    Dim Rapport_OUT
'    Dim Sheet As Worksheet
'    Dim bestaat As Boolean
'    Dim rcell As Range
'
'
'For Each rcell In Range("SET.UserGroup").Cells
'      If rcell.Value <> "" Then
'          If clmRange(rcell.Row - 1, 1).Value = "X" Or clmRange(rcell.Row - 1, 1).Value <> "" Then
'PAGEBREAK:   'Page length defined in sheet (Print) SETTINGS
'              If pbrRange(rcell.Row - 1, 1).Value = "Y" Then
'                 Set rngMyRange_column = Range("i." & Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1))
'                'Set rngMyRange_data = Range("i." & rngMyRange_column)
'              Else
'              End If
'          Else
'          End If
'      Else
'      End If
'      rcell = (rcell.Row + 1)
'Next rcell
'    UserGroup = Role       ''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Rapport_OUT = "RAPPORT" & "_" & UserGroup
'
'    For Each Sheet In ThisWorkbook.Sheets
'        If Sheet.Name = Rapport_OUT Then bestaat = True: Exit For
'    Next Sheet
'
'    If bestaat = True Then
'        'Clear the contents of "RAPPORT_OUT" sheet
'        Worksheets(Rapport_OUT).Select
'        If ActiveSheet.AutoFilterMode = True Then
'        ActiveSheet.AutoFilterMode = False
'        End If
'       'Make all Columns / Rows Visible
'        Range("A:DZ").EntireColumn.Hidden = False
'        Range("1:65000").EntireRow.Hidden = False
'        Worksheets(Rapport_OUT).Cells.Clear
'        MsgBox = "Het tabblad " & Rapport_OUT & " bestaat al."
'
'    ElseIf bestaat = False Then
'        Worksheets.Add after:=Sheets(Sheets.Count)
'        ActiveSheet.Name = Rapport_OUT
''   if error aantal tekens tellen max 31 tekens
'    End If
'
''Copy the contents of "DATA" sheet
'    Worksheets("Accordering").Visible = xlSheetVisible
'    Worksheets("Accordering").Select
'    Range("A1", ActiveCell.SpecialCells(xlLastCell)).Select
'    Selection.COPY
'
'    Worksheets(Rapport_OUT).Select
'    Range("A1").Select
'
'''    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone,
'''        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'
'    ActiveWorkbook.Worksheets("Accordering").UsedRange.COPY Destination:=ThisWorkbook.Worksheets(Rapport_OUT).Range("A1")
'   'Worksheets("Accordering").Visible = xlSheetHidden
'''===============================||====================================
''DRAGON filtercolom: This has impact on a new sheet and NOT source sheet
'    Worksheets(Rapport_OUT).Select
'    Range("A6:DZ65000").Select
'    Selection.AutoFilter Field:=1, Criteria1:="<>"
'''===============================||====================================
'If UserGroup = "Manager" Then
'          'If UserLevel = 1 Then
'              GoTo Alle_Taken 'Excl_Open_Taak
'       ElseIf UserGroup = "GM_Vakgroep" Or UserGroup = "YPT0_Hoofdopdrachten" Then
'              GoTo Alleen_Laatste_Taak
'      'ElseIf UserGroup = "YPT0_Hoofdopdrachten" Then
'      '       GoTo Alleen_YPT0_Hoofdopdrachten
'       Else:  GoTo Alle_Taken
'End If
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Excl_Open_Taak:
'''''FILTER
'    Selection.AutoFilter Field:=Range("i.WERKPLEK").Column, Criteria1:="<>"
'''  Selection.AutoFilter Field:=Range("i.TAAKSTATUS").Column, Criteria1:="<>OPEN"
'''''SORT
'    Range("A2:CI65000").Sort _
'    Key1:=Range("i.WERKPLEK"), Order1:=xlAscending, DataOption1:=xlSortNormal, _
'    Key2:=Range("i.WEEK"), Order2:=xlAscending, DataOption2:=xlSortTextAsNumbers, _
'    Key3:=Range("i.DAGEN"), Order3:=xlAscending, DataOption3:=xlSortNormal, _
'    Header:=xlNo, OrderCustom:=1, MatchCase:=True, Orientation:=xlTopToBottom
'GoTo FIRST_ONLY_END
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Alle_Taken:
'''===============================||====================================
'FIRST_ONLY_END:
''Reset all Page Breaks
'ActiveSheet.ResetAllPageBreaks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''delete alle code
'SORTING_START_1: ''WERKT NOG NIET svp. TESTEN
'For Each rcell In Range("SET.RANGE_ALL").Cells
'      If rcell.Value <> "" Then
'          If clmRange(rcell.Row - 1, 1).Value = "X" Or clmRange(rcell.Row - 1, 1).Value <> "" Then
'SORTING_1:     'Each Column Sorting Value defined by 3 Keys in sheet SETTINGS
'              If srtRange(rcell.Row, 1).Value = "Y" Then
'              ''srtCell
'              Else
'              End If
'          Else
'          End If
'      Else
'      End If
'Next rcell
'SORTING_END_1:
'''===============================||====================================
'''===============================||====================================
'''===============================||====================================
'PAGEBREAK_START:
'Dim rngMyRange_column  As Variant
'
'For Each rcell In Range("SET.RANGE_ALL").Cells
'      If rcell.Value <> "" Then
'          If clmRange(rcell.Row - 1, 1).Value = "X" Or clmRange(rcell.Row - 1, 1).Value <> "" Then
'PAGEBREAK:   'Page length defined in sheet (Print) SETTINGS
'              If pbrRange(rcell.Row - 1, 1).Value = "Y" Then
'                 Set rngMyRange_column = Range("i." & Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1))
'                'Set rngMyRange_data = Range("i." & rngMyRange_column)
'              Else
'              End If
'          Else
'          End If
'      Else
'      End If
'      rcell = (rcell.Row + 1)
'Next rcell
'PAGEBREAK_END:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Page Break Insert Horizontal
'With Worksheets(Rapport_OUT) 'rngMyRange_data
'    For Each rngCell_data In rngMyRange_column
'       'loop through the range
'       If (rngCell_data.Value <> rngCell_data.Offset(1, 0).Value) And rngCell_data.Offset(1, 0).Value <> "" Then
'          .HPageBreaks.Add Before:=rngCell_data.Offset(1, 0)
'       End If
'    Next
'
' 'Define last column
'  Dim LastColumn, LastColumn_new As Integer
'  Dim LastColumn_new_Letter As String
'  If WorksheetFunction.CountA(Cells) > 0 Then
'    'Search for any entry, by searching backwards by Columns
'     LastColumn = Cells.Find(What:="*", after:=[Z1], _
'                            searchorder:=xlByColumns, _
'                            searchdirection:=xlPrevious).Column
'     LastColumn_new = LastColumn + 1
'     LastColumn_new_Letter = Chr(LastColumn_new + 64)
'  End If
'
' 'Page Break Insert Vertical
'  ActiveSheet.VPageBreaks.Add Before:=ActiveSheet.Range(LastColumn_new_Letter & "1")
'End With
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GoTo LASTLINE_SETTINGS
'''===============================||====================================
'LASTLINE_SETTINGS:
''Verwijderen text "i."  Range benaming
'        Worksheets(Rapport_OUT).Visible = xlSheetVisible
'        Worksheets(Rapport_OUT).Select
'
'        Worksheets(Rapport_OUT).Rows("1:1").Select
'        Dim di
'
'        For Each di In Selection
'            If di.Value <> "" Then
'            Range("1:1").Replace What:="i.", Replacement:="", _
'            lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False
'            End If
'        Next di
'''===============================||====================================
'''===============================||====================================
'''===============================||====================================
'PAGESETUP_START:
'    For Each rcell In Range("SET.RANGE_ALL").Cells
'    If rcell.Value <> "" Then
'        If clmRange(rcell.Row - 1, 1).Value = "X" Or clmRange(rcell.Row - 1, 1).Value <> "" Then
'        If pstRange(rcell.Row - 1, 1).Value = "Y" Then
'PAGESETUP:
'Dim UserGroup_Orientation As String
'Dim UserGroup_PWide As Integer
'Dim UserGroup_PTall As Integer
'Dim UserGroup_PageSize As String
'
'            If Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1) = "PageSize" Then
'            UserGroup_PageSize = clmRange(rcell.Row - 1, 1).Value
'            ''Range("SET.RANGE_ALL").Range("PageSize").Row = ""
'            ElseIf Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1) = "Orientation" Then
'            UserGroup_Orientation = clmRange(rcell.Row - 1, 1).Value
'            ElseIf Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1) = "PWide" Then
'            UserGroup_PWide = clmRange(rcell.Row - 1, 1).Value
''            ElseIf Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1) = "PTall" Then
''            Set UserGroup_PTall = clmRange(rcell.Row - 1, 1).Value
'            Else
'            End If
''***********
'            Else
'            End If
'          Else
'          End If
'      Else
'      End If
'Next rcell
'PAGESETUP_END:
'''===============================||====================================
''Define Attributes Standaard Name & Value (into ARRAY/RANGE defined in sheet SETTINGS)
''Use Attributes Columns & Values into ARRAY/RANGE defined in sheet SETTINGS
''    Dim Std_Column_01 As Range
'Dim Std_Column_01_Name As Variant
'Dim Std_Column_02_Name As Variant
'Dim Std_Column_03_Name As Variant
'Dim Std_Column_04_Name As Variant
'Dim Std_Column_05_Name As Variant
'Dim Std_Column_06_Name As Variant
'Dim Std_Column_07_Name As Variant
'Dim Std_Column_01_Value As Variant
'Dim Std_Column_02_Value As Variant
'Dim Std_Column_03_Value As Variant
'Dim Std_Column_04_Value As Variant
'Dim Std_Column_05_Value As Variant
'Dim Std_Column_06_Value As Variant
'Dim Std_Column_07_Value As Variant
'
'Dim OH_Column_01_Name As Variant
'Dim OH_Column_02_Name As Variant
'Dim OH_Column_03_Name As Variant
'Dim OH_Column_04_Name As Variant
'Dim OH_Column_05_Name As Variant
'Dim OH_Column_01_Value As Variant
'Dim OH_Column_02_Value As Variant
'Dim OH_Column_03_Value As Variant
'Dim OH_Column_04_Value As Variant
'Dim OH_Column_05_Value As Variant
'
'Dim OH_Spare_01_Name As Variant
'Dim OH_Spare_02_Name As Variant
'Dim OH_Spare_03_Name As Variant
'Dim OH_Spare_04_Name As Variant
'Dim OH_Spare_05_Name As Variant
'Dim OH_Spare_01_Value As Variant
'Dim OH_Spare_02_Value As Variant
'Dim OH_Spare_03_Value As Variant
'Dim OH_Spare_04_Value As Variant
'Dim OH_Spare_05_Value As Variant
'
''ATTRIBUTES_START:
'    For Each rcell In Range("SET.RANGE_ALL").Cells
'        If rcell.Value <> "" Then
'            If atrRange(rcell.Row - 1, 1).Value = "Y" Then
''ATTRIBUTES:    'Each Column Filter Value defined in sheet SETTINGS
'                If Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "Name" Then
'                       Std_Column_01_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       Std_Column_01_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "Title" Then
'                       Std_Column_02_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       Std_Column_02_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "Aspect" Then
'                       Std_Column_03_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       Std_Column_03_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "Soort" Then
'                       Std_Column_04_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       Std_Column_04_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "Afdeling" Then
'                       Std_Column_05_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       Std_Column_05_Value = clmRange(rcell.Row - 1, 1).Value
'
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.planningsgroep" Then
'                       OH_Column_01_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Column_01_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.planner" Then
'                       OH_Column_02_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Column_02_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.klant" Then
'                       OH_Column_03_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Column_03_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.gebruiker" Then
'                       OH_Column_04_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Column_04_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "SAP.doctype" Then
'                       OH_Column_05_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Column_05_Value = clmRange(rcell.Row - 1, 1).Value
'
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.Spare1" Then
'                       OH_Spare_01_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Spare_01_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.Spare2" Then
'                       OH_Spare_02_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Spare_02_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.Spare3" Then
'                       OH_Spare_03_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Spare_03_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.Spare4" Then
'                       OH_Spare_04_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Spare_04_Value = clmRange(rcell.Row - 1, 1).Value
'                ElseIf Range("SET.RANGE_ALL")(rcell.Row - 1, 1).Value = "OH.Spare5" Then
'                       OH_Spare_05_Name = Range("SET.RANGE_ALL").Range("A" & rcell.Row - 1)
'                       OH_Spare_05_Value = clmRange(rcell.Row - 1, 1).Value
'
'                Else
'                End If
'            Else
'            End If
'        Else
'        End If
'    rcell = (rcell.Row + 1)
'Next rcell
''ATTRIBUTES_END:
'''===============================||====================================
'Worksheets(Rapport_OUT).Select
'    With ActiveSheet.PAGESETUP
'        .PrintTitleRows = "$1:$1"
'      ''.PrintTitleColumns = ""
'      ''.PrintArea = "$A2:$BD65000"
'        .LeftHeader = "&L&8&F.xls : [&A]"           '"Arial""http://dafportal.eu.paccar.com/operations/departments/maintenance/Pages/us.aspx"
'        .CenterHeader = ""
'        .RightHeader = UserGroup & " (" & USERLEVEL & ")"
'        .LeftFooter = ""                                                    '\"Map: " & ActiveWorkbook.Path
'        .CenterFooter = "&C&8&""Arial""Pagina:  &P / &N"
'        .RightFooter = "&R&8&""Arial""Afdruk: " & Format(Date, "dd-mm-yyyy") '\Format(Date, "dd-mm-yyyy") (Now, "dd-mm-yyyy   hh:mm")
'        .LeftMargin = Application.InchesToPoints(0.3)
'        .RightMargin = Application.InchesToPoints(0.3)
'        .TopMargin = Application.InchesToPoints(1#)
'        .BottomMargin = Application.InchesToPoints(0.5)
'        .HeaderMargin = Application.InchesToPoints(0.6)
'        .FooterMargin = Application.InchesToPoints(0.25) '\ 0.75(1.9) 0.70(1.8) 0.65(1.7) 0.6(1.5) 0.55(1.4) 0.5(1.3) 0.45(1.1)
'        .PrintQuality = 600
'        .PrintHeadings = False                           '\true
'        .PrintGridlines = True                           '\True
'        .PrintComments = xlPrintNoComments
'      ''.PaperSize = PaperSize                           '\User Defined Paper Size
'        .CenterHorizontally = True
'        .CenterVertically = False                       '\FALSE xlCenter
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = False
'        .Draft = False
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False                          '\True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            If UserGroup_Orientation = "L" Then         'User Defined Page Orientation
'            .Orientation = xlLandscape
'            ElseIf UserGroup_Orientation = "P" Then
'            .Orientation = xlPortrait
'            ElseIf UserGroup_Orientation = "" Then
'            .Orientation = xlPortrait
'            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            If UserGroup_PageSize = "A4" Then
'                   .PaperSize = 9
'            ElseIf UserGroup_PageSize = "A3" Then
'                   .PaperSize = 8                                    'íf error then A4 Papersize = 9 (A3 local printer does not exist: Papersize = 8)
'            ElseIf UserGroup_PageSize = "A5" Then
'                   .PaperSize = 11
'            ElseIf UserGroup_PageSize = "Legal" Then
'                   .PaperSize = 5
'            ElseIf UserGroup_PageSize = "Letter" Then
'                   .PaperSize = 1
'            ElseIf UserGroup_PageSize = "Quarto" Then
'                   .PaperSize = 15
'            ElseIf UserGroup_PageSize = "Executive" Then
'                   .PaperSize = 7
'            ElseIf UserGroup_PageSize = "B4" Then
'                   .PaperSize = 12
'            ElseIf UserGroup_PageSize = "B5" Then
'                   .PaperSize = 13
'            ElseIf UserGroup_PageSize = "10x14" Then
'                   .PaperSize = 16
'            ElseIf UserGroup_PageSize = "11x17" Then
'                   .PaperSize = 17
'            ElseIf UserGroup_PageSize = "Csheet" Then
'                   .PaperSize = 24
'            ElseIf UserGroup_PageSize = "Dsheet" Then
'                   .PaperSize = 25
'            Else
'                   .PaperSize = 9 'Defaults to A4
'            End If
''***********************************************************************************************
''''''''''            .CenterHeader = Header    'User Defined Header (Shift to Left or Right as required)
''''''''''            .LeftFooter = Footer      'User Defined Footer (Shift to Left or Right as required)
''''''''''            .FitToPagesWide = PWide     'User Defined No Pages Wide
''''''''''            .FitToPagesTall = PTall     'User Defined No Pages Tall
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       .Zoom = False '100  '70
'       .FitToPagesWide = 1
'       .FitToPagesTall = False
'       .PrintErrors = False
'      '.PrintErrors = xlPrintErrorsDisplayed
'     End With
'''===============================||====================================
'''PAGINA-EINDVOORBEELD
'    Worksheets(Rapport_OUT).Select
'    If ActiveWindow.View = xlNormalView Then
'       ActiveWindow.View = xlPageBreakPreview
'    Else
'       ActiveWindow.View = xlPageBreakPreview
'    End If
'
'    'Activate if Rapport_Out.xls without filter preffered
'    ''Selection.AutoFilter
'    ''1ste rij bevriezen
'        With ActiveWindow
'        .SplitColumn = 0
'        .SplitRow = 1
'        End With
'    ActiveWindow.FreezePanes = True
'''===============================||====================================
'   'Makes a copy of the active sheet and save it to a temporary file
'    Dim FileName, wb, Path
'
'    Worksheets("RAPPORT_" & UserGroup).COPY
'    Set wb = Worksheets("RAPPORT_" & UserGroup)
'    FileName = UserGroup & ".xlsx"
'    Path = "C:\Rapport_Out\"
'  ''MsgBox "Look in <" & Path & "> ... for the files!" & vbNewLine & "Look if map <" & Path & "> is present!"
'''''''''''''''''''''''''''''''''''
'    Dim nm As Name
'
'    On Error Resume Next
'    For Each nm In ActiveWorkbook.Names
'        nm.Delete
'    Next
'    On Error GoTo 0
'''''''''''''''''''''''''''''''''''
''Herhaling 1ste rij in de nieuwe document
'    With ActiveSheet.PAGESETUP
'        .PrintTitleRows = "$1:$1"
'     End With
''==============================================================
''Filter inzetten in erste rij           NEW TEST DRAGAN
''Activate if Rapport_Out.xls with filter preffered
'
''                If Not ActiveSheet.AutoFilterMode Then
''                  ActiveSheet.Range("A1").AutoFilter
''                End If
''==============================================================
''Teksten toevoegen aan de bestaande BuiltIn Document Properties m.b.v. VBA
'ActiveWorkbook.BuiltinDocumentProperties("Title").Value = FileName
'ActiveWorkbook.BuiltinDocumentProperties("Subject").Value = "SAP Rapportage SharePoint"
'ActiveWorkbook.BuiltinDocumentProperties("Company").Value = "DAF Trucks Eindhoven"
'ActiveWorkbook.BuiltinDocumentProperties("Author").Value = "Dragan Straleger"
'ActiveWorkbook.BuiltinDocumentProperties("Manager").Value = "Maintenance Manager Engineering & Planning"
'ActiveWorkbook.BuiltinDocumentProperties("Comments").Value = "No comments at this moment"
'
''Je eigen Custom Document Properties toevoegen m.b.v. VBA
'ActiveWorkbook.CustomDocumentProperties.Add Name:=Std_Column_03_Name, LinkToContent:=False, Value:=Std_Column_03_Value, Type:=msoPropertyTypeString
'ActiveWorkbook.CustomDocumentProperties.Add Name:=Std_Column_04_Name, LinkToContent:=False, Value:=Std_Column_04_Value, Type:=msoPropertyTypeString
'ActiveWorkbook.CustomDocumentProperties.Add Name:=Std_Column_05_Name, LinkToContent:=False, Value:=Std_Column_05_Value, Type:=msoPropertyTypeString
'''GoTo Spare_Items:
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Column_01_Name, LinkToContent:=False, Value:=OH_Column_01_Value, Type:=msoPropertyTypeString
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Column_02_Name, LinkToContent:=False, Value:=OH_Column_02_Value, Type:=msoPropertyTypeString
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Column_03_Name, LinkToContent:=False, Value:=OH_Column_03_Value, Type:=msoPropertyTypeString
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Column_04_Name, LinkToContent:=False, Value:=OH_Column_04_Value, Type:=msoPropertyTypeString
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Column_05_Name, LinkToContent:=False, Value:=OH_Column_05_Value, Type:=msoPropertyTypeString
'Spare_Items:
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Spare_01_Name, LinkToContent:=False, Value:=OH_Spare_01_Value, Type:=msoPropertyTypeString
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Spare_02_Name, LinkToContent:=False, Value:=OH_Spare_02_Value, Type:=msoPropertyTypeString
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Spare_03_Name, LinkToContent:=False, Value:=OH_Spare_03_Value, Type:=msoPropertyTypeString
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Spare_04_Name, LinkToContent:=False, Value:=OH_Spare_04_Value, Type:=msoPropertyTypeString
'    'ActiveWorkbook.CustomDocumentProperties.Add Name:=OH_Spare_05_Name, LinkToContent:=False, Value:=OH_Spare_05_Value, Type:=msoPropertyTypeString
'
''Je eigen Custom Document Properties toevoegen m.b.v. VBA
''                    ActiveWorkbook.CustomDocumentProperties.Add Name:="Aspect", LinkToContent:=False, Value:="Nr 1a", Type:=msoPropertyTypeString
''                    ActiveWorkbook.CustomDocumentProperties.Add Name:="Soort", LinkToContent:=False, Value:="Nr 2a", Type:=msoPropertyTypeString
''                    ActiveWorkbook.CustomDocumentProperties.Add Name:="Afdeling", LinkToContent:=False, Value:="Nr 3a", Type:=msoPropertyTypeString
''                    ActiveWorkbook.CustomDocumentProperties.Add Name:="Version", LinkToContent:=False, Value:="Nr 4a", Type:=msoPropertyTypeString
''                    ActiveWorkbook.CustomDocumentProperties.Add Name:="OH.planningsgroep", LinkToContent:=False, Value:="Nr 5a", Type:=msoPropertyTypeString
''                    ActiveWorkbook.CustomDocumentProperties.Add Name:="OH.klant", LinkToContent:=False, Value:="Nr 6a", Type:=msoPropertyTypeString
''                    ActiveWorkbook.CustomDocumentProperties.Add Name:="OH.gebruiker", LinkToContent:=False, Value:="Nr 7a", Type:=msoPropertyTypeString
''                    ActiveWorkbook.CustomDocumentProperties.Add Name:="SAP.doctype", LinkToContent:=False, Value:="Nr 8a", Type:=msoPropertyTypeString
''This properties are to find under MENU:Ontwikkelaars|| Documentpaneel ||Geavanceerde Eigenschappen van Document(Keuzemenu)
''==============================================================
''Replace existing files
'Application.DisplayAlerts = False   'replacing
'    wb.SaveAs Path & FileName, FileFormat:=51
'    ActiveWorkbook.Close
'Application.DisplayAlerts = True    'replacing
'''MsgBox "Look in <" & Path & "> ... for the files!" & vbNewLine & "Look if map <" & Path & "> is present!"
'   'ActiveWorkbook.SaveAs "C:\ron.xlsm", FileFormat:=52
'   '50 = xlExcel12 (Excel Binary Workbook in 2007-2013 with or without macro's, xlsb)
'   '51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
'   '52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
'   '56 = xlExcel8 (97-2003 format in Excel 2007-2013, xls)
'   'FileExtStr = ".csv": FileFormatNum = 6
'   'FileExtStr = ".txt": FileFormatNum = -4158
'   'FileExtStr = ".prn": FileFormatNum = 36
''==============================================================
''SpeedUp the macro
'Application.EnableEvents = True
'Application.ScreenUpdating = True
''==============================================================
'    Worksheets("Accordering").Select
'
'    If ActiveSheet.AutoFilterMode = True Then
'       ActiveSheet.AutoFilterMode = False
'    End If
'
'    Range("A:DZ").EntireColumn.Hidden = False
'    Range("1:65000").EntireRow.Hidden = False
'''===============================||====================================
'End Sub
'
