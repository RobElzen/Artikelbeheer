Attribute VB_Name = "M16_Init_ACC_col_mail"

Sub Init_Columns_ACC_mail()
''===============================||====================================
'Make all Collumns / Rows Visible
    
'    Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Visible = xlSheetVisible
'    Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Activate
'    Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Select
    
    If Workbooks("Artikelbeheer.xlsm").ActiveSheet.Name <> "Accordering" And _
       Workbooks("Artikelbeheer.xlsm").ActiveSheet.Name <> "OUT" Then
       If Role = "ME" Then
          Workbooks("Artikelbeheer.xlsm").Sheets("OUT").Select
       End If
       If Role <> "ME" Then
          Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Select
       End If
    End If

    Call Affix_Case
    
    Call ProtectOff
    Call SpeedOn
    
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    Range("A:DZ").EntireColumn.Hidden = False
    Range("1:65000").EntireRow.Hidden = False
    
    Dim Sheet_Cells_All As Range
    Set Sheet_Cells_All = Range("A1", ActiveCell.SpecialCells(xlLastCell))
''===============================||====================================

Set wbL = Workbooks("Lijsten_New.xlsm")
Set wsL = wbL.Worksheets("SETTINGS")
'    wbL.Activate
'    wbL.Worksheets ("SETTINGS")
    
    Dim rCell As Variant
    Dim clmInvisible As Range   ''C_olumn_zichtbaar
    Dim clmRange As Range       ''C_olumn_User
    Dim clhRange As Range       ''C_olumn LAY_out
    
    Set clmRange = wsL.Range("SET." & Role)
    Set clhRange = wsL.Range("SET.ColumnHide")
''===========================COLUMNHIDE================================
'Call SpeedOn
Dim Column_Name As String
                     
For Each rCell In wsL.Range("SET.RANGE_ALL").Cells
      If rCell.Value <> "" Then
           If clhRange(rCell.Row - 1, 1).Value <> "" Then
           
                    Column_Name = wsL.Range("SET.RANGE_ALL").Range("A" & rCell.Row - 1)
                    Workbooks("Artikelbeheer.xlsm").Activate
                    Workbooks("Artikelbeheer.xlsm").ActiveSheet.Select
                    Call ProtectOff
                    'Set clmInvisible = Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Range("ACC_" & Column_Name)
                    'Set clmInvisible = Range("ACC_" & Column_Name)
                     Set clmInvisible = Range(Affix & Column_Name)
                     
                  If clmRange(rCell.Row - 1, 1).Value = "" Then                              ''HIDDEN LOCKED
'                     clmInvisible.Select
                     clmInvisible.Locked = True
                     clmInvisible.EntireColumn.Hidden = True
              ElseIf clmRange(rCell.Row - 1, 1).Value = "H" Then                             ''HIDDEN UNLOCKED
'                     clmInvisible.Select
                     clmInvisible.Locked = False
                     clmInvisible.EntireColumn.Hidden = True
              ElseIf clmRange(rCell.Row - 1, 1).Value = "R" Then                             ''LOCKED
                     clmInvisible.Select
                     clmInvisible.Locked = True
                     clmInvisible.EntireColumn.Hidden = False
              ElseIf clmRange(rCell.Row - 1, 1).Value = "W" Then                             ''UNLOCKED
'                     clmInvisible.Select
                     clmInvisible.Locked = False
                     clmInvisible.EntireColumn.Hidden = False
              End If
                 
           End If
    End If
     rCell = (rCell.Row + 1)
Next rCell

Call SpeedOff
Call ProtectOnALL
End Sub

