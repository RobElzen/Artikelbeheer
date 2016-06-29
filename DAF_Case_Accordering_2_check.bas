Attribute VB_Name = "DAF_Case_Accordering_2_check"

Sub Init_Columns_ACCoord_check()
''===============================||====================================
'Make all Collumns / Rows Visible
    
    Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Visible = xlSheetVisible
    Workbooks("Artikelbeheer.xlsm").Sheets("Accordering").Select
    
    Call ProtectOff
    
'    If ActiveSheet.AutoFilterMode = True Then
'       ActiveSheet.AutoFilterMode = False
'    End If
'
'    Range("A:DZ").EntireColumn.Hidden = False
'    Range("1:65000").EntireRow.Hidden = False
Dim Sheet_Cells_All As Range
'                          Range("A1", ActiveCell.SpecialCells(xlLastCell)).Select
    Set Sheet_Cells_All = Range("A1", ActiveCell.SpecialCells(xlLastCell))
''===============================||====================================
    Dim rCell As Variant
    Dim clmInvisible As Range   ''C_olumn_zichtbaar
    Dim clmRange As Range       ''C_olumn_User
    Dim clhRange As Range       ''C_olumn LAY_out
    
'    Set clmRange = Range("SET." & Role)
'    Set clhRange = Range("SET.ColumnHide")
''===========================COLUMNHIDE================================
On Error Resume Next

If Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) <> Aanvraag_level_69 Then
  
   Call SpeedOn
'====================== ULTIEME CHECK ACCOORD ========================
'Aanvraag afgewezen door Databeheerder
   If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
      Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Then
'     MsgBox "afgewezen"
      Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "NEE"
      Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbRed
      Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_67
      GoTo Accoord_gereed_DB
   ElseIf Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
      Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "JA" Then
      Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone    ''checken
      Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_61
   End If
'=================== 25000 € =< X ====================================
   If Range("ACC_Aanvraagbedrag").Range("A" & ActiveCell.Row - HeadingRows) >= 25000 Then
   
     If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
      ((Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "NL") Or _
       (Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "BE")) Then

         If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
          ((Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "NL") Or _
           (Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "BE")) Then
'           MsgBox "gelukt met goedkeur"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "JA"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbGreen
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_64
       
     ElseIf Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
          ((Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows) = "NEE") And _
            Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "NL") Or _
          ((Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows) = "NEE") And _
            Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "BE") Then
'           MsgBox "afgewezen"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "NEE"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbRed
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_67
'    Else   MsgBox "onvolledig"
          End If
          
      Else      ''Moet deze iets hoger geplaatst worden     MsgBox "onvolledig"
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).ClearContents
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone
      End If
     
     'DIT STUKJE CODE AAN HET EINDE PLAATSEN EN NIET MEERDERE KEREN HERHALEN
     'Mailing geel verwijderen als geaccordeerd EN GEEN GRIJS KLEUR IS
      If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      
      If Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "NL" Then
         If Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
         End If
            Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows).Select
            Call Columns_ACCoord_avoid
         If Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
         End If
            Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows).Select
            Call Columns_ACCoord_avoid
      
      ElseIf Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "BE" Then
         If Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
         End If
         If Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
         End If
         If Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
         End If
         If Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
         End If
      End If
'''
GoTo Accoord_gereed_DB
   End If
'=================== 12500 € =< X < 25000 € ==========================
   If Range("ACC_Aanvraagbedrag").Range("A" & ActiveCell.Row - HeadingRows) >= 12500 And _
      Range("ACC_Aanvraagbedrag").Range("A" & ActiveCell.Row - HeadingRows) < 25000 Then
'NL

     If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
      ((Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "NL") Or _
       (Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "BE")) Then

         If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
          ((Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "NL") Or _
           (Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "BE")) Then
'           MsgBox "gelukt met goedkeur"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "JA"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbGreen
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_64
       
     ElseIf Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
         (((Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" And _
            Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "NL") Or _
          ((Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows) = "NEE") And _
            Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "BE"))) Then
'           MsgBox "afgewezen"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "NEE"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbRed
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_67
'    Else   MsgBox "onvolledig"
          End If
          
      Else      ''Moet deze iets hoger geplaatst worden     MsgBox "onvolledig"
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).ClearContents
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone
      End If
     
     'DIT STUKJE CODE AAN HET EINDE PLAATSEN EN NIET MEERDERE KEREN HERHALEN
     'Mailing geel verwijderen als geaccordeerd EN GEEN GRIJS KLEUR IS
      If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      
      If Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "NL" Then
         If Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
            
            Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows).Select
            Call Columns_ACCoord_avoid
         End If
      ElseIf Range("ACC_Vestiging").Range("A" & ActiveCell.Row - HeadingRows) = "BE" Then
         If Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
         End If
         If Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
            Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
         End If
      End If
''''''Grijze veld bij geen active cell
      Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
'''
GoTo Accoord_gereed_DB
   End If
'==================== 2500 € =< X < 12500 € ==========================
   If Range("ACC_Aanvraagbedrag").Range("A" & ActiveCell.Row - HeadingRows) >= 2500 And _
      Range("ACC_Aanvraagbedrag").Range("A" & ActiveCell.Row - HeadingRows) < 12500 Then
      
      If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
         Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
         Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
         Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
         Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
         Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         
         If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) = "JA" Then
'           MsgBox "gelukt met goedkeur"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "JA"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbGreen
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_64
       
     ElseIf Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Then
'            MsgBox "afgewezen"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "NEE"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbRed
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_67
'    Else   MsgBox "onvolledig"
          End If
          
      Else      ''Moet deze iets hoger geplaatst worden     MsgBox "onvolledig"
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).ClearContents
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone
      End If
      
     'Mailing geel verwijderen als geaccordeerd
      If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
''''''Grijze veld bij geen active cell
      Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
'''
GoTo Accoord_gereed_DB
   End If
'==================== 1250 € =< X < 2500 € ===========================
   If Range("ACC_Aanvraagbedrag").Range("A" & ActiveCell.Row - HeadingRows) >= 1250 And _
      Range("ACC_Aanvraagbedrag").Range("A" & ActiveCell.Row - HeadingRows) < 2500 Then
      
     If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then

         If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) = "JA" Then
'           MsgBox "gelukt met goedkeur"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "JA"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbGreen
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_64
       
     ElseIf Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Then
'            MsgBox "afgewezen"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "NEE"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbRed
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_67
'    Else   MsgBox "onvolledig"
          End If
          
      Else      ''Moet deze iets hoger geplaatst worden     MsgBox "onvolledig"
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).ClearContents
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone
      End If

     'Mailing geel verwijderen als geaccordeerd
      If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
''''''Grijze veld bij geen active cell
      Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
'''
GoTo Accoord_gereed_DB
   End If
'=========================== X < 1250 ================================
   If Range("ACC_Aanvraagbedrag").Range("A" & ActiveCell.Row - HeadingRows) < 1250 Then
   
     If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" And _
        Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then

         If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "JA" And _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "JA" Then
'           MsgBox "gelukt met goedkeur"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "JA"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbGreen
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_64
       
     ElseIf Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Or _
            Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) = "NEE" Then
'            MsgBox "afgewezen"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1) = "NEE"
            Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = vbRed
            Range("ACC_Aanvraag.code").Cells(ActiveCell.Row - HeadingRows, 1) = Aanvraag_level_67
'    Else   MsgBox "onvolledig"
          End If
          
      Else      ''Moet deze iets hoger geplaatst worden     MsgBox "onvolledig"
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).ClearContents
          Range("ACC_Gereed_voor_Upload.SAP").Cells(ActiveCell.Row - HeadingRows, 1).Interior.Color = xlNone
      End If

     'Mailing geel verwijderen als geaccordeerd
      If Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.DB").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.ICM").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
      If Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows) <> "" Then
         Range("ACC_Screening.MMP").Range("A" & ActiveCell.Row - HeadingRows).Interior.Color = xlNone
      End If
''''''Grijze veld bij geen active cell
      Range("ACC_Screening.MMR").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.CMO").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.MMO").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.COE").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.DOE").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.COW").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
      Range("ACC_Screening.DOW").Range("A" & ActiveCell.Row - HeadingRows).Select
      Call Columns_ACCoord_avoid
'''
GoTo Accoord_gereed_DB
   End If
'===============================||====================================
Else
Dim Aanvraagcode_new As String
Dim Aanvraagcode_old As String
    
    Application.EnableEvents = False
    Aanvraagcode_new = ActiveCell.Value
    Application.Undo
    Aanvraagcode_old = ActiveCell.Value
    Application.EnableEvents = True
End If

Accoord_gereed_DB:

Call Generate_Ranges_ALL
Call ProtectOnALL
Call SpeedOff
End Sub

''Define Grey color cells
Sub Columns_ACCoord_avoid()
    ActiveCell.ClearContents
    With Selection.Interior
        .Pattern = xlSolid                      ''xlGray50      xlGray8     xlSolid     xlUp
        .PatternColorIndex = xlAutomatic
        .ColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1         ''
        .TintAndShade = -0.1                    ''0 (licht)     90 (donker)
        .PatternTintAndShade = 0                ''
    End With
End Sub


