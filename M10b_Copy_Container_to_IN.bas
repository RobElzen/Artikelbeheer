Attribute VB_Name = "M10b_Copy_Container_to_IN"

Sub copy_Container_to_IN()
''===============================||====================================
Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
''===============================||====================================
''Define Filter Values into ARRAY/RANGE defined in sheet SETTINGS
    Dim Aanvraagbestand_naam As String
    Dim statusRange As Range
    Dim Column_Name_Werkbestand_CHECK As Integer
    Dim Column_Name_Werkbestand As String
    Dim Column_Name_Container As Range                  ''Doe ik iets hiermee?
    Dim rCell As Variant
    Dim I As Integer
    I = 0
    Dim j As Integer
    j = 0
    ''HeadingRows = 5
    mylastRow_Container = 0
''===============================||====================================
    Aanvraagbestand_naam = ActiveWorkbook.Name
''===============================||====================================
''===============================||====================================
''===============================||====================================
Dim xlApp As Excel.Application
    Dim strFile As String
    strFile = Path & "/" & "Container.xlsm"
Dim Stempel_copy As Boolean

If Workbooks.CanCheckOut(strFile) = True Then
   Application.Workbooks.Open strFile
   Workbooks.CheckOut strFile
   Stempel_copy = True
Else
    MsgBox "Uitchecken van het bestand is niet mogelijk. Probeer later nog een keer."
    Stempel_copy = False
    Exit Sub
End If
''===============================||====================================
''===============================||====================================
''===============================||====================================
'rCell is positie in RangeName excl. HeadingRows
COPY_START:
    Workbooks("Container").Worksheets("Container").Activate
SpeedOn
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
SpeedOn
    Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
SpeedOn
'SETTINGS CALL voor status check "ME_inleveren"
    Set statusRange = Range("CNT_Aanvraag.code")
    
    mylastRow_Container = Workbooks("Container").Worksheets("Container").UsedRange.Rows.Count - HeadingRows
    
''===============================||====================================
''===============================||====================================
''===============================||====================================
    For Each rCell In statusRange.Cells
    Workbooks("Container").Worksheets("Container").Activate
    If statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_24 And _
       mylastRow_Container > 0 Then
''===============================||====================================
SpeedOn
           statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_29
           Set Network = CreateObject("wscript.network")
           Cells(rCell.Row, Range("CNT_Aanvraagbeheerder").Column).Value = Naam       'inwisellen met variable (Affix & "_Aanvraagbeheerder")
           Cells(rCell.Row, Range("CNT_Datum_IN_AB").Column).Value = Now()
'SpeedOff
''''''''''''''''''''''''''''''''''
'      Call Generate_Ranges_ALL                           ''ERGENS ANDERS PLAATSEN
   'Define Container last row   '3 regels zijn leeg en worden niet opgeteld dus *** LET OP ".UsedRange.Rows.Count" telt alleen NonEempty regels
''===============================||====================================
   'vind en copy range to range
   'check hoeveel RECORDS zijn er om in te dienen anders open geen Container
''''''''''''''
    Workbooks(Aanvraagbestand_naam).Worksheets("IN").Activate
SpeedOn
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
    mylastRow_IN = Workbooks(Aanvraagbestand_naam).Worksheets("IN").UsedRange.Rows.Count
'''''''''''''''
''duplicate in IN_to_DATA
'    Set wb = Workbooks("Container.xlsm")
    HeaderNameColumn = Workbooks("Container").Worksheets("Container").Cells(1, 1).SpecialCells(xlLastCell).Column
        
        Workbooks("Container").Worksheets("Container").Activate

SpeedOn ''Hiermee wordt macro "Worksheet_Change" gedeactiveerd
           Range(Cells(rCell.Row, 1), Cells(rCell.Row, HeaderNameColumn)).COPY Destination:=Workbooks(Aanvraagbestand_naam).Worksheets("IN").Cells(mylastRow_IN + 1, 1)
''''''''
    Workbooks(Aanvraagbestand_naam).Worksheets("IN").Activate
    Cells(mylastRow_IN + 1, Range("IN_Aanvraag.code").Column).Value = Aanvraag_level_30
    
    Cells(mylastRow_IN + 1, Range("IN_Aanvraagbeheerder").Column) = Naam    ''dubbeling zie boven
    Cells(mylastRow_IN + 1, Range("IN_Datum_IN_AB").Column) = Now()                     ''dubbeling zie boven
''''''''''''''''''''''''''''''''''
Else
'MsgBox "No records to copy!"
End If
''========================================================================''
CHECK_OVER:
    rCell = (rCell.Row + 1)
'SpeedOff
Next rCell
''========================================================================''
''========================================================================''
''========================================================================''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
       Application.Run ("'Lijsten_new.xlsm'!ProtectOnALL")
COPY_END:
''===============================||====================================
''===============================||====================================
''===============================||====================================
'    If wbCheck = True Then
    
           Workbooks("Container").Worksheets("Container").Activate
           Application.Run ("'Lijsten_new.xlsm'!ProtectOnALL")
    strFile = Path & "/" & ActiveWorkbook.Name
           
        If ActiveWorkbook.CanCheckIn = True Then
           ActiveWorkbook.CheckIn (strFile)
        '    MsgBox "The file has been checked in."
           Stempel_copy = False
        End If
    
'    End If
''===============================||====================================
''===============================||====================================
''===============================||====================================
SpeedOff
       Workbooks(Aanvraagbestand_naam).Worksheets("IN").Activate
    If Stempel_copy = False Then
       Workbooks(Aanvraagbestand_naam).Save
    Else
       MsgBox "Overzetten van aanvragen is mislukt. Probeer het opnieuw."
    End If
End Sub

