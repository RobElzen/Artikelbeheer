Attribute VB_Name = "M10a_Copy_WERK_to_Container"

Sub copy_WERK_to_Container()
''===============================||====================================
Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
''===============================||====================================
''Define Filter Values into ARRAY/RANGE defined in sheet SETTINGS
    Dim Werkbestand_naam As String
    Dim statusRange As Range
    
    Dim Column_Name_Werkbestand As String
'    Dim Column_Name_Werkbestand_CHECK As Integer
    Dim Column_Name_Container As String
    Dim Column_Name_Container_NR As Integer
    
    Dim rCell As Variant
    Dim I As Integer
    Dim j As Integer
    I = 0
    j = 0
    mylastRow_Container = 0
''===============================||====================================
    Werkbestand_naam = ActiveWorkbook.Name
''===============================||====================================
''===============================||====================================
''====================== CONTROLE OP FOUTEN ===========================
''===============================||====================================
''===============================||====================================
'Count data rows present in source file (only pure data)
    Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Activate
    Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Select
    Set statusRange = Range("WB_Aanvraag.code")
''===============================||====================================
'rCell is positie in RangeName excl. HeadingRows
CHECK_START:
SpeedOn
   For Each rCell In statusRange.Cells
   statusRange(rCell.Row - HeadingRows, 1).Select
If statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_14 Then
''========================================================================''
'If Range("WB_ABC.code").Cells(rCell.Row, 1).Value = "C: Onderdeel zonder relatie tot machine" Then
If (Cells(rCell.Row, Range("WB_ABC.code").Column) = "C: Onderdeel zonder relatie tot machine" And _
   Cells(rCell.Row, Range("WB_Mach.nr.Boom.Aantal").Column) = "") Or _
   Cells(rCell.Row, Range("WB_ABC.code").Column) = "NPG" Then
   Cells(rCell.Row, Range("WB_Mach.nr.Boom.Aantal").Column) = "Boom nvt"
  'MsgBox "Het werkt  " & (rCell.Row) & "  :  " & Cells(rCell.Row, Range("WB_ABC.code").Column).Value & " : " & " Machineboom niet invullen"
End If

If Cells(rCell.Row, Range("WB_Type").Column) = "Handelsartikel" And _
   Cells(rCell.Row, Range("WB_Offerte").Column) = "" Then
   Cells(rCell.Row, Range("WB_Offerte").Column) = "Databeheerder vraagt offerte aan!"
'  MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("WB_Type").Column).Value & " : " & "  Offerte door DataBeheerder laten aanvragen"
End If

If Cells(rCell.Row, Range("WB_Website.producent").Column) = "" And _
   Cells(rCell.Row, Range("WB_Type").Column) <> "Handelsartikel" Then
   Cells(rCell.Row, Range("WB_Website.producent").Column) = "Machinedelen zijn leveranciers maatwerk!"
ElseIf Cells(rCell.Row, Range("WB_Offerte").Column) <> "" And _
   Cells(rCell.Row, Range("WB_Type").Column) = "Handelsartikel" Then
   Cells(rCell.Row, Range("WB_Website.producent").Column) = "nvt."
ElseIf Cells(rCell.Row, Range("WB_Offerte").Column) = "" And _
   Cells(rCell.Row, Range("WB_Type").Column) = "Handelsartikel" And _
   Cells(rCell.Row, Range("WB_Website.producent").Column) = "" Then
   Cells(rCell.Row, Range("WB_Website.producent").Column) = ""
End If

If Cells(rCell.Row, Range("WB_Opmerking.ME").Column) = "" Then
   Cells(rCell.Row, Range("WB_Opmerking.ME").Column) = "nvt."
End If
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
    Set requRange = Workbooks("Lijsten_new").Worksheets("SETTINGS").Range("SET.COL_REQUIRED_WB")
    Set formRange = Workbooks("Lijsten_new").Worksheets("SETTINGS").Range("SET.COL_FORMAT")
    Set charRange = Workbooks("Lijsten_new").Worksheets("SETTINGS").Range("SET.COL_CHAR")
    Set ALL_Range = Workbooks("Lijsten_new").Worksheets("SETTINGS").Range("SET.RANGE_ALL")
''===============================||====================================
    For j = 1 To Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Cells(1, 1).SpecialCells(xlLastCell).Column

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
         Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Cells(rCell.Row, j).Select
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
            
            Selection.NumberFormat = "#,#0.0_ ;-#,#0.0 "
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Selection.NumberFormat = "#,#0.0_ ;-#,#0.0 "
         
         End If
         
         If formRange(hCell.Row - 1, 1).Value = "NE" Then    ''Numeric Exact
         End If
         
         If formRange(hCell.Row - 1, 1).Value = "D" Then    ''Date time
            Selection.NumberFormat = "[$-13]dd-mm-yyyy;@" ''Format(Date, "dd-mm-yyyy")
'            Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Selection.NumberFormat = "[$-13]dd-mm-yyyy;@" ''Format(Date, "dd-mm-yyyy")
         GoTo COL_CHAR
         End If
        
         If requRange(hCell.Row - 1, 1).Value = "V" Then    ''Valuta
            Selection.NumberFormat = "#,##0.00"
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
               Else
                  Selection.Interior.Color = xlNone
               End If
            End If

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
   statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_17
   GoTo CHECK_OVER
Else
   statusRange(rCell.Row - HeadingRows, 1).Select
   statusRange(rCell.Row - HeadingRows, 1).Value = Aanvraag_level_14
   GoTo CHECK_OVER
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CHECK_CHAR_GESLAAGD:
CHECK_OVER:
End If
     rCell = (rCell.Row + 1)
Next rCell



''Controle hoeveel regels zijn er om over te zetten naar Container
        If Application.WorksheetFunction.CountIf(statusRange, Aanvraag_level_14) > 0 Then
           GoTo COPY
        Else
           MsgBox "Er zijn geen regels om over te zetten."
           Exit Sub
        End If
COPY:
''''''''
SpeedOff
SpeedOn
''''''''
    Set Network = CreateObject("wscript.network")
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
Dim xlApp As Excel.Application

'Dim strFile As String
    strFile = Path & "/" & "Container.xlsm"
Dim Stempel_copy As Boolean
    Stempel_copy = False
 
If Workbooks.CanCheckOut(strFile) = True Then
   Workbooks.CheckOut strFile
   Application.Workbooks.Open strFile
   Stempel_copy = True
Else
    MsgBox "Uitchecken van het bestand is niet mogelijk. Probeer later nog een keer."
    Stempel_copy = False
    Exit Sub
End If
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''controle invoeren als bestand al geopend
    Workbooks("Container").Worksheets("Container").Activate
    
    Application.Run ("'Lijsten_new.xlsm'!ProtectOff")
    Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
   'Define Container last row
    mylastRow_Container = Workbooks("Container").Worksheets("Container").UsedRange.Rows.Count
''===============================||====================================
   'vind en copy cell to cell || range to range
   'check hoeveel RECORDS zijn er om in te dienen anders open geen Container
''''''''
SpeedOff
SpeedOn
''''''''
'Count data rows present in source file (only pure data)
    Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Activate
    Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Select
'SETTINGS CALL voor status check "ME_inleveren"
    Set statusRange = Range("WB_Aanvraag.code")
''===============================||====================================
'rCelly is positie in RangeName excl. HeadingRows
COPY_START:
Dim rCelly As Variant
    For Each rCelly In statusRange.Cells
    statusRange(rCelly.Row - HeadingRows, 1).Select
    If statusRange(rCelly.Row - HeadingRows, 1).Value = Aanvraag_level_14 Then
       statusRange(rCelly.Row - HeadingRows, 1).Select
''========================================================================''
    For I = 1 To Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Cells(1, 1).SpecialCells(xlLastCell).Column
''''''''''''''
        Column_Name_Werkbestand = Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Cells(1, I)
        Column_Name_Container = "CNT_" & Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Cells(1, I)
''''''''''''''
        Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Cells(rCelly.Row, I).COPY
''''''''''''''
        Workbooks("Container").Worksheets("Container").Activate ''
        Column_Name_Container_NR = Range(Column_Name_Container).Column
        ActiveSheet.Cells(mylastRow_Container + 1, Column_Name_Container_NR).Select
''With Selection
       'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
       'SkipBlanks:=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
''End With
    Next I
    Cells(mylastRow_Container + 1, Range("CNT_" & "Aanvraag.code").Column) = Aanvraag_level_24
    Cells(mylastRow_Container + 1, Range("CNT_" & "Aanvrager.Naam").Column) = Naam
    Cells(mylastRow_Container + 1, Range("CNT_" & "Datum_OUT_ME").Column) = Now()
    
    Cells(mylastRow_Container + 1, Range("CNT_" & "Aanvrager.Vestiging").Column) = Vestiging
    Cells(mylastRow_Container + 1, Range("CNT_" & "Aanvrager.Afdeling").Column) = Afdeling
    
    mylastRow_Container = mylastRow_Container + 1
''''''''''''''''''''''''''''''''''
    Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Activate
    Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Select
    statusRange(rCelly.Row - HeadingRows, 1).Value = Aanvraag_level_19
Else
'MsgBox "No records to copy!"
End If
''========================================================================''
'CHECK_OVER:
    rCelly = (rCelly.Row + 1)
Next rCelly
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
COPY_END:
''===============================||====================================
    Workbooks("Container").Worksheets("Container").Activate
    mylastRow_Container = Workbooks("Container").Worksheets("Container").UsedRange.Rows.Count
    mylastColumn_Container = Workbooks("Container").Worksheets("Container").UsedRange.Columns.Count

Workbooks("Container").Worksheets("Container").Range(Cells(5, 1), Cells(5, mylastColumn_Container)).Select
Workbooks("Container").Worksheets("Container").Range(Cells(5, 1), Cells(5, mylastColumn_Container)).COPY
Workbooks("Container").Worksheets("Container").Range(Cells(6, 1), Cells(mylastRow_Container, mylastColumn_Container)).Select

    Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
              SkipBlanks:=False, Transpose:=False

    Application.Run ("'Lijsten_new.xlsm'!Generate_Ranges_ALL")
    Application.Run ("'Lijsten_new.xlsm'!ProtectOnALL")
SpeedOff
''===============================||====================================
''===============================||====================================
''===============================||====================================
   Workbooks("Container").Worksheets("Container").Activate
   Workbooks("Container").Worksheets("Container").Select
   strFile = Path & "/" & ActiveWorkbook.Name

If ActiveWorkbook.CanCheckIn = True Then
   ActiveWorkbook.CheckIn strFile
'   ActiveWorkbook.Close (True)
'    MsgBox "The file has been checked in."
   Stempel_copy = False
End If
''===============================||====================================
''===============================||====================================
''===============================||====================================
       Workbooks(Werkbestand_naam).Activate
       Workbooks(Werkbestand_naam).Worksheets("Werkbestand").Select
    If Stempel_copy = False Then
       Workbooks(Werkbestand_naam).Save
    Else
       MsgBox "Overzetten van aanvragen is mislukt. Probeer het opnieuw."
    End If
End Sub







''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
''===============================||====================================
'         If Column_Name_Werkbestand_CHECK > 0 Then ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            If Len(Cells(rCell.Row, j)) > Column_Name_Werkbestand_CHECK Then
'               Cells(rCell.Row, j).Select
'               Selection.Interior.Color = vbRed
'               CHECK_CHAR_FLAG = True
''               MsgBox "Er zijn te lange teksten aanwezig. Svp. aanpassen."
'            Else
'               Cells(rCell.Row, j).Select
'               If Selection.Interior.Color = vbYellow Then
'               Else
'                  Selection.Interior.Color = xlNone
'               End If
'            End If
'         End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Relatiegroep
'               Cells(1, j).Select
'            If Cells(1, j) = "Relatiegroep" Then
'               Cells(rCell.Row, Range("WB_Relatiegroep").Column).Select
'
'               If Not IsNumeric(Selection) Then
'                  'MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("WB_Type").Column).Value & " : " & "  Offerte door DataBeheerder laten aanvragen"
'                  Selection.Interior.Color = vbRed
'                  CHECK_EMPTY_FLAG = True
'               End If
'            End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Veiligheidsvoorraad.Bestelpunt
'               Cells(1, j).Select
'            If Cells(1, j) = "Veiligheidsvoorraad.Bestelpunt" Then
'               Cells(rCell.Row, Range("WB_Veiligheidsvoorraad.Bestelpunt").Column).Select
'
'               If Not IsNumeric(Selection) Then
'                  'MsgBox "Het werkt" & " : " & Cells(rCell.Row - HeadingRows, Range("WB_Type").Column).Value & " : " & "  Offerte door DataBeheerder laten aanvragen"
'                  Selection.Interior.Color = vbRed
'                  CHECK_EMPTY_FLAG = True
'               End If
'            End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Voorraad
'               Cells(1, j).Select
'            If Cells(1, j) = "Voorraad" Then
'               Cells(rCell.Row, Range("WB_Voorraad").Column).Select
'
'               If Not IsNumeric(Selection) Then
'                  Selection.Interior.Color = vbRed
'                  CHECK_EMPTY_FLAG = True
'               End If
'            End If
''''''''''''''


