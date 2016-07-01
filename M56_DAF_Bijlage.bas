Attribute VB_Name = "M56_DAF_Bijlage"


''===============================||====================================
Sub Bijlage_bestand()

''===============================||====================================
''    Call Apply_UserNames
''===============================||====================================

''===============================||====================================
'SpeedUp the macro
Call SpeedOn
Application.EnableEvents = True
Application.ScreenUpdating = True
''===============================||====================================
'Copy all rows from Worksheet "DATA" to Worksheet RAPPORT_UserGroup
    Dim RAPPORT_OUT
    Dim Sheet As Worksheet
    Dim bestaat As Boolean
    
  ''  set WRSA
    
'    RAPPORT_OUT = "UPLOAD" & "_" & UserGroup
'    RAPPORT_OUT = Range("SET.Bestandsnaam").

    For Each Sheet In Workbooks("ARTIKELBEHEER").Sheets ''ThisWorkbook.Sheets
        If Sheet.Name = RAPPORT_OUT Then bestaat = True: Exit For
    Next Sheet

    If bestaat = True Then
        'Clear the contents of "RAPPORT_OUT" sheet
        Worksheets(RAPPORT_OUT).Select
'        If ActiveSheet.AutoFilterMode = True Then
'        ActiveSheet.AutoFilterMode = False
'        End If
       'Make all Columns / Rows Visible
'        Range("A:DZ").EntireColumn.Hidden = False
'        Range("1:65000").EntireRow.Hidden = False
'        Worksheets(RAPPORT_OUT).Cells.Clear
       ' MsgBox = "Het tabblad " & RAPPORT_OUT & " bestaat al."
    
    ElseIf bestaat = False Then
        Worksheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = RAPPORT_OUT
'   if error aantal tekens tellen max 31 tekens
    End If

'Copy the contents of "DATA" sheet
wbRS.Activate
wbRS.Select
'    Worksheets("Accordering").Visible = xlSheetVisible
'    Worksheets("Accordering").Select
    Range("A1", ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.COPY
    
    
    Worksheets(RAPPORT_OUT).Select
    Range("A1").Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
''    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
''        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
    
'    ActiveWorkbook.Worksheets("Accordering").UsedRange.COPY Destination:=ActiveWorkbook.Worksheets(RAPPORT_OUT).Range("A1")
   'Worksheets("DATA").Visible = xlSheetHidden
''===============================||====================================
'DRAGON filtercolom: This has impact on a new sheet and NOT source sheet
    Worksheets(RAPPORT_OUT).Select
    Range("A1:DZ65000").Select
    Selection.AutoFilter Field:=1, Criteria1:="<>"
    ActiveSheet.Rows("2:5").Delete
''===============================||====================================
'SpeedUp the macro
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
''===============================||====================================


'''===============================||====================================
'   'Makes a copy of the active sheet and save it to a temporary file
'    Dim filename, wb, Path As String
'
'    Worksheets("RAPPORT_" & UserGroup).COPY
'    Set wb = Worksheets("RAPPORT_" & UserGroup)
'    filename = "RAPPORT_" & UserGroup & ".xlsx"
'    Path = "C:\Temp\"
'  ''MsgBox "Look in <" & Path & "> ... for the files!" & vbNewLine & "Look if map <" & Path & "> is present!"
''==============================================================
''Replace existing files
'Application.DisplayAlerts = False   'replacing
'    wb.SaveAs Path & filename, FileFormat:=51
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
'SpeedUp the macroCall SpeedOn
'Call SpeedOff
'Application.EnableEvents = True
'Application.ScreenUpdating = True
'==============================================================
End Sub

