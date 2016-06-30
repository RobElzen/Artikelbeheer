Attribute VB_Name = "DAF_Case_NIEUW"
Option Explicit

Sub Aanvraag_code_Nieuw()

Dim Werkbestand_Aanvrager As String
    Werkbestand_Aanvrager = ActiveWorkbook.Name
    
    Workbooks(Werkbestand_Aanvrager).Worksheets("Werkbestand").Activate
    Workbooks(Werkbestand_Aanvrager).Worksheets("Werkbestand").Select

    mylastRow_Werkbestand = Workbooks(Werkbestand_Aanvrager).Worksheets("Werkbestand").UsedRange.Rows.Count
    mylastColumn_Werkbestand = Workbooks(Werkbestand_Aanvrager).Worksheets("Werkbestand").UsedRange.Columns.Count
'===============================||====================================
SpeedOn
'===============================||====================================
ActiveSheet.Range(Cells(5, 1), Cells(5, mylastColumn_Werkbestand)).COPY
'ActiveSheet.Range(Cells(mylastRow_Werkbestand, 1), Cells(mylastRow_Werkbestand, mylastColumn_Werkbestand)).COPY
ActiveSheet.Range(Cells(mylastRow_Werkbestand + 1, 1), Cells(mylastRow_Werkbestand + 1, mylastColumn_Werkbestand)).Select
'ActiveSheet.Range(Cells(6, 1), Cells(mylastRow_Werkbestand + 1, mylastColumn_Werkbestand)).Select
With Selection
    Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteLink, Operation:=xlNone,
'        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone,
'        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone,
'        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone,
'        SkipBlanks:=False, Transpose:=False
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone,
'        SkipBlanks:=False, Transpose:=False

   'Delete content met opmaak behoud (Content: vorige regel inhoud deleten en alleen eerste twee velden vullen)
    Selection.Cells.ClearContents
   'Selection.EntireRow.Delete
End With
   
Range("A4") = Range("A4") + 1
'create variable
 Dim CurrentFileNameNoExtension As String
 'set variable
 CurrentFileNameNoExtension = Left(ActiveWorkbook.Name, (InStrRev(ActiveWorkbook.Name, ".", -1, vbTextCompare) - 1))
 'place file name in cell A3 without file extension
ActiveSheet.Range("A" & mylastRow_Werkbestand + 1).Value = Mid(CurrentFileNameNoExtension, 6, 15) & "-" & Range("A4")
'
'''For each    "Blanco"    vernieuw    Network.UserName  <<<<<<<<<<<<<<<<<<<
ActiveSheet.Range("B" & mylastRow_Werkbestand + 1).Value = "NIEUW"
'Set Network = CreateObject("wscript.network")
'ActiveSheet.Range("C" & mylastRow_Werkbestand + 1).Value = Network.UserName         ''Format(Dag, "dd-mm-yyyy")
''ActiveSheet.Range("C" & mylastRow_Werkbestand + 1).Value = Date  ''Format(Now, "dd-mm-yyyy")''FormatDateTime(Nu, vbShortTime)
ActiveSheet.Range("B" & mylastRow_Werkbestand + 1).Select
'===============================||====================================
SpeedOff
'===============================||====================================
End Sub

