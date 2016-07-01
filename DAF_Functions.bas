Attribute VB_Name = "DAF_Functions"
Option Explicit

'LetOp
'Als in "str" een teken voorkomt ("<" of ">") dan moeten er twee opgevoerd zijn
'voorbeeld    <>ACC_OUT,<>*
'Anders mogen er meerdere voorwaarden opgevoerd zijn (onbeperkt aantal)
'voorbeeld    ACC_IN,ACC_in_behandeling,ACC_inleveren

Function ContainsGTLT(str) 'GreaTer Lower
    Dim plaats As Integer
    plaats = 0
    plaats = InStr(str, "<") + InStr(str, ">") ''InStr(str, ",") ''
    If plaats > 0 Then
       ContainsGTLT = True
    Else
       ContainsGTLT = False
    End If
End Function
''===============================||====================================
Function RangetoHTML(rngHTML As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim FSO As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rngHTML.COPY
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ts = FSO.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set FSO = Nothing
    Set TempWB = Nothing
End Function


Function ColumnLetter(Col As Long)
     '-----------------------------------------------------------------
    Dim sColumn As String
    On Error Resume Next
    sColumn = Split(Columns(Col).Address(, False), ":")(1)
    On Error GoTo 0
    ColumnLetter = sColumn
End Function


Function Col_Letter(lngCol As Long) As String
Dim vArr
vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)
End Function

'Public Function ColumnLetter1(Column As Integer) As String
'    If Column < 1 Then Exit Function
'    ColumnLetter = ColumnLetter(Int((Column - 1) / 26)) & Chr(((Column - 1) Mod 26) + Asc("A"))
'End Function

Function ColName(colNum As Integer) As String
    ColName = Split(Worksheets(1).Cells(1, colNum).Address, "$")(1)
End Function


'Function ConvertToLetter1(iCol As Integer) As String
'   Dim iAlpha As Integer
'   Dim iRemainder As Integer
'   iAlpha = Int(iCol / 27)
'   iRemainder = iCol - (iAlpha * 26)
'   If iAlpha > 0 Then
'      ConvertToLetter = Chr(iAlpha + 64)
'   End If
'   If iRemainder > 0 Then
'      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
'   End If
'End Function
