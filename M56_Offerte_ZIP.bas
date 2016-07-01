Attribute VB_Name = "M56_Offerte_ZIP"

'http://www.rondebruin.nl/win/s7/win001.htm
'Create empty Zip File
'Changed by keepITcool Dec-12-2005

'Browse to the folder you want and select the file or files

Sub Offerte_TOEVOEGEN_zip()
    Call Affix_Case
   'Zip_File_Or_Files()
    Dim strDate As String, sFName As String
    Dim DefPath As String, DefPath_NL As String, DefPath_BE As String
    Dim oApp As Object, iCtr As Long, I As Integer
    Dim FName, vArr, FileNameZip, filename
    
'   DefPath = Application.DefaultFilePath
'    DefPath = "W:\SAP PM docs Ehv\Offertes aanvraag artikelen"
'    DefPath_NL = "W:\SAP PM docs Ehv\Offertes aanvraag artikelen"
'  DefPath_BE = "W:\SAP PM docs Ehv\Offertes aanvraag artikelen"
    DefPath = "\\eu.paccar.com\dafehv\SAP PM docs Ehv\Offertes aanvraag artikelen\"
      If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    filename = ActiveSheet.Cells(ActiveCell.Row, 1) & ".zip"
    If ActiveSheet.Cells(ActiveCell.Row, 1) = "" Or _
       ActiveCell.Row < 6 Or _
       Application.Intersect(ActiveCell, Range(Affix & "Offerte")) Is Nothing Then
       MsgBox "Selecteer een geldige aanvraagregel in kolom Offerte."
      Exit Sub
    End If
'   FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"
    FileNameZip = DefPath & filename

    'Browse to the file(s), use the Ctrl key to select more files
                                                 ':="Excel Files (*.xl*), *.xl*",
    FName = Application.GetOpenFilename(filefilter:="All Files (*.*), *.*", _
                    MultiSelect:=True, Title:="Select the files you want to zip")
    If IsArray(FName) = False Then
        'do nothing
'    ElseIf FName.SelectedItems.Count = 1   Then "enkel bestand toevoegen"
'    ElseIf UBound(FName) = 1               Then "enkel bestand toevoegen"
       MsgBox "Geen bestanden geselecteerd."
   Else
        'Create empty Zip File
        NewZip (FileNameZip)
        Set oApp = CreateObject("Shell.Application")
        I = 0
        For iCtr = LBound(FName) To UBound(FName)
            vArr = Split97(FName(iCtr), "\")
            sFName = vArr(UBound(vArr))
            If bIsBookOpen(sFName) Then
                MsgBox "You can't zip a file that is open!" & vbLf & _
                       "Please close it and try again: " & FName(iCtr)
            Else
                'Copy the file to the compressed folder
                I = I + 1
                oApp.Namespace(FileNameZip).CopyHere FName(iCtr)

                'Keep script waiting until Compressing is done
                On Error Resume Next
                Do Until oApp.Namespace(FileNameZip).items.Count = I
                    Application.Wait (Now + TimeValue("0:00:01"))
                Loop
                On Error GoTo 0
            End If
        Next iCtr

                                                                                      'FileName & ".zip"
                                                                                      'FileNameZip
    ActiveCell.Select                                                                 '"LINK to files"
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=FileNameZip, TextToDisplay:=FileNameZip
    MsgBox "You find the zip file here: " & vbNewLine & vbNewLine & _
            FileNameZip
    End If
End Sub
''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''
'Code used by every example macro on this page
'Every macro use the sub NewZip and the first example also use both functions.
Sub NewZip(sPath)
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


Function Split97(sStr As Variant, sdelim As String) As Variant
'Tom Ogilvy
    Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function



