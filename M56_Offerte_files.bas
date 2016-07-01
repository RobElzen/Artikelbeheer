Attribute VB_Name = "M56_Offerte_files"
Sub Offerte_TOEVOEGEN_files()
''Select_multiple_files()
''http://analystcave.com/vba-application-filedialog-select-file/
'Quite common is a scenario when you are asking the user to select one or more files.
'The code below does just that. Notice that you need to set AllowMultiSelect to True.

Dim fDialog As FileDialog, result As Integer
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
     
'IMPORTANT!
fDialog.AllowMultiSelect = True
 
'Optional FileDialog properties
fDialog.Title = "Select a file"
fDialog.InitialFileName = "W:\SAP PM docs Ehv\Offertes aanvraag artikelen"
'Optional: Add filters
fDialog.Filters.Clear
fDialog.Filters.Add "All files", "*.*"
'fDialog.Filters.Add "Excel files", "*.xlsx;*.xls;*.xlsm"
'fDialog.Filters.Add "Text/CSV files", "*.txt;*.csv"
'
'Show the dialog. -1 means success!
If fDialog.Show = -1 Then
   
      aantal = fDialog.SelectedItems.Count
   If aantal > 2 Then
      MsgBox "Aantal selected bestanden > 2" & vbNewLine & vbNewLine & _
             "Selecteer max.  2 bestanden!"
      Exit Sub
   End If
teller = 0
  For Each it In fDialog.SelectedItems
'  Debug.Print it
    teller = teller + 1
    ActiveCell.Select
                                                                             'it
                                                                             '"LINK " & teller
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=it, TextToDisplay:=it
    ActiveCell.Offset(0, 2).Select
  Next it
End If

End Sub


  Sub GetImportFileName21()
      Dim FileNames As Variant
      Dim Msg As String
      Dim I As Integer
      FileNames = Application.GetOpenFilename(MultiSelect:=True)
                  With FileNames
                  If UBound(FileNames) > 2 Then
                     MsgBox "SELECTED > 2 FILES"
                     Exit Sub
                  End If
                  End With
      If IsArray(FileNames) Then
          Msg = "You selected:" & vbNewLine
          For I = LBound(FileNames) To UBound(FileNames)
              ''ActiveCell.Select
              ActiveSheet.Hyperlinks.Add Anchor:=ActiveCell, Address:=FileNames(I)
              ActiveCell.Offset(0, 2).Select
'              Msg = Msg & FileNames(i) & vbNewLine
          Next I
'          MsgBox Msg
      Else
          MsgBox "No files were selected."
      End If
  End Sub



