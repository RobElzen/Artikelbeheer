Attribute VB_Name = "M102_ExportModules"

'In VBA, go to the Tools menu, choose References, and select
' "Microsoft Visual Basic For Applications Extensibility Library".

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.vbcomponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If

    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)

    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If

    szExportPath = FolderWithVBAProjectFiles & "\"

    For Each cmpComponent In wkbSource.VBProject.VBComponents

        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select

        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName

        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent

        End If

    Next cmpComponent

    MsgBox "Export is ready"
End Sub



Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String
    Dim DefPath As String
    
    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")
    SpecialPath = "W:\MRO\05) Projecten overig\5.2) Herontwerp artikelaanvragen\VBA\bas temp"
'   SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If

    If FSO.FolderExists(SpecialPath & Format(Now, "ddmmyyyy") & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & Format(Date, "ddmmyyyy") & "VBAProjectFiles"
        On Error GoTo 0
    End If

    If FSO.FolderExists(SpecialPath & Format(Date, "ddmmyyyy") & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & Format(Date, "ddmmyyyy") & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If

End Function
