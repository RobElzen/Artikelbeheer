Attribute VB_Name = "M91a_DAF_Ranges_Lijsten"

Sub Generate_Ranges_Lijsten()
    SpeedOn
    Workbooks("Lijsten_new.xlsm").Worksheets("SAVE Blad").Activate
    Workbooks("Lijsten_new.xlsm").Worksheets("SAVE Blad").Select
''========================================================
For Each ws In ActiveWorkbook.Worksheets
'*********************************************************
    ws.Select
    ActiveSheet.Rows("1:1").Select
    Dim di  As Variant
      ''**************
    For Each di In Application.Selection
           If di.Value <> "" Then
      Range("1:1").Replace What:=" ", Replacement:=".", _
          lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False
      Range("1:1").Replace What:="-", Replacement:="_", _
          lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False
   '   Range("1:1").Replace What:=".", Replacement:="",
   '       LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
      Range("1:1").Replace What:="/", Replacement:="", _
          lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False
   '   Range("1:1").Replace What:="""*""", Replacement:="",        ''steretje werkt niet
   '       lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False
      Range("1:1").Replace What:="(", Replacement:="", _
          lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False
      Range("1:1").Replace What:=")", Replacement:="", _
          lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False
      Range("1:1").Replace What:=Chr(10), Replacement:="_", _
          lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False
           End If
     Next di
           
    For Each di In Application.Selection
        If di.Value <> "" Then
        di.Value = Replace(di, " ", ".")
        di.Value = Replace(di, "-", "_")
'       di.Value = Replace(di, ".", "")
        di.Value = Replace(di, "/", "")
        di.Value = Replace(di, "*", "")
        di.Value = Replace(di, "(", "")
        di.Value = Replace(di, ")", "")
        di.Value = Replace(di, Chr(10), "_")
        di.Value = di.Value                '= "EHV." & di.Value 'Adjust Name Heading 1st Row by Adding Affix
        End If
    Next di
'***********************************************************
    Dim r           As Long
    Dim lcol        As Byte
    Dim c           As Byte
        r = 0
        lcol = 0
        c = "0"
       'Count rows present in worksheet Lists per column
        lcol = Cells(1, Columns.Count).End(xlToLeft).Column
    For c = 1 To lcol
        r = Cells(Rows.Count, c).End(xlUp).Row
''''''''
        If ActiveSheet.Name = "UserNames" Then
           r = Cells(Rows.Count, 1).End(xlUp).Row
           Range(Cells(2, c), Cells(r, c)).Name = "USER." & Cells(1, c).Value
    ElseIf ActiveSheet.Name = "SETTINGS" Then
           r = Cells(Rows.Count, 1).End(xlUp).Row
           Range(Cells(2, c), Cells(r, c)).Name = "SET." & Cells(1, c).Value
        Else
           Range(Cells(2, c), Cells(r, c)).Name = "Lst_" & Cells(1, c).Value
        End If
''''''''
    Next
'*********************************************************
Next ws
''========================================================
    SpeedOff
End Sub
