Attribute VB_Name = "M91__DAF_Ranges_ALL"
''' Define Ranges al available workbooks except Lijsten_new.xlsm
Sub Generate_Ranges_ALL()
''========================================================
Call SpeedOn
Call Affix_Case
''========================================================
    ActiveWorkbook.ActiveSheet.Activate
    ActiveWorkbook.ActiveSheet.Select
    
    Application.Run "'Lijsten_new.xlsm'!ProtectOff"
    
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
        di.Value = di.Value                     '= Affix & di.Value 'Adjust Name Heading 1st Row by Adding Affix
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
        r = Cells(Rows.Count, 1).End(xlUp).Row
    For c = 1 To lcol
    
        If r = 5 Then
           r = 6
        Else
           r = r
        End If

        Range(Cells(6, c), Cells(r, c)).Name = Affix & Cells(1, c).Value
    Next
''========================================================
Call SpeedOff
     Application.Run "'Lijsten_new.xlsm'!ProtectOnALL"
End Sub

