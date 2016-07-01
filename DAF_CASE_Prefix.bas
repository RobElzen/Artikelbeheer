Attribute VB_Name = "DAF_CASE_Prefix"

Sub Affix_Case()
''========================================================
    If ActiveWorkbook.Name <> "Artikelbeheer.xlsm" Then
    
    If ActiveSheet.Name = "Werkbestand" Then
       Affix = "WB_"
ElseIf ActiveSheet.Name = "Container" Then
       Affix = "CNT_"
ElseIf ActiveSheet.Name = "Databestand" Then
       Affix = "DB_"
End If
Else

    If ActiveSheet.Name = "IN" Then
       Affix = "IN_"
ElseIf ActiveSheet.Name = "Accordering" Then
       Affix = "ACC_"
ElseIf ActiveSheet.Name = "OUT" Then
       Affix = "OUT_"
End If

End If
''========================================================
End Sub

