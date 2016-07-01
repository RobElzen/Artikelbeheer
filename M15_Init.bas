Attribute VB_Name = "M15_Init"
Option Explicit
    
'Diverse macros gedefinieerd
'
'Apply_UserNames
'Apply_UserNames
'SheetsVeryHidden
'LijstenVeryHidden
'RetrieveUser
'FirstRun
'
'''''''In uitwerking: Column Visibility'''''TODO'''''''''''''''
'Define_Worksheets_Visibility
'
'
'
Public tmpMatchHeader As Variant
Public tmpIndexHeader As Variant
Public tmpHeaderName As Integer
Public tmpHeaderNameColumn As Integer

''===============================||====================================
Sub Apply_UserNames()

Application.Run ("'Lijsten_new.xlsm'!SpeedOn")

Dim UserName        As String
Dim Computername    As String
Dim Password        As String
Dim MyError         As Integer
    
Dim Network As Object
Set Network = CreateObject("wscript.network")
    Err = 0
    MyError = 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set wbA = Workbooks(ActiveWorkbook.Name)
'Set wbAS = wbA.Worksheets(ActiveSheet.Name)

Set wbL = Workbooks("Lijsten_New.xlsm")
Set wsLS = wbL.Worksheets("SETTINGS")
Set wsLU = wbL.Worksheets("UserNames")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Positie van Column "UserName" in TabelUsers bepalen waar de namen moeten worden vergeleken met Gebruikers Inlognaam
    wsLU.Activate
    ''''
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If
    ''''
    Dim HeaderName As Integer
    HeaderName = 0
    For HeaderName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If wsLU.Cells(1, HeaderName).Value = "UserName" Then
        HeaderNameColumn = HeaderName
        End If
    Next
    wbL.Activate

    tmpIndexHeader = Application.Index(Range("TableUsers"), , HeaderNameColumn)
    tmpMatchHeader = Application.Match(Network.UserName, tmpIndexHeader, 0)
               If IsError(tmpMatchHeader) Then
                  MsgBox "Onbekend Gebruiker" & Network.UserName
                  
                  Niveau = 2
                  Naam = "ONBEKEND: " & Network.UserName
                  Vestiging = "NL"
                  Role = "ME"
                  Afdeling = "Onbekend"
                  Computername = Network.Computername
'                  ActiveWorkbook.Close
                  Exit Sub
               End If
    
    UserName = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
'''''''
   'Positie van Column "Niveau" in TabelUsers bepalen waar de Niveau moeten worden afgeleid worden van Gebruikers Inlognaam
        HeaderName = 0
    For HeaderName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If wsLU.Cells(1, HeaderName).Value = "Niveau" Then
'       If Range("TableUsers")[[#Headers],[Niveau]] = "Niveau" Then
        HeaderNameColumn = HeaderName
        End If
    Next
    
    Niveau = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
'''''''
   'Positie van Column "Role" in TabelUsers bepalen waar de Niveau moeten worden afgeleid worden van Gebruikers Inlognaam
        HeaderName = 0
    For HeaderName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If wsLU.Cells(1, HeaderName).Value = "Role" Then
        HeaderNameColumn = HeaderName
        End If
    Next

    Role = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
'''''''
   'Positie van Column "Vestiging" in TabelUsers bepalen waar de Vestiging moet worden afgeleid worden van Gebruikers Inlognaam
        HeaderName = 0
    For HeaderName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If wsLU.Cells(1, HeaderName).Value = "Vestiging" Then
        HeaderNameColumn = HeaderName
        End If
    Next

    Vestiging = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
'''''''
   'Positie van Column "Afdeling" in TabelUsers bepalen waar de Afdeling moet worden afgeleid worden van Gebruikers Inlognaam
        HeaderName = 0
    For HeaderName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If wsLU.Cells(1, HeaderName).Value = "Afdeling" Then
        HeaderNameColumn = HeaderName
        End If
    Next

    Afdeling = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
'''''''
   'Positie van Column "Afdeling" in TabelUsers bepalen waar de Afdeling moet worden afgeleid worden van Gebruikers Inlognaam
        HeaderName = 0
    For HeaderName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If wsLU.Cells(1, HeaderName).Value = "Naam" Then
        HeaderNameColumn = HeaderName
        End If
    Next

    Naam = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
'''''''
   'Positie van Column "Afdeling" in TabelUsers bepalen waar de Afdeling moet worden afgeleid worden van Gebruikers Inlognaam
        HeaderName = 0
    For HeaderName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If wsLU.Cells(1, HeaderName).Value = "Email" Then
        HeaderNameColumn = HeaderName
        End If
    Next

    Email = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Positie van Column "PC_Nummer" in TabelUsers bepalen waar de Computername moeten worden afgeleid worden van Gebruikers PC_Nummer
        HeaderName = 0
    For HeaderName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If Cells(1, HeaderName).Value = "PC_Nummer" Then
        HeaderNameColumn = HeaderName
        End If
    Next

Computername = Application.WorksheetFunction.Index(Range("TableUsers"), tmpMatchHeader, HeaderNameColumn)
'''''''
'   http://peltiertech.com/structured-referencing-excel-tables/
'   Niveau = "TableUsers"[[#23], [Niveau]]
'''''''
''===============================||====================================
'          Hide werkbladen in Lijsten_new.xlsm
           Application.Run ("'Lijsten_new.xlsm'!SpeedOn")
           Application.Run ("SheetsVeryHidden_Lijsten")
           Application.Run ("'Lijsten_new.xlsm'!SpeedOn")
''===============================||====================================
'           Application.Run ("'Lijsten_new.xlsm'!SpeedOn")
'           Application.Run ("SheetsVeryHidden_wbA")
'           Application.Run ("'Lijsten_new.xlsm'!SpeedOn")
'wbAS.Select
End Sub


'''============================Define:Worksheets HIDDEN==============''
Sub SheetsVeryHidden_wbA()
Application.Run ("'Lijsten_new.xlsm'!SpeedOn")
   'Hier moet logica komen van welke sheets verbergen en welke niet
   'ook voor "Onbekende gebruiker"
    Set wbA = Workbooks("Artikelbeheer.xlsm")
    wbA.Activate
    wbA.Worksheets("INTRO").Visible = xlSheetVisible
''wbA.Worksheets("Legenda").Visible = xlSheetVeryHidden

    wbA.Worksheets("IN").Visible = xlSheetVeryHidden
    wbA.Worksheets("Accordering").Visible = xlSheetVeryHidden
    wbA.Worksheets("OUT").Visible = xlSheetVeryHidden
''''''''''''''
    wbA.Worksheets("InfoButtons").Visible = xlSheetVeryHidden
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim LaatsteRegel As Integer
Dim LaatsteColumn As Integer
Dim BladNaam As String
Dim tmpIndexSheet As Variant
Dim m As Integer

LaatsteRegel = 0
LaatsteColumn = 0
LaatsteRegel = wsLU.UsedRange.Rows.Count
LaatsteColumn = wsLU.UsedRange.Columns.Count
''===
    tmpIndexSheet = 0
    m = 0
    BladNaam = ""
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.Run ("'Lijsten_new.xlsm'!SpeedOn")

For Each wbA In Workbooks
 If wbA.Name = "Artikelbeheer.xlsm" Then
 
    WorksheetName = 0
    FirstWorksheetName = 0
    For WorksheetName = 1 To wsLU.Cells(1, 1).SpecialCells(xlLastCell).Column
        If wsLU.Cells(1, WorksheetName).Value = "INTRO" Then
        FirstWorksheetName = WorksheetName
        End If
    Next WorksheetName
    
       For m = FirstWorksheetName To LaatsteColumn
'           wsLU.Activate
           tmpIndexSheet = Application.WorksheetFunction.Index(wsLU.Range("TableUsers"), tmpMatchHeader, m)
                           BladNaam = wsLU.Cells(1, m).Value

                        If tmpIndexSheet <> "" Then
'                               BladNaam = Application.WorksheetFunction.Index(Range("TableUsers"), HeaderNameColumn, m).Offset(-1, 0)
                               Workbooks(wbA.Name).Worksheets(BladNaam).Visible = xlSheetVisible
                               Workbooks(wbA.Name).Worksheets(BladNaam).Activate
                               Workbooks(wbA.Name).Worksheets(BladNaam).Select

                               If BladNaam = "IN" Then
                                  If Niveau = 1 Then
                                       ActiveSheet.Shapes("btnCNT_to_IN").Visible = True
                                       ActiveSheet.Shapes("btnAanvragenDELETE_IN").Visible = True
                                  Else
                                       ActiveSheet.Shapes("btnCNT_to_IN").Visible = False
                                       ActiveSheet.Shapes("btnAanvragenDELETE_IN").Visible = False
                                  End If
                           ElseIf BladNaam = "Accordering" Then
                                  If Niveau = 1 Then
                                       ActiveSheet.Shapes("btnProtectOFF").Visible = True
                                       ActiveSheet.Shapes("btnRapportACC").Visible = True
                                       ActiveSheet.Shapes("btnAanvragenDELETE_ACC").Visible = True
                                       ActiveSheet.Shapes("btnACC_to_OUT").Visible = True
                                  Else
                                       ActiveSheet.Shapes("btnProtectOFF").Visible = False
                                       ActiveSheet.Shapes("btnRapportACC").Visible = True
                                       ActiveSheet.Shapes("btnAanvragenDELETE_ACC").Visible = False
                                       ActiveSheet.Shapes("btnACC_to_OUT").Visible = False
                                  End If

                           ElseIf BladNaam = "OUT" Then
                                  If Niveau = 1 Then
                                       ActiveSheet.Shapes("btnRapportOUT").Visible = True
                                       ActiveSheet.Shapes("btnAanvragenDELETE_OUT").Visible = True
                                       ActiveSheet.Shapes("btnInitColumns").Visible = True
                                  Else
                                       ActiveSheet.Shapes("btnRapportOUT").Visible = True
                                       ActiveSheet.Shapes("btnAanvragenDELETE_OUT").Visible = False
                                       ActiveSheet.Shapes("btnInitColumns").Visible = False
                                  End If
                           Else
                           End If
'
'                                       ActiveSheet.Unprotect Password:="123"
'                                       ActiveSheet.Range("A17") = Niveau
'                                       ActiveSheet.Protect Password:="123"
                        Else
                                  wbA.Worksheets(BladNaam).Visible = xlSheetVeryHidden
                        End If
'                Call Init_Columns
       Next m
Else
End If

Next wbA
           
End Sub
''===============================||====================================

'''============================Define:Worksheets HIDDEN==============''
Sub SheetsVeryHidden_Lijsten()
Application.Run ("'Lijsten_new.xlsm'!SpeedOn")

Set wbL = Workbooks("Lijsten_New.xlsm")
    wbL.Activate
SpeedOn
If Role = "MMP" Or Role = "DB" Then
   
   'Hier moet logica komen van welke sheets verbergen en welke niet
   'ook voor "Onbekende gebruiker"
    wbL.Worksheets("SETTINGS").Visible = xlSheetVisible
    wbL.Worksheets("UserNames").Visible = xlSheetVisible
    wbL.Worksheets("User").Visible = xlSheetVisible
    wbL.Worksheets("Aanvraag_code").Visible = xlSheetVisible
    wbL.Worksheets("Algemeen").Visible = xlSheetVisible

    wbL.Worksheets("Leverancier").Visible = xlSheetVisible
    wbL.Worksheets("Producent").Visible = xlSheetVisible
    wbL.Worksheets("Statistieknr").Visible = xlSheetVisible
''''''''''''''
    wbL.Worksheets("Interface").Visible = xlSheetVisible
Else
    wbL.Worksheets("SETTINGS").Visible = xlSheetVisible
    wbL.Worksheets("UserNames").Visible = xlSheetVisible
'    wbL.Worksheets("SETTINGS").Visible = xlSheetVeryHidden
'    wbL.Worksheets("UserNames").Visible = xlSheetVeryHidden

    wbL.Worksheets("User").Visible = xlSheetVeryHidden
    wbL.Worksheets("Aanvraag_code").Visible = xlSheetVeryHidden
    wbL.Worksheets("Algemeen").Visible = xlSheetVeryHidden

    wbL.Worksheets("Leverancier").Visible = xlSheetVeryHidden
    wbL.Worksheets("Producent").Visible = xlSheetVeryHidden
    wbL.Worksheets("Statistieknr").Visible = xlSheetVeryHidden
''''''''''''''
    wbL.Worksheets("Interface").Visible = xlSheetVeryHidden
End If

    wbL.Worksheets("SAVE Blad").Visible = xlSheetVisible
    
SpeedOff
End Sub
''===============================||====================================


Sub FirstRun()
''===============================||====================================''
'Voorkomen bij het openen van Artikelbeheer.xlsm om bij Niveau 1 rapporten te draaien
    If Niveau = 1 Then
Else
        
'          Workbooks("Artikelbeheer.xlsm").Activate
'          Worksheets("Accordering").Select

        If Role = "ME" Then
           Worksheets("OUT").Activate
           Worksheets("OUT").Select
    ElseIf Role <> "ME" Then
           Worksheets("Accordering").Activate
           Worksheets("Accordering").Select
    End If

    Application.Run "'Lijsten_new.xlsm'!RAPPORT_OUT"
End If
''===============================||====================================''
End Sub

'
'Private Sub RetrieveUser()
''Er zijn meerdere manieren om de gebruikersnaam te achterhalen:
''- Application.Username geeft de Gebruikersnaam ingevuld onder Excel\Opties\Algemeen
''- Environ("username") geeft de Windows inlognaam, echter gedetacheerden / stagiairs beginnen dan met "a-"!
'
'UsName = Application.UserName
'UsNameRev = StrReverse(UsName)
'
''Gebruikers initialen bepalenop basis van eerste naamletters
'On Error GoTo User_XX
'FirstName = Left(UsName, Application.Find(" ", UsName) - 1)
'LastName = Right(UsName, Application.Find(" ", UsNameRev) - 1)
'UserInitials = (Left(FirstName, 1) & Left(LastName, 1))
'Exit Sub
'
'User_XX:
'    FirstName = "X"
'    LastName = "X"
'    UserInitials = "XX"
'    'Overbodig stukje code (we geven vaste initialen aan gebruiker)
'End Sub
'

