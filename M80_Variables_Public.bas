Attribute VB_Name = "M80_Variables_Public"
Option Explicit
    Public WB As Workbook
    Public ws As Worksheet
    Public wbA As Workbook         ''Aanvraagbeheer.xlsm
    Public wbAS As Worksheet       ''ActiveSheet
    
    Public wsL As Worksheet
    Public wbS As Worksheet        ''SETTINGS
    Public wbL As Workbook         ''Lijsten_New.xlsm
    Public wsLS As Worksheet       ''SETTINGS
    Public wsLU As Worksheet       ''UserNames

    Public wbR As Workbook         ''Rapport.xlsm
    Public wbRS As Worksheet       ''RapportSheet
    Public wbk As Workbook         ''Container.xlsm
    
    Public USERLEVEL As Variant
    Public USERMAIL As Variant
    Public FirstUserColumn As Integer
    Public LastUserColumn As Integer
    Public UserColumn As Integer
    Public UserGroup As String
   'Public UserGroupVestiging As String
''''''''''
    Public WorksheetName As Integer
    Public FirstWorksheetName As Integer
    
    Public HeaderName As Integer
    Public HeaderNameColumn As Integer
    Public Const HeadingRows As Integer = 5
''''''''''
    'UserGroup
    Public RAPPORT_ALL As Variant
    Public chk_RapportALL As Boolean      ''test
    Public chk_MailUsers As Boolean       ''test
    Public chk_MailAllUsers As Boolean    ''test
    Public chk_KPIexport As Boolean       ''test
    Public Mail_ALL_Users As Variant
    Public MAIL_User As Variant
    Public MAIL_ALL As Variant
    Public KPIexport As Variant
    '
    Public HiddenColumnName As Range
    Public mylastDATA As Integer
    
    Public mylastRow_Werkbestand As Long    ''Integer
    Public mylastRow_Container As Integer
    Public mylastRow_IN As Integer
    Public mylastRow_Accordering As Integer
    Public mylastRow_OUT As Integer
    Public mylastRow_Databestand As Integer

    Public mylastColumn_Werkbestand As Long    ''Integer
    Public mylastColumn_Container As Integer
    Public mylastColumn_IN As Integer
    Public mylastColumn_Accordering As Integer
    Public mylastColumn_OUT As Integer
    Public mylastColumn_Databestand As Integer
'    Dim Buttons() As New Class1
    Public Network As Object
    Public filename As String
    Public Const Path As String = "http://dafshare-org.eu.paccar.com/organization/ops-mtc/Artikelaanvraag"
   'Public Const Path As String = "http://dafshare-org.eu.paccar.com/organization/ops-mtc/MMP/ArtikelBeheer"
    Public rngHTML As Range
    Public Const Path_C As String = "C:\Temp\"
    Public Interrupt As Boolean
''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''
Public Affix As String
Public strFile As String

Public UsName As String
Public UsNameRev As String
Public FirstName As String
Public LastName As String
Public UserInitials As String
''''''
Public UserName As String
Public Initials As String
Public Naam As String
Public Voornaam As String
Public Achternaam As String
Public Team As String
Public Afdeling As String
Public Vestiging As String
Public Role As String
Public Functie As String
Public Email As String
Public PC_Nummer As String
Public Werkblad_IN_DB As String
Public Persnr As Integer
Public Niveau As Integer
''''''''''''''''''''''''''''''
''Public UserME As Boolean
''Public CSVmap As String
''''''''''''''''''''''''''''''

'Public LastRowME, LastColumnME, FirstRecRowME, RecordsME, NextRowME As Integer
'Public ME_FileName, ME_FileNameDef, SheetME  As String
'Public KolME_RecdNr, KolME_NaarAV, KolME_NaamAV, KolME_Offert, KolME_Opmkng, KolME_Statnr, KolME_Gewich, KOLME_GewBru, KOLME_GewEen As String 'toegevoegd voor statistieknummer'
'
'Public LastRowAV, LastColumnAV, FirstRecRowAV, RecordsAV, NextRowAV As Integer
'Public Cnt_LDafOm, Cnt_LInkBs, Cnt_LProdu, Cnt_PntKom, Cnt_PrsPnt, Cnt_EmpCel, CntRelGrp As Integer
'Public Cnt_RecdNr, CntTot_RecdNr, Cnt_DafOms, CntTot_DafOms As Integer
'Public AV_FileName, SheetAV As String
'Public KolAV_RecdNr, KolAV_NaamAV, KolAV_DatmAV, KolAV_AbcTek, KolAV_PlnSap, KolAV_RelGrp, KolAV_DafOms, KolAV_LDafOm, KolAV_InkBes, KolAV_LInkBs, KolAV_Produc, KolAV_LProdu, KolAV_VeiVrd, KolAV_InkPrs, KolAV_LevrNr, KolAV_LevrNm, KolAV_ArtNrL, KolAV_PCNcde, KolAV_Status, KolAV_Offert, KolAV_Opmkng, KolAV_Statnr, KolAV_GewBru, KolAV_Gewich, KolAV_GewEen As String 'toegevoegd voor statistieknummer'
'Public AV_BufVar As String
'Public AantNwRijen As Integer
'Public MarkKleurAV As Integer
'Public Lim_DafOms, Lim_InkBes, Lim_Produc, Lim_RelGrp As Integer
'
'Public LastRowTO, LastColumnTO, TitelRowTO, FirstRecRowTO, RecordsTO, NextRowTO As Integer
'Public TO_FileName, SheetTO As String
'Public KolTO_RecdNr, KolTO_MatrNr, KolTO_Status, KolTO_Offert, KolTO_Opmkng, KolTO_ScrnOk, KolTO_Akkor1, KolTO_Akkor2, KolTO_Akkor3, KolTO_Akkor4, KolTO_RdyUpl, KolTO_Opgvrd, KolTO_Cntrct, KolTO_Datum1 As String
'Public KolTO_AfdeAV, KolTO_BedrAV As String
'Public KolTO_NaamAV, KolTO_DatmAV, KolTO_ArtTyp, KolTO_AbcTek, KolTO_PlnSap, KolTO_RelGrp, KolTO_DafOms, KolTO_LDafOm, KolTO_InkBes, KolTO_LInkBs, KolTO_Produc, KolTO_LProdu As String
'Public KolTO_VeiVrd, KolTO_AfrWrd, KolTO_BesEnh, KolTO_BasEnh, KolTO_MinSer, KolTO_Levrtd, KolTO_InkPrs, KolTO_PrsPer, KolTO_InkGrp, KolTO_LevrNr, KolTO_ArtNrL, KolTO_PCNcde, KolTO_Locati, KolTO_Statnr, KolTO_GewBru, KolTO_Gewich, KOLTO_GewEen As String 'toegevoegd voor statistieknummer'
'Public TO_FilterVar As String
'



