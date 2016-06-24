Attribute VB_Name = "Module6_LSMW_Voorbereiding"
Private Sub SB_Voorbereiding_tbv_LSMW()


'Hier komt een stukje code
''
''
If MsgBox("Weet u zeker dat u CSV bestanden wilt genereren voor de op te voeren artikelen?" & Chr(13) & _
  "Als u doorgaat zal er een nieuw excel bestand worden aangemaakt.", vbYesNo) = vbNo Then Exit Sub Else

'----- 1. Declaratie Variabelen -----
Application.Run "DeclVar_AV" 'Vanwege AV bestandsnaam
Application.Run "DeclVar_TO"

Dim i, j, k As Integer
Dim LastRowLS_Ma, LastColumnLS_Ma, TitelRowLS_Ma, FirstRecRowLS_Ma, RecordsLS_Ma, NextRowLS_Ma As Integer
Dim WbookLSMW As String
Dim SheetLS_Ma, SheetLS_St, SheetLS_Tk, SheetLS_Ib, SheetLS_If, SheetLS_Rd, SheetLS_Ct, SheetsLS_Pk As String '(SheetLS_Ct toegevoegd voor statistieknr)'
Dim wsLS_Ma, wsLS_St, wsLS_Tk, wsLS_Ib, wsLS_If, wsLS_Rd, wsLS_Ct, wsLS_Pk, wsTO As Worksheet '(wsLS_Ct toegevoegd als statistieknummer)'

Dim KolLS_Ma_MatrNr, KolLS_Ma_ArtTyp, KolLS_Ma_AbcTek, KolLS_Ma_PlnSap, KolLS_Ma_RelGrp, KolLS_Ma_DafOms, KolLS_Ma_InkBes, KolLS_Ma_Produc As String
Dim KolLS_Ma_VeiVrd, KolLS_Ma_AfrWrd, KolLS_Ma_BesEnh, KolLS_Ma_BasEnh, KolLS_Ma_MinSer, KolLS_Ma_Levrtd, KolLS_Ma_InkPrs, KolLS_Ma_PrsPer As String
Dim KolLS_Ma_InkGrp, KolLS_Ma_LevrNr, KolLS_Ma_ArtNrL, KolLS_Ma_PCNcde, KolLS_Ma_Locati, KolLS_MA_Statnr, KolLS_Ma_GewBru, KolLS_Ma_Gewich, KolLS_Ma_GewEen As String 'toegevoegd voor statistieknummer'

Dim KolLS_St_MatrNr, KolLS_St_AbcTek, KolLS_St_PlnSap, KolLS_St_RelGrp, KolLS_St_DafOms, KolLS_St_InkBes, KolLS_St_Produc As String
Dim KolLS_St_VeiVrd, KolLS_St_AfrWrd, KolLS_St_BesEnh, KolLS_St_BasEnh, KolLS_St_MinSer, KolLS_St_Levrtd, KolLS_St_InkPrsS, KolLS_St_InkPrsV, KolLS_St_PrsPer As String
Dim KolLS_St_InkGrp, KolLS_St_LevrNr, KolLS_St_ArtNrL, KolLS_St_PCNcde, KolLS_St_Locati, KolLS_St_Statnr, kolls_St_GewBru, KolLS_St_Gewich, KolLS_St_GewEen As String 'toegevoegd voor statistieknummer'

Dim KolLS_Tk_MatrNr, KolLS_Tk_DafOms As String
Dim KolLS_Ib_MatrNr, KolLS_Ib_InkBes As String
Dim KolLS_If_LevrNr, KolLS_If_MatrNr, KolLS_If_ArtNrL, KolLS_If_InkPrs, KolLS_If_PrsPer As String
Dim KolLS_Rd_MatrNr, KolLS_Rd_InkPrsS, KolLS_Rd_PrsPer As String

Dim KolLS_Ct_MatrNr As String 'toegevoegd voor statistieknummer'
Dim KolLS_Pk_MatrNr As String 'toegevoegd voor plankenmerk'
Dim KolLS_Pk_VeiVrd As String 'toegevoegd voor plankenmerk'

'----- 2. Initiele waarden Variabelen -----
WbookLSMW = "LSMW materiaal " & Date & ".xls"
SheetLS_Ma = "Master"
SheetLS_St = "Stam"
SheetLS_Tk = "Tkt EN-NL"
SheetLS_Ib = "InkBestTkt"
SheetLS_If = "Inforecord"
SheetLS_Rd = "Repdelen"
SheetLS_Ct = "Statistieknr" '(toegevoegd voor statistieknummer)'
SheetLS_Pk = "V1bestuur" '(toegevoegd voor plankenmerk)'

TitelRowLS_Ma = 1
FirstRecRowLS_Ma = TitelRowLS_Ma + 1
NextRowLS_Ma = TitelRowLS_Ma + 1 'wordt later geherdefinieerd

KolLS_Ma_MatrNr = KolTO_MatrNr
KolLS_Ma_ArtTyp = KolTO_ArtTyp
KolLS_Ma_AbcTek = KolTO_AbcTek
KolLS_Ma_PlnSap = KolTO_PlnSap
KolLS_Ma_RelGrp = KolTO_RelGrp
KolLS_Ma_DafOms = KolTO_DafOms
KolLS_Ma_InkBes = KolTO_InkBes
KolLS_Ma_Produc = KolTO_Produc
KolLS_Ma_VeiVrd = KolTO_VeiVrd
KolLS_Ma_AfrWrd = KolTO_AfrWrd
KolLS_Ma_BesEnh = KolTO_BesEnh
KolLS_Ma_BasEnh = KolTO_BasEnh
KolLS_Ma_MinSer = KolTO_MinSer
KolLS_Ma_Levrtd = KolTO_Levrtd
KolLS_Ma_InkPrs = KolTO_InkPrs
KolLS_Ma_PrsPer = KolTO_PrsPer
KolLS_Ma_InkGrp = KolTO_InkGrp
KolLS_Ma_LevrNr = KolTO_LevrNr
KolLS_Ma_ArtNrL = KolTO_ArtNrL
KolLS_Ma_PCNcde = KolTO_PCNcde
KolLS_Ma_Locati = KolTO_Locati
KolLS_MA_Statnr = KolTO_Statnr 'toegevoegd voor statistieknummer'
KolLS_Ma_GewBru = KolTO_GewBru 'Toegevoegd voor statistieknummer'
KolLS_Ma_Gewich = KolTO_Gewich 'toegevoegd voor statistieknummer'
KolLS_Ma_GewEen = KOLTO_GewEen 'toegevoegd voor statistieknummer'

KolLS_St_MatrNr = "A"
KolLS_St_AbcTek = "W"
KolLS_St_PlnSap = "Z"
'KolLS_St_RelGrp wordt niet gebruikt
KolLS_St_DafOms = "D"
KolLS_St_InkBes = "BE"
KolLS_St_Produc = "U"
KolLS_St_VeiVrd = "AJ"
KolLS_St_AfrWrd = "AD"
'KolLS_St_BesEnh wordt niet gebruikt
KolLS_St_BasEnh = "E"
KolLS_St_MinSer = "AB"
KolLS_St_Levrtd = "AH"
KolLS_St_InkPrsS = "AT" 'Prijs bij prijssturing Standaard
KolLS_St_InkPrsV = "AU" 'Prijs bij prijssturing Voortschrijdend gemiddelde
KolLS_St_PrsPer = "AS"
KolLS_St_InkGrp = "L"
KolLS_St_LevrNr = "AV"
KolLS_St_PCNcde = "T"
KolLS_St_Locati = "AO"

KolLS_Tk_MatrNr = "A"
KolLS_Tk_DafOms = "C"

KolLS_Ib_MatrNr = "A"
KolLS_Ib_InkBes = "B"

KolLS_If_LevrNr = "A"
KolLS_If_MatrNr = "B"
KolLS_If_ArtNrL = "E"
KolLS_If_InkPrs = "H"
KolLS_If_PrsPer = "I"

KolLS_Rd_MatrNr = "A"
KolLS_Rd_InkPrsS = "H"
KolLS_Rd_PrsPer = "I"

KolLS_Ct_MatrNr = "B" 'toegevoegd voor statistieknummer'
KolLS_Ct_Statnr = "F" 'toegevoegd voor statistieknummer'
KolLS_Ct_GewBru = "C" 'toegevoegd voor statistieknummer'
KolLS_Ct_Gewich = "D" 'toegevoegd voor statistieknummer'
KolLS_Ct_GewEen = "E" 'toegevoegd voor statistieknummer'

KolLS_Pk_MatrNr = "A" 'toegevoegd voor plankenmerk'
KolLS_Pk_VeiVrd = "F" 'toegevoegd voor plankenmerk'

'----- 3.1. Code Controles -----

'Quotes omzetten
ActiveSheet.Unprotect
    Range("AP:AP").Select
    'Range("AP1").Activate
    Selection.Replace What:="""", Replacement:="”", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
ActiveSheet.Unprotect
    Range("AN:AN").Select
    Selection.Replace What:="""", Replacement:="”", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'Check of LSMW workbook al geopend is, eindig de macro als deze reeds geopend is
    On Error GoTo LSMW_voorbereiden
    If Windows(WbookLSMW).Visible = False Then Else
        MsgBox ("Het bestand " & WbookLSMW & " is reeds geopend. Sluit deze eerst.")
        Exit Sub
    
LSMW_voorbereiden:
'Aanmaken LSMW bestand & worksheets
    Application.DisplayAlerts = False
    Workbooks.Add.SaveAs Filename:="C:\Windows\Temp\" & WbookLSMW
    Workbooks(WbookLSMW).Sheets.Add.Move After:=Sheets("Blad3") 'aanmaak extra tabblad'
    Workbooks(WbookLSMW).Sheets.Add.Move After:=Sheets("Blad3") 'aanmaak extra tabblad'
    Workbooks(WbookLSMW).Sheets.Add.Move After:=Sheets("Blad3") 'aanmaak extra tabblad'
    Workbooks(WbookLSMW).Sheets.Add.Move After:=Sheets("Blad3") 'toegevoegd voor aanmaak extra tabblad statistiek'
    Workbooks(WbookLSMW).Sheets.Add.Move After:=Sheets("Blad3") 'toegevoegd voor aanmaak extra tabblad plankenmerk'
    Workbooks(WbookLSMW).Sheets(1).Name = SheetLS_Ma
    Workbooks(WbookLSMW).Sheets(2).Name = SheetLS_St
    Workbooks(WbookLSMW).Sheets(3).Name = SheetLS_Tk
    Workbooks(WbookLSMW).Sheets(4).Name = SheetLS_Ib
    Workbooks(WbookLSMW).Sheets(5).Name = SheetLS_If
    Workbooks(WbookLSMW).Sheets(6).Name = SheetLS_Rd
    Workbooks(WbookLSMW).Sheets(7).Name = SheetLS_Ct 'toegevoegd voor statistiek'
    Workbooks(WbookLSMW).Sheets(8).Name = SheetLS_Pk 'toegevoegd voor plankenmerk'
    
    Application.DisplayAlerts = True

    Set wsLS_Ma = Workbooks(WbookLSMW).Sheets(SheetLS_Ma)
    Set wsLS_St = Workbooks(WbookLSMW).Sheets(SheetLS_St)
    Set wsLS_Tk = Workbooks(WbookLSMW).Sheets(SheetLS_Tk)
    Set wsLS_Ib = Workbooks(WbookLSMW).Sheets(SheetLS_Ib)
    Set wsLS_If = Workbooks(WbookLSMW).Sheets(SheetLS_If)
    Set wsLS_Rd = Workbooks(WbookLSMW).Sheets(SheetLS_Rd)
    Set wsLS_Ct = Workbooks(WbookLSMW).Sheets(SheetLS_Ct) 'toegevoegd voor statistiek'
    Set wsLS_Pk = Workbooks(WbookLSMW).Sheets(SheetLS_Pk) 'toegevoegd voor plankenmerk'
    Set wsTO = Workbooks(AV_FileName).Sheets(SheetTO)

'----- 3.2 Code Activiteiten -----

'Master werkblad
    wsLS_Ma.Columns("A:BL").EntireColumn.NumberFormat = "@" 'Zorgt voor tekstformat van de cellen zodat voorloopnullen blijven staan'
    wsLS_Ma.Rows(TitelRowLS_Ma) = wsTO.Rows(TitelRowTO).Value
    For i = FirstRecRowTO To LastRowTO
        If (UCase(wsTO.Range(KolTO_Opgvrd & i))) = "IN PROGRESS" Then
            wsLS_Ma.Rows(NextRowLS_Ma) = wsTO.Rows(i).Value
            NextRowLS_Ma = NextRowLS_Ma + 1
        End If
    Next i
    wsLS_Ma.UsedRange.AutoFilter

    LastRowLS_Ma = wsLS_Ma.UsedRange.Rows.Count
    LastColumnLS_Ma = wsLS_Ma.UsedRange.Columns.Count
    RecordsLS_Ma = LastRowLS_Ma - FirstRecRowLS_Ma + 1
    NextRowLS_Ma = LastRowLS_Ma + 1

'Indien geen records exit
If RecordsLS_Ma = 0 Then
    MsgBox ("Geen records gevonden. U kunt het gemaakte bestand " & WbookLSMW & " sluiten.")
    Exit Sub
End If

' Stam
    wsLS_St.Columns("A:BL").EntireColumn.NumberFormat = "@"
    wsLS_St.Range("A1:BL1") = Array("MATNR", "MBRSH", "MTART", "MAKTX", "MEINS", "MATKL", "BISMT", "LABOR", "MTPOS", "WERKS", "BSTME", "EKGRP", "MMSTA", "MMSTD", "EKWSL", "WEBAZ", "INSMK", "KZKRI", "KORDB", "MFRPN", "MFRNR", "DISGR", "MAABC", "DISMM", "MINBE", "DISPO", "DISLS", "BSTMI", "BSTFE", "BSTRF", "BESKZ", "LGPRO", "LGFSB", "PLIFZ", "FHORI", "EISBE", "PERKZ", "MTVFP", "SBDKZ", "KZBED", "LGPBE", "LWMKB", "BKLAS", "VPRSV", "PEINH", "VERPR", "STPRS", "LIFNR", "IDNLF", "LGORT", "BUKRS", "WBKLA", "FXHOR", "ADDIT", "UMREN", "UMREZ", "INKTK", "STAWN", "NTGEW", "STOFF", "PROFL", "GEWEI", "BWTTY", "SPART")

    j = 2 '(1e record begint altijd op rij 2)
    For i = FirstRecRowLS_Ma To LastRowLS_Ma
        'Verschil tussen standaardwaarden artikelen NL / BE
        If Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "E" Then
            wsLS_St.Range("A" & j & ":AZ" & j) = Array(, "M", "ERSA", , , "PM_SP", , , , "NL01", , , , , "3", "0", , , "X", , , , , "PD", , , "EX", , , , "F", "0001", "0001", , "000", , "M", "02", , "T", , , "3040", "V", , , , , , "0001", "7002", "3040")
            ElseIf Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "W" Then
            wsLS_St.Range("A" & j & ":AZ" & j) = Array(, "M", "ERSA", , , "PM_SP", , , , "BE01", , , , , "3", "0", , , "X", , , , , "PD", , , "EX", , , , "F", "0001", "0001", , "000", , "M", "02", , "T", , , "3040", "V", , , , , , "0001", "7019", "3040")
            Else
            MsgBox ("Inkoopgroep onbekend voor artikelnummer " & wsLS_Ma.Range(KolLS_Ma_MatrNr & i).Value & ".")
            Exit Sub
        End If
        If wsLS_Ma.Range(KolLS_Ma_ArtTyp & i) = "Ruildeel" Then 'Ruildelen specifieke waarden (worden herkend adhv Type Artikel)
            wsLS_St.Range("BK" & j) = "C"
            wsLS_St.Range("BL" & j) = "RD"
        End If
        j = j + 1
    Next i
        
    'Kopieren van artikelspecifieke waarden
    wsLS_Ma.Range(KolLS_Ma_MatrNr & FirstRecRowLS_Ma & ":" & KolLS_Ma_MatrNr & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_MatrNr & "2")
    wsLS_Ma.Range(KolLS_Ma_DafOms & FirstRecRowLS_Ma & ":" & KolLS_Ma_DafOms & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_DafOms & "2")
    wsLS_Ma.Range(KolLS_Ma_BasEnh & FirstRecRowLS_Ma & ":" & KolLS_Ma_BasEnh & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_BasEnh & "2")
    wsLS_Ma.Range(KolLS_Ma_Produc & FirstRecRowLS_Ma & ":" & KolLS_Ma_Produc & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_Produc & "2")
    wsLS_Ma.Range(KolLS_Ma_AbcTek & FirstRecRowLS_Ma & ":" & KolLS_Ma_AbcTek & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_AbcTek & "2")
    wsLS_Ma.Range(KolLS_Ma_PlnSap & FirstRecRowLS_Ma & ":" & KolLS_Ma_PlnSap & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_PlnSap & "2")
    wsLS_Ma.Range(KolLS_Ma_MinSer & FirstRecRowLS_Ma & ":" & KolLS_Ma_MinSer & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_MinSer & "2")
    wsLS_Ma.Range(KolLS_Ma_AfrWrd & FirstRecRowLS_Ma & ":" & KolLS_Ma_AfrWrd & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_AfrWrd & "2")
    wsLS_Ma.Range(KolLS_Ma_Levrtd & FirstRecRowLS_Ma & ":" & KolLS_Ma_Levrtd & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_Levrtd & "2")
    wsLS_Ma.Range(KolLS_Ma_VeiVrd & FirstRecRowLS_Ma & ":" & KolLS_Ma_VeiVrd & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_VeiVrd & "2")
    wsLS_Ma.Range(KolLS_Ma_Locati & FirstRecRowLS_Ma & ":" & KolLS_Ma_Locati & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_Locati & "2")
    wsLS_Ma.Range(KolLS_Ma_PrsPer & FirstRecRowLS_Ma & ":" & KolLS_Ma_PrsPer & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_PrsPer & "2")
    wsLS_Ma.Range(KolLS_Ma_InkPrs & FirstRecRowLS_Ma & ":" & KolLS_Ma_InkPrs & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_InkPrsS & "2")
    wsLS_Ma.Range(KolLS_Ma_InkPrs & FirstRecRowLS_Ma & ":" & KolLS_Ma_InkPrs & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_InkPrsV & "2")
    wsLS_Ma.Range(KolLS_Ma_LevrNr & FirstRecRowLS_Ma & ":" & KolLS_Ma_LevrNr & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_LevrNr & "2")
    wsLS_Ma.Range(KolLS_Ma_InkBes & FirstRecRowLS_Ma & ":" & KolLS_Ma_InkBes & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_InkBes & "2")
    wsLS_Ma.Range(KolLS_Ma_InkGrp & FirstRecRowLS_Ma & ":" & KolLS_Ma_InkGrp & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_InkGrp & "2")
    wsLS_Ma.Range(KolLS_Ma_PCNcde & FirstRecRowLS_Ma & ":" & KolLS_Ma_PCNcde & LastRowLS_Ma).Copy Destination:=wsLS_St.Range(KolLS_St_PCNcde & "2")
    'KolLS_Ma_RelGrp = KolTO_RelGrp nog erbij?
    'KolLS_Ma_BesEnh = KolTO_BesEnh nog erbij?
    'KolLS_Ma_PCNcde = KolTO_PCNcde nog erbij?
    'wsLS_St.Range("A2:A" & (AantalRijen)) = wsLS_Ma.Range("AB2:AB" & (AantalRijen)).Value 'Materiaalnummer

' TKT EN-NL
    wsLS_Tk.Columns("A:C").EntireColumn.NumberFormat = "@"
    wsLS_Tk.Range("A1:C1") = Array("MATNR", "SPRAS", "MAKTX")
    wsLS_Tk.Range("A2:B" & (RecordsLS_Ma + 1)) = Array(, "EN")

    wsLS_Ma.Range(KolLS_Ma_MatrNr & FirstRecRowLS_Ma & ":" & KolLS_Ma_MatrNr & LastRowLS_Ma).Copy Destination:=wsLS_Tk.Range(KolLS_Tk_MatrNr & "2")
    wsLS_Ma.Range(KolLS_Ma_DafOms & FirstRecRowLS_Ma & ":" & KolLS_Ma_DafOms & LastRowLS_Ma).Copy Destination:=wsLS_Tk.Range(KolLS_Tk_DafOms & "2")

' Ink.best.tkt
 
    wsLS_Ib.Columns("A:C").EntireColumn.NumberFormat = "@"
    wsLS_Ib.Range("A1:C1") = Array("MATNR", "INKTK", "WERKS")
    
     j = 2 '(1e record begint altijd op rij 2)
    For i = FirstRecRowLS_Ma To LastRowLS_Ma
        'Verschil tussen standaardwaarden artikelen NL / BE
        If Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "E" Then
            wsLS_Ib.Range("A" & j & ":C" & j) = Array(, , "NL01")
            ElseIf Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "W" Then
            wsLS_Ib.Range("A" & j & ":C" & j) = Array(, , "BE01")
        End If
        j = j + 1
    Next i

    wsLS_Ma.Range(KolLS_Ma_MatrNr & FirstRecRowLS_Ma & ":" & KolLS_Ma_MatrNr & LastRowLS_Ma).Copy Destination:=wsLS_Ib.Range(KolLS_Ib_MatrNr & "2")
    wsLS_Ma.Range(KolLS_Ma_InkBes & FirstRecRowLS_Ma & ":" & KolLS_Ma_InkBes & LastRowLS_Ma).Copy Destination:=wsLS_Ib.Range(KolLS_Ib_InkBes & "2")

'Inforecord
    wsLS_If.Columns("A:K").EntireColumn.NumberFormat = "@"
    wsLS_If.Range("A1:K1") = Array("LIFNR", "MATNR", "EKORG", "WERKS", "IDNLF", "NORBM", "MINBM", "NETPR", "PEINH", "LTEX1", "LTEX2")

    j = 2 '(1e record begint altijd op rij 2)
    For i = FirstRecRowLS_Ma To LastRowLS_Ma
        'Verschil tussen standaardwaarden artikelen NL / BE
        If Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "E" Then
            wsLS_If.Range("A" & j & ":G" & j) = Array(, , "NL01", "NL01", , "1", "1")
            ElseIf Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "W" Then
            wsLS_If.Range("A" & j & ":G" & j) = Array(, , "BE01", "BE01", , "1", "1")
        End If
        j = j + 1
    Next i

    wsLS_Ma.Range(KolLS_Ma_LevrNr & FirstRecRowLS_Ma & ":" & KolLS_Ma_LevrNr & LastRowLS_Ma).Copy Destination:=wsLS_If.Range(KolLS_If_LevrNr & "2")
    wsLS_Ma.Range(KolLS_Ma_MatrNr & FirstRecRowLS_Ma & ":" & KolLS_Ma_MatrNr & LastRowLS_Ma).Copy Destination:=wsLS_If.Range(KolLS_If_MatrNr & "2")
    wsLS_Ma.Range(KolLS_Ma_ArtNrL & FirstRecRowLS_Ma & ":" & KolLS_Ma_ArtNrL & LastRowLS_Ma).Copy Destination:=wsLS_If.Range(KolLS_If_ArtNrL & "2")
    wsLS_Ma.Range(KolLS_Ma_InkPrs & FirstRecRowLS_Ma & ":" & KolLS_Ma_InkPrs & LastRowLS_Ma).Copy Destination:=wsLS_If.Range(KolLS_If_InkPrs & "2")
    wsLS_Ma.Range(KolLS_Ma_PrsPer & FirstRecRowLS_Ma & ":" & KolLS_Ma_PrsPer & LastRowLS_Ma).Copy Destination:=wsLS_If.Range(KolLS_If_PrsPer & "2")

'Reparatiedelen
    wsLS_Rd.Columns("A:J").EntireColumn.NumberFormat = "@"
    wsLS_Rd.Range("A1:J1") = Array("MATNR", "MTART", "WERKS", "BWTAR-NEW", "BWTAR-REF", "VPRSV", "VERPR", "STPRS", "PEINH", "BKLAS")
    
    j = 2 '(1e record begint altijd op rij 2)
    For i = FirstRecRowLS_Ma To LastRowLS_Ma
        If wsLS_Ma.Range(KolLS_Ma_ArtTyp & i) = "Ruildeel" Then
            'Verschil tussen standaardwaarden artikelen NL / BE
            If Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "E" Then
               wsLS_Rd.Range("A" & j & ":J" & j) = Array(, "ERSA", "NL01", "NIEUW", , "V", "0,01", , , "3040")
               ElseIf Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "W" Then
               wsLS_Rd.Range("A" & j & ":J" & j) = Array(, "ERSA", "BE01", "NIEUW", , "V", "0,01", , , "3040")
               Else
            End If
        wsLS_Ma.Range(KolLS_Ma_MatrNr & i).Copy Destination:=wsLS_Rd.Range(KolLS_Rd_MatrNr & j)
        wsLS_Ma.Range(KolLS_Ma_InkPrs & i).Copy Destination:=wsLS_Rd.Range(KolLS_Rd_InkPrsS & j)
        wsLS_Ma.Range(KolLS_Ma_PrsPer & i).Copy Destination:=wsLS_Rd.Range(KolLS_Rd_PrsPer & j)
        j = j + 1
        End If
    Next i
    
'Reparatiedelen V1 besturing
    wsLS_Pk.Columns("A:G").EntireColumn.NumberFormat = "@"
    wsLS_Pk.Range("A1:G1") = Array("MATNR", "WERKS", "LGORT", "EISBE", "DISMM", "MINBE", "BKLAS")
    
    j = 2 '(1e record begint altijd op rij 2)
    For i = FirstRecRowLS_Ma To LastRowLS_Ma
        If wsLS_Ma.Range(KolLS_Ma_ArtTyp & i) = "Ruildeel" And wsLS_Ma.Range(KolLS_Ma_VeiVrd & i) > 0 Then
            'Verschil tussen standaardwaarden artikelen NL / BE
            If Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "E" Then
               wsLS_Pk.Range("A" & j & ":G" & j) = Array(, "NL01", "0001", "0", "V1", , "3040")
               ElseIf Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "W" Then
               wsLS_Pk.Range("A" & j & ":G" & j) = Array(, "BE01", , "0", "V1", , "3040")
               Else
            End If
        wsLS_Ma.Range(KolLS_Ma_MatrNr & i).Copy Destination:=wsLS_Pk.Range(KolLS_Pk_MatrNr & j)
        wsLS_Ma.Range(KolLS_Ma_VeiVrd & i).Copy Destination:=wsLS_Pk.Range(KolLS_Pk_VeiVrd & j)
        j = j + 1
        End If
    Next i
    
'Statistieknummer (Toegevoegd voor Statistieknummer)
    wsLS_Ct.Columns("A:F").EntireColumn.NumberFormat = "@"
    wsLS_Ct.Range("A1:F1") = Array("WERKS", "MATNR", "BRGEW", "NTGEW", "GEWEI", "STAWN")

    j = 2 '(1e record begint altijd op rij 2)
    For i = FirstRecRowLS_Ma To LastRowLS_Ma
        'Verschil tussen standaardwaarden artikelen NL / BE
        If Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "E" Then
            wsLS_Ct.Range("A" & j & ":F" & j) = Array("NL01", , , , , "1")
            ElseIf Left(wsLS_Ma.Range(KolLS_Ma_InkGrp & i), 1) = "W" Then
            wsLS_Ct.Range("A" & j & ":F" & j) = Array("BE01", , , , , "1")
            Else
            MsgBox ("Inkoopgroep onbekend voor artikelnummer " & wsLS_Ct.Range(KolLS_Ct_MatrNr & i).Value & ".")
            Exit Sub
        End If
        
        j = j + 1
    Next i
        
    'Kopieren van artikelspecifieke waarden
    wsLS_Ma.Range(KolLS_Ma_MatrNr & FirstRecRowLS_Ma & ":" & KolLS_Ma_MatrNr & LastRowLS_Ma).Copy Destination:=wsLS_Ct.Range(KolLS_Ct_MatrNr & "2")
    wsLS_Ma.Range(KolLS_Ma_GewBru & FirstRecRowLS_Ma & ":" & KolLS_Ma_GewBru & LastRowLS_Ma).Copy Destination:=wsLS_Ct.Range(KolLS_Ct_GewBru & "2")
    wsLS_Ma.Range(KolLS_Ma_Gewich & FirstRecRowLS_Ma & ":" & KolLS_Ma_Gewich & LastRowLS_Ma).Copy Destination:=wsLS_Ct.Range(KolLS_Ct_Gewich & "2")
    wsLS_Ma.Range(KolLS_Ma_GewEen & FirstRecRowLS_Ma & ":" & KolLS_Ma_GewEen & LastRowLS_Ma).Copy Destination:=wsLS_Ct.Range(KolLS_Ct_GewEen & "2")
    wsLS_Ma.Range(KolLS_MA_Statnr & FirstRecRowLS_Ma & ":" & KolLS_MA_Statnr & LastRowLS_Ma).Copy Destination:=wsLS_Ct.Range(KolLS_Ct_Statnr & "2")

'Formatting
Sheets(2).UsedRange.EntireColumn.AutoFit
Sheets(3).UsedRange.EntireColumn.AutoFit
Sheets(4).UsedRange.EntireColumn.AutoFit
Sheets(5).UsedRange.EntireColumn.AutoFit
Sheets(6).UsedRange.EntireColumn.AutoFit
Sheets(7).UsedRange.EntireColumn.AutoFit 'toegevoegd voor statistieknummer'
Sheets(8).UsedRange.EntireColumn.AutoFit 'toegevoegd voor plankenmerk'

'
wsLS_Ma.Activate
MsgBox ("Tijdelijk bestand aangemaakt voor upload SAP. Aantal artikels: " & RecordsLS_Ma & ".")

'Opslaan als CSV
Load UserForm1
UserForm1.Show

Application.DisplayAlerts = False
Application.ScreenUpdating = False

For k = 2 To 8 '6 vervangen door 7 voor statistieknummer csvbestand en 7 door 8 voor plankenmerk'
    Workbooks(WbookLSMW).Sheets(k).Copy
    ActiveWorkbook.SaveAs Filename:=CSVmap & "\" & ActiveSheet.Name & ".csv", FileFormat:=xlCSV, local:=True
    ActiveWindow.Close
Next k

Application.DisplayAlerts = True
Application.ScreenUpdating = True

MsgBox ("CSV bestanden opgeslagen in: " & CSVmap & ".")

End Sub

