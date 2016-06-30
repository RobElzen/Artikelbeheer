Attribute VB_Name = "M57_DAF_Mailing"
''Option Explicit


''===============================||====================================
Sub Mailing_ACC_link()
''===============================||====================================
    Call Apply_UserNames
''===============================||====================================
    Dim rCell As Variant
    Dim clmRange As Range
    Dim mailRange As Range
    Dim mailingRange As Range
    Dim mailonlyRange As Range
    Dim shtMailAdres As String
    Dim I                       ''First cell - user name

'''''''''''''''Mail Users defined in sheet SETTINGS: SINGLE or PLURAL
'mailing:
'              If mailRange(rCell.Row - 1, 1).Value = "Y" Then
''Mailing in Office 2000-2010
    Dim Email_Subject, Email_Send_From, Email_Send_To, Email_Cc, Email_Bcc, Email_Body As String
    Dim Mail_Object, Mail_Single As Variant
    Dim Email_Signature, Email_Note As String
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    
    Dim strbody As String
''===============================||====================================
'SpeedUp the macro
Application.DisplayAlerts = False    'replacing
Application.EnableEvents = False
Application.ScreenUpdating = False
''===============================||====================================
Set Network = CreateObject("wscript.network")
''===============================||====================================
'Create an Outlook object and new mail message
Email_Send_From = Email                                     '"dragan.straleger@daftrucks.com"
Email_Send_To = USERMAIL                                    'UserEmail_var ''shtMailAdres
Email_Cc = "" ' "dragan.straleger@paccar.com" ' ""
Email_Bcc = ""
                                                                                                                 ''         Now, "dd-mm-yyyy   hh:mm:ss"
Email_Subject = "Accorderingsverzoek Databeheer: " & _
                " '' " & UserGroup & "  ''   d.d. " & Format(Now, "dd mmm yyyy   hh:nn") & "    "                ''  Format(Now, "yyyy-mm-dd hh:nn:ss)

'Email_Body = "Hallo"                                                       ''Data-DS220679.xlsm#'Databestand'!D6       Artikelbeheer.xlsm#Sheet4!BY5
''''''''''''''
'If ActiveSheet = "Accordering" Then
If UserGroup = "MMO" Then
strbody = "Hallo," & "<br>" & "<br>" & _
          "Bij deze ontvangt u link naar allerlaatste rapportage betreffende uw accordering." & "<br>" & _
          "Gaarne deze accordering verwerken." & "<br>" & "<br>" & _
          "Click on this link to open the file : " & _
          "<A HREF=""http://dafshare-org.eu.paccar.com/organization/ops-mtc/Artikelaanvraag/Artikelbeheer.xlsm#'Accordering'!BY5" & _
          """>Link to the file</A>" & _
          "<br>" & "<br>" & _
          "ps. Voor eventuele vragen kunt u zonder afspraak even langs komen." & "<br>" & _
          "    Dit is een automatisch gegenereerd bericht."
Else
strbody = "Hallo," & "<br>" & "<br>" & _
          "Bij deze ontvangt u link naar allerlaatste rapportage betreffende uw accordering." & "<br>" & _
          "Gaarne deze accordering verwerken." & "<br>" & "<br>" & _
          "Click on this link to open the file : " & _
          "<A HREF=""http://dafshare-org.eu.paccar.com/organization/ops-mtc/Artikelaanvraag/Artikelbeheer.xlsm#'Accordering'!BY5" & _
          """>Link to the file</A>" & "<br>" & "<br>" & _
          "ps. Voor eventuele vragen kunt u zonder afspraak even langs komen." & "<br>" & _
          "    Dit is een automatisch gegenereerd bericht."
End If
''''''''''''''''Network.UserName  +  Telefoonnummer TOEVOEGEN als variable
Email_Signature = "Met vriendelijke groeten," & "<br>" & _
                Naam & "<br>" & _
                "_____________________________" & "<br>" & _
                "Maintenance Material Planning" & "<br>" & _
                "Onderhoudsondersteuning      " & "<br>" & _
                "_____________________________" & "<br>" & _
                "Interne Postcode: E55.00.395 "

Email_Note = "CONFIDENTIALITY NOTICE: " & "This e-mail and any attachments," & _
"files or previous e-mail messages attached to it may contain confidential information " & _
"and is for the intended recipient(s) only. Confidential information should not be disclosed," & _
"copied, distributed or used without the permission of the sender or PACCAR." & _
"If you are not the intended recipient, please notify me by reply e-mail and destroy the original transmission."
''===============================||====================================
''===============================||====================================
''===============================||====================================
    On Error GoTo debugs
    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)
''===============================||====================================
    With Mail_Single
        .To = Email_Send_To
        .cc = Email_Cc
        .BCC = Email_Bcc
        .Subject = Email_Subject
        
       '.BodyFormat = olFormatHTML
        .HTMLBody = strbody & "<br>" & "<br>" & "<br>" & _
                    Email_Signature & "<br>" & "<br>" & "<br>" & Email_Note & "<br>" & "<br>" & "<br>" & RangetoHTML(rngHTML)
        'HTMLBody accepteert geen vbNewLine maar "<br>"
       '.Body = Email_Body & vbNewLine & vbNewLine & Email_Signature & vbNewLine & vbNewLine & vbNewLine & vbNewLine & Email_Note & vbNewLine & vbNewLine
        .Importance = 2                     ''(0 = Low, 2 = High, 1 = Normal)
        .ReadReceiptRequested = True
        .Attachments.Add Path_C & filename    ''("C:\Text.txt")
        .SentOnBehalfOfName = """Dragan Straleger"" <dragan.straleger@daftrucks.com>" '<dstraleger@home.nl>" 'Change sender name and reply address
        ''Stay in the outbox untill this date and time
        .DeferredDeliveryTime = DateAdd("h", 1000, Now) ''h  ''n  ''s
       '.DeferredDeliveryTime = Now + 3 ''3 uur "1/1/2013 10:40:00 AM"
        'Uitstel aanmaak email (bij elke loop - nieuwe mailbericht)
        'Application.Wait (Now + TimeValue("00:01:05"))
        'Workbooks(filename).Close
        .send
       '.Display    'SLECHTS WEERGAVE, ZONDER OPSLAAN WORDT HET NIET VERZONDEN
    End With
debugs:
If Err.Description <> "" Then MsgBox Err.Description
''===============================||====================================
'SpeedUp the macro
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
''===============================||====================================
''<><><><><><><><><><><><>
      Set Mail_Single = Nothing
      Set Mail_Object = Nothing
''===============================||====================================
'              Else
'              End If
'          Else
'          End If
'      Else
'      End If
'      rCell = (rCell.Row + 1)
'Next rCell
MAILING_END:
'End If
'Next

End Sub


''===============================||====================================
Sub Mail_All_Sheets()
    Dim rCell As Variant
    Dim clmRange As Range
    Dim mailRange As Range
    Dim mailingRange As Range
    Dim mailonlyRange As Range
    Dim shtMailAdres As String
    Dim I                       ''First cell - user name

'  For i = FirstUserColumn To Worksheets("SETTINGS").Cells(1, 1).SpecialCells(xlLastCell).Column
  For I = 10 To Worksheets("SETTINGS").Cells(1, 1).SpecialCells(xlLastCell).Column
  
  If Worksheets("SETTINGS").Cells(1, I) <> 0 And Worksheets("SETTINGS").Cells(1, I) <> "" Then

  UserGroup = Cells(1, I).Value
  Set clmRange = Range("SET." & UserGroup)
'Define mail range
  Set mailingRange = Range("SET.Mailing")
  Set mailonlyRange = Range("SET.MailOnly")
'      If Mail_ALL_Users = False Then
'         Set mailRange = mailonlyRange
'         Else
'         Set mailRange = mailingRange
'      End If

MAILING_START:
For Each rCell In Range("SET.RANGE_ALL").Cells
      If rCell.Value <> "" Then
''''''
          If clmRange(rCell.Row - 1, 1).Value = "X" Or clmRange(rCell.Row - 1, 1).Value <> "" Then
             shtMailAdres = Range("SET.RANGE_ALL").Range("A" & rCell.Row - 1) ''rcell.Row - 1)
          GoTo mailing
''''''''''''''Mail Users defined in sheet SETTINGS: SINGLE or PLURAL
mailing:
              If mailRange(rCell.Row - 1, 1).Value = "Y" Then
''Mailing in Office 2000-2010
    Dim Email_Subject, Email_Send_From, Email_Send_To, Email_Cc, Email_Bcc, Email_Body As String
    Dim Mail_Object, Mail_Single As Variant
    Dim Email_Signature, Email_Note As String
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
''===============================||====================================
'SpeedUp the macro
Application.EnableEvents = False
Application.ScreenUpdating = False
''===============================||====================================
'Create an Outlook object and new mail message
Email_Send_From = "dragan.straleger@daftrucks.com"          '"dstraleger@gmail.com" '
Email_Send_To = shtMailAdres
Email_Cc = "" ' "dragan.straleger@paccar.com" ' ""
Email_Bcc = ""
                                                                                                                 ''         Now, "dd-mm-yyyy   hh:mm:ss"
Email_Subject = "Newsletter TD Rayon Centraal: " & _
                "RAPPORT_" & UserGroup & " d.d. " & Format(Now, "dd mmm yyyy   hh:nn") & "    " & shtMailAdres   ''  Format(Now, "yyyy-mm-dd hh:nn:ss)
Email_Body = "Hallo," & vbNewLine & vbNewLine & _
              "Bij deze ontvangt u, als bijlage, allerlaatste rapportage betreffende uw werkzaamheden." & vbNewLine & _
              "Gaarne deze lijst(en) doornemen." & vbNewLine & vbNewLine & _
              "ps. Voor eventuele vragen kunt u zonder afspraak even langs komen." & vbNewLine
Email_Signature = "Met vriendelijke groeten," & vbNewLine & vbNewLine & _
                "Dragan Straleger" & vbNewLine & _
                UserName & UserName & UserName & _
                "Maintenance Engineering & Planning" & vbNewLine & _
                "Rayon Centraal: Tools & Utility Services" & vbNewLine & _
                "_________________________" & vbNewLine & _
                "Interne Postcode: E55.00.395" & vbNewLine & _
                "Telefoon: +31 (0)40 214 3226" & vbNewLine

Email_Note = "CONFIDENTIALITY NOTICE: " & "This e-mail and any attachments," & _
"files or previous e-mail messages attached to it may contain confidential information " & _
"and is for the intended recipient(s) only. Confidential information should not be disclosed," & _
"copied, distributed or used without the permission of the sender or PACCAR." & _
"If you are not the intended recipient, please notify me by reply e-mail and destroy the original transmission."
''===============================||====================================
    On Error GoTo debugs
    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)
''===============================||====================================
   'Makes a copy of the active sheet and save it to a temporary file
    Dim filename, WB '' Path
    
    Worksheets("RAPPORT_" & UserGroup).COPY
    Set WB = Worksheets("RAPPORT_" & UserGroup)
    filename = "RAPPORT_" & UserGroup & " " & Format(Now, "dd mmm yyyy hh.nn") & "     " & shtMailAdres & ".xls"
   'Path = "C:\Rapport_Out\"
    WB.SaveAs Path & filename
  ''MsgBox "Look in <" & Path & "> ... for the files!" & vbNewLine & "Look if map <" & Path & "> is present!"
    
    With Mail_Single
        .To = Email_Send_To
        .cc = Email_Cc
        .BCC = Email_Bcc
        .Subject = Email_Subject
        .Body = Email_Body & vbNewLine & vbNewLine & Email_Signature & vbNewLine & vbNewLine & vbNewLine & vbNewLine & Email_Note & vbNewLine & vbNewLine
        .Importance = 2                     ''(0 = Low, 2 = High, 1 = Normal)
        .ReadReceiptRequested = True
        .Attachments.Add Path & filename    ''("C:\Text.txt")
        .SentOnBehalfOfName = """Dragan Straleger"" <dragan.straleger@daftrucks.com>" '<dstraleger@home.nl>" 'Change sender name and reply address
        ''Stay in the outbox untill this date and time
        .DeferredDeliveryTime = DateAdd("s", 10, Now) ''h  ''n  ''s
        ''.DeferredDeliveryTime = Now + 3 ''3 uur "1/1/2013 10:40:00 AM"
        ''Uitstel aanmaak email (bij elke loop - nieuwe mailbericht)
        ''Application.Wait (Now + TimeValue("00:01:05"))
         Workbooks(filename).Close
        .send
       '.Display    'SLECHTS WEERGAVE, ZONDER OPSLAAN WORDT HET NIET VERZONDEN
    End With
debugs:
If Err.Description <> "" Then MsgBox Err.Description
''===============================||====================================
'SpeedUp the macro
Application.EnableEvents = True
Application.ScreenUpdating = True
''===============================||====================================
''<><><><><><><><><><><><>
      Set Mail_Single = Nothing
      Set Mail_Object = Nothing
''===============================||====================================
              Else
              End If
          Else
          End If
      Else
      End If
      rCell = (rCell.Row + 1)
Next rCell
MAILING_END:
End If
Next

End Sub
