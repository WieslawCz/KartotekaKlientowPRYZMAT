Imports System.Drawing.Printing
Imports System.Data.OleDb
Imports System
Imports System.Reflection ' For Missing.Value and BindingFlags
Imports System.Runtime.InteropServices ' For COMException
'------------------------------------------
'UWAGA :
'Do referencji trzeba DODAC
'   Microsoft Excel 11.0 Object Library
'   Microsoft Office 11.0 Object Library
'-------------------------------------------
'Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Outlook

Module modOutlook

    '****************************************************
    ' moduł integracji z MS Outlook
    ' 1. Dpdać do Referencji Microsoft.Office.Interop.Outlook
    ' 2. zdefiniować obiekty oOutlook i oMessage]
    '****************************************************

    Dim oOutlook As Microsoft.Office.Interop.Outlook.Application
    Dim oMessage As Microsoft.Office.Interop.Outlook.MailItem

    Dim UsunOutllokZPamieci As Boolean = False

    Public Sub SendByOutlook(ByVal pAdresOdbiorcy As String, _
                             ByVal pTemat As String, _
                             ByVal pTresc As String, _
                             ByVal pZalacznik As String, _
                             ByVal pKomenda As String)

        Try
            'czy Outlook jest uruchomiony
            oOutlook = CType(GetObject(, "Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
            UsunOutllokZPamieci = False
        Catch ex As System.Exception
            If oOutlook Is Nothing Then
                'nie jest uruchomiony - uruchom...
                oOutlook = New Microsoft.Office.Interop.Outlook.Application
                UsunOutllokZPamieci = True
            End If
        End Try

        If oOutlook Is Nothing Then
            'nuie udało się uruchomic - nie jest zainstalowany
            MessageBox.Show("Program Outlook nie jest zainstalowany na tym komputerze.", "Uwaga :", MessageBoxButtons.OK)
        Else
            '----------------------------
            '' Create an Outlook application.
            ' ''Me.Cursor = Cursors.WaitCursor
            'Dim oApp As Microsoft.Office.Interop.Outlook._Application
            'oApp = New Microsoft.Office.Interop.Outlook.Application()

            ' Create a new MailItem.
            Dim oMsg As Microsoft.Office.Interop.Outlook._MailItem
            oMsg = oOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)

            oMsg.Subject = pTemat       '"Send Attachment Using OOM in Visual Basic .NET"
            oMsg.Body = pTresc          '"Hello World" & vbCr & vbCr
            oMsg.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain
            oMsg.To = pAdresOdbiorcy    '"user@example.com"
            oMsg.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal

            ' Add an attachment
            ' TODO: Replace with a valid attachment path.
            If Len(pZalacznik) > 0 Then
                'Dim sSource As String = "C:\Temp\Hello.txt"
                '' TODO: Replace with attachment name
                'Dim sDisplayName As String = "Hello.txt"

                'Dim sBodyLen As String = oMsg.Body.Length
                'Dim oAttachs As Microsoft.Office.Interop.Outlook.Attachments = oMsg.Attachments
                'Dim oAttach As Microsoft.Office.Interop.Outlook.Attachment
                'oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)
                oMsg.Attachments.Add(pZalacznik, _
                                Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, _
                                oMsg.Body.Length + 1, _
                                IO.Path.GetFileName(pZalacznik))
            End If

            ' Send
            If pKomenda = "Wyslij" Then
                oMsg.Send()
            Else
                oMsg.Display(False) 'True is modal
            End If

            ' Clean up
            oOutlook = Nothing
            oMsg = Nothing
            'Me.Cursor = Cursors.Default
            '----------------------------



            ''-------------------
            ''Me.Cursor = Cursors.WaitCursor
            'oMessage = oOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
            'With oMessage
            '    .To = pAdresOdbiorcy
            '    .Subject = pTemat
            '    .BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain
            '    .Body = pTresc
            '    .Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal

            '    If Len(pZalacznik) > 0 Then
            '        .Attachments.Add(pZalacznik, _
            '                        Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, _
            '                        1, _
            '                        IO.Path.GetFileName(pZalacznik))
            '    End If


            '    If pKomenda = "Wyslij" Then
            '        .Send()
            '    Else
            '        .Display(False) 'True is modal
            '    End If
            'End With
            ''----------------------------



            '----------------------------
            'Me.Cursor = Cursors.Default
            If UsunOutllokZPamieci = True Then
                If Not oOutlook Is Nothing Then
                    oOutlook.Quit()
                    oOutlook = Nothing
                End If
            End If
            '----------------------------
            oOutlook = Nothing
        End If
    End Sub


    '*************************
    'wysyłąnie emaili automatycznie....
    '***********************************


    'Private Sub SendEmailtoContacts()
    '    Dim subjectEmail As String = "Meeting has been rescheduled."
    '    Dim bodyEmail As String = "Meeting is one hour later."
    '    Dim sentContacts As Microsoft.Office.Interop.Outlook.MAPIFolder = Me.ActiveExplorer() _
    '        .Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook _
    '        .OlDefaultFolders.olFolderContacts)
    '    For Each contact As Microsoft.Office.Interop.Outlook.ContactItem In sentContacts.Items()
    '        If contact.Email1Address.Contains("example.com") Then
    '            CreateEmailItem(subjectEmail, contact _
    '            .Email1Address, bodyEmail)
    '        End If
    '    Next
    'End Sub

    'Private Sub CreateEmailItem(ByVal subjectEmail As String, _
    '    ByVal toEmail As String, ByVal bodyEmail As String)
    '    Dim eMail As Microsoft.Office.Interop.Outlook.MailItem = Me.CreateItem _
    '        (Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
    '    With eMail
    '        .Subject = subjectEmail
    '        .To = toEmail
    '        .Body = bodyEmail
    '        .Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceLow
    '        .Send()
    '    End With
    'End Sub


    '********************
    'czyta nieodczytane waidomosci
    '************************

    'Public Sub ReadByOutlook()
    '    Dim inbox As Microsoft.Office.Interop.Outlook.MAPIFolder = _
    '        Me.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox)

    '    Dim unreadItems As Microsoft.Office.Interop.Outlook.Items = inbox.Items.Restrict("[Unread]=true")

    '    MessageBox.Show(String.Format("Unread items in Inbox = {0}", unreadItems.Count))
    'End Sub


End Module
