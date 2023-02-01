Module modOutlook

    '****************************************************
    ' moduł integracji z MS Outlook
    ' 1. Dpdać do Referencji Microsoft.Office.Interop.Outlook
    ' 2. zdefiniować obiekty oOutlook i oMessage]
    '****************************************************

    Dim oOutlook As Microsoft.Office.Interop.Outlook.Application
    Dim oMessage As Microsoft.Office.Interop.Outlook.MailItem

    Dim UsunOutlookZPamieci As Boolean = False

    Public Sub SendByOutlook(ByVal pAdresOdbiorcy As String, _
                             ByVal pTemat As String, _
                             ByVal pTresc As String, _
                             ByVal pTrescPic As String, _
                             ByVal pZalacznik As String, _
                             ByVal pKomenda As String)
        SendByOutlook(pAdresOdbiorcy, pTemat, pTresc, pTrescPic, pZalacznik, "", "", pKomenda)
    End Sub




    Public Sub SendByOutlook(ByVal pAdresOdbiorcy As String, _
                             ByVal pTemat As String, _
                             ByVal pTresc As String, _
                             ByVal pTrescPic As String, _
                             ByVal pZalacznik As String, _
                             ByVal pDW As String, _
                             ByVal pUDW As String, _
                             ByVal pKomenda As String)
        Try
            'czy Outlook jest uruchomiony
            oOutlook = CType(GetObject(, "Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
            UsunOutlookZPamieci = False
        Catch ex As System.Exception
            If oOutlook Is Nothing Then
                'nie jest uruchomiony - uruchom...
                oOutlook = New Microsoft.Office.Interop.Outlook.Application
                UsunOutlookZPamieci = True
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
            oMsg.To = pAdresOdbiorcy    '"user@example.com"
            oMsg.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal
            oMsg.BCC = pUDW
            oMsg.CC = pDW

            If Len(pTrescPic) = 0 Then
                oMsg.Body = pTresc          '"Hello World" & vbCr & vbCr
                oMsg.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain
            Else
                oMsg.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML
                ' Add image attachment from local disk
                Dim oAttachment As Microsoft.Office.Interop.Outlook.Attachment = oMsg.Attachments.Add(pTrescPic)
                ' Specifies the attachment as an embedded picture
                ' contentid can be any string.
                Dim contentID As String = IO.Path.GetFileName(pTrescPic)
                'oAttachment.DisplayName = contentID
                'oMsg.HTMLBody = "<html><body>this is a <img src=""cid:" _
                '         + contentID + """> embedded picture.</body></html>"
                oMsg.HTMLBody = "<html><body><img src=""cid:" _
                         + contentID + """></body></html>"
            End If

            ' Add an attachment
            ' TODO: Replace with a valid attachment path.
            If Len(pZalacznik) > 0 Then
                Dim PlikZal As String = Dir(pZalacznik)
                Do While Len(PlikZal) > 0
                    oMsg.Attachments.Add(IO.Path.GetDirectoryName(pZalacznik) & "\" & PlikZal, _
                                    Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, _
                                    1, _
                                    PlikZal)
                    PlikZal = Dir()
                Loop
                'Dim sSource As String = "C:\Temp\Hello.txt"
                '' TODO: Replace with attachment name
                'Dim sDisplayName As String = "Hello.txt"

                'Dim sBodyLen As String = oMsg.Body.Length
                'Dim oAttachs As Microsoft.Office.Interop.Outlook.Attachments = oMsg.Attachments
                'Dim oAttach As Microsoft.Office.Interop.Outlook.Attachment
                'oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)

                'oMsg.Attachments.Add(pZalacznik, _
                '                Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, _
                '                oMsg.Body.Length + 1, _
                '                IO.Path.GetFileName(pZalacznik))
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
            If UsunOutlookZPamieci = True Then
                If Not oOutlook Is Nothing Then
                    oOutlook.Quit()
                    oOutlook = Nothing
                End If
            End If
            '----------------------------
            oOutlook = Nothing
        End If
    End Sub





    'Public Sub SendByOutlook(ByVal pAdresOdbiorcy As String, _
    '                         ByVal pTemat As String, _
    '                         ByVal pTresc As String, _
    '                         ByVal pZalacznik As String)

    '    Dim PlikZalacznik As String = ""
    '    Dim PathZalacznik As String = ""
    '    Dim LicznikZalacznikow As Integer = 0

    '    Try
    '        'czy Outlook jest uruchomiony
    '        oOutlook = CType(GetObject(, "Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
    '        UsunOutllokZPamieci = False
    '    Catch ex As Exception
    '        If oOutlook Is Nothing Then
    '            'nie jest uruchomiony - uruchom...
    '            oOutlook = New Microsoft.Office.Interop.Outlook.Application
    '            UsunOutllokZPamieci = True
    '        End If
    '    End Try

    '    If oOutlook Is Nothing Then
    '        'nuie udało się uruchomic - nie jest zainstalowany
    '        MessageBox.Show("Program Outlook nie jest zainstalowany na tym komputerze.", "Uwaga :", MessageBoxButtons.OK)
    '    Else

    '        'Me.Cursor = Cursors.WaitCursor
    '        oMessage = oOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
    '        With oMessage
    '            .To = pAdresOdbiorcy
    '            .Subject = pTemat
    '            .BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain
    '            .Body = pTresc
    '            .Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal
    '            If Len(pZalacznik) > 0 Then
    '                ' pobierz wszystkie załączniki
    '                LicznikZalacznikow = 0
    '                PathZalacznik = IO.Path.GetDirectoryName(pZalacznik)
    '                PlikZalacznik = Dir(pZalacznik)
    '                Do While Len(PlikZalacznik) > 0
    '                    LicznikZalacznikow += 1
    '                    .Attachments.Add(PathZalacznik & "\" & PlikZalacznik, _
    '                                    Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, _
    '                                    LicznikZalacznikow, _
    '                                    PlikZalacznik)
    '                    PlikZalacznik = Dir()
    '                Loop
    '            End If
    '            .Display(False) 'True is modal
    '        End With

    '        'usuwamy pliki z dysku...
    '        Dim PathJPG As String = IO.Path.GetDirectoryName(pZalacznik)
    '        Dim plikJPG As String = Dir(pZalacznik)
    '        Do While Len(plikJPG) > 0
    '            IO.File.Delete(PathJPG & "\" & plikJPG)
    '            plikJPG = Dir()
    '        Loop


    '        'Me.Cursor = Cursors.Default
    '        '----------------------------
    '        'If UsunOutllokZPamieci = True Then
    '        '    If Not oOutlook Is Nothing Then
    '        '        oOutlook.Quit()
    '        '        oOutlook = Nothing
    '        '    End If
    '        'End If
    '        oOutlook = Nothing
    '    End If
    'End Sub









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








    '*************************
    'analiza folderów OUTLOOK
    '***********************************

    'http://stackoverflow.com/questions/13376831/vb-net-read-emails-from-outlook-own-folder
    'Sub Main()

    '    Dim otkApp As Outlook.Application = New Outlook.Application
    '    Dim otkMailItem = "IPM.Note"
    '    Dim otkNameSpace As Outlook.NameSpace = otkApp.GetNamespace("MAPI")
    '    Dim otkInboxFolder As Outlook.MAPIFolder = otkNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
    '    Dim otkMailItems As Outlook.Items = otkInboxFolder.Items
    '    Dim otkMessage As Outlook.MailItem
    '    Dim iCntr As Integer

    '    MsgBox(otkMailItems.Count)
    '    For iCntr = 1 To otkMailItems.Count
    '        If otkMailItems.Item(iCntr).MessageClass = otkMailItem Then
    '            otkMessage = otkMailItems.Item(iCntr)

    '            Console.WriteLine(iCntr)
    '            Console.WriteLine(otkMessage.SenderName)
    '            Console.WriteLine(otkMessage.Subject)
    '            Console.WriteLine(otkMessage.ReceivedTime)
    '            Console.WriteLine(otkMessage.Body)
    '            Console.WriteLine("______________________________")
    '        End If
    '    Next

    '    otkApp = Nothing
    '    otkNameSpace = Nothing
    '    otkMailItems = Nothing
    '    otkMessage = Nothing
    'End Sub

    '----------------
    'Dim otkInboxFolder As Outlook.MAPIFolder = otkNameSpace.Folders("support@domain.com").Folders("domain.com").Folders("Inbox")


    Public Sub InicjujStruktureDataSetOutlook(ByRef pDataSetOutlook As DataSet)
        Dim pDataTableOutlook As New System.Data.DataTable

        Dim robCol1 As DataColumn = pDataTableOutlook.Columns.Add("FOLDEROUTLOOK", GetType(System.String))
        Dim robCol2 As DataColumn = pDataTableOutlook.Columns.Add("DATAOTRZYMANIA", GetType(System.String))
        Dim robCol3 As DataColumn = pDataTableOutlook.Columns.Add("ADRESNADAWCY", GetType(System.String))
        Dim robCol4 As DataColumn = pDataTableOutlook.Columns.Add("NAZWANADAWCY", GetType(System.String))
        Dim robCol5 As DataColumn = pDataTableOutlook.Columns.Add("ADRESODBIORCY", GetType(System.String))
        Dim robCol6 As DataColumn = pDataTableOutlook.Columns.Add("NAZWAODBIORCY", GetType(System.String))
        Dim robCol7 As DataColumn = pDataTableOutlook.Columns.Add("TRESC", GetType(System.String))
        'definiuj Dataset
        pDataSetOutlook.Tables.Add(pDataTableOutlook)
        pDataSetOutlook.Locale = New System.Globalization.CultureInfo("pl-PL")
    End Sub


    Public Sub AnalizujFolderOutlook(ByRef pDataSetOutlook As DataSet,
                                     ByVal pAdreseeMailKlienta As String, _
                                     ByVal pFolderOutlook As Integer)
        Dim newrow As DataRow = Nothing
        Dim AdresNadawcy As String = ""
        Dim AdresOdbiorcy As String = ""

        'zapełniamy DATASET
        Dim otkApp As Microsoft.Office.Interop.Outlook.Application = New Microsoft.Office.Interop.Outlook.Application
        Dim otkMailItem = "IPM.Note"
        Dim otkNameSpace As Microsoft.Office.Interop.Outlook.NameSpace = otkApp.GetNamespace("MAPI")

        'standardowy filder outlook
        'Dim otkRootFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = pFolderOutlook

        Dim otkRootFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = _
                    otkNameSpace.GetDefaultFolder(pFolderOutlook)

        'foldery osobiste
        'Dim otkInboxFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = _
        '            otkNameSpace.Folders(pFolderOutlook)   '.Folders("______Private")  '.Folders("Inbox")

        Dim otkMailItems As Microsoft.Office.Interop.Outlook.Items = otkRootFolder.Items
        Dim otkMailFolders As Microsoft.Office.Interop.Outlook.Folders = otkRootFolder.Folders
        Dim otkMessage As Microsoft.Office.Interop.Outlook.MailItem
        Dim otkFolder As String = ""
        Dim iCntr As Integer

        ''analizuj podfoldery
        'For iCntr = 1 To otkMailFolders.Count
        '    otkFolder = otkMailFolders.Item(iCntr).Name
        'Next

        'analizuj przesyłki
        '--------------------------------------
        Dim FormPokazPostep As New _frmPokazIleZrobiles("Analizuję pocztę", otkMailItems.Count)
        FormPokazPostep.Show()
        '--------------------------------------
        For iCntr = 1 To otkMailItems.Count
            '--------------------------------------
            FormPokazPostep.PokazPostep(iCntr)
            '--------------------------------------
            If otkMailItems.Item(iCntr).MessageClass = otkMailItem Then
                otkMessage = otkMailItems.Item(iCntr)



                Select Case pFolderOutlook
                    Case Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail
                        AdresOdbiorcy = otkMessage.To
                        AdresNadawcy = otkMessage.SenderEmailAddress

                        If (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Or _
                            (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Then

                            newrow = pDataSetOutlook.Tables(0).NewRow()
                            newrow("FOLDEROUTLOOK") = otkRootFolder.Name
                            newrow("DATAOTRZYMANIA") = otkMessage.ReceivedTime
                            newrow("ADRESNADAWCY") = otkMessage.SenderEmailAddress
                            newrow("NAZWANADAWCY") = otkMessage.SenderName
                            newrow("ADRESODBIORCY") = otkMessage.To
                            newrow("NAZWAODBIORCY") = otkMessage.To
                            newrow("TRESC") = otkMessage.Body
                            pDataSetOutlook.Tables(0).Rows.Add(newrow)
                        End If

                    Case Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox
                        AdresOdbiorcy = otkMessage.Recipients.Session.CurrentUser.Address
                        AdresNadawcy = otkMessage.SenderEmailAddress

                        If (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Or _
                            (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Then

                            newrow = pDataSetOutlook.Tables(0).NewRow()
                            newrow("FOLDEROUTLOOK") = otkRootFolder.Name
                            newrow("DATAOTRZYMANIA") = otkMessage.ReceivedTime
                            newrow("ADRESNADAWCY") = otkMessage.SenderEmailAddress
                            newrow("NAZWANADAWCY") = otkMessage.SenderName
                            newrow("ADRESODBIORCY") = otkMessage.Recipients.Session.CurrentUser.Address
                            newrow("NAZWAODBIORCY") = otkMessage.Recipients.Session.CurrentUser.Name
                            newrow("TRESC") = otkMessage.Body
                            pDataSetOutlook.Tables(0).Rows.Add(newrow)
                        End If

                    Case Else
                        AdresOdbiorcy = otkMessage.Recipients.Session.CurrentUser.Address
                        AdresNadawcy = otkMessage.SenderEmailAddress

                        If (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Or _
                            (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Then

                            newrow = pDataSetOutlook.Tables(0).NewRow()
                            newrow("FOLDEROUTLOOK") = otkRootFolder.Name
                            newrow("DATAOTRZYMANIA") = otkMessage.ReceivedTime
                            newrow("ADRESNADAWCY") = otkMessage.SenderEmailAddress
                            newrow("NAZWANADAWCY") = otkMessage.SenderName
                            newrow("ADRESODBIORCY") = otkMessage.Recipients.Session.CurrentUser.Address
                            newrow("NAZWAODBIORCY") = otkMessage.Recipients.Session.CurrentUser.Name
                            newrow("TRESC") = otkMessage.Body
                            pDataSetOutlook.Tables(0).Rows.Add(newrow)
                        End If

                End Select


            End If
        Next
        '--------------------------------------
        FormPokazPostep.KoniecPracy()
        '--------------------------------------
        otkApp = Nothing
        otkNameSpace = Nothing
        otkMailItems = Nothing
        otkMessage = Nothing
        '--------------------------------------
    End Sub



    Public Sub AnalizujFolderyOutlookRekurencyjnie(ByRef pDataSetOutlook As DataSet,
                                                   ByVal pAdreseeMailKlienta As String, _
                                                   ByVal pFolderOutlook As String)
        Dim newrow As DataRow = Nothing
        Dim AdresNadawcy As String = ""
        Dim AdresOdbiorcy As String = ""

        'zapełniamy DATASET
        Dim otkApp As Microsoft.Office.Interop.Outlook.Application = New Microsoft.Office.Interop.Outlook.Application
        Dim otkMailItem = "IPM.Note"
        Dim otkNameSpace As Microsoft.Office.Interop.Outlook.NameSpace = otkApp.GetNamespace("MAPI")

        'standardowy filder outlook
        'Dim otkRootFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = pFolderOutlook

        Dim otkRootFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = _
                    otkNameSpace.GetDefaultFolder(pFolderOutlook)

        'foldery osobiste
        'Dim otkInboxFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = _
        '            otkNameSpace.Folders(pFolderOutlook)   '.Folders("______Private")  '.Folders("Inbox")

        Dim otkMailItems As Microsoft.Office.Interop.Outlook.Items = otkRootFolder.Items
        Dim otkMailFolders As Microsoft.Office.Interop.Outlook.Folders = otkRootFolder.Folders
        Dim otkMessage As Microsoft.Office.Interop.Outlook.MailItem
        Dim otkFolder As String = ""
        Dim iCntr As Integer

        ''analizuj podfoldery
        'For iCntr = 1 To otkMailFolders.Count
        '    otkFolder = otkMailFolders.Item(iCntr).Name
        'Next

        'analizuj przesyłki
        '--------------------------------------
        Dim FormPokazPostep As New _frmPokazIleZrobiles("Analizuję pocztę", otkMailItems.Count)
        FormPokazPostep.Show()
        '--------------------------------------
        For iCntr = 1 To otkMailItems.Count
            '--------------------------------------
            FormPokazPostep.PokazPostep(iCntr)
            '--------------------------------------
            If otkMailItems.Item(iCntr).MessageClass = otkMailItem Then
                otkMessage = otkMailItems.Item(iCntr)



                Select Case pFolderOutlook
                    Case Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail
                        AdresOdbiorcy = otkMessage.To
                        AdresNadawcy = otkMessage.SenderEmailAddress

                        If (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Or _
                            (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Then

                            newrow = pDataSetOutlook.Tables(0).NewRow()
                            newrow("FOLDEROUTLOOK") = otkRootFolder.Name
                            newrow("DATAOTRZYMANIA") = otkMessage.ReceivedTime
                            newrow("ADRESNADAWCY") = otkMessage.SenderEmailAddress
                            newrow("NAZWANADAWCY") = otkMessage.SenderName
                            newrow("ADRESODBIORCY") = otkMessage.To
                            newrow("NAZWAODBIORCY") = otkMessage.To
                            newrow("TRESC") = otkMessage.Body
                            pDataSetOutlook.Tables(0).Rows.Add(newrow)
                        End If

                    Case Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox
                        AdresOdbiorcy = otkMessage.Recipients.Session.CurrentUser.Address
                        AdresNadawcy = otkMessage.SenderEmailAddress

                        If (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Or _
                            (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Then

                            newrow = pDataSetOutlook.Tables(0).NewRow()
                            newrow("FOLDEROUTLOOK") = otkRootFolder.Name
                            newrow("DATAOTRZYMANIA") = otkMessage.ReceivedTime
                            newrow("ADRESNADAWCY") = otkMessage.SenderEmailAddress
                            newrow("NAZWANADAWCY") = otkMessage.SenderName
                            newrow("ADRESODBIORCY") = otkMessage.Recipients.Session.CurrentUser.Address
                            newrow("NAZWAODBIORCY") = otkMessage.Recipients.Session.CurrentUser.Name
                            newrow("TRESC") = otkMessage.Body
                            pDataSetOutlook.Tables(0).Rows.Add(newrow)
                        End If

                    Case Else
                        AdresOdbiorcy = otkMessage.Recipients.Session.CurrentUser.Address
                        AdresNadawcy = otkMessage.SenderEmailAddress

                        If (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Or _
                            (UCase(pAdreseeMailKlienta) = UCase(AdresNadawcy)) Then

                            newrow = pDataSetOutlook.Tables(0).NewRow()
                            newrow("FOLDEROUTLOOK") = otkRootFolder.Name
                            newrow("DATAOTRZYMANIA") = otkMessage.ReceivedTime
                            newrow("ADRESNADAWCY") = otkMessage.SenderEmailAddress
                            newrow("NAZWANADAWCY") = otkMessage.SenderName
                            newrow("ADRESODBIORCY") = otkMessage.Recipients.Session.CurrentUser.Address
                            newrow("NAZWAODBIORCY") = otkMessage.Recipients.Session.CurrentUser.Name
                            newrow("TRESC") = otkMessage.Body
                            pDataSetOutlook.Tables(0).Rows.Add(newrow)
                        End If

                End Select


            End If
        Next
        '--------------------------------------
        FormPokazPostep.KoniecPracy()
        '--------------------------------------
        otkApp = Nothing
        otkNameSpace = Nothing
        otkMailItems = Nothing
        otkMessage = Nothing
        '--------------------------------------
    End Sub


End Module
