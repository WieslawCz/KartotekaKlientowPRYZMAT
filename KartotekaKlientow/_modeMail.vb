Imports System.IO
Imports System
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Threading
Imports System.ComponentModel

Module _modeMail

    Public Sub WyslijeMail2(ByVal eMailOdbiorcy As String, _
                           ByVal Temat As String, _
                           ByVal Tresc As String, _
                           ByVal Zalacznik As String)
        'Dim client As SmtpClient
        'Dim fromAddr As MailAddress
        'Dim toAddr As MailAddress

        If Not IsValidEmail(Trim(ParL_eMailAdres)) Then
            MessageBox.Show("Nie wys³a³em poczty..." & vbCrLf & _
                            "Adres nadawcy nie jest poprawny" & vbCrLf & _
                            Trim(ParL_eMailAdres), _
                            "Przykro mi", _
                            MessageBoxButtons.OK, _
                            MessageBoxIcon.Exclamation)
        ElseIf Not IsValidEmail(eMailOdbiorcy) Then
            MessageBox.Show("Nie wys³a³em poczty..." & vbCrLf & _
                            "Adres odbiorcy nie jest poprawny" & vbCrLf & _
                            eMailOdbiorcy, _
                            "Przykro mi", _
                            MessageBoxButtons.OK, _
                            MessageBoxIcon.Exclamation)
        Else
            'client = New System.Net.Mail.SmtpClient()
            'fromAddr = New MailAddress(Trim(ParL_eMailAdres))
            'toAddr = New MailAddress(eMailOdbiorcy)

            ''Specify the message content.
            'Dim message As MailMessage = New MailMessage(fromAddr, toAddr)
            'message.Body = Tresc
            'message.BodyEncoding = System.Text.Encoding.UTF8
            'message.Subject = Temat
            'message.SubjectEncoding = System.Text.Encoding.UTF8
            'message.Priority = MailPriority.Normal

            'If Len(Zalacznik) > 0 Then
            '    message.Attachments.Add(New Attachment(Trim(Zalacznik)))
            'End If

            'Try
            '    client.DeliveryMethod = SmtpDeliveryMethod.Network
            '    client.Host = ParL_SMTP
            '    client.Port = IIf(ParL_SMTPPort = "", 25, CInt(ParL_SMTPPort))
            '    client.UseDefaultCredentials = True
            '    client.EnableSsl = (ParL_SSLProtocol = "TAK")
            '    client.UseDefaultCredentials = True
            '    client.Credentials = New NetworkCredential(ParL_eMailUser, ParL_eMailPass)
            '    client.Send(message)
            '    MessageBox.Show("Wiadomoœæ zosta³a wys³ana", _
            '                    "OK", _
            '                    MessageBoxButtons.OK, _
            '                    MessageBoxIcon.Exclamation)
            'Catch ex As System.Exception
            '    MessageBox.Show("Nie wys³a³em poczty..." & vbCrLf & ex.Message, _
            '                    "Przykro mi", _
            '                    MessageBoxButtons.OK, _
            '                    MessageBoxIcon.Exclamation)
            'Finally
            '    'Clean up.
            '    message.Dispose()
            'End Try





            'Command line argument must the the SMTP host.
            Dim client As SmtpClient = New System.Net.Mail.SmtpClient()
            With client
                .Host = ParL_SMTP
                .Port = IIf(ParL_SMTPPort = "", 25, CInt(ParL_SMTPPort))
                .DeliveryMethod = SmtpDeliveryMethod.Network
                .Port = IIf(ParL_SMTPPort = "", 25, CInt(ParL_SMTPPort))
                .EnableSsl = IIf(ParL_SSLProtocol = "TAK", True, False)
                .UseDefaultCredentials = True
                .Credentials = New NetworkCredential(ParL_eMailUser, ParL_eMailPass)
            End With


            '--------------------------------------
            ' dodanie Time Stemp powoduje problemy z polskimi znakami w dacie (Œroda, PaŸdziernik etc)
            '--------------------------------------
            ''Add time stamp information for the file.
            'Dim disposition As ContentDisposition = Zalacznik.ContentDisposition
            'disposition.CreationDate = System.IO.File.GetCreationTime(lblZalacznik.Text)
            'disposition.ModificationDate = System.IO.File.GetLastWriteTime(lblZalacznik.Text)
            'disposition.ReadDate = System.IO.File.GetLastAccessTime(lblZalacznik.Text)
            'disposition.DispositionType = DispositionTypeNames.Attachment
            '--------------------------------------

            'Specify the e-mail sender.
            Dim fromAddr As MailAddress = New MailAddress(Trim(ParL_eMailAdres))
            'Set destinations for the e-mail message.
            Dim toAddr As MailAddress = New MailAddress(Trim(eMailOdbiorcy))

            'Specify the message content.
            Dim message As MailMessage = New MailMessage(fromAddr, toAddr)
            message.Body = Tresc
            message.BodyEncoding = System.Text.Encoding.Default
            message.Subject = Temat
            message.SubjectEncoding = System.Text.Encoding.Default  'UTF8
            message.Priority = MailPriority.Normal
            'message.IsBodyHtml = True

            'lista zalacznikow
            'Dim Zalacznik As Attachment = New System.Net.Mail.Attachment(lblZalacznik.Text)
            Dim ZalacznikListu As Attachment = Nothing
            If Len(Trim(Zalacznik)) > 0 Then
                ZalacznikListu = New System.Net.Mail.Attachment(Zalacznik)
                message.Attachments.Add(ZalacznikListu)
            End If

            Try
                client.Send(message)
                MessageBox.Show("Wiadomoœæ zosta³a wys³ana", _
                                    "OK", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            Catch ex As System.Exception
                MessageBox.Show("Nie wys³a³em poczty..." & vbCrLf & _
                                ex.Message & vbCrLf & _
                                ex.StackTrace, _
                                "Przykro mi", _
                                MessageBoxButtons.OK, _
                                MessageBoxIcon.Exclamation)
            Finally
                'Clean up.
                message.Dispose()
            End Try

        End If
    End Sub





    ''' <summary>
    ''' method for determining is the user provided a valid email address
    ''' We use regular expressions in this check, as it is a more thorough
    ''' way of checking the address provided
    ''' </summary>
    ''' <param name="email">email address to validate</param>
    ''' <returns>true is valid, false if not valid</returns>
    Public Function IsValidEmail(ByVal email As String) As Boolean
        'regular expression pattern for valid email
        'addresses, allows for the following domains:
        'com,edu,info,gov,int,mil,net,org,biz,name,museum,coop,aero,pro,tv
        Dim pattern As String = "^[-a-zA-Z0-9_][-.a-zA-Z0-9_]*@[-.a-zA-Z0-9_]+(\.[-.a-zA-Z0-9_]+)*\." & _
        "(com|edu|info|gov|int|mil|net|org|biz|name|museum|coop|aero|pro|tv|[a-zA-Z]{2})$"

        'Regular expression object
        Dim check As New Text.RegularExpressions.Regex(pattern, Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace)
        'boolean variable to return to calling method
        Dim valid As Boolean = False

        'make sure an email address was provided
        If String.IsNullOrEmpty(email) Then
            valid = False
        Else
            'use IsMatch to validate the address
            valid = check.IsMatch(email)
        End If
        'return the value to the calling method
        Return valid
    End Function


End Module
