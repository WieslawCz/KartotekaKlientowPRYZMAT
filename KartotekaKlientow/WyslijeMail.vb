Imports System
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Threading
Imports System.ComponentModel


Public Class WyslijeMail
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

    Private _Adres As String = ""
    Private _Klient As String = ""
    Private _TematListu As String = ""
    Private _TrescListu As String = ""
    Private _ZalacznikListu As String = ""
    Private _Auto As Boolean = False


#Region " Windows Form Designer generated code "


    Public Sub New(ByVal eAdres As String, _
                   ByVal eKlient As String, _
                   ByVal eTemat As String, _
                   ByVal eTresc As String, _
                   ByVal eZalacznik As String, _
                   ByVal eAuto As Boolean)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _Adres = eAdres
        _Klient = eKlient
        _TematListu = eTemat
        _TrescListu = eTresc
        _ZalacznikListu = eZalacznik
        _Auto = eauto
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents tmrStart As System.Windows.Forms.Timer
    Friend WithEvents lblAdres As System.Windows.Forms.Label
    Friend WithEvents lblKlient As System.Windows.Forms.Label
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WyslijeMail))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.lblAdres = New System.Windows.Forms.Label()
        Me.lblKlient = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.tmrStart = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.lblAdres)
        Me.Panel1.Controls.Add(Me.lblKlient)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(472, 56)
        Me.Panel1.TabIndex = 0
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(392, 0)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(72, 56)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 19
        Me.PictureBox3.TabStop = False
        '
        'lblAdres
        '
        Me.lblAdres.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAdres.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdres.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAdres.ForeColor = System.Drawing.Color.Purple
        Me.lblAdres.Location = New System.Drawing.Point(72, 32)
        Me.lblAdres.Name = "lblAdres"
        Me.lblAdres.Size = New System.Drawing.Size(312, 16)
        Me.lblAdres.TabIndex = 15
        '
        'lblKlient
        '
        Me.lblKlient.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblKlient.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKlient.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKlient.ForeColor = System.Drawing.Color.Purple
        Me.lblKlient.Location = New System.Drawing.Point(72, 8)
        Me.lblKlient.Name = "lblKlient"
        Me.lblKlient.Size = New System.Drawing.Size(312, 16)
        Me.lblKlient.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Adres eMail . . . . . "
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Klient . . . . . "
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 69)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 25)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 18
        Me.PictureBox2.TabStop = False
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(341, 69)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(136, 24)
        Me.cmdPowrot.TabIndex = 19
        Me.cmdPowrot.Text = "OK, skoñczy³em..."
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tmrStart
        '
        '
        'WyslijeMail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(493, 100)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WyslijeMail"
        Me.ShowInTaskbar = False
        Me.Text = "Wysy³am eMail"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub WysyanieeMaili_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblKlient.Text = _Klient
        lblAdres.Text = _Adres

        cmdPowrot.Enabled = False
        cmdPowrot.Visible = False

        tmrStart.Interval = 5
        tmrStart.Enabled = True
    End Sub

    Private Sub tmrStart_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStart.Tick
        Dim NazwaKlienta As String
        Dim AdresKlienta As String
        Dim TxtInfo As String

        'get the product name, version and description from the assembly
        TxtInfo = vbCrLf + vbCrLf + _
                  System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).ProductName + _
                  " v" + System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).ProductVersion + vbCrLf + _
                  System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).Comments + vbCrLf + vbCrLf



        tmrStart.Enabled = False

        NazwaKlienta = _Klient
        AdresKlienta = _Adres

        If Len(Trim(AdresKlienta)) > 0 Then
            lblKlient.Text = NazwaKlienta
            lblAdres.Text = AdresKlienta
            System.Windows.Forms.Application.DoEvents()
            ''------------------------------
            ''wysylaj email do tego klienta....
            'Dim MyMessage As New MailMessage
            'If Len(_ZalacznikListu) > 0 Then
            '    Dim myAttachment As New MailAttachment(Trim(_ZalacznikListu), MailEncoding.Base64)
            '    MyMessage.Attachments.Add(myAttachment)
            'End If

            'MyMessage.To = AdresKlienta
            ''MyMessage.From = "me@somewhere.com"    'aktualny adres eMail
            'MyMessage.Subject = _TematListu
            'MyMessage.Priority = MailPriority.High
            'MyMessage.BodyFormat = MailFormat.Text
            'MyMessage.Body = _TrescListu
            'MyMessage.BodyEncoding = System.Text.Encoding.Default
            'Try
            '    ' Change this line to use your mail server!
            '    'SmtpMail.SmtpServer = "test.mailserver.com"
            '    SmtpMail.Send(MyMessage)
            '    'Me.Close()
            'Catch ex As System.Exception
            '    MessageBox.Show("Nie wys³a³em listu do " & AdresKlienta & vbCrLf & _
            '                    "z powodu b³êdu " & ex.ToString & vbCrLf & _
            '                    ex.Message, _
            '                    "Przykro mi", _
            '                    MessageBoxButtons.OK, _
            '                    MessageBoxIcon.Exclamation)
            'End Try
            ''------------------------------








            'Dim message As MailMessage
            'Try
            '    Dim client As SmtpClient = New System.Net.Mail.SmtpClient()
            '    Dim fromAddr As MailAddress = New MailAddress(Trim(ParL_eMailAdres))
            '    Dim toAddr As MailAddress = New MailAddress(Trim(AdresKlienta))

            '    'Specify the message content.
            '    message = New MailMessage(fromAddr, toAddr)
            '    message.Body = _TrescListu
            '    message.BodyEncoding = System.Text.Encoding.UTF8
            '    message.Subject = _TematListu
            '    message.SubjectEncoding = System.Text.Encoding.UTF8
            '    message.Priority = MailPriority.Normal

            '    If Len(_ZalacznikListu) > 0 Then
            '        message.Attachments.Add(New Attachment(Trim(_ZalacznikListu)))
            '    End If

            '    With client
            '        .DeliveryMethod = SmtpDeliveryMethod.Network
            '        .Host = ParL_SMTP
            '        .Port = IIf(ParL_SMTPPort = "", 25, CInt(ParL_SMTPPort))
            '        .UseDefaultCredentials = True
            '        .EnableSsl = IIf(ParL_SSLProtocol = "TAK", True, False)
            '        .UseDefaultCredentials = True
            '        .Credentials = New NetworkCredential(ParL_eMailAdres, ParL_eMailPass)
            '        .Send(message)
            '    End With

            '    'Clean up.
            '    message.Dispose()

            '    If Not _Auto Then
            '        MessageBox.Show("Wiadomoœæ zosta³a wys³ana", _
            '                        "OK", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '    End If


            'Catch ex As System.Exception
            '    MessageBox.Show("Nie wys³a³em poczty..." & vbCrLf & _
            '                    NazwaKlienta & vbCrLf & AdresKlienta & vbCrLf & ex.Message, _
            '                    "Przykro mi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            'Finally
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
            Dim toAddr As MailAddress = New MailAddress(Trim(AdresKlienta))

            'Specify the message content.
            Dim message As MailMessage = New MailMessage(fromAddr, toAddr)
            message.Body = _TrescListu
            message.BodyEncoding = System.Text.Encoding.Default
            message.Subject = _TematListu
            message.SubjectEncoding = System.Text.Encoding.Default  'UTF8
            message.Priority = MailPriority.Normal
            'message.IsBodyHtml = True

            'lista zalacznikow
            'Dim Zalacznik As Attachment = New System.Net.Mail.Attachment(lblZalacznik.Text)
            Dim Zalacznik As Attachment = Nothing
            If Len(Trim(_ZalacznikListu)) > 0 Then
                Zalacznik = New System.Net.Mail.Attachment(_ZalacznikListu)
                message.Attachments.Add(Zalacznik)
            End If

            Try
                client.Send(message)
                If Not _Auto Then
                    MessageBox.Show("Wiadomoœæ zosta³a wys³ana", _
                                    "OK", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If

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







            ''------------------------------------------
            '' Create new email message
            'Dim sm As SimpleMailMessage = New SimpleMailMessage

            'sm.From.Add(New MailBox(ParL_eMailAdres, ParL_eMailAdres))
            'sm.To.Add(New MailBox(AdresKlienta, AdresKlienta))
            'sm.Subject = _TematListu
            'sm.Date = Now
            'sm.TextDataString = _TrescListu

            'If Len(_ZalacznikListu) > 0 Then
            '    ' Read image from disk...
            '    Dim Atmt As MimeData = New MimeData
            '    Atmt.ContentType = New ContentType(MimeType.Other, MimeSubtype.Other)
            '    Atmt.ContentId = "Zalacznik1"
            '    Atmt.LoadFromFile(Trim(_ZalacznikListu))
            '    sm.Attachments.Add(Atmt)
            'End If

            '' Send it
            'Dim smtp As Smtp = New Smtp
            'smtp.User = ParL_eMailUser               '  Set user name and password
            'smtp.Password = ParL_eMailPass
            'Try
            '    smtp.Connect(ParL_SMTP, Nothing, True)       ' Connect to server and login
            '    'smtp.Ehlo(HeloType.EhloHelo, Nothing)    ' Say hello to the server
            '    'smtp.StartTLS()                             ' Starts SSL connection
            '    smtp.Login()

            '    Dim smtpMail As SmtpMail = New SmtpMail(sm.CreateMail())
            '    smtp.SendMessage(smtpMail)          ' Send email message
            'Catch ex As ServerException
            '    MessageBox.Show("Nie wys³a³em listu do " & AdresKlienta & vbCrLf & _
            '                    "z powodu b³êdu " & ex.ToString & vbCrLf & _
            '                    ex.Message, _
            '                    "Przykro mi", _
            '                    MessageBoxButtons.OK, _
            '                    MessageBoxIcon.Exclamation)
            'Finally
            '    smtp.Close(False)                   'Close connection
            'End Try
            ''------------------------------------------






        End If

        If _Auto Then
            Me.Close()
        Else
            cmdPowrot.Enabled = True
            cmdPowrot.Visible = True
        End If
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub WyslijeMail_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

End Class
