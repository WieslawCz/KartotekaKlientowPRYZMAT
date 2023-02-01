Imports System
'Imports System.Net
'Imports System.Net.Mail
'Imports System.Net.Mime
'Imports System.Threading
'Imports System.ComponentModel


Public Class WysylanieeMaili
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

    Private dvTabelaAdresow As DataView
    Private TematListu As String
    Private TrescListu As String
    Private ZalacznikListu As String

#Region " Windows Form Designer generated code "


    Public Sub New(ByRef eDataView As DataView, _
                   ByVal eTemat As String, _
                   ByVal eTresc As String, _
                   ByVal eZalacznik As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        dvTabelaAdresow = eDataView
        TematListu = eTemat
        TrescListu = eTresc
        ZalacznikListu = eZalacznik
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents tmrStart As System.Windows.Forms.Timer
    Friend WithEvents lblAdres As System.Windows.Forms.Label
    Friend WithEvents lblKlient As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblIleKlientow As System.Windows.Forms.Label
    Friend WithEvents lblIleeMaili As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WysylanieeMaili))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.lblIleeMaili = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblAdres = New System.Windows.Forms.Label()
        Me.lblKlient = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblIleKlientow = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.tmrStart = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.PictureBox4)
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.lblIleeMaili)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.lblAdres)
        Me.Panel1.Controls.Add(Me.lblKlient)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.lblIleKlientow)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(400, 104)
        Me.Panel1.TabIndex = 0
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(160, 0)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(72, 56)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 20
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(320, 0)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(72, 56)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 19
        Me.PictureBox3.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(240, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(72, 56)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 18
        Me.PictureBox1.TabStop = False
        '
        'lblIleeMaili
        '
        Me.lblIleeMaili.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblIleeMaili.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleeMaili.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIleeMaili.ForeColor = System.Drawing.Color.Purple
        Me.lblIleeMaili.Location = New System.Drawing.Point(72, 8)
        Me.lblIleeMaili.Name = "lblIleeMaili"
        Me.lblIleeMaili.Size = New System.Drawing.Size(80, 16)
        Me.lblIleeMaili.TabIndex = 17
        Me.lblIleeMaili.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 16)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "eMail nr . . . . . "
        '
        'lblAdres
        '
        Me.lblAdres.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAdres.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdres.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAdres.ForeColor = System.Drawing.Color.Purple
        Me.lblAdres.Location = New System.Drawing.Point(72, 80)
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
        Me.lblKlient.Location = New System.Drawing.Point(72, 56)
        Me.lblKlient.Name = "lblKlient"
        Me.lblKlient.Size = New System.Drawing.Size(312, 16)
        Me.lblKlient.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Adres eMail . . . . . "
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Klient . . . . . "
        '
        'lblIleKlientow
        '
        Me.lblIleKlientow.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblIleKlientow.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleKlientow.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIleKlientow.ForeColor = System.Drawing.Color.Purple
        Me.lblIleKlientow.Location = New System.Drawing.Point(72, 32)
        Me.lblIleKlientow.Name = "lblIleKlientow"
        Me.lblIleKlientow.Size = New System.Drawing.Size(80, 16)
        Me.lblIleKlientow.TabIndex = 11
        Me.lblIleKlientow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Klient nr . . . . . "
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 117)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 18
        Me.PictureBox2.TabStop = False
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(273, 117)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(136, 23)
        Me.cmdPowrot.TabIndex = 19
        Me.cmdPowrot.Text = "OK, skoñczy³em..."
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tmrStart
        '
        '
        'WysylanieeMaili
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(425, 147)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WysylanieeMaili"
        Me.ShowInTaskbar = False
        Me.Text = "Wysy³am eMaile do klientów"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private licznikKlientow As Long = 0
    Private licznikeMaili As Long = 0

    Private Sub WysyanieeMaili_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblKlient.Text = ""
        lblAdres.Text = ""
        lblIleKlientow.Text = Str(licznikKlientow)
        lblIleeMaili.Text = Str(licznikeMaili)
        cmdPowrot.Enabled = False
        cmdPowrot.Visible = False

        tmrStart.Interval = 5
        tmrStart.Enabled = True
    End Sub

    Private Sub tmrStart_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStart.Tick
        Dim NazwaKlienta As String
        Dim AdresKlienta As String

        tmrStart.Enabled = False
        For licznikKlientow = 0 To dvTabelaAdresow.Count - 1
            lblIleKlientow.Text = Str(licznikKlientow)
            lblKlient.Text = ""
            lblAdres.Text = ""

            NazwaKlienta = dvTabelaAdresow.Item(licznikKlientow).Item("NAZWAKLIENTA")
            AdresKlienta = dvTabelaAdresow.Item(licznikKlientow).Item("eMAIL")
            If Len(Trim(AdresKlienta)) > 0 Then
                licznikeMaili += 1
                lblIleeMaili.Text = Str(licznikeMaili)
                lblKlient.Text = NazwaKlienta
                lblAdres.Text = AdresKlienta
                System.Windows.Forms.Application.DoEvents()


                Select Case ParL_eMailService
                    Case "Wbudowany"
                        'Dim Form As New WyslijeMail(AdresKlienta & ";" & AdresKlienta & ";" & AdresKlienta & ";" & AdresKlienta & ";" & AdresKlienta, AdresKlienta, TematListu, TrescListu, ZalacznikListu)
                        Dim Form As New WyslijeMail(AdresKlienta, NazwaKlienta, TematListu, TrescListu, ZalacznikListu, True)
                        Form.ShowDialog()
                    Case "MS Outlook"
                        SendByOutlook(AdresKlienta, TematListu, TrescListu, ZalacznikListu, "Wyslij")
                    Case "Mozilla Thunderbird"
                        SendByThunderbird(AdresKlienta, "", "", TematListu, TrescListu, ZalacznikListu)

                    Case Else
                End Select


                'If Wysylaj(AdresKlienta, TematListu, TrescListu, ZalacznikListu) Then
                'End If


            End If
        Next
        cmdPowrot.Enabled = True
        cmdPowrot.Visible = True
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub WysylanieeMaili_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub









    'Private Function Wysylaj(ByVal _Adres As String, _
    '                         ByVal _Temat As String, _
    '                         ByVal _Tresc As String, _
    '                         ByVal _Zalacznik As String) As Boolean

    '    Dim client As SmtpClient = New System.Net.Mail.SmtpClient()
    '    Dim fromAddr As MailAddress = New MailAddress(Trim(ParL_eMailAdres))
    '    Dim toAddr As MailAddress = New MailAddress(Trim(_Adres))

    '    'Specify the message content.
    '    Dim message As MailMessage = New MailMessage(fromAddr, toAddr)
    '    message.Body = _Tresc
    '    message.BodyEncoding = System.Text.Encoding.UTF8
    '    message.Subject = _Temat
    '    message.SubjectEncoding = System.Text.Encoding.UTF8
    '    message.Priority = MailPriority.Normal

    '    If Len(_Zalacznik) > 0 Then
    '        message.Attachments.Add(New Attachment(Trim(_Zalacznik)))
    '    End If

    '    Try
    '        client.UseDefaultCredentials = True
    '        client.DeliveryMethod = SmtpDeliveryMethod.Network
    '        client.Host = ParL_SMTP
    '        client.Credentials = New NetworkCredential(ParL_eMailAdres, ParL_eMailPass)
    '        client.Send(message)
    '        '---------MessageBox.Show("Wiadomoœæ zosta³a wys³ana", "OK", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '    Catch ex As System.Exception
    '        MessageBox.Show("Nie wys³a³em poczty..." & vbCrLf & ex.Message, "Przykro mi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '    Finally
    '        'Clean up.
    '        message.Dispose()
    '    End Try

    'End Function


End Class
