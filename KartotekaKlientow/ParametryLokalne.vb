Public Class ParametryLokalne
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

    ' Declare the API function.
    Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long

    ' Define the possible types of connections.
    Private Enum ConnectStates
        LAN = &H2
        Modem = &H1
        Proxy = &H4
        Offline = &H20
        Configured = &H40
        RasInstalled = &H10
    End Enum

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents tbcParametry As System.Windows.Forms.TabControl
    Friend WithEvents Instalacja As System.Windows.Forms.TabPage
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents cmdZapisz As System.Windows.Forms.Button
    Friend WithEvents cmdOdswiez As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtOpisInstalacji As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dlgOpenFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtAdres As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSMTP As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPOP3 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtLogin As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtPass As System.Windows.Forms.TextBox
    Friend WithEvents TabBaza As System.Windows.Forms.TabPage
    Friend WithEvents cmdPobierzKatalog As System.Windows.Forms.Button
    Friend WithEvents lblKatalog As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblJakaBaza As System.Windows.Forms.Label
    Friend WithEvents cbxBazaDanych As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSQLBaza As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSQLServer As System.Windows.Forms.ComboBox
    Friend WithEvents btnTestuj As System.Windows.Forms.Button
    Friend WithEvents rbtSQLServer As System.Windows.Forms.RadioButton
    Friend WithEvents rbtWindows As System.Windows.Forms.RadioButton
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtSQLPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtSQLLogin As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblIP As System.Windows.Forms.Label
    Friend WithEvents lblNazwa As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtNazwaArchiwum As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents btnNowaBaza As System.Windows.Forms.Button
    Friend WithEvents txtPOP3Port As System.Windows.Forms.TextBox
    Friend WithEvents txtSMTPPort As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents chkSSL As System.Windows.Forms.CheckBox
    Friend WithEvents cbxeMailService As System.Windows.Forms.ComboBox
    Friend WithEvents Label69 As System.Windows.Forms.Label
    Friend WithEvents cmdPokazHasloeMail As System.Windows.Forms.Button
    Friend WithEvents cmdPokazSQLPass As System.Windows.Forms.Button
    Friend WithEvents cmdPobierzKatalogArchiwum As System.Windows.Forms.Button
    Friend WithEvents txtKatalogArchiwum As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents cmdPobierzKatalogRaporty As System.Windows.Forms.Button
    Friend WithEvents txtKatalogRaporty As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cmdPobierzKatalogArchZIP As System.Windows.Forms.Button
    Friend WithEvents txtKatalogArchiwumZIP As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TabBazaImport As TabPage
    Friend WithEvents lblHQIP As Label
    Friend WithEvents lblHQNazwa As Label
    Friend WithEvents cbxHQSQLBaza As ComboBox
    Friend WithEvents txtHQSQLPassword As TextBox
    Friend WithEvents txtHQSQLLogin As TextBox
    Friend WithEvents cbxHQSQLServer As ComboBox
    Friend WithEvents rbtHQWindows As RadioButton
    Friend WithEvents Label83 As Label
    Friend WithEvents cmdHQSQLPokazHaslo As Button
    Friend WithEvents Label25 As Label
    Friend WithEvents Label26 As Label
    Friend WithEvents cmdHQTestuj As Button
    Friend WithEvents rbtHQSQLServer As RadioButton
    Friend WithEvents Label27 As Label
    Friend WithEvents Label28 As Label
    Friend WithEvents Label29 As Label
    Friend WithEvents Label30 As Label
    Friend WithEvents Label31 As Label
    Friend WithEvents cmdHQPobierz As Button
    Friend WithEvents cmdHQWersjaBazyDanych As Button
    Friend WithEvents txtHQWersja As TextBox
    Friend WithEvents Label23 As Label
    Friend WithEvents txtHQWersjaWymagana As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents Label33 As Label
    Friend WithEvents Label32 As Label
    Friend WithEvents txtPZnakIdPracownika As TextBox
    Friend WithEvents Label35 As Label
    Friend WithEvents txtPZnakIdKlienta As TextBox
    Friend WithEvents Label34 As Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ParametryLokalne))
        Me.tbcParametry = New System.Windows.Forms.TabControl()
        Me.Instalacja = New System.Windows.Forms.TabPage()
        Me.cmdPobierzKatalogArchZIP = New System.Windows.Forms.Button()
        Me.txtKatalogArchiwumZIP = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.cmdPobierzKatalogRaporty = New System.Windows.Forms.Button()
        Me.txtKatalogRaporty = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.cmdPobierzKatalogArchiwum = New System.Windows.Forms.Button()
        Me.txtKatalogArchiwum = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.cbxeMailService = New System.Windows.Forms.ComboBox()
        Me.Label69 = New System.Windows.Forms.Label()
        Me.cmdPokazHasloeMail = New System.Windows.Forms.Button()
        Me.chkSSL = New System.Windows.Forms.CheckBox()
        Me.txtPOP3Port = New System.Windows.Forms.TextBox()
        Me.txtSMTPPort = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtNazwaArchiwum = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.cbxBazaDanych = New System.Windows.Forms.ComboBox()
        Me.txtPass = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtLogin = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtPOP3 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSMTP = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtAdres = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtOpisInstalacji = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabBaza = New System.Windows.Forms.TabPage()
        Me.cmdPokazSQLPass = New System.Windows.Forms.Button()
        Me.btnNowaBaza = New System.Windows.Forms.Button()
        Me.lblIP = New System.Windows.Forms.Label()
        Me.lblNazwa = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cbxSQLBaza = New System.Windows.Forms.ComboBox()
        Me.cbxSQLServer = New System.Windows.Forms.ComboBox()
        Me.btnTestuj = New System.Windows.Forms.Button()
        Me.rbtSQLServer = New System.Windows.Forms.RadioButton()
        Me.rbtWindows = New System.Windows.Forms.RadioButton()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtSQLPassword = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtSQLLogin = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmdPobierzKatalog = New System.Windows.Forms.Button()
        Me.lblKatalog = New System.Windows.Forms.Label()
        Me.lblJakaBaza = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TabBazaImport = New System.Windows.Forms.TabPage()
        Me.txtPZnakIdPracownika = New System.Windows.Forms.TextBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.txtPZnakIdKlienta = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.txtHQWersjaWymagana = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtHQWersja = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.cmdHQWersjaBazyDanych = New System.Windows.Forms.Button()
        Me.cmdHQPobierz = New System.Windows.Forms.Button()
        Me.lblHQIP = New System.Windows.Forms.Label()
        Me.lblHQNazwa = New System.Windows.Forms.Label()
        Me.cbxHQSQLBaza = New System.Windows.Forms.ComboBox()
        Me.txtHQSQLPassword = New System.Windows.Forms.TextBox()
        Me.txtHQSQLLogin = New System.Windows.Forms.TextBox()
        Me.cbxHQSQLServer = New System.Windows.Forms.ComboBox()
        Me.rbtHQWindows = New System.Windows.Forms.RadioButton()
        Me.Label83 = New System.Windows.Forms.Label()
        Me.cmdHQSQLPokazHaslo = New System.Windows.Forms.Button()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.cmdHQTestuj = New System.Windows.Forms.Button()
        Me.rbtHQSQLServer = New System.Windows.Forms.RadioButton()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.cmdZapisz = New System.Windows.Forms.Button()
        Me.cmdOdswiez = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.dlgOpenFolder = New System.Windows.Forms.FolderBrowserDialog()
        Me.tbcParametry.SuspendLayout()
        Me.Instalacja.SuspendLayout()
        Me.TabBaza.SuspendLayout()
        Me.TabBazaImport.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbcParametry
        '
        Me.tbcParametry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbcParametry.Controls.Add(Me.Instalacja)
        Me.tbcParametry.Controls.Add(Me.TabBaza)
        Me.tbcParametry.Controls.Add(Me.TabBazaImport)
        Me.tbcParametry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.tbcParametry.Location = New System.Drawing.Point(8, 8)
        Me.tbcParametry.Name = "tbcParametry"
        Me.tbcParametry.SelectedIndex = 0
        Me.tbcParametry.Size = New System.Drawing.Size(774, 413)
        Me.tbcParametry.TabIndex = 0
        '
        'Instalacja
        '
        Me.Instalacja.BackColor = System.Drawing.SystemColors.Control
        Me.Instalacja.Controls.Add(Me.cmdPobierzKatalogArchZIP)
        Me.Instalacja.Controls.Add(Me.txtKatalogArchiwumZIP)
        Me.Instalacja.Controls.Add(Me.Label22)
        Me.Instalacja.Controls.Add(Me.cmdPobierzKatalogRaporty)
        Me.Instalacja.Controls.Add(Me.txtKatalogRaporty)
        Me.Instalacja.Controls.Add(Me.Label21)
        Me.Instalacja.Controls.Add(Me.cmdPobierzKatalogArchiwum)
        Me.Instalacja.Controls.Add(Me.txtKatalogArchiwum)
        Me.Instalacja.Controls.Add(Me.Label20)
        Me.Instalacja.Controls.Add(Me.cbxeMailService)
        Me.Instalacja.Controls.Add(Me.Label69)
        Me.Instalacja.Controls.Add(Me.cmdPokazHasloeMail)
        Me.Instalacja.Controls.Add(Me.chkSSL)
        Me.Instalacja.Controls.Add(Me.txtPOP3Port)
        Me.Instalacja.Controls.Add(Me.txtSMTPPort)
        Me.Instalacja.Controls.Add(Me.Label19)
        Me.Instalacja.Controls.Add(Me.Label18)
        Me.Instalacja.Controls.Add(Me.txtNazwaArchiwum)
        Me.Instalacja.Controls.Add(Me.Label17)
        Me.Instalacja.Controls.Add(Me.cbxBazaDanych)
        Me.Instalacja.Controls.Add(Me.txtPass)
        Me.Instalacja.Controls.Add(Me.Label8)
        Me.Instalacja.Controls.Add(Me.txtLogin)
        Me.Instalacja.Controls.Add(Me.Label7)
        Me.Instalacja.Controls.Add(Me.txtPOP3)
        Me.Instalacja.Controls.Add(Me.Label6)
        Me.Instalacja.Controls.Add(Me.txtSMTP)
        Me.Instalacja.Controls.Add(Me.Label5)
        Me.Instalacja.Controls.Add(Me.txtAdres)
        Me.Instalacja.Controls.Add(Me.Label4)
        Me.Instalacja.Controls.Add(Me.Label3)
        Me.Instalacja.Controls.Add(Me.Label2)
        Me.Instalacja.Controls.Add(Me.txtOpisInstalacji)
        Me.Instalacja.Controls.Add(Me.Label1)
        Me.Instalacja.ForeColor = System.Drawing.Color.Navy
        Me.Instalacja.Location = New System.Drawing.Point(4, 23)
        Me.Instalacja.Name = "Instalacja"
        Me.Instalacja.Size = New System.Drawing.Size(766, 386)
        Me.Instalacja.TabIndex = 0
        Me.Instalacja.Text = "Instalacja"
        '
        'cmdPobierzKatalogArchZIP
        '
        Me.cmdPobierzKatalogArchZIP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPobierzKatalogArchZIP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPobierzKatalogArchZIP.Image = CType(resources.GetObject("cmdPobierzKatalogArchZIP.Image"), System.Drawing.Image)
        Me.cmdPobierzKatalogArchZIP.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdPobierzKatalogArchZIP.Location = New System.Drawing.Point(718, 77)
        Me.cmdPobierzKatalogArchZIP.Name = "cmdPobierzKatalogArchZIP"
        Me.cmdPobierzKatalogArchZIP.Size = New System.Drawing.Size(40, 24)
        Me.cmdPobierzKatalogArchZIP.TabIndex = 345
        Me.cmdPobierzKatalogArchZIP.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtKatalogArchiwumZIP
        '
        Me.txtKatalogArchiwumZIP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtKatalogArchiwumZIP.ForeColor = System.Drawing.Color.Purple
        Me.txtKatalogArchiwumZIP.Location = New System.Drawing.Point(132, 79)
        Me.txtKatalogArchiwumZIP.Name = "txtKatalogArchiwumZIP"
        Me.txtKatalogArchiwumZIP.Size = New System.Drawing.Size(626, 20)
        Me.txtKatalogArchiwumZIP.TabIndex = 344
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(8, 82)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(129, 20)
        Me.Label22.TabIndex = 343
        Me.Label22.Text = "Katalog Archiwum ZIP . . . . . . . . . . . "
        '
        'cmdPobierzKatalogRaporty
        '
        Me.cmdPobierzKatalogRaporty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPobierzKatalogRaporty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPobierzKatalogRaporty.Image = CType(resources.GetObject("cmdPobierzKatalogRaporty.Image"), System.Drawing.Image)
        Me.cmdPobierzKatalogRaporty.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdPobierzKatalogRaporty.Location = New System.Drawing.Point(718, 101)
        Me.cmdPobierzKatalogRaporty.Name = "cmdPobierzKatalogRaporty"
        Me.cmdPobierzKatalogRaporty.Size = New System.Drawing.Size(40, 24)
        Me.cmdPobierzKatalogRaporty.TabIndex = 342
        Me.cmdPobierzKatalogRaporty.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtKatalogRaporty
        '
        Me.txtKatalogRaporty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtKatalogRaporty.ForeColor = System.Drawing.Color.Purple
        Me.txtKatalogRaporty.Location = New System.Drawing.Point(132, 103)
        Me.txtKatalogRaporty.Name = "txtKatalogRaporty"
        Me.txtKatalogRaporty.Size = New System.Drawing.Size(626, 20)
        Me.txtKatalogRaporty.TabIndex = 341
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(8, 106)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(129, 20)
        Me.Label21.TabIndex = 340
        Me.Label21.Text = "Katalog na Raporty . . . . . . . . . . . "
        '
        'cmdPobierzKatalogArchiwum
        '
        Me.cmdPobierzKatalogArchiwum.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPobierzKatalogArchiwum.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPobierzKatalogArchiwum.Image = CType(resources.GetObject("cmdPobierzKatalogArchiwum.Image"), System.Drawing.Image)
        Me.cmdPobierzKatalogArchiwum.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdPobierzKatalogArchiwum.Location = New System.Drawing.Point(718, 53)
        Me.cmdPobierzKatalogArchiwum.Name = "cmdPobierzKatalogArchiwum"
        Me.cmdPobierzKatalogArchiwum.Size = New System.Drawing.Size(40, 24)
        Me.cmdPobierzKatalogArchiwum.TabIndex = 339
        Me.cmdPobierzKatalogArchiwum.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtKatalogArchiwum
        '
        Me.txtKatalogArchiwum.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtKatalogArchiwum.ForeColor = System.Drawing.Color.Purple
        Me.txtKatalogArchiwum.Location = New System.Drawing.Point(132, 55)
        Me.txtKatalogArchiwum.Name = "txtKatalogArchiwum"
        Me.txtKatalogArchiwum.Size = New System.Drawing.Size(626, 20)
        Me.txtKatalogArchiwum.TabIndex = 338
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(8, 58)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(129, 20)
        Me.Label20.TabIndex = 337
        Me.Label20.Text = "Katalog Archiwum . . . . . . . . . . . "
        '
        'cbxeMailService
        '
        Me.cbxeMailService.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cbxeMailService.FormattingEnabled = True
        Me.cbxeMailService.Location = New System.Drawing.Point(447, 215)
        Me.cbxeMailService.Name = "cbxeMailService"
        Me.cbxeMailService.Size = New System.Drawing.Size(187, 22)
        Me.cbxeMailService.TabIndex = 336
        '
        'Label69
        '
        Me.Label69.ForeColor = System.Drawing.Color.Navy
        Me.Label69.Location = New System.Drawing.Point(8, 218)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(546, 16)
        Me.Label69.TabIndex = 335
        Me.Label69.Text = "Serwis wykorzystywany do wysy³ania i odbierania poczty email przez program . . . " &
    ". . . . . . . . "
        '
        'cmdPokazHasloeMail
        '
        Me.cmdPokazHasloeMail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazHasloeMail.Location = New System.Drawing.Point(550, 313)
        Me.cmdPokazHasloeMail.Name = "cmdPokazHasloeMail"
        Me.cmdPokazHasloeMail.Size = New System.Drawing.Size(84, 23)
        Me.cmdPokazHasloeMail.TabIndex = 323
        Me.cmdPokazHasloeMail.Text = "Poka¿ has³o"
        '
        'chkSSL
        '
        Me.chkSSL.Location = New System.Drawing.Point(11, 337)
        Me.chkSSL.Name = "chkSSL"
        Me.chkSSL.Size = New System.Drawing.Size(605, 17)
        Me.chkSSL.TabIndex = 74
        Me.chkSSL.Text = "(SSL) Transmisja z u¿yciem bezpiecznego protoko³u wykorzystujacego certyfikaty"
        Me.chkSSL.UseVisualStyleBackColor = True
        '
        'txtPOP3Port
        '
        Me.txtPOP3Port.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPOP3Port.ForeColor = System.Drawing.Color.Purple
        Me.txtPOP3Port.Location = New System.Drawing.Point(550, 290)
        Me.txtPOP3Port.Name = "txtPOP3Port"
        Me.txtPOP3Port.Size = New System.Drawing.Size(208, 20)
        Me.txtPOP3Port.TabIndex = 73
        '
        'txtSMTPPort
        '
        Me.txtSMTPPort.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSMTPPort.ForeColor = System.Drawing.Color.Purple
        Me.txtSMTPPort.Location = New System.Drawing.Point(550, 266)
        Me.txtSMTPPort.Name = "txtSMTPPort"
        Me.txtSMTPPort.Size = New System.Drawing.Size(208, 20)
        Me.txtSMTPPort.TabIndex = 72
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(487, 293)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(129, 20)
        Me.Label19.TabIndex = 71
        Me.Label19.Text = "Port POP3 . . . . . . . . . . . "
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(487, 269)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(129, 20)
        Me.Label18.TabIndex = 70
        Me.Label18.Text = "Port SMTP . . . . . . . . . . . "
        '
        'txtNazwaArchiwum
        '
        Me.txtNazwaArchiwum.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNazwaArchiwum.ForeColor = System.Drawing.Color.Purple
        Me.txtNazwaArchiwum.Location = New System.Drawing.Point(132, 31)
        Me.txtNazwaArchiwum.Name = "txtNazwaArchiwum"
        Me.txtNazwaArchiwum.Size = New System.Drawing.Size(626, 20)
        Me.txtNazwaArchiwum.TabIndex = 69
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(8, 34)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(129, 20)
        Me.Label17.TabIndex = 68
        Me.Label17.Text = "Nazwa pliku Archiwum . . . . . . . . . . . "
        '
        'cbxBazaDanych
        '
        Me.cbxBazaDanych.ForeColor = System.Drawing.Color.Purple
        Me.cbxBazaDanych.Location = New System.Drawing.Point(258, 143)
        Me.cbxBazaDanych.Name = "cbxBazaDanych"
        Me.cbxBazaDanych.Size = New System.Drawing.Size(166, 22)
        Me.cbxBazaDanych.TabIndex = 67
        '
        'txtPass
        '
        Me.txtPass.ForeColor = System.Drawing.Color.Purple
        Me.txtPass.Location = New System.Drawing.Point(388, 314)
        Me.txtPass.Name = "txtPass"
        Me.txtPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPass.Size = New System.Drawing.Size(156, 20)
        Me.txtPass.TabIndex = 28
        Me.txtPass.Text = "TextBox1"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(331, 317)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(129, 20)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "Has³o . . . . . . . . . . . . . . . . "
        '
        'txtLogin
        '
        Me.txtLogin.ForeColor = System.Drawing.Color.Purple
        Me.txtLogin.Location = New System.Drawing.Point(132, 314)
        Me.txtLogin.Name = "txtLogin"
        Me.txtLogin.Size = New System.Drawing.Size(193, 20)
        Me.txtLogin.TabIndex = 26
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 317)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(129, 20)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "Login . . . . . . . . . . . . . . . . . "
        '
        'txtPOP3
        '
        Me.txtPOP3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPOP3.ForeColor = System.Drawing.Color.Purple
        Me.txtPOP3.Location = New System.Drawing.Point(132, 290)
        Me.txtPOP3.Name = "txtPOP3"
        Me.txtPOP3.Size = New System.Drawing.Size(419, 20)
        Me.txtPOP3.TabIndex = 24
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 293)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(129, 20)
        Me.Label6.TabIndex = 23
        Me.Label6.Text = "Serwer POP3 . . . . . . . . . . . "
        '
        'txtSMTP
        '
        Me.txtSMTP.ForeColor = System.Drawing.Color.Purple
        Me.txtSMTP.Location = New System.Drawing.Point(132, 266)
        Me.txtSMTP.Name = "txtSMTP"
        Me.txtSMTP.Size = New System.Drawing.Size(349, 20)
        Me.txtSMTP.TabIndex = 22
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 269)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(129, 20)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Serwer SMTP . . . . . . . . . . . "
        '
        'txtAdres
        '
        Me.txtAdres.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAdres.ForeColor = System.Drawing.Color.Purple
        Me.txtAdres.Location = New System.Drawing.Point(132, 242)
        Me.txtAdres.Name = "txtAdres"
        Me.txtAdres.Size = New System.Drawing.Size(626, 20)
        Me.txtAdres.TabIndex = 20
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 245)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(129, 20)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Adres eMail . . . . . . . . . . . "
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Teal
        Me.Label3.Location = New System.Drawing.Point(0, 184)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(766, 20)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Parametry poczty elektronicznej wysy³anej i odbieranej z tego komputera"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 146)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(366, 20)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Typ Bazy Danych obs³ugiwanej przez program . . . . . . . . . . . "
        '
        'txtOpisInstalacji
        '
        Me.txtOpisInstalacji.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOpisInstalacji.ForeColor = System.Drawing.Color.Purple
        Me.txtOpisInstalacji.Location = New System.Drawing.Point(132, 7)
        Me.txtOpisInstalacji.Name = "txtOpisInstalacji"
        Me.txtOpisInstalacji.Size = New System.Drawing.Size(626, 20)
        Me.txtOpisInstalacji.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(140, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Opis instalacji . . . . . . . . . . . "
        '
        'TabBaza
        '
        Me.TabBaza.BackColor = System.Drawing.SystemColors.Control
        Me.TabBaza.Controls.Add(Me.cmdPokazSQLPass)
        Me.TabBaza.Controls.Add(Me.btnNowaBaza)
        Me.TabBaza.Controls.Add(Me.lblIP)
        Me.TabBaza.Controls.Add(Me.lblNazwa)
        Me.TabBaza.Controls.Add(Me.Label15)
        Me.TabBaza.Controls.Add(Me.Label16)
        Me.TabBaza.Controls.Add(Me.cbxSQLBaza)
        Me.TabBaza.Controls.Add(Me.cbxSQLServer)
        Me.TabBaza.Controls.Add(Me.btnTestuj)
        Me.TabBaza.Controls.Add(Me.rbtSQLServer)
        Me.TabBaza.Controls.Add(Me.rbtWindows)
        Me.TabBaza.Controls.Add(Me.Label9)
        Me.TabBaza.Controls.Add(Me.Label11)
        Me.TabBaza.Controls.Add(Me.txtSQLPassword)
        Me.TabBaza.Controls.Add(Me.Label12)
        Me.TabBaza.Controls.Add(Me.txtSQLLogin)
        Me.TabBaza.Controls.Add(Me.Label13)
        Me.TabBaza.Controls.Add(Me.Label14)
        Me.TabBaza.Controls.Add(Me.cmdPobierzKatalog)
        Me.TabBaza.Controls.Add(Me.lblKatalog)
        Me.TabBaza.Controls.Add(Me.lblJakaBaza)
        Me.TabBaza.Controls.Add(Me.Label10)
        Me.TabBaza.Location = New System.Drawing.Point(4, 23)
        Me.TabBaza.Name = "TabBaza"
        Me.TabBaza.Size = New System.Drawing.Size(766, 386)
        Me.TabBaza.TabIndex = 1
        Me.TabBaza.Text = "Baza Danych"
        '
        'cmdPokazSQLPass
        '
        Me.cmdPokazSQLPass.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazSQLPass.Location = New System.Drawing.Point(272, 127)
        Me.cmdPokazSQLPass.Name = "cmdPokazSQLPass"
        Me.cmdPokazSQLPass.Size = New System.Drawing.Size(84, 23)
        Me.cmdPokazSQLPass.TabIndex = 324
        Me.cmdPokazSQLPass.Text = "Poka¿ has³o"
        '
        'btnNowaBaza
        '
        Me.btnNowaBaza.ForeColor = System.Drawing.Color.Navy
        Me.btnNowaBaza.Location = New System.Drawing.Point(16, 210)
        Me.btnNowaBaza.Name = "btnNowaBaza"
        Me.btnNowaBaza.Size = New System.Drawing.Size(340, 23)
        Me.btnNowaBaza.TabIndex = 39
        Me.btnNowaBaza.Text = "Za³ó¿ now¹ Bazê Danych na serwerze SQL"
        '
        'lblIP
        '
        Me.lblIP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblIP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIP.ForeColor = System.Drawing.Color.Purple
        Me.lblIP.Location = New System.Drawing.Point(430, 104)
        Me.lblIP.Name = "lblIP"
        Me.lblIP.Size = New System.Drawing.Size(324, 269)
        Me.lblIP.TabIndex = 38
        '
        'lblNazwa
        '
        Me.lblNazwa.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa.Location = New System.Drawing.Point(430, 76)
        Me.lblNazwa.Name = "lblNazwa"
        Me.lblNazwa.Size = New System.Drawing.Size(324, 16)
        Me.lblNazwa.TabIndex = 37
        '
        'Label15
        '
        Me.Label15.ForeColor = System.Drawing.Color.Navy
        Me.Label15.Location = New System.Drawing.Point(369, 105)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(104, 16)
        Me.Label15.TabIndex = 36
        Me.Label15.Text = "Adres IP . . . . . . . . . . . "
        '
        'Label16
        '
        Me.Label16.ForeColor = System.Drawing.Color.Navy
        Me.Label16.Location = New System.Drawing.Point(369, 76)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(104, 16)
        Me.Label16.TabIndex = 35
        Me.Label16.Text = "Komputer . . . . . . . . . . . "
        '
        'cbxSQLBaza
        '
        Me.cbxSQLBaza.ForeColor = System.Drawing.Color.Purple
        Me.cbxSQLBaza.FormattingEnabled = True
        Me.cbxSQLBaza.Location = New System.Drawing.Point(100, 153)
        Me.cbxSQLBaza.Name = "cbxSQLBaza"
        Me.cbxSQLBaza.Size = New System.Drawing.Size(256, 22)
        Me.cbxSQLBaza.TabIndex = 34
        '
        'cbxSQLServer
        '
        Me.cbxSQLServer.ForeColor = System.Drawing.Color.Purple
        Me.cbxSQLServer.FormattingEnabled = True
        Me.cbxSQLServer.Location = New System.Drawing.Point(100, 74)
        Me.cbxSQLServer.Name = "cbxSQLServer"
        Me.cbxSQLServer.Size = New System.Drawing.Size(256, 22)
        Me.cbxSQLServer.TabIndex = 33
        '
        'btnTestuj
        '
        Me.btnTestuj.ForeColor = System.Drawing.Color.Navy
        Me.btnTestuj.Location = New System.Drawing.Point(16, 181)
        Me.btnTestuj.Name = "btnTestuj"
        Me.btnTestuj.Size = New System.Drawing.Size(340, 23)
        Me.btnTestuj.TabIndex = 32
        Me.btnTestuj.Text = "Testuj po³¹czenie z serwerem SQL"
        '
        'rbtSQLServer
        '
        Me.rbtSQLServer.Location = New System.Drawing.Point(220, 55)
        Me.rbtSQLServer.Name = "rbtSQLServer"
        Me.rbtSQLServer.Size = New System.Drawing.Size(104, 16)
        Me.rbtSQLServer.TabIndex = 31
        Me.rbtSQLServer.Text = "SQL Server"
        '
        'rbtWindows
        '
        Me.rbtWindows.Location = New System.Drawing.Point(100, 55)
        Me.rbtWindows.Name = "rbtWindows"
        Me.rbtWindows.Size = New System.Drawing.Size(104, 16)
        Me.rbtWindows.TabIndex = 30
        Me.rbtWindows.Text = "Windows"
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(13, 55)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(120, 16)
        Me.Label9.TabIndex = 29
        Me.Label9.Text = "Autentykacja . . . . . . . . . . . "
        '
        'Label11
        '
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(13, 156)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(120, 16)
        Me.Label11.TabIndex = 28
        Me.Label11.Text = "Baza Danych . . . . . . . . . . . . . . "
        '
        'txtSQLPassword
        '
        Me.txtSQLPassword.ForeColor = System.Drawing.Color.Purple
        Me.txtSQLPassword.Location = New System.Drawing.Point(100, 128)
        Me.txtSQLPassword.Name = "txtSQLPassword"
        Me.txtSQLPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSQLPassword.Size = New System.Drawing.Size(162, 20)
        Me.txtSQLPassword.TabIndex = 25
        Me.txtSQLPassword.Text = "TextBox1"
        '
        'Label12
        '
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(13, 131)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(120, 16)
        Me.Label12.TabIndex = 27
        Me.Label12.Text = "Password . . . . . . . . . . . . . . "
        '
        'txtSQLLogin
        '
        Me.txtSQLLogin.ForeColor = System.Drawing.Color.Purple
        Me.txtSQLLogin.Location = New System.Drawing.Point(100, 102)
        Me.txtSQLLogin.Name = "txtSQLLogin"
        Me.txtSQLLogin.Size = New System.Drawing.Size(256, 20)
        Me.txtSQLLogin.TabIndex = 24
        '
        'Label13
        '
        Me.Label13.ForeColor = System.Drawing.Color.Navy
        Me.Label13.Location = New System.Drawing.Point(13, 105)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(120, 16)
        Me.Label13.TabIndex = 26
        Me.Label13.Text = "Login . . . . . . . . . . . . . . .. . . ."
        '
        'Label14
        '
        Me.Label14.ForeColor = System.Drawing.Color.Navy
        Me.Label14.Location = New System.Drawing.Point(13, 77)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(120, 16)
        Me.Label14.TabIndex = 23
        Me.Label14.Text = "Serwer SQL . . . . . . . . . . . "
        '
        'cmdPobierzKatalog
        '
        Me.cmdPobierzKatalog.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPobierzKatalog.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPobierzKatalog.Image = CType(resources.GetObject("cmdPobierzKatalog.Image"), System.Drawing.Image)
        Me.cmdPobierzKatalog.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdPobierzKatalog.Location = New System.Drawing.Point(714, 23)
        Me.cmdPobierzKatalog.Name = "cmdPobierzKatalog"
        Me.cmdPobierzKatalog.Size = New System.Drawing.Size(40, 24)
        Me.cmdPobierzKatalog.TabIndex = 19
        Me.cmdPobierzKatalog.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblKatalog
        '
        Me.lblKatalog.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblKatalog.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalog.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalog.Location = New System.Drawing.Point(125, 27)
        Me.lblKatalog.Name = "lblKatalog"
        Me.lblKatalog.Size = New System.Drawing.Size(618, 17)
        Me.lblKatalog.TabIndex = 20
        '
        'lblJakaBaza
        '
        Me.lblJakaBaza.ForeColor = System.Drawing.Color.Navy
        Me.lblJakaBaza.Location = New System.Drawing.Point(13, 8)
        Me.lblJakaBaza.Name = "lblJakaBaza"
        Me.lblJakaBaza.Size = New System.Drawing.Size(536, 20)
        Me.lblJakaBaza.TabIndex = 21
        Me.lblJakaBaza.Text = "Program obs³uguje baze danych MS ACCESS (.MDB)"
        '
        'Label10
        '
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(13, 28)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(120, 20)
        Me.Label10.TabIndex = 18
        Me.Label10.Text = "Katalog Bazy Danych . . . . . . . . . . . "
        '
        'TabBazaImport
        '
        Me.TabBazaImport.BackColor = System.Drawing.SystemColors.Control
        Me.TabBazaImport.Controls.Add(Me.txtPZnakIdPracownika)
        Me.TabBazaImport.Controls.Add(Me.Label35)
        Me.TabBazaImport.Controls.Add(Me.txtPZnakIdKlienta)
        Me.TabBazaImport.Controls.Add(Me.Label34)
        Me.TabBazaImport.Controls.Add(Me.Label33)
        Me.TabBazaImport.Controls.Add(Me.Label32)
        Me.TabBazaImport.Controls.Add(Me.txtHQWersjaWymagana)
        Me.TabBazaImport.Controls.Add(Me.Label24)
        Me.TabBazaImport.Controls.Add(Me.txtHQWersja)
        Me.TabBazaImport.Controls.Add(Me.Label23)
        Me.TabBazaImport.Controls.Add(Me.cmdHQWersjaBazyDanych)
        Me.TabBazaImport.Controls.Add(Me.cmdHQPobierz)
        Me.TabBazaImport.Controls.Add(Me.lblHQIP)
        Me.TabBazaImport.Controls.Add(Me.lblHQNazwa)
        Me.TabBazaImport.Controls.Add(Me.cbxHQSQLBaza)
        Me.TabBazaImport.Controls.Add(Me.txtHQSQLPassword)
        Me.TabBazaImport.Controls.Add(Me.txtHQSQLLogin)
        Me.TabBazaImport.Controls.Add(Me.cbxHQSQLServer)
        Me.TabBazaImport.Controls.Add(Me.rbtHQWindows)
        Me.TabBazaImport.Controls.Add(Me.Label83)
        Me.TabBazaImport.Controls.Add(Me.cmdHQSQLPokazHaslo)
        Me.TabBazaImport.Controls.Add(Me.Label25)
        Me.TabBazaImport.Controls.Add(Me.Label26)
        Me.TabBazaImport.Controls.Add(Me.cmdHQTestuj)
        Me.TabBazaImport.Controls.Add(Me.rbtHQSQLServer)
        Me.TabBazaImport.Controls.Add(Me.Label27)
        Me.TabBazaImport.Controls.Add(Me.Label28)
        Me.TabBazaImport.Controls.Add(Me.Label29)
        Me.TabBazaImport.Controls.Add(Me.Label30)
        Me.TabBazaImport.Controls.Add(Me.Label31)
        Me.TabBazaImport.Location = New System.Drawing.Point(4, 23)
        Me.TabBazaImport.Name = "TabBazaImport"
        Me.TabBazaImport.Size = New System.Drawing.Size(766, 386)
        Me.TabBazaImport.TabIndex = 2
        Me.TabBazaImport.Text = "Baza Danych Importowanych"
        '
        'txtPZnakIdPracownika
        '
        Me.txtPZnakIdPracownika.ForeColor = System.Drawing.Color.Purple
        Me.txtPZnakIdPracownika.Location = New System.Drawing.Point(707, 301)
        Me.txtPZnakIdPracownika.Name = "txtPZnakIdPracownika"
        Me.txtPZnakIdPracownika.Size = New System.Drawing.Size(45, 20)
        Me.txtPZnakIdPracownika.TabIndex = 392
        Me.txtPZnakIdPracownika.Visible = False
        '
        'Label35
        '
        Me.Label35.ForeColor = System.Drawing.Color.Navy
        Me.Label35.Location = New System.Drawing.Point(412, 304)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(340, 16)
        Me.Label35.TabIndex = 393
        Me.Label35.Text = "Pierwszy znak ID PRACOWNIKA z Bazy Importowanej . . . . . . . . . . . . . . "
        Me.Label35.Visible = False
        '
        'txtPZnakIdKlienta
        '
        Me.txtPZnakIdKlienta.ForeColor = System.Drawing.Color.Purple
        Me.txtPZnakIdKlienta.Location = New System.Drawing.Point(312, 301)
        Me.txtPZnakIdKlienta.Name = "txtPZnakIdKlienta"
        Me.txtPZnakIdKlienta.Size = New System.Drawing.Size(45, 20)
        Me.txtPZnakIdKlienta.TabIndex = 390
        '
        'Label34
        '
        Me.Label34.ForeColor = System.Drawing.Color.Navy
        Me.Label34.Location = New System.Drawing.Point(17, 304)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(340, 16)
        Me.Label34.TabIndex = 391
        Me.Label34.Text = "Pierwszy znak NRU KARTY KLIENTA z Bazy Importowanej . . . . . . . . . . . . . . "
        '
        'Label33
        '
        Me.Label33.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label33.ForeColor = System.Drawing.Color.Navy
        Me.Label33.Location = New System.Drawing.Point(409, 251)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(343, 57)
        Me.Label33.TabIndex = 389
        Me.Label33.Text = "Pierwszy znak w IDENTYFIKATORZE PRACOWNIKA dodawany w identyfikatorach z Bazy Imp" &
    "ortowanej, aby zapewniæ unikalnoœæ identyfikatorów pracowników w Bazie Wynikowej" &
    " "
        Me.Label33.Visible = False
        '
        'Label32
        '
        Me.Label32.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label32.ForeColor = System.Drawing.Color.Navy
        Me.Label32.Location = New System.Drawing.Point(14, 251)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(343, 57)
        Me.Label32.TabIndex = 388
        Me.Label32.Text = "Pierwszy znak w ID / NUMERZE KARTY KLIENTA dodawany w identyfikatorach z Bazy Imp" &
    "ortowanej, aby zapewniæ unikalnoœæ numeru karty w Bazie Wynikowej "
        '
        'txtHQWersjaWymagana
        '
        Me.txtHQWersjaWymagana.ForeColor = System.Drawing.Color.Purple
        Me.txtHQWersjaWymagana.Location = New System.Drawing.Point(297, 155)
        Me.txtHQWersjaWymagana.Name = "txtHQWersjaWymagana"
        Me.txtHQWersjaWymagana.ReadOnly = True
        Me.txtHQWersjaWymagana.Size = New System.Drawing.Size(60, 20)
        Me.txtHQWersjaWymagana.TabIndex = 386
        Me.txtHQWersjaWymagana.Text = "????"
        '
        'Label24
        '
        Me.Label24.ForeColor = System.Drawing.Color.Navy
        Me.Label24.Location = New System.Drawing.Point(186, 158)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(120, 16)
        Me.Label24.TabIndex = 387
        Me.Label24.Text = "Wymagana Wersja . . . . . . . . . . . . . . "
        '
        'txtHQWersja
        '
        Me.txtHQWersja.ForeColor = System.Drawing.Color.Purple
        Me.txtHQWersja.Location = New System.Drawing.Point(101, 155)
        Me.txtHQWersja.Name = "txtHQWersja"
        Me.txtHQWersja.ReadOnly = True
        Me.txtHQWersja.Size = New System.Drawing.Size(60, 20)
        Me.txtHQWersja.TabIndex = 384
        Me.txtHQWersja.Text = "????"
        '
        'Label23
        '
        Me.Label23.ForeColor = System.Drawing.Color.Navy
        Me.Label23.Location = New System.Drawing.Point(14, 158)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(120, 16)
        Me.Label23.TabIndex = 385
        Me.Label23.Text = "Wersja Bazy . . . . . . . . . . . . . . "
        '
        'cmdHQWersjaBazyDanych
        '
        Me.cmdHQWersjaBazyDanych.ForeColor = System.Drawing.Color.Navy
        Me.cmdHQWersjaBazyDanych.Location = New System.Drawing.Point(17, 210)
        Me.cmdHQWersjaBazyDanych.Name = "cmdHQWersjaBazyDanych"
        Me.cmdHQWersjaBazyDanych.Size = New System.Drawing.Size(340, 23)
        Me.cmdHQWersjaBazyDanych.TabIndex = 383
        Me.cmdHQWersjaBazyDanych.Text = "Aktualizuj wersje Bazy Danych Importowanych"
        '
        'cmdHQPobierz
        '
        Me.cmdHQPobierz.ForeColor = System.Drawing.Color.Navy
        Me.cmdHQPobierz.Location = New System.Drawing.Point(17, 336)
        Me.cmdHQPobierz.Name = "cmdHQPobierz"
        Me.cmdHQPobierz.Size = New System.Drawing.Size(735, 23)
        Me.cmdHQPobierz.TabIndex = 382
        Me.cmdHQPobierz.Text = "Pobierz dane o klientach z Bazy Danych Importowanych i dopisz do aktualnej Bazy D" &
    "anych"
        '
        'lblHQIP
        '
        Me.lblHQIP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblHQIP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblHQIP.ForeColor = System.Drawing.Color.Purple
        Me.lblHQIP.Location = New System.Drawing.Point(431, 78)
        Me.lblHQIP.Name = "lblHQIP"
        Me.lblHQIP.Size = New System.Drawing.Size(323, 155)
        Me.lblHQIP.TabIndex = 378
        '
        'lblHQNazwa
        '
        Me.lblHQNazwa.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblHQNazwa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblHQNazwa.ForeColor = System.Drawing.Color.Purple
        Me.lblHQNazwa.Location = New System.Drawing.Point(431, 50)
        Me.lblHQNazwa.Name = "lblHQNazwa"
        Me.lblHQNazwa.Size = New System.Drawing.Size(321, 16)
        Me.lblHQNazwa.TabIndex = 377
        '
        'cbxHQSQLBaza
        '
        Me.cbxHQSQLBaza.ForeColor = System.Drawing.Color.Purple
        Me.cbxHQSQLBaza.FormattingEnabled = True
        Me.cbxHQSQLBaza.Location = New System.Drawing.Point(101, 127)
        Me.cbxHQSQLBaza.Name = "cbxHQSQLBaza"
        Me.cbxHQSQLBaza.Size = New System.Drawing.Size(256, 22)
        Me.cbxHQSQLBaza.TabIndex = 374
        '
        'txtHQSQLPassword
        '
        Me.txtHQSQLPassword.ForeColor = System.Drawing.Color.Purple
        Me.txtHQSQLPassword.Location = New System.Drawing.Point(101, 102)
        Me.txtHQSQLPassword.Name = "txtHQSQLPassword"
        Me.txtHQSQLPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtHQSQLPassword.Size = New System.Drawing.Size(166, 20)
        Me.txtHQSQLPassword.TabIndex = 365
        Me.txtHQSQLPassword.Text = "xxx"
        '
        'txtHQSQLLogin
        '
        Me.txtHQSQLLogin.ForeColor = System.Drawing.Color.Purple
        Me.txtHQSQLLogin.Location = New System.Drawing.Point(101, 76)
        Me.txtHQSQLLogin.Name = "txtHQSQLLogin"
        Me.txtHQSQLLogin.Size = New System.Drawing.Size(256, 20)
        Me.txtHQSQLLogin.TabIndex = 364
        '
        'cbxHQSQLServer
        '
        Me.cbxHQSQLServer.ForeColor = System.Drawing.Color.Purple
        Me.cbxHQSQLServer.FormattingEnabled = True
        Me.cbxHQSQLServer.Location = New System.Drawing.Point(101, 48)
        Me.cbxHQSQLServer.Name = "cbxHQSQLServer"
        Me.cbxHQSQLServer.Size = New System.Drawing.Size(256, 22)
        Me.cbxHQSQLServer.TabIndex = 373
        '
        'rbtHQWindows
        '
        Me.rbtHQWindows.Location = New System.Drawing.Point(101, 29)
        Me.rbtHQWindows.Name = "rbtHQWindows"
        Me.rbtHQWindows.Size = New System.Drawing.Size(104, 16)
        Me.rbtHQWindows.TabIndex = 370
        Me.rbtHQWindows.Text = "Windows"
        '
        'Label83
        '
        Me.Label83.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label83.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label83.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label83.ForeColor = System.Drawing.Color.Teal
        Me.Label83.Location = New System.Drawing.Point(0, 0)
        Me.Label83.Name = "Label83"
        Me.Label83.Size = New System.Drawing.Size(765, 16)
        Me.Label83.TabIndex = 380
        Me.Label83.Text = "Parametry dostêpu do Bazy Danych z której bêd¹ importowane dane klientów "
        Me.Label83.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdHQSQLPokazHaslo
        '
        Me.cmdHQSQLPokazHaslo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdHQSQLPokazHaslo.Location = New System.Drawing.Point(273, 101)
        Me.cmdHQSQLPokazHaslo.Name = "cmdHQSQLPokazHaslo"
        Me.cmdHQSQLPokazHaslo.Size = New System.Drawing.Size(84, 23)
        Me.cmdHQSQLPokazHaslo.TabIndex = 379
        Me.cmdHQSQLPokazHaslo.Text = "Poka¿ has³o"
        '
        'Label25
        '
        Me.Label25.ForeColor = System.Drawing.Color.Navy
        Me.Label25.Location = New System.Drawing.Point(370, 79)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(104, 16)
        Me.Label25.TabIndex = 376
        Me.Label25.Text = "Adresy IP . . . . . . . . . . . "
        '
        'Label26
        '
        Me.Label26.ForeColor = System.Drawing.Color.Navy
        Me.Label26.Location = New System.Drawing.Point(370, 50)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(104, 16)
        Me.Label26.TabIndex = 375
        Me.Label26.Text = "Komputer . . . . . . . . . . . "
        '
        'cmdHQTestuj
        '
        Me.cmdHQTestuj.ForeColor = System.Drawing.Color.Navy
        Me.cmdHQTestuj.Location = New System.Drawing.Point(17, 181)
        Me.cmdHQTestuj.Name = "cmdHQTestuj"
        Me.cmdHQTestuj.Size = New System.Drawing.Size(340, 23)
        Me.cmdHQTestuj.TabIndex = 372
        Me.cmdHQTestuj.Text = "Testuj po³¹czenie z SQL Server"
        '
        'rbtHQSQLServer
        '
        Me.rbtHQSQLServer.Location = New System.Drawing.Point(221, 29)
        Me.rbtHQSQLServer.Name = "rbtHQSQLServer"
        Me.rbtHQSQLServer.Size = New System.Drawing.Size(104, 16)
        Me.rbtHQSQLServer.TabIndex = 371
        Me.rbtHQSQLServer.Text = "SQL Server"
        '
        'Label27
        '
        Me.Label27.ForeColor = System.Drawing.Color.Navy
        Me.Label27.Location = New System.Drawing.Point(14, 29)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(120, 16)
        Me.Label27.TabIndex = 369
        Me.Label27.Text = "Autentykacja . . . . . . . . . . . "
        '
        'Label28
        '
        Me.Label28.ForeColor = System.Drawing.Color.Navy
        Me.Label28.Location = New System.Drawing.Point(14, 130)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(120, 16)
        Me.Label28.TabIndex = 368
        Me.Label28.Text = "Baza Danych . . . . . . . . . . . . . . "
        '
        'Label29
        '
        Me.Label29.ForeColor = System.Drawing.Color.Navy
        Me.Label29.Location = New System.Drawing.Point(14, 105)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(120, 16)
        Me.Label29.TabIndex = 367
        Me.Label29.Text = "Password . . . . . . . . . . . . . . "
        '
        'Label30
        '
        Me.Label30.ForeColor = System.Drawing.Color.Navy
        Me.Label30.Location = New System.Drawing.Point(14, 79)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(120, 16)
        Me.Label30.TabIndex = 366
        Me.Label30.Text = "Login . . . . . . . . . . . . . . .. . . ."
        '
        'Label31
        '
        Me.Label31.ForeColor = System.Drawing.Color.Navy
        Me.Label31.Location = New System.Drawing.Point(14, 51)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(120, 16)
        Me.Label31.TabIndex = 363
        Me.Label31.Text = "Serwer SQL . . . . . . . . . . . "
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(702, 429)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(75, 23)
        Me.cmdPowrot.TabIndex = 2
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdZapisz
        '
        Me.cmdZapisz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZapisz.Image = CType(resources.GetObject("cmdZapisz.Image"), System.Drawing.Image)
        Me.cmdZapisz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZapisz.Location = New System.Drawing.Point(617, 429)
        Me.cmdZapisz.Name = "cmdZapisz"
        Me.cmdZapisz.Size = New System.Drawing.Size(79, 23)
        Me.cmdZapisz.TabIndex = 3
        Me.cmdZapisz.Text = "Zapisz"
        Me.cmdZapisz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOdswiez
        '
        Me.cmdOdswiez.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOdswiez.Image = CType(resources.GetObject("cmdOdswiez.Image"), System.Drawing.Image)
        Me.cmdOdswiez.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOdswiez.Location = New System.Drawing.Point(520, 429)
        Me.cmdOdswiez.Name = "cmdOdswiez"
        Me.cmdOdswiez.Size = New System.Drawing.Size(91, 23)
        Me.cmdOdswiez.TabIndex = 4
        Me.cmdOdswiez.Text = "Przywróæ"
        Me.cmdOdswiez.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 427)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 12
        Me.PictureBox2.TabStop = False
        '
        'ParametryLokalne
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(790, 457)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdOdswiez)
        Me.Controls.Add(Me.cmdZapisz)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.tbcParametry)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ParametryLokalne"
        Me.ShowInTaskbar = False
        Me.Text = "Parametry lokalne"
        Me.tbcParametry.ResumeLayout(False)
        Me.Instalacja.ResumeLayout(False)
        Me.Instalacja.PerformLayout()
        Me.TabBaza.ResumeLayout(False)
        Me.TabBaza.PerformLayout()
        Me.TabBazaImport.ResumeLayout(False)
        Me.TabBazaImport.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ParametryLokalne_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        'cbxBazaDanych.Items.Add("MS ACCESS")
        cbxBazaDanych.Items.Add("MS SQL")

        cbxeMailService.Items.Add("Wbudowany")
        cbxeMailService.Items.Add("MS Outlook")
        cbxeMailService.Items.Add("Mozilla Thunderbird")
        'PobierzParametry()
        '-----------------
        If _MaUprawnieniaAdministratora() Then
            'POKAZUJ zak³adke "Baza Danych do Importu"
        Else
            'nie pokazuj zak³adki "Baza Danych do Importu"
            tbcParametry.TabPages.RemoveAt(2)
        End If
        '-----------------
        Dim HostName As String = System.Net.Dns.GetHostName()
        'Dim IPAddress As String = System.Net.Dns.GetHostEntry(HostName).AddressList(0).ToString()
        Dim IPAddress As String = ""
        Dim ii As Integer = 0
        For ii = 0 To System.Net.Dns.GetHostEntry(HostName).AddressList.Length - 1
            IPAddress &= System.Net.Dns.GetHostEntry(HostName).AddressList(ii).ToString() & vbCrLf
        Next
        lblIP.Text = IPAddress
        lblNazwa.Text = HostName
        lblHQIP.Text = IPAddress
        lblHQNazwa.Text = HostName
        ''-------------
        ''WER.1 pobierz liste serverów SQL
        Dim instance As System.Data.Sql.SqlDataSourceEnumerator = System.Data.Sql.SqlDataSourceEnumerator.Instance
        Try
            Dim table As System.Data.DataTable = instance.GetDataSources()

            cbxSQLServer.Items.Add("(local)")
            cbxHQSQLServer.Items.Add("(local)")
            For Each row As DataRow In table.Rows
                'row(0) - nazwa serwera
                'row(1) - nazwa instancji
                'row(2) - clustered/unclastered
                'row(3) - nr wersji
                If IsDBNull(row(1)) Then
                    'brak nazwy instancji
                    cbxSQLServer.Items.Add(row(0))
                    cbxHQSQLServer.Items.Add(row(0))
                Else
                    If Len(Trim(row(1))) = 0 Then
                        'brak nazwy instancji
                        cbxSQLServer.Items.Add(row(0))
                        cbxHQSQLServer.Items.Add(row(0))
                    Else
                        cbxHQSQLServer.Items.Add(row(0) & "\" & row(1))
                        cbxSQLServer.Items.Add(row(0) & "\" & row(1))
                    End If
                End If
            Next
            cbxSQLServer.SelectedIndex = 0
            cbxHQSQLServer.SelectedIndex = 0

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        '-------------
        '*  WER. pobierz liste serverów SQL
        '*  dla SQL 2000 - dodac referencje do Microsoft SQLDMO Object Library 8.0. 
        '*  dla SQL 2005 - to wszystko robi SMO
        'Dim sqlApp As New Sqldmo.Application
        'Dim NL As Sqldmo.NameList
        'Dim Index As Integer
        'NL = sqlApp.ListAvailableSQLServers
        'For Index = 1 To NL.Count
        '    Me.cbxServer.Items.Add(NL.Item(Index))
        'Next



        '---------
        ' Get the connected status.
        'Dim dwFlags As Long
        'Dim Connected As Boolean = CBool(InternetGetConnectedState(dwFlags, 0&))
        'If Connected Then
        '    ' Display all connection flags.
        '    Dim ConnectionType As ConnectStates
        '    For Each ConnectionType In System.Enum.GetValues(GetType(ConnectStates))
        '        If (ConnectionType And dwFlags) = ConnectionType Then
        '            'lblSposobPolaczenia.Text = ConnectionType.ToString()
        '            Exit For
        '        End If
        '    Next

        '    'Dim HostName As String
        '    'Dim IPAddress As String
        '    'HostName = System.Net.Dns.GetHostName()
        '    'IPAddress = System.Net.Dns.GetHostEntry(HostName).AddressList(0).ToString()
        '    'lblIP.Text = IPAddress
        '    'lblNazwa.Text = HostName
        'Else
        '    lblIP.Text = "Brak po³¹czenia"
        'End If

        '------------------------
        PobierzParametry()
    End Sub

    Private Sub ParametryLokalne_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    '*****************************************
    '** obsluga komend
    '*****************************************

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        ZapiszParametry()
        If cbxBazaDanych.Text = "MS ACCESS" Then
            If SrodowiskoOK() Then
                Me.Close()
            Else
                End
            End If
        Else
            Me.Close()
        End If
    End Sub

    Private Sub cmdPobierzKatalog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPobierzKatalog.Click
        With dlgOpenFolder
            .Description = "Proszê wybraæ katalog dyskowy w którym umieszczono pliki .MDB z danymi programu Magazyn..."
            '.RootFolder = "c:\"
            .ShowNewFolderButton = True

            'Me.Visible = False
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                lblKatalog.Text = .SelectedPath
            End If
        End With
    End Sub


    Private Sub cmdPobierzKatalogArchiwum_Click(sender As Object, e As EventArgs) Handles cmdPobierzKatalogArchiwum.Click
        With dlgOpenFolder
            .Description = "Proszê wybraæ katalog dyskowy w którym zapisywany bedzia plik Archiwum..."
            '.RootFolder = "c:\"
            .ShowNewFolderButton = True

            'Me.Visible = False
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                txtKatalogArchiwum.Text = .SelectedPath
            End If
        End With
    End Sub

    Private Sub cmdPobierzKatalogArchZIP_Click(sender As Object, e As EventArgs) Handles cmdPobierzKatalogArchZIP.Click
        With dlgOpenFolder
            .Description = "Proszê wybraæ katalog dyskowy w którym zapisywany bedzia plik Archiwum ZIP..."
            '.RootFolder = "c:\"
            .ShowNewFolderButton = True

            'Me.Visible = False
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                txtKatalogArchiwumZIP.Text = .SelectedPath
            End If
        End With
    End Sub

    Private Sub cmdPobierzKatalogRaporty_Click(sender As Object, e As EventArgs) Handles cmdPobierzKatalogRaporty.Click
        With dlgOpenFolder
            .Description = "Proszê wybraæ katalog dyskowy w którym zapisywane bêd¹ Raporty Excel."
            '.RootFolder = "c:\"
            .ShowNewFolderButton = True

            'Me.Visible = False
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                txtKatalogRaporty.Text = .SelectedPath
            End If
        End With
    End Sub


    Private Sub cmdOdswiez_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOdswiez.Click
        PobierzParametry()
    End Sub

    Private Sub cmdZapisz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZapisz.Click
        ZapiszParametry()
    End Sub

    Private Sub btnTestuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestuj.Click
        ZapiszParametry()
        'sprawdz polaczenie z Baza Danych
        'Dim forma As New TestujPolaczenieZBazaDanych
        Dim forma As New _frmTestujPolaczenieZBazaDanych(ParL_Autentykacja,
                                                         ParL_Serwer,
                                                         ParL_Login,
                                                         ParL_Password,
                                                         ParL_DataBase)
        System.Windows.Forms.Application.DoEvents()
        forma.Show()
    End Sub

    Private Sub cmdHQTestuj_Click(sender As Object, e As EventArgs) Handles cmdHQTestuj.Click
        ZapiszParametry()
        'sprawdz polaczenie z Baza Danych
        Dim forma As New _frmTestujPolaczenieZBazaDanych(ParL_HQAutentykacja,
                                                         ParL_HQSerwer,
                                                         ParL_HQLogin,
                                                         ParL_HQPassword,
                                                         ParL_HQDataBase)
        System.Windows.Forms.Application.DoEvents()
        forma.Show()
    End Sub


    '*****************************************
    '** edycja parametrow
    '*****************************************

    Private Sub txtOpisInstalacji_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpisInstalacji.GotFocus
        StartEdycjiTxt(txtOpisInstalacji)
    End Sub
    Private Sub txtOpisInstalacji_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpisInstalacji.LostFocus
        KoniecEdycjiTxt(txtOpisInstalacji)
    End Sub

    Private Sub txtNazwaArchiwum_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNazwaArchiwum.GotFocus
        StartEdycjiTxt(txtNazwaArchiwum)
    End Sub
    Private Sub txtNazwaArchiwum_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNazwaArchiwum.LostFocus
        KoniecEdycjiTxt(txtNazwaArchiwum)
    End Sub

    Private Sub txtKatalogArchiwum_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKatalogArchiwum.GotFocus
        StartEdycjiTxt(txtKatalogArchiwum)
    End Sub
    Private Sub txtKatalogArchiwum_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKatalogArchiwum.LostFocus
        KoniecEdycjiTxt(txtKatalogArchiwum)
    End Sub

    Private Sub txtKatalogArchiwumZIP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKatalogArchiwumZIP.GotFocus
        StartEdycjiTxt(txtKatalogArchiwumZIP)
    End Sub
    Private Sub txtKatalogArchiwumZIP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKatalogArchiwumZIP.LostFocus
        KoniecEdycjiTxt(txtKatalogArchiwumZIP)
    End Sub

    Private Sub txtKatalogRaporty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKatalogRaporty.GotFocus
        StartEdycjiTxt(txtKatalogRaporty)
    End Sub
    Private Sub txtKatalogRaporty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKatalogRaporty.LostFocus
        KoniecEdycjiTxt(txtKatalogRaporty)
    End Sub


    Private Sub cbxeMailService_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxeMailService.GotFocus
        StartEdycjiCbx(cbxeMailService)
    End Sub
    Private Sub cbxeMailService_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxeMailService.LostFocus
        KoniecEdycjiCbx(cbxeMailService)
    End Sub

    Private Sub txtAdres_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdres.GotFocus
        StartEdycjiTxt(txtAdres)
    End Sub
    Private Sub txtAdres_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdres.LostFocus
        KoniecEdycjiTxt(txtAdres)
        '-------------------------
        If String.IsNullOrEmpty(Trim(txtAdres.Text)) Then
            'adres eMail niezdef - jesk OK
        Else
            'adres eMail zdef - czy jest poprawny
            If Not IsValidEmail(Trim(txtAdres.Text)) Then
                MessageBox.Show("Adres email nie jest poprawny" & vbCrLf &
                                Trim(txtAdres.Text),
                                "Uwaga.",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation)
                txtAdres.Focus()
            Else
            End If
        End If
    End Sub

    Private Sub txtSMTP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSMTP.GotFocus
        StartEdycjiTxt(txtSMTP)
    End Sub
    Private Sub txtSMTP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSMTP.LostFocus
        KoniecEdycjiTxt(txtSMTP)
    End Sub

    Private Sub txtSMTPPort_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSMTPPort.GotFocus
        StartEdycjiTxt(txtSMTPPort)
    End Sub
    Private Sub txtSMTPPort_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSMTPPort.LostFocus
        KoniecEdycjiTxt(txtSMTPPort)
    End Sub

    Private Sub txtPOP3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPOP3.GotFocus
        StartEdycjiTxt(txtPOP3)
    End Sub
    Private Sub txtPOP3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPOP3.LostFocus
        KoniecEdycjiTxt(txtPOP3)
    End Sub

    Private Sub txtPOP3Port_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPOP3Port.GotFocus
        StartEdycjiTxt(txtPOP3Port)
    End Sub
    Private Sub txtPOP3Port_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPOP3Port.LostFocus
        KoniecEdycjiTxt(txtPOP3Port)
    End Sub

    Private Sub txtLogin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLogin.GotFocus
        StartEdycjiTxt(txtLogin)
    End Sub
    Private Sub txtLogin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLogin.LostFocus
        KoniecEdycjiTxt(txtLogin)
    End Sub

    Private Sub txtPass_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPass.GotFocus
        StartEdycjiTxt(txtPass)
    End Sub
    Private Sub txtPass_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPass.LostFocus
        KoniecEdycjiTxt(txtPass)
    End Sub

    Private Sub cbxSQLServer_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSQLServer.GotFocus
        StartEdycjiCbx(cbxSQLServer)
    End Sub
    Private Sub cbxSQLServer_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSQLServer.LostFocus
        KoniecEdycjiCbx(cbxSQLServer)
    End Sub

    Private Sub cbxSQLBaza_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSQLBaza.GotFocus
        StartEdycjiCbx(cbxSQLBaza)
        '---------------------------------
        PobierzListeBazDanych()
    End Sub
    Private Sub cbxSQLBaza_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSQLBaza.LostFocus
        KoniecEdycjiCbx(cbxSQLBaza)
    End Sub

    Private Sub txtSQLLogin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSQLLogin.GotFocus
        StartEdycjiTxt(txtSQLLogin)
    End Sub
    Private Sub txtSQLLogin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSQLLogin.LostFocus
        KoniecEdycjiTxt(txtSQLLogin)
    End Sub

    Private Sub txtSQLPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSQLPassword.GotFocus
        StartEdycjiTxt(txtSQLPassword)
    End Sub
    Private Sub txtSQLPassword_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSQLPassword.LostFocus
        KoniecEdycjiTxt(txtSQLPassword)
    End Sub
    '---------------------------------------------


    Private Sub cbxHQSQLServer_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxHQSQLServer.GotFocus
        StartEdycjiCbx(cbxHQSQLServer)
    End Sub
    Private Sub cbxHQSQLServer_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxHQSQLServer.LostFocus
        KoniecEdycjiCbx(cbxHQSQLServer)
    End Sub

    Private Sub cbxHQSQLBaza_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxHQSQLBaza.GotFocus
        StartEdycjiCbx(cbxHQSQLBaza)
        '---------------------------------
        HQPobierzListeBazDanych()
    End Sub
    Private Sub cbxHQSQLBaza_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxHQSQLBaza.LostFocus
        KoniecEdycjiCbx(cbxHQSQLBaza)
        '----------------------
        ParL_HQDataBase = cbxHQSQLBaza.Text
        oWersjaBazyDanych = HQWersjaBazyDanych()
        If oWersjaBazyDanych = 0 Then
            txtHQWersja.Text = "????"
        Else
            txtHQWersja.Text = (oWersjaBazyDanych / 100).ToString("N2")
        End If
    End Sub

    Private Sub txtHQSQLLogin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHQSQLLogin.GotFocus
        StartEdycjiTxt(txtHQSQLLogin)
    End Sub
    Private Sub txtHQSQLLogin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHQSQLLogin.LostFocus
        KoniecEdycjiTxt(txtHQSQLLogin)
    End Sub

    Private Sub txtHQSQLPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHQSQLPassword.GotFocus
        StartEdycjiTxt(txtHQSQLPassword)
    End Sub
    Private Sub txtHQSQLPassword_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHQSQLPassword.LostFocus
        KoniecEdycjiTxt(txtHQSQLPassword)
    End Sub
    '---------------------------------------------




    Private Sub PobierzParametry()
        PobierzParametryLokalne()
        '----------------------------------------
        txtOpisInstalacji.Text = ParL_OpisInstalacji
        txtNazwaArchiwum.Text = ParL_NazwaArchiwum
        txtKatalogArchiwum.Text = ParL_KatalogArchiwum
        txtKatalogArchiwumZIP.Text = ParL_KatalogArchiwumZIP
        txtKatalogRaporty.Text = ParL_KatalogRaporty


        Dim Jest As Boolean = False
        Dim i As Integer = 0

        'typ bazy danych
        For i = 0 To cbxBazaDanych.Items.Count() - 1
            If ParL_DataBaseType = cbxBazaDanych.Items.Item(i) Then
                cbxBazaDanych.Select(i, 1)
                cbxBazaDanych.Text = ParL_DataBaseType
                Jest = True
                Exit For
            End If
        Next

        lblKatalog.Text = ParL_KatZbiorow

        If ParL_Autentykacja = "Windows" Then
            rbtWindows.Checked = True
            cbxSQLServer.Enabled = True
            txtSQLLogin.ReadOnly = True
            txtSQLPassword.ReadOnly = True
            cbxSQLBaza.Enabled = True

            ParL_Login = ""
            ParL_Password = ""
        Else
            rbtSQLServer.Checked = True
            cbxSQLServer.Enabled = True
            txtSQLLogin.ReadOnly = False
            txtSQLPassword.ReadOnly = False
            cbxSQLBaza.Enabled = True
        End If

        cbxSQLServer.Text = ParL_Serwer
        For i = 0 To cbxSQLServer.Items.Count() - 1
            If ParL_Serwer = cbxSQLServer.Items.Item(i) Then
                'cbxSQLServer.Select(0, 1)
                cbxSQLServer.Text = ParL_Serwer
                Jest = True
                Exit For
            End If
        Next
        If Not Jest Then
            'nie znalaz³ takiej bazy - dopisz na koniec
            cbxSQLServer.Items.Add(ParL_Serwer)
            cbxSQLServer.Text = ParL_Serwer
        End If

        txtSQLLogin.Text = ParL_Login
        txtSQLPassword.Text = ParL_Password
        cbxSQLBaza.Text = ParL_DataBase







        'For i = 0 To cbxHQBazaDanych.Items.Count() - 1
        '    If ParL_HQDataBaseType = cbxHQBazaDanych.Items.Item(i) Then
        '        'cbxBazaDanych.Select(i, 1)
        '        cbxHQBazaDanych.Text = ParL_HQDataBaseType
        '        Exit For
        '    End If
        'Next

        'lblHQPlikMDB.Text = ParL_HQPlikMDB

        If ParL_Autentykacja = "Windows" Then
            rbtHQWindows.Checked = True
            cbxHQSQLServer.Enabled = True
            txtHQSQLLogin.ReadOnly = True
            txtHQSQLPassword.ReadOnly = True
            cbxHQSQLBaza.Enabled = True

            ParL_HQLogin = ""
            ParL_HQPassword = ""
        Else
            rbtHQSQLServer.Checked = True
            cbxHQSQLServer.Enabled = True
            txtHQSQLLogin.ReadOnly = False
            txtHQSQLPassword.ReadOnly = False
            cbxHQSQLBaza.Enabled = True
        End If

        cbxHQSQLServer.Text = ParL_HQSerwer
        For i = 0 To cbxHQSQLServer.Items.Count() - 1
            If ParL_HQSerwer = cbxHQSQLServer.Items.Item(i) Then
                'cbxHQSQLServer.Select(0, 1)
                cbxHQSQLServer.Text = ParL_HQSerwer
                Jest = True
                Exit For
            End If
        Next
        If Not Jest Then
            'nie znalaz³ takiej bazy - dopisz na koniec
            cbxHQSQLServer.Items.Add(ParL_HQSerwer)
            cbxHQSQLServer.Text = ParL_HQSerwer
        End If

        txtHQSQLLogin.Text = ParL_HQLogin
        txtHQSQLPassword.Text = ParL_HQPassword
        cbxHQSQLBaza.Text = ParL_HQDataBase
        '----------------------
        oWersjaBazyDanych = HQWersjaBazyDanych()
        If oWersjaBazyDanych = 0 Then
            txtHQWersja.Text = "????"
        Else
            txtHQWersja.Text = (oWersjaBazyDanych / 100).ToString("N2")
        End If
        txtHQWersjaWymagana.Text = (oAktualnaWersjaBazyDanych / 100).ToString("N2")









        'cbxeMailService.Text = Par_eMailService
        For i = 0 To cbxeMailService.Items.Count() - 1
            If ParL_eMailService = cbxeMailService.Items.Item(i) Then
                cbxeMailService.Select(i, 10)
                cbxeMailService.Text = ParL_eMailService
                Exit For
            End If
        Next

        txtAdres.Text = ParL_eMailAdres
        txtSMTP.Text = ParL_SMTP
        txtPOP3.Text = ParL_POP3
        txtSMTPPort.Text = ParL_SMTPPort
        txtPOP3Port.Text = ParL_POP3Port
        txtLogin.Text = ParL_eMailUser
        txtPass.Text = ParL_eMailPass
        If ParL_SSLProtocol = "TAK" Then
            chkSSL.Checked = True
        Else
            chkSSL.Checked = False
        End If
    End Sub

    Private Sub ZapiszParametry()
        ParL_OpisInstalacji = txtOpisInstalacji.Text
        ParL_NazwaArchiwum = txtNazwaArchiwum.Text
        ParL_KatalogArchiwum = txtKatalogArchiwum.Text
        ParL_KatalogArchiwumZIP = txtKatalogArchiwumZIP.Text
        ParL_KatalogRaporty = txtKatalogRaporty.Text

        ParL_DataBaseType = cbxBazaDanych.Text
        ParL_KatZbiorow = lblKatalog.Text

        If rbtWindows.Checked Then
            ParL_Autentykacja = "Windows"
        Else
            ParL_Autentykacja = "SQL Server"
        End If
        ParL_Serwer = cbxSQLServer.Text
        ParL_Login = txtSQLLogin.Text
        ParL_Password = txtSQLPassword.Text
        ParL_DataBase = cbxSQLBaza.Text

        ParL_eMailService = cbxeMailService.Text
        ParL_eMailAdres = txtAdres.Text
        ParL_SMTP = txtSMTP.Text
        ParL_POP3 = txtPOP3.Text
        ParL_SMTPPort = txtSMTPPort.Text
        ParL_POP3Port = txtPOP3Port.Text
        ParL_eMailUser = txtLogin.Text
        ParL_eMailPass = txtPass.Text
        If chkSSL.Checked Then
            ParL_SSLProtocol = "TAK"
        Else
            ParL_SSLProtocol = "NIE"
        End If


        ParL_HQ = "TAK"

        'ParL_HQDataBaseType = cbxHQBazaDanych.Text
        'ParL_HQPlikMDB = lblHQPlikMDB.Text

        If rbtHQWindows.Checked Then
            ParL_HQAutentykacja = "Windows"
        Else
            ParL_HQAutentykacja = "SQL Server"
        End If
        ParL_HQSerwer = cbxHQSQLServer.Text
        ParL_HQLogin = txtHQSQLLogin.Text
        ParL_HQPassword = txtHQSQLPassword.Text
        ParL_HQDataBase = cbxHQSQLBaza.Text




        ZapiszParametryLokalne()
    End Sub

    Private Sub tbcParametry_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbcParametry.SelectedIndexChanged
        Select Case tbcParametry.SelectedIndex
            Case 0
            Case 1
                lblJakaBaza.Text = "Program obs³uguje Baze Danych typu " & cbxBazaDanych.Text
                'If cbxBazaDanych.Text = "MS ACCESS" Then
                '    Label10.Visible = True
                '    lblKatalog.Visible = True
                '    cmdPobierzKatalog.Visible = True
                '    '--------------------------------------
                '    Label9.Visible = False
                '    Label11.Visible = False
                '    Label12.Visible = False
                '    Label13.Visible = False
                '    Label14.Visible = False
                '    Label15.Visible = False
                '    Label16.Visible = False

                '    rbtWindows.Visible = False
                '    rbtSQLServer.Visible = False
                '    cbxSQLServer.Visible = False
                '    txtSQLLogin.Visible = False
                '    txtSQLPassword.Visible = False
                '    cmdPokazSQLPass.Visible = False
                '    cbxSQLBaza.Visible = False
                '    lblNazwa.Visible = False
                '    lblIP.Visible = False

                '    btnTestuj.Visible = False
                '    btnNowaBaza.Visible = False
                'Else
                Label10.Visible = False
                lblKatalog.Visible = False
                cmdPobierzKatalog.Visible = False
                '--------------------------------------
                Label9.Visible = True
                Label11.Visible = True
                Label12.Visible = True
                Label13.Visible = True
                Label14.Visible = True
                Label15.Visible = True
                Label16.Visible = True

                rbtWindows.Visible = True
                rbtSQLServer.Visible = True
                cbxSQLServer.Visible = True
                txtSQLLogin.Visible = True
                txtSQLPassword.Visible = True
                cmdPokazSQLPass.Visible = True
                cbxSQLBaza.Visible = True
                lblNazwa.Visible = True
                lblIP.Visible = True

                btnTestuj.Visible = True
                btnNowaBaza.Visible = True
                'End If
        End Select
    End Sub


    Private Sub PobierzListeBazDanych()
        Dim Cur As Cursor
        Cur = Windows.Forms.Cursor.Current
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '------------------------------------------------
        'pobierz dostêpne bazy danych
        cbxSQLBaza.Items.Clear()

        Dim sConnectionStr As String = ""
        Dim dbSerwer As String = ""
        Dim dbLogin As String = ""
        Dim dbPassword As String = ""
        Dim dbBaza As String = ""

        dbSerwer = Trim(cbxSQLServer.Text)
        dbLogin = Trim(txtSQLLogin.Text)
        dbPassword = Trim(txtSQLPassword.Text)
        dbBaza = "master"

        If rbtWindows.Checked Then
            sConnectionStr = "Data Source=" & dbSerwer & ";" &
                   "Integrated Security=SSPI;" &
                   "Initial Catalog=" & dbBaza & ""
        Else
            sConnectionStr = "Persist Security Info=False;" &
                   "User ID=" & dbLogin & ";" &
                   "Password=" & dbPassword & ";" &
                   "Initial Catalog=" & dbBaza & ";" &
                   "Data Source=" & dbSerwer & ";" &
                   "Packet Size=4096"
        End If

        Dim sSelectSQL As String = "exec sp_catalogs_rowset;2 NULL"
        Dim dbConnection As New SqlClient.SqlConnection(sConnectionStr)
        Dim dbCommandSelect As New SqlClient.SqlCommand(sSelectSQL, dbConnection)
        Dim dbDataAdapter As New SqlClient.SqlDataAdapter(dbCommandSelect)
        Dim dbDataSet As New DataSet
        Dim dbDataView As New DataView
        dbDataSet.Locale = New System.Globalization.CultureInfo("pl-PL")
        Dim MapowanieTabeli As System.Data.Common.DataTableMapping
        MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Database")
        'wypelnij DATASET

        Try
            dbConnection.Open()
        Catch Ex As System.Exception
        End Try

        If dbConnection.State = ConnectionState.Open Then
            Try
                dbDataAdapter.FillSchema(dbDataSet, SchemaType.Mapped, "TABELA_Database")
                dbDataAdapter.Fill(dbDataSet)
                dbConnection.Close()

                'definiuj dbDataView
                dbDataView = New DataView(dbDataSet.Tables(0))
                If dbDataView.Count > 0 Then
                    Dim dr As DataRow
                    Dim db As String
                    For Each dr In dbDataView.Table.Rows
                        db = dr.Item("CATALOG_NAME")
                        cbxSQLBaza.Items.Add(db)
                    Next
                End If
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End Try
        End If
        '------------------------------------------------
        Windows.Forms.Cursor.Current = Cur
    End Sub

    Private Sub HQPobierzListeBazDanych()
        Dim Cur As Cursor
        Cur = Cursor.Current
        Cursor.Current = Cursors.WaitCursor
        '------------------------------------------------
        'pobierz dostêpne bazy danych
        cbxHQSQLBaza.Items.Clear()

        Dim sConnectionStr As String = ""
        Dim dbSerwer As String = ""
        Dim dbLogin As String = ""
        Dim dbPassword As String = ""
        Dim dbBaza As String = ""

        dbSerwer = Trim(cbxHQSQLServer.Text)
        dbLogin = Trim(txtHQSQLLogin.Text)
        dbPassword = Trim(txtHQSQLPassword.Text)
        dbBaza = "master"

        If rbtHQWindows.Checked Then
            sConnectionStr = "Data Source=" & dbSerwer & ";" &
                   "Integrated Security=SSPI;" &
                   "Initial Catalog=" & dbBaza & ""
        Else
            sConnectionStr = "Persist Security Info=False;" &
                   "User ID=" & dbLogin & ";" &
                   "Password=" & dbPassword & ";" &
                   "Initial Catalog=" & dbBaza & ";" &
                   "Data Source=" & dbSerwer & ";" &
                   "Packet Size=4096"
        End If

        Dim sSelectSQL As String = "exec sp_catalogs_rowset;2 NULL"
        Dim sqlConnection As New SqlClient.SqlConnection(sConnectionStr)
        Dim sqlCommandSelect As New SqlClient.SqlCommand(sSelectSQL, sqlConnection)
        Dim sqlDataAdapter As New SqlClient.SqlDataAdapter(sqlCommandSelect)
        Dim sqlDataSet As New DataSet
        Dim sqlDataView As New DataView
        sqlDataSet.Locale = New System.Globalization.CultureInfo("pl-PL")

        Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
        DBMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Database")

        'wypelnij DATASET
        Try
            sqlConnection.Open()
        Catch Ex As System.Exception
        End Try

        If sqlConnection.State = ConnectionState.Open Then
            Try
                sqlDataAdapter.FillSchema(sqlDataSet, SchemaType.Mapped, "TABELA_Database")
                sqlDataAdapter.Fill(sqlDataSet)
                sqlConnection.Close()

                'definiuj sqlDataView
                sqlDataView = New DataView(sqlDataSet.Tables(0))
                If sqlDataView.Count > 0 Then
                    Dim dr As DataRow
                    Dim db As String
                    For Each dr In sqlDataView.Table.Rows
                        db = dr.Item("CATALOG_NAME")
                        cbxHQSQLBaza.Items.Add(db)
                    Next
                End If
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End Try
        End If
        '------------------------------------------------
        Cursor.Current = Cur
    End Sub













    Private Sub rbtSQLServer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtSQLServer.CheckedChanged
        If rbtSQLServer.Checked Then
            'rbtSQLServer.Checked = True
            cbxSQLServer.Enabled = True
            txtSQLLogin.ReadOnly = False
            txtSQLPassword.ReadOnly = False
            cbxSQLBaza.Enabled = True
        Else
            'rbtWindows.Checked = True
            cbxSQLServer.Enabled = True
            txtSQLLogin.ReadOnly = True
            txtSQLPassword.ReadOnly = True
            cbxSQLBaza.Enabled = True
        End If
    End Sub

    Private Sub rbtWindows_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtWindows.CheckedChanged
        If rbtWindows.Checked Then
            'rbtWindows.Checked = True
            cbxSQLServer.Enabled = True
            txtSQLLogin.ReadOnly = True
            txtSQLPassword.ReadOnly = True
            cbxSQLBaza.Enabled = True
        Else
            'rbtSQLServer.Checked = True
            cbxSQLServer.Enabled = True
            txtSQLLogin.ReadOnly = False
            txtSQLPassword.ReadOnly = False
            cbxSQLBaza.Enabled = True
        End If
    End Sub






    Private Sub rbtHQWindows_CheckedChanged(sender As Object, e As EventArgs) Handles rbtHQWindows.CheckedChanged
        If rbtHQWindows.Checked Then
            'rbthqWindows.Checked = True
            cbxHQSQLServer.Enabled = True
            txtHQSQLLogin.ReadOnly = True
            txtHQSQLPassword.ReadOnly = True
            cbxHQSQLBaza.Enabled = True
        Else
            'rbthqSQLServer.Checked = True
            cbxHQSQLServer.Enabled = True
            txtHQSQLLogin.ReadOnly = False
            txtHQSQLPassword.ReadOnly = False
            cbxHQSQLBaza.Enabled = True
        End If
    End Sub

    Private Sub rbtHQSQLServer_CheckedChanged(sender As Object, e As EventArgs) Handles rbtHQSQLServer.CheckedChanged
        If rbtHQSQLServer.Checked Then
            'rbthqSQLServer.Checked = True
            cbxHQSQLServer.Enabled = True
            txtHQSQLLogin.ReadOnly = False
            txtHQSQLPassword.ReadOnly = False
            cbxHQSQLBaza.Enabled = True
        Else
            'rbthqWindows.Checked = True
            cbxHQSQLServer.Enabled = True
            txtHQSQLLogin.ReadOnly = True
            txtHQSQLPassword.ReadOnly = True
            cbxHQSQLBaza.Enabled = True
        End If
    End Sub



    Private Sub btnNowaBaza_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNowaBaza.Click
        ZapiszParametry()
        '----------------------------
        Dim forma As New ZalozBazeSQL
        forma.Show()
    End Sub

    Private Sub cmdPokazHasloeMail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokazHasloeMail.Click
        If txtPass.PasswordChar = "*" Then
            txtPass.PasswordChar = ""
            cmdPokazHasloeMail.Text = "Ukryj has³o"
        Else
            txtPass.PasswordChar = "*"
            cmdPokazHasloeMail.Text = "Poka¿ has³o"
        End If
    End Sub

    Private Sub cmdPokazSQLPass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokazSQLPass.Click
        If txtSQLPassword.PasswordChar = "*" Then
            txtSQLPassword.PasswordChar = ""
            cmdPokazSQLPass.Text = "Ukryj has³o"
        Else
            txtSQLPassword.PasswordChar = "*"
            cmdPokazSQLPass.Text = "Poka¿ has³o"
        End If
    End Sub

    Private Sub cmdHQSQLPokazHaslo_Click(sender As Object, e As EventArgs) Handles cmdHQSQLPokazHaslo.Click
        If txtHQSQLPassword.PasswordChar = "*" Then
            txtHQSQLPassword.PasswordChar = ""
            cmdHQSQLPokazHaslo.Text = "Ukryj has³o"
        Else
            txtHQSQLPassword.PasswordChar = "*"
            cmdHQSQLPokazHaslo.Text = "Poka¿ has³o"
        End If
    End Sub






















    '====================
    ' pobierz dane o klientach z Bazy Danych do Importu
    '=======================

    Private Sub cmdHQPobierz_Click(sender As Object, e As EventArgs) Handles cmdHQPobierz.Click
        ZapiszParametry()
        '-----------------------------
        Dim Form As New frmImportujKlientow(txtPZnakIdKlienta.Text, txtPZnakIdPracownika.Text)
        System.Windows.Forms.Application.DoEvents()
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing

    End Sub

    Private Sub cmdHQWersjaBazyDanych_Click(sender As Object, e As EventArgs) Handles cmdHQWersjaBazyDanych.Click
        ZapiszParametry()
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New AktualizacjaStrukturBazyDanych("HQ")
        Form.ShowDialog()
        '----------------------
        oWersjaBazyDanych = HQWersjaBazyDanych()
        If oWersjaBazyDanych = 0 Then
            txtHQWersja.Text = "????"
        Else
            txtHQWersja.Text = (oWersjaBazyDanych / 100).ToString("N2")
        End If
    End Sub

    Private Sub txtPZnakIdKlienta_TextChanged(sender As Object, e As EventArgs) Handles txtPZnakIdKlienta.TextChanged
        TestLen(txtPZnakIdKlienta, "PIERWSZY ZNAK ID KLIENTA", 1)
    End Sub
    Private Sub txtPZnakIdKlienta_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPZnakIdKlienta.GotFocus
        StartEdycjiTxt(txtPZnakIdKlienta)
    End Sub
    Private Sub txtPZnakIdKlienta_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPZnakIdKlienta.LostFocus
        KoniecEdycjiTxt(txtPZnakIdKlienta)
    End Sub




    Private Sub txtPZnakIdPracownika_TextChanged(sender As Object, e As EventArgs) Handles txtPZnakIdPracownika.TextChanged
        TestLen(txtPZnakIdPracownika, "PIERWSZY ZNAK ID PRACOWNIKA", 1)
    End Sub
    Private Sub txtPZnakIdPracownika_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPZnakIdPracownika.GotFocus
        StartEdycjiTxt(txtPZnakIdPracownika)
    End Sub
    Private Sub txtPZnakIdPracownika_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPZnakIdPracownika.LostFocus
        KoniecEdycjiTxt(txtPZnakIdPracownika)
    End Sub


End Class
