Public Class Parametry
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents tbcParametry As System.Windows.Forms.TabControl
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Uzytkownik As System.Windows.Forms.TabPage
    Friend WithEvents Techniczne As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents dlgOpenFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents txtBankUzytkownika As System.Windows.Forms.TextBox
    Friend WithEvents txtKontoUzytkownika As System.Windows.Forms.TextBox
    Friend WithEvents txtNIPUzytkownika As System.Windows.Forms.TextBox
    Friend WithEvents txtMiejscowoscUzytkownika As System.Windows.Forms.TextBox
    Friend WithEvents txtAdresUzytkownika As System.Windows.Forms.TextBox
    Friend WithEvents txtNazwaUzytkownika As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblKatalogOferty As System.Windows.Forms.Label
    Friend WithEvents cmdWybierzKatalog As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdDelKatalog As System.Windows.Forms.Button
    Friend WithEvents txtIdentUzytkownika As System.Windows.Forms.TextBox
    Friend WithEvents TabPAYBACK As System.Windows.Forms.TabPage
    Friend WithEvents txtWWWPayBack As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cmdPokazWWW As System.Windows.Forms.Button
    Friend WithEvents TabRaportAkt As System.Windows.Forms.TabPage
    Friend WithEvents txtEXCEL10 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL09 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL08 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL07 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL06 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL05 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL04 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL03 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL02 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL01 As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtIdentOddzialu As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TabWartGran As TabPage
    Friend WithEvents txtWartGranP As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents txtWartGranW As TextBox
    Friend WithEvents Label23 As Label
    Friend WithEvents Label26 As Label
    Friend WithEvents txtWartGranMatP As TextBox
    Friend WithEvents Label27 As Label
    Friend WithEvents txtWartGranMatW As TextBox
    Friend WithEvents Label28 As Label
    Friend WithEvents Label25 As Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Parametry))
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.tbcParametry = New System.Windows.Forms.TabControl()
        Me.Uzytkownik = New System.Windows.Forms.TabPage()
        Me.txtIdentOddzialu = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.txtIdentUzytkownika = New System.Windows.Forms.TextBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.txtBankUzytkownika = New System.Windows.Forms.TextBox()
        Me.txtKontoUzytkownika = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtNIPUzytkownika = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtMiejscowoscUzytkownika = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtAdresUzytkownika = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtNazwaUzytkownika = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Techniczne = New System.Windows.Forms.TabPage()
        Me.cmdDelKatalog = New System.Windows.Forms.Button()
        Me.cmdWybierzKatalog = New System.Windows.Forms.Button()
        Me.lblKatalogOferty = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TabPAYBACK = New System.Windows.Forms.TabPage()
        Me.cmdPokazWWW = New System.Windows.Forms.Button()
        Me.txtWWWPayBack = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TabRaportAkt = New System.Windows.Forms.TabPage()
        Me.txtEXCEL10 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL09 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL08 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL07 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL06 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL05 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL04 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL03 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL02 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL01 = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.TabWartGran = New System.Windows.Forms.TabPage()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.txtWartGranMatP = New System.Windows.Forms.TextBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.txtWartGranMatW = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.txtWartGranP = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtWartGranW = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.dlgOpenFolder = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbcParametry.SuspendLayout()
        Me.Uzytkownik.SuspendLayout()
        Me.Techniczne.SuspendLayout()
        Me.TabPAYBACK.SuspendLayout()
        Me.TabRaportAkt.SuspendLayout()
        Me.TabWartGran.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 299)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 17
        Me.PictureBox2.TabStop = False
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(551, 297)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(84, 23)
        Me.cmdPowrot.TabIndex = 40
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tbcParametry
        '
        Me.tbcParametry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbcParametry.Controls.Add(Me.Uzytkownik)
        Me.tbcParametry.Controls.Add(Me.Techniczne)
        Me.tbcParametry.Controls.Add(Me.TabPAYBACK)
        Me.tbcParametry.Controls.Add(Me.TabRaportAkt)
        Me.tbcParametry.Controls.Add(Me.TabWartGran)
        Me.tbcParametry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.tbcParametry.Location = New System.Drawing.Point(8, 8)
        Me.tbcParametry.Name = "tbcParametry"
        Me.tbcParametry.SelectedIndex = 0
        Me.tbcParametry.Size = New System.Drawing.Size(632, 283)
        Me.tbcParametry.TabIndex = 42
        '
        'Uzytkownik
        '
        Me.Uzytkownik.BackColor = System.Drawing.SystemColors.Control
        Me.Uzytkownik.Controls.Add(Me.txtIdentOddzialu)
        Me.Uzytkownik.Controls.Add(Me.Label22)
        Me.Uzytkownik.Controls.Add(Me.txtIdentUzytkownika)
        Me.Uzytkownik.Controls.Add(Me.Label46)
        Me.Uzytkownik.Controls.Add(Me.txtBankUzytkownika)
        Me.Uzytkownik.Controls.Add(Me.txtKontoUzytkownika)
        Me.Uzytkownik.Controls.Add(Me.Label7)
        Me.Uzytkownik.Controls.Add(Me.Label6)
        Me.Uzytkownik.Controls.Add(Me.txtNIPUzytkownika)
        Me.Uzytkownik.Controls.Add(Me.Label5)
        Me.Uzytkownik.Controls.Add(Me.txtMiejscowoscUzytkownika)
        Me.Uzytkownik.Controls.Add(Me.Label4)
        Me.Uzytkownik.Controls.Add(Me.txtAdresUzytkownika)
        Me.Uzytkownik.Controls.Add(Me.Label3)
        Me.Uzytkownik.Controls.Add(Me.txtNazwaUzytkownika)
        Me.Uzytkownik.Controls.Add(Me.Label1)
        Me.Uzytkownik.ForeColor = System.Drawing.Color.Navy
        Me.Uzytkownik.Location = New System.Drawing.Point(4, 23)
        Me.Uzytkownik.Name = "Uzytkownik"
        Me.Uzytkownik.Size = New System.Drawing.Size(624, 256)
        Me.Uzytkownik.TabIndex = 0
        Me.Uzytkownik.Text = "U¿ytkownik"
        '
        'txtIdentOddzialu
        '
        Me.txtIdentOddzialu.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtIdentOddzialu.ForeColor = System.Drawing.Color.Purple
        Me.txtIdentOddzialu.Location = New System.Drawing.Point(128, 172)
        Me.txtIdentOddzialu.Name = "txtIdentOddzialu"
        Me.txtIdentOddzialu.Size = New System.Drawing.Size(487, 20)
        Me.txtIdentOddzialu.TabIndex = 44
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(8, 172)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(136, 16)
        Me.Label22.TabIndex = 43
        Me.Label22.Text = "Ident. Oddzia³u . . . . . . . . . . . "
        '
        'txtIdentUzytkownika
        '
        Me.txtIdentUzytkownika.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtIdentUzytkownika.ForeColor = System.Drawing.Color.Purple
        Me.txtIdentUzytkownika.Location = New System.Drawing.Point(128, 9)
        Me.txtIdentUzytkownika.Name = "txtIdentUzytkownika"
        Me.txtIdentUzytkownika.Size = New System.Drawing.Size(172, 20)
        Me.txtIdentUzytkownika.TabIndex = 42
        '
        'Label46
        '
        Me.Label46.Location = New System.Drawing.Point(8, 9)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(136, 16)
        Me.Label46.TabIndex = 41
        Me.Label46.Text = "Ident. u¿ytkownika . . . . . . . . . . . "
        '
        'txtBankUzytkownika
        '
        Me.txtBankUzytkownika.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBankUzytkownika.ForeColor = System.Drawing.Color.Purple
        Me.txtBankUzytkownika.Location = New System.Drawing.Point(128, 105)
        Me.txtBankUzytkownika.Name = "txtBankUzytkownika"
        Me.txtBankUzytkownika.Size = New System.Drawing.Size(487, 20)
        Me.txtBankUzytkownika.TabIndex = 5
        '
        'txtKontoUzytkownika
        '
        Me.txtKontoUzytkownika.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtKontoUzytkownika.ForeColor = System.Drawing.Color.Purple
        Me.txtKontoUzytkownika.Location = New System.Drawing.Point(128, 129)
        Me.txtKontoUzytkownika.Name = "txtKontoUzytkownika"
        Me.txtKontoUzytkownika.Size = New System.Drawing.Size(487, 20)
        Me.txtKontoUzytkownika.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 129)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(136, 16)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Konto bankowe . . . . . . . . . . . "
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 105)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 16)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Bank u¿ytkownika . . . . . . . . . . . "
        '
        'txtNIPUzytkownika
        '
        Me.txtNIPUzytkownika.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNIPUzytkownika.ForeColor = System.Drawing.Color.Purple
        Me.txtNIPUzytkownika.Location = New System.Drawing.Point(480, 81)
        Me.txtNIPUzytkownika.Name = "txtNIPUzytkownika"
        Me.txtNIPUzytkownika.Size = New System.Drawing.Size(135, 20)
        Me.txtNIPUzytkownika.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.Location = New System.Drawing.Point(444, 81)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 16)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "NIP . . . . . . . . . . . "
        '
        'txtMiejscowoscUzytkownika
        '
        Me.txtMiejscowoscUzytkownika.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMiejscowoscUzytkownika.ForeColor = System.Drawing.Color.Purple
        Me.txtMiejscowoscUzytkownika.Location = New System.Drawing.Point(128, 81)
        Me.txtMiejscowoscUzytkownika.Name = "txtMiejscowoscUzytkownika"
        Me.txtMiejscowoscUzytkownika.Size = New System.Drawing.Size(288, 20)
        Me.txtMiejscowoscUzytkownika.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 81)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 16)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Miejscowoœæ/siedziba . . . . . . . . . . . "
        '
        'txtAdresUzytkownika
        '
        Me.txtAdresUzytkownika.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAdresUzytkownika.ForeColor = System.Drawing.Color.Purple
        Me.txtAdresUzytkownika.Location = New System.Drawing.Point(128, 57)
        Me.txtAdresUzytkownika.Name = "txtAdresUzytkownika"
        Me.txtAdresUzytkownika.Size = New System.Drawing.Size(487, 20)
        Me.txtAdresUzytkownika.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Adres u¿ytkownika . . . . . . . . . . . "
        '
        'txtNazwaUzytkownika
        '
        Me.txtNazwaUzytkownika.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNazwaUzytkownika.ForeColor = System.Drawing.Color.Purple
        Me.txtNazwaUzytkownika.Location = New System.Drawing.Point(128, 33)
        Me.txtNazwaUzytkownika.Name = "txtNazwaUzytkownika"
        Me.txtNazwaUzytkownika.Size = New System.Drawing.Size(487, 20)
        Me.txtNazwaUzytkownika.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Nazwa u¿ytkownika . . . . . . . . . . . "
        '
        'Techniczne
        '
        Me.Techniczne.BackColor = System.Drawing.SystemColors.Control
        Me.Techniczne.Controls.Add(Me.cmdDelKatalog)
        Me.Techniczne.Controls.Add(Me.cmdWybierzKatalog)
        Me.Techniczne.Controls.Add(Me.lblKatalogOferty)
        Me.Techniczne.Controls.Add(Me.Label14)
        Me.Techniczne.Controls.Add(Me.Label2)
        Me.Techniczne.Location = New System.Drawing.Point(4, 23)
        Me.Techniczne.Name = "Techniczne"
        Me.Techniczne.Size = New System.Drawing.Size(624, 256)
        Me.Techniczne.TabIndex = 1
        Me.Techniczne.Text = "Oferty"
        '
        'cmdDelKatalog
        '
        Me.cmdDelKatalog.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDelKatalog.Image = CType(resources.GetObject("cmdDelKatalog.Image"), System.Drawing.Image)
        Me.cmdDelKatalog.Location = New System.Drawing.Point(559, 26)
        Me.cmdDelKatalog.Name = "cmdDelKatalog"
        Me.cmdDelKatalog.Size = New System.Drawing.Size(49, 22)
        Me.cmdDelKatalog.TabIndex = 106
        '
        'cmdWybierzKatalog
        '
        Me.cmdWybierzKatalog.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzKatalog.Image = CType(resources.GetObject("cmdWybierzKatalog.Image"), System.Drawing.Image)
        Me.cmdWybierzKatalog.Location = New System.Drawing.Point(520, 26)
        Me.cmdWybierzKatalog.Name = "cmdWybierzKatalog"
        Me.cmdWybierzKatalog.Size = New System.Drawing.Size(33, 22)
        Me.cmdWybierzKatalog.TabIndex = 89
        '
        'lblKatalogOferty
        '
        Me.lblKatalogOferty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblKatalogOferty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalogOferty.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalogOferty.Location = New System.Drawing.Point(56, 28)
        Me.lblKatalogOferty.Name = "lblKatalogOferty"
        Me.lblKatalogOferty.Size = New System.Drawing.Size(459, 16)
        Me.lblKatalogOferty.TabIndex = 15
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Navy
        Me.Label14.Location = New System.Drawing.Point(8, 8)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(600, 16)
        Me.Label14.TabIndex = 13
        Me.Label14.Text = "Katalog dyskowy do którego wpisywane bêd¹ oferty sk³adane klientom :"
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 27)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Katalog . . . . . . . . . . . "
        '
        'TabPAYBACK
        '
        Me.TabPAYBACK.BackColor = System.Drawing.SystemColors.Control
        Me.TabPAYBACK.Controls.Add(Me.cmdPokazWWW)
        Me.TabPAYBACK.Controls.Add(Me.txtWWWPayBack)
        Me.TabPAYBACK.Controls.Add(Me.Label9)
        Me.TabPAYBACK.Controls.Add(Me.Label8)
        Me.TabPAYBACK.Location = New System.Drawing.Point(4, 23)
        Me.TabPAYBACK.Name = "TabPAYBACK"
        Me.TabPAYBACK.Size = New System.Drawing.Size(624, 256)
        Me.TabPAYBACK.TabIndex = 2
        Me.TabPAYBACK.Text = "PAYBACK"
        '
        'cmdPokazWWW
        '
        Me.cmdPokazWWW.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazWWW.ForeColor = System.Drawing.Color.Black
        Me.cmdPokazWWW.Image = CType(resources.GetObject("cmdPokazWWW.Image"), System.Drawing.Image)
        Me.cmdPokazWWW.Location = New System.Drawing.Point(569, 23)
        Me.cmdPokazWWW.Name = "cmdPokazWWW"
        Me.cmdPokazWWW.Size = New System.Drawing.Size(39, 23)
        Me.cmdPokazWWW.TabIndex = 21
        Me.cmdPokazWWW.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtWWWPayBack
        '
        Me.txtWWWPayBack.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtWWWPayBack.ForeColor = System.Drawing.Color.Purple
        Me.txtWWWPayBack.Location = New System.Drawing.Point(102, 24)
        Me.txtWWWPayBack.Name = "txtWWWPayBack"
        Me.txtWWWPayBack.Size = New System.Drawing.Size(461, 20)
        Me.txtWWWPayBack.TabIndex = 16
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(8, 24)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(136, 16)
        Me.Label9.TabIndex = 15
        Me.Label9.Text = "Stroma www. . . . . . . . . . . "
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(8, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(600, 16)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Strona Internetowa programu PAYBACK :"
        '
        'TabRaportAkt
        '
        Me.TabRaportAkt.BackColor = System.Drawing.SystemColors.Control
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL10)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL09)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL08)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL07)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL06)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL05)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL04)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL03)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL02)
        Me.TabRaportAkt.Controls.Add(Me.txtEXCEL01)
        Me.TabRaportAkt.Controls.Add(Me.Label15)
        Me.TabRaportAkt.Controls.Add(Me.Label10)
        Me.TabRaportAkt.Controls.Add(Me.Label13)
        Me.TabRaportAkt.Controls.Add(Me.Label12)
        Me.TabRaportAkt.Controls.Add(Me.Label11)
        Me.TabRaportAkt.Controls.Add(Me.Label16)
        Me.TabRaportAkt.Controls.Add(Me.Label17)
        Me.TabRaportAkt.Controls.Add(Me.Label18)
        Me.TabRaportAkt.Controls.Add(Me.Label19)
        Me.TabRaportAkt.Controls.Add(Me.Label20)
        Me.TabRaportAkt.Controls.Add(Me.Label21)
        Me.TabRaportAkt.Location = New System.Drawing.Point(4, 23)
        Me.TabRaportAkt.Name = "TabRaportAkt"
        Me.TabRaportAkt.Size = New System.Drawing.Size(624, 256)
        Me.TabRaportAkt.TabIndex = 3
        Me.TabRaportAkt.Text = "Raport Aktywnoœci"
        '
        'txtEXCEL10
        '
        Me.txtEXCEL10.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL10.Location = New System.Drawing.Point(167, 223)
        Me.txtEXCEL10.Name = "txtEXCEL10"
        Me.txtEXCEL10.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL10.TabIndex = 467
        Me.txtEXCEL10.Text = "0"
        '
        'txtEXCEL09
        '
        Me.txtEXCEL09.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL09.Location = New System.Drawing.Point(167, 201)
        Me.txtEXCEL09.Name = "txtEXCEL09"
        Me.txtEXCEL09.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL09.TabIndex = 466
        Me.txtEXCEL09.Text = "0"
        '
        'txtEXCEL08
        '
        Me.txtEXCEL08.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL08.Location = New System.Drawing.Point(167, 179)
        Me.txtEXCEL08.Name = "txtEXCEL08"
        Me.txtEXCEL08.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL08.TabIndex = 465
        Me.txtEXCEL08.Text = "0"
        '
        'txtEXCEL07
        '
        Me.txtEXCEL07.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL07.Location = New System.Drawing.Point(167, 157)
        Me.txtEXCEL07.Name = "txtEXCEL07"
        Me.txtEXCEL07.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL07.TabIndex = 464
        Me.txtEXCEL07.Text = "0"
        '
        'txtEXCEL06
        '
        Me.txtEXCEL06.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL06.Location = New System.Drawing.Point(167, 135)
        Me.txtEXCEL06.Name = "txtEXCEL06"
        Me.txtEXCEL06.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL06.TabIndex = 463
        Me.txtEXCEL06.Text = "0"
        '
        'txtEXCEL05
        '
        Me.txtEXCEL05.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL05.Location = New System.Drawing.Point(167, 113)
        Me.txtEXCEL05.Name = "txtEXCEL05"
        Me.txtEXCEL05.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL05.TabIndex = 462
        Me.txtEXCEL05.Text = "0"
        '
        'txtEXCEL04
        '
        Me.txtEXCEL04.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL04.Location = New System.Drawing.Point(167, 91)
        Me.txtEXCEL04.Name = "txtEXCEL04"
        Me.txtEXCEL04.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL04.TabIndex = 461
        Me.txtEXCEL04.Text = "0"
        '
        'txtEXCEL03
        '
        Me.txtEXCEL03.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL03.Location = New System.Drawing.Point(167, 69)
        Me.txtEXCEL03.Name = "txtEXCEL03"
        Me.txtEXCEL03.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL03.TabIndex = 460
        Me.txtEXCEL03.Text = "0"
        '
        'txtEXCEL02
        '
        Me.txtEXCEL02.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL02.Location = New System.Drawing.Point(167, 47)
        Me.txtEXCEL02.Name = "txtEXCEL02"
        Me.txtEXCEL02.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL02.TabIndex = 459
        Me.txtEXCEL02.Text = "0"
        '
        'txtEXCEL01
        '
        Me.txtEXCEL01.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL01.Location = New System.Drawing.Point(167, 25)
        Me.txtEXCEL01.Name = "txtEXCEL01"
        Me.txtEXCEL01.Size = New System.Drawing.Size(345, 20)
        Me.txtEXCEL01.TabIndex = 458
        Me.txtEXCEL01.Text = "0"
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Navy
        Me.Label15.Location = New System.Drawing.Point(14, 226)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(199, 16)
        Me.Label15.TabIndex = 457
        Me.Label15.Text = "Dodatkowa kolumna Nr 10. . . . . . . "
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(14, 204)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(199, 16)
        Me.Label10.TabIndex = 456
        Me.Label10.Text = "Dodatkowa kolumna Nr 09. . . . . . . "
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Navy
        Me.Label13.Location = New System.Drawing.Point(14, 182)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(199, 16)
        Me.Label13.TabIndex = 455
        Me.Label13.Text = "Dodatkowa kolumna Nr 08. . . . . . . "
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(14, 160)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(199, 16)
        Me.Label12.TabIndex = 454
        Me.Label12.Text = "Dodatkowa kolumna Nr 07. . . . . . . "
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(14, 138)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(199, 16)
        Me.Label11.TabIndex = 453
        Me.Label11.Text = "Dodatkowa kolumna Nr 06. . . . . . . "
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Navy
        Me.Label16.Location = New System.Drawing.Point(14, 116)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(199, 16)
        Me.Label16.TabIndex = 452
        Me.Label16.Text = "Dodatkowa kolumna Nr 05. . . . . . . "
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Navy
        Me.Label17.Location = New System.Drawing.Point(14, 94)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(199, 16)
        Me.Label17.TabIndex = 451
        Me.Label17.Text = "Dodatkowa kolumna Nr 04. . . . . . . "
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(14, 72)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(199, 16)
        Me.Label18.TabIndex = 450
        Me.Label18.Text = "Dodatkowa kolumna Nr 03. . . . . . . "
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Navy
        Me.Label19.Location = New System.Drawing.Point(14, 50)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(199, 16)
        Me.Label19.TabIndex = 449
        Me.Label19.Text = "Dodatkowa kolumna Nr 02. . . . . . . "
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Navy
        Me.Label20.Location = New System.Drawing.Point(14, 28)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(196, 16)
        Me.Label20.TabIndex = 448
        Me.Label20.Text = "Dodatkowa kolumna Nr 01. . . . . . . "
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Navy
        Me.Label21.Location = New System.Drawing.Point(14, 9)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(596, 19)
        Me.Label21.TabIndex = 447
        Me.Label21.Text = "Nag³ówki dodatkowych kolumn dodane do Raportu Aktywnoœci  emitowanego w EXCEL'u :" &
    ""
        '
        'TabWartGran
        '
        Me.TabWartGran.BackColor = System.Drawing.SystemColors.Control
        Me.TabWartGran.Controls.Add(Me.Label26)
        Me.TabWartGran.Controls.Add(Me.txtWartGranMatP)
        Me.TabWartGran.Controls.Add(Me.Label27)
        Me.TabWartGran.Controls.Add(Me.txtWartGranMatW)
        Me.TabWartGran.Controls.Add(Me.Label28)
        Me.TabWartGran.Controls.Add(Me.Label25)
        Me.TabWartGran.Controls.Add(Me.txtWartGranP)
        Me.TabWartGran.Controls.Add(Me.Label24)
        Me.TabWartGran.Controls.Add(Me.txtWartGranW)
        Me.TabWartGran.Controls.Add(Me.Label23)
        Me.TabWartGran.Location = New System.Drawing.Point(4, 23)
        Me.TabWartGran.Name = "TabWartGran"
        Me.TabWartGran.Size = New System.Drawing.Size(624, 256)
        Me.TabWartGran.TabIndex = 4
        Me.TabWartGran.Text = "Wart. Graniczna"
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Navy
        Me.Label26.Location = New System.Drawing.Point(8, 95)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(600, 16)
        Me.Label26.TabIndex = 52
        Me.Label26.Text = "Wartoœæ graniczna wyliczona na podstawie sprzeda¿y materia³ów eksploatacyjnych :"
        '
        'txtWartGranMatP
        '
        Me.txtWartGranMatP.ForeColor = System.Drawing.Color.Purple
        Me.txtWartGranMatP.Location = New System.Drawing.Point(407, 139)
        Me.txtWartGranMatP.Name = "txtWartGranMatP"
        Me.txtWartGranMatP.Size = New System.Drawing.Size(108, 20)
        Me.txtWartGranMatP.TabIndex = 51
        '
        'Label27
        '
        Me.Label27.ForeColor = System.Drawing.Color.Navy
        Me.Label27.Location = New System.Drawing.Point(8, 142)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(446, 16)
        Me.Label27.TabIndex = 50
        Me.Label27.Text = "Procent wartoœci sprzeda¿y wykorzystywany do obliczania wartoœci granicznej. . . " &
    ". . . . . . . . . . . . . . . . . . . . . . . . . . ."
        '
        'txtWartGranMatW
        '
        Me.txtWartGranMatW.ForeColor = System.Drawing.Color.Purple
        Me.txtWartGranMatW.Location = New System.Drawing.Point(407, 117)
        Me.txtWartGranMatW.Name = "txtWartGranMatW"
        Me.txtWartGranMatW.Size = New System.Drawing.Size(108, 20)
        Me.txtWartGranMatW.TabIndex = 49
        '
        'Label28
        '
        Me.Label28.ForeColor = System.Drawing.Color.Navy
        Me.Label28.Location = New System.Drawing.Point(8, 120)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(446, 16)
        Me.Label28.TabIndex = 48
        Me.Label28.Text = "Wartoœæ graniczna za poprzedni rok ze sprzeda¿y materia³ów eksploatacyjnych . . ." &
    " . . . . . . . . . . . . . . . . . . . . . . . . . . . . ."
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.Navy
        Me.Label25.Location = New System.Drawing.Point(8, 8)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(600, 16)
        Me.Label25.TabIndex = 47
        Me.Label25.Text = "Wartoœæ graniczna wyliczona na podstawie sprzeda¿y ogó³em :"
        '
        'txtWartGranP
        '
        Me.txtWartGranP.ForeColor = System.Drawing.Color.Purple
        Me.txtWartGranP.Location = New System.Drawing.Point(407, 52)
        Me.txtWartGranP.Name = "txtWartGranP"
        Me.txtWartGranP.Size = New System.Drawing.Size(108, 20)
        Me.txtWartGranP.TabIndex = 46
        '
        'Label24
        '
        Me.Label24.ForeColor = System.Drawing.Color.Navy
        Me.Label24.Location = New System.Drawing.Point(8, 55)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(446, 16)
        Me.Label24.TabIndex = 45
        Me.Label24.Text = "Procent wartoœci sprzeda¿y wykorzystywany do obliczania wartoœci granicznej. . . " &
    ". . . . . . . . . . . . . . . . . . . . . . . . . . ."
        '
        'txtWartGranW
        '
        Me.txtWartGranW.ForeColor = System.Drawing.Color.Purple
        Me.txtWartGranW.Location = New System.Drawing.Point(407, 30)
        Me.txtWartGranW.Name = "txtWartGranW"
        Me.txtWartGranW.Size = New System.Drawing.Size(108, 20)
        Me.txtWartGranW.TabIndex = 44
        '
        'Label23
        '
        Me.Label23.ForeColor = System.Drawing.Color.Navy
        Me.Label23.Location = New System.Drawing.Point(8, 33)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(446, 16)
        Me.Label23.TabIndex = 43
        Me.Label23.Text = "Wartoœæ graniczna ogó³em za poprzedni rok (Wartoœæ) . . . . . . . . . . . . . . ." &
    " . . . . . . . . . . . . . . . . ."
        '
        'Parametry
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(648, 329)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.tbcParametry)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Parametry"
        Me.ShowInTaskbar = False
        Me.Text = "Parametry programu"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbcParametry.ResumeLayout(False)
        Me.Uzytkownik.ResumeLayout(False)
        Me.Uzytkownik.PerformLayout()
        Me.Techniczne.ResumeLayout(False)
        Me.TabPAYBACK.ResumeLayout(False)
        Me.TabPAYBACK.PerformLayout()
        Me.TabRaportAkt.ResumeLayout(False)
        Me.TabRaportAkt.PerformLayout()
        Me.TabWartGran.ResumeLayout(False)
        Me.TabWartGran.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Parametry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        PobierzParametry()
        '-----------------------------------
        WyswietlForme()
    End Sub

    Private Sub Parametry_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        PobierzZFormy()
        '-------------------------------------
        ZapiszParametry(Me)
        Me.Close()
    End Sub




    Private Sub WyswietlForme()
        txtIdentUzytkownika.Text = Par_IdentUzytkownika
        txtNazwaUzytkownika.Text = Par_NazwaUzytkownika
        txtAdresUzytkownika.Text = Par_AdresUzytkownika
        txtBankUzytkownika.Text = Par_BankUzytkownika
        txtKontoUzytkownika.Text = Par_KontoUzytkownika
        txtMiejscowoscUzytkownika.Text = Par_MiejscowoscUzytkownika
        txtNIPUzytkownika.Text = Par_NIPUzytkownika
        txtIdentOddzialu.Text = Par_IdentOddzialu

        lblKatalogOferty.Text = Par_KatalogOferty
        txtWWWPayBack.Text = Par_wwwPAYBACK

        txtEXCEL01.Text = Par_RAKolumna01
        txtEXCEL02.Text = Par_RAKolumna02
        txtEXCEL03.Text = Par_RAKolumna03
        txtEXCEL04.Text = Par_RAKolumna04
        txtEXCEL05.Text = Par_RAKolumna05
        txtEXCEL06.Text = Par_RAKolumna06
        txtEXCEL07.Text = Par_RAKolumna07
        txtEXCEL08.Text = Par_RAKolumna08
        txtEXCEL09.Text = Par_RAKolumna09
        txtEXCEL10.Text = Par_RAKolumna10

        txtWartGranW.Text = Par_WartGranWartosc.ToString("F2")
        txtWartGranP.Text = Par_WartGranProcent.ToString("F0")
        txtWartGranMatW.Text = Par_WartGranMatWartosc.ToString("F2")
        txtWartGranMatP.Text = Par_WartGranMatProcent.ToString("F0")
    End Sub

    Private Sub PobierzZFormy()
        Par_IdentUzytkownika = txtIdentUzytkownika.Text
        Par_NazwaUzytkownika = txtNazwaUzytkownika.Text
        Par_AdresUzytkownika = txtAdresUzytkownika.Text
        Par_BankUzytkownika = txtBankUzytkownika.Text
        Par_KontoUzytkownika = txtKontoUzytkownika.Text
        Par_MiejscowoscUzytkownika = txtMiejscowoscUzytkownika.Text
        Par_NIPUzytkownika = txtNIPUzytkownika.Text
        Par_IdentOddzialu = txtIdentOddzialu.Text

        Par_KatalogOferty = lblKatalogOferty.Text
        Par_wwwPAYBACK = txtWWWPayBack.Text

        Par_RAKolumna01 = txtEXCEL01.Text
        Par_RAKolumna02 = txtEXCEL02.Text
        Par_RAKolumna03 = txtEXCEL03.Text
        Par_RAKolumna04 = txtEXCEL04.Text
        Par_RAKolumna05 = txtEXCEL05.Text
        Par_RAKolumna06 = txtEXCEL06.Text
        Par_RAKolumna07 = txtEXCEL07.Text
        Par_RAKolumna08 = txtEXCEL08.Text
        Par_RAKolumna09 = txtEXCEL09.Text
        Par_RAKolumna10 = txtEXCEL10.Text

        Par_WartGranWartosc = GetDblFromText(txtWartGranW.Text)
        Par_WartGranProcent = GetDblFromText(txtWartGranP.Text)
        Par_WartGranMatWartosc = GetDblFromText(txtWartGranMatW.Text)
        Par_WartGranMatProcent = GetDblFromText(txtWartGranMatP.Text)
    End Sub

    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    Private Sub txtIdentOddzialu_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdentOddzialu.TextChanged
        TestLen(txtIdentOddzialu, "Ident Oddzialu", 50)
    End Sub
    Private Sub txtIdentOddzialu_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdentOddzialu.GotFocus
        StartEdycjiTxt(txtIdentOddzialu)
    End Sub
    Private Sub txtIdentOddzialu_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdentOddzialu.LostFocus
        KoniecEdycjiTxt(txtIdentOddzialu)
    End Sub
    '-------------------------------------------
    Private Sub txtIdentUzytkownika_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdentUzytkownika.TextChanged
        TestLen(txtIdentUzytkownika, "Ident UZYTKOWNIKA", 20)
    End Sub
    Private Sub txtIdentUzytkownika_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdentUzytkownika.GotFocus
        StartEdycjiTxt(txtIdentUzytkownika)
    End Sub
    Private Sub txtIdentUzytkownika_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdentUzytkownika.LostFocus
        KoniecEdycjiTxt(txtIdentUzytkownika)
    End Sub
    '-------------------------------------------
    Private Sub txtNazwaUzytkownika_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazwaUzytkownika.TextChanged
        TestLen(txtNazwaUzytkownika, "NAZWA UZYTKOWNIKA", 100)
    End Sub
    Private Sub txtNazwaUzytkownika_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNazwaUzytkownika.GotFocus
        StartEdycjiTxt(txtNazwaUzytkownika)
    End Sub
    Private Sub txtNazwaUzytkownika_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNazwaUzytkownika.LostFocus
        KoniecEdycjiTxt(txtNazwaUzytkownika)
    End Sub
    '-------------------------------------------
    Private Sub txtAdresUzytkownika_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAdresUzytkownika.TextChanged
        TestLen(txtAdresUzytkownika, "ADRES UZYTKOWNIKA", 100)
    End Sub
    Private Sub txtAdresUzytkownika_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdresUzytkownika.GotFocus
        StartEdycjiTxt(txtAdresUzytkownika)
    End Sub
    Private Sub txtAdresUzytkownika_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdresUzytkownika.LostFocus
        KoniecEdycjiTxt(txtAdresUzytkownika)
    End Sub
    '-------------------------------------------
    Private Sub txtMiejscowoscUzytkownika_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMiejscowoscUzytkownika.TextChanged
        TestLen(txtMiejscowoscUzytkownika, "MIEJSCOWOŒÆ UZYTKOWNIKA", 100)
    End Sub
    Private Sub txtMiejscowoscUzytkownika_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMiejscowoscUzytkownika.GotFocus
        StartEdycjiTxt(txtMiejscowoscUzytkownika)
    End Sub
    Private Sub txtMiejscowoscUzytkownika_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMiejscowoscUzytkownika.LostFocus
        KoniecEdycjiTxt(txtMiejscowoscUzytkownika)
    End Sub
    '-------------------------------------------
    Private Sub txtNIPUzytkownika_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNIPUzytkownika.TextChanged
        TestLen(txtNIPUzytkownika, "NIP UZYTKOWNIKA", 15)
    End Sub
    Private Sub txtNIPUzytkownika_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIPUzytkownika.GotFocus
        StartEdycjiTxt(txtNIPUzytkownika)
    End Sub
    Private Sub txtNIPUzytkownika_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIPUzytkownika.LostFocus
        KoniecEdycjiTxt(txtNIPUzytkownika)
    End Sub
    '-------------------------------------------
    Private Sub txtBankUzytkownika_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBankUzytkownika.TextChanged
        TestLen(txtBankUzytkownika, "BANK UZYTKOWNIKA", 100)
    End Sub
    Private Sub txtBankUzytkownika_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBankUzytkownika.GotFocus
        StartEdycjiTxt(txtBankUzytkownika)
    End Sub
    Private Sub txtBankUzytkownika_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBankUzytkownika.LostFocus
        KoniecEdycjiTxt(txtBankUzytkownika)
    End Sub
    '-------------------------------------------
    Private Sub txtKontoUzytkownika_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKontoUzytkownika.TextChanged
        TestLen(txtKontoUzytkownika, "KONTO UZYTKOWNIKA", 100)
    End Sub
    Private Sub txtKontoUzytkownika_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKontoUzytkownika.GotFocus
        StartEdycjiTxt(txtKontoUzytkownika)
    End Sub
    Private Sub txtKontoUzytkownika_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKontoUzytkownika.LostFocus
        KoniecEdycjiTxt(txtKontoUzytkownika)
    End Sub



    '-------------------------------------------
    Private Sub txtEXCEL01_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL01.TextChanged
        TestLen(txtEXCEL01, "Nag³ówek Kolumny 01", 100)
    End Sub
    Private Sub txtEXCEL01_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL01.GotFocus
        StartEdycjiTxt(txtEXCEL01)
    End Sub
    Private Sub txtEXCEL01_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL01.LostFocus
        KoniecEdycjiTxt(txtEXCEL01)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL02_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL02.TextChanged
        TestLen(txtEXCEL02, "Nag³ówek Kolumny 02", 100)
    End Sub
    Private Sub txtEXCEL02_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL02.GotFocus
        StartEdycjiTxt(txtEXCEL02)
    End Sub
    Private Sub txtEXCEL02_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL02.LostFocus
        KoniecEdycjiTxt(txtEXCEL02)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL03_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL03.TextChanged
        TestLen(txtEXCEL03, "Nag³ówek Kolumny 03", 100)
    End Sub
    Private Sub txtEXCEL03_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL03.GotFocus
        StartEdycjiTxt(txtEXCEL03)
    End Sub
    Private Sub txtEXCEL03_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL03.LostFocus
        KoniecEdycjiTxt(txtEXCEL03)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL04_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL04.TextChanged
        TestLen(txtEXCEL04, "Nag³ówek Kolumny 04", 100)
    End Sub
    Private Sub txtEXCEL04_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL04.GotFocus
        StartEdycjiTxt(txtEXCEL04)
    End Sub
    Private Sub txtEXCEL04_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL04.LostFocus
        KoniecEdycjiTxt(txtEXCEL04)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL05_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL05.TextChanged
        TestLen(txtEXCEL05, "Nag³ówek Kolumny 05", 100)
    End Sub
    Private Sub txtEXCEL05_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL05.GotFocus
        StartEdycjiTxt(txtEXCEL05)
    End Sub
    Private Sub txtEXCEL05_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL05.LostFocus
        KoniecEdycjiTxt(txtEXCEL05)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL06_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL06.TextChanged
        TestLen(txtEXCEL06, "Nag³ówek Kolumny 06", 100)
    End Sub
    Private Sub txtEXCEL06_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL06.GotFocus
        StartEdycjiTxt(txtEXCEL06)
    End Sub
    Private Sub txtEXCEL06_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL06.LostFocus
        KoniecEdycjiTxt(txtEXCEL06)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL07_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL07.TextChanged
        TestLen(txtEXCEL07, "Nag³ówek Kolumny 07", 100)
    End Sub
    Private Sub txtEXCEL07_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL07.GotFocus
        StartEdycjiTxt(txtEXCEL07)
    End Sub
    Private Sub txtEXCEL07_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL07.LostFocus
        KoniecEdycjiTxt(txtEXCEL07)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL08_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL08.TextChanged
        TestLen(txtEXCEL08, "Nag³ówek Kolumny 08", 100)
    End Sub
    Private Sub txtEXCEL08_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL08.GotFocus
        StartEdycjiTxt(txtEXCEL08)
    End Sub
    Private Sub txtEXCEL08_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL08.LostFocus
        KoniecEdycjiTxt(txtEXCEL08)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL09_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL09.TextChanged
        TestLen(txtEXCEL09, "Nag³ówek Kolumny 09", 100)
    End Sub
    Private Sub txtEXCEL09_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL09.GotFocus
        StartEdycjiTxt(txtEXCEL09)
    End Sub
    Private Sub txtEXCEL09_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL09.LostFocus
        KoniecEdycjiTxt(txtEXCEL09)
    End Sub
    '-------------------------------------------
    Private Sub txtEXCEL10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL10.TextChanged
        TestLen(txtEXCEL10, "Nag³ówek Kolumny 10", 100)
    End Sub
    Private Sub txtEXCEL10_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL10.GotFocus
        StartEdycjiTxt(txtEXCEL10)
    End Sub
    Private Sub txtEXCEL10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL10.LostFocus
        KoniecEdycjiTxt(txtEXCEL10)
    End Sub




    Private Sub txtWartGranW_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWartGranW.TextChanged
        If TestNum(txtWartGranW, "WARTOSC GRANICZNA-WARTOSC") Then
        End If
    End Sub
    Private Sub txtWartGranW_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWartGranW.GotFocus
        StartEdycjiTxt(txtWartGranW)
    End Sub
    Private Sub txtWartGranW_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWartGranW.LostFocus
        KoniecEdycjiTxt(txtWartGranW)
    End Sub

    Private Sub txtWartGranP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWartGranP.TextChanged
        If TestNum(txtWartGranP, "WARTOSC GRANICZNA-PROCENT") Then
        End If
    End Sub
    Private Sub txtWartGranP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWartGranP.GotFocus
        StartEdycjiTxt(txtWartGranP)
    End Sub
    Private Sub txtWartGranP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWartGranP.LostFocus
        KoniecEdycjiTxt(txtWartGranP)
    End Sub





    Private Sub txtWartGranMatW_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWartGranMatW.TextChanged
        If TestNum(txtWartGranMatW, "WARTOSC GRANICZNA-WARTOSC") Then
        End If
    End Sub
    Private Sub txtWartGranMatW_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWartGranMatW.GotFocus
        StartEdycjiTxt(txtWartGranMatW)
    End Sub
    Private Sub txtWartGranMatW_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWartGranMatW.LostFocus
        KoniecEdycjiTxt(txtWartGranMatW)
    End Sub

    Private Sub txtWartGranMatP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWartGranMatP.TextChanged
        If TestNum(txtWartGranMatP, "WARTOSC GRANICZNA-PROCENT") Then
        End If
    End Sub
    Private Sub txtWartGranMatP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWartGranMatP.GotFocus
        StartEdycjiTxt(txtWartGranMatP)
    End Sub
    Private Sub txtWartGranMatP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWartGranMatP.LostFocus
        KoniecEdycjiTxt(txtWartGranMatP)
    End Sub








    Private Sub cmdWybierzHurt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierzKatalog.Click
        With dlgOpenFolder
            .Description = "Proszê wybraæ katalog dyskowy przeznaczony na pliki Ofert dls klientów"
            '.RootFolder = System.Environment.SpecialFolder.Programs
            .ShowNewFolderButton = True

            'Me.Visible = False
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                lblKatalogOferty.Text = Mid(.SelectedPath, 1, 100)
            End If
        End With
    End Sub



    Private Sub cmdDelHurt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelKatalog.Click
        lblKatalogOferty.Text = ""
    End Sub


    Private Sub cmdPokazWWW_Click(sender As Object, e As EventArgs) Handles cmdPokazWWW.Click
        System.Diagnostics.Process.Start(
       "C:\Program Files\Internet Explorer\IExplore.exe",
       Trim(txtWWWPayBack.Text))
    End Sub

End Class
