Public Class KatalogKontaktow
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Dim _IdentKlienta As String = ""
    Public Sub New(ByVal pIdentK As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _IdentKlienta = pIdentK
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
    Friend WithEvents stbKontakty As System.Windows.Forms.StatusBar
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblUlica As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblTOFI As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblUwagi As System.Windows.Forms.Label
    Friend WithEvents lbleMail As System.Windows.Forms.Label
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents lblTelefon As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblMiejscowosc As System.Windows.Forms.Label
    Friend WithEvents lblNrDomu As System.Windows.Forms.Label
    Friend WithEvents lblNIP As System.Windows.Forms.Label
    Friend WithEvents lblAdres As System.Windows.Forms.Label
    Friend WithEvents lblKod As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa3 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa2 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa1 As System.Windows.Forms.Label
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblData As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lblTemat As System.Windows.Forms.Label
    Friend WithEvents lblNotatka As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents lblIDDomu As System.Windows.Forms.Label
    Friend WithEvents cmdAnaliza As System.Windows.Forms.Button
    Friend WithEvents HorizScrollTimer As Timer
    Friend WithEvents menuWybieraj As ContextMenuStrip
    Friend WithEvents menuWEdytuj As ToolStripMenuItem
    Friend WithEvents menuWDodaj As ToolStripMenuItem
    Friend WithEvents menuWUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents menuWWybierz As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents menuWSingleL As ToolStripMenuItem
    Friend WithEvents menuWMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents menuWOdswiez As ToolStripMenuItem
    Friend WithEvents menuEdytuj As ContextMenuStrip
    Friend WithEvents menuEEdytuj As ToolStripMenuItem
    Friend WithEvents menuEDodaj As ToolStripMenuItem
    Friend WithEvents menuEUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystko As Button
    Friend WithEvents stbFiltry As StatusBar
    Friend WithEvents dagKontakty As DataGridView
    Friend WithEvents lblNazwa4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogKontaktow))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.stbKontakty = New System.Windows.Forms.StatusBar()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblNazwa4 = New System.Windows.Forms.Label()
        Me.lblIDDomu = New System.Windows.Forms.Label()
        Me.lblNrDomu = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.lblNotatka = New System.Windows.Forms.Label()
        Me.lblTemat = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblData = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblUlica = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblTOFI = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lblUwagi = New System.Windows.Forms.Label()
        Me.lbleMail = New System.Windows.Forms.Label()
        Me.lblFax = New System.Windows.Forms.Label()
        Me.lblTelefon = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblMiejscowosc = New System.Windows.Forms.Label()
        Me.lblNIP = New System.Windows.Forms.Label()
        Me.lblAdres = New System.Windows.Forms.Label()
        Me.lblKod = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblNazwa3 = New System.Windows.Forms.Label()
        Me.lblNazwa2 = New System.Windows.Forms.Label()
        Me.lblNazwa1 = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmdAnaliza = New System.Windows.Forms.Button()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.menuWybieraj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuWEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuWDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuWUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWWybierz = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWSingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuWMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEdytuj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuESingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.dagKontakty = New System.Windows.Forms.DataGridView()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.menuWybieraj.SuspendLayout()
        Me.menuEdytuj.SuspendLayout()
        CType(Me.dagKontakty, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'stbKontakty
        '
        Me.stbKontakty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbKontakty.Dock = System.Windows.Forms.DockStyle.None
        Me.stbKontakty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbKontakty.Location = New System.Drawing.Point(8, 448)
        Me.stbKontakty.Name = "stbKontakty"
        Me.stbKontakty.ShowPanels = True
        Me.stbKontakty.Size = New System.Drawing.Size(888, 23)
        Me.stbKontakty.TabIndex = 67
        Me.stbKontakty.Text = "StatusBar1"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 480)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 25)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 64
        Me.PictureBox1.TabStop = False
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(800, 480)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 24)
        Me.cmdStop.TabIndex = 60
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.ForeColor = System.Drawing.Color.Black
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(712, 480)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 24)
        Me.cmdEdytuj.TabIndex = 63
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.ForeColor = System.Drawing.Color.Black
        Me.cmdUsuñ.Image = CType(resources.GetObject("cmdUsuñ.Image"), System.Drawing.Image)
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(624, 480)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 24)
        Me.cmdUsuñ.TabIndex = 62
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.ForeColor = System.Drawing.Color.Black
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(536, 480)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 24)
        Me.cmdDodaj.TabIndex = 61
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblNazwa4)
        Me.Panel1.Controls.Add(Me.lblIDDomu)
        Me.Panel1.Controls.Add(Me.lblNrDomu)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.lblNotatka)
        Me.Panel1.Controls.Add(Me.lblTemat)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lblData)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.lblUlica)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblTOFI)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.lblUwagi)
        Me.Panel1.Controls.Add(Me.lbleMail)
        Me.Panel1.Controls.Add(Me.lblFax)
        Me.Panel1.Controls.Add(Me.lblTelefon)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.lblMiejscowosc)
        Me.Panel1.Controls.Add(Me.lblNIP)
        Me.Panel1.Controls.Add(Me.lblAdres)
        Me.Panel1.Controls.Add(Me.lblKod)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblNazwa3)
        Me.Panel1.Controls.Add(Me.lblNazwa2)
        Me.Panel1.Controls.Add(Me.lblNazwa1)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(888, 208)
        Me.Panel1.TabIndex = 58
        '
        'lblNazwa4
        '
        Me.lblNazwa4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa4.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa4.Location = New System.Drawing.Point(352, 72)
        Me.lblNazwa4.Name = "lblNazwa4"
        Me.lblNazwa4.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwa4.TabIndex = 37
        '
        'lblIDDomu
        '
        Me.lblIDDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIDDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIDDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblIDDomu.Location = New System.Drawing.Point(537, 24)
        Me.lblIDDomu.Name = "lblIDDomu"
        Me.lblIDDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblIDDomu.TabIndex = 36
        '
        'lblNrDomu
        '
        Me.lblNrDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNrDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNrDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblNrDomu.Location = New System.Drawing.Point(494, 24)
        Me.lblNrDomu.Name = "lblNrDomu"
        Me.lblNrDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblNrDomu.TabIndex = 11
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Navy
        Me.Label17.Location = New System.Drawing.Point(8, 120)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(56, 16)
        Me.Label17.TabIndex = 35
        Me.Label17.Text = "Opis . . . . . . . . "
        '
        'lblNotatka
        '
        Me.lblNotatka.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNotatka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNotatka.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNotatka.ForeColor = System.Drawing.Color.Purple
        Me.lblNotatka.Location = New System.Drawing.Point(64, 120)
        Me.lblNotatka.Name = "lblNotatka"
        Me.lblNotatka.Size = New System.Drawing.Size(816, 80)
        Me.lblNotatka.TabIndex = 34
        '
        'lblTemat
        '
        Me.lblTemat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTemat.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTemat.ForeColor = System.Drawing.Color.Purple
        Me.lblTemat.Location = New System.Drawing.Point(254, 104)
        Me.lblTemat.Name = "lblTemat"
        Me.lblTemat.Size = New System.Drawing.Size(626, 16)
        Me.lblTemat.TabIndex = 32
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(160, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(118, 16)
        Me.Label9.TabIndex = 33
        Me.Label9.Text = "Rodzaj kontaktu . . . . . . . . "
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(8, 104)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 16)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Data . . . . . . . . "
        '
        'lblData
        '
        Me.lblData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblData.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblData.ForeColor = System.Drawing.Color.Purple
        Me.lblData.Location = New System.Drawing.Point(64, 104)
        Me.lblData.Name = "lblData"
        Me.lblData.Size = New System.Drawing.Size(80, 16)
        Me.lblData.TabIndex = 30
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 96)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(880, 1)
        Me.Panel2.TabIndex = 29
        '
        'lblUlica
        '
        Me.lblUlica.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUlica.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUlica.ForeColor = System.Drawing.Color.Purple
        Me.lblUlica.Location = New System.Drawing.Point(352, 24)
        Me.lblUlica.Name = "lblUlica"
        Me.lblUlica.Size = New System.Drawing.Size(144, 16)
        Me.lblUlica.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(312, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Ulica . . "
        '
        'lblTOFI
        '
        Me.lblTOFI.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTOFI.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTOFI.ForeColor = System.Drawing.Color.Purple
        Me.lblTOFI.Location = New System.Drawing.Point(216, 8)
        Me.lblTOFI.Name = "lblTOFI"
        Me.lblTOFI.Size = New System.Drawing.Size(80, 24)
        Me.lblTOFI.TabIndex = 27
        Me.lblTOFI.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(160, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Nr TOFI . . . . . . . . "
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Nazwa . . . . . . . . "
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Navy
        Me.Label14.Location = New System.Drawing.Point(592, 56)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(56, 16)
        Me.Label14.TabIndex = 24
        Me.Label14.Text = "Uwagi"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(592, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(56, 16)
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "eMail"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(592, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 16)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "Fax"
        '
        'lblUwagi
        '
        Me.lblUwagi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUwagi.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUwagi.ForeColor = System.Drawing.Color.Purple
        Me.lblUwagi.Location = New System.Drawing.Point(648, 56)
        Me.lblUwagi.Name = "lblUwagi"
        Me.lblUwagi.Size = New System.Drawing.Size(232, 16)
        Me.lblUwagi.TabIndex = 20
        '
        'lbleMail
        '
        Me.lbleMail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbleMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lbleMail.ForeColor = System.Drawing.Color.Purple
        Me.lbleMail.Location = New System.Drawing.Point(648, 40)
        Me.lbleMail.Name = "lbleMail"
        Me.lbleMail.Size = New System.Drawing.Size(232, 16)
        Me.lbleMail.TabIndex = 18
        '
        'lblFax
        '
        Me.lblFax.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblFax.ForeColor = System.Drawing.Color.Purple
        Me.lblFax.Location = New System.Drawing.Point(648, 24)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(232, 16)
        Me.lblFax.TabIndex = 17
        '
        'lblTelefon
        '
        Me.lblTelefon.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTelefon.ForeColor = System.Drawing.Color.Purple
        Me.lblTelefon.Location = New System.Drawing.Point(648, 8)
        Me.lblTelefon.Name = "lblTelefon"
        Me.lblTelefon.Size = New System.Drawing.Size(232, 16)
        Me.lblTelefon.TabIndex = 16
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(592, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 16)
        Me.Label10.TabIndex = 15
        Me.Label10.Text = "Telefon"
        '
        'lblMiejscowosc
        '
        Me.lblMiejscowosc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMiejscowosc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblMiejscowosc.ForeColor = System.Drawing.Color.Purple
        Me.lblMiejscowosc.Location = New System.Drawing.Point(424, 8)
        Me.lblMiejscowosc.Name = "lblMiejscowosc"
        Me.lblMiejscowosc.Size = New System.Drawing.Size(158, 16)
        Me.lblMiejscowosc.TabIndex = 12
        '
        'lblNIP
        '
        Me.lblNIP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNIP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNIP.ForeColor = System.Drawing.Color.Purple
        Me.lblNIP.Location = New System.Drawing.Point(64, 72)
        Me.lblNIP.Name = "lblNIP"
        Me.lblNIP.Size = New System.Drawing.Size(232, 16)
        Me.lblNIP.TabIndex = 10
        '
        'lblAdres
        '
        Me.lblAdres.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdres.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAdres.ForeColor = System.Drawing.Color.Purple
        Me.lblAdres.Location = New System.Drawing.Point(352, 40)
        Me.lblAdres.Name = "lblAdres"
        Me.lblAdres.Size = New System.Drawing.Size(232, 16)
        Me.lblAdres.TabIndex = 9
        '
        'lblKod
        '
        Me.lblKod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKod.ForeColor = System.Drawing.Color.Purple
        Me.lblKod.Location = New System.Drawing.Point(352, 8)
        Me.lblKod.Name = "lblKod"
        Me.lblKod.Size = New System.Drawing.Size(64, 16)
        Me.lblKod.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(312, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Adres"
        '
        'lblNazwa3
        '
        Me.lblNazwa3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa3.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa3.Location = New System.Drawing.Point(352, 56)
        Me.lblNazwa3.Name = "lblNazwa3"
        Me.lblNazwa3.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwa3.TabIndex = 4
        '
        'lblNazwa2
        '
        Me.lblNazwa2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2.Location = New System.Drawing.Point(64, 56)
        Me.lblNazwa2.Name = "lblNazwa2"
        Me.lblNazwa2.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwa2.TabIndex = 3
        '
        'lblNazwa1
        '
        Me.lblNazwa1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1.Location = New System.Drawing.Point(64, 40)
        Me.lblNazwa1.Name = "lblNazwa1"
        Me.lblNazwa1.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwa1.TabIndex = 2
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(64, 8)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(80, 24)
        Me.lblIdent.TabIndex = 1
        Me.lblIdent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Klient . . . . . . . . "
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "NIP"
        '
        'cmdAnaliza
        '
        Me.cmdAnaliza.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAnaliza.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdAnaliza.ForeColor = System.Drawing.Color.Black
        Me.cmdAnaliza.Image = CType(resources.GetObject("cmdAnaliza.Image"), System.Drawing.Image)
        Me.cmdAnaliza.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAnaliza.Location = New System.Drawing.Point(362, 480)
        Me.cmdAnaliza.Name = "cmdAnaliza"
        Me.cmdAnaliza.Size = New System.Drawing.Size(168, 22)
        Me.cmdAnaliza.TabIndex = 68
        Me.cmdAnaliza.Text = "Analiza czest.kontaktów"
        Me.cmdAnaliza.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'HorizScrollTimer
        '
        '
        'menuWybieraj
        '
        Me.menuWybieraj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuWEdytuj, Me.menuWDodaj, Me.menuWUsun, Me.ToolStripSeparator1, Me.menuWWybierz, Me.ToolStripSeparator4, Me.menuWSingleL, Me.menuWMultiL, Me.ToolStripSeparator5, Me.menuWOdswiez})
        Me.menuWybieraj.Name = "menuWybieraj"
        Me.menuWybieraj.Size = New System.Drawing.Size(187, 176)
        '
        'menuWEdytuj
        '
        Me.menuWEdytuj.Name = "menuWEdytuj"
        Me.menuWEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.menuWEdytuj.Text = "Edytuj"
        '
        'menuWDodaj
        '
        Me.menuWDodaj.Name = "menuWDodaj"
        Me.menuWDodaj.Size = New System.Drawing.Size(186, 22)
        Me.menuWDodaj.Text = "Dodaj"
        '
        'menuWUsun
        '
        Me.menuWUsun.Name = "menuWUsun"
        Me.menuWUsun.Size = New System.Drawing.Size(186, 22)
        Me.menuWUsun.Text = "Usuñ"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(183, 6)
        '
        'menuWWybierz
        '
        Me.menuWWybierz.Name = "menuWWybierz"
        Me.menuWWybierz.Size = New System.Drawing.Size(186, 22)
        Me.menuWWybierz.Text = "Wybierz"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(183, 6)
        '
        'menuWSingleL
        '
        Me.menuWSingleL.Name = "menuWSingleL"
        Me.menuWSingleL.Size = New System.Drawing.Size(186, 22)
        Me.menuWSingleL.Text = "Poka¿ w jednej linii"
        '
        'menuWMultiL
        '
        Me.menuWMultiL.Name = "menuWMultiL"
        Me.menuWMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuWMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(183, 6)
        '
        'menuWOdswiez
        '
        Me.menuWOdswiez.Name = "menuWOdswiez"
        Me.menuWOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuWOdswiez.Text = "Odswie¿ Tabelê"
        '
        'menuEdytuj
        '
        Me.menuEdytuj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuEEdytuj, Me.menuEDodaj, Me.menuEUsun, Me.ToolStripSeparator2, Me.menuESingleL, Me.menuEMultiL, Me.ToolStripSeparator3, Me.menuEOdswiez})
        Me.menuEdytuj.Name = "menuEdytuj"
        Me.menuEdytuj.Size = New System.Drawing.Size(187, 148)
        '
        'menuEEdytuj
        '
        Me.menuEEdytuj.Name = "menuEEdytuj"
        Me.menuEEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.menuEEdytuj.Text = "Edytuj"
        '
        'menuEDodaj
        '
        Me.menuEDodaj.Name = "menuEDodaj"
        Me.menuEDodaj.Size = New System.Drawing.Size(186, 22)
        Me.menuEDodaj.Text = "Dodaj"
        '
        'menuEUsun
        '
        Me.menuEUsun.Name = "menuEUsun"
        Me.menuEUsun.Size = New System.Drawing.Size(186, 22)
        Me.menuEUsun.Text = "Usuñ"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(183, 6)
        '
        'menuESingleL
        '
        Me.menuESingleL.Name = "menuESingleL"
        Me.menuESingleL.Size = New System.Drawing.Size(186, 22)
        Me.menuESingleL.Text = "Poka¿ w jednej linii"
        '
        'menuEMultiL
        '
        Me.menuEMultiL.Name = "menuEMultiL"
        Me.menuEMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuEMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(183, 6)
        '
        'menuEOdswiez
        '
        Me.menuEOdswiez.Name = "menuEOdswiez"
        Me.menuEOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuEOdswiez.Text = "Odœwie¿ Tabelê"
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(546, 224)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 186
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(603, 222)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystko.TabIndex = 184
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 222)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(888, 22)
        Me.stbFiltry.TabIndex = 185
        Me.stbFiltry.Text = "StatusBar1"
        '
        'dagKontakty
        '
        Me.dagKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagKontakty.ContextMenuStrip = Me.menuEdytuj
        Me.dagKontakty.Location = New System.Drawing.Point(8, 244)
        Me.dagKontakty.Name = "dagKontakty"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagKontakty.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagKontakty.Size = New System.Drawing.Size(888, 202)
        Me.dagKontakty.TabIndex = 187
        '
        'KatalogKontaktow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(904, 511)
        Me.Controls.Add(Me.dagKontakty)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.cmdAnaliza)
        Me.Controls.Add(Me.stbKontakty)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "KatalogKontaktow"
        Me.ShowInTaskbar = False
        Me.Text = "Katalog Kontaktów Handlowych"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.menuWybieraj.ResumeLayout(False)
        Me.menuEdytuj.ResumeLayout(False)
        CType(Me.dagKontakty, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim StartFormy As Boolean = True
    Dim _OCoMamRobic As String = oCoMamRobic

    '---------------------------------------------
    'SCROLL-parametry obslugi scrollingu poziomego
    Dim StartPointInCell00 As System.Drawing.Point             'pocz¹tek ektanu przed scrollingiem
    Dim ScrollXOffset As Integer = 0            'bie¿¹cy punkt ekranu wskazywany przez kursor podczas scrollingu
    Dim HorizScrollLastTime As String = ""      'Chwila rozpoczecia scrollingu - filtry wyswietlaner bêd¹ po 1 sekundzie...
    ' HorizScrollTimer - zegar odliczaj¹cy czas od poczatku scrollingu
    '---------------------------------------------


    Dim dbSelectSQLKontakty As String = ""

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbCommandDelete As OleDb.OleDbCommand
    Dim dbCommandUpdate As OleDb.OleDbCommand
    Dim dbCommandInsert As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As SqlClient.SqlConnection
    Dim sqlCommandSelect As SqlClient.SqlCommand
    Dim sqlCommandDelete As SqlClient.SqlCommand
    Dim sqlCommandUpdate As SqlClient.SqlCommand
    Dim sqlCommandInsert As SqlClient.SqlCommand
    Dim sqlDataAdapter As SqlClient.SqlDataAdapter

    'DataSet
    Dim DataSetKontakty As New DataSet
    Dim DataViewKontakty As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim _CoMamRobic As String = ""

    Private Sub KatalogOsob_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        _CoMamRobic = oCoMamRobic

        IdentUzytkownika(Program_IdUzytkownika)
        If OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownik Then
            cmdUsuñ.Enabled = False
            menuEUsun.Enabled = False
            menuWUsun.Enabled = False

            cmdEdytuj.Text = "Poka¿"
            menuEdytuj.Text = "Poka¿"
        End If


        '--------------------------------
        If Len(_IdentKlienta) = 0 Then
            cmdDodaj.Enabled = False
            dbSelectSQLKontakty = sSelectSQLKontakty & " ORDER BY IDENTKLIENTA, DATAKONTAKTU DESC"
        Else
            cmdDodaj.Enabled = True
            dbSelectSQLKontakty = sSelectSQLKontakty & " WHERE IDENTKLIENTA='" & _IdentKlienta & "' ORDER BY DATAKONTAKTU DESC"
        End If

        'DataSet
        DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")




        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectSQLKontakty, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Kontakty")

            'DBDeleteKontaktyHandlowe(dbCommandDelete, dbDataAdapter)
            'DBUpdateKontaktyHandlowe(dbCommandUpdate, dbDataAdapter)
            'DBInsertKontaktyHandlowe(dbCommandInsert, dbDataAdapter)


            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnection.State
            'End Try
        Else
            sqlConnection = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelect = New SqlClient.SqlCommand(dbSelectSQLKontakty, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLKontakty, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLKontakty, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLKontakty, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Kontakty")

            SQLDeleteKontaktyHandlowe(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateKontaktyHandlowe(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertKontaktyHandlowe(sqlCommandInsert, sqlDataAdapter)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapter.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnection.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapter.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    'dbDataAdapter.Fill(DataSetKontakty)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    sqlDataAdapter.Fill(DataSetKontakty)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("UNIQID")}

                'definiuj DataView
                DataViewKontakty = New DataView(DataSetKontakty.Tables(0))
                DataViewKontakty.AllowDelete = False
                DataViewKontakty.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagKontakty.BackgroundColor = System.Drawing.SystemColors.Control
                dagKontakty.GridColor = System.Drawing.SystemColors.ControlDark
                dagKontakty.ForeColor = System.Drawing.Color.Purple
                dagKontakty.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagKontakty.BorderStyle = BorderStyle.Fixed3D
                'dagKontakty.Dock = DockStyle.Fill
                dagKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagKontakty.AllowUserToAddRows = False
                dagKontakty.AllowUserToDeleteRows = False
                dagKontakty.AllowUserToOrderColumns = True
                dagKontakty.AllowUserToResizeColumns = True
                dagKontakty.AllowUserToResizeRows = True
                dagKontakty.ReadOnly = True
                dagKontakty.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagKontakty.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagKontakty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagKontakty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagKontakty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagKontakty.ColumnHeadersVisible = True
                dagKontakty.ColumnHeadersHeight = 40
                dagKontakty.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagKontakty.RowHeadersVisible = True
                dagKontakty.RowHeadersWidth = 50
                dagKontakty.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagKontakty.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagKontakty.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagKontakty.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagKontakty.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagKontakty.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagKontakty.DefaultCellStyle.NullValue = ""
                dagKontakty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.True         'multi


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagKontakty.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagKontakty.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagKontakty.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagKontakty.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagKontakty.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagKontakty.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagKontakty.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagKontakty.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagKontakty.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagKontakty.DataMember = DataSetKontakty.Tables(0).TableName
                'dagKontakty.DataSource = DataViewKontakty
                '-----------------------------------

                TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(1).ColumnName, "Klient", 60, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(2).ColumnName, "Data", 70, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(3).ColumnName, "Pracownik", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(4).ColumnName, "Rodzaj kontaktu", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(5).ColumnName, "Opis", 750, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(6).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                Me.Cursor = Cursors.WaitCursor
                dagKontakty.DataSource = DataViewKontakty
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetKontakty.Tables(0).Rows.Count > 0 Then
                    dagKontakty.CurrentCell = dagKontakty(1, 0)
                    dagKontakty.CurrentCell.Selected = True
                End If

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            '------------------------------------
            'dodaj StatusBar na koncu
            stbKontakty.BackColor = System.Drawing.SystemColors.Control
            stbKontakty.ForeColor = System.Drawing.Color.DarkGreen
            stbKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbKontakty.Panels.Add("stbPanelIleRec")
            stbKontakty.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbKontakty.Panels(0).Width = 80
            stbKontakty.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKontakty.Panels(0).Text = "Iloœæ rec.: " & DataSetKontakty.Tables(0).Rows.Count.ToString

            stbKontakty.Panels.Add("stbPanelWiersz")
            stbKontakty.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbKontakty.Panels(1).Width = 80
            stbKontakty.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagKontakty.CurrentCell Is Nothing Then
                stbKontakty.Panels(1).Text = "Wiersz : "
            Else
                stbKontakty.Panels(1).Text = "Wiersz : " & dagKontakty.CurrentCell.RowIndex.ToString
            End If

            stbKontakty.Panels.Add("stbPanelKolumna")
            stbKontakty.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbKontakty.Panels(2).Width = 80
            stbKontakty.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagKontakty.CurrentCell Is Nothing Then
                stbKontakty.Panels(2).Text = "Kolumna : "
            Else
                stbKontakty.Panels(2).Text = "Kolumna : " & dagKontakty.CurrentCell.ColumnIndex.ToString
            End If

            stbKontakty.Panels.Add("stbPanelSort")
            stbKontakty.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbKontakty.Panels(3).Width = 150
            stbKontakty.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKontakty.Panels(3).Text = "Sort : "

            stbKontakty.Panels.Add("stbPanelFiltr")
            stbKontakty.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbKontakty.Panels(4).Width = 150
            stbKontakty.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKontakty.Panels(4).Text = "Filtr : "

            stbKontakty.Panels.Add("stbPanelReszta")
            stbKontakty.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbKontakty.Panels(5).Width = 150
            stbKontakty.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKontakty.Panels(5).Text = ""

            stbKontakty.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagKontakty.TableStyles(0).RowHeaderWidth
            stbFiltry.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbFiltry.Panels(0).Text = "Filtry"

            stbFiltry.Panels.Add("stbFiltryReszta")
            stbFiltry.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
            'stbFiltry.Panels(1).Width = 150
            stbFiltry.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbFiltry.Panels(1).Text = ""

            stbFiltry.ShowPanels = True

            'dodaj ctrl do pobierania szablonu
            'cmdWszystko.Size = New Size(20, 20)
            'cmdWszystko.Location = New system.drawing.Point(stbFiltry.Location.X + _
            '                                 stbFiltry.Size.Width - _
            '                                 cmdWszystko.Size.Width, _
            '                                 stbFiltry.Location.Y + 2)
            'cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
            '                            System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            cmdWszystko.Size = New Size(20, 22)
            cmdWszystko.Location = New System.Drawing.Point(stbFiltry.Location.X + 50 - 20,
                                             stbFiltry.Location.Y)
            cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

        End If
        InitKontakty() 'inicjuj zmienne
        '-----------------------------------------------------------------
        If _OCoMamRobic = "WYBIERAJ" Then
            dagKontakty.ContextMenuStrip = Me.menuWybieraj
            cmdStop.Text = "Wybierz"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = True
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = False
        Else
            dagKontakty.ContextMenuStrip = Me.menuEdytuj
            cmdStop.Text = "Powrót"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = False
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = True
        End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagKontakty.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagKontakty.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetKontakty.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagKontakty.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagKontakty.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagKontakty.FirstDisplayedScrollingColumnIndex +
                        dagKontakty.GetCellDisplayRectangle(dagKontakty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagKontakty.Left + 4), (dagKontakty.Top + 4))
            ScrollXOffset = 0
        End If
        '--------------------------------------------------
        'inicjowanie zegara dla scrollingu poziomego
        HorizScrollLastTime = ""
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
        '--------------------------------------------------
        'Stworz zmienne do filtrowania
        Dim ii As Integer = 0
        For ii = 0 To DataViewKontakty.Table.Columns.Count - 1
            'stworz zmienn¹ TXTBOX
            Dim CtrlT As New System.Windows.Forms.TextBox
            CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlT.Visible = False
            Me.Controls.Add(CtrlT)
            AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXX_TextChanged)
            AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXX_GotFocus)
            AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXX_LostFocus)
            AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

            'stworz zmienn¹ BUTTON
            Dim CtrlB As New System.Windows.Forms.Button
            CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlB.Image = cmdFi.Image
            CtrlB.ImageAlign = ContentAlignment.MiddleCenter
            CtrlB.Visible = False
            Me.Controls.Add(CtrlB)
            AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXX_Click)
            AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXX_GotFocus)
            AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXX_LostFocus)
        Next
        Me.Refresh()
        '--------------------------------------------------
        StartFormy = False
        dagKontakty_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub

    Private Sub Katalogkontakty_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub







    Private Sub dagKontakty_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKontakty.CurrentCellChanged
        If Not StartFormy Then
            If dagKontakty.CurrentCell Is Nothing Then
                stbKontakty.Panels(1).Text = "Wi: "
                stbKontakty.Panels(2).Text = "Ko: "

                lblIdent.Text = ""
                lblTOFI.Text = ""
                lblNazwa1.Text = ""
                lblNazwa2.Text = ""
                lblNazwa3.Text = ""

                lblKod.Text = ""
                lblMiejscowosc.Text = ""
                lblUlica.Text = ""
                lblNrDomu.Text = ""
                lblIDDomu.Text = ""
                lblAdres.Text = ""
                lblNIP.Text = ""

                lblTelefon.Text = ""
                lblFax.Text = ""
                lbleMail.Text = ""
                lblUwagi.Text = ""

                lblData.Text = ""
                lblTemat.Text = ""
                lblNotatka.Text = ""
            Else
                stbKontakty.Panels(1).Text = "Wi: " & dagKontakty.CurrentCell.RowIndex.ToString
                stbKontakty.Panels(2).Text = "Ko: " & dagKontakty.CurrentCell.ColumnIndex.ToString

                If DataViewKontakty.Count = 0 Then
                    lblIdent.Text = ""
                    lblTOFI.Text = ""
                    lblNazwa1.Text = ""
                    lblNazwa2.Text = ""
                    lblNazwa3.Text = ""

                    lblKod.Text = ""
                    lblMiejscowosc.Text = ""
                    lblUlica.Text = ""
                    lblNrDomu.Text = ""
                    lblIDDomu.Text = ""
                    lblAdres.Text = ""
                    lblNIP.Text = ""

                    lblTelefon.Text = ""
                    lblFax.Text = ""
                    lbleMail.Text = ""
                    lblUwagi.Text = ""

                    lblData.Text = ""
                    lblTemat.Text = ""
                    lblNotatka.Text = ""
                Else
                    'Public oIdentOso As String            'IDENTKLIENTA     Text, 6
                    'Public oOsobaOso As String            'OSOBA            Text, 50
                    'Public oStanowiskoOso As String       'STANOWISKO       Text, 50
                    'Public oKompetencjeOso As String      'KOMPETENCJE      Text, 50
                    'Public oTelefonOso As String          'TELEFON          Text, 50
                    'Public oeMailOso As String            'EMAIL            Text, 50
                    'Public oVIPOso As Boolean              'VIP              boolean
                    'Public oWersjaOso As Integer          'WERSJA           Integer

                    IdentKlienta(GetTxtField(dagKontakty, 1))

                    lblIdent.Text = GetTxtField(dagKontakty, 1)
                    lblTOFI.Text = oNrTOFITxtKli
                    lblNazwa1.Text = ""
                    lblNazwa2.Text = oNazwa1Kli
                    lblNazwa3.Text = ""

                    lblKod.Text = oKodPoczKli
                    lblMiejscowosc.Text = oMiejscKli
                    lblUlica.Text = oUlicaKli
                    lblNrDomu.Text = oNrDomuKli
                    lblIDDomu.Text = oIDDomuKli
                    lblAdres.Text = ""
                    lblNIP.Text = oNIPKli

                    lblTelefon.Text = oTelefonKli
                    lblFax.Text = oFaxKli
                    lbleMail.Text = oeMailKli
                    lblUwagi.Text = ""


                    'Public oUniqIdKon As String           'UNIQID            varchar(40)
                    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
                    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
                    'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
                    'Public oTematKon As String            'TEMAT            Text, 50
                    'Public oUwagiKon As String            'UWAGI            Memo
                    'Public oWersjaKon As Integer          'WERSJA           Integer

                    lblData.Text = GetTxtField(dagKontakty, 2)
                    lblTemat.Text = GetTxtField(dagKontakty, 4)
                    lblNotatka.Text = GetTxtField(dagKontakty, 5)
                End If
            End If
        End If
    End Sub




    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
        HorizScrollTimer.Enabled = False
        If Len(HorizScrollLastTime) = 0 Then
            'nie zainicjowano scrollingu
        Else
            'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
            If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
                If dagKontakty.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagKontakty.FirstDisplayedScrollingColumnIndex +
                                    dagKontakty.GetCellDisplayRectangle(dagKontakty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagKontakty.FirstDisplayedScrollingColumnIndex +
                                    dagKontakty.GetCellDisplayRectangle(dagKontakty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagKontakty.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagKontakty.Rows(nextrow).Selected = True
            End If
        End If
    End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXX_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormy Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagKontakty, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, DataViewKontakty, stbKontakty)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagKontakty, Mid(sender.name, 6, 3), "Kontakty")
    End Sub
    Private Sub cmdFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me, sender)
    End Sub
    Private Sub cmdFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystko.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormy = True
            CzyscFiltryDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, DataViewKontakty, stbKontakty)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbKontakty.Panels(0).Text = "Il.zap.: " & DataViewKontakty.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagKontakty_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagKontakty.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagKontakty.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagKontakty.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltry_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltry.PanelClick
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagKontakty_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagKontakty.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagKontakty, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagKontakty, e.RowIndex, 4)

    '        Select Case UserName = Program_IdUzytkownika
    '            Case True
    '                e.CellStyle.ForeColor = Color.Red
    '            Case Else
    '                Select Case Mid(Upr, _Uprawnnienia_NrZnaku, 1)
    '                    Case "A"
    '                        e.CellStyle.ForeColor = Color.Green
    '                    Case "S"
    '                        e.CellStyle.ForeColor = Color.Purple
    '                    Case "U"
    '                        e.CellStyle.ForeColor = Color.Navy
    '                    Case Else
    '                        e.CellStyle.ForeColor = Color.Black
    '                End Select
    '        End Select
    '        '-----------------------
    '        ''zmieniamy BackColor
    '        'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
    '        'Select Case Wal
    '        '    Case "EUR"
    '        '        e.CellStyle.BackColor = Color.White
    '        '    Case Else
    '        'End Select
    '        '-----------------------
    '    End If
    'End Sub




    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagKontakty_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagKontakty.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKontakty_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKontakty.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKontakty_Scroll(sender As Object, e As ScrollEventArgs) Handles dagKontakty.Scroll
        'If (Not StartFormy) And (DataViewKontakty.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewKontakty.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagKontakty_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKontakty.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKontakty_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKontakty.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormy Then
                '    PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagKontakty, DataViewKontakty.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagKontakty_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKontakty.MouseUp
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Dim message As String = "You clicked "

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None
                message &= "the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                message &= "cell at row " & hti.RowIndex & ", col " & hti.ColumnIndex
                stbKontakty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbKontakty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagKontakty(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbKontakty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbKontakty.Panels(3).Text = "Sort: " & DataSetKontakty.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbKontakty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagKontakty(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbKontakty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbKontakty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub








    Private Sub dagKontakty_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKontakty.DoubleClick
        If _CoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagKontakty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagKontakty.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _CoMamRobic = "WYBIERAJ" Then
                    cmdStop_Click(sender, e)
                Else
                    cmdEdytuj_Click(sender, e)
                End If
            Case Keys.Insert
                cmdDodaj_Click(sender, e)
            Case Keys.Delete
                cmdUsuñ_Click(sender, e)
            Case Else
        End Select
    End Sub


    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************
    'Public oUniqIdKon As String           'UNIQID            varchar(40)
    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
    'Public oTematKon As String            'TEMAT            Text, 50
    'Public oUwagiKon As String            'UWAGI            Memo
    'Public oWersjaKon As Integer          'WERSJA           Integer


    Private Sub InitKontakty()
        oUniqIdKon = Guid.NewGuid.ToString
        oIdentKon = _IdentKlienta
        oDataKon = SysData()
        oPracownikKon = ""
        oTematKon = ""
        oUwagiKon = ""
        oWersjaKon = 1          'WERSJA         Integer, 2
    End Sub

    Private Sub PobierzKontakty()
        'pobierz wartosci przed aktualizacja
        oUniqIdKon = GetTxtField(dagKontakty, 0)
        oIdentKon = GetTxtField(dagKontakty, 1)
        oDataKon = GetTxtField(dagKontakty, 2)
        oPracownikKon = GetTxtField(dagKontakty, 3)
        oTematKon = GetTxtField(dagKontakty, 4)
        oUwagiKon = GetTxtField(dagKontakty, 5)
        oWersjaKon = GetNumField(dagKontakty, 6)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetKontakty.Tables(0).Rows.Count <= 0 Then
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzKontakty()
            Dim Form As New EdytujKatalogKontaktow
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oUniqIdKon
                foundRow = DataSetKontakty.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("UNIQID") = oUniqIdKon
                    foundRow("IDENTKLIENTA") = oIdentKli
                    foundRow("DATAKONTAKTU") = oDataKon
                    foundRow("PRACOWNIK") = oPracownikKon
                    foundRow("TEMAT") = oTematKon
                    foundRow("UWAGI") = oUwagiKon
                    foundRow("WERSJA") = ZmienWersje(oWersjaKon) 'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbKontakty.Panels(0).Text = "Il.rec.: " & DataViewKontakty.Count.ToString
                    AktualizujBaze()
                    DopiszDoLoguAktywnosci("",
                                           "Katalog Kontaktów Handlowych",
                                           SysTime(),
                                           Program_IdUzytkownika,
                                           LOG_OperacjaEdytuj,
                                           "Edytowano zapis kontaktu pracownika " & Trim(oPracownikKon) & " z klientem " & oIdentKon & vbCrLf &
                                           "Temat : " & oTematKon & vbCrLf &
                                           "Data kontaktu : " & oDataKon & "     Funkcja : KatalogKontaktow" & vbCrLf &
                                           "")


                    'aktualizuj DataGrid
                    dagKontakty.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetKontakty.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis ?",
                        "Proszê o potwierdzenie :",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
                oOperacja = "USUN"
                PobierzKontakty()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oUniqIdKon
                foundRow = DataSetKontakty.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbKontakty.Panels(0).Text = "Il.rec.: " & DataViewKontakty.Count.ToString
                    AktualizujBaze()
                    DopiszDoLoguAktywnosci("",
                                           "Katalog Kontaktów Handlowych",
                                           SysTime(),
                                           Program_IdUzytkownika,
                                           LOG_OperacjaUsun,
                                           "Usunieto zapis kontaktu pracownika " & Trim(oPracownikKon) & " z klientem " & oIdentKon & vbCrLf &
                                           "Temat : " & oTematKon & vbCrLf &
                                           "Data kontaktu : " & oDataKon & "     Funkcja : KatalogKontaktow" & vbCrLf &
                                           oUwagiKon)
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitKontakty()
        '----------------------------------------------
        Dim Form As New EdytujKatalogKontaktow
        Do While True
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oUniqIdKon
                foundRow = DataSetKontakty.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "UNIQID = " & oUniqIdKon,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetKontakty.Tables(0).NewRow()

                    NewRow("UNIQID") = oUniqIdKon
                    NewRow("IDENTKLIENTA") = oIdentKon
                    NewRow("DATAKONTAKTU") = oDataKon
                    NewRow("PRACOWNIK") = oPracownikKon
                    NewRow("TEMAT") = oTematKon
                    NewRow("UWAGI") = oUwagiKon
                    NewRow("WERSJA") = 1 'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetKontakty.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbKontakty.Panels(0).Text = "Il.rec.: " & DataViewKontakty.Count.ToString
                    AktualizujBaze()
                    DopiszDoLoguAktywnosci("",
                                           "Katalog Kontaktów Handlowych",
                                           SysTime(),
                                           Program_IdUzytkownika,
                                           LOG_OperacjaDodaj,
                                           "Dodano zapis kontaktu pracownika " & Trim(oPracownikKon) & " z klientem " & oIdentKon & vbCrLf &
                                           "Temat : " & oTematKon & vbCrLf &
                                           "Data kontaktu : " & oDataKon & "     Funkcja : KatalogKontaktow" & vbCrLf &
                                           "")




                    'aktualizuj DataGrid
                    dagKontakty.Update()
                    Exit Do
                End If
            End If
        Loop
    End Sub


    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnection.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapter.Update(DataSetKontakty, "TABELA_Kontakty")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetKontakty)
            '    End If
            '    dbConnection.Close()
            'End If
        Else
            Try
                sqlConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnection.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapter.Update(DataSetKontakty, "TABELA_Kontakty")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetKontakty)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub

    Private Sub OdswiezBaze()
        DataSetKontakty.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetKontakty)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetKontakty)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub




    '****************************************
    '** obsluga Menu Kontekstowego
    '****************************************

    Private Sub menuWEdytuj_Click(sender As Object, e As EventArgs) Handles menuWEdytuj.Click
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWDodaj_Click(sender As Object, e As EventArgs) Handles menuWDodaj.Click
        cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWUsun_Click(sender As Object, e As EventArgs) Handles menuWUsun.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWWybierz_Click(sender As Object, e As EventArgs) Handles menuWWybierz.Click
        cmdStop_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWSingleL_Click(sender As Object, e As EventArgs) Handles menuWSingleL.Click
        dagKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagKontakty.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        dagKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagKontakty.Refresh()
    End Sub

    Private Sub menuWOdswiez_Click(sender As Object, e As EventArgs) Handles menuWOdswiez.Click
        OdswiezBaze()
    End Sub




    Private Sub menuEEdytuj_Click(sender As Object, e As EventArgs) Handles menuEEdytuj.Click
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuEDodaj_Click(sender As Object, e As EventArgs) Handles menuEDodaj.Click
        cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuEUsun_Click(sender As Object, e As EventArgs) Handles menuEUsun.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuESingleL_Click(sender As Object, e As EventArgs) Handles menuESingleL.Click
        dagKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagKontakty.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagKontakty.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub







    Private Sub cmdAnaliza_Click(sender As Object, e As EventArgs) Handles cmdAnaliza.Click
        Dim Form As New frmAnalizaCzestKontaktow(DataViewKontakty, lblIdent.Text)
        Form.ShowDialog()
    End Sub


End Class

