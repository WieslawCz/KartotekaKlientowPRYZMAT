Public Class KatalogObrotowMies
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Dim _EdytujObroty As Boolean = False

    Public Sub New(ByVal ParEdytuj As Boolean)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _EdytujObroty = ParEdytuj
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
    Friend WithEvents lblNazwa2 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa1 As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHandlowa As System.Windows.Forms.Label
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents stbObrotyMies As System.Windows.Forms.StatusBar
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents dagObrotyMies As DataGridView
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystko As Button
    Friend WithEvents stbFiltry As StatusBar
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
    Friend WithEvents lblIDDomu As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogObrotowMies))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblIDDomu = New System.Windows.Forms.Label()
        Me.lblNrDomu = New System.Windows.Forms.Label()
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
        Me.lblNazwa2 = New System.Windows.Forms.Label()
        Me.lblNazwa1 = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.stbObrotyMies = New System.Windows.Forms.StatusBar()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.dagObrotyMies = New System.Windows.Forms.DataGridView()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
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
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagObrotyMies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuWybieraj.SuspendLayout()
        Me.menuEdytuj.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblIDDomu)
        Me.Panel1.Controls.Add(Me.lblNrDomu)
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
        Me.Panel1.Controls.Add(Me.lblNazwa2)
        Me.Panel1.Controls.Add(Me.lblNazwa1)
        Me.Panel1.Controls.Add(Me.lblNazwaHandlowa)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(810, 96)
        Me.Panel1.TabIndex = 48
        '
        'lblIDDomu
        '
        Me.lblIDDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIDDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIDDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblIDDomu.Location = New System.Drawing.Point(537, 23)
        Me.lblIDDomu.Name = "lblIDDomu"
        Me.lblIDDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblIDDomu.TabIndex = 29
        '
        'lblNrDomu
        '
        Me.lblNrDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNrDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNrDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblNrDomu.Location = New System.Drawing.Point(494, 23)
        Me.lblNrDomu.Name = "lblNrDomu"
        Me.lblNrDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblNrDomu.TabIndex = 11
        '
        'lblUlica
        '
        Me.lblUlica.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUlica.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUlica.ForeColor = System.Drawing.Color.Purple
        Me.lblUlica.Location = New System.Drawing.Point(352, 23)
        Me.lblUlica.Name = "lblUlica"
        Me.lblUlica.Size = New System.Drawing.Size(146, 16)
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
        Me.lblNIP.Location = New System.Drawing.Point(352, 56)
        Me.lblNIP.Name = "lblNIP"
        Me.lblNIP.Size = New System.Drawing.Size(230, 16)
        Me.lblNIP.TabIndex = 10
        '
        'lblAdres
        '
        Me.lblAdres.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdres.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAdres.ForeColor = System.Drawing.Color.Purple
        Me.lblAdres.Location = New System.Drawing.Point(352, 40)
        Me.lblAdres.Name = "lblAdres"
        Me.lblAdres.Size = New System.Drawing.Size(230, 16)
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
        'lblNazwa2
        '
        Me.lblNazwa2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2.Location = New System.Drawing.Point(64, 72)
        Me.lblNazwa2.Name = "lblNazwa2"
        Me.lblNazwa2.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwa2.TabIndex = 4
        '
        'lblNazwa1
        '
        Me.lblNazwa1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1.Location = New System.Drawing.Point(64, 56)
        Me.lblNazwa1.Name = "lblNazwa1"
        Me.lblNazwa1.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwa1.TabIndex = 3
        '
        'lblNazwaHandlowa
        '
        Me.lblNazwaHandlowa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa.Location = New System.Drawing.Point(64, 40)
        Me.lblNazwaHandlowa.Name = "lblNazwaHandlowa"
        Me.lblNazwaHandlowa.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwaHandlowa.TabIndex = 2
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
        Me.Label2.Location = New System.Drawing.Point(312, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "NIP"
        '
        'stbObrotyMies
        '
        Me.stbObrotyMies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbObrotyMies.Dock = System.Windows.Forms.DockStyle.None
        Me.stbObrotyMies.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbObrotyMies.Location = New System.Drawing.Point(8, 370)
        Me.stbObrotyMies.Name = "stbObrotyMies"
        Me.stbObrotyMies.ShowPanels = True
        Me.stbObrotyMies.Size = New System.Drawing.Size(807, 22)
        Me.stbObrotyMies.TabIndex = 66
        Me.stbObrotyMies.Text = "StatusBar1"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(16, 402)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 63
        Me.PictureBox1.TabStop = False
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(722, 402)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 23)
        Me.cmdStop.TabIndex = 59
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
        Me.cmdEdytuj.Location = New System.Drawing.Point(634, 402)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytuj.TabIndex = 62
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
        Me.cmdUsuñ.Location = New System.Drawing.Point(546, 402)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 23)
        Me.cmdUsuñ.TabIndex = 61
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
        Me.cmdDodaj.Location = New System.Drawing.Point(458, 402)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodaj.TabIndex = 60
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dagObrotyMies
        '
        Me.dagObrotyMies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagObrotyMies.Location = New System.Drawing.Point(8, 132)
        Me.dagObrotyMies.Name = "dagObrotyMies"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagObrotyMies.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagObrotyMies.Size = New System.Drawing.Size(806, 236)
        Me.dagObrotyMies.TabIndex = 181
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(546, 112)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 186
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(603, 110)
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
        Me.stbFiltry.Location = New System.Drawing.Point(8, 110)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(806, 22)
        Me.stbFiltry.TabIndex = 185
        Me.stbFiltry.Text = "StatusBar1"
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
        'KatalogObrotowMies
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(826, 432)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.dagObrotyMies)
        Me.Controls.Add(Me.stbObrotyMies)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "KatalogObrotowMies"
        Me.ShowInTaskbar = False
        Me.Text = "Katalog Obrotów w ostatnim miesi¹cu"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagObrotyMies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuWybieraj.ResumeLayout(False)
        Me.menuEdytuj.ResumeLayout(False)
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




    '---------------------------------------------------------------------
    ''ObrotyMies
    'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
    'Public oMiesiacMies As String          'MIESIAC          Text,7
    'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
    'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
    'Public oLMARSprzedMies As Double       'LMARSPRZED       Double
    'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
    'Public oAIloSprzedMies As Double       'AILOSPRZED       Double
    'Public oAMARSprzedMies As Double       'AMARSPRZED       Double
    'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
    'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
    'Public oLOMARSprzedMies As Double       'LOMARSPRZED       Double
    'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
    'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double
    'Public oAOMARSprzedMies As Double       'AOMARSPRZED       Double
    'Public oWersjaMies As Integer          'WERSJA           Integer

    Dim baseSelectSQLObrotyMies As String = "SELECT " &
                                                 "ObrotyMies.IDENTKLIENTA," &
                                                 "Klienci.NRTOFITXT," &
                                                 "ObrotyMies.MIESIAC," &
                                                 "ObrotyMies.AILOSPRZED," &
                                                 "ObrotyMies.AWARTSPRZED," &
                                                 "ObrotyMies.AMARSPRZED," &
                                                 "ObrotyMies.LILOSPRZED," &
                                                 "ObrotyMies.LWARTSPRZED," &
                                                 "ObrotyMies.LMARSPRZED," &
                                                 "ObrotyMies.AOILOSPRZED," &
                                                 "ObrotyMies.AOWARTSPRZED," &
                                                 "ObrotyMies.AOMARSPRZED," &
                                                 "ObrotyMies.LOILOSPRZED," &
                                                 "ObrotyMies.LOWARTSPRZED," &
                                                 "ObrotyMies.LOMARSPRZED," &
                                                   "ObrotyMies.NAJEMILOSPRZED," &
                                                   "ObrotyMies.NAJEMWARTSPRZED," &
                                                   "ObrotyMies.NAJEMMARSPRZED," &
                                                 "ObrotyMies.STRONYILOSPRZED," &
                                                 "ObrotyMies.STRONYWARTSPRZED," &
                                                 "ObrotyMies.STRONYMARSPRZED," &
                                                   "ObrotyMies.DRUKARKIILOSPRZED," &
                                                   "ObrotyMies.DRUKARKIWARTSPRZED," &
                                                   "ObrotyMies.DRUKARKIMARSPRZED," &
                                                 "ObrotyMies.SKUPILOSPRZED," &
                                                 "ObrotyMies.SKUPWARTSPRZED," &
                                                 "ObrotyMies.SKUPMARSPRZED," &
                                                   "ObrotyMies.AILOSPRZED + ObrotyMies.LILOSPRZED + " &
                                                   "ObrotyMies.AOILOSPRZED + ObrotyMies.LOILOSPRZED + " &
                                                   "ObrotyMies.NAJEMILOSPRZED + ObrotyMies.STRONYILOSPRZED + " &
                                                   "ObrotyMies.DRUKARKIILOSPRZED + ObrotyMies.SKUPILOSPRZED AS SUMAILO, " &
                                                     "ObrotyMies.AWARTSPRZED + ObrotyMies.LWARTSPRZED + " &
                                                     "ObrotyMies.AOWARTSPRZED + ObrotyMies.LOWARTSPRZED + " &
                                                     "ObrotyMies.NAJEMWARTSPRZED + ObrotyMies.STRONYWARTSPRZED + " &
                                                     "ObrotyMies.DRUKARKIWARTSPRZED + ObrotyMies.SKUPWARTSPRZED AS SUMAWART, " &
                                                       "ObrotyMies.AMARSPRZED + ObrotyMies.LMARSPRZED + " &
                                                       "ObrotyMies.AOMARSPRZED + ObrotyMies.LOMARSPRZED + " &
                                                       "ObrotyMies.NAJEMMARSPRZED + ObrotyMies.STRONYMARSPRZED + " &
                                                       "ObrotyMies.DRUKARKIMARSPRZED + ObrotyMies.SKUPMARSPRZED AS SUMAMAR, " &
                                                   "ObrotyMies.WERSJA " &
                                            "FROM ObrotyMies INNER JOIN Klienci ON ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA "

    Dim dbSelectSQLObrotyMies As String = baseSelectSQLObrotyMies & " WHERE ObrotyMies.IDENTKLIENTA='" & oIdentKli & "'"

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
    Dim DataSetObrotyMies As New DataSet
    Dim DataViewObrotyMies As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Private Sub KatalogObrotowMies_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If _EdytujObroty Then
        Else
            NieEdytowac(Me)
        End If
    End Sub

    Private Sub KatalogOsob_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        'DataSet
        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")

        If Len(oIdentKli) > 0 Then
            dbSelectSQLObrotyMies = baseSelectSQLObrotyMies & " WHERE ObrotyMies.IDENTKLIENTA='" & oIdentKli & "'"
        Else
            dbSelectSQLObrotyMies = baseSelectSQLObrotyMies
            cmdDodaj.Visible = False
        End If

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnection = New OleDb.OleDbConnection(sConnectionObroty)
            dbCommandSelect = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnection)
            dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLObrotyMies, dbConnection)
            dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLObrotyMies, dbConnection)
            dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLObrotyMies, dbConnection)
            dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_ObrotyMies")

            DBDeleteObrotyMies(dbCommandDelete, dbDataAdapter)
            DBUpdateObrotyMies(dbCommandUpdate, dbDataAdapter)
            DBInsertObrotyMies(dbCommandInsert, dbDataAdapter)

            ' Add the RowUpdated event handler.
            AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnection.State
            End Try
        Else
            sqlConnection = New SqlClient.SqlConnection(sConnectionObrotyMies)
            sqlCommandSelect = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLObrotyMies, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLObrotyMies, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_ObrotyMies")

            SQLDeleteObrotyMies(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateObrotyMies(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertObrotyMies(sqlCommandInsert, sqlDataAdapter)

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
                    dbDataAdapter.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    dbDataAdapter.Fill(DataSetObrotyMies)
                    dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapter.Fill(DataSetObrotyMies)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"),
                                                                           DataSetObrotyMies.Tables(0).Columns("MIESIAC")}

                'definiuj DataView
                DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))
                DataViewObrotyMies.AllowDelete = False
                DataViewObrotyMies.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagObrotyMies.BackgroundColor = System.Drawing.SystemColors.Control
                dagObrotyMies.GridColor = System.Drawing.SystemColors.ControlDark
                dagObrotyMies.ForeColor = System.Drawing.Color.Purple
                dagObrotyMies.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagObrotyMies.BorderStyle = BorderStyle.Fixed3D
                'dagObrotyMies.Dock = DockStyle.Fill
                dagObrotyMies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagObrotyMies.AllowUserToAddRows = False
                dagObrotyMies.AllowUserToDeleteRows = False
                dagObrotyMies.AllowUserToOrderColumns = True
                dagObrotyMies.AllowUserToResizeColumns = True
                dagObrotyMies.AllowUserToResizeRows = True
                dagObrotyMies.ReadOnly = True
                dagObrotyMies.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagObrotyMies.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagObrotyMies.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagObrotyMies.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagObrotyMies.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagObrotyMies.ColumnHeadersVisible = True
                dagObrotyMies.ColumnHeadersHeight = 60
                dagObrotyMies.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagObrotyMies.RowHeadersVisible = True
                dagObrotyMies.RowHeadersWidth = 50
                dagObrotyMies.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagObrotyMies.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagObrotyMies.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagObrotyMies.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagObrotyMies.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagObrotyMies.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagObrotyMies.DefaultCellStyle.NullValue = ""
                dagObrotyMies.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagObrotyMies.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagObrotyMies.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagObrotyMies.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagObrotyMies.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagObrotyMies.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagObrotyMies.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagObrotyMies.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagObrotyMies.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagObrotyMies.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagObrotyMies.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagObrotyMies.DataMember = DataSetObrotyMies.Tables(0).TableName
                'dagObrotyMies.DataSource = DataViewObrotyMies
                '-----------------------------------

                'Dim baseSelectSQLObrotyMies As String = "SELECT " &
                '                                 "ObrotyMies.IDENTKLIENTA," &
                '                                 "Klienci.NRTOFITXT," &
                '                                 "ObrotyMies.MIESIAC," &
                '                                 "ObrotyMies.AILOSPRZED," &
                '                                 "ObrotyMies.AWARTSPRZED," &
                '                                 "ObrotyMies.AMARSPRZED," &
                '                                 "ObrotyMies.LILOSPRZED," &
                '                                 "ObrotyMies.LWARTSPRZED," &
                '                                 "ObrotyMies.LMARSPRZED," &
                '                                 "ObrotyMies.AOILOSPRZED," &
                '                                 "ObrotyMies.AOWARTSPRZED," &
                '                                 "ObrotyMies.AOMARSPRZED," &
                '                                 "ObrotyMies.LOILOSPRZED," &
                '                                 "ObrotyMies.LOWARTSPRZED," &
                '                                 "ObrotyMies.LOMARSPRZED," &
                '                                   "Obroty.NAJEMILOSPRZED," &
                '                                   "Obroty.NAJEMWARTSPRZED," &
                '                                   "Obroty.NAJEMMARSPRZED," &
                '                                 "Obroty.STRONYILOSPRZED," &
                '                                 "Obroty.STRONYWARTSPRZED," &
                '                                 "Obroty.STRONYMARSPRZED," &
                '                                   "Obroty.DRUKARKIILOSPRZED," &
                '                                   "Obroty.DRUKARKIWARTSPRZED," &
                '                                   "Obroty.DRUKARKIMARSPRZED," &
                '                                 "Obroty.SKUPILOSPRZED," &
                '                                 "Obroty.SKUPWARTSPRZED," &
                '                                 "Obroty.SKUPMARSPRZED," &
                '                                   "Obroty.AILOSPRZED + Obroty.LILOSPRZED + " &
                '                                   "Obroty.AOILOSPRZED + Obroty.LOILOSPRZED + " &
                '                                   "Obroty.NAJEMILOSPRZED + Obroty.STRONYILOSPRZED + " &
                '                                   "Obroty.DRUKARKIILOSPRZED + Obroty.SKUPILOSPRZED AS SUMAILO, " &
                '                                     "Obroty.AWARTSPRZED + Obroty.LWARTSPRZED + " &
                '                                     "Obroty.AOWARTSPRZED + Obroty.LOWARTSPRZED + " &
                '                                     "Obroty.NAJEMWARTSPRZED + Obroty.STRONYWARTSPRZED + " &
                '                                     "Obroty.DRUKARKIWARTSPRZED + Obroty.SKUPWARTSPRZED AS SUMAWART, " &
                '                                       "Obroty.AMARSPRZED + Obroty.LMARSPRZED + " &
                '                                       "Obroty.AOMARSPRZED + Obroty.LOMARSPRZED + " &
                '                                       "Obroty.NAJEMMARSPRZED + Obroty.STRONYMARSPRZED + " &
                '                                       "Obroty.DRUKARKIMARSPRZED + Obroty.SKUPMARSPRZED AS SUMAMAR " &
                '                                   "ObrotyMies.WERSJA " &
                '                            "FROM ObrotyMies INNER JOIN Klienci ON ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA "


                TxtColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(0).ColumnName, "Klient", 50, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(1).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(2).ColumnName, "Miesiac", 80, DataGridViewContentAlignment.MiddleLeft, True)

                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(8).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(7).ColumnName, "Sprzeda¿ Laser Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(6).ColumnName, "Sprzeda¿ Laser Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)

                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(5).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(4).ColumnName, "Sprzeda¿ Atrament Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(3).ColumnName, "Sprzeda¿ Atrament Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)

                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(14).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(13).ColumnName, "Sprzeda¿ Oryg Laser Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(12).ColumnName, "Sprzeda¿ Oryg Laser Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)

                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(11).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(10).ColumnName, "Sprzeda¿ Oryg Atrament Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(9).ColumnName, "Sprzeda¿ Oryg Atrament Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)


                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(17).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(16).ColumnName, "Najem Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(15).ColumnName, "Najem Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)

                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(20).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(19).ColumnName, "Strony Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(18).ColumnName, "Strony Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)

                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(23).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(22).ColumnName, "Drukarki Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(21).ColumnName, "Drukarki Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)

                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(26).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(25).ColumnName, "Skup Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(24).ColumnName, "Skup Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)

                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(29).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(28).ColumnName, "Suma Wartoœci", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(27).ColumnName, "Suma Iloœci", 100, DataGridViewContentAlignment.MiddleRight, "N0", True)

                TxtColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(30).ColumnName, "", 0, HorizontalAlignment.Left, False)


                Me.Cursor = Cursors.WaitCursor
                dagObrotyMies.DataSource = DataViewObrotyMies
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetObrotyMies.Tables(0).Rows.Count > 0 Then
                    dagObrotyMies.CurrentCell = dagObrotyMies(0, 0)
                    dagObrotyMies.CurrentCell.Selected = True
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
            stbObrotyMies.BackColor = System.Drawing.SystemColors.Control
            stbObrotyMies.ForeColor = System.Drawing.Color.DarkGreen
            stbObrotyMies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbObrotyMies.Panels.Add("stbPanelIleRec")
            stbObrotyMies.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbObrotyMies.Panels(0).Width = 80
            stbObrotyMies.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbObrotyMies.Panels(0).Text = "Iloœæ rec.: " & DataSetObrotyMies.Tables(0).Rows.Count.ToString

            stbObrotyMies.Panels.Add("stbPanelWiersz")
            stbObrotyMies.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbObrotyMies.Panels(1).Width = 80
            stbObrotyMies.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagObrotyMies.CurrentCell Is Nothing Then
                stbObrotyMies.Panels(1).Text = "Wi: "
                'stbObrotyMies.Panels(2).Text = "Ko: "
            Else
                stbObrotyMies.Panels(1).Text = "Wi: " & dagObrotyMies.CurrentCell.RowIndex.ToString
                'stbObrotyMies.Panels(2).Text = "Ko: " & dagObrotyMies.CurrentCell.ColumnIndex.ToString
            End If

            stbObrotyMies.Panels.Add("stbPanelKolumna")
            stbObrotyMies.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbObrotyMies.Panels(2).Width = 80
            stbObrotyMies.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagObrotyMies.CurrentCell Is Nothing Then
                'stbObrotyMies.Panels(1).Text = "Wi: "
                stbObrotyMies.Panels(2).Text = "Ko: "
            Else
                'stbObrotyMies.Panels(1).Text = "Wi: " & dagObrotyMies.CurrentCell.RowIndex.ToString
                stbObrotyMies.Panels(2).Text = "Ko: " & dagObrotyMies.CurrentCell.ColumnIndex.ToString
            End If

            stbObrotyMies.Panels.Add("stbPanelSort")
            stbObrotyMies.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbObrotyMies.Panels(3).Width = 150
            stbObrotyMies.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbObrotyMies.Panels(3).Text = "Sort : "

            stbObrotyMies.Panels.Add("stbPanelFiltr")
            stbObrotyMies.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbObrotyMies.Panels(4).Width = 150
            stbObrotyMies.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbObrotyMies.Panels(4).Text = "Filtr : "

            stbObrotyMies.Panels.Add("stbPanelReszta")
            stbObrotyMies.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbObrotyMies.Panels(5).Width = 150
            stbObrotyMies.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbObrotyMies.Panels(5).Text = ""

            stbObrotyMies.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagObrotyMies.TableStyles(0).RowHeaderWidth
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
        InitObrotyMies() 'inicjuj zmienne
        '-----------------------------------------------------------------
        If _OCoMamRobic = "WYBIERAJ" Then
            dagObrotyMies.ContextMenuStrip = Me.menuWybieraj
            cmdStop.Text = "Wybierz"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = True
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = False
        Else
            dagObrotyMies.ContextMenuStrip = Me.menuEdytuj
            cmdStop.Text = "Powrót"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = False
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = True
        End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagObrotyMies.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagObrotyMies.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetObrotyMies.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagObrotyMies.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagObrotyMies.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagObrotyMies.FirstDisplayedScrollingColumnIndex +
                        dagObrotyMies.GetCellDisplayRectangle(dagObrotyMies.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagObrotyMies.Left + 4), (dagObrotyMies.Top + 4))
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
        For ii = 0 To DataViewObrotyMies.Table.Columns.Count - 1
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
        dagObrotyMies_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub


    Private Sub KatalogObrotowMies_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub













    Private Sub dagObrotyMies_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObrotyMies.CurrentCellChanged
        If Not StartFormy Then
            If dagObrotyMies.CurrentCell Is Nothing Then
                stbObrotyMies.Panels(1).Text = "Wi: "
                stbObrotyMies.Panels(2).Text = "Ko: "

                lblIdent.Text = ""
                lblTOFI.Text = ""
                lblNazwaHandlowa.Text = ""
                lblNazwa1.Text = ""
                lblNazwa2.Text = ""

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
            Else
                stbObrotyMies.Panels(1).Text = "Wi: " & dagObrotyMies.CurrentCell.RowIndex.ToString
                stbObrotyMies.Panels(2).Text = "Ko: " & dagObrotyMies.CurrentCell.ColumnIndex.ToString

                If DataViewObrotyMies.Count = 0 Then
                    lblIdent.Text = ""
                    lblTOFI.Text = ""
                    lblNazwaHandlowa.Text = ""
                    lblNazwa1.Text = ""
                    lblNazwa2.Text = ""

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
                Else
                    'Public oIdentOso As String            'IDENTKLIENTA     Text, 6
                    'Public oOsobaOso As String            'OSOBA            Text, 50
                    'Public oStanowiskoOso As String       'STANOWISKO       Text, 50
                    'Public oKompetencjeOso As String      'KOMPETENCJE      Text, 50
                    'Public oTelefonOso As String          'TELEFON          Text, 50
                    'Public oeMailOso As String            'EMAIL            Text, 50
                    'Public oVIPOso As Boolean              'VIP              boolean
                    'Public oWersjaOso As Integer          'WERSJA           Integer

                    IdentKlienta(GetTxtField(dagObrotyMies, 0))

                    lblIdent.Text = GetTxtField(dagObrotyMies, 0)
                    lblTOFI.Text = oNrTOFITxtKli
                    lblNazwaHandlowa.Text = ""
                    lblNazwa1.Text = oNazwa1Kli
                    lblNazwa2.Text = ""

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
                If dagObrotyMies.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagObrotyMies.FirstDisplayedScrollingColumnIndex +
                                    dagObrotyMies.GetCellDisplayRectangle(dagObrotyMies.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagObrotyMies.FirstDisplayedScrollingColumnIndex +
                                    dagObrotyMies.GetCellDisplayRectangle(dagObrotyMies.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagObrotyMies.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagObrotyMies.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagObrotyMies, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, DataViewObrotyMies, stbObrotyMies)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagObrotyMies, Mid(sender.name, 6, 3), "ObrotyMies")
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
            CzyscFiltryDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, DataViewObrotyMies, stbObrotyMies)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbObrotyMies.Panels(0).Text = "Il.zap.: " & DataViewObrotyMies.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagObrotyMies_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagObrotyMies.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagObrotyMies.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagObrotyMies.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagObrotyMies_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagObrotyMies.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagObrotyMies, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagObrotyMies, e.RowIndex, 4)

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

    Private Sub dagObrotyMies_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagObrotyMies.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagObrotyMies_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObrotyMies.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagObrotyMies_Scroll(sender As Object, e As ScrollEventArgs) Handles dagObrotyMies.Scroll
        'If (Not StartFormy) And (DataViewObrotyMies.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewObrotyMies.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagObrotyMies_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObrotyMies.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagObrotyMies_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagObrotyMies.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagObrotyMies_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagObrotyMies.MouseUp
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
                stbObrotyMies.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbObrotyMies.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagObrotyMies(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbObrotyMies.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbObrotyMies.Panels(3).Text = "Sort: " & DataSetObrotyMies.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbObrotyMies.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagObrotyMies(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbObrotyMies.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbObrotyMies.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub




    Private Sub dagObrotyMies_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObrotyMies.DoubleClick
        If _OCoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            If _EdytujObroty Then
                cmdEdytuj_Click(sender, e)
            Else
            End If
        End If
    End Sub

    Private Sub dagObrotyMies_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagObrotyMies.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _OCoMamRobic = "WYBIERAJ" Then
                    cmdStop_Click(sender, e)
                Else
                    If _EdytujObroty Then
                        cmdEdytuj_Click(sender, e)
                    Else
                    End If
                End If
            Case Keys.Insert
                If _EdytujObroty Then
                    cmdDodaj_Click(sender, e)
                Else
                End If
            Case Keys.Delete
                If _EdytujObroty Then
                    cmdUsuñ_Click(sender, e)
                Else
                End If
            Case Else
        End Select
    End Sub

    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************

    Private Sub InitObrotyMies()
        oIdentMies = oIdentKli
        'IdentIdKlienta(oIdentMies)
        oMiesiacMies = Mid(SysData(), 1, 7)

        oAIloSprzedMies = 0
        oAWartSprzedMies = 0
        oAMARSprzedMies = 0
        oLIloSprzedMies = 0
        oLWartSprzedMies = 0
        oLMARSprzedMies = 0

        oAOIloSprzedMies = 0
        oAOWartSprzedMies = 0
        oAOMARSprzedMies = 0
        oLOIloSprzedMies = 0
        oLOWartSprzedMies = 0
        oLOMARSprzedMies = 0

        oNAJEMWartSprzedMies = 0
        oNAJEMIloSprzedMies = 0
        oNAJEMMARSprzedMies = 0

        oSTRONYWartSprzedMies = 0
        oSTRONYIloSprzedMies = 0
        oSTRONYMARSprzedMies = 0

        oDRUKARKIWartSprzedMies = 0
        oDRUKARKIIloSprzedMies = 0
        oDRUKARKIMARSprzedMies = 0

        oSKUPWartSprzedMies = 0
        oSKUPIloSprzedMies = 0
        oSKUPMARSprzedMies = 0

        oWersjaMies = 1          'WERSJA         Integer, 2
    End Sub

    'TxtColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(0).ColumnName, "Klient", 50, HorizontalAlignment.Left)
    'TxtColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(1).ColumnName, "Nr TOFI", 100, HorizontalAlignment.Left)
    'TxtColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(2).ColumnName, "Miesiac", 50, HorizontalAlignment.Left)
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(8).ColumnName, "", 0, HorizontalAlignment.Right, "N2")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(7).ColumnName, "Sprz.L Wart.", 100, HorizontalAlignment.Right, "N2")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(6).ColumnName, "Sprz.L Iloœæ", 100, HorizontalAlignment.Right, "N0")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(5).ColumnName, "", 0, HorizontalAlignment.Right, "N2")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(4).ColumnName, "Sprz.A Wart.", 100, HorizontalAlignment.Right, "N2")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(3).ColumnName, "Sprz.A Iloœæ", 100, HorizontalAlignment.Right, "N0")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(14).ColumnName, "", 0, HorizontalAlignment.Right, "N2")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(13).ColumnName, "Sprz.Oryg L Wart.", 100, HorizontalAlignment.Right, "N2")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(12).ColumnName, "Sprz.Oryg L Iloœæ", 100, HorizontalAlignment.Right, "N0")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(11).ColumnName, "", 0, HorizontalAlignment.Right, "N2")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(10).ColumnName, "Sprz.Oryg A Wart.", 100, HorizontalAlignment.Right, "N2")
    'NumColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(9).ColumnName, "Sprz.Oryg A Iloœæ", 100, HorizontalAlignment.Right, "N0")
    'TxtColumn(TSObroty, DataSetObrotyMies.Tables(0).Columns(15).ColumnName, "", 0, HorizontalAlignment.Left)

    Private Sub PobierzObroty()
        'pobierz wartosci przed aktualizacja
        'pobierz wartosci przed aktualizacja
        oIdentMies = GetTxtField(dagObrotyMies, 0)
        'IdentIdKlienta(oIdentMies)
        oMiesiacMies = GetTxtField(dagObrotyMies, 2)

        oLMARSprzedMies = GetDblField(dagObrotyMies, 3)
        oLWartSprzedMies = GetDblField(dagObrotyMies, 4)
        oLIloSprzedMies = GetDblField(dagObrotyMies, 5)
        oAMARSprzedMies = GetDblField(dagObrotyMies, 6)
        oAWartSprzedMies = GetDblField(dagObrotyMies, 7)
        oAIloSprzedMies = GetDblField(dagObrotyMies, 8)

        oLOMARSprzedMies = GetDblField(dagObrotyMies, 9)
        oLOWartSprzedMies = GetDblField(dagObrotyMies, 10)
        oLOIloSprzedMies = GetDblField(dagObrotyMies, 11)
        oAOMARSprzedMies = GetDblField(dagObrotyMies, 12)
        oAOWartSprzedMies = GetDblField(dagObrotyMies, 13)
        oAOIloSprzedMies = GetDblField(dagObrotyMies, 14)

        oNAJEMMARSprzedMies = GetDblField(dagObrotyMies, 15)
        oNAJEMWartSprzedMies = GetDblField(dagObrotyMies, 16)
        oNAJEMIloSprzedMies = GetDblField(dagObrotyMies, 17)

        oSTRONYMARSprzedMies = GetDblField(dagObrotyMies, 18)
        oSTRONYWartSprzedMies = GetDblField(dagObrotyMies, 19)
        oSTRONYIloSprzedMies = GetDblField(dagObrotyMies, 20)

        oDRUKARKIMARSprzedMies = GetDblField(dagObrotyMies, 21)
        oDRUKARKIWartSprzedMies = GetDblField(dagObrotyMies, 22)
        oDRUKARKIIloSprzedMies = GetDblField(dagObrotyMies, 23)

        oSKUPMARSprzedMies = GetDblField(dagObrotyMies, 24)
        oSKUPWartSprzedMies = GetDblField(dagObrotyMies, 25)
        oSKUPIloSprzedMies = GetDblField(dagObrotyMies, 26)

        oWersjaMies = GetIntField(dagObrotyMies, 30)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetObrotyMies.Tables(0).Rows.Count <= 0 Then
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzObroty()
            Dim Form As New EdytujKatalogObrotowMies
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentMies
                findTheseVals(1) = oMiesiacMies
                foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDENTKLIENTA") = oIdentmies
                    foundRow("NRTOFITXT") = oNrTOFITxtKli
                    'foundRow("MIESIAC") = oMiesiacMies

                    foundRow("LILOSPRZED") = oLIloSprzedMies
                    foundRow("LWARTSPRZED") = oLWartSprzedMies
                    foundRow("LMARSPRZED") = oLMARSprzedMies

                    foundRow("AILOSPRZED") = oAIloSprzedMies
                    foundRow("AWARTSPRZED") = oAWartSprzedMies
                    foundRow("AMARSPRZED") = oAMARSprzedMies

                    foundRow("LOILOSPRZED") = oLOIloSprzedMies
                    foundRow("LOWARTSPRZED") = oLOWartSprzedMies
                    foundRow("LOMARSPRZED") = oLOMARSprzedMies

                    foundRow("AOILOSPRZED") = oAOIloSprzedMies
                    foundRow("AOWARTSPRZED") = oAOWartSprzedMies
                    foundRow("AOMARSPRZED") = oAOMARSprzedMies

                    foundRow("NAJEMILOSPRZED") = oNAJEMIloSprzedMies
                    foundRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedMies
                    foundRow("NAJEMMARSPRZED") = oNAJEMMARSprzedMies

                    foundRow("STRONYILOSPRZED") = oSTRONYIloSprzedMies
                    foundRow("STRONYWARTSPRZED") = oSTRONYWartSprzedMies
                    foundRow("STRONYMARSPRZED") = oSTRONYMARSprzedMies

                    foundRow("DRUKARKIILOSPRZED") = oDRUKARKIIloSprzedMies
                    foundRow("DRUKARKIWARTSPRZED") = oDRUKARKIWartSprzedMies
                    foundRow("DRUKARKIMARSPRZED") = oDRUKARKIMARSprzedMies

                    foundRow("SKUPILOSPRZED") = oSKUPIloSprzedMies
                    foundRow("SKUPWARTSPRZED") = oSKUPWartSprzedMies
                    foundRow("SKUPMARSPRZED") = oSKUPMARSprzedMies

                    'foundRow("SUMAILO") = oLIloSprzedMies + oAIloSprzedMies +
                    '                      oLOIloSprzedMies + oAOIloSprzedMies +
                    '                      oNAJEMIloSprzedMies + oSTRONYIloSprzedMies +
                    '                      oDRUKARKIIloSprzedMies + oSKUPIloSprzedMies
                    'foundRow("SUMAWART") = oLWartSprzedMies + oAWartSprzedMies +
                    '                      oLOWartSprzedMies + oAOWartSprzedMies +
                    '                      oNAJEMWartSprzedMies + oSTRONYWartSprzedMies +
                    '                      oDRUKARKIWartSprzedMies + oSKUPWartSprzedMies
                    'foundRow("SUMAMAR") = oLMARSprzedMies + oAMARSprzedMies +
                    '                      oLOMARSprzedMies + oAOMARSprzedMies +
                    '                      oNAJEMMARSprzedMies + oSTRONYMARSprzedMies +
                    '                      oDRUKARKIMARSprzedMies + oSKUPMARSprzedMies

                    foundRow("WERSJA") = ZmienWersje(oWersjaMies) 'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbObrotyMies.Panels(0).Text = "Il.rec.: " & DataViewObrotyMies.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagObrotyMies.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetObrotyMies.Tables(0).Rows.Count <= 0 Then
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
                PobierzObroty()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentMies
                findTheseVals(1) = oMiesiacMies
                foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbObrotyMies.Panels(0).Text = "Il.rec.: " & DataViewObrotyMies.Count.ToString
                    AktualizujBaze()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitObrotyMies()
        '----------------------------------------------
        Do While True
            oAktualizuj = False
            Dim Form As New EdytujKatalogObrotowMies
            Form.ShowDialog()
            Form.Dispose()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentMies
                findTheseVals(1) = oMiesiacMies
                foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Klient = " & oIdentMies & vbCrLf &
                        "Miesi¹c = " & oMiesiacMies,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetObrotyMies.Tables(0).NewRow()

                    NewRow("IDENTKLIENTA") = oIdentMies
                    NewRow("NRTOFITXT") = oNrTOFITxtKli
                    NewRow("MIESIAC") = oMiesiacMies

                    NewRow("LILOSPRZED") = oLIloSprzedMies
                    NewRow("LWARTSPRZED") = oLWartSprzedMies
                    NewRow("LMARSPRZED") = oLMARSprzedMies

                    NewRow("AILOSPRZED") = oAIloSprzedMies
                    NewRow("AWARTSPRZED") = oAWartSprzedMies
                    NewRow("AMARSPRZED") = oAMARSprzedMies

                    NewRow("LOILOSPRZED") = oLOIloSprzedMies
                    NewRow("LOWARTSPRZED") = oLOWartSprzedMies
                    NewRow("LOMARSPRZED") = oLOMARSprzedMies

                    NewRow("AOILOSPRZED") = oAOIloSprzedMies
                    NewRow("AOWARTSPRZED") = oAOWartSprzedMies
                    NewRow("AOMARSPRZED") = oAOMARSprzedMies

                    NewRow("NAJEMILOSPRZED") = oNAJEMIloSprzedMies
                    NewRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedMies
                    NewRow("NAJEMMARSPRZED") = oNAJEMMARSprzedMies

                    NewRow("STRONYILOSPRZED") = oSTRONYIloSprzedMies
                    NewRow("STRONYWARTSPRZED") = oSTRONYWartSprzedMies
                    NewRow("STRONYMARSPRZED") = oSTRONYMARSprzedMies

                    NewRow("DRUKARKIILOSPRZED") = oDRUKARKIIloSprzedMies
                    NewRow("DRUKARKIWARTSPRZED") = oDRUKARKIWartSprzedMies
                    NewRow("DRUKARKIMARSPRZED") = oDRUKARKIMARSprzedMies

                    NewRow("SKUPILOSPRZED") = oSKUPIloSprzedMies
                    NewRow("SKUPWARTSPRZED") = oSKUPWartSprzedMies
                    NewRow("SKUPMARSPRZED") = oSKUPMARSprzedMies

                    NewRow("SUMAILO") = oLIloSprzedMies + oAIloSprzedMies +
                                          oLOIloSprzedMies + oAOIloSprzedMies +
                                          oNAJEMIloSprzedMies + oSTRONYIloSprzedMies +
                                          oDRUKARKIIloSprzedMies + oSKUPIloSprzedMies
                    NewRow("SUMAWART") = oLWartSprzedMies + oAWartSprzedMies +
                                          oLOWartSprzedMies + oAOWartSprzedMies +
                                          oNAJEMWartSprzedMies + oSTRONYWartSprzedMies +
                                          oDRUKARKIWartSprzedMies + oSKUPWartSprzedMies
                    NewRow("SUMAMAR") = oLMARSprzedMies + oAMARSprzedMies +
                                          oLOMARSprzedMies + oAOMARSprzedMies +
                                          oNAJEMMARSprzedMies + oSTRONYMARSprzedMies +
                                          oDRUKARKIMARSprzedMies + oSKUPMARSprzedMies

                    NewRow("WERSJA") = 1     'Wersja         Integer, 2

                    'dodaj rekord do DataSet
                    DataSetObrotyMies.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbObrotyMies.Panels(0).Text = "Il.rec.: " & DataViewObrotyMies.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagObrotyMies.Update()
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
            Try
                dbConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnection.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapter.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapter.Fill(DataSetObrotyMies)
                End If
                dbConnection.Close()
            End If
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
                    sqlDataAdapter.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetObrotyMies)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub


    Private Sub OdswiezBaze()
        DataSetObrotyMies.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnection.Open()
                If dbConnection.State = ConnectionState.Open Then
                    dbDataAdapter.Fill(DataSetObrotyMies)
                    dbConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetObrotyMies)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub





    '*****************************************************
    '** obs³uga menu kontekstowego
    '*****************************************************

    Private Sub menuEEdytuj_Click(sender As Object, e As EventArgs) Handles menuEEdytuj.Click
        If _EdytujObroty Then
            cmdEdytuj_Click(Me, System.EventArgs.Empty)
        Else
        End If
    End Sub

    Private Sub menuEDodaj_Click(sender As Object, e As EventArgs) Handles menuEDodaj.Click
        If _EdytujObroty Then
            cmdDodaj_Click(Me, System.EventArgs.Empty)
        Else
        End If
    End Sub

    Private Sub menuEUsun_Click(sender As Object, e As EventArgs) Handles menuEUsun.Click
        If _EdytujObroty Then
            cmdUsuñ_Click(Me, System.EventArgs.Empty)
        Else
        End If
    End Sub

    Private Sub menuESingleL_Click(sender As Object, e As EventArgs) Handles menuESingleL.Click
        dagObrotyMies.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagObrotyMies.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagObrotyMies.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagObrotyMies.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub




    Private Sub menuWEdytuj_Click(sender As Object, e As EventArgs) Handles menuWEdytuj.Click
        If _EdytujObroty Then
            cmdEdytuj_Click(Me, System.EventArgs.Empty)
        Else
        End If
    End Sub

    Private Sub menuWDodaj_Click(sender As Object, e As EventArgs) Handles menuWDodaj.Click
        If _EdytujObroty Then
            cmdDodaj_Click(Me, System.EventArgs.Empty)
        Else
        End If
    End Sub

    Private Sub menuWUsun_Click(sender As Object, e As EventArgs) Handles menuWUsun.Click
        If _EdytujObroty Then
            cmdUsuñ_Click(Me, System.EventArgs.Empty)
        Else
        End If
    End Sub

    Private Sub menuWWybierz_Click(sender As Object, e As EventArgs) Handles menuWWybierz.Click
        cmdStop_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWSingleL_Click(sender As Object, e As EventArgs) Handles menuWSingleL.Click
        dagObrotyMies.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagObrotyMies.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        dagObrotyMies.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagObrotyMies.Refresh()
    End Sub

    Private Sub menuWOdswiez_Click(sender As Object, e As EventArgs) Handles menuWOdswiez.Click
        OdswiezBaze()
    End Sub
End Class
