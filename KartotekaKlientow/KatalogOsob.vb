Public Class KatalogOsob
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

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
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents stbOsoby As System.Windows.Forms.StatusBar
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents lblIDDomu As System.Windows.Forms.Label
    Friend WithEvents HorizScrollTimer As Timer
    Friend WithEvents menuEdytuj As ContextMenuStrip
    Friend WithEvents menuEEdytuj As ToolStripMenuItem
    Friend WithEvents menuEDodaj As ToolStripMenuItem
    Friend WithEvents menuEUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents stbFiltry As StatusBar
    Friend WithEvents dagOsoby As DataGridView
    Friend WithEvents cmdFi As Button
    Friend WithEvents lblTelefonKom As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents cmdWszystko As Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogOsob))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblTelefonKom = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
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
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.stbOsoby = New System.Windows.Forms.StatusBar()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.menuEdytuj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuESingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.dagOsoby = New System.Windows.Forms.DataGridView()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuEdytuj.SuspendLayout()
        CType(Me.dagOsoby, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblTelefonKom)
        Me.Panel1.Controls.Add(Me.Label8)
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
        Me.Panel1.Size = New System.Drawing.Size(751, 96)
        Me.Panel1.TabIndex = 47
        '
        'lblTelefonKom
        '
        Me.lblTelefonKom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefonKom.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTelefonKom.ForeColor = System.Drawing.Color.Purple
        Me.lblTelefonKom.Location = New System.Drawing.Point(648, 24)
        Me.lblTelefonKom.Name = "lblTelefonKom"
        Me.lblTelefonKom.Size = New System.Drawing.Size(232, 16)
        Me.lblTelefonKom.TabIndex = 31
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(592, 24)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 16)
        Me.Label8.TabIndex = 30
        Me.Label8.Text = "Tel. kom."
        '
        'lblIDDomu
        '
        Me.lblIDDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIDDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIDDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblIDDomu.Location = New System.Drawing.Point(539, 24)
        Me.lblIDDomu.Name = "lblIDDomu"
        Me.lblIDDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblIDDomu.TabIndex = 29
        '
        'lblNrDomu
        '
        Me.lblNrDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNrDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNrDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblNrDomu.Location = New System.Drawing.Point(493, 24)
        Me.lblNrDomu.Name = "lblNrDomu"
        Me.lblNrDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblNrDomu.TabIndex = 11
        '
        'lblUlica
        '
        Me.lblUlica.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUlica.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUlica.ForeColor = System.Drawing.Color.Purple
        Me.lblUlica.Location = New System.Drawing.Point(352, 24)
        Me.lblUlica.Name = "lblUlica"
        Me.lblUlica.Size = New System.Drawing.Size(143, 16)
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
        Me.Label14.Location = New System.Drawing.Point(592, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(56, 16)
        Me.Label14.TabIndex = 24
        Me.Label14.Text = "Uwagi"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(592, 56)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(56, 16)
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "eMail"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(592, 40)
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
        Me.lblUwagi.Location = New System.Drawing.Point(648, 72)
        Me.lblUwagi.Name = "lblUwagi"
        Me.lblUwagi.Size = New System.Drawing.Size(232, 16)
        Me.lblUwagi.TabIndex = 20
        '
        'lbleMail
        '
        Me.lbleMail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbleMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lbleMail.ForeColor = System.Drawing.Color.Purple
        Me.lbleMail.Location = New System.Drawing.Point(648, 56)
        Me.lbleMail.Name = "lbleMail"
        Me.lbleMail.Size = New System.Drawing.Size(232, 16)
        Me.lbleMail.TabIndex = 18
        '
        'lblFax
        '
        Me.lblFax.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblFax.ForeColor = System.Drawing.Color.Purple
        Me.lblFax.Location = New System.Drawing.Point(648, 40)
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
        Me.lblMiejscowosc.Size = New System.Drawing.Size(160, 16)
        Me.lblMiejscowosc.TabIndex = 12
        '
        'lblNIP
        '
        Me.lblNIP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNIP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNIP.ForeColor = System.Drawing.Color.Purple
        Me.lblNIP.Location = New System.Drawing.Point(352, 56)
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
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(663, 344)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 22)
        Me.cmdStop.TabIndex = 49
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
        Me.cmdEdytuj.Location = New System.Drawing.Point(575, 344)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 22)
        Me.cmdEdytuj.TabIndex = 52
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
        Me.cmdUsuñ.Location = New System.Drawing.Point(487, 344)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 22)
        Me.cmdUsuñ.TabIndex = 51
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
        Me.cmdDodaj.Location = New System.Drawing.Point(399, 344)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 22)
        Me.cmdDodaj.TabIndex = 50
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 344)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 54
        Me.PictureBox1.TabStop = False
        '
        'stbOsoby
        '
        Me.stbOsoby.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbOsoby.Dock = System.Windows.Forms.DockStyle.None
        Me.stbOsoby.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbOsoby.Location = New System.Drawing.Point(8, 312)
        Me.stbOsoby.Name = "stbOsoby"
        Me.stbOsoby.ShowPanels = True
        Me.stbOsoby.Size = New System.Drawing.Size(751, 21)
        Me.stbOsoby.TabIndex = 57
        Me.stbOsoby.Text = "StatusBar1"
        '
        'HorizScrollTimer
        '
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
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 110)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(751, 22)
        Me.stbFiltry.TabIndex = 185
        Me.stbFiltry.Text = "StatusBar1"
        '
        'dagOsoby
        '
        Me.dagOsoby.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagOsoby.ContextMenuStrip = Me.menuEdytuj
        Me.dagOsoby.Location = New System.Drawing.Point(8, 134)
        Me.dagOsoby.Name = "dagOsoby"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagOsoby.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagOsoby.Size = New System.Drawing.Size(751, 176)
        Me.dagOsoby.TabIndex = 187
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(323, 112)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 189
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(380, 110)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystko.TabIndex = 188
        '
        'KatalogOsob
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(767, 374)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.dagOsoby)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.stbOsoby)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "KatalogOsob"
        Me.ShowInTaskbar = False
        Me.Text = "Osoby kontaktowe"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuEdytuj.ResumeLayout(False)
        CType(Me.dagOsoby, System.ComponentModel.ISupportInitialize).EndInit()
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


    Dim dbSelectSQLOsoby As String = ""

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

    Dim DataSetOsoby As New DataSet
    Dim DataViewOsoby As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    ''Osoby Kontaktowe
    'Public oIdentOso As String            'IDENTKLIENTA     Text, 6
    'Public oOsobaOso As String            'OSOBA            Text, 50
    'Public oVIPOso As Boolean              'VIP              boolean
    'Public oeMailOso As String            'EMAIL            Text, 50
    'Public oTelefonOso As String          'TELEFON          Text, 50
    'Public oStanowiskoOso As String       'STANOWISKO       Text, 50
    'Public oKompetencjeOso As String      'KOMPETENCJE      Text, 50
    'Public oWersjaOso As Integer          'WERSJA           Integer


    Private Sub KatalogOsob_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        'DataSet
        DataSetOsoby.Locale = New System.Globalization.CultureInfo("pl-PL")

        If Len(oIdentKli) > 0 Then
            dbSelectSQLOsoby = sSelectSQLOsoby & " WHERE IDENTKLIENTA='" & oIdentKli & "'"
        Else
            dbSelectSQLOsoby = sSelectSQLOsoby
            cmdDodaj.Enabled = False
            menuEDodaj.Enabled = False
        End If

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionOsoby)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectSQLOsoby, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLOsoby, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLOsoby, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLOsoby, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Osoby")

            'DBDeleteOsoby(dbCommandDelete, dbDataAdapter)
            'DBUpdateOsoby(dbCommandUpdate, dbDataAdapter)
            'DBInsertOsoby(dbCommandInsert, dbDataAdapter)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionOsoby)
            sqlCommandSelect = New SqlClient.SqlCommand(dbSelectSQLOsoby, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLOsoby, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLOsoby, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLOsoby, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Osoby")

            SQLDeleteOsoby(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateOsoby(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertOsoby(sqlCommandInsert, sqlDataAdapter)

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
                    'dbDataAdapter.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                    'dbDataAdapter.Fill(DataSetOsoby)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                    sqlDataAdapter.Fill(DataSetOsoby)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetOsoby.Tables(0).PrimaryKey = New DataColumn() {DataSetOsoby.Tables(0).Columns("IDENTKLIENTA"),
                                                                      DataSetOsoby.Tables(0).Columns("OSOBA")}

                'definiuj DataView
                DataViewOsoby = New DataView(DataSetOsoby.Tables(0))
                DataViewOsoby.AllowDelete = False
                DataViewOsoby.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagOsoby.BackgroundColor = System.Drawing.SystemColors.Control
                dagOsoby.GridColor = System.Drawing.SystemColors.ControlDark
                dagOsoby.ForeColor = System.Drawing.Color.Purple
                dagOsoby.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagOsoby.BorderStyle = BorderStyle.Fixed3D
                'dagOsoby.Dock = DockStyle.Fill
                dagOsoby.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagOsoby.AllowUserToAddRows = False
                dagOsoby.AllowUserToDeleteRows = False
                dagOsoby.AllowUserToOrderColumns = True
                dagOsoby.AllowUserToResizeColumns = True
                dagOsoby.AllowUserToResizeRows = True
                dagOsoby.ReadOnly = True
                dagOsoby.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagOsoby.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagOsoby.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagOsoby.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagOsoby.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagOsoby.ColumnHeadersVisible = True
                dagOsoby.ColumnHeadersHeight = 40
                dagOsoby.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagOsoby.RowHeadersVisible = True
                dagOsoby.RowHeadersWidth = 50
                dagOsoby.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagOsoby.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagOsoby.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagOsoby.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagOsoby.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagOsoby.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagOsoby.DefaultCellStyle.NullValue = ""
                dagOsoby.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagOsoby.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagOsoby.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagOsoby.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagOsoby.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagOsoby.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagOsoby.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagOsoby.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagOsoby.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagOsoby.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagOsoby.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagOsoby.DataMember = DataSetOsoby.Tables(0).TableName
                'dagOsoby.DataSource = DataViewOsoby
                '-----------------------------------

                TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(0).ColumnName, "Klient", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(1).ColumnName, "Osoba kontaktowa", 250, DataGridViewContentAlignment.MiddleLeft, True)
                LogColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(2).ColumnName, "VIP", 40, True)
                TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(3).ColumnName, "eMail", 250, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(4).ColumnName, "Telefon", 250, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(5).ColumnName, "Telefon komórkowy", 250, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(6).ColumnName, "Stanowisko", 250, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(7).ColumnName, "Kompetencje", 250, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(8).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                Me.Cursor = Cursors.WaitCursor
                dagOsoby.DataSource = DataViewOsoby
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetOsoby.Tables(0).Rows.Count > 0 Then
                    dagOsoby.CurrentCell = dagOsoby(0, 0)
                    dagOsoby.CurrentCell.Selected = True
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
            stbOsoby.BackColor = System.Drawing.SystemColors.Control
            stbOsoby.ForeColor = System.Drawing.Color.DarkGreen
            stbOsoby.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbOsoby.Panels.Add("stbPanelIleRec")
            stbOsoby.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbOsoby.Panels(0).Width = 80
            stbOsoby.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbOsoby.Panels(0).Text = "Iloœæ rec.: " & DataSetOsoby.Tables(0).Rows.Count.ToString

            stbOsoby.Panels.Add("stbPanelWiersz")
            stbOsoby.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbOsoby.Panels(1).Width = 80
            stbOsoby.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagOsoby.CurrentCell Is Nothing Then
                stbOsoby.Panels(1).Text = "Wiersz : "
            Else
                stbOsoby.Panels(1).Text = "Wiersz : " & dagOsoby.CurrentCell.RowIndex.ToString
            End If

            stbOsoby.Panels.Add("stbPanelKolumna")
            stbOsoby.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbOsoby.Panels(2).Width = 80
            stbOsoby.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagOsoby.CurrentCell Is Nothing Then
                stbOsoby.Panels(2).Text = "Kolumna : "
            Else
                stbOsoby.Panels(2).Text = "Kolumna : " & dagOsoby.CurrentCell.ColumnIndex.ToString
            End If

            stbOsoby.Panels.Add("stbPanelSort")
            stbOsoby.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbOsoby.Panels(3).Width = 150
            stbOsoby.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbOsoby.Panels(3).Text = "Sort : "

            stbOsoby.Panels.Add("stbPanelFiltr")
            stbOsoby.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbOsoby.Panels(4).Width = 150
            stbOsoby.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbOsoby.Panels(4).Text = "Filtr : "

            stbOsoby.Panels.Add("stbPanelReszta")
            stbOsoby.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbOsoby.Panels(5).Width = 150
            stbOsoby.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbOsoby.Panels(5).Text = ""

            stbOsoby.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagOsoby.TableStyles(0).RowHeaderWidth
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
        InitOsoby() 'inicjuj zmienne
        '-----------------------------------------------------------------
        'If _OCoMamRobic = "WYBIERAJ" Then
        '    dagOsoby.ContextMenuStrip = Me.menuWybieraj
        '    cmdStop.Text = "Wybierz"
        '    menuWybieraj.Visible = False
        '    menuWybieraj.Enabled = True
        '    MenuEdytuj.Visible = False
        '    MenuEdytuj.Enabled = False
        'Else
        '    dagOsoby.ContextMenuStrip = Me.MenuEdytuj
        '    cmdStop.Text = "Powrót"
        '    menuWybieraj.Visible = False
        '    menuWybieraj.Enabled = False
        '    MenuEdytuj.Visible = False
        '    MenuEdytuj.Enabled = True
        'End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagOsoby.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagOsoby.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetOsoby.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagOsoby.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagOsoby.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagOsoby.FirstDisplayedScrollingColumnIndex +
                        dagOsoby.GetCellDisplayRectangle(dagOsoby.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagOsoby.Left + 4), (dagOsoby.Top + 4))
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
        For ii = 0 To DataViewOsoby.Table.Columns.Count - 1
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
        dagOsoby_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub

    Private Sub KatalogAkcjiOpis_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub









    Private Sub dagOsoby_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagOsoby.CurrentCellChanged
        If Not StartFormy Then
            If dagOsoby.CurrentCell Is Nothing Then
                stbOsoby.Panels(1).Text = "Wi: "
                stbOsoby.Panels(2).Text = "Ko: "

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
                lblTelefonKom.Text = ""
                lblFax.Text = ""
                lbleMail.Text = ""
                lblUwagi.Text = ""
            Else
                stbOsoby.Panels(1).Text = "Wi: " & dagOsoby.CurrentCell.RowIndex.ToString
                stbOsoby.Panels(2).Text = "Ko: " & dagOsoby.CurrentCell.ColumnIndex.ToString

                If DataViewOsoby.Count = 0 Then
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
                    lblTelefonKom.Text = ""
                    lblFax.Text = ""
                    lbleMail.Text = ""
                    lblUwagi.Text = ""
                Else
                    'Public oIdentOso As String            'IDENTKLIENTA     Text, 6
                    'Public oOsobaOso As String            'OSOBA            Text, 50
                    'Public oVIPOso As Boolean              'VIP              boolean
                    'Public oeMailOso As String            'EMAIL            Text, 50
                    'Public oTelefonOso As String          'TELEFON          Text, 50
                    'Public oTelefonKomOso As String       'TELEFONKOM          Text, 50
                    'Public oStanowiskoOso As String       'STANOWISKO       Text, 50
                    'Public oKompetencjeOso As String      'KOMPETENCJE      Text, 50
                    'Public oWersjaOso As Integer          'WERSJA           Integer

                    IdentKlienta(GetTxtField(dagOsoby, 0))

                    lblIdent.Text = oIdentKli
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

                    lblTelefon.Text = GetTxtField(dagOsoby, 4)      'oTelefonKli
                    lblTelefonKom.Text = GetTxtField(dagOsoby, 5)      'oTelefonKomOso
                    lblFax.Text = oFaxKli
                    lbleMail.Text = GetTxtField(dagOsoby, 3)        'oeMailKli
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
                If dagOsoby.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagOsoby.FirstDisplayedScrollingColumnIndex +
                                    dagOsoby.GetCellDisplayRectangle(dagOsoby.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagOsoby.FirstDisplayedScrollingColumnIndex +
                                    dagOsoby.GetCellDisplayRectangle(dagOsoby.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagOsoby.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagOsoby.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagOsoby, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, DataViewOsoby, stbOsoby)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagOsoby, Mid(sender.name, 6, 3), "Osoby")
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
            CzyscFiltryDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, DataViewOsoby, stbOsoby)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbOsoby.Panels(0).Text = "Il.zap.: " & DataViewOsoby.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagOsoby_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagOsoby.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagOsoby.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagOsoby.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagOsoby_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagOsoby.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagOsoby, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagOsoby, e.RowIndex, 4)

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

    Private Sub dagOsoby_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagOsoby.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagOsoby_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagOsoby.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagOsoby_Scroll(sender As Object, e As ScrollEventArgs) Handles dagOsoby.Scroll
        'If (Not StartFormy) And (DataViewOsoby.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewOsoby.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagOsoby_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagOsoby.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagOsoby_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagOsoby.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagOsoby_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagOsoby.MouseUp
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
                stbOsoby.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbOsoby.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagOsoby(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbOsoby.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbOsoby.Panels(3).Text = "Sort: " & DataSetOsoby.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbOsoby.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagOsoby(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbOsoby.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbOsoby.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub



    Private Sub dagOsoby_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagOsoby.DoubleClick
        If _OCoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagOsoby_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagOsoby.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _OCoMamRobic = "WYBIERAJ" Then
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

    Private Sub InitOsoby()
        oIdentOso = oIdentKli
        oOsobaOso = ""
        oVIPOso = False
        oeMailOso = ""
        oTelefonOso = ""
        oTelefonKomOso = ""
        oStanowiskoOso = ""
        oKompetencjeOso = ""
        oWersjaOso = 1          'WERSJA         Integer, 2
    End Sub

    Private Sub PobierzOsoby()
        'pobierz wartosci przed aktualizacja
        oIdentOso = GetTxtField(dagOsoby, 0)
        oOsobaOso = GetTxtField(dagOsoby, 1)
        oVIPOso = GetLogField(dagOsoby, 2)
        oeMailOso = GetTxtField(dagOsoby, 3)
        oTelefonOso = GetTxtField(dagOsoby, 4)
        oTelefonKomOso = GetTxtField(dagOsoby, 5)
        oStanowiskoOso = GetTxtField(dagOsoby, 6)
        oKompetencjeOso = GetTxtField(dagOsoby, 7)
        oWersjaOso = GetNumField(dagOsoby, 8)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetOsoby.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzOsoby()
            oAktualizuj = False
            Dim Form As New EdytujKatalogOsob
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentOso
                findTheseVals(1) = oOsobaOso
                foundRow = DataSetOsoby.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDENTKLIENTA") = oIdentOso
                    'foundRow("OSOBA") = oOsobaOso
                    foundRow("VIP") = oVIPOso
                    foundRow("EMAIL") = oeMailOso
                    foundRow("TELEFON") = oTelefonOso
                    foundRow("TELEFONKOM") = oTelefonKomOso
                    foundRow("STANOWISKO") = oStanowiskoOso
                    foundRow("KOMPETENCJE") = oKompetencjeOso
                    foundRow("WERSJA") = ZmienWersje(oWersjaOso) 'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbOsoby.Panels(0).Text = "Il.rec.: " & DataViewOsoby.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagOsoby.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetOsoby.Tables(0).Rows.Count <= 0 Then
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
                PobierzOsoby()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentOso
                findTheseVals(1) = oOsobaOso
                foundRow = DataSetOsoby.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbOsoby.Panels(0).Text = "Il.rec.: " & DataViewOsoby.Count.ToString
                    AktualizujBaze()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitOsoby()
        '----------------------------------------------
        Do While True
            oAktualizuj = False
            Dim Form As New EdytujKatalogOsob
            Form.ShowDialog()
            Form.Dispose()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentOso
                findTheseVals(1) = oOsobaOso
                foundRow = DataSetOsoby.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Klient = " & oIdentOso & vbCrLf &
                        "Osoba = " & oOsobaOso,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetOsoby.Tables(0).NewRow()

                    NewRow("IDENTKLIENTA") = oIdentOso
                    NewRow("OSOBA") = oOsobaOso
                    NewRow("VIP") = oVIPOso
                    NewRow("EMAIL") = oeMailOso
                    NewRow("TELEFON") = oTelefonOso
                    NewRow("TELEFONKOM") = oTelefonKomOso
                    NewRow("STANOWISKO") = oStanowiskoOso
                    NewRow("KOMPETENCJE") = oKompetencjeOso
                    NewRow("WERSJA") = 1     'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetOsoby.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbOsoby.Panels(0).Text = "Il.rec.: " & DataViewOsoby.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagOsoby.Update()
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
            '        dbDataAdapter.Update(DataSetOsoby, "TABELA_Osoby")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetOsoby)
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
                    sqlDataAdapter.Update(DataSetOsoby, "TABELA_Osoby")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetOsoby)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub


    Private Sub OdswiezBaze()
        DataSetOsoby.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetOsoby)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetOsoby)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub




    '*************************************************
    '** obs³uga menu kontekstowego
    '*************************************************

    Private Sub KatalogOsob_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
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
        dagOsoby.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagOsoby.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagOsoby.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagOsoby.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub

End Class
