Public Class KatalogHarmonogramWizyt
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Dim _Opiekun As String = ""

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _Opiekun = ""
    End Sub


    Public Sub New(ByVal parOpiekun As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _Opiekun = parOpiekun
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
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblOpiekun As System.Windows.Forms.Label
    Friend WithEvents lblDzien As System.Windows.Forms.Label
    Friend WithEvents lblRokTydzien As System.Windows.Forms.Label
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

    Friend WithEvents stbKlienci As System.Windows.Forms.StatusBar
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents lblIDDomu As System.Windows.Forms.Label
    Friend WithEvents lblCoDalej As System.Windows.Forms.Label
    Friend WithEvents cmdFi As System.Windows.Forms.Button
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    Friend WithEvents stbFiltry As System.Windows.Forms.StatusBar
    Friend WithEvents HorizScrollTimer As System.Windows.Forms.Timer
    Friend WithEvents dagKlienci As DataGridView
    Friend WithEvents menuEdytuj As ContextMenuStrip
    Friend WithEvents menuEEdytuj As ToolStripMenuItem
    Friend WithEvents menuEDodaj As ToolStripMenuItem
    Friend WithEvents menuEUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents Label8 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogHarmonogramWizyt))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblDzien = New System.Windows.Forms.Label()
        Me.lblOpiekun = New System.Windows.Forms.Label()
        Me.lblRokTydzien = New System.Windows.Forms.Label()
        Me.lblCoDalej = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblIDDomu = New System.Windows.Forms.Label()
        Me.lblNrDomu = New System.Windows.Forms.Label()
        Me.lblUlica = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblTOFI = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
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
        Me.stbKlienci = New System.Windows.Forms.StatusBar()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.dagKlienci = New System.Windows.Forms.DataGridView()
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
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuEdytuj.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblDzien)
        Me.Panel1.Controls.Add(Me.lblOpiekun)
        Me.Panel1.Controls.Add(Me.lblRokTydzien)
        Me.Panel1.Controls.Add(Me.lblCoDalej)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.lblIDDomu)
        Me.Panel1.Controls.Add(Me.lblNrDomu)
        Me.Panel1.Controls.Add(Me.lblUlica)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblTOFI)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label11)
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
        Me.Panel1.Size = New System.Drawing.Size(868, 96)
        Me.Panel1.TabIndex = 47
        '
        'lblDzien
        '
        Me.lblDzien.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDzien.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblDzien.ForeColor = System.Drawing.Color.Purple
        Me.lblDzien.Location = New System.Drawing.Point(667, 40)
        Me.lblDzien.Name = "lblDzien"
        Me.lblDzien.Size = New System.Drawing.Size(102, 16)
        Me.lblDzien.TabIndex = 18
        '
        'lblOpiekun
        '
        Me.lblOpiekun.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOpiekun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOpiekun.ForeColor = System.Drawing.Color.Purple
        Me.lblOpiekun.Location = New System.Drawing.Point(667, 56)
        Me.lblOpiekun.Name = "lblOpiekun"
        Me.lblOpiekun.Size = New System.Drawing.Size(102, 16)
        Me.lblOpiekun.TabIndex = 19
        '
        'lblRokTydzien
        '
        Me.lblRokTydzien.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRokTydzien.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblRokTydzien.ForeColor = System.Drawing.Color.Purple
        Me.lblRokTydzien.Location = New System.Drawing.Point(667, 24)
        Me.lblRokTydzien.Name = "lblRokTydzien"
        Me.lblRokTydzien.Size = New System.Drawing.Size(102, 16)
        Me.lblRokTydzien.TabIndex = 17
        '
        'lblCoDalej
        '
        Me.lblCoDalej.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCoDalej.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCoDalej.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblCoDalej.ForeColor = System.Drawing.Color.Purple
        Me.lblCoDalej.Location = New System.Drawing.Point(352, 72)
        Me.lblCoDalej.Name = "lblCoDalej"
        Me.lblCoDalej.Size = New System.Drawing.Size(506, 16)
        Me.lblCoDalej.TabIndex = 30
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(305, 72)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(60, 16)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Co dalej . . . . ."
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
        Me.Label3.Location = New System.Drawing.Point(305, 24)
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
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Navy
        Me.Label13.Location = New System.Drawing.Point(592, 56)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(104, 16)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "Opiekun . . . . ."
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(592, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(104, 16)
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "Dzieñ tyg. . . . ."
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(592, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(104, 16)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "Rok / Tydzieñ . . . . ."
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(592, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(177, 16)
        Me.Label10.TabIndex = 15
        Me.Label10.Text = "Planowana wizyta u klienta :"
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
        Me.Label7.Location = New System.Drawing.Point(305, 8)
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
        Me.Label2.Location = New System.Drawing.Point(305, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 16)
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
        Me.cmdStop.Location = New System.Drawing.Point(780, 431)
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
        Me.cmdEdytuj.Location = New System.Drawing.Point(692, 431)
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
        Me.cmdUsuñ.Location = New System.Drawing.Point(604, 431)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 22)
        Me.cmdUsuñ.TabIndex = 51
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdUsuñ.Visible = False
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.ForeColor = System.Drawing.Color.Black
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(516, 431)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 22)
        Me.cmdDodaj.TabIndex = 50
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdDodaj.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 431)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 54
        Me.PictureBox1.TabStop = False
        '
        'stbKlienci
        '
        Me.stbKlienci.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbKlienci.Dock = System.Windows.Forms.DockStyle.None
        Me.stbKlienci.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbKlienci.Location = New System.Drawing.Point(8, 399)
        Me.stbKlienci.Name = "stbKlienci"
        Me.stbKlienci.ShowPanels = True
        Me.stbKlienci.Size = New System.Drawing.Size(868, 21)
        Me.stbKlienci.TabIndex = 57
        Me.stbKlienci.Text = "StatusBar1"
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(706, 112)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(20, 20)
        Me.cmdFi.TabIndex = 177
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWszystko.ForeColor = System.Drawing.Color.Black
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(732, 113)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 20)
        Me.cmdWszystko.TabIndex = 176
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 111)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(868, 22)
        Me.stbFiltry.TabIndex = 175
        Me.stbFiltry.Text = "StatusBar1"
        '
        'HorizScrollTimer
        '
        '
        'dagKlienci
        '
        Me.dagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagKlienci.ContextMenuStrip = Me.menuEdytuj
        Me.dagKlienci.Location = New System.Drawing.Point(8, 134)
        Me.dagKlienci.Name = "dagKlienci"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagKlienci.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagKlienci.Size = New System.Drawing.Size(869, 265)
        Me.dagKlienci.TabIndex = 181
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
        'KatalogHarmonogramWizyt
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(884, 461)
        Me.Controls.Add(Me.dagKlienci)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.stbKlienci)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "KatalogHarmonogramWizyt"
        Me.Text = "Harmonogram kontaktów z  klientami"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).EndInit()
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





    Dim dbSelectKlienci As String = "SELECT * FROM Klienci"

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    'Dim dbCommandDelete As OleDb.OleDbCommand
    Dim dbCommandUpdate As OleDb.OleDbCommand
    'Dim dbCommandInsert As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienci As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienci As SqlClient.SqlCommand
    'Dim sqlCommandDeleteKlienci As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienci As SqlClient.SqlCommand
    'Dim sqlCommandInsertKlienci As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienci As SqlClient.SqlDataAdapter
    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim CurYear As String = ""      'txt 4
    Dim CurWeekNo As Integer = 0
    Dim CurDayOfWeek As String = "" 'txt12

    Private Sub KatalogHarmonogramWizyt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        Me.Icon = My.Resources.MrJones

        Me.TopMost = True   'ustaw siê na szczycie
        Me.TopMost = False  'pozwol innym tez byc na szczycie...
        '-------------------------------
        'okreœl aktualny rok i tydzieñ
        Dim FirstDayOfWeek As Integer

        Dim DayOfYear As Integer = CType(SysData(), Date).DayOfYear
        Dim DayOfWeek As Integer = CType(SysData(), Date).DayOfWeek
        DayOfWeek = IIf(DayOfWeek = 0, 7, DayOfWeek)

        If (DayOfYear < DayOfWeek) Then
            FirstDayOfWeek = (DayOfWeek - DayOfYear + 1)
        Else
            FirstDayOfWeek = (DayOfYear - DayOfWeek) Mod 7
        End If
        FirstDayOfWeek = IIf(FirstDayOfWeek = 0, 7, FirstDayOfWeek)

        CurYear = Mid(SysData, 1, 4)
        CurWeekNo = Int((DayOfYear - DayOfWeek + (FirstDayOfWeek - 1)) / 7) + 1
        '-------------------------------
        'DataSet
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")


        'TxtColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, "", 0, HorizontalAlignment.Center)
        'NumColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, "", 0, HorizontalAlignment.Right, "F0")
        'TxtColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, "", 0, HorizontalAlignment.Center)

        'TxtColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, "Rok", 50, HorizontalAlignment.Center)
        'NumColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, "Tydz.", 50, HorizontalAlignment.Right, "F0")
        'TxtColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, "Dzieñ", 50, HorizontalAlignment.Center)

        'TxtColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, "Opiekun", 50, HorizontalAlignment.Center)
        'TxtColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, "Opiekun Co dalej ?", 200, HorizontalAlignment.Left)
        'TxtColoredColumnF(deleg, TSKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, "Opiekun Uwagi", 200, HorizontalAlignment.Left)




        '  Public oIdentKli As String         'IDENTKLIENTA   Text, 6
        'Public oNrTOFIKli As Long          'NRTOFI         Long
        '  Public oNrTOFITxtKli As String     'NRTOFITXT      Text 500

        '  Public oKartaPaybackKli As Boolean 'KARTAPAYBACK   Logical
        '  Public oNryKartPaybackKli As String 'NRYKARTPAYBACK   memo

        '  Public oNazwa1Kli As String        'NAZWA1         Text, 100
        '  Public oKodPoczKli As String       'KODPOCZTOWY    Text, 10
        '  Public oMiejscKli As String        'MIEJSCOWOSC    Text, 50
        '  Public oUlicaKli As String         'ULICA          Text, 50
        '  Public oNumNrDomuKli As Integer    'NUMNRDOMU      INTEGER
        '  Public oNrDomuKli As String        'NRDOMU         Text, 10
        '  Public oIDDomuKli As String        'IDDOMU         Text, 10
        '  Public oNIPKli As String           'NIP            Text, 15
        '  Public oTelefonKli As String       'TELEFON        Text, 50
        '  Public oFaxKli As String           'FAX            Text, 50
        '  Public oeMailKli As String         'EMAIL          Text, 50
        '  Public oWpisalKli As String        'KTOWPISAL      Text, 10   uzytkownik
        '  Public oRejonKli As String         'REJKLIENTA     Text, 20   
        '  Public oPKontaktKli As String      'PKONTAKT       Text, 10   data rrrr-mm-dd
        ''..............................
        'Public oWykazUrzadzenKli As String              'WYKLAZURZADZEN    Memo

        'Public oSposobWyboruDostawcyKli As String       'SPOSOBWYBORUDOSTAWCY    Text 30
        'Public oZainstalowanoMonitorKli As Boolean      'ZAINSTALOWANOMONITOR    Logical
        'Public oLiczbaUrzadzenKli As Integer            'LICZBAURZADZEN     Integer
        'Public oLiczbaWydrukowKli As Integer            'LICZBAWYDRUKOW     Integer
        'Public oPotencjalDrukuKli As String             'POTENCJALDRUKU     Text 2

        'Public oAktZakupMatEkspKli As Boolean           'AKTZAKUPMATEKSP    Logical
        'Public oAktRozliczaStronyDrukuKli As Boolean    'AKTROZLICZASTRONYDRUKU    Logical
        'Public oAktKorzystaZNajmuKli As Boolean         'AKTKORZYSTAZNAJMU  Logical
        'Public oZaintRozliczaniemStronDrukuKli As Boolean   'ZAINTROZLICZANIEMSTRONDRUKU    Logical
        'Public oZaintKorzystaniemZNajmuKli As Boolean   'ZAINTKORZYSTANIEMZNAJMU    Logical

        'Public oDataWeryfSegmentacjiKli As String          'DATAWERYFSEGMENTACJI  Text 10

        'Public oPlanDlugookresowyKli As Integer         'PLANDLUGOOKRESOWY     Integer
        'Public oPlanKrotkookresowyKli As Integer        'PLANKROTKOOKRESOWY    Integer
        ''......................................
        '  Public oRodzSprzetuLKli As Boolean 'RODZSPRZETUL Logical
        '  Public oRodzSprzetuAKli As Boolean 'RODZSPRZETUA Logical
        '  Public oRodzSprzetuTKli As Boolean 'RODZSPRZETUT Logical
        '  Public oOfertaTZlozeniaKli As String        'OFERTATZLOZENIA         Text, 4
        '  Public oOfertaPlikKli As String    'OFERTAPLIK     Text, 100
        '  Public oUwagikli As String         'UWAGI          Memo

        '  Public oSkupUwagikli As String        'SKUPUWAGI      Memo
        '  Public oSprzedOpiekunkli As String    'SPRZEDOPIEKUN    Text, 10   uzytkownik

        '  Public oSprzedOKontaktRKli As String  'SPRZEDOKONTAKTR  Text, 4    rok
        '  Public oSprzedOKontaktTKli As String  'SPRZEDOKONTAKTT  Integer    nr tygodnia
        '  Public oSprzedOKontaktDKli As String  'SPRZEDOKONTAKTD  Text, 12   dzien tygodnia
        '  Public oSprzedNKontaktRKli As String  'SPRZEDNKONTAKTR  Text, 4    rok
        '  Public oSprzedNKontaktTKli As String  'SPRZEDNKONTAKTT  Integer    nr tygodnia
        '  Public oSprzedNKontaktDKli As String  'SKUPNKONTAKTT    Text, 12    nr tygodnia

        '  Public oSprzedUwagiKli As String      'SPRZEDUWAGI      Memo
        '  Public oWersjaKli As Integer          'WERSJA           Integer
        '  Public oMarkerKli As Boolean          'MARKER           boolean
        '  Public oMarkerPKli As Boolean         'MARKERP          boolean
        'Public oAktywnyKli As Boolean         'AKTYWNY          boolean

        'Public oZakupPopRokKli As Double       'ZAKUPPOPROK        Double
        'Public oUslugiDodKli As String         'USLUGIDOD          Text, 200
        'Public oRZURap01 As String              'RZURAP01     Text, 100
        'Public oRZURap02 As String              'RZURAP02     Text, 100
        'Public oRZURap03 As String              'RZURAP03     Text, 100
        'Public oRZURap04 As String              'RZURAP04     Text, 100
        'Public oRZURap05 As String              'RZURAP05     Text, 100
        'Public oRZURap06 As String              'RZURAP06     Text, 100
        'Public oRZURap07 As String              'RZURAP07     Text, 100
        'Public oRZURap08 As String              'RZURAP08     Text, 100
        'Public oRZURap09 As String              'RZURAP09     Text, 100

        dbSelectKlienci = "SELECT " &
                                                 "SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," &
                                                 "SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," &
                                                 "SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," &
                                                 "SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," &
                                                 "SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," &
                                                 "SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," &
                                                 "SPRZEDOPIEKUN AS OPIEKUN," &
                                                 "SPRZEDUWAGI AS OPIEKUNUWAGI," &
                                               "IDENTKLIENTA AS NRKARTY, " &
                                               "NIP," &
                                               "NRTOFITXT," &
                                               "KARTAPAYBACK," &
                                               "NRYKARTPAYBACK," &
                                               "NAZWA1 AS NAZWAKLIENTA," &
                                               "MIEJSCOWOSC," &
                                               "KODPOCZTOWY," &
                                               "ULICA," &
                                               "NUMNRDOMU, "

        If ParL_DataBaseType = "MS ACCESS" Then
            dbSelectKlienci &= "IIF([NUMNRDOMU] MOD 2=0,'True','False') AS PARZYSTE, "
        Else
            dbSelectKlienci &= "CASE WHEN [NUMNRDOMU]%2=0 THEN CAST('True' AS BIT) ELSE CAST('False' AS BIT) END AS PARZYSTE, "
        End If

        dbSelectKlienci &=
                                               "NRDOMU," &
                                               "IDDOMU," &
                                               "REJKLIENTA AS REJONKLIENTA," &
                                               "TELEFON," &
                                               "FAX," &
                                               "EMAIL," &
                                               "RODZSPRZETUL," &
                                               "RODZSPRZETUA," &
                                               "RODZSPRZETUT," &
                                               "OFERTATZLOZENIA," &
                                               "OFERTAPLIK," &
                                               "PKONTAKT," &
                                               "SKUPUWAGI AS PROMOTORUWAGI," &
                                               "KTOWPISAL," &
                                               "UWAGI," &
                                               "MARKER, " &
                                               "MARKERP, " &
                                               "WERSJA " &
                                            "FROM Klienci "

        If Len(_Opiekun) > 0 Then
            dbSelectKlienci &= " WHERE (AKTYWNY='TRUE') AND (SPRZEDOPIEKUN='" & _Opiekun & "')  ORDER BY OPIEKUNKOLKONTAKTROK,OPIEKUNKOLKONTAKTTYDZ "
        Else
            dbSelectKlienci &= " WHERE AKTYWNY='TRUE' ORDER BY OPIEKUNKOLKONTAKTROK,OPIEKUNKOLKONTAKTTYDZ "
        End If




        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectKlienci, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            Dim objParam As SqlClient.SqlParameter
            ''------------------------------------------
            ''komenda DELETE
            ''parametry filtrowania
            'objParam = sqlCommandDeleteKlienci.Parameters.Add("@orygSymbol", SqlDbType.Char, 10, "NRKARTY")
            'objParam.SourceVersion = DataRowVersion.Original
            'objParam = sqlCommandDeleteKlienci.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            'objParam.SourceVersion = DataRowVersion.Original
            'sqlDataAdapterKlienci.DeleteCommand = sqlCommandDeleteKlienci



            '------------------------------------------
            'komenda UPDATE
            'parametry aktualizacji
            sqlCommandUpdateKlienci.Parameters.Add("@oSprzedOKontaktRKli", SqlDbType.Char, 4, "OPIEKUNOSTKONTAKTROK")
            sqlCommandUpdateKlienci.Parameters.Add("@oSprzedOKontaktTKli", SqlDbType.Int, Nothing, "OPIEKUNOSTKONTAKTTYDZ")
            sqlCommandUpdateKlienci.Parameters.Add("@oSprzedOKontaktDKli", SqlDbType.Char, 12, "OPIEKUNOSTKONTAKTDZIEN")

            sqlCommandUpdateKlienci.Parameters.Add("@oSprzedNKontaktRKli", SqlDbType.Char, 4, "OPIEKUNKOLKONTAKTROK")
            sqlCommandUpdateKlienci.Parameters.Add("@oSprzedNKontaktTKli", SqlDbType.Int, Nothing, "OPIEKUNKOLKONTAKTTYDZ")
            sqlCommandUpdateKlienci.Parameters.Add("@oSprzedNKontaktDKli", SqlDbType.Char, 12, "OPIEKUNKOLKONTAKTDZIEN")

            sqlCommandUpdateKlienci.Parameters.Add("@oSprzedOpiekunKli", SqlDbType.Char, 10, "OPIEKUN")
            sqlCommandUpdateKlienci.Parameters.Add("@oSprzedUwagiKli", SqlDbType.Text, Nothing, "OPIEKUNUWAGI")
            'sqlCommandUpdateKlienci.Parameters.Add("@oIdentKli", sqldbtype.Char, 6, "NRKARTY")
            sqlCommandUpdateKlienci.Parameters.Add("@oNIPKli", SqlDbType.Char, 15, "NIP")
            sqlCommandUpdateKlienci.Parameters.Add("@oNrTofiTxtKli", SqlDbType.VarChar, 500, "NRTOFITXT")
            sqlCommandUpdateKlienci.Parameters.Add("@oKartaPaybackKli", SqlDbType.Bit, Nothing, "KARTAPAYBACK")
            sqlCommandUpdateKlienci.Parameters.Add("@oNryKartPaybackKli", SqlDbType.Text, Nothing, "NRYKARTPAYBACK")
            sqlCommandUpdateKlienci.Parameters.Add("@oNazwa1Kli", SqlDbType.VarChar, 100, "NAZWAKLIENTA")
            sqlCommandUpdateKlienci.Parameters.Add("@oMiejscKli", SqlDbType.VarChar, 50, "MIEJSCOWOSC")
            sqlCommandUpdateKlienci.Parameters.Add("@oKodPoczKli", SqlDbType.Char, 10, "KODPOCZTOWY")
            sqlCommandUpdateKlienci.Parameters.Add("@oUlicaKli", SqlDbType.VarChar, 50, "ULICA")
            sqlCommandUpdateKlienci.Parameters.Add("@oNumNrDomuKli", SqlDbType.Int, Nothing, "NUMNRDOMU")
            'parzyste
            sqlCommandUpdateKlienci.Parameters.Add("@oNrDomuKli", SqlDbType.Char, 10, "NRDOMU")
            sqlCommandUpdateKlienci.Parameters.Add("@oIDDomuKli", SqlDbType.Char, 10, "IDDOMU")
            sqlCommandUpdateKlienci.Parameters.Add("@oRejonKli", SqlDbType.Char, 20, "REJONKLIENTA")
            sqlCommandUpdateKlienci.Parameters.Add("@oTelefonKli", SqlDbType.VarChar, 50, "TELEFON")
            sqlCommandUpdateKlienci.Parameters.Add("@oFaxKli", SqlDbType.VarChar, 50, "FAX")
            sqlCommandUpdateKlienci.Parameters.Add("@oeMailKli", SqlDbType.VarChar, 50, "EMAIL")
            sqlCommandUpdateKlienci.Parameters.Add("@oRodzSprzetuLKli", SqlDbType.Bit, Nothing, "RODZSPRZETUL")
            sqlCommandUpdateKlienci.Parameters.Add("@oRodzSprzetuAKli", SqlDbType.Bit, Nothing, "RODZSPRZETUA")
            sqlCommandUpdateKlienci.Parameters.Add("@oRodzSprzetuTKli", SqlDbType.Bit, Nothing, "RODZSPRZETUT")

            sqlCommandUpdateKlienci.Parameters.Add("@oOfertaTZLOZENIAKli", SqlDbType.VarChar, 4, "OFERTATZLOZENIA")
            sqlCommandUpdateKlienci.Parameters.Add("@oOfertaPLIKKli", SqlDbType.VarChar, 100, "OFERTAPLIK")
            sqlCommandUpdateKlienci.Parameters.Add("@oPKontaktKli", SqlDbType.Char, 10, "PKONTAKT")

            sqlCommandUpdateKlienci.Parameters.Add("@oSkupUwagiKli", SqlDbType.Text, Nothing, "PROMOTORUWAGI")
            sqlCommandUpdateKlienci.Parameters.Add("@oWpisalKli", SqlDbType.Char, 10, "KTOWPISAL")
            sqlCommandUpdateKlienci.Parameters.Add("@oUwagiKli", SqlDbType.Text, Nothing, "UWAGI")

            sqlCommandUpdateKlienci.Parameters.Add("@oMarkerKli", SqlDbType.Bit, Nothing, "MARKER")
            sqlCommandUpdateKlienci.Parameters.Add("@oMarkerPKli", SqlDbType.Bit, Nothing, "MARKERP")
            sqlCommandUpdateKlienci.Parameters.Add("@oWersjaKli", SqlDbType.Int, Nothing, "WERSJA")

            'parametry filtrowania
            objParam = sqlCommandUpdateKlienci.Parameters.Add("@orygSymbol", SqlDbType.Char, 10, "NRKARTY")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandUpdateKlienci.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapterKlienci.UpdateCommand = sqlCommandUpdateKlienci



            ''------------------------------------------
            ''komenda INSERT
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oIdentKli", SqlDbType.Char, 6, "NRKARTY")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oNIPKli", SqlDbType.Char, 15, "NIP")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oNrTofiTxtKli", SqlDbType.VarChar, 500, "NRTOFITXT")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oKartaPaybackKli", SqlDbType.Bit, Nothing, "KARTAPAYBACK")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oNryKartPaybackKli", SqlDbType.Text, Nothing, "NRYKARTPAYBACK")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oNazwa1Kli", SqlDbType.VarChar, 100, "NAZWAKLIENTA")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oMiejscKli", SqlDbType.VarChar, 50, "MIEJSCOWOSC")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oKodPoczKli", SqlDbType.Char, 10, "KODPOCZTOWY")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oUlicaKli", SqlDbType.VarChar, 50, "ULICA")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oNumNrDomuKli", SqlDbType.Int, Nothing, "NUMNRDOMU")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oNrDomuKli", SqlDbType.Char, 10, "NRDOMU")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oIDDomuKli", SqlDbType.Char, 10, "IDDOMU")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oRejonKli", SqlDbType.Char, 20, "REJONKLIENTA")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oTelefonKli", SqlDbType.VarChar, 50, "TELEFON")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oFaxKli", SqlDbType.VarChar, 50, "FAX")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oeMailKli", SqlDbType.VarChar, 50, "EMAIL")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oRodzSprzetuLKli", SqlDbType.Bit, Nothing, "RODZSPRZETUL")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oRodzSprzetuAKli", SqlDbType.Bit, Nothing, "RODZSPRZETUA")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oRodzSprzetuTKli", SqlDbType.Bit, Nothing, "RODZSPRZETUT")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oOfertaTZLOZENIAKli", SqlDbType.VarChar, 4, "OFERTATZLOZENIA")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oOfertaPLIKKli", SqlDbType.VarChar, 100, "OFERTAPLIK")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oPKontaktKli", SqlDbType.Char, 10, "PKONTAKT")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSkupUwagiKli", SqlDbType.Text, Nothing, "PROMOTORUWAGI")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSprzedOpiekunKli", SqlDbType.Char, 10, "OPIEKUN")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSprzedOKontaktRKli", SqlDbType.Char, 4, "OPIEKUNOSTKONTAKTROK")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSprzedOKontaktTKli", SqlDbType.Int, Nothing, "OPIEKUNOSTKONTAKTTYDZ")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSprzedOKontaktDKli", SqlDbType.Char, 12, "OPIEKUNOSTKONTAKTDZIEN")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSprzedNKontaktRKli", SqlDbType.Char, 4, "OPIEKUNKOLKONTAKTROK")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSprzedNKontaktTKli", SqlDbType.Int, Nothing, "OPIEKUNKOLKONTAKTTYDZ")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSprzedNKontaktDKli", SqlDbType.Char, 12, "OPIEKUNKOLKONTAKTDZIEN")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oSprzedUwagiKli", SqlDbType.Text, Nothing, "OPIEKUNUWAGI")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oWpisalKli", SqlDbType.Char, 10, "KTOWPISAL")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oUwagiKli", SqlDbType.Text, Nothing, "UWAGI")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oMarkerKli", SqlDbType.Bit, Nothing, "MARKER")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oMarkerPKli", SqlDbType.Bit, Nothing, "MARKERP")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertKlienci.Parameters.Add("@oWersjaKli", SqlDbType.Int, Nothing, "WERSJA")
            'objParam.SourceVersion = DataRowVersion.Current
            'sqlDataAdapterKlienci.InsertCommand = sqlCommandInsertKlienci

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienci.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKlienci.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapter.Fill(DataSetKlienci)
                    'dbConnection.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}

                'definiuj DataView
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                DataViewKlienci.AllowDelete = False
                DataViewKlienci.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagKlienci.BackgroundColor = System.Drawing.SystemColors.Control
                dagKlienci.GridColor = System.Drawing.SystemColors.ControlDark
                dagKlienci.ForeColor = System.Drawing.Color.Purple
                dagKlienci.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagKlienci.BorderStyle = BorderStyle.Fixed3D
                'dagKlienci.Dock = DockStyle.Fill
                dagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagKlienci.AllowUserToAddRows = False
                dagKlienci.AllowUserToDeleteRows = False
                dagKlienci.AllowUserToOrderColumns = True
                dagKlienci.AllowUserToResizeColumns = True
                dagKlienci.AllowUserToResizeRows = True
                dagKlienci.ReadOnly = True
                dagKlienci.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagKlienci.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagKlienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagKlienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagKlienci.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagKlienci.ColumnHeadersVisible = True
                dagKlienci.ColumnHeadersHeight = 40
                dagKlienci.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagKlienci.RowHeadersVisible = True
                dagKlienci.RowHeadersWidth = 50
                dagKlienci.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagKlienci.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagKlienci.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagKlienci.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagKlienci.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagKlienci.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagKlienci.DefaultCellStyle.NullValue = ""
                dagKlienci.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagKlienci.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagKlienci.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagKlienci.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagKlienci.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagKlienci.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagKlienci.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagKlienci.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagKlienci.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagKlienci.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagKlienci.DataMember = DataSetKlienci.Tables(0).TableName
                'dagKlienci.DataSource = DataViewKlienci
                '-----------------------------------

                'dbSelectKlienci = "SELECT " &
                '                                 "SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," &
                '                                 "SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," &
                '                                 "SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," &
                '                                 "SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," &
                '                                 "SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," &
                '                                 "SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," &
                '                                 "SPRZEDOPIEKUN AS OPIEKUN," &
                '                                 "SPRZEDUWAGI AS OPIEKUNUWAGI," &
                '                               "IDENTKLIENTA AS NRKARTY, " &
                '                               "NIP," &
                '                               "NRTOFITXT," &
                '                               "KARTAPAYBACK," &
                '                               "NRYKARTPAYBACK," &
                '                               "NAZWA1 AS NAZWAKLIENTA," &
                '                               "MIEJSCOWOSC," &
                '                               "KODPOCZTOWY," &
                '                               "ULICA," &
                '                               "NUMNRDOMU, "
                ''parzyste
                'dbSelectKlienci &=
                '                               "NRDOMU," &
                '                               "IDDOMU," &
                '                               "REJKLIENTA AS REJONKLIENTA," &
                '                               "TELEFON," &
                '                               "FAX," &
                '                               "EMAIL," &
                '                               "RODZSPRZETUL," &
                '                               "RODZSPRZETUA," &
                '                               "RODZSPRZETUT," &
                '                               "OFERTATZLOZENIA," &
                '                               "OFERTAPLIK," &
                '                               "PKONTAKT," &
                '                               "SKUPUWAGI AS PROMOTORUWAGI," &
                '                               "KTOWPISAL," &
                '                               "UWAGI," &
                '                               "MARKER, " &
                '                               "MARKERP, " &
                '                               "WERSJA " &
                '                            "FROM Klienci "

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F0", False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "Rok", 50, DataGridViewContentAlignment.MiddleCenter, True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "Tydz.", 50, DataGridViewContentAlignment.MiddleRight, "F0", True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Dzieñ", 50, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "Opiekun", 50, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                '---------------------------
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "", 0, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Parzysty", 40, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "", 0, False)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "", 0, False)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "", 0, False)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "", 0, False)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "", 0, False)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F0", False)


                Me.Cursor = Cursors.WaitCursor
                dagKlienci.DataSource = DataViewKlienci
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                    dagKlienci.CurrentCell = dagKlienci(3, 0)
                    dagKlienci.CurrentCell.Selected = True
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
            stbKlienci.BackColor = System.Drawing.SystemColors.Control
            stbKlienci.ForeColor = System.Drawing.Color.DarkGreen
            stbKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbKlienci.Panels.Add("stbPanelIleRec")
            stbKlienci.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(0).Width = 80
            stbKlienci.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(0).Text = "Iloœæ rec.: " & DataSetKlienci.Tables(0).Rows.Count.ToString

            stbKlienci.Panels.Add("stbPanelWiersz")
            stbKlienci.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(1).Width = 80
            stbKlienci.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(1).Text = "Wi: "
                'stbKlienci.Panels(2).Text = "Ko: "
            Else
                stbKlienci.Panels(1).Text = "Wi: " & dagKlienci.CurrentCell.RowIndex.ToString
                'stbKlienci.Panels(2).Text = "Ko: " & dagKlienci.CurrentCell.ColumnIndex.ToString
            End If

            stbKlienci.Panels.Add("stbPanelKolumna")
            stbKlienci.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(2).Width = 80
            stbKlienci.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagKlienci.CurrentCell Is Nothing Then
                'stbKlienci.Panels(1).Text = "Wi: "
                stbKlienci.Panels(2).Text = "Ko: "
            Else
                'stbKlienci.Panels(1).Text = "Wi: " & dagKlienci.CurrentCell.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & dagKlienci.CurrentCell.ColumnIndex.ToString
            End If

            stbKlienci.Panels.Add("stbPanelSort")
            stbKlienci.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(3).Width = 150
            stbKlienci.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(3).Text = "Sort : "

            stbKlienci.Panels.Add("stbPanelFiltr")
            stbKlienci.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(4).Width = 150
            stbKlienci.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(4).Text = "Filtr : "

            stbKlienci.Panels.Add("stbPanelReszta")
            stbKlienci.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbKlienci.Panels(5).Width = 150
            stbKlienci.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(5).Text = ""

            stbKlienci.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagKlienci.TableStyles(0).RowHeaderWidth
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



            'przesun do bie¿¹cego roku
            Dim bRok As String = ""
            Dim bWeek As String = ""
            Dim irow As Integer = 0
            Try
                If DataViewKlienci.Count > 0 Then
                    For irow = 0 To DataViewKlienci.Count - 1
                        bRok = IIf(IsDBNull(DataViewKlienci.Item(irow).Item("OPIEKUNKOLKONTAKTROK")), "", DataViewKlienci.Item(irow).Item("OPIEKUNKOLKONTAKTROK"))
                        bWeek = IIf(IsDBNull(DataViewKlienci.Item(irow).Item("OPIEKUNKOLKONTAKTTYDZ")), 7, DataViewKlienci.Item(irow).Item("OPIEKUNKOLKONTAKTTYDZ"))

                        If bRok < CurYear Then
                        ElseIf bRok = CurYear Then
                            If bWeek < CurWeekNo Then
                            Else
                                'irow - pokazuje pierwszy wiersz do wyswietlnia...
                                Exit For
                            End If
                        Else
                        End If
                    Next

                    dagKlienci.CurrentCell = dagKlienci(3, irow)
                    dagKlienci.CurrentCell.Selected = True
                End If

            Catch ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Finally
            End Try
            '------------------------
            dagKlienci.Focus()
            dagKlienci_CurrentCellChanged(Me, System.EventArgs.Empty)
        End If

        InitKlienci() 'inicjuj zmienne
        '-----------------------------------------------------------------
        'If _OCoMamRobic = "WYBIERAJ" Then
        '    dagKlienci.ContextMenuStrip = Me.menuWybieraj
        '    cmdStop.Text = "Wybierz"
        '    menuWybieraj.Visible = False
        '    menuWybieraj.Enabled = True
        '    menuEdytuj.Visible = False
        '    menuEdytuj.Enabled = False
        'Else
        '    dagKlienci.ContextMenuStrip = Me.menuEdytuj
        '    cmdStop.Text = "Powrót"
        '    menuWybieraj.Visible = False
        '    menuWybieraj.Enabled = False
        '    menuEdytuj.Visible = False
        '    menuEdytuj.Enabled = True
        'End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagKlienci.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagKlienci.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetKlienci.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagKlienci.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagKlienci.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                        dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagKlienci.Left + 4), (dagKlienci.Top + 4))
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
        For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
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
        dagKlienci_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub



    Private Sub KatalogHarmonogramWizyt_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub











    Private Sub dagKlienci_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.CurrentCellChanged
        If Not StartFormy Then
            If dagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(1).Text = "Wi: "
                stbKlienci.Panels(2).Text = "Ko: "

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
                lblNIP.Text = ""
                lblCoDalej.Text = ""
                lblRokTydzien.Text = ""
                lblDzien.Text = ""
                lblOpiekun.Text = ""
            Else
                stbKlienci.Panels(1).Text = "Wi: " & dagKlienci.CurrentCell.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & dagKlienci.CurrentCell.ColumnIndex.ToString

                If DataViewKlienci.Count = 0 Then
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
                    lblNIP.Text = ""
                    lblCoDalej.Text = ""
                    lblRokTydzien.Text = ""
                    lblDzien.Text = ""
                    lblOpiekun.Text = ""
                Else
                    ''--------------------------------
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F0", False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "Rok", 50, DataGridViewContentAlignment.MiddleCenter, True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "Tydz.", 50, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Dzieñ", 50, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "Opiekun", 50, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    ''---------------------------
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "", 0, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Parzysty", 40, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "", 0, False)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "", 0, False)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "", 0, False)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "", 0, False)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "", 0, False)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F0", False)


                    lblIdent.Text = GetTxtField(dagKlienci, 8)
                    lblTOFI.Text = GetTxtField(dagKlienci, 10)
                    lblNazwaHandlowa.Text = GetTxtField(dagKlienci, 13)
                    lblNazwa1.Text = "" 'GetTxtField(dagKlienci, 13)
                    lblNazwa2.Text = "" 'GetTxtField(dagKlienci, 8)
                    lblKod.Text = GetTxtField(dagKlienci, 15)
                    lblMiejscowosc.Text = GetTxtField(dagKlienci, 14)
                    lblUlica.Text = GetTxtField(dagKlienci, 16)
                    lblNrDomu.Text = GetIntField(dagKlienci, 17).ToString("F0")
                    lblIDDomu.Text = GetTxtField(dagKlienci, 20)
                    lblAdres.Text = ""
                    lblNIP.Text = GetTxtField(dagKlienci, 9)

                    lblCoDalej.Text = GetTxtField(dagKlienci, 7)
                    lblRokTydzien.Text = GetTxtField(dagKlienci, 3) & " / " & GetTxtField(dagKlienci, 4)
                    lblDzien.Text = GetTxtField(dagKlienci, 5)
                    lblOpiekun.Text = GetTxtField(dagKlienci, 6)
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
                If dagKlienci.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                                    dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                                    dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagKlienci.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagKlienci.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagKlienci, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, DataViewKlienci, stbKlienci)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagKlienci, Mid(sender.name, 6, 3), "Klienci")
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
            CzyscFiltryDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, DataViewKlienci, stbKlienci)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbKlienci.Panels(0).Text = "Il.zap.: " & DataViewKlienci.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagKlienci_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagKlienci.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagKlienci.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagKlienci.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    '**************************************************************
    '** procedura okreslania koloru w DataGrid
    '**************************************************************

    Private Sub dagKlienci_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagKlienci.CellFormatting
        Dim bRok As String = ""
        Dim bWeek As String = ""

        If Not StartFormy Then
            '-----------------------
            'zmieniamy ForeColor

            '"SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," & _
            '"SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," & _
            '"SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," & _

            bRok = IIf(IsDBNull(DataViewKlienci.Item(e.RowIndex).Item("OPIEKUNKOLKONTAKTROK")), "", DataViewKlienci.Item(e.RowIndex).Item("OPIEKUNKOLKONTAKTROK"))
            bWeek = IIf(IsDBNull(DataViewKlienci.Item(e.RowIndex).Item("OPIEKUNKOLKONTAKTTYDZ")), 7, DataViewKlienci.Item(e.RowIndex).Item("OPIEKUNKOLKONTAKTTYDZ"))

            If bRok < CurYear Then
                e.CellStyle.ForeColor = Color.Black
            ElseIf bRok = CurYear Then
                If bWeek < CurWeekNo Then
                    e.CellStyle.ForeColor = Color.Black
                ElseIf bWeek = CurWeekNo Then
                    e.CellStyle.ForeColor = Color.Red
                Else
                    e.CellStyle.ForeColor = Color.Navy
                End If
            Else
                e.CellStyle.ForeColor = Color.Navy
            End If
            '-----------------------
            ''zmieniamy BackColor
            'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
            'Select Case Wal
            '    Case "EUR"
            '        e.CellStyle.BackColor = Color.White
            '    Case Else
            'End Select
            '-----------------------
        End If
    End Sub




    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagKlienci_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagKlienci.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKlienci_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKlienci_Scroll(sender As Object, e As ScrollEventArgs) Handles dagKlienci.Scroll
        'If (Not StartFormy) And (DataViewKlienci.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewKlienci.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagKlienci_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKlienci_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKlienci.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagKlienci_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKlienci.MouseUp
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
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagKlienci(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbKlienci.Panels(3).Text = "Sort: " & DataSetKlienci.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagKlienci(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub








    Private Sub dagKlienci_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.DoubleClick
        If oCoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagKlienci_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagKlienci.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If oCoMamRobic = "WYBIERAJ" Then
                    cmdStop_Click(sender, e)
                Else
                    cmdEdytuj_Click(sender, e)
                End If
            Case Keys.Insert
                'cmdDodaj_Click(sender, e)
            Case Keys.Delete
                'cmdUsuñ_Click(sender, e)
            Case Else
        End Select
    End Sub

    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************

    Private Sub InitKlienci()
        oIdentKli = ""
        oNIPKli = ""
        oNrTOFIKli = 0
        oNrTOFITxtKli = ""
        oKartaPaybackKli = False
        oNryKartPaybackKli = ""

        oNazwa1Kli = ""
        oMiejscKli = ""
        oKodPoczKli = ""
        oUlicaKli = ""
        oNumNrDomuKli = 0
        oNrDomuKli = ""
        oIDDomuKli = ""
        oRejonKli = ""
        oTelefonKli = ""
        oFaxKli = ""
        oeMailKli = ""

        oWykazUrzadzenKli = ""
        oSposobWyboruDostawcyKli = ""
        oZainstalowanoMonitorKli = False
        oLiczbaUrzadzenKli = 0
        oLiczbaWydrukowKli = 0
        oPotencjalDrukuKli = ""
        oAktZakupMatEkspKli = False
        oAktRozliczaStronyDrukuKli = False
        oAktKorzystaZNajmuKli = False
        oZaintRozliczaniemStronDrukuKli = False
        oZaintKorzystaniemZNajmuKli = False
        oDataWeryfSegmentacjiKli = ""
        oPlanDlugookresowyKli = 0
        oPlanKrotkookresowyKli = 0

        oRodzSprzetuLKli = False
        oRodzSprzetuAKli = False
        oRodzSprzetuTKli = False
        oOfertaTZlozeniaKli = ""
        oOfertaPlikKli = ""
        oPKontaktKli = ""
        oSkupUwagikli = ""
        oSprzedOpiekunkli = ""
        oSprzedOKontaktRKli = Mid(SysData, 1, 4)
        oSprzedOKontaktTKli = ""
        oSprzedOKontaktDKli = ""
        oSprzedNKontaktRKli = Mid(SysData, 1, 4)
        oSprzedNKontaktTKli = ""
        oSprzedNKontaktDKli = ""
        oSprzedUwagiKli = ""
        oWpisalKli = ""
        oUwagikli = ""
        oMarkerKli = False
        oMarkerPKli = False
        oAktywnyKli = True
        oZakupPopRokKli = 0
        oUslugiDodKli = ""
        oRZURap01 = ""
        oRZURap02 = ""
        oRZURap03 = ""
        oRZURap04 = ""
        oRZURap05 = ""
        oRZURap06 = ""
        oRZURap07 = ""
        oRZURap08 = ""
        oRZURap09 = ""

        oWersjaKli = 1          'WERSJA         Integer, 2
    End Sub

    Private Sub PobierzKlienci()
        ''--------------------------------
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
        'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F0", False)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "Rok", 50, DataGridViewContentAlignment.MiddleCenter, True)
        'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "Tydz.", 50, DataGridViewContentAlignment.MiddleRight, "F0", True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Dzieñ", 50, DataGridViewContentAlignment.MiddleCenter, True)

        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "Opiekun", 50, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
        ''---------------------------
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "", 0, False)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
        'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Parzysty", 40, True)

        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "", 0, False)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "", 0, False)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "", 0, False)

        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
        'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "", 0, False)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "", 0, False)
        'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F0", False)


        'pobierz wartosci przed aktualizacja
        oSprzedOKontaktRKli = GetTxtField(dagKlienci, 0)            'GetTxtField(dagKlienci, 44)
        oSprzedOKontaktTKli = GetTxtField(dagKlienci, 1)            'GetTxtField(dagKlienci, 45)
        oSprzedOKontaktDKli = GetTxtField(dagKlienci, 2)            'GetTxtField(dagKlienci, 46)

        oSprzedNKontaktRKli = GetTxtField(dagKlienci, 3)            'GetTxtField(dagKlienci, 47)
        oSprzedNKontaktTKli = GetTxtField(dagKlienci, 4)            'GetTxtField(dagKlienci, 48)
        oSprzedNKontaktDKli = GetTxtField(dagKlienci, 5)            'GetTxtField(dagKlienci, 49)

        oSprzedOpiekunkli = GetTxtField(dagKlienci, 6)              'GetTxtField(dagKlienci, 42)
        oSprzedUwagiKli = GetTxtField(dagKlienci, 7)                'GetTxtField(dagKlienci, 51)
        '-----------------------------------------
        oIdentKli = GetTxtField(dagKlienci, 8)                      'GetTxtField(dagKlienci, 0)
        oNIPKli = GetTxtField(dagKlienci, 9)                       'GetTxtField(dagKlienci, 1)
        oNrTOFITxtKli = GetTxtField(dagKlienci, 10)                 'GetTxtField(dagKlienci, 2)
        oKartaPaybackKli = GetLogField(dagKlienci, 11)              'GetlOGField(dagKlienci, 1)
        oNryKartPaybackKli = GetTxtField(dagKlienci, 12)

        oNazwa1Kli = GetTxtField(dagKlienci, 13)
        oMiejscKli = GetTxtField(dagKlienci, 14)
        oKodPoczKli = GetTxtField(dagKlienci, 15)
        oUlicaKli = GetTxtField(dagKlienci, 16)
        oNumNrDomuKli = GetIntField(dagKlienci, 17)
        '18 parzyste/nieparzyste
        oNrDomuKli = GetTxtField(dagKlienci, 19)
        oIDDomuKli = GetTxtField(dagKlienci, 20)
        oRejonKli = GetTxtField(dagKlienci, 21)
        oTelefonKli = GetTxtField(dagKlienci, 22)
        oFaxKli = GetTxtField(dagKlienci, 23)
        oeMailKli = GetTxtField(dagKlienci, 24)
        oRodzSprzetuLKli = GetLogField(dagKlienci, 25)
        oRodzSprzetuAKli = GetLogField(dagKlienci, 26)
        oRodzSprzetuTKli = GetLogField(dagKlienci, 27)

        oOfertaTZlozeniaKli = GetTxtField(dagKlienci, 28)
        oOfertaPlikKli = GetTxtField(dagKlienci, 29)
        oPKontaktKli = GetTxtField(dagKlienci, 30)

        oSkupUwagikli = GetTxtField(dagKlienci, 31)
        oWpisalKli = GetTxtField(dagKlienci, 32)
        oUwagikli = GetTxtField(dagKlienci, 33)
        oMarkerKli = GetLogField(dagKlienci, 34)
        oMarkerPKli = GetLogField(dagKlienci, 35)
        oWersjaKli = GetDblField(dagKlienci, 36)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie przegl¹daæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            Do While True
                oOperacja = "Poka¿"
                '====================================
                'OdswiezBaze()
                '====================================
                PobierzKlienci()
                Dim Form As New EdytujKatalogKlientow
                oAktualizuj = False
                Form.ShowDialog()

                If Not oAktualizuj Then
                Else
                    'Dim foundRow As DataRow
                    '' Create an array for the key values to find.
                    'Dim findTheseVals(0) As Object
                    'findTheseVals(0) = oIdentKli
                    'foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
                    '' sprawdz czy jest...
                    'If Not (foundRow Is Nothing) Then
                    '    'foundRow("NRKARTY") = oIdentKli
                    '    foundRow("NIP") = oNIPKli
                    '    foundRow("NRTOFITXT") = oNrTOFITxtKli
                    '    foundRow("KARTAPAYBACK") = oKartaPaybackKli
                    '    foundRow("NRYKARTPAYBACK") = oNryKartPaybackKli
                    '    foundRow("NAZWAKLIENTA") = oNazwa1Kli
                    '    foundRow("MIEJSCOWOSC") = oMiejscKli
                    '    foundRow("KODPOCZTOWY") = oKodPoczKli
                    '    foundRow("ULICA") = oUlicaKli
                    '    foundRow("NUMNRDOMU") = oNumNrDomuKli
                    '    foundRow("NRDOMU") = oNrDomuKli
                    '    foundRow("IDDOMU") = oIDDomuKli
                    '    foundRow("REJONKLIENTA") = oRejonKli
                    '    foundRow("TELEFON") = oTelefonKli
                    '    foundRow("FAX") = oFaxKli
                    '    foundRow("EMAIL") = oeMailKli
                    '    foundRow("RODZSPRZETUL") = oRodzSprzetuLKli
                    '    foundRow("RODZSPRZETUA") = oRodzSprzetuAKli
                    '    foundRow("RODZSPRZETUT") = oRodzSprzetuTKli
                    '    foundRow("OFERTATZLOZENIA") = oOfertaTZlozeniaKli
                    '    foundRow("OFERTAPLIK") = oOfertaPlikKli
                    '    foundRow("PKONTAKT") = oPKontaktKli
                    '    foundRow("PROMOTORUWAGI") = oSkupUwagikli
                    '    foundRow("OPIEKUN") = oSprzedOpiekunkli
                    '    foundRow("OPIEKUNUWAGI") = oSprzedUwagiKli
                    '    foundRow("KTOWPISAL") = oWpisalKli
                    '    foundRow("UWAGI") = oUwagikli
                    '    foundRow("MARKER") = oMarkerKli
                    '    foundRow("MARKERP") = oMarkerPKli
                    '    foundRow("WERSJA") = ZmienWersje(oWersjaKli) 'Wersja         Integer, 2

                    '    'wyswietl ilosc rekordow
                    '    stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
                    '    AktualizujBaze()
                    '    'aktualizuj DataGrid
                    '    dagKlienci.Update()
                    '    '---------------
                    '    dagKlienci_CurrentCellChanged(sender, e)
                    'End If
                End If
                '--------------------------
                Select Case oCoDalej
                    Case "STÓJ"
                        Exit Do

                    Case "NASTÊPNY"
                        If DataViewKlienci.Count() > 0 Then
                            If dagKlienci.CurrentCell.RowIndex < DataViewKlienci.Count() - 1 Then
                                dagKlienci.CurrentCell = dagKlienci(0, dagKlienci.CurrentCell.RowIndex + 1)
                                dagKlienci.CurrentCell.Selected = True
                            End If
                        End If
                    Case "POPRZEDNI"
                        If DataViewKlienci.Count() > 0 Then
                            If dagKlienci.CurrentCell.RowIndex > 0 Then
                                dagKlienci.CurrentCell = dagKlienci(0, dagKlienci.CurrentCell.RowIndex - 1)
                                dagKlienci.CurrentCell.Selected = True
                            End If
                        End If
                End Select
                '--------------------------
            Loop
        End If
    End Sub


    'Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
    '    If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
    '        'Raiseevent(Dodaj,Click )
    '        MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf & _
    '            "Proszê najpierw dopisaæ rekord do tabeli...", _
    '            "Uwaga :", _
    '            System.Windows.Forms.MessageBoxButtons.OK, _
    '            MessageBoxIcon.Information, _
    '            MessageBoxDefaultButton.Button1)
    '    Else

    '        If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis ?", _
    '                    "Proszê o potwierdzenie :", _
    '                    System.Windows.Forms.MessageBoxButtons.YesNo, _
    '                    MessageBoxIcon.Question, _
    '                    MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

    '            oOperacja = "USUN"
    '            PobierzKlienci()
    '            '====================================
    '            'OdswiezBaze()
    '            '====================================
    '            Dim foundRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(0) As Object
    '            findTheseVals(0) = oIdentKli
    '            foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                foundRow.Delete()
    '                stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
    '                AktualizujBaze()
    '                '---------------
    '                dagKlienci_CurrentCellChanged(sender, e)
    '                DataSetKlienci.AcceptChanges()
    '                '--------------------------------
    '                'usun zapisy z tablic podleglych...
    '                UsunOsoby(oIdentKli)
    '                UsunKontakty(oIdentKli)
    '                UsunObrotyMies(oIdentKli)
    '                UsunObroty(oIdentKli)
    '                UsunAkcjeSpec(oIdentKli)
    '            Else
    '                MessageBox.Show("Nie znalaz³em w katalogu zapisu dla którego " & vbCrLf & _
    '                    "NrKarty = " & oIdentKli & vbCrLf & _
    '                    "Ktoœ musia³ ju¿ ten zapis usun¹æ...", _
    '                    "Uwaga :", _
    '                    System.Windows.Forms.MessageBoxButtons.OK, _
    '                    MessageBoxIcon.Information, _
    '                    MessageBoxDefaultButton.Button1)
    '            End If
    '        End If
    '    End If
    'End Sub

    'Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
    '    oOperacja = "DODAJ"
    '    InitKlienci()
    '    '----------------------------------------------
    '    ''generuj IdentKlienta
    '    'Dim objRow As DataRow
    '    'Dim LastIdent As Long = 0
    '    'For Each objRow In DataSetKlienci.Tables(0).Rows
    '    '    If LastIdent < Val(objRow.Item(0)) Then
    '    '        LastIdent = Val(objRow.Item(0))
    '    '    End If
    '    'Next
    '    'oIdentKli = Mid(Trim(Str(1000000 + LastIdent + 1)), 2, 6)
    '    '----------------------------------------------
    '    Do While True
    '        oAktualizuj = False
    '        Dim Form As New EdytujKatalogKlientow
    '        Form.ShowDialog()
    '        Form.Dispose()
    '        If Not oAktualizuj Then
    '            Exit Do
    '        Else
    '            '====================================
    '            'OdswiezBaze()
    '            '====================================
    '            Dim foundRow As DataRow
    '            Dim NewRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(0) As Object
    '            findTheseVals(0) = oIdentKli
    '            foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf & _
    '                    "NrKarty = " & oIdentKli & vbCrLf & _
    '                    "Nie mogê dopisaæ tego rekordu...", _
    '                    "Uwaga :", _
    '                    System.Windows.Forms.MessageBoxButtons.OK, _
    '                    MessageBoxIcon.Information, _
    '                    MessageBoxDefaultButton.Button1)
    '                'Exit Do
    '            Else
    '                'stworz nowy rekord
    '                NewRow = DataSetKlienci.Tables(0).NewRow()

    '                NewRow("NRKARTY") = oIdentKli
    '                NewRow("NIP") = oNIPKli
    '                NewRow("NRTOFITXT") = oNrTOFITxtKli
    '                NewRow("KARTAPAYBACK") = oKartaPaybackKli
    '                NewRow("NRYKARTPAYBACK") = oNryKartPaybackKli
    '                NewRow("NAZWAKLIENTA") = oNazwa1Kli
    '                NewRow("MIEJSCOWOSC") = oMiejscKli
    '                NewRow("KODPOCZTOWY") = oKodPoczKli
    '                NewRow("ULICA") = oUlicaKli
    '                NewRow("NUMNRDOMU") = oNumNrDomuKli
    '                NewRow("NRDOMU") = oNrDomuKli
    '                NewRow("IDDOMU") = oIDDomuKli
    '                NewRow("REJONKLIENTA") = oRejonKli
    '                NewRow("TELEFON") = oTelefonKli
    '                NewRow("FAX") = oFaxKli
    '                NewRow("EMAIL") = oeMailKli
    '                NewRow("RODZSPRZETUL") = oRodzSprzetuLKli
    '                NewRow("RODZSPRZETUA") = oRodzSprzetuAKli
    '                NewRow("RODZSPRZETUT") = oRodzSprzetuTKli
    '                NewRow("OFERTATZLOZENIA") = oOfertaTZlozeniaKli
    '                NewRow("OFERTAPLIK") = oOfertaPlikKli
    '                NewRow("PKONTAKT") = oPKontaktKli
    '                NewRow("PROMOTORUWAGI") = oSkupUwagikli
    '                NewRow("OPIEKUN") = oSprzedOpiekunkli
    '                NewRow("OPIEKUNUWAGI") = oSprzedUwagiKli
    '                NewRow("KTOWPISAL") = oWpisalKli
    '                NewRow("UWAGI") = oUwagikli
    '                NewRow("MARKER") = False
    '                NewRow("MARKERP") = False
    '                NewRow("WERSJA") = 1                     'Wersja         Integer, 2
    '                'dodaj rekord do DataSet
    '                DataSetKlienci.Tables(0).Rows.Add(NewRow)

    '                'wyswietl ilosc rekordow
    '                stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
    '                AktualizujBaze()
    '                'aktualizuj DataGrid
    '                dagKlienci.Update()
    '                '---------------
    '                dagKlienci_CurrentCellChanged(sender, e)
    '                Exit Do
    '            End If
    '        End If
    '    Loop
    'End Sub


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
            '        dbDataAdapter.Update(DataSetKlienci, "TABELA_Klienci")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetKlienci)
            '    End If
            '    dbConnection.Close()
            'End If
        Else
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienci.Update(DataSetKlienci, "TABELA_Klienci")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                End If
                sqlConnectionKlienci.Close()
            End If
        End If
    End Sub



    Private Sub OdswiezBaze()
        DataSetKlienci.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetKlienci)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionKlienci.Open()
                If sqlConnectionKlienci.State = ConnectionState.Open Then
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub












    '****************************************
    '** obsluga Menu Kontekstowego
    '****************************************



    Private Sub menuEEdytuj_Click(sender As Object, e As EventArgs) Handles menuEEdytuj.Click
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuEDodaj_Click(sender As Object, e As EventArgs) Handles menuEDodaj.Click
        '    cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuEUsun_Click(sender As Object, e As EventArgs) Handles menuEUsun.Click
        '    cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuESingleL_Click(sender As Object, e As EventArgs) Handles menuESingleL.Click
        dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagKlienci.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagKlienci.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub

End Class
