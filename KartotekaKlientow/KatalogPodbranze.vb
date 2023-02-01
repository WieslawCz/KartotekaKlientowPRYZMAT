Imports System.Drawing.Printing

Public Class KatalogPodBranze
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
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents stbPodBranze As System.Windows.Forms.StatusBar
    Friend WithEvents cmdWycofajsie As System.Windows.Forms.Button
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    Friend WithEvents cmdFi As System.Windows.Forms.Button
    Friend WithEvents cmdHarmonogramWizyt As System.Windows.Forms.Button
    Friend WithEvents dagPodBranze As DataGridView
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
    Friend WithEvents cmdPobierzZPliku As Button
    Friend WithEvents stbFiltry As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogPodBranze))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbPodBranze = New System.Windows.Forms.StatusBar()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdWycofajsie = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdHarmonogramWizyt = New System.Windows.Forms.Button()
        Me.dagPodBranze = New System.Windows.Forms.DataGridView()
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
        Me.cmdPobierzZPliku = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagPodBranze, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuWybieraj.SuspendLayout()
        Me.menuEdytuj.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(690, 333)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 23)
        Me.cmdStop.TabIndex = 37
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'stbPodBranze
        '
        Me.stbPodBranze.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbPodBranze.Dock = System.Windows.Forms.DockStyle.None
        Me.stbPodBranze.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbPodBranze.Location = New System.Drawing.Point(8, 303)
        Me.stbPodBranze.Name = "stbPodBranze"
        Me.stbPodBranze.ShowPanels = True
        Me.stbPodBranze.Size = New System.Drawing.Size(762, 22)
        Me.stbPodBranze.TabIndex = 42
        Me.stbPodBranze.Text = "StatusBar1"
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(490, 333)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytuj.TabIndex = 41
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.Image = CType(resources.GetObject("cmdUsuñ.Image"), System.Drawing.Image)
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(402, 333)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 23)
        Me.cmdUsuñ.TabIndex = 40
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(314, 333)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodaj.TabIndex = 39
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(16, 333)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 38
        Me.PictureBox1.TabStop = False
        '
        'cmdWycofajsie
        '
        Me.cmdWycofajsie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajsie.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWycofajsie.Image = CType(resources.GetObject("cmdWycofajsie.Image"), System.Drawing.Image)
        Me.cmdWycofajsie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajsie.Location = New System.Drawing.Point(578, 333)
        Me.cmdWycofajsie.Name = "cmdWycofajsie"
        Me.cmdWycofajsie.Size = New System.Drawing.Size(104, 23)
        Me.cmdWycofajsie.TabIndex = 45
        Me.cmdWycofajsie.Text = "Wycofaj siê"
        Me.cmdWycofajsie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(552, 7)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 20)
        Me.cmdWszystko.TabIndex = 66
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 8)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(770, 22)
        Me.stbFiltry.TabIndex = 85
        Me.stbFiltry.Text = "StatusBar1"
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(528, 8)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 175
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdHarmonogramWizyt
        '
        Me.cmdHarmonogramWizyt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdHarmonogramWizyt.ForeColor = System.Drawing.Color.Black
        Me.cmdHarmonogramWizyt.Image = CType(resources.GetObject("cmdHarmonogramWizyt.Image"), System.Drawing.Image)
        Me.cmdHarmonogramWizyt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdHarmonogramWizyt.Location = New System.Drawing.Point(126, 332)
        Me.cmdHarmonogramWizyt.Name = "cmdHarmonogramWizyt"
        Me.cmdHarmonogramWizyt.Size = New System.Drawing.Size(24, 23)
        Me.cmdHarmonogramWizyt.TabIndex = 205
        Me.cmdHarmonogramWizyt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dagPodBranze
        '
        Me.dagPodBranze.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagPodBranze.Location = New System.Drawing.Point(8, 32)
        Me.dagPodBranze.Name = "dagPodBranze"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagPodBranze.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagPodBranze.Size = New System.Drawing.Size(770, 269)
        Me.dagPodBranze.TabIndex = 206
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
        'cmdPobierzZPliku
        '
        Me.cmdPobierzZPliku.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdPobierzZPliku.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPobierzZPliku.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPobierzZPliku.Location = New System.Drawing.Point(156, 333)
        Me.cmdPobierzZPliku.Name = "cmdPobierzZPliku"
        Me.cmdPobierzZPliku.Size = New System.Drawing.Size(86, 23)
        Me.cmdPobierzZPliku.TabIndex = 207
        Me.cmdPobierzZPliku.Text = "Pobierz z pliku"
        '
        'KatalogPodBranze
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 361)
        Me.Controls.Add(Me.cmdPobierzZPliku)
        Me.Controls.Add(Me.dagPodBranze)
        Me.Controls.Add(Me.cmdHarmonogramWizyt)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.cmdWycofajsie)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.stbPodBranze)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "KatalogPodBranze"
        Me.Text = "Podran¿e"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagPodBranze, System.ComponentModel.ISupportInitialize).EndInit()
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

    Dim DataSetPodBranze As New DataSet
    Dim DataViewPodBranze As New DataView
    Dim ConnectionState As System.Data.ConnectionState


    '---------------------------------------------------------------------
    'PodBranze
    '---------------------------------------------------------------------



    Private Sub KatalogPodBranze_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        DataSetPodBranze.Locale = New System.Globalization.CultureInfo("pl-PL")

        ''PodPodBranze
        'Public pbrIdBranzy As String            'IDBRANZY         Text, 
        'Public pbrIdPodBranzy As String         'IDPODBRANZY         Text, 
        'Public pbrWersja As Integer             'WERSJA           Integer

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionPodBranze)
            'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLPodBranze, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLPodBranze, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLPodBranze, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLPodBranze, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
            'DBMapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_PodBranze")

            'DBDeletePodBranze(dbCommandDelete, dbDataAdapter)
            'DBUpdatePodBranze(dbCommandUpdate, dbDataAdapter)
            'DBInsertPodBranze(dbCommandInsert, dbDataAdapter)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionPodBranze)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQLPodBranze & IIf(Len(oIdBranzy) > 0, " WHERE IDBRANZY='" & oIdBranzy & "' ", ""), sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLPodBranze, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLPodBranze, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLPodBranze, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_PodBranze")

            SQLDeletePodBranze(sqlCommandDelete, sqlDataAdapter)
            SQLUpdatePodBranze(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertPodBranze(sqlCommandInsert, sqlDataAdapter)

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
                    dbDataAdapter.FillSchema(DataSetPodBranze, SchemaType.Mapped, "TABELA_PodBranze")
                    dbDataAdapter.Fill(DataSetPodBranze)
                    dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetPodBranze, SchemaType.Mapped, "TABELA_PodBranze")
                    sqlDataAdapter.Fill(DataSetPodBranze)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetPodBranze.Tables(0).PrimaryKey = New DataColumn() {DataSetPodBranze.Tables(0).Columns("IDBRANZY"),
                                                                          DataSetPodBranze.Tables(0).Columns("IDPODBRANZY")}

                'definiuj DataView
                DataViewPodBranze = New DataView(DataSetPodBranze.Tables(0))
                DataViewPodBranze.AllowDelete = False
                DataViewPodBranze.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagPodBranze.BackgroundColor = System.Drawing.SystemColors.Control
                dagPodBranze.GridColor = System.Drawing.SystemColors.ControlDark
                dagPodBranze.ForeColor = System.Drawing.Color.Purple
                dagPodBranze.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagPodBranze.BorderStyle = BorderStyle.Fixed3D
                'dagPodBranze.Dock = DockStyle.Fill
                dagPodBranze.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagPodBranze.AllowUserToAddRows = False
                dagPodBranze.AllowUserToDeleteRows = False
                dagPodBranze.AllowUserToOrderColumns = True
                dagPodBranze.AllowUserToResizeColumns = True
                dagPodBranze.AllowUserToResizeRows = True
                dagPodBranze.ReadOnly = True
                dagPodBranze.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagPodBranze.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagPodBranze.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagPodBranze.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagPodBranze.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagPodBranze.ColumnHeadersVisible = True
                dagPodBranze.ColumnHeadersHeight = 40
                dagPodBranze.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagPodBranze.RowHeadersVisible = True
                dagPodBranze.RowHeadersWidth = 50
                dagPodBranze.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagPodBranze.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagPodBranze.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagPodBranze.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagPodBranze.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagPodBranze.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagPodBranze.DefaultCellStyle.NullValue = ""
                dagPodBranze.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagPodBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagPodBranze.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagPodBranze.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagPodBranze.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagPodBranze.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagPodBranze.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagPodBranze.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagPodBranze.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagPodBranze.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagPodBranze.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagPodBranze.DataMember = DataSetPodBranze.Tables(0).TableName
                'dagPodBranze.DataSource = DataViewPodBranze
                '-----------------------------------

                ' Add a GridColumnStyle and set the MappingName 
                ' to the name of a DataColumn in the DataTable. 
                ' Set the HeaderText and Width properties. 

                ''PodPodBranze
                'Public pbrIdBranzy As String            'IDBRANZY         Text, 
                'Public pbrIdPodBranzy As String         'IDPODBRANZY         Text, 
                'Public pbrWersja As Integer             'WERSJA           Integer

                TxtColumnView(dagPodBranze, DataSetPodBranze.Tables(0).Columns(0).ColumnName, "Bran¿a", 400, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagPodBranze, DataSetPodBranze.Tables(0).Columns(1).ColumnName, "Podbran¿a", 400, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(dagPodBranze, DataSetPodBranze.Tables(0).Columns(2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N0", False)

                Me.Cursor = Cursors.WaitCursor
                dagPodBranze.DataSource = DataViewPodBranze
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetPodBranze.Tables(0).Rows.Count > 0 Then
                    dagPodBranze.CurrentCell = dagPodBranze(0, 0)
                    dagPodBranze.CurrentCell.Selected = True
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
            stbPodBranze.BackColor = System.Drawing.SystemColors.Control
            stbPodBranze.ForeColor = System.Drawing.Color.DarkGreen
            stbPodBranze.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbPodBranze.Panels.Add("stbPanelIleRec")
            stbPodBranze.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbPodBranze.Panels(0).Width = 80
            stbPodBranze.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbPodBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetPodBranze.Tables(0).Rows.Count.ToString

            stbPodBranze.Panels.Add("stbPanelWiersz")
            stbPodBranze.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbPodBranze.Panels(1).Width = 80
            stbPodBranze.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagPodBranze.CurrentCell Is Nothing Then
                stbPodBranze.Panels(1).Text = "Wiersz : "
            Else
                stbPodBranze.Panels(1).Text = "Wiersz : " & dagPodBranze.CurrentCell.RowIndex.ToString
            End If

            stbPodBranze.Panels.Add("stbPanelKolumna")
            stbPodBranze.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbPodBranze.Panels(2).Width = 80
            stbPodBranze.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagPodBranze.CurrentCell Is Nothing Then
                stbPodBranze.Panels(2).Text = "Kolumna : "
            Else
                stbPodBranze.Panels(2).Text = "Kolumna : " & dagPodBranze.CurrentCell.ColumnIndex.ToString
            End If

            stbPodBranze.Panels.Add("stbPanelSort")
            stbPodBranze.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbPodBranze.Panels(3).Width = 150
            stbPodBranze.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbPodBranze.Panels(3).Text = "Sort : "

            stbPodBranze.Panels.Add("stbPanelFiltr")
            stbPodBranze.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbPodBranze.Panels(4).Width = 150
            stbPodBranze.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbPodBranze.Panels(4).Text = "Filtr : "

            stbPodBranze.Panels.Add("stbPanelReszta")
            stbPodBranze.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbPodBranze.Panels(5).Width = 150
            stbPodBranze.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbPodBranze.Panels(5).Text = ""

            stbPodBranze.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagPodBranze.TableStyles(0).RowHeaderWidth
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
        InitPodBranze() 'inicjuj zmienne
        '-----------------------------------------------------------------
        If _OCoMamRobic = "WYBIERAJ" Then
            dagPodBranze.ContextMenuStrip = Me.menuWybieraj
            cmdStop.Text = "Wybierz"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = True
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = False
        Else
            dagPodBranze.ContextMenuStrip = Me.menuEdytuj
            cmdStop.Text = "Powrót"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = False
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = True
        End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagPodBranze.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagPodBranze.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetPodBranze.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagPodBranze.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagPodBranze.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagPodBranze.FirstDisplayedScrollingColumnIndex +
                        dagPodBranze.GetCellDisplayRectangle(dagPodBranze.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagPodBranze.Left + 4), (dagPodBranze.Top + 4))
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
        For ii = 0 To DataViewPodBranze.Table.Columns.Count - 1
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
        dagPodBranze_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub


    Private Sub KatalogPodBranze_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub




















    Private Sub dagPodBranze_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagPodBranze.CurrentCellChanged
        If Not StartFormy Then
            If dagPodBranze.CurrentCell Is Nothing Then
                stbPodBranze.Panels(1).Text = "Wi: "
                stbPodBranze.Panels(2).Text = "Ko: "
            Else
                stbPodBranze.Panels(1).Text = "Wi: " & dagPodBranze.CurrentCell.RowIndex.ToString
                stbPodBranze.Panels(2).Text = "Ko: " & dagPodBranze.CurrentCell.ColumnIndex.ToString
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
                If dagPodBranze.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagPodBranze.FirstDisplayedScrollingColumnIndex +
                                    dagPodBranze.GetCellDisplayRectangle(dagPodBranze.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagPodBranze.FirstDisplayedScrollingColumnIndex +
                                    dagPodBranze.GetCellDisplayRectangle(dagPodBranze.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagPodBranze.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagPodBranze.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagPodBranze, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, DataViewPodBranze, stbPodBranze)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagPodBranze, Mid(sender.name, 6, 3), "PodBranze")
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
            CzyscFiltryDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, DataViewPodBranze, stbPodBranze)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbPodBranze.Panels(0).Text = "Il.zap.: " & DataViewPodBranze.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagPodBranze_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagPodBranze.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagPodBranze.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagPodBranze.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagPodBranze_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagPodBranze.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagPodBranze, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagPodBranze, e.RowIndex, 4)

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

    Private Sub dagPodBranze_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagPodBranze.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagPodBranze_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagPodBranze.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagPodBranze_Scroll(sender As Object, e As ScrollEventArgs) Handles dagPodBranze.Scroll
        'If (Not StartFormy) And (DataViewPodBranze.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewPodBranze.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagPodBranze_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagPodBranze.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagPodBranze_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagPodBranze.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagPodBranze, DataViewPodBranze.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagPodBranze_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagPodBranze.MouseUp
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
                stbPodBranze.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbPodBranze.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagPodBranze(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbPodBranze.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbPodBranze.Panels(3).Text = "Sort: " & DataSetPodBranze.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbPodBranze.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagPodBranze(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbPodBranze.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbPodBranze.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub


    Private Sub dagPodBranze_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagPodBranze.DoubleClick
        If _OCoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagPodBranze_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagPodBranze.KeyDown
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
    ''PodPodBranze
    'Public pbrIdBranzy As String            'IDBRANZY         Text, 
    'Public pbrIdPodBranzy As String         'IDPODBRANZY         Text, 
    'Public pbrWersja As Integer             'WERSJA           Integer


    Private Sub InitPodBranze()
        pbrIdBranzy = ""
        pbrIdPodBranzy = ""
        pbrWersja = 1
    End Sub

    Private Sub PobierzPodBranze()
        'pobierz wartosci przed aktualizacja
        pbrIdBranzy = GetTxtField(dagPodBranze, 0)
        pbrIdPodBranzy = GetTxtField(dagPodBranze, 1)
        pbrWersja = GetIntField(dagPodBranze, 2)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        If DataSetPodBranze.Tables(0).Rows.Count <= 0 Then
            oIdBranzy = ""
            oIdPodBranzy = ""
        Else
            oIdBranzy = GetTxtField(dagPodBranze, 0)
            oIdPodBranzy = GetTxtField(dagPodBranze, 1)
        End If
        oAktualizuj = True
        Me.Close()
    End Sub

    Private Sub cmdWycofajsie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajsie.Click
        oIdBranzy = ""
        oIdPodBranzy = ""
        oAktualizuj = False
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetPodBranze.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            'Public brIdBranzy As String            'IDBRANZY         Text, 
            'Public brWersja As Integer             'WERSJA           Integer

            oOperacja = "EDYTUJ"
            PobierzPodBranze()
            Dim Form As New EdytujKatalogPodbranze
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = pbrIdBranzy
                findTheseVals(1) = pbrIdPodBranzy
                foundRow = DataSetPodBranze.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDBRANZY") = pbrIDBRANZY
                    'foundRow("IDPODBRANZY") = pbrIDPODBRANZY
                    foundRow("WERSJA") = ZmienWersje(pbrWersja)    'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbPodBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetPodBranze.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagPodBranze.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetPodBranze.Tables(0).Rows.Count <= 0 Then
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
                        MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                oOperacja = "USUN"
                PobierzPodBranze()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = pbrIdBranzy
                findTheseVals(1) = pbrIdPodBranzy
                foundRow = DataSetPodBranze.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbPodBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetPodBranze.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitPodBranze()
        Dim Form As New EdytujKatalogPodbranze
        Do While True
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = pbrIdBranzy
                findTheseVals(1) = pbrIdPodBranzy
                foundRow = DataSetPodBranze.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Bran¿a = " & pbrIdBranzy & vbCrLf &
                        "Podbran¿a = " & pbrIdPodBranzy,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetPodBranze.Tables(0).NewRow()

                    NewRow("IDBRANZY") = pbrIdBranzy
                    NewRow("IDPODBRANZY") = pbrIdPodBranzy
                    NewRow("WERSJA") = 1                     'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetPodBranze.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbPodBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetPodBranze.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagPodBranze.Update()
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
                    dbDataAdapter.Update(DataSetPodBranze, "TABELA_PodBranze")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapter.Fill(DataSetPodBranze)
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
                    sqlDataAdapter.Update(DataSetPodBranze, "TABELA_PodBranze")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetPodBranze)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub


    Private Sub OdswiezBaze()
        DataSetPodBranze.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnection.Open()
                If dbConnection.State = ConnectionState.Open Then
                    dbDataAdapter.Fill(DataSetPodBranze)
                    dbConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetPodBranze)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub




    Private Sub cmdHarmonogramWizyt_Click(sender As Object, e As EventArgs) Handles cmdHarmonogramWizyt.Click
        Dim Form As New KatalogHarmonogramWizyt
        Form.Show()
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
        dagPodBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagPodBranze.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        dagPodBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagPodBranze.Refresh()
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
        dagPodBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagPodBranze.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagPodBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagPodBranze.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub





    Private Sub cmdPobierzZPliku_Click(sender As Object, e As EventArgs) Handles cmdPobierzZPliku.Click
        Dim Form As New frmPobierzZPlikuBranze("PODBRANZE", DataSetPodBranze)
        Form.ShowDialog()
        '-----------
        AktualizujBaze()
        'aktualizuj DataGrid
        dagPodBranze.Update()
        stbPodBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetPodBranze.Tables(0).Rows.Count.ToString
    End Sub


End Class
