Imports System.Drawing.Printing

Public Class KatalogBranze
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
    Friend WithEvents stbBranze As System.Windows.Forms.StatusBar
    Friend WithEvents cmdWycofajsie As System.Windows.Forms.Button
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    Friend WithEvents cmdFi As System.Windows.Forms.Button
    Friend WithEvents cmdHarmonogramWizyt As System.Windows.Forms.Button
    Friend WithEvents dagBranze As DataGridView
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogBranze))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbBranze = New System.Windows.Forms.StatusBar()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdWycofajsie = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdHarmonogramWizyt = New System.Windows.Forms.Button()
        Me.dagBranze = New System.Windows.Forms.DataGridView()
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
        CType(Me.dagBranze, System.ComponentModel.ISupportInitialize).BeginInit()
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
        'stbBranze
        '
        Me.stbBranze.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbBranze.Dock = System.Windows.Forms.DockStyle.None
        Me.stbBranze.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbBranze.Location = New System.Drawing.Point(8, 303)
        Me.stbBranze.Name = "stbBranze"
        Me.stbBranze.ShowPanels = True
        Me.stbBranze.Size = New System.Drawing.Size(762, 22)
        Me.stbBranze.TabIndex = 42
        Me.stbBranze.Text = "StatusBar1"
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
        'dagBranze
        '
        Me.dagBranze.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagBranze.Location = New System.Drawing.Point(8, 32)
        Me.dagBranze.Name = "dagBranze"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagBranze.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagBranze.Size = New System.Drawing.Size(770, 269)
        Me.dagBranze.TabIndex = 206
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
        'KatalogBranze
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 361)
        Me.Controls.Add(Me.cmdPobierzZPliku)
        Me.Controls.Add(Me.dagBranze)
        Me.Controls.Add(Me.cmdHarmonogramWizyt)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.cmdWycofajsie)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.stbBranze)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "KatalogBranze"
        Me.Text = "Bran¿e"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagBranze, System.ComponentModel.ISupportInitialize).EndInit()
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

    Dim DataSetBranze As New DataSet
    Dim DataViewBranze As New DataView
    Dim ConnectionState As System.Data.ConnectionState


    '---------------------------------------------------------------------
    'Branze
    '---------------------------------------------------------------------



    Private Sub KatalogBranze_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        DataSetBranze.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionBranze)
            'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLBranze, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLBranze, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLBranze, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLBranze, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
            'DBMapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Branze")

            'DBDeleteBranze(dbCommandDelete, dbDataAdapter)
            'DBUpdateBranze(dbCommandUpdate, dbDataAdapter)
            'DBInsertBranze(dbCommandInsert, dbDataAdapter)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionBranze)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQLBranze, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLBranze, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLBranze, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLBranze, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Branze")

            SQLDeleteBranze(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateBranze(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertBranze(sqlCommandInsert, sqlDataAdapter)

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
                    dbDataAdapter.FillSchema(DataSetBranze, SchemaType.Mapped, "TABELA_Branze")
                    dbDataAdapter.Fill(DataSetBranze)
                    dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetBranze, SchemaType.Mapped, "TABELA_Branze")
                    sqlDataAdapter.Fill(DataSetBranze)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetBranze.Tables(0).PrimaryKey = New DataColumn() {DataSetBranze.Tables(0).Columns("IDBRANZY")}

                'definiuj DataView
                DataViewBranze = New DataView(DataSetBranze.Tables(0))
                DataViewBranze.AllowDelete = False
                DataViewBranze.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagBranze.BackgroundColor = System.Drawing.SystemColors.Control
                dagBranze.GridColor = System.Drawing.SystemColors.ControlDark
                dagBranze.ForeColor = System.Drawing.Color.Purple
                dagBranze.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagBranze.BorderStyle = BorderStyle.Fixed3D
                'dagBranze.Dock = DockStyle.Fill
                dagBranze.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagBranze.AllowUserToAddRows = False
                dagBranze.AllowUserToDeleteRows = False
                dagBranze.AllowUserToOrderColumns = True
                dagBranze.AllowUserToResizeColumns = True
                dagBranze.AllowUserToResizeRows = True
                dagBranze.ReadOnly = True
                dagBranze.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagBranze.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagBranze.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagBranze.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagBranze.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagBranze.ColumnHeadersVisible = True
                dagBranze.ColumnHeadersHeight = 40
                dagBranze.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagBranze.RowHeadersVisible = True
                dagBranze.RowHeadersWidth = 50
                dagBranze.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagBranze.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagBranze.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagBranze.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagBranze.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagBranze.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagBranze.DefaultCellStyle.NullValue = ""
                dagBranze.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagBranze.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagBranze.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagBranze.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagBranze.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagBranze.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagBranze.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagBranze.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagBranze.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagBranze.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagBranze.DataMember = DataSetBranze.Tables(0).TableName
                'dagBranze.DataSource = DataViewBranze
                '-----------------------------------

                ' Add a GridColumnStyle and set the MappingName 
                ' to the name of a DataColumn in the DataTable. 
                ' Set the HeaderText and Width properties. 

                'Public brIdBranzy As String            'IDBRANZY         Text, 
                'Public brWersja As Integer             'WERSJA           Integer

                TxtColumnView(dagBranze, DataSetBranze.Tables(0).Columns(0).ColumnName, "Bran¿a", 400, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(dagBranze, DataSetBranze.Tables(0).Columns(1).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N0", False)

                Me.Cursor = Cursors.WaitCursor
                dagBranze.DataSource = DataViewBranze
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetBranze.Tables(0).Rows.Count > 0 Then
                    dagBranze.CurrentCell = dagBranze(0, 0)
                    dagBranze.CurrentCell.Selected = True
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
            stbBranze.BackColor = System.Drawing.SystemColors.Control
            stbBranze.ForeColor = System.Drawing.Color.DarkGreen
            stbBranze.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbBranze.Panels.Add("stbPanelIleRec")
            stbBranze.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbBranze.Panels(0).Width = 80
            stbBranze.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetBranze.Tables(0).Rows.Count.ToString

            stbBranze.Panels.Add("stbPanelWiersz")
            stbBranze.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbBranze.Panels(1).Width = 80
            stbBranze.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagBranze.CurrentCell Is Nothing Then
                stbBranze.Panels(1).Text = "Wiersz : "
            Else
                stbBranze.Panels(1).Text = "Wiersz : " & dagBranze.CurrentCell.RowIndex.ToString
            End If

            stbBranze.Panels.Add("stbPanelKolumna")
            stbBranze.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbBranze.Panels(2).Width = 80
            stbBranze.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagBranze.CurrentCell Is Nothing Then
                stbBranze.Panels(2).Text = "Kolumna : "
            Else
                stbBranze.Panels(2).Text = "Kolumna : " & dagBranze.CurrentCell.ColumnIndex.ToString
            End If

            stbBranze.Panels.Add("stbPanelSort")
            stbBranze.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbBranze.Panels(3).Width = 150
            stbBranze.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbBranze.Panels(3).Text = "Sort : "

            stbBranze.Panels.Add("stbPanelFiltr")
            stbBranze.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbBranze.Panels(4).Width = 150
            stbBranze.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbBranze.Panels(4).Text = "Filtr : "

            stbBranze.Panels.Add("stbPanelReszta")
            stbBranze.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbBranze.Panels(5).Width = 150
            stbBranze.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbBranze.Panels(5).Text = ""

            stbBranze.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagBranze.TableStyles(0).RowHeaderWidth
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
        InitBranze() 'inicjuj zmienne
        '-----------------------------------------------------------------
        If _OCoMamRobic = "WYBIERAJ" Then
            dagBranze.ContextMenuStrip = Me.menuWybieraj
            cmdStop.Text = "Wybierz"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = True
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = False
        Else
            dagBranze.ContextMenuStrip = Me.menuEdytuj
            cmdStop.Text = "Powrót"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = False
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = True
        End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagBranze.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagBranze.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetBranze.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagBranze.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagBranze.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagBranze.FirstDisplayedScrollingColumnIndex +
                        dagBranze.GetCellDisplayRectangle(dagBranze.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagBranze.Left + 4), (dagBranze.Top + 4))
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
        For ii = 0 To DataViewBranze.Table.Columns.Count - 1
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
        dagBranze_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub


    Private Sub KatalogBranze_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub




















    Private Sub dagBranze_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagBranze.CurrentCellChanged
        If Not StartFormy Then
            If dagBranze.CurrentCell Is Nothing Then
                stbBranze.Panels(1).Text = "Wi: "
                stbBranze.Panels(2).Text = "Ko: "
            Else
                stbBranze.Panels(1).Text = "Wi: " & dagBranze.CurrentCell.RowIndex.ToString
                stbBranze.Panels(2).Text = "Ko: " & dagBranze.CurrentCell.ColumnIndex.ToString
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
                If dagBranze.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagBranze.FirstDisplayedScrollingColumnIndex +
                                    dagBranze.GetCellDisplayRectangle(dagBranze.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagBranze.FirstDisplayedScrollingColumnIndex +
                                    dagBranze.GetCellDisplayRectangle(dagBranze.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagBranze.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagBranze.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagBranze, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, DataViewBranze, stbBranze)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagBranze, Mid(sender.name, 6, 3), "Branze")
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
            CzyscFiltryDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, DataViewBranze, stbBranze)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbBranze.Panels(0).Text = "Il.zap.: " & DataViewBranze.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagBranze_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagBranze.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagBranze.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagBranze.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagBranze_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagBranze.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagBranze, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagBranze, e.RowIndex, 4)

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

    Private Sub dagBranze_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagBranze.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagBranze_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagBranze.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagBranze_Scroll(sender As Object, e As ScrollEventArgs) Handles dagBranze.Scroll
        'If (Not StartFormy) And (DataViewBranze.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewBranze.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagBranze_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagBranze.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagBranze_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagBranze.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagBranze, DataViewBranze.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagBranze_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagBranze.MouseUp
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
                stbBranze.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbBranze.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagBranze(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbBranze.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbBranze.Panels(3).Text = "Sort: " & DataSetBranze.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbBranze.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagBranze(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbBranze.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbBranze.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub


    Private Sub dagBranze_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagBranze.DoubleClick
        If _OCoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagBranze_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagBranze.KeyDown
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
    'Public brIdBranzy As String            'IDBRANZY         Text, 
    'Public brWersja As Integer             'WERSJA           Integer


    Private Sub InitBranze()
        brIdBranzy = ""
        brWersja = 1
    End Sub

    Private Sub PobierzBranze()
        'pobierz wartosci przed aktualizacja
        brIdBranzy = GetTxtField(dagBranze, 0)
        brWersja = GetIntField(dagBranze, 1)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        If DataSetBranze.Tables(0).Rows.Count <= 0 Then
            oIdBranzy = ""
        Else
            oIdBranzy = GetTxtField(dagBranze, 0)
        End If
        oAktualizuj = True
        Me.Close()
    End Sub

    Private Sub cmdWycofajsie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajsie.Click
        oIdBranzy = ""
        oAktualizuj = False
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetBranze.Tables(0).Rows.Count <= 0 Then
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
            PobierzBranze()
            Dim Form As New EdytujKatalogBranze
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = brIdBranzy
                foundRow = DataSetBranze.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDBRANZY") = brIDBRANZY
                    foundRow("WERSJA") = ZmienWersje(brWersja)    'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetBranze.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagBranze.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetBranze.Tables(0).Rows.Count <= 0 Then
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
                PobierzBranze()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = brIdBranzy
                foundRow = DataSetBranze.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetBranze.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitBranze()
        Dim Form As New EdytujKatalogBranze
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
                findTheseVals(0) = brIdBranzy
                foundRow = DataSetBranze.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Bran¿a = " & brIdBranzy,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetBranze.Tables(0).NewRow()

                    NewRow("IDBRANZY") = brIdBranzy
                    NewRow("WERSJA") = 1                     'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetBranze.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetBranze.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagBranze.Update()
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
                    dbDataAdapter.Update(DataSetBranze, "TABELA_Branze")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapter.Fill(DataSetBranze)
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
                    sqlDataAdapter.Update(DataSetBranze, "TABELA_Branze")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetBranze)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub


    Private Sub OdswiezBaze()
        DataSetBranze.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnection.Open()
                If dbConnection.State = ConnectionState.Open Then
                    dbDataAdapter.Fill(DataSetBranze)
                    dbConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetBranze)
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
        dagBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagBranze.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        dagBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagBranze.Refresh()
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
        dagBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagBranze.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagBranze.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagBranze.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub





    Private Sub cmdPobierzZPliku_Click(sender As Object, e As EventArgs) Handles cmdPobierzZPliku.Click
        Dim Form As New frmPobierzZPlikuBranze("BRANZE", DataSetBranze)
        Form.ShowDialog()
        '-----------
        AktualizujBaze()
        'aktualizuj DataGrid
        dagBranze.Update()
        stbBranze.Panels(0).Text = "Iloœæ rec.: " & DataSetBranze.Tables(0).Rows.Count.ToString
    End Sub
End Class
