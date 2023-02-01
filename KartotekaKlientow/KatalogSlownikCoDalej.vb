Imports System.Drawing.Printing

Public Class KatalogSlownikCoDalej
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
    Friend WithEvents stbSlownikCoDalej As System.Windows.Forms.StatusBar
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
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
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystko As Button
    Friend WithEvents stbFiltry As StatusBar
    Friend WithEvents dagSlownikCoDalej As DataGridView
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogSlownikCoDalej))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbSlownikCoDalej = New System.Windows.Forms.StatusBar()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
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
        Me.dagSlownikCoDalej = New System.Windows.Forms.DataGridView()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuWybieraj.SuspendLayout()
        Me.menuEdytuj.SuspendLayout()
        CType(Me.dagSlownikCoDalej, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(689, 329)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(89, 22)
        Me.cmdStop.TabIndex = 37
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'stbSlownikCoDalej
        '
        Me.stbSlownikCoDalej.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbSlownikCoDalej.Dock = System.Windows.Forms.DockStyle.None
        Me.stbSlownikCoDalej.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbSlownikCoDalej.Location = New System.Drawing.Point(8, 300)
        Me.stbSlownikCoDalej.Name = "stbSlownikCoDalej"
        Me.stbSlownikCoDalej.ShowPanels = True
        Me.stbSlownikCoDalej.Size = New System.Drawing.Size(770, 23)
        Me.stbSlownikCoDalej.TabIndex = 42
        Me.stbSlownikCoDalej.Text = "StatusBar1"
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(604, 329)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 22)
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
        Me.cmdUsuñ.Location = New System.Drawing.Point(518, 329)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 22)
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
        Me.cmdDodaj.Location = New System.Drawing.Point(432, 329)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 22)
        Me.cmdDodaj.TabIndex = 39
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 329)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 23)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 38
        Me.PictureBox1.TabStop = False
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
        Me.cmdFi.Location = New System.Drawing.Point(546, 14)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 187
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(603, 12)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystko.TabIndex = 185
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
        Me.stbFiltry.TabIndex = 186
        Me.stbFiltry.Text = "StatusBar1"
        '
        'dagSlownikCoDalej
        '
        Me.dagSlownikCoDalej.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagSlownikCoDalej.Location = New System.Drawing.Point(8, 30)
        Me.dagSlownikCoDalej.Name = "dagSlownikCoDalej"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagSlownikCoDalej.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagSlownikCoDalej.Size = New System.Drawing.Size(770, 269)
        Me.dagSlownikCoDalej.TabIndex = 184
        '
        'KatalogSlownikCoDalej
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 361)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.dagSlownikCoDalej)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.stbSlownikCoDalej)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "KatalogSlownikCoDalej"
        Me.ShowInTaskbar = False
        Me.Text = "S³ownik Co Dalej"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuWybieraj.ResumeLayout(False)
        Me.menuEdytuj.ResumeLayout(False)
        CType(Me.dagSlownikCoDalej, System.ComponentModel.ISupportInitialize).EndInit()
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


    Dim DataSetSlownikCoDalej As New DataSet
    Dim DataViewSlownikCoDalej As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim _CoMamRobic As String = ""

    '---------------------------------------------------------------------
    'SlownikCoDalej
    'Public scdIDENT As String             'IDENT       Text, 40
    'Public scdOPIS As String              'OPIS        memo
    'Public scdWersja As Integer           'WERSJA      Integer

    Private Sub KatalogSlownikCoDalej_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        _CoMamRobic = oCoMamRobic
        '--------------------------------
        DataSetSlownikCoDalej.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionSlownikCoDalej)
            'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLSlownikCoDalej, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLSlownikCoDalej, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLSlownikCoDalej, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLSlownikCoDalej, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_SlownikCoDalej")

            'DBDeleteSlownikCoDalej(dbCommandDelete, dbDataAdapter)
            'DBUpdateSlownikCoDalej(dbCommandUpdate, dbDataAdapter)
            'DBInsertSlownikCoDalej(dbCommandInsert, dbDataAdapter)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionSlownikCoDalej)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQLSlownikCoDalej, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLSlownikCoDalej, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLSlownikCoDalej, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLSlownikCoDalej, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_SlownikCoDalej")

            SQLDeleteSlownikCoDalej(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateSlownikCoDalej(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertSlownikCoDalej(sqlCommandInsert, sqlDataAdapter)

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
                    dbDataAdapter.FillSchema(DataSetSlownikCoDalej, SchemaType.Mapped, "TABELA_SlownikCoDalej")
                    dbDataAdapter.Fill(DataSetSlownikCoDalej)
                    dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetSlownikCoDalej, SchemaType.Mapped, "TABELA_SlownikCoDalej")
                    sqlDataAdapter.Fill(DataSetSlownikCoDalej)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetSlownikCoDalej.Tables(0).PrimaryKey = New DataColumn() {DataSetSlownikCoDalej.Tables(0).Columns("IDENT")}

                'definiuj DataView
                DataViewSlownikCoDalej = New DataView(DataSetSlownikCoDalej.Tables(0))
                DataViewSlownikCoDalej.AllowDelete = False
                DataViewSlownikCoDalej.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagSlownikCoDalej.BackgroundColor = System.Drawing.SystemColors.Control
                dagSlownikCoDalej.GridColor = System.Drawing.SystemColors.ControlDark
                dagSlownikCoDalej.ForeColor = System.Drawing.Color.Purple
                dagSlownikCoDalej.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagSlownikCoDalej.BorderStyle = BorderStyle.Fixed3D
                'dagSlownikCoDalej.Dock = DockStyle.Fill
                dagSlownikCoDalej.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagSlownikCoDalej.AllowUserToAddRows = False
                dagSlownikCoDalej.AllowUserToDeleteRows = False
                dagSlownikCoDalej.AllowUserToOrderColumns = True
                dagSlownikCoDalej.AllowUserToResizeColumns = True
                dagSlownikCoDalej.AllowUserToResizeRows = True
                dagSlownikCoDalej.ReadOnly = True
                dagSlownikCoDalej.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagSlownikCoDalej.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagSlownikCoDalej.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagSlownikCoDalej.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagSlownikCoDalej.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagSlownikCoDalej.ColumnHeadersVisible = True
                dagSlownikCoDalej.ColumnHeadersHeight = 40
                dagSlownikCoDalej.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagSlownikCoDalej.RowHeadersVisible = True
                dagSlownikCoDalej.RowHeadersWidth = 50
                dagSlownikCoDalej.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagSlownikCoDalej.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagSlownikCoDalej.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagSlownikCoDalej.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagSlownikCoDalej.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagSlownikCoDalej.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagSlownikCoDalej.DefaultCellStyle.NullValue = ""
                dagSlownikCoDalej.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagSlownikCoDalej.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagSlownikCoDalej.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagSlownikCoDalej.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagSlownikCoDalej.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagSlownikCoDalej.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagSlownikCoDalej.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagSlownikCoDalej.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagSlownikCoDalej.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagSlownikCoDalej.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagSlownikCoDalej.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagSlownikCoDalej.DataMember = DataSetSlownikCoDalej.Tables(0).TableName
                'dagSlownikCoDalej.DataSource = DataViewSlownikCoDalej
                '-----------------------------------

                TxtColumnView(dagSlownikCoDalej, DataSetSlownikCoDalej.Tables(0).Columns(0).ColumnName, "Identyfikator", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagSlownikCoDalej, DataSetSlownikCoDalej.Tables(0).Columns(1).ColumnName, "Kiedy stosowaæ", 400, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagSlownikCoDalej, DataSetSlownikCoDalej.Tables(0).Columns(2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                Me.Cursor = Cursors.WaitCursor
                dagSlownikCoDalej.DataSource = DataViewSlownikCoDalej
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetSlownikCoDalej.Tables(0).Rows.Count > 0 Then
                    dagSlownikCoDalej.CurrentCell = dagSlownikCoDalej(0, 0)
                    dagSlownikCoDalej.CurrentCell.Selected = True
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
            stbSlownikCoDalej.BackColor = System.Drawing.SystemColors.Control
            stbSlownikCoDalej.ForeColor = System.Drawing.Color.DarkGreen
            stbSlownikCoDalej.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbSlownikCoDalej.Panels.Add("stbPanelIleRec")
            stbSlownikCoDalej.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbSlownikCoDalej.Panels(0).Width = 80
            stbSlownikCoDalej.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbSlownikCoDalej.Panels(0).Text = "Iloœæ rec.: " & DataSetSlownikCoDalej.Tables(0).Rows.Count.ToString

            stbSlownikCoDalej.Panels.Add("stbPanelWiersz")
            stbSlownikCoDalej.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbSlownikCoDalej.Panels(1).Width = 80
            stbSlownikCoDalej.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagSlownikCoDalej.CurrentCell Is Nothing Then
                stbSlownikCoDalej.Panels(1).Text = "Wiersz : "
            Else
                stbSlownikCoDalej.Panels(1).Text = "Wiersz : " & dagSlownikCoDalej.CurrentCell.RowIndex.ToString
            End If

            stbSlownikCoDalej.Panels.Add("stbPanelKolumna")
            stbSlownikCoDalej.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbSlownikCoDalej.Panels(2).Width = 80
            stbSlownikCoDalej.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagSlownikCoDalej.CurrentCell Is Nothing Then
                stbSlownikCoDalej.Panels(2).Text = "Kolumna : "
            Else
                stbSlownikCoDalej.Panels(2).Text = "Kolumna : " & dagSlownikCoDalej.CurrentCell.ColumnIndex.ToString
            End If

            stbSlownikCoDalej.Panels.Add("stbPanelSort")
            stbSlownikCoDalej.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbSlownikCoDalej.Panels(3).Width = 150
            stbSlownikCoDalej.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbSlownikCoDalej.Panels(3).Text = "Sort : "

            stbSlownikCoDalej.Panels.Add("stbPanelFiltr")
            stbSlownikCoDalej.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbSlownikCoDalej.Panels(4).Width = 150
            stbSlownikCoDalej.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbSlownikCoDalej.Panels(4).Text = "Filtr : "

            stbSlownikCoDalej.Panels.Add("stbPanelReszta")
            stbSlownikCoDalej.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbSlownikCoDalej.Panels(5).Width = 150
            stbSlownikCoDalej.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbSlownikCoDalej.Panels(5).Text = ""

            stbSlownikCoDalej.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagSlownikCoDalej.TableStyles(0).RowHeaderWidth
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
        InitSlownikCoDalej() 'inicjuj zmienne
        '-----------------------------------------------------------------
        If _OCoMamRobic = "WYBIERAJ" Then
            dagSlownikCoDalej.ContextMenuStrip = Me.menuWybieraj
            cmdStop.Text = "Wybierz"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = True
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = False
        Else
            dagSlownikCoDalej.ContextMenuStrip = Me.menuEdytuj
            cmdStop.Text = "Powrót"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = False
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = True
        End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagSlownikCoDalej.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagSlownikCoDalej.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetSlownikCoDalej.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagSlownikCoDalej.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagSlownikCoDalej.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagSlownikCoDalej.FirstDisplayedScrollingColumnIndex +
                        dagSlownikCoDalej.GetCellDisplayRectangle(dagSlownikCoDalej.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagSlownikCoDalej.Left + 4), (dagSlownikCoDalej.Top + 4))
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
        For ii = 0 To DataViewSlownikCoDalej.Table.Columns.Count - 1
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
        dagSlownikCoDalej_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub



    Private Sub KatalogSlownikCoDalej_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub




















    Private Sub dagSlownikCoDalej_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagSlownikCoDalej.CurrentCellChanged
        If Not StartFormy Then
            If dagSlownikCoDalej.CurrentCell Is Nothing Then
                stbSlownikCoDalej.Panels(1).Text = "Wi: "
                stbSlownikCoDalej.Panels(2).Text = "Ko: "
            Else
                stbSlownikCoDalej.Panels(1).Text = "Wi: " & dagSlownikCoDalej.CurrentCell.RowIndex.ToString
                stbSlownikCoDalej.Panels(2).Text = "Ko: " & dagSlownikCoDalej.CurrentCell.ColumnIndex.ToString
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
                If dagSlownikCoDalej.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagSlownikCoDalej.FirstDisplayedScrollingColumnIndex +
                                    dagSlownikCoDalej.GetCellDisplayRectangle(dagSlownikCoDalej.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagSlownikCoDalej.FirstDisplayedScrollingColumnIndex +
                                    dagSlownikCoDalej.GetCellDisplayRectangle(dagSlownikCoDalej.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagSlownikCoDalej.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagSlownikCoDalej.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagSlownikCoDalej, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, DataViewSlownikCoDalej, stbSlownikCoDalej)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagSlownikCoDalej, Mid(sender.name, 6, 3), "SlownikCoDalej")
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
            CzyscFiltryDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, DataViewSlownikCoDalej, stbSlownikCoDalej)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbSlownikCoDalej.Panels(0).Text = "Il.zap.: " & DataViewSlownikCoDalej.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagSlownikCoDalej_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagSlownikCoDalej.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagSlownikCoDalej.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagSlownikCoDalej.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagSlownikCoDalej_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagSlownikCoDalej.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagSlownikCoDalej, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagSlownikCoDalej, e.RowIndex, 4)

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

    Private Sub dagSlownikCoDalej_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagSlownikCoDalej.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagSlownikCoDalej_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagSlownikCoDalej.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagSlownikCoDalej_Scroll(sender As Object, e As ScrollEventArgs) Handles dagSlownikCoDalej.Scroll
        'If (Not StartFormy) And (DataViewSlownikCoDalej.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewSlownikCoDalej.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagSlownikCoDalej_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagSlownikCoDalej.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagSlownikCoDalej_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagSlownikCoDalej.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagSlownikCoDalej, DataViewSlownikCoDalej.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagSlownikCoDalej_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagSlownikCoDalej.MouseUp
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
                stbSlownikCoDalej.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbSlownikCoDalej.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagSlownikCoDalej(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbSlownikCoDalej.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbSlownikCoDalej.Panels(3).Text = "Sort: " & DataSetSlownikCoDalej.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbSlownikCoDalej.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagSlownikCoDalej(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbSlownikCoDalej.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbSlownikCoDalej.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub




    Private Sub dagSlownikCoDalej_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagSlownikCoDalej.DoubleClick
        If _CoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagSlownikCoDalej_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagSlownikCoDalej.KeyDown
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

    Private Sub InitSlownikCoDalej()
        scdIDENT = ""
        scdOPIS = ""
        scdWersja = 1        'WERSJA      Integer, 2
    End Sub

    Private Sub PobierzSlownikCoDalej()
        'pobierz wartosci przed aktualizacja
        scdIDENT = GetTxtField(dagSlownikCoDalej, 0)
        scdOPIS = GetTxtField(dagSlownikCoDalej, 1)
        scdWersja = GetIntField(dagSlownikCoDalej, 2)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        If DataSetSlownikCoDalej.Tables(0).Rows.Count > 0 Then
            oCoDalej = GetTxtField(dagSlownikCoDalej, 0)
        Else
            oCoDalej = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetSlownikCoDalej.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzSlownikCoDalej()
            Dim Form As New EdytujKatalogSlownikCoDalej
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = scdIDENT
                foundRow = DataSetSlownikCoDalej.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDENT") = scdIDENT
                    foundRow("OPIS") = scdOPIS
                    foundRow("WERSJA") = ZmienWersje(scdWersja)    'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbSlownikCoDalej.Panels(0).Text = "Iloœæ rec.: " & DataSetSlownikCoDalej.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagSlownikCoDalej.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetSlownikCoDalej.Tables(0).Rows.Count <= 0 Then
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
                PobierzSlownikCoDalej()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = scdIDENT
                foundRow = DataSetSlownikCoDalej.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbSlownikCoDalej.Panels(0).Text = "Iloœæ rec.: " & DataSetSlownikCoDalej.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitSlownikCoDalej()
        Dim Form As New EdytujKatalogSlownikCoDalej
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
                findTheseVals(0) = scdIDENT
                foundRow = DataSetSlownikCoDalej.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Identyfikator = " & scdIDENT,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetSlownikCoDalej.Tables(0).NewRow()

                    NewRow("IDENT") = scdIDENT
                    NewRow("OPIS") = scdOPIS
                    NewRow("WERSJA") = 1                     'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetSlownikCoDalej.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbSlownikCoDalej.Panels(0).Text = "Iloœæ rec.: " & DataSetSlownikCoDalej.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagSlownikCoDalej.Update()
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
                    dbDataAdapter.Update(DataSetSlownikCoDalej, "TABELA_SlownikCoDalej")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapter.Fill(DataSetSlownikCoDalej)
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
                    sqlDataAdapter.Update(DataSetSlownikCoDalej, "TABELA_SlownikCoDalej")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetSlownikCoDalej)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub


    Private Sub OdswiezBaze()
        DataSetSlownikCoDalej.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnection.Open()
                If dbConnection.State = ConnectionState.Open Then
                    dbDataAdapter.Fill(DataSetSlownikCoDalej)
                    dbConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetSlownikCoDalej)
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
        dagSlownikCoDalej.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagSlownikCoDalej.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagSlownikCoDalej.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagSlownikCoDalej.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub




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
        dagSlownikCoDalej.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagSlownikCoDalej.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        dagSlownikCoDalej.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagSlownikCoDalej.Refresh()
    End Sub

    Private Sub menuWOdswiez_Click(sender As Object, e As EventArgs) Handles menuWOdswiez.Click
        OdswiezBaze()
    End Sub

End Class
