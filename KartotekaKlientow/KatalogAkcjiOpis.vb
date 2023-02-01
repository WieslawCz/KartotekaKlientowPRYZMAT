Imports System.Drawing.Printing

Public Class KatalogAkcjiOpis
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
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents stbAkcjeOpis As System.Windows.Forms.StatusBar
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystko As Button
    Friend WithEvents stbFiltry As StatusBar
    Friend WithEvents dagAKcjeOpis As DataGridView
    Friend WithEvents HorizScrollTimer As Timer
    Friend WithEvents menuWybieraj As ContextMenuStrip
    Friend WithEvents menuWEdytuj As ToolStripMenuItem
    Friend WithEvents menuWDodaj As ToolStripMenuItem
    Friend WithEvents menuWUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents menuWWybierz As ToolStripMenuItem
    Friend WithEvents menuEdytuj As ContextMenuStrip
    Friend WithEvents menuEEdytuj As ToolStripMenuItem
    Friend WithEvents menuEDodaj As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents menuWSingleL As ToolStripMenuItem
    Friend WithEvents menuWMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents menuWOdswiez As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents menuEUsun As ToolStripMenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogAkcjiOpis))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbAkcjeOpis = New System.Windows.Forms.StatusBar()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.dagAKcjeOpis = New System.Windows.Forms.DataGridView()
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
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagAKcjeOpis, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.cmdStop.Location = New System.Drawing.Point(690, 329)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 23)
        Me.cmdStop.TabIndex = 37
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'stbAkcjeOpis
        '
        Me.stbAkcjeOpis.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbAkcjeOpis.Dock = System.Windows.Forms.DockStyle.None
        Me.stbAkcjeOpis.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbAkcjeOpis.Location = New System.Drawing.Point(8, 301)
        Me.stbAkcjeOpis.Name = "stbAkcjeOpis"
        Me.stbAkcjeOpis.ShowPanels = True
        Me.stbAkcjeOpis.Size = New System.Drawing.Size(770, 22)
        Me.stbAkcjeOpis.TabIndex = 42
        Me.stbAkcjeOpis.Text = "StatusBar1"
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(604, 329)
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
        Me.cmdUsuñ.Location = New System.Drawing.Point(518, 329)
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
        Me.cmdDodaj.Location = New System.Drawing.Point(433, 329)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodaj.TabIndex = 39
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 329)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 38
        Me.PictureBox1.TabStop = False
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(546, 10)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 183
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(603, 8)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystko.TabIndex = 181
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
        Me.stbFiltry.TabIndex = 182
        Me.stbFiltry.Text = "StatusBar1"
        '
        'dagAKcjeOpis
        '
        Me.dagAKcjeOpis.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagAKcjeOpis.Location = New System.Drawing.Point(8, 30)
        Me.dagAKcjeOpis.Name = "dagAKcjeOpis"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagAKcjeOpis.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagAKcjeOpis.Size = New System.Drawing.Size(770, 269)
        Me.dagAKcjeOpis.TabIndex = 180
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
        'KatalogAkcjiOpis
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 361)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.dagAKcjeOpis)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.stbAkcjeOpis)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "KatalogAkcjiOpis"
        Me.ShowInTaskbar = False
        Me.Text = "Katalog Akcji Marketingowych"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagAKcjeOpis, System.ComponentModel.ISupportInitialize).EndInit()
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


    'Akcje Handlowe - Opis
    'Public aoIdentAkcji As String            'IDENT     Text 20
    'Public aoDataAkcji As String             'DATA      Text,10
    'Public aoOpisAkcji As String             'OPIS      Text,100
    'Public aoUwagiAkcji As String            'UWAGI     Text,10
    'Public aoMarkerAkcji As String           'MARKER    logical
    'Public aoWersjaAkcji As Integer           'WERSJA    Integer

    Dim dbSelect As String = sSelectSQLAkcjeOpis

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

    Dim DataSetAkcjeOpis As New DataSet
    Dim DataViewAkcjeOpis As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim _CoMamRobic As String = ""

    Private Sub KatalogAkcjiOpis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        _CoMamRobic = oCoMamRobic
        '-------------------------------------
        DataSetAkcjeOpis.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionAkcjeOpis)
            'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLAkcjeOpis, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLAkcjeOpis, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLAkcjeOpis, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLAkcjeOpis, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_AkcjeOpis")

            'DBDeleteAkcjeOpis(dbCommandDelete, dbDataAdapter)
            'DBUpdateAkcjeOpis(dbCommandUpdate, dbDataAdapter)
            'DBInsertAkcjeOpis(dbCommandInsert, dbDataAdapter)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionAkcjeOpis)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQLAkcjeOpis, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLAkcjeOpis, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLAkcjeOpis, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLAkcjeOpis, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_AkcjeOpis")

            SQLDeleteAkcjeOpis(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateAkcjeOpis(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertAkcjeOpis(sqlCommandInsert, sqlDataAdapter)

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
                    'dbDataAdapter.FillSchema(DataSetAkcjeOpis, SchemaType.Mapped, "TABELA_AkcjeOpis")
                    'dbDataAdapter.Fill(DataSetAkcjeOpis)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetAkcjeOpis, SchemaType.Mapped, "TABELA_AkcjeOpis")
                    sqlDataAdapter.Fill(DataSetAkcjeOpis)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetAkcjeOpis.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeOpis.Tables(0).Columns("IDENT")}

                'definiuj DataView
                DataViewAkcjeOpis = New DataView(DataSetAkcjeOpis.Tables(0))
                DataViewAkcjeOpis.AllowDelete = False
                DataViewAkcjeOpis.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagAKcjeOpis.BackgroundColor = System.Drawing.SystemColors.Control
                dagAKcjeOpis.GridColor = System.Drawing.SystemColors.ControlDark
                dagAKcjeOpis.ForeColor = System.Drawing.Color.Purple
                dagAKcjeOpis.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagAKcjeOpis.BorderStyle = BorderStyle.Fixed3D
                'dagAkcjeOpis.Dock = DockStyle.Fill
                dagAKcjeOpis.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagAKcjeOpis.AllowUserToAddRows = False
                dagAKcjeOpis.AllowUserToDeleteRows = False
                dagAKcjeOpis.AllowUserToOrderColumns = True
                dagAKcjeOpis.AllowUserToResizeColumns = True
                dagAKcjeOpis.AllowUserToResizeRows = True
                dagAKcjeOpis.ReadOnly = True
                dagAKcjeOpis.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagAKcjeOpis.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagAkcjeOpis.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagAKcjeOpis.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagAKcjeOpis.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagAKcjeOpis.ColumnHeadersVisible = True
                dagAKcjeOpis.ColumnHeadersHeight = 40
                dagAKcjeOpis.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagAKcjeOpis.RowHeadersVisible = True
                dagAKcjeOpis.RowHeadersWidth = 50
                dagAKcjeOpis.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagAKcjeOpis.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagAKcjeOpis.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagAKcjeOpis.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagAKcjeOpis.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagAKcjeOpis.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagAKcjeOpis.DefaultCellStyle.NullValue = ""
                dagAKcjeOpis.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagAKcjeOpis.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagAKcjeOpis.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagAKcjeOpis.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagAKcjeOpis.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagAKcjeOpis.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagAKcjeOpis.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagAKcjeOpis.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagAKcjeOpis.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagAKcjeOpis.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagAKcjeOpis.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagAKcjeOpis.DataMember = DataSetAkcjeOpis.Tables(0).TableName
                'dagAkcjeOpis.DataSource = DataViewAkcjeOpis
                '-----------------------------------

                TxtColumnView(dagAKcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(0).ColumnName, "Ident", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagAKcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(1).ColumnName, "Data", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagAKcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(2).ColumnName, "Opis", 300, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagAKcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(3).ColumnName, "Uwagi", 300, DataGridViewContentAlignment.MiddleLeft, True)
                LogColumnView(dagAKcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(4).ColumnName, "", 0, False)
                TxtColumnView(dagAKcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(5).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                Me.Cursor = Cursors.WaitCursor
                dagAKcjeOpis.DataSource = DataViewAkcjeOpis
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetAkcjeOpis.Tables(0).Rows.Count > 0 Then
                    dagAKcjeOpis.CurrentCell = dagAKcjeOpis(0, 0)
                    dagAKcjeOpis.CurrentCell.Selected = True
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
            stbAkcjeOpis.BackColor = System.Drawing.SystemColors.Control
            stbAkcjeOpis.ForeColor = System.Drawing.Color.DarkGreen
            stbAkcjeOpis.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbAkcjeOpis.Panels.Add("stbPanelIleRec")
            stbAkcjeOpis.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(0).Width = 80
            stbAkcjeOpis.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(0).Text = "Iloœæ rec.: " & DataSetAkcjeOpis.Tables(0).Rows.Count.ToString

            stbAkcjeOpis.Panels.Add("stbPanelWiersz")
            stbAkcjeOpis.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(1).Width = 80
            stbAkcjeOpis.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagAKcjeOpis.CurrentCell Is Nothing Then
                stbAkcjeOpis.Panels(1).Text = "Wiersz : "
            Else
                stbAkcjeOpis.Panels(1).Text = "Wiersz : " & dagAKcjeOpis.CurrentCell.RowIndex.ToString
            End If

            stbAkcjeOpis.Panels.Add("stbPanelKolumna")
            stbAkcjeOpis.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(2).Width = 80
            stbAkcjeOpis.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagAKcjeOpis.CurrentCell Is Nothing Then
                stbAkcjeOpis.Panels(2).Text = "Kolumna : "
            Else
                stbAkcjeOpis.Panels(2).Text = "Kolumna : " & dagAKcjeOpis.CurrentCell.ColumnIndex.ToString
            End If

            stbAkcjeOpis.Panels.Add("stbPanelSort")
            stbAkcjeOpis.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(3).Width = 150
            stbAkcjeOpis.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(3).Text = "Sort : "

            stbAkcjeOpis.Panels.Add("stbPanelFiltr")
            stbAkcjeOpis.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(4).Width = 150
            stbAkcjeOpis.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(4).Text = "Filtr : "

            stbAkcjeOpis.Panels.Add("stbPanelReszta")
            stbAkcjeOpis.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbAkcjeOpis.Panels(5).Width = 150
            stbAkcjeOpis.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(5).Text = ""

            stbAkcjeOpis.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagAkcjeOpis.TableStyles(0).RowHeaderWidth
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
        InitAkcjeOpis() 'inicjuj zmienne
        '-----------------------------------------------------------------
        If _OCoMamRobic = "WYBIERAJ" Then
            dagAKcjeOpis.ContextMenuStrip = Me.menuWybieraj
            cmdStop.Text = "Wybierz"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = True
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = False
        Else
            dagAKcjeOpis.ContextMenuStrip = Me.menuEdytuj
            cmdStop.Text = "Powrót"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = False
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = True
        End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagAKcjeOpis.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagAkcjeOpis.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetAkcjeOpis.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagAKcjeOpis.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagAKcjeOpis.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagAKcjeOpis.FirstDisplayedScrollingColumnIndex +
                        dagAKcjeOpis.GetCellDisplayRectangle(dagAKcjeOpis.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagAKcjeOpis.Left + 4), (dagAKcjeOpis.Top + 4))
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
        For ii = 0 To DataViewAkcjeOpis.Table.Columns.Count - 1
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
        dagAkcjeOpis_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub

    Private Sub KatalogAkcjiOpis_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub









    Private Sub dagAkcjeOpis_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAKcjeOpis.CurrentCellChanged
        If Not StartFormy Then
            If dagAKcjeOpis.CurrentCell Is Nothing Then
                stbAkcjeOpis.Panels(1).Text = "Wi: "
                stbAkcjeOpis.Panels(2).Text = "Ko: "
            Else
                stbAkcjeOpis.Panels(1).Text = "Wi: " & dagAKcjeOpis.CurrentCell.RowIndex.ToString
                stbAkcjeOpis.Panels(2).Text = "Ko: " & dagAKcjeOpis.CurrentCell.ColumnIndex.ToString
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
                If dagAKcjeOpis.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagAKcjeOpis.FirstDisplayedScrollingColumnIndex +
                                    dagAKcjeOpis.GetCellDisplayRectangle(dagAKcjeOpis.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagAKcjeOpis.FirstDisplayedScrollingColumnIndex +
                                    dagAKcjeOpis.GetCellDisplayRectangle(dagAKcjeOpis.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagAKcjeOpis.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagAKcjeOpis.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagAKcjeOpis, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, DataViewAkcjeOpis, stbAkcjeOpis)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagAKcjeOpis, Mid(sender.name, 6, 3), "AkcjeOpis")
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
            CzyscFiltryDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, DataViewAkcjeOpis, stbAkcjeOpis)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbAkcjeOpis.Panels(0).Text = "Il.zap.: " & DataViewAkcjeOpis.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagAkcjeOpis_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagAKcjeOpis.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagAKcjeOpis.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagAKcjeOpis.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagAkcjeOpis_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagAkcjeOpis.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagAkcjeOpis, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagAkcjeOpis, e.RowIndex, 4)

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

    Private Sub dagAkcjeOpis_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagAKcjeOpis.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagAkcjeOpis_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAKcjeOpis.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagAkcjeOpis_Scroll(sender As Object, e As ScrollEventArgs) Handles dagAKcjeOpis.Scroll
        'If (Not StartFormy) And (DataViewAkcjeOpis.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagAkcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagAkcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewAkcjeOpis.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagAkcjeOpis_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAKcjeOpis.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagAkcjeOpis_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAKcjeOpis.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagAkcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagAKcjeOpis, DataViewAkcjeOpis.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagAkcjeOpis_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAKcjeOpis.MouseUp
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
                stbAkcjeOpis.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbAkcjeOpis.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagAKcjeOpis(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbAkcjeOpis.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbAkcjeOpis.Panels(3).Text = "Sort: " & DataSetAkcjeOpis.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbAkcjeOpis.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagAkcjeOpis(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbAkcjeOpis.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbAkcjeOpis.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub




    Private Sub dagAkcjeOpis_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAKcjeOpis.DoubleClick
        If _CoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagAkcjeOpis_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagAKcjeOpis.KeyDown
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

    Private Sub InitAkcjeOpis()
        aoIdentAkcji = ""        'IDENT       Text, 10
        aoDataAkcji = SysData()
        aoOpisAkcji = ""
        aoUwagiAkcji = ""
        aoMarkerAkcji = False
        aoWersjaAkcji = 1        'WERSJA      Integer, 2
    End Sub

    Private Sub PobierzAkcjeOpis()
        'pobierz wartosci przed aktualizacja
        aoIdentAkcji = GetTxtField(dagAKcjeOpis, 0)
        aoDataAkcji = GetTxtField(dagAKcjeOpis, 1)
        aoOpisAkcji = GetTxtField(dagAKcjeOpis, 2)
        aoUwagiAkcji = GetTxtField(dagAKcjeOpis, 3)
        aoMarkerAkcji = GetLogField(dagAKcjeOpis, 4)
        aoWersjaAkcji = GetIntField(dagAKcjeOpis, 5)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        If DataSetAkcjeOpis.Tables(0).Rows.Count > 0 Then
            PobierzAkcjeOpis()
            oAkcja = GetTxtField(dagAKcjeOpis, 0)
        Else
            InitAkcjeOpis()
            oAkcja = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetAkcjeOpis.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzAkcjeOpis()
            Dim Form As New EdytujKatalogAkcjiOpis
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = aoIdentAkcji
                foundRow = DataSetAkcjeOpis.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDENT") = aoidentAkcji
                    foundRow("DATA") = aoDataAkcji           'OPIS      Text 50
                    foundRow("OPIS") = aoOpisAkcji           'OPIS      Text 50
                    foundRow("UWAGI") = aoUwagiAkcji           'OPIS      Text 50
                    foundRow("MARKER") = aoMarkerAkcji           'OPIS      Text 50
                    foundRow("WERSJA") = ZmienWersje(aoWersjaAkcji)    'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbAkcjeOpis.Panels(0).Text = "Iloœæ rec.: " & DataSetAkcjeOpis.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagAKcjeOpis.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetAkcjeOpis.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis z katalogu ?",
                        "Proszê o potwierdzenie :",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
                oOperacja = "USUN"
                PobierzAkcjeOpis()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = aoIdentAkcji
                foundRow = DataSetAkcjeOpis.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbAkcjeOpis.Panels(0).Text = "Iloœæ rec.: " & DataSetAkcjeOpis.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    dagAkcjeOpis_CurrentCellChanged(sender, e)
                    '--------------------------------
                    'usun zapisy z tablic podleglych...
                    UsunAkcjeSpec(aoIdentAkcji)
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitAkcjeOpis()
        Dim Form As New EdytujKatalogAkcjiOpis
        Do While True
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
                '--------------------------------
                'usun zapisy z tablic podleglych...
                UsunAkcjeSpec(aoIdentAkcji)
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = aoIdentAkcji
                foundRow = DataSetAkcjeOpis.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Ident.akcji = " & aoIdentAkcji,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetAkcjeOpis.Tables(0).NewRow()

                    NewRow("IDENT") = aoIdentAkcji
                    NewRow("DATA") = aoDataAkcji           'OPIS      Text 50
                    NewRow("OPIS") = aoOpisAkcji           'OPIS      Text 50
                    NewRow("UWAGI") = aoUwagiAkcji           'OPIS      Text 50
                    NewRow("MARKER") = aoMarkerAkcji           'OPIS      Text 50
                    NewRow("WERSJA") = ZmienWersje(aoWersjaAkcji)    'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetAkcjeOpis.Tables(0).Rows.Add(NewRow)

                    stbAkcjeOpis.Panels(0).Text = "Iloœæ rec.: " & DataSetAkcjeOpis.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    dagAKcjeOpis.Update()
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
            '        dbDataAdapter.Update(DataSetAkcjeOpis, "TABELA_AkcjeOpis")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetAkcjeOpis)
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
                    sqlDataAdapter.Update(DataSetAkcjeOpis, "TABELA_AkcjeOpis")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetAkcjeOpis)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub



    Private Sub OdswiezBaze()
        DataSetAkcjeOpis.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetAkcjeOpis)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetAkcjeOpis)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub





    '*************************************************
    '** usuwanie zapisow z katalogu AkcjeSpec
    '*************************************************

    Dim dbSelectSQLAkcjeSpec As String = sSelectSQLAkcjeSpec

    Dim dbConnectionAkcjeSpec As OleDb.OleDbConnection
    Dim dbCommandSelectAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandDeleteAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandUpdateAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandInsertAkcjeSpec As OleDb.OleDbCommand
    Dim dbDataAdapterAkcjeSpec As OleDb.OleDbDataAdapter

    Dim sqlConnectionAkcjeSpec As SqlClient.SqlConnection
    Dim sqlCommandSelectAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandDeleteAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandUpdateAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandInsertAkcjeSpec As SqlClient.SqlCommand
    Dim sqlDataAdapterAkcjeSpec As SqlClient.SqlDataAdapter

    Dim DataSetAkcjeSpec As DataSet
    Dim DataViewAkcjeSpec As DataView


    Private Sub UsunAkcjeSpec(ByVal IdTra As String)
        dbSelectSQLAkcjeSpec = sSelectSQLAkcjeSpec & " WHERE IDENTAKCJI='" & IdTra & "' "

        DataSetAkcjeSpec = New DataSet
        DataViewAkcjeSpec = New DataView

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
            'dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
            'dbCommandDeleteAkcjeSpec = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnectionAkcjeSpec)
            'dbCommandUpdateAkcjeSpec = New OleDb.OleDbCommand(sUpdateSQLAkcjeSpec, dbConnectionAkcjeSpec)
            'dbCommandInsertAkcjeSpec = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnectionAkcjeSpec)
            'dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_usuwanie")

            'Dim dbobjParam As OleDb.OleDbParameter
            ''------------------------------------------
            ''komenda DELETE
            ''parametry filtrowania
            'dbobjParam = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
            'dbobjParam.SourceVersion = DataRowVersion.Original
            'dbobjParam = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygKlient", OleDb.OleDbType.Char, 6, "IDENTKLI")
            'dbobjParam.SourceVersion = DataRowVersion.Original
            'dbobjParam = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            'dbobjParam.SourceVersion = DataRowVersion.Original
            'dbDataAdapterAkcjeSpec.DeleteCommand = dbCommandDeleteAkcjeSpec

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterAkcjeSpec.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionAkcjeSpec.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionAkcjeSpec.State
            'End Try
        Else
            sqlConnectionAkcjeSpec = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
            sqlCommandSelectAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionAkcjeSpec)
            sqlCommandDeleteAkcjeSpec = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnectionAkcjeSpec)
            sqlCommandUpdateAkcjeSpec = New SqlClient.SqlCommand(sUpdateSQLAkcjeSpec, sqlConnectionAkcjeSpec)
            sqlCommandInsertAkcjeSpec = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnectionAkcjeSpec)
            sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_usuwanie")

            Dim sqlobjParam As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            sqlobjParam = sqlCommandDeleteAkcjeSpec.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENTAKCJI")
            sqlobjParam.SourceVersion = DataRowVersion.Original
            sqlobjParam = sqlCommandDeleteAkcjeSpec.Parameters.Add("@orygKlient", SqlDbType.Char, 6, "IDENTKLI")
            sqlobjParam.SourceVersion = DataRowVersion.Original
            sqlobjParam = sqlCommandDeleteAkcjeSpec.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            sqlobjParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapterAkcjeSpec.DeleteCommand = sqlCommandDeleteAkcjeSpec


            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterAkcjeSpec.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionAkcjeSpec.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionAkcjeSpec.State
            End Try
        End If


















        'If ParL_DataBaseType = "MS ACCESS" Then
        '    dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
        '    dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
        '    dbCommandDeleteAkcjeSpec = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnectionAkcjeSpec)
        '    'dbCommandUpdateAkcjeSpec = New OleDb.OleDbCommand(sUpdateSQLAkcjeSpec, dbConnectionAkcjeSpec)
        '    'dbCommandInsertAkcjeSpec = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnectionAkcjeSpec)
        '    dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

        '    Dim dbMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
        '    dbMapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_usuwanie")

        '    Dim dbobjParamAkcjeSpec As OleDb.OleDbParameter
        '    '------------------------------------------
        '    'komenda DELETE
        '    'parametry filtrowania
        '    dbobjParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
        '    dbobjParamAkcjeSpec.SourceVersion = DataRowVersion.Original
        '    dbobjParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygKlient", OleDb.OleDbType.Char, 6, "IDENTKLI")
        '    dbobjParamAkcjeSpec.SourceVersion = DataRowVersion.Original
        '    dbobjParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        '    dbobjParamAkcjeSpec.SourceVersion = DataRowVersion.Original

        '    dbDataAdapter.DeleteCommand = dbCommandDeleteAkcjeSpec

        '    ' Add the RowUpdated event handler.
        '    AddHandler dbDataAdapterAkcjeSpec.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

        '    '------------------------------------------
        '    'wypelnij DATASET
        '    Try
        '        dbConnectionAkcjeSpec.Open()
        '    Catch Ex As System.Exception
        '        ViewInLog(Ex, Me.Name(), Nothing)
        '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
        '        '    System.Windows.Forms.MessageBoxButtons.OK, _
        '        '    MessageBoxIcon.Information, _
        '        '    MessageBoxDefaultButton.Button1)
        '    Finally
        '        ConnectionState = dbConnectionAkcjeSpec.State
        '    End Try
        'Else
        '    sqlConnectionAkcjeSpec = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
        '    sqlCommandSelectAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionAkcjeSpec)
        '    sqlCommandDeleteAkcjeSpec = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnectionAkcjeSpec)
        '    'sqlCommandUpdateAkcjeSpec = New SqlClient.SqlCommand(sUpdateSQLAkcjeSpec, sqlConnectionAkcjeSpec)
        '    'sqlCommandInsertAkcjeSpec = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnectionAkcjeSpec)
        '    sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

        '    Dim sqlMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
        '    sqlMapowanieTabeliAkcjeSpec = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_usuwanie")

        '    Dim sqlobjParamAkcjeSpec As SqlClient.SqlParameter
        '    '------------------------------------------
        '    'komenda DELETE
        '    'parametry filtrowania
        '    sqlobjParamAkcjeSpec = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENTAKCJI")
        '    sqlobjParamAkcjeSpec.SourceVersion = DataRowVersion.Original
        '    sqlobjParamAkcjeSpec = sqlCommandDelete.Parameters.Add("@orygKlient", SqlDbType.Char, 6, "IDENTKLI")
        '    sqlobjParamAkcjeSpec.SourceVersion = DataRowVersion.Original
        '    sqlobjParamAkcjeSpec = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        '    sqlobjParamAkcjeSpec.SourceVersion = DataRowVersion.Original
        '    sqlDataAdapterAkcjeSpec.DeleteCommand = sqlCommandDeleteAkcjeSpec

        '    ' Add the RowUpdated event handler.
        '    AddHandler sqlDataAdapterAkcjeSpec.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)
        '    '------------------------------------------
        '    'wypelnij DATASET
        '    Try
        '        sqlConnectionAkcjeSpec.Open()
        '    Catch Ex As System.Exception
        '        ViewInLog(Ex, Me.Name(), Nothing)
        '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
        '        '    System.Windows.Forms.MessageBoxButtons.OK, _
        '        '    MessageBoxIcon.Information, _
        '        '    MessageBoxDefaultButton.Button1)
        '    Finally
        '        ConnectionState = sqlConnectionAkcjeSpec.State
        '    End Try
        'End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_usuwanie")
                    'dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                    'dbConnectionAkcjeSpec.Close()
                Else
                    sqlDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_usuwanie")
                    sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                    sqlConnectionAkcjeSpec.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

                'definiuj DataView
                'DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                'DataViewAkcjeSpec.AllowDelete = True

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
        End If

        Dim i As Integer
        If DataSetAkcjeSpec.Tables(0).Rows.Count > 0 Then
            For i = DataSetAkcjeSpec.Tables(0).Rows.Count - 1 To 0 Step -1
                DataSetAkcjeSpec.Tables(0).Rows.Item(i).Delete()
            Next
            AktualizujBazeAkcjeSpec()
        End If
    End Sub




    Private Sub AktualizujBazeAkcjeSpec()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionAkcjeSpec.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionAkcjeSpec.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_usuwanie")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
            '    End If
            '    dbConnectionAkcjeSpec.Close()
            'End If
        Else
            Try
                sqlConnectionAkcjeSpec.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionAkcjeSpec.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_usuwanie")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                End If
                sqlConnectionAkcjeSpec.Close()
            End If
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
        dagAKcjeOpis.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagAKcjeOpis.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        dagAKcjeOpis.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagAKcjeOpis.Refresh()
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
        dagAKcjeOpis.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagAKcjeOpis.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagAKcjeOpis.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagAKcjeOpis.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub
End Class
