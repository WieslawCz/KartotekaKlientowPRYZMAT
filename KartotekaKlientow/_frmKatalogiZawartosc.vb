
Public Class _frmKatalogiZawartosc
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Dim _Select As String = ""
    Public Sub New(ByVal pSelect As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _Select = pSelect
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
    Friend WithEvents dagTabela As System.Windows.Forms.DataGridView
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents pnlClear As Panel
    Friend WithEvents cmdClear As Button
    Friend WithEvents rbtDoTegoMiesiaca As RadioButton
    Friend WithEvents rbtTylkoMiesiac As RadioButton
    Friend WithEvents cbxMiesiac As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents cbxRok As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents stbTabela As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dagTabela = New System.Windows.Forms.DataGridView()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.stbTabela = New System.Windows.Forms.StatusBar()
        Me.pnlClear = New System.Windows.Forms.Panel()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.rbtDoTegoMiesiaca = New System.Windows.Forms.RadioButton()
        Me.rbtTylkoMiesiac = New System.Windows.Forms.RadioButton()
        Me.cbxMiesiac = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxRok = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.dagTabela, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlClear.SuspendLayout()
        Me.SuspendLayout()
        '
        'dagTabela
        '
        Me.dagTabela.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagTabela.Location = New System.Drawing.Point(8, 8)
        Me.dagTabela.Name = "dagTabela"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagTabela.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagTabela.Size = New System.Drawing.Size(824, 454)
        Me.dagTabela.TabIndex = 0
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(743, 504)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(89, 22)
        Me.cmdPowrot.TabIndex = 15
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 497)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 16
        Me.picLogo.TabStop = False
        '
        'stbTabela
        '
        Me.stbTabela.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbTabela.Dock = System.Windows.Forms.DockStyle.None
        Me.stbTabela.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbTabela.Location = New System.Drawing.Point(8, 462)
        Me.stbTabela.Name = "stbTabela"
        Me.stbTabela.ShowPanels = True
        Me.stbTabela.Size = New System.Drawing.Size(824, 21)
        Me.stbTabela.TabIndex = 21
        Me.stbTabela.Text = "StatusBar1"
        '
        'pnlClear
        '
        Me.pnlClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pnlClear.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlClear.Controls.Add(Me.cmdClear)
        Me.pnlClear.Controls.Add(Me.rbtDoTegoMiesiaca)
        Me.pnlClear.Controls.Add(Me.rbtTylkoMiesiac)
        Me.pnlClear.Controls.Add(Me.cbxMiesiac)
        Me.pnlClear.Controls.Add(Me.Label2)
        Me.pnlClear.Controls.Add(Me.cbxRok)
        Me.pnlClear.Controls.Add(Me.Label1)
        Me.pnlClear.Location = New System.Drawing.Point(126, 489)
        Me.pnlClear.Name = "pnlClear"
        Me.pnlClear.Size = New System.Drawing.Size(611, 45)
        Me.pnlClear.TabIndex = 71
        '
        'cmdClear
        '
        Me.cmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClear.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdClear.ForeColor = System.Drawing.Color.Black
        Me.cmdClear.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClear.Location = New System.Drawing.Point(515, 13)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(80, 23)
        Me.cmdClear.TabIndex = 69
        Me.cmdClear.Text = "Usuñ zapisy"
        Me.cmdClear.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'rbtDoTegoMiesiaca
        '
        Me.rbtDoTegoMiesiaca.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtDoTegoMiesiaca.ForeColor = System.Drawing.Color.Navy
        Me.rbtDoTegoMiesiaca.Location = New System.Drawing.Point(296, 20)
        Me.rbtDoTegoMiesiaca.Name = "rbtDoTegoMiesiaca"
        Me.rbtDoTegoMiesiaca.Size = New System.Drawing.Size(200, 18)
        Me.rbtDoTegoMiesiaca.TabIndex = 75
        Me.rbtDoTegoMiesiaca.TabStop = True
        Me.rbtDoTegoMiesiaca.Text = "Do wybranego miesi¹ca w³¹cznie"
        Me.rbtDoTegoMiesiaca.UseVisualStyleBackColor = True
        '
        'rbtTylkoMiesiac
        '
        Me.rbtTylkoMiesiac.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtTylkoMiesiac.ForeColor = System.Drawing.Color.Navy
        Me.rbtTylkoMiesiac.Location = New System.Drawing.Point(296, 0)
        Me.rbtTylkoMiesiac.Name = "rbtTylkoMiesiac"
        Me.rbtTylkoMiesiac.Size = New System.Drawing.Size(162, 18)
        Me.rbtTylkoMiesiac.TabIndex = 74
        Me.rbtTylkoMiesiac.TabStop = True
        Me.rbtTylkoMiesiac.Text = "Tylko wybrany miesi¹c/rok"
        Me.rbtTylkoMiesiac.UseVisualStyleBackColor = True
        '
        'cbxMiesiac
        '
        Me.cbxMiesiac.Location = New System.Drawing.Point(193, 9)
        Me.cbxMiesiac.Name = "cbxMiesiac"
        Me.cbxMiesiac.Size = New System.Drawing.Size(88, 21)
        Me.cbxMiesiac.TabIndex = 72
        Me.cbxMiesiac.Text = "<wszystkie>"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(139, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 73
        Me.Label2.Text = "Miesi¹c . . . . . . . . . . ."
        '
        'cbxRok
        '
        Me.cbxRok.Location = New System.Drawing.Point(38, 9)
        Me.cbxRok.Name = "cbxRok"
        Me.cbxRok.Size = New System.Drawing.Size(88, 21)
        Me.cbxRok.TabIndex = 71
        Me.cbxRok.Text = "<wszystkie>"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(10, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 70
        Me.Label1.Text = "Rok . . . . . . . . . . ."
        '
        '_frmKatalogiZawartosc
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(839, 541)
        Me.Controls.Add(Me.pnlClear)
        Me.Controls.Add(Me.stbTabela)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.dagTabela)
        Me.Name = "_frmKatalogiZawartosc"
        Me.Text = "Katalog zawartosc"
        CType(Me.dagTabela, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlClear.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim StartFormy As Boolean = True

    Dim sConnection As String = DataBaseConnection()
    Dim sSelectSQL As String = "SELECT * FROM " & oTabela
    Dim sDeleteSQL As String = ""

    'Dim dbConnection As OleDb.OleDbConnection
    'Dim dbCommandSelect As OleDb.OleDbCommand
    'Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As SqlClient.SqlConnection
    Dim sqlCommandSelect As SqlClient.SqlCommand
    Dim sqlCommandDelete As SqlClient.SqlCommand
    Dim sqlDataAdapter As SqlClient.SqlDataAdapter

    Dim DataSetTabela As New System.Data.DataSet
    Dim DataViewTabela As New System.Data.DataView
    Dim DataTableTabela As System.Data.DataTable
    Dim ConnectionState As System.Data.ConnectionState


    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub _frmKatalogiZawartosc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.MrJones
        Me.picLogo.Image = My.Resources.logomini
        Me.cmdPowrot.Image = My.Resources._EXIT.ToBitmap
        '-------------------------------
        If Len(_Select) > 0 Then
            sSelectSQL = _Select
        End If
        '----------------------------------
        If ParL_DataBaseType = "MS ACCESS" Then
        Else

            sSelectSQL = "SELECT * FROM " & oTabela
            Select Case UCase(oTabela)
                Case "OBROTY"
                    sDeleteSQL = sDeleteSQLObroty
                Case "OBROTYMIES"
                    sDeleteSQL = sDeleteSQLObrotyMies
                Case "KONTAKTYHANDLOWE"
                    sDeleteSQL = sDeleteSQLKontakty
                Case Else
            End Select
            '------------------------------------------
            DataSetTabela = New DataSet
            DataSetTabela.Locale = New System.Globalization.CultureInfo("pl-PL")

            sConnection = ConnectionStrings()

            sqlConnection = New SqlClient.SqlConnection(sConnection)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQL, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQL, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'Odwzorowanie KursyWalut dla tabeli Kursy
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA")

            '------------------------------------------
            'komenda DELETE
            Select Case UCase(oTabela)
                Case "OBROTY"
                    SQLDeleteObroty(sqlCommandDelete, sqlDataAdapter)
                Case "OBROTYMIES"
                    SQLDeleteObrotyMies(sqlCommandDelete, sqlDataAdapter)
                Case "KONTAKTYHANDLOWE"
                    SQLDeleteKontaktyHandlowe(sqlCommandDelete, sqlDataAdapter)
                Case Else
            End Select

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapter.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)
            '------------------------------------------
            Me.Text = "Zawartoœæ tabeli " & oTabela
            Try
                sqlConnection.Open()
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnection.State
            End Try
        End If


        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapter.FillSchema(DataSetTabela, SchemaType.Mapped, "TABELA")
                    'dbDataAdapter.Fill(DataSetTabela)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetTabela, SchemaType.Mapped, "TABELA")
                    sqlDataAdapter.Fill(DataSetTabela)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetTabela.Tables(0).PrimaryKey = New DataColumn() {DataSetTabela.Tables(0).Columns("DATANOTOW"), _
                '                                                      DataSetTabela.Tables(0).Columns("WALUTA")}


                'definiuj DataView
                DataViewTabela = New DataView(DataSetTabela.Tables(0))
                DataViewTabela.AllowDelete = False
                DataViewTabela.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagTabela.BackgroundColor = System.Drawing.SystemColors.Control
                dagTabela.GridColor = System.Drawing.SystemColors.ControlDark
                dagTabela.ForeColor = System.Drawing.Color.Purple
                dagTabela.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagTabela.BorderStyle = BorderStyle.Fixed3D
                'dagTabela.Dock = DockStyle.Fill
                dagTabela.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagTabela.AllowUserToAddRows = False
                dagTabela.AllowUserToDeleteRows = False
                dagTabela.AllowUserToOrderColumns = True
                dagTabela.AllowUserToResizeColumns = True
                dagTabela.AllowUserToResizeRows = True
                dagTabela.ReadOnly = True
                dagTabela.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagTabela.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagTabela.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagTabela.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagTabela.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagTabela.ColumnHeadersVisible = True
                dagTabela.ColumnHeadersHeight = 40
                dagTabela.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagTabela.RowHeadersVisible = True
                dagTabela.RowHeadersWidth = 50
                dagTabela.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagTabela.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagTabela.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagTabela.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagTabela.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagTabela.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagTabela.DefaultCellStyle.NullValue = ""
                dagTabela.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagTabela.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagTabela.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagTabela.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagTabela.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagTabela.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagTabela.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagTabela.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagTabela.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagTabela.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagTabela.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagTabela.DataMember = DataSetTabela.Tables(0).TableName
                dagTabela.DataSource = DataViewTabela
                '-----------------------------------

                DataTableTabela = DataSetTabela.Tables("TABELA")
                Dim objColumn As DataColumn
                Dim PNazwa As String = ""
                Dim PDlug As Long = 0
                Dim PTyp As String

                For Each objColumn In DataTableTabela.Columns
                    PNazwa = objColumn.ColumnName
                    PDlug = IIf(objColumn.MaxLength < 10, 10, objColumn.MaxLength)
                    PTyp = objColumn.DataType.Name

                    Select Case PTyp
                        Case "String"
                            TxtColumnView(dagTabela, PNazwa, PNazwa, 4 * IIf(PDlug > 100, 100, PDlug), DataGridViewContentAlignment.MiddleLeft, True)
                        Case "Integer", "Int64", "Int16", "Int32", "Single", "Long"
                            NumColumnView(dagTabela, PNazwa, PNazwa, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                        Case "Double", "Float"
                            NumColumnView(dagTabela, PNazwa, PNazwa, 80, DataGridViewContentAlignment.MiddleRight, "N4", True)
                        Case "Boolean"
                            LogColumnView(dagTabela, PNazwa, PNazwa, 40, True)
                        Case Else
                            'TxtColumnView(dagTabela, PNazwa, PNazwa, 5 * IIf(PDlug > 100, 100, PDlug), DataGridViewContentAlignment.MiddleLeft, True)
                            TxtColumnView(dagTabela, PNazwa, PNazwa, 80, DataGridViewContentAlignment.MiddleLeft, True)
                    End Select
                Next

                'ustaw sie na pierwszym zapisie
                If DataSetTabela.Tables(0).Rows.Count > 0 Then
                    dagTabela.CurrentCell = dagTabela(0, 0)
                    dagTabela.CurrentCell.Selected = True
                End If


                'dodaj StatusBar na koncu
                stbTabela.BackColor = System.Drawing.SystemColors.Control
                'stbTabela.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
                stbTabela.ForeColor = System.Drawing.Color.DarkGreen
                stbTabela.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                stbTabela.Panels.Add("stbPanelIleRec")
                stbTabela.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                stbTabela.Panels(0).Width = 80
                stbTabela.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                stbTabela.Panels(0).Text = "Iloœæ rec.: " & DataSetTabela.Tables(0).Rows.Count.ToString

                stbTabela.Panels.Add("stbPanelWiersz")
                stbTabela.Panels(1).AutoSize = StatusBarPanelAutoSize.None
                stbTabela.Panels(1).Width = 80
                stbTabela.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                If dagTabela.CurrentCell Is Nothing Then
                    stbTabela.Panels(1).Text = "Wi: "
                Else
                    stbTabela.Panels(1).Text = "Wi: " & dagTabela.CurrentCell.ColumnIndex.ToString
                End If

                stbTabela.Panels.Add("stbPanelKolumna")
                stbTabela.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                stbTabela.Panels(2).Width = 80
                stbTabela.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                If dagTabela.CurrentCell Is Nothing Then
                    stbTabela.Panels(2).Text = "Ko: "
                Else
                    stbTabela.Panels(2).Text = "Ko: " & dagTabela.CurrentCell.RowIndex.ToString
                End If '

                stbTabela.Panels.Add("stbPanelSort")
                stbTabela.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                stbTabela.Panels(3).Width = 150
                stbTabela.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                stbTabela.Panels(3).Text = "Sort: "

                stbTabela.Panels.Add("stbPanelFiltr")
                stbTabela.Panels(4).AutoSize = StatusBarPanelAutoSize.None
                stbTabela.Panels(4).Width = 150
                stbTabela.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
                stbTabela.Panels(4).Text = "Filtr : "

                stbTabela.Panels.Add("stbPanelReszta")
                stbTabela.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
                'stbTabela.Panels(5).Width = 150
                stbTabela.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
                stbTabela.Panels(5).Text = ""

                stbTabela.ShowPanels = True

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End Try
        End If

        Select Case UCase(oTabela)
            Case "OBROTY"
                pnlClear.Visible = True
            Case "OBROTYMIES"
                pnlClear.Visible = True
            Case "KONTAKTYHANDLOWE"
                pnlClear.Visible = True
            Case Else
                pnlClear.Visible = False
        End Select

        Dim Ir As Integer = 0
        cbxRok.Items.Add(Trim("<wszystkie>"))
        For Ir = 1900 To 2100
            cbxRok.Items.Add(Trim(Str(Ir)))
        Next
        cbxRok.Text = "<wszystkie>"

        'Lista miesiêcy
        cbxMiesiac.Items.Add(Trim("<wszystkie>"))
        cbxMiesiac.Items.Add("01 Styczeñ")
        cbxMiesiac.Items.Add("02 Luty")
        cbxMiesiac.Items.Add("03 Marzec")
        cbxMiesiac.Items.Add("04 Kwiecieñ")
        cbxMiesiac.Items.Add("05 Maj")
        cbxMiesiac.Items.Add("06 Czerwiec")
        cbxMiesiac.Items.Add("07 Lipiec")
        cbxMiesiac.Items.Add("08 Sierpieñ")
        cbxMiesiac.Items.Add("09 Wrzesieñ")
        cbxMiesiac.Items.Add("10 PaŸdziernik")
        cbxMiesiac.Items.Add("11 Listopad")
        cbxMiesiac.Items.Add("12 Grudzieñ")
        cbxMiesiac.Text = "<wszystkie>"

        rbtTylkoMiesiac.Checked = True
        '----------------------
        StartFormy = False
    End Sub

    Private Sub dagTabela_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagTabela.CurrentCellChanged
        If StartFormy Then
        Else
            If dagTabela.CurrentCell Is Nothing Then
                stbTabela.Panels(1).Text = "Wi: "
                stbTabela.Panels(2).Text = "Ko: "
            Else
                stbTabela.Panels(1).Text = "Wi: " & dagTabela.CurrentCell.RowIndex.ToString
                stbTabela.Panels(2).Text = "Ko: " & dagTabela.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    Private Sub dagTabela_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagTabela.MouseUp
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
                stbTabela.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbTabela.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagTabela(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbTabela.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbTabela.Panels(3).Text = "Sort: " & DataSetTabela.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbTabela.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                myGridView.CurrentCell = dagTabela(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbTabela.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbTabela.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagTabela_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles dagTabela.Paint
        RysujGradient(Me)
    End Sub

    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagTabela_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagTabela.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagTabela.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagTabela.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub











    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        Dim dr As DataRow = Nothing
        Dim I As Integer = 0
        Dim rok, mies As String
        Dim oMiesiac As String = ""
        Dim IleDel As Integer = 0

        If rbtTylkoMiesiac.Checked Then
            If MessageBox.Show("Czy na pewno mam usun¹æ zapisy z tabeli " & vbCrLf &
                               "dotycz¹ce wybranego roku/miesi¹ca ?",
                    "Proszê o potwierdzenie :",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

                Me.Cursor = Cursors.WaitCursor

                If cbxRok.Text = "<wszystkie>" Then
                    'wszystkie lata...
                    If cbxMiesiac.Text = "<wszystkie>" Then
                        'USUN WSZYSTKO....
                        If DataSetTabela.Tables(0).Rows.Count > 0 Then
                            For I = DataSetTabela.Tables(0).Rows.Count - 1 To 0 Step -1
                                dr = DataSetTabela.Tables(0).Rows.Item(I)

                                IleDel += 1
                                dr.Delete()
                            Next
                        End If

                    Else
                        'USUN MIESIAC Z WSZYSTKICH LAT
                        If DataSetTabela.Tables(0).Rows.Count > 0 Then
                            For I = DataSetTabela.Tables(0).Rows.Count - 1 To 0 Step -1
                                dr = DataSetTabela.Tables(0).Rows.Item(I)

                                Select Case UCase(oTabela)
                                    Case "OBROTY"
                                        oMiesiac = dr.Item("DATA")
                                    Case "OBROTYMIES"
                                        oMiesiac = dr.Item("MIESIAC")
                                    Case "KONTAKTYHANDLOWE"
                                        oMiesiac = dr.Item("DATAKONTAKTU")
                                    Case Else
                                End Select

                                rok = Mid(oMiesiac, 1, 4)
                                mies = Mid(oMiesiac, 6, 2)
                                If Mid(cbxMiesiac.Text, 1, 2) = mies Then
                                    IleDel += 1
                                    dr.Delete()
                                End If
                            Next
                        End If
                    End If
                Else
                    'wybrany rok
                    If cbxMiesiac.Text = "<wszystkie>" Then
                        'USUN WSZYSTKIE MIESIACE WYBRANEGO ROKU
                        If DataSetTabela.Tables(0).Rows.Count > 0 Then
                            For I = DataSetTabela.Tables(0).Rows.Count - 1 To 0 Step -1
                                dr = DataSetTabela.Tables(0).Rows.Item(I)

                                Select Case UCase(oTabela)
                                    Case "OBROTY"
                                        oMiesiac = dr.Item("DATA")
                                    Case "OBROTYMIES"
                                        oMiesiac = dr.Item("MIESIAC")
                                    Case "KONTAKTYHANDLOWE"
                                        oMiesiac = dr.Item("DATAKONTAKTU")
                                    Case Else
                                End Select

                                rok = Mid(oMiesiac, 1, 4)
                                mies = Mid(oMiesiac, 6, 2)
                                If cbxRok.Text = rok Then
                                    IleDel += 1
                                    dr.Delete()
                                End If
                            Next
                        End If
                    Else
                        'USUN WYBRANY ROK I MIESIAC
                        If DataSetTabela.Tables(0).Rows.Count > 0 Then
                            For I = DataSetTabela.Tables(0).Rows.Count - 1 To 0 Step -1
                                dr = DataSetTabela.Tables(0).Rows.Item(I)

                                Select Case UCase(oTabela)
                                    Case "OBROTY"
                                        oMiesiac = dr.Item("DATA")
                                    Case "OBROTYMIES"
                                        oMiesiac = dr.Item("MIESIAC")
                                    Case "KONTAKTYHANDLOWE"
                                        oMiesiac = dr.Item("DATAKONTAKTU")
                                    Case Else
                                End Select

                                rok = Mid(oMiesiac, 1, 4)
                                mies = Mid(oMiesiac, 6, 2)
                                If (cbxRok.Text = rok) And (Mid(cbxMiesiac.Text, 1, 2) = mies) Then
                                    IleDel += 1
                                    dr.Delete()
                                End If
                            Next
                        End If

                    End If
                End If
                AktualizujBaze()
                'DataSetTabela.Clear()
                stbTabela.Panels(0).Text = "Iloœæ rec.: " & DataSetTabela.Tables(0).Rows.Count.ToString

                Me.Cursor = Cursors.Default
                MessageBox.Show("Usun¹³em " & Trim(Str(IleDel)) & " zapisów.",
                                    "OK, Skoñczy³em.",
                                    System.Windows.Forms.MessageBoxButtons.OK,
                                    MessageBoxIcon.Information,
                                    MessageBoxDefaultButton.Button1)
            End If


        Else
            If MessageBox.Show("Czy na pewno mam usun¹æ zapisy z tabeli " & vbCrLf &
                               "do wybranego roku/miesi¹ca w³¹cznie ?",
                    "Proszê o potwierdzenie :",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

                Me.Cursor = Cursors.WaitCursor

                If cbxRok.Text = "<wszystkie>" Then
                    'wszystkie lata...
                    If cbxMiesiac.Text = "<wszystkie>" Then
                        'USUN WSZYSTKO....
                        If DataSetTabela.Tables(0).Rows.Count > 0 Then
                            For I = DataSetTabela.Tables(0).Rows.Count - 1 To 0 Step -1
                                dr = DataSetTabela.Tables(0).Rows.Item(I)

                                IleDel += 1
                                dr.Delete()
                            Next
                        End If
                    Else
                        MessageBox.Show("Proszê zdefiniowaæ ROK i MIESI¥" & vbCrLf &
                                        "do którego w³¹cznie mam usu¹æ zapisy z Tabeli.",
                                        "Nie mogê usun¹¿ zapisów z tabeli...",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1)
                        Return
                    End If
                Else
                    'wybrany rok
                    If cbxMiesiac.Text = "<wszystkie>" Then
                        'USUN WSZYSTKIE MIESIACE WYBRANEGO ROKU
                        If DataSetTabela.Tables(0).Rows.Count > 0 Then
                            For I = DataSetTabela.Tables(0).Rows.Count - 1 To 0 Step -1
                                dr = DataSetTabela.Tables(0).Rows.Item(I)

                                Select Case UCase(oTabela)
                                    Case "OBROTY"
                                        oMiesiac = dr.Item("DATA")
                                    Case "OBROTYMIES"
                                        oMiesiac = dr.Item("MIESIAC")
                                    Case "KONTAKTYHANDLOWE"
                                        oMiesiac = dr.Item("DATAKONTAKTU")
                                    Case Else
                                End Select

                                rok = Mid(oMiesiac, 1, 4)
                                mies = Mid(oMiesiac, 6, 2)
                                If rok <= cbxRok.Text Then
                                    IleDel += 1
                                    dr.Delete()
                                End If
                            Next
                        End If

                    Else
                        'USUN WYBRANY ROK I MIESIAC
                        If DataSetTabela.Tables(0).Rows.Count > 0 Then
                            For I = DataSetTabela.Tables(0).Rows.Count - 1 To 0 Step -1
                                dr = DataSetTabela.Tables(0).Rows.Item(I)

                                Select Case UCase(oTabela)
                                    Case "OBROTY"
                                        oMiesiac = dr.Item("DATA")
                                    Case "OBROTYMIES"
                                        oMiesiac = dr.Item("MIESIAC")
                                    Case "KONTAKTYHANDLOWE"
                                        oMiesiac = dr.Item("DATAKONTAKTU")
                                    Case Else
                                End Select

                                mies = Mid(oMiesiac, 1, 7)
                                If (mies <= cbxRok.Text & "-" & Mid(cbxMiesiac.Text, 1, 2)) Then
                                    IleDel += 1
                                    dr.Delete()
                                End If
                            Next
                        End If
                    End If
                End If
                AktualizujBaze()
                'DataSetTabela.Clear()
                stbTabela.Panels(0).Text = "Iloœæ rec.: " & DataSetTabela.Tables(0).Rows.Count.ToString


                Me.Cursor = Cursors.Default
                MessageBox.Show("Usun¹³em " & Trim(Str(IleDel)) & " zapisów.",
                                    "OK, Skoñczy³em.",
                                    System.Windows.Forms.MessageBoxButtons.OK,
                                    MessageBoxIcon.Information,
                                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub


    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        Try
            sqlConnection.Open()
        Catch Ex As System.Exception
            ViewInLog(Ex, Me.Name(), Nothing)
        End Try

        If sqlConnection.State = ConnectionState.Open Then
            oWystapilConcurrencyException = False
            '------------------------------------------------------------
            Try
                sqlDataAdapter.Update(DataSetTabela, "TABELA")
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
            '------------------------------------------------------------
            If oWystapilConcurrencyException = True Then
                sqlDataAdapter.Fill(DataSetTabela)
            End If
            sqlConnection.Close()
        End If
    End Sub


End Class
