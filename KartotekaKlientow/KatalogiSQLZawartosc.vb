Public Class KatalogiSQLZawartosc
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
    Friend WithEvents dagTabela As System.Windows.Forms.DataGrid
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents pnlClear As System.Windows.Forms.Panel
    Friend WithEvents cbxMiesiac As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbxRok As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents rbtDoTegoMiesiaca As RadioButton
    Friend WithEvents rbtTylkoMiesiac As RadioButton
    Friend WithEvents stbTabela As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogiSQLZawartosc))
        Me.dagTabela = New System.Windows.Forms.DataGrid()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
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
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlClear.SuspendLayout()
        Me.SuspendLayout()
        '
        'dagTabela
        '
        Me.dagTabela.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagTabela.DataMember = ""
        Me.dagTabela.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dagTabela.Location = New System.Drawing.Point(8, 8)
        Me.dagTabela.Name = "dagTabela"
        Me.dagTabela.Size = New System.Drawing.Size(813, 437)
        Me.dagTabela.TabIndex = 0
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(731, 480)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(89, 22)
        Me.cmdStop.TabIndex = 15
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 481)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 16
        Me.PictureBox1.TabStop = False
        '
        'stbTabela
        '
        Me.stbTabela.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbTabela.Dock = System.Windows.Forms.DockStyle.None
        Me.stbTabela.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbTabela.Location = New System.Drawing.Point(8, 444)
        Me.stbTabela.Name = "stbTabela"
        Me.stbTabela.ShowPanels = True
        Me.stbTabela.Size = New System.Drawing.Size(812, 20)
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
        Me.pnlClear.Location = New System.Drawing.Point(122, 469)
        Me.pnlClear.Name = "pnlClear"
        Me.pnlClear.Size = New System.Drawing.Size(588, 40)
        Me.pnlClear.TabIndex = 70
        '
        'cmdClear
        '
        Me.cmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClear.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdClear.ForeColor = System.Drawing.Color.Black
        Me.cmdClear.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClear.Location = New System.Drawing.Point(492, 8)
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
        'KatalogiSQLZawartosc
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(827, 517)
        Me.Controls.Add(Me.pnlClear)
        Me.Controls.Add(Me.stbTabela)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.dagTabela)
        Me.Name = "KatalogiSQLZawartosc"
        Me.ShowInTaskbar = False
        Me.Text = "Katalog zawartosc"
        CType(Me.dagTabela, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlClear.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim sConnection As String = ConnectionStrings()

    Dim sSelectSQL As String = ""
    Dim sDeleteSQL As String = ""

    Dim dbConnection As SqlClient.SqlConnection
    Dim dbCommandSelect As SqlClient.SqlCommand
    Dim dbCommandDelete As SqlClient.SqlCommand

    Dim dbDataAdapter As SqlClient.SqlDataAdapter
    Dim DataSetTabela As New DataSet
    Dim MapowanieTabeli As System.Data.Common.DataTableMapping

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub KatalogiSQLZawartosc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
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

        dbConnection = New SqlClient.SqlConnection(sConnection)
        dbCommandSelect = New SqlClient.SqlCommand(sSelectSQL, dbConnection)
        dbCommandDelete = New SqlClient.SqlCommand(sDeleteSQL, dbConnection)
        dbDataAdapter = New SqlClient.SqlDataAdapter(dbCommandSelect)

        Dim MapowanieTabeli As System.Data.Common.DataTableMapping
        'Odwzorowanie KursyWalut dla tabeli Kursy
        MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA")
        'DataSet

        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        Select Case UCase(oTabela)
            Case "OBROTY"
                SQLDeleteObroty(dbCommandDelete, dbDataAdapter)
            Case "OBROTYMIES"
                SQLDeleteObrotyMies(dbCommandDelete, dbDataAdapter)
            Case "KONTAKTYHANDLOWE"
                SQLDeleteKontaktyHandlowe(dbCommandDelete, dbDataAdapter)
            Case Else
        End Select

        ' Add the RowUpdated event handler.
        AddHandler dbDataAdapter.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)
        '------------------------------------------
        'wypelnij DATASET
        Me.Text = "Zawartoœæ tabeli " & oTabela
        Try
            dbConnection.Open()
        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        End Try

        If dbConnection.State = ConnectionState.Open Then
            Try
                dbDataAdapter.FillSchema(DataSetTabela, SchemaType.Mapped, "TABELA")
                dbDataAdapter.Fill(DataSetTabela)
                dbConnection.Close()

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetTabela.Tables(0).PrimaryKey = New DataColumn() {DataSetTabela.Tables(0).Columns("DATANOTOW"), _
                '                                                      DataSetTabela.Tables(0).Columns("WALUTA")}





                'parametry DataGrid
                dagTabela.CaptionVisible = False
                dagTabela.CaptionText = DataSetTabela.Tables(0).TableName
                dagTabela.ColumnHeadersVisible = True
                dagTabela.RowHeadersVisible = True
                dagTabela.BackgroundColor = System.Drawing.SystemColors.Control
                dagTabela.BorderStyle = BorderStyle.Fixed3D
                dagTabela.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                dagTabela.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
                dagTabela.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
                dagTabela.ReadOnly = True

                'wypelnienie DataGrid
                dagTabela.DataSource = DataSetTabela
                dagTabela.DataMember = DataSetTabela.Tables(0).TableName
                dagTabela.SetDataBinding(DataSetTabela, "TABELA")












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
                stbTabela.Panels(1).Text = "Wiersz : " & dagTabela.CurrentCell.ColumnNumber.ToString

                stbTabela.Panels.Add("stbPanelKolumna")
                stbTabela.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                stbTabela.Panels(2).Width = 80
                stbTabela.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                stbTabela.Panels(2).Text = "Kolumna : " & dagTabela.CurrentCell.RowNumber.ToString

                stbTabela.Panels.Add("stbPanelSort")
                stbTabela.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                stbTabela.Panels(3).Width = 150
                stbTabela.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                stbTabela.Panels(3).Text = "Sort : "

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
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
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
    End Sub

    Private Sub dagTabela_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagTabela.CurrentCellChanged
        stbTabela.Panels(1).Text = "Wiersz : " & dagTabela.CurrentCell.RowNumber.ToString
        stbTabela.Panels(2).Text = "Kolumna : " & dagTabela.CurrentCell.ColumnNumber.ToString
    End Sub

    Private Sub dagTabela_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagTabela.MouseUp
        'przypisz biezacy DataGrid do myGrid
        Dim myGrid As DataGrid = CType(sender, DataGrid)
        Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
        hti = myGrid.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                stbTabela.Panels(3).Text = "Sort : " & hti.Column.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
    End Sub

    Private Sub dagTabela_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles dagTabela.Paint
        RysujGradient(Me)
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
            dbConnection.Open()
        Catch Ex As System.Exception
            ViewInLog(ex, Me.Name(), Nothing)
        End Try

        If dbConnection.State = ConnectionState.Open Then
            oWystapilConcurrencyException = False
            '------------------------------------------------------------
            Try
                dbDataAdapter.Update(DataSetTabela, "TABELA")
            Catch Ex As System.Exception
                ViewInLog(ex, Me.Name(), Nothing)
            End Try
            '------------------------------------------------------------
            If oWystapilConcurrencyException = True Then
                dbDataAdapter.Fill(DataSetTabela)
            End If
            dbConnection.Close()
        End If
    End Sub


    '********************************************************************
    '*** testowanie wyst¹pienia b³êdu wspóbie¿noœci
    '*******************************************************************

    'Private Shared Sub OnRowUpdated(ByVal sender As Object, ByVal args As OleDb.OleDbRowUpdatedEventArgs)
    '    Dim argsKomenda As String = ""
    '    'sprawdzamy liczbe aktualizowanych rekordow
    '    If args.RecordsAffected = 0 Then
    '        Select Case args.StatementType
    '            Case StatementType.Delete
    '                argsKomenda = "DELETE"
    '            Case StatementType.Insert
    '                argsKomenda = "INSERT"
    '            Case StatementType.Select
    '                argsKomenda = "SELECT"
    '            Case StatementType.Update
    '                argsKomenda = "UPDATE"
    '        End Select

    '        Dim aTXT As String
    '        aTXT = args.Command.ToString
    '        aTXT = args.Errors.ToString
    '        aTXT = args.Row.RowError.ToString
    '        aTXT = args.Status.ToString
    '        aTXT = args.TableMapping.ToString

    '        'opis bledu = args.Errors.ToString()
    '        MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf & _
    '                                    "wyst¹pi³ b³¹d wspólnego dostêpu do bazy danych..." & vbCrLf & _
    '                                    "Aktualizacje rekordu " & args.Row(0) & " zostan¹ utracone" & vbCrLf & _
    '                                    "i pobrany bêdzie bie¿¹cy rekord z bazy danych..." & vbCrLf & vbCrLf & _
    '                                    "Proszê powtórzyæ aktualizacje i spróbowaæ ponownie zapisaæ rekord do bazy.", _
    '            "B³¹d podczas aktualizacji bazy :", _
    '            System.Windows.Forms.MessageBoxButtons.OK, _
    '            MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

    '        'args.Row.RowError ="Optimistic Concurrency Violation Encountered"
    '        args.Status = UpdateStatus.SkipCurrentRow

    '        oWystapilConcurrencyException = True
    '    End If
    'End Sub
End Class
