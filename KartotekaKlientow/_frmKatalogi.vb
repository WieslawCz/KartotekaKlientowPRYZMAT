Public Class _frmKatalogi
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblPlik As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents ListView As System.Windows.Forms.ListView
    Friend WithEvents NazwaPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents TypPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents DlugoscPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents UnikalnoscPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents ROPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents AllowNULL As System.Windows.Forms.ColumnHeader
    Friend WithEvents Caption As System.Windows.Forms.ColumnHeader
    Friend WithEvents ListTabela As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblNazwaTabeli As System.Windows.Forms.Label
    Friend WithEvents cmdPokazZawartosc As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(_frmKatalogi))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblNazwaTabeli = New System.Windows.Forms.Label()
        Me.ListTabela = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ListView = New System.Windows.Forms.ListView()
        Me.NazwaPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.TypPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.DlugoscPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.UnikalnoscPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ROPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.AllowNULL = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Caption = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblPlik = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.cmdPokazZawartosc = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblNazwaTabeli)
        Me.Panel1.Controls.Add(Me.ListTabela)
        Me.Panel1.Controls.Add(Me.ListView)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.lblPlik)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(896, 296)
        Me.Panel1.TabIndex = 0
        '
        'lblNazwaTabeli
        '
        Me.lblNazwaTabeli.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaTabeli.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaTabeli.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaTabeli.ForeColor = System.Drawing.Color.Navy
        Me.lblNazwaTabeli.Location = New System.Drawing.Point(372, 32)
        Me.lblNazwaTabeli.Name = "lblNazwaTabeli"
        Me.lblNazwaTabeli.Size = New System.Drawing.Size(508, 20)
        Me.lblNazwaTabeli.TabIndex = 7
        Me.lblNazwaTabeli.Text = "NazwaTabeli"
        Me.lblNazwaTabeli.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ListTabela
        '
        Me.ListTabela.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ListTabela.BackColor = System.Drawing.SystemColors.Control
        Me.ListTabela.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.ListTabela.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.ListTabela.ForeColor = System.Drawing.Color.Purple
        Me.ListTabela.Location = New System.Drawing.Point(60, 32)
        Me.ListTabela.Name = "ListTabela"
        Me.ListTabela.Size = New System.Drawing.Size(306, 251)
        Me.ListTabela.TabIndex = 6
        Me.ListTabela.UseCompatibleStateImageBehavior = False
        Me.ListTabela.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Nazwa Tabeli"
        Me.ColumnHeader1.Width = 182
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Iloœæ rek."
        Me.ColumnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader2.Width = 90
        '
        'ListView
        '
        Me.ListView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView.BackColor = System.Drawing.SystemColors.Control
        Me.ListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.NazwaPola, Me.TypPola, Me.DlugoscPola, Me.UnikalnoscPola, Me.ROPola, Me.AllowNULL, Me.Caption})
        Me.ListView.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.ListView.ForeColor = System.Drawing.Color.Purple
        Me.ListView.Location = New System.Drawing.Point(372, 52)
        Me.ListView.Name = "ListView"
        Me.ListView.Size = New System.Drawing.Size(508, 231)
        Me.ListView.TabIndex = 5
        Me.ListView.UseCompatibleStateImageBehavior = False
        Me.ListView.View = System.Windows.Forms.View.Details
        '
        'NazwaPola
        '
        Me.NazwaPola.Text = "Nazwa Pola"
        Me.NazwaPola.Width = 120
        '
        'TypPola
        '
        Me.TypPola.Text = "Typ"
        '
        'DlugoscPola
        '
        Me.DlugoscPola.Text = "D³ugoœæ"
        '
        'UnikalnoscPola
        '
        Me.UnikalnoscPola.Text = "Uni"
        Me.UnikalnoscPola.Width = 30
        '
        'ROPola
        '
        Me.ROPola.Text = "R/O"
        Me.ROPola.Width = 30
        '
        'AllowNULL
        '
        Me.AllowNULL.Text = "Null"
        Me.AllowNULL.Width = 30
        '
        'Caption
        '
        Me.Caption.Text = "Opis pola"
        Me.Caption.Width = 200
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(16, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Tabele :"
        '
        'lblPlik
        '
        Me.lblPlik.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPlik.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPlik.ForeColor = System.Drawing.Color.Navy
        Me.lblPlik.Location = New System.Drawing.Point(88, 8)
        Me.lblPlik.Name = "lblPlik"
        Me.lblPlik.Size = New System.Drawing.Size(792, 16)
        Me.lblPlik.TabIndex = 1
        Me.lblPlik.Text = "Plik na dysku"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Plik na dysku"
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(824, 312)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 14
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 311)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 16
        Me.picLogo.TabStop = False
        '
        'cmdPokazZawartosc
        '
        Me.cmdPokazZawartosc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazZawartosc.Image = CType(resources.GetObject("cmdPokazZawartosc.Image"), System.Drawing.Image)
        Me.cmdPokazZawartosc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazZawartosc.Location = New System.Drawing.Point(650, 312)
        Me.cmdPokazZawartosc.Name = "cmdPokazZawartosc"
        Me.cmdPokazZawartosc.Size = New System.Drawing.Size(159, 23)
        Me.cmdPokazZawartosc.TabIndex = 17
        Me.cmdPokazZawartosc.Text = "Poka¿ zawartoœæ tabeli"
        Me.cmdPokazZawartosc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        '_frmKatalogi
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(914, 340)
        Me.Controls.Add(Me.cmdPokazZawartosc)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "_frmKatalogi"
        Me.Text = "Baza Danych programu"
        Me.Panel1.ResumeLayout(False)
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Dim sConnection As String = DataBaseConnection()
    Dim sSelectSQL As String = "SELECT * FROM " & NazwaTabeli

    'Dim dbConnection As OleDb.OleDbConnection
    'Dim dbCommandSelect As OleDb.OleDbCommand
    'Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As New SqlClient.SqlConnection
    Dim sqlCommandSelect As New SqlClient.SqlCommand
    Dim sqlDataAdapter As New SqlClient.SqlDataAdapter

    Dim DataSetTabela As New System.Data.DataSet
    Dim DataTableTabela As System.Data.DataTable
    Dim DataViewTabela As New System.Data.DataView

    Dim ConnectionState As System.Data.ConnectionState

    Dim NazwaTabeli As String


    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub _frmKatalogi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.MrJones
        Me.picLogo.Image = My.Resources.logomini
        Me.cmdPowrot.Image = My.Resources._EXIT.ToBitmap
        Me.cmdPokazZawartosc.Image = My.Resources.Tabela.ToBitmap
        '-------------------------------
        Dim ir As Long = 0
        Dim it As Integer = 0
        Dim nTab As String = ""
        Dim rTab As Long = 0
        Dim wTab(1) As String
        'Me.Text = oKomunikat


        If ParL_DataBaseType = "MS ACCESS" Then
            'Label1.Text = "Plik Bazy Danych . . . . ."
            'lblPlik.Text = ParL_PlikMDB

            'Dim AccessConnection As New System.Data.OleDb.OleDbConnection()
            'Dim SchemaTable As System.Data.DataTable

            'AccessConnection.ConnectionString = sConnection
            'AccessConnection.Open()

            ''Retrieve schema information about Table1.
            ''SchemaTable = AccessConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, _
            ''                New Object() {Nothing, Nothing, "TABLE"})

            'SchemaTable = AccessConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, _
            '                New Object() {Nothing, Nothing, Nothing, "TABLE"})

            'If SchemaTable.Rows.Count <> 0 Then
            '    'Console.WriteLine("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " Exists")

            '    ListTabela.ForeColor = System.Drawing.Color.Navy
            '    it = 0
            '    For ir = 0 To SchemaTable.Rows.Count - 1
            '        If SchemaTable.Rows(ir)!TABLE_TYPE.ToString = "TABLE" Then
            '            nTab = SchemaTable.Rows(ir)!TABLE_NAME.ToString
            '            rTab = IloscRekordowWTabeli(nTab)

            '            'wTab(0) = nTab
            '            'wTab(1) = Trim(Str(rTab))
            '            'lbxTabele.Items.Add(wTab)

            '            ListTabela.Items.Add(nTab, it)
            '            ListTabela.Items(it).SubItems.Add(Trim(Str(rTab)))

            '            If 2 * Int(it / 2) = it Then
            '                'parzyste
            '                ListTabela.Items(it).BackColor = System.Drawing.SystemColors.ControlLight
            '            Else
            '                'nieparzyste
            '                ListTabela.Items(it).BackColor = System.Drawing.SystemColors.Control
            '            End If

            '            it += 1

            '        End If
            '    Next
            'End If
            'AccessConnection.Close()
        Else
            Label1.Text = "Baza SQL . . . . ."
            lblPlik.Text = ParL_DataBase

            Dim sSelectSQL As String = "SELECT * FROM sysobjects " & _
                                       "WHERE xtype='U'" & _
                                       "ORDER BY name"
            Dim sConnection As String = DataBaseConnection()

            'Dim dbConnection As OleDb.OleDbConnection = Nothing
            'Dim dbCommandSelect As OleDb.OleDbCommand = Nothing
            'Dim dbDataAdapter As OleDb.OleDbDataAdapter = Nothing

            Dim sqlConnection As SqlClient.SqlConnection = Nothing
            Dim sqlCommandSelect As SqlClient.SqlCommand = Nothing
            Dim sqlDataAdapter As SqlClient.SqlDataAdapter = Nothing

            Dim DataSetTabela As New DataSet
            Dim DataViewTabela As New DataView
            Dim ConnectionState As System.Data.ConnectionState

            sqlConnection = New SqlClient.SqlConnection(sConnection)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQL, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnection.Open()
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnection.State
            End Try

            If ConnectionState = ConnectionState.Open Then
                Try
                    sqlDataAdapter.FillSchema(DataSetTabela, SchemaType.Mapped, "TABELA")
                    sqlDataAdapter.Fill(DataSetTabela)
                    sqlConnection.Close()

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    'DataSetTabela.Tables(0).PrimaryKey = New DataColumn() {DataSetTabela.Tables(0).Columns("DATANOTOW"), _
                    '---------------------------------

                    'definiuj DataView
                    DataViewTabela = New DataView(DataSetTabela.Tables(0))
                    DataViewTabela.AllowNew = False
                    DataViewTabela.AllowDelete = False

                    If DataViewTabela.Count > 0 Then
                        ListTabela.ForeColor = System.Drawing.Color.Navy
                        it = 0
                        For ir = 0 To DataViewTabela.Count - 1
                            'lbxTabele.Items.Add(DataViewTabela.Item(ir).Item("NAME"))
                            nTab = DataViewTabela.Item(ir).Item("NAME")
                            rTab = IloscRekordowWTabeli(nTab)

                            'wTab(0) = nTab
                            'wTab(1) = Trim(Str(rTab))
                            'lbxTabele.Items.Add(wTab)

                            ListTabela.Items.Add(nTab, it)
                            ListTabela.Items(it).SubItems.Add(Trim(Str(rTab)))

                            If 2 * Int(it / 2) = it Then
                                'parzyste
                                ListTabela.Items(it).BackColor = System.Drawing.SystemColors.ControlLight
                            Else
                                'nieparzyste
                                ListTabela.Items(it).BackColor = System.Drawing.SystemColors.Control
                            End If

                            it += 1
                        Next
                    End If

                Catch Ex As System.Exception
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                End Try
            End If
        End If
        ListTabela.Items(0).Selected = True

        'zawartosc tylko dla wybranych
        If _MaUprawnieniaAdministratora() Or _MaUprawnieniaSzefa() Then
            cmdPokazZawartosc.Enabled = True
        Else
            cmdPokazZawartosc.Enabled = False
        End If
    End Sub

    Private Sub lbxTabele_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PokazTabelePodKursorem()
    End Sub


    Private Sub ListTabela_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListTabela.SelectedIndexChanged
        PokazTabelePodKursorem()
    End Sub

    Private Sub cmdPokazZawartosc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokazZawartosc.Click
        'NazwaTabeli = lbxTabele.SelectedItem.ToString
        If ListTabela.SelectedItems.Count > 0 Then
            NazwaTabeli = ListTabela.SelectedItems(0).SubItems(0).Text
            lblNazwaTabeli.Text = NazwaTabeli
            If Len(Trim(NazwaTabeli)) > 0 Then
                oTabela = NazwaTabeli
                '----------------------------------------------
                Dim Form As New _frmKatalogiZawartosc("")
                Using Form
                    Form.ShowDialog()
                End Using
                Form = Nothing
                '----------------------------------------------
            End If
        Else
            lblNazwaTabeli.Text = ""
        End If
    End Sub

    Private Sub _frmKatalogi_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    '****************************
    ' pokazuje strukture tabeli pod kursorem
    '****************************

    Private Sub PokazTabelePodKursorem()
        DataSetTabela = New DataSet
        'NazwaTabeli = lbxTabele.SelectedItem.ToString
        If ListTabela.SelectedItems.Count > 0 Then
            NazwaTabeli = ListTabela.SelectedItems(0).SubItems(0).Text
            lblNazwaTabeli.Text = NazwaTabeli
            If ParL_DataBaseType = "MS ACCESS" Then
                'dbConnection = New OleDb.OleDbConnection(sConnection)
                'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQL & NazwaTabeli, dbConnection)
                'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)
                ''------------------------------------------
                ''wypelnij DATASET
                'Try
                '    dbConnection.Open()
                'Catch Ex As System.Exception
                '    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                '        System.Windows.Forms.MessageBoxButtons.OK, _
                '        MessageBoxIcon.Information, _
                '        MessageBoxDefaultButton.Button1)
                'Finally
                '    ConnectionState = dbConnection.State
                'End Try
            Else
                sqlConnection = New SqlClient.SqlConnection(sConnection)
                sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQL & NazwaTabeli, sqlConnection)
                sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)
                '------------------------------------------
                'wypelnij DATASET
                Try
                    sqlConnection.Open()
                Catch Ex As System.Exception
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
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

                    DataTableTabela = DataSetTabela.Tables("TABELA")
                    Dim objColumn As DataColumn
                    Dim i As Integer = 0

                    'usuwamy poprzednie zapisy na liscie
                    If ListView.Items.Count > 0 Then
                        For i = ListView.Items.Count - 1 To 0 Step -1
                            ListView.Items.RemoveAt(i)
                        Next
                    End If

                    'zapisujemy liste informacjami o strukturze tabeli
                    i = 0
                    ListView.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(192, Byte))
                    For Each objColumn In DataTableTabela.Columns
                        ListView.Items.Add(objColumn.ColumnName, i)
                        ListView.Items(i).SubItems.Add(objColumn.DataType.Name)
                        ListView.Items(i).SubItems.Add(IIf(objColumn.MaxLength < 0, "", objColumn.MaxLength.ToString))
                        ListView.Items(i).SubItems.Add(IIf(objColumn.Unique, "T", "N"))
                        ListView.Items(i).SubItems.Add(IIf(objColumn.ReadOnly, "T", "N"))
                        ListView.Items(i).SubItems.Add(IIf(objColumn.AllowDBNull, "T", "N"))
                        ListView.Items(i).SubItems.Add(objColumn.Caption)
                        i += 1
                    Next

                Catch Ex As System.Exception
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                End Try
            End If
        Else
            lblNazwaTabeli.Text = ""

            Dim j As Integer = 0
            'usuwamy poprzednie zapisy na liscie
            If ListView.Items.Count > 0 Then
                For j = ListView.Items.Count - 1 To 0 Step -1
                    ListView.Items.RemoveAt(j)
                Next
            End If
        End If
    End Sub



    '**************************
    ' pokazuje ilosc rekordow (zapisow) w tabeli
    '***************************

    Private Function IloscRekordowWTabeli(ByVal pNazwaTabeli As String) As Long
        Dim RetIleRec As Long = 0
        Dim sConn As String = DataBaseConnection()
        Dim dsTabela As New DataSet
        Dim ConnState As System.Data.ConnectionState

        'Dim dbConn As OleDb.OleDbConnection = Nothing
        'Dim dbComm As OleDb.OleDbCommand = Nothing
        'Dim dbAdapter As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConn As SqlClient.SqlConnection = Nothing
        Dim sqlComm As SqlClient.SqlCommand = Nothing
        Dim sqlAdapter As SqlClient.SqlDataAdapter = Nothing

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConn = New OleDb.OleDbConnection(sConn)
            'dbComm = New OleDb.OleDbCommand("SELECT count(*) AS ILOSCREC FROM " & pNazwaTabeli, dbConn)
            'dbAdapter = New OleDb.OleDbDataAdapter(dbComm)

            ''mapujemy domyslna nazwe tabeli
            'Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
            'DBMapowanieTabeli = dbAdapter.TableMappings.Add("Table", "TABELA")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConn.Open()
            'Catch Ex As System.Exception
            '    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
            '        System.Windows.Forms.MessageBoxButtons.OK, _
            '        MessageBoxIcon.Information, _
            '        MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnState = dbConn.State
            'End Try
        Else
            sqlConn = New SqlClient.SqlConnection(sConn)
            sqlComm = New SqlClient.SqlCommand("SELECT count(*) AS 'ILOSCREC' FROM " & pNazwaTabeli, sqlConn)
            sqlAdapter = New SqlClient.SqlDataAdapter(sqlComm)

            'mapujemy domyslna nazwe tabeli
            Dim SQLMapowanieTabeli As System.Data.Common.DataTableMapping
            SQLMapowanieTabeli = sqlAdapter.TableMappings.Add("Table", "TABELA")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConn.Open()
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Finally
                ConnState = sqlConn.State
            End Try
        End If

        If ConnState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbAdapter.FillSchema(dsTabela, SchemaType.Mapped, "TABELA")
                    'dbAdapter.Fill(dsTabela)
                    'dbConn.Close()
                Else
                    sqlAdapter.FillSchema(dsTabela, SchemaType.Mapped, "TABELA")
                    sqlAdapter.Fill(dsTabela)
                    sqlConn.Close()
                End If

                If dsTabela.Tables(0).Rows.Count > 0 Then
                    RetIleRec = GetNumDRField(dsTabela.Tables(0).Rows.Item(0), "ILOSCREC")
                End If

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End Try
        End If
        Return (RetIleRec)
    End Function


End Class
