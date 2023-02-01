Public Class KatalogiZawartosc
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
    Friend WithEvents dagTabela As System.Windows.Forms.DataGrid
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents stbTabela As System.Windows.Forms.StatusBar
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents cbxRok As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbxMiesiac As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pnlClear As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogiZawartosc))
        Me.dagTabela = New System.Windows.Forms.DataGrid()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.stbTabela = New System.Windows.Forms.StatusBar()
        Me.pnlClear = New System.Windows.Forms.Panel()
        Me.cbxMiesiac = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxRok = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdClear = New System.Windows.Forms.Button()
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
        Me.dagTabela.Size = New System.Drawing.Size(777, 343)
        Me.dagTabela.TabIndex = 0
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(689, 391)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 22)
        Me.cmdStop.TabIndex = 15
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(16, 391)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 23)
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
        Me.stbTabela.Location = New System.Drawing.Point(8, 351)
        Me.stbTabela.Name = "stbTabela"
        Me.stbTabela.ShowPanels = True
        Me.stbTabela.Size = New System.Drawing.Size(777, 21)
        Me.stbTabela.TabIndex = 21
        Me.stbTabela.Text = "StatusBar1"
        '
        'pnlClear
        '
        Me.pnlClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlClear.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlClear.Controls.Add(Me.cbxMiesiac)
        Me.pnlClear.Controls.Add(Me.Label2)
        Me.pnlClear.Controls.Add(Me.cbxRok)
        Me.pnlClear.Controls.Add(Me.Label1)
        Me.pnlClear.Controls.Add(Me.cmdClear)
        Me.pnlClear.Location = New System.Drawing.Point(289, 383)
        Me.pnlClear.Name = "pnlClear"
        Me.pnlClear.Size = New System.Drawing.Size(392, 40)
        Me.pnlClear.TabIndex = 69
        '
        'cbxMiesiac
        '
        Me.cbxMiesiac.Location = New System.Drawing.Point(192, 8)
        Me.cbxMiesiac.Name = "cbxMiesiac"
        Me.cbxMiesiac.Size = New System.Drawing.Size(88, 21)
        Me.cbxMiesiac.TabIndex = 72
        Me.cbxMiesiac.Text = "<wszystkie>"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(144, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 73
        Me.Label2.Text = "Miesi¹c . . . . . . . . . . ."
        '
        'cbxRok
        '
        Me.cbxRok.Location = New System.Drawing.Point(40, 8)
        Me.cbxRok.Name = "cbxRok"
        Me.cbxRok.Size = New System.Drawing.Size(88, 21)
        Me.cbxRok.TabIndex = 71
        Me.cbxRok.Text = "<wszystkie>"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 70
        Me.Label1.Text = "Rok . . . . . . . . . . ."
        '
        'cmdClear
        '
        Me.cmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClear.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdClear.ForeColor = System.Drawing.Color.Black
        Me.cmdClear.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClear.Location = New System.Drawing.Point(296, 8)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(80, 23)
        Me.cmdClear.TabIndex = 69
        Me.cmdClear.Text = "Czyœæ obroty"
        Me.cmdClear.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'KatalogiZawartosc
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(794, 429)
        Me.Controls.Add(Me.pnlClear)
        Me.Controls.Add(Me.stbTabela)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.dagTabela)
        Me.Name = "KatalogiZawartosc"
        Me.ShowInTaskbar = False
        Me.Text = "Katalog zawartosc"
        CType(Me.dagTabela, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlClear.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim sConnection As String = ConnectionStrings()

    Dim sSelectSQL As String
    Dim sDeleteSQL As String

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbCommandDelete As OleDb.OleDbCommand

    Dim dbDataAdapter As OleDb.OleDbDataAdapter
    Dim DataSetTabela As New DataSet
    Dim MapowanieTabeli As System.Data.Common.DataTableMapping

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub KatalogiZawartosc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        sSelectSQL = "SELECT * FROM " & oTabela
        Select Case oTabela
            Case "OBROTY"
                sDeleteSQL = "DELETE FROM Obroty " & _
                             "WHERE (IDENTKLIENTA=@orygSymbol) AND " & _
                                   "(DATA=@orygData) AND " & _
                                   "(WERSJA=@orygWersja)"
            Case "OBROTYMIES"
                sDeleteSQL = "DELETE FROM ObrotyMies " & _
                             "WHERE (IDENTKLIENTA=@orygSymbol) AND " & _
                                   "(MIESIAC=@orygMies) AND " & _
                                   "(WERSJA=@orygWersja)"
            Case Else
        End Select
        dbConnection = New OleDb.OleDbConnection(sConnection)
        dbCommandSelect = New OleDb.OleDbCommand(sSelectSQL, dbConnection)
        dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQL, dbConnection)
        dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)


        MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA")
        DataSetTabela.Locale = New System.Globalization.CultureInfo("pl-PL")
        'komendy DataAdapter
        Dim objParam As OleDb.OleDbParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        Select Case oTabela
            Case "OBROTY"
                objParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
                objParam.SourceVersion = DataRowVersion.Original
                objParam = dbCommandDelete.Parameters.Add("@orygData", OleDb.OleDbType.Char, 10, "DATA")
                objParam.SourceVersion = DataRowVersion.Original
                objParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                objParam.SourceVersion = DataRowVersion.Original
                dbDataAdapter.DeleteCommand = dbCommandDelete
            Case "OBROTYMIES"
                objParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
                objParam.SourceVersion = DataRowVersion.Original
                objParam = dbCommandDelete.Parameters.Add("@orygMies", OleDb.OleDbType.Char, 7, "MIESIAC")
                objParam.SourceVersion = DataRowVersion.Original
                objParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                objParam.SourceVersion = DataRowVersion.Original
                dbDataAdapter.DeleteCommand = dbCommandDelete
            Case Else
        End Select
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

        Select Case oTabela
            Case "OBROTY"
                pnlClear.Visible = True
            Case "OBROTYMIES"
                pnlClear.Visible = True
            Case Else
                pnlClear.Visible = False
        End Select

        'Lista lat
        Dim Ir As Integer

        cbxRok.Items.Add(Trim("<wszystkie>"))
        For Ir = 1900 To 2100
            cbxRok.Items.Add(Trim(Str(Ir)))
        Next

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

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        If MessageBox.Show("Czy na pewno mam usun¹æ informacje o obrotach ?", _
                    "Proszê o potwierdzenie :", _
                    System.Windows.Forms.MessageBoxButtons.YesNo, _
                    MessageBoxIcon.Question, _
                    MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

            Dim dr As DataRow
            Dim rok, mies As String
            Dim oMiesiac As String
            If cbxRok.Text = "<wszystkie>" Then
                If cbxMiesiac.Text = "<wszystkie>" Then
                    'USUN WSZYSTKO....
                    For Each dr In DataSetTabela.Tables(0).Rows
                        dr.Delete()
                    Next
                Else
                    'USUN MIESIAC Z WSZYSTKICH LAT
                    For Each dr In DataSetTabela.Tables(0).Rows
                        If oTabela = "OBROTYMIES" Then
                            oMiesiac = dr.Item("MIESIAC")
                        Else
                            oMiesiac = dr.Item("DATA")
                        End If
                        rok = Mid(oMiesiac, 1, 4)
                        mies = Mid(oMiesiac, 6, 2)
                        If Mid(cbxMiesiac.Text, 1, 2) = mies Then
                            dr.Delete()
                        End If
                    Next
                End If
            Else
                If cbxMiesiac.Text = "<wszystkie>" Then
                    'USUN WSZYSTKIE MIESIACE WYBRANEGO ROKU
                    For Each dr In DataSetTabela.Tables(0).Rows
                        If oTabela = "OBROTYMIES" Then
                            oMiesiac = dr.Item("MIESIAC")
                        Else
                            oMiesiac = dr.Item("DATA")
                        End If
                        rok = Mid(oMiesiac, 1, 4)
                        mies = Mid(oMiesiac, 6, 2)
                        If cbxRok.Text = rok Then
                            dr.Delete()
                        End If
                    Next
                Else
                    'USUN WYBRANY ROK I MIESIAC
                    For Each dr In DataSetTabela.Tables(0).Rows
                        If oTabela = "OBROTYMIES" Then
                            oMiesiac = dr.Item("MIESIAC")
                        Else
                            oMiesiac = dr.Item("DATA")
                        End If
                        rok = Mid(oMiesiac, 1, 4)
                        mies = Mid(oMiesiac, 6, 2)
                        If (cbxRok.Text = rok) And (Mid(cbxMiesiac.Text, 1, 2) = mies) Then
                            dr.Delete()
                        End If
                    Next
                End If
            End If
            AktualizujBaze()
            'DataSetTabela.Clear()
            stbTabela.Panels(0).Text = "Iloœæ rec.: " & DataSetTabela.Tables(0).Rows.Count.ToString
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

    Private Shared Sub OnRowUpdated(ByVal sender As Object, ByVal args As OleDb.OleDbRowUpdatedEventArgs)
        Dim argsKomenda As String = ""
        'sprawdzamy liczbe aktualizowanych rekordow
        If args.RecordsAffected = 0 Then
            Select Case args.StatementType
                Case StatementType.Delete
                    argsKomenda = "DELETE"
                Case StatementType.Insert
                    argsKomenda = "INSERT"
                Case StatementType.Select
                    argsKomenda = "SELECT"
                Case StatementType.Update
                    argsKomenda = "UPDATE"
            End Select

            Dim aTXT As String
            aTXT = args.Command.ToString
            aTXT = args.Errors.ToString
            aTXT = args.Row.RowError.ToString
            aTXT = args.Status.ToString
            aTXT = args.TableMapping.ToString

            'opis bledu = args.Errors.ToString()
            MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf & _
                                        "wyst¹pi³ b³¹d wspólnego dostêpu do bazy danych..." & vbCrLf & _
                                        "Aktualizacje rekordu " & args.Row(0) & " zostan¹ utracone" & vbCrLf & _
                                        "i pobrany bêdzie bie¿¹cy rekord z bazy danych..." & vbCrLf & vbCrLf & _
                                        "Proszê powtórzyæ aktualizacje i spróbowaæ ponownie zapisaæ rekord do bazy.", _
                "B³¹d podczas aktualizacji bazy :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

            'args.Row.RowError ="Optimistic Concurrency Violation Encountered"
            args.Status = UpdateStatus.SkipCurrentRow

            oWystapilConcurrencyException = True
        End If
    End Sub

    Private Sub KatalogiZawartosc_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

End Class
