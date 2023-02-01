Public Class TestujPolaczenieZBazaDanych
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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents txtRaport As System.Windows.Forms.TextBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TestujPolaczenieZBazaDanych))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtRaport = New System.Windows.Forms.TextBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.txtRaport)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(504, 271)
        Me.Panel1.TabIndex = 15
        '
        'txtRaport
        '
        Me.txtRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport.BackColor = System.Drawing.SystemColors.Control
        Me.txtRaport.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport.Location = New System.Drawing.Point(8, 8)
        Me.txtRaport.Multiline = True
        Me.txtRaport.Name = "txtRaport"
        Me.txtRaport.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtRaport.Size = New System.Drawing.Size(488, 256)
        Me.txtRaport.TabIndex = 23
        Me.txtRaport.Text = "TextBox1"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 287)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 39
        Me.PictureBox1.TabStop = False
        '
        'Timer1
        '
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(424, 287)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 24)
        Me.cmdStop.TabIndex = 56
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TestujPolaczenieZBazaDanych
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(520, 318)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "TestujPolaczenieZBazaDanych"
        Me.ShowInTaskbar = False
        Me.Text = "SQL Server"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    ' lista serverow
    '
    ' informacje o serwerze
    '   exec xp_msver
    ' lista baz danych w SQL Serwerze
    '   exec sp_catalogs_rowset;2 NULL


    Private Sub TestujPolaczenieZBazaDanych_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        Timer1.Interval = 10
        Timer1.Enabled = True
    End Sub

    Private Sub TestujPolaczenieZBazaDanych_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim sConnectionStr As String
        If ParL_Autentykacja = "Windows" Then
            sConnectionStr = "Data Source=" & ParL_Serwer & ";" & _
                   "Integrated Security=SSPI;" & _
                   "Initial Catalog=" & ParL_DataBase & ""
        Else
            sConnectionStr = "Persist Security Info=False;" & _
                   "User ID=" & ParL_Login & ";" & _
                   "Password=" & ParL_Password & ";" & _
                   "Initial Catalog=" & ParL_DataBase & ";" & _
                   "Data Source=" & ParL_Serwer & ";" & _
                   "Packet Size=4096"
        End If




        'DbConnection-polaczenie
        Dim dbConnection As New SqlClient.SqlConnection(sConnectionStr)
        'Dim dbConnection As New OleDb.OleDbConnection(sConnectionStr)
        Dim dbCommand As SqlClient.SqlCommand = New SqlClient.SqlCommand
        'Dim dbCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand

        Dim Wynik As String

        Timer1.Enabled = False
        '--------------------------------------
        txtRaport.Text = ""
        txtRaport.ReadOnly = True
        txtRaport.SelectionStart = Len(txtRaport.Text)
        txtRaport.SelectionLength = 0
        txtRaport.Focus()
        System.Windows.Forms.Application.DoEvents()
        txtRaport.ScrollToCaret()

        dbCommand.Connection = dbConnection
        dbCommand.CommandType = CommandType.Text


        Try
            dbConnection.Open()
            '------------------------------------------------
            dbCommand.CommandText = "SELECT @@SERVERNAME"
            Wynik = dbCommand.ExecuteScalar()
            txtRaport.Text &= "NAZWA SQL SERWERA : " & Wynik & vbCrLf & vbCrLf
            txtRaport.SelectionStart = Len(txtRaport.Text)
            txtRaport.SelectionLength = 0
            System.Windows.Forms.Application.DoEvents()
            txtRaport.ScrollToCaret()
            '------------------------------------------------
            dbCommand.CommandText = "SELECT @@SERVICENAME"
            Wynik = dbCommand.ExecuteScalar()
            txtRaport.Text &= "INSTANCJA : " & Wynik & vbCrLf & vbCrLf
            txtRaport.SelectionStart = Len(txtRaport.Text)
            txtRaport.SelectionLength = 0
            System.Windows.Forms.Application.DoEvents()
            txtRaport.ScrollToCaret()
            '------------------------------------------------
            dbCommand.CommandText = "SELECT @@VERSION"
            Wynik = dbCommand.ExecuteScalar()
            txtRaport.Text &= "WERSJA : "

            Dim iPos As Integer
            Dim Wiersz As String

            iPos = InStr(Wynik, vbLf)
            Do While iPos > 0
                Wiersz = Mid(Wynik, 1, iPos - 1)
                txtRaport.Text &= Wiersz & vbCrLf
                Wynik = Mid(Wynik, iPos + 1)
                iPos = InStr(Wynik, vbLf)
            Loop
            txtRaport.Text &= Wynik & vbCrLf
            txtRaport.SelectionStart = Len(txtRaport.Text)
            txtRaport.SelectionLength = 0
            System.Windows.Forms.Application.DoEvents()
            txtRaport.ScrollToCaret()
            '------------------------------------------------
            dbCommand.CommandText = "SELECT @@CONNECTIONS"
            Wynik = dbCommand.ExecuteScalar()
            txtRaport.Text &= "ILOŒÆ PO£¥CZEÑ : " & Wynik & vbCrLf & vbCrLf
            txtRaport.SelectionStart = Len(txtRaport.Text)
            txtRaport.SelectionLength = 0
            System.Windows.Forms.Application.DoEvents()
            txtRaport.ScrollToCaret()


            '------------------------------------------------
            Dim sSelectSQL As String = "exec sp_catalogs_rowset;2 NULL"
            'Dim dbConnection As New SqlClient.SqlConnection(sConnectionStr)
            Dim dbCommandSelect As New SqlClient.SqlCommand(sSelectSQL, dbConnection)
            Dim dbDataAdapter As New SqlClient.SqlDataAdapter(dbCommandSelect)
            Dim dbDataSet As New DataSet
            Dim dbDataView As New DataView
            dbDataSet.Locale = New System.Globalization.CultureInfo("pl-PL")
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Database")
            'wypelnij DATASET

            Try
                dbConnection.Open()
            Catch Ex As System.Exception
            End Try

            txtRaport.Text &= "BAZY DANYCH : " & vbCrLf & vbTab
            If dbConnection.State = ConnectionState.Open Then
                Try
                    dbDataAdapter.FillSchema(dbDataSet, SchemaType.Mapped, "TABELA_Database")
                    dbDataAdapter.Fill(dbDataSet)
                    dbConnection.Close()

                    'definiuj dbDataView
                    dbDataView = New DataView(dbDataSet.Tables(0))
                    If dbDataView.Count > 0 Then
                        Dim dr As DataRow
                        Dim db As String
                        For Each dr In dbDataView.Table.Rows
                            db = dr.Item("CATALOG_NAME")
                            'txtRaport.Text &= vbTab & db & "  " & SprawdzBazeDanych(db) & vbCrLf & vbTab
                            txtRaport.Text &= db & "  " & SprawdzBazeDanych(db) & vbCrLf & vbTab
                        Next
                    End If
                Catch Ex As System.Exception
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                End Try
            End If
            '------------------------------------------------


        Catch ex As Exception
            MessageBox.Show("B³¹d po³¹czenia z baz¹ danych SQL" & vbCrLf & "na serwerze " & ParL_Serwer & "...", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub



    '********************************************************************
    '*** sprawdza czy Baza Danych z tego programu...
    '********************************************************************

    Dim qConnectionString As String = ""
    Dim qSelectSQL As String = "SELECT * FROM sysobjects " & _
                               "WHERE xtype='P'" & _
                               "ORDER BY name"
    Dim qConnection As SqlClient.SqlConnection
    Dim qCommandSelect As SqlClient.SqlCommand
    Dim qDataAdapter As SqlClient.SqlDataAdapter
    Dim qDataSetTabela As DataSet
    Dim qDataViewTabela As DataView

    Private Function SprawdzBazeDanych(ByVal NazwaDB As String) As String
        Dim ir As Long = 0
        Dim OK As Boolean = False
        Dim Odpowiedz As String = ""

        'pomin w analizie bazy systemowe....
        If NazwaDB = "master" Or _
           NazwaDB = "model" Or _
           NazwaDB = "msdb" Or _
           NazwaDB = "tempdb" Then
        Else
            If ParL_Autentykacja = "Windows" Then
                qConnectionString = "Data Source=" & ParL_Serwer & ";" & _
                       "Integrated Security=SSPI;" & _
                       "Initial Catalog=" & NazwaDB & ""
            Else
                qConnectionString = "Persist Security Info=False;" & _
                       "User ID=" & ParL_Login & ";" & _
                       "Password=" & ParL_Password & ";" & _
                       "Initial Catalog=" & NazwaDB & ";" & _
                       "Data Source=" & ParL_Serwer & ";" & _
                       "Packet Size=4096"
            End If

            qSelectSQL = "SELECT * FROM sysobjects " & _
                         "WHERE xtype='P'" & _
                         "ORDER BY name"

            qConnection = New SqlClient.SqlConnection(qConnectionString)
            qCommandSelect = New SqlClient.SqlCommand(qSelectSQL, qConnection)
            qDataAdapter = New SqlClient.SqlDataAdapter(qCommandSelect)
            qDataSetTabela = New DataSet
            qDataViewTabela = New DataView

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = qDataAdapter.TableMappings.Add("Table", "TABELA")
            qDataSetTabela.Locale = New System.Globalization.CultureInfo("pl-PL")
            '------------------------------------------
            'wypelnij DATASET
            Try
                qConnection.Open()
            Catch Ex As System.Exception
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If qConnection.State = ConnectionState.Open Then
                Try
                    qDataAdapter.FillSchema(qDataSetTabela, SchemaType.Mapped, "TABELA")
                    qDataAdapter.Fill(qDataSetTabela)
                    qConnection.Close()

                    'definiuj DataView
                    qDataViewTabela = New DataView(qDataSetTabela.Tables(0))
                    If qDataViewTabela.Count > 0 Then
                        'sprawdz czy to baza danych softart
                        For ir = 0 To qDataViewTabela.Count - 1
                            If qDataViewTabela.Item(ir).Item("NAME") = "softart_nazwaprogramu" Then
                                OK = True
                                Exit For
                            End If
                        Next
                    End If
                Catch Ex As System.Exception
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                End Try

                'sprawdz czy to baza danych programu SOFTART-MAGAZYN
                If OK Then  'to jest baza danych softart
                    Odpowiedz = " :--> " & NazwaProgramu(NazwaDB) & WersjaProgramu(NazwaDB)
                Else
                    Odpowiedz = ""
                End If
            End If
        End If
        Return Odpowiedz
    End Function

    Private Function NazwaProgramu(ByVal BazaD As String) As String
        Dim Wynik As String = ""

        Dim xConnectionString As String = ""
        If ParL_Autentykacja = "Windows" Then
            xConnectionString = "Data Source=" & ParL_Serwer & ";" & _
                   "Integrated Security=SSPI;" & _
                   "Initial Catalog=" & BazaD & ""
        Else
            xConnectionString = "Persist Security Info=False;" & _
                   "User ID=" & ParL_Login & ";" & _
                   "Password=" & ParL_Password & ";" & _
                   "Initial Catalog=" & BazaD & ";" & _
                   "Data Source=" & ParL_Serwer & ";" & _
                   "Packet Size=4096"
        End If

        Dim xStoredProc As String = "softart_nazwaprogramu"

        Dim xConnection As New SqlClient.SqlConnection(xConnectionString)
        Dim xCommand As New SqlClient.SqlCommand
        Dim xParam As SqlClient.SqlParameter

        xCommand.CommandText = xStoredProc  'xSelectSQL
        xCommand.CommandType = CommandType.StoredProcedure
        xParam = xCommand.Parameters.Add(New SqlClient.SqlParameter("@name", SqlDbType.VarChar, 50))
        xParam.Direction = ParameterDirection.Output
        xCommand.Connection = xConnection

        Try
            xConnection.Open()
            xCommand.ExecuteNonQuery()
            xConnection.Close()

            'podstaw parametry wyjsciowe
            Wynik = CType(xCommand.Parameters("@name").Value, String)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return Wynik
    End Function


    Private Function WersjaProgramu(ByVal BazaD As String) As String
        Dim Wersja As Int32 = 0
        Dim Wynik As String = ""

        Dim xConnectionString As String = ""
        If ParL_Autentykacja = "Windows" Then
            xConnectionString = "Data Source=" & ParL_Serwer & ";" & _
                   "Integrated Security=SSPI;" & _
                   "Initial Catalog=" & BazaD & ""
        Else
            xConnectionString = "Persist Security Info=False;" & _
                   "User ID=" & ParL_Login & ";" & _
                   "Password=" & ParL_Password & ";" & _
                   "Initial Catalog=" & BazaD & ";" & _
                   "Data Source=" & ParL_Serwer & ";" & _
                   "Packet Size=4096"
        End If

        Dim xStoredProc As String = "softart_wersjaprogramu"

        Dim xConnection As New SqlClient.SqlConnection(xConnectionString)
        Dim xCommand As New SqlClient.SqlCommand
        Dim xParam As SqlClient.SqlParameter

        xCommand.CommandText = xStoredProc  'xSelectSQL
        xCommand.CommandType = CommandType.StoredProcedure
        xParam = xCommand.Parameters.Add(New SqlClient.SqlParameter("@wersja", SqlDbType.Int, Nothing))
        xParam.Direction = ParameterDirection.Output
        xCommand.Connection = xConnection

        Try
            xConnection.Open()
            xCommand.ExecuteNonQuery()
            xConnection.Close()

            'podstaw parametry wyjsciowe
            Wersja = CType(xCommand.Parameters("@wersja").Value, Int32)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        If Wersja > 0 Then
            Wynik = " Wersja " & Trim((Wersja / 100).ToString("N2"))
        Else
            Wynik = ""
        End If

        Return Wynik
    End Function
End Class
