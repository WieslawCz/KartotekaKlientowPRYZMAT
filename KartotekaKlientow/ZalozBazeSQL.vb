Imports System.IO
'Imports System.Text

Public Class ZalozBazeSQL
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
    Friend WithEvents cmdZaloz As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtRaport As System.Windows.Forms.TextBox
    Friend WithEvents txtNazwaBazy As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ZalozBazeSQL))
        Me.cmdZaloz = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtNazwaBazy = New System.Windows.Forms.TextBox()
        Me.txtRaport = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdZaloz
        '
        Me.cmdZaloz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZaloz.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdZaloz.Image = CType(resources.GetObject("cmdZaloz.Image"), System.Drawing.Image)
        Me.cmdZaloz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZaloz.Location = New System.Drawing.Point(204, 378)
        Me.cmdZaloz.Name = "cmdZaloz"
        Me.cmdZaloz.Size = New System.Drawing.Size(162, 24)
        Me.cmdZaloz.TabIndex = 32
        Me.cmdZaloz.Text = "Za³ó¿ Bazê Danych"
        Me.cmdZaloz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(372, 378)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(94, 24)
        Me.cmdPowrot.TabIndex = 31
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(9, 376)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 26)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 30
        Me.PictureBox2.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.txtNazwaBazy)
        Me.Panel1.Controls.Add(Me.txtRaport)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(9, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(456, 362)
        Me.Panel1.TabIndex = 29
        '
        'txtNazwaBazy
        '
        Me.txtNazwaBazy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNazwaBazy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtNazwaBazy.ForeColor = System.Drawing.Color.Purple
        Me.txtNazwaBazy.Location = New System.Drawing.Point(126, 5)
        Me.txtNazwaBazy.Name = "txtNazwaBazy"
        Me.txtNazwaBazy.Size = New System.Drawing.Size(314, 20)
        Me.txtNazwaBazy.TabIndex = 52
        '
        'txtRaport
        '
        Me.txtRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport.Location = New System.Drawing.Point(8, 30)
        Me.txtRaport.Multiline = True
        Me.txtRaport.Name = "txtRaport"
        Me.txtRaport.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtRaport.Size = New System.Drawing.Size(432, 326)
        Me.txtRaport.TabIndex = 51
        Me.txtRaport.WordWrap = False
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 16)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Nazwa Bazy Danych . . . . . .. . ."
        '
        'ZalozBazeSQL
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(472, 408)
        Me.Controls.Add(Me.cmdZaloz)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "ZalozBazeSQL"
        Me.ShowInTaskbar = False
        Me.Text = "Zak³adanie nowej bazy SQL"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Dim PlikSQL As String = ""

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub ImportKlienci_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        txtNazwaBazy.Text = ""
    End Sub

    Private Sub ZalozBazeSQL_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub




    '*****************************************************
    '** Importujemy dane
    '*****************************************************

    'Private Sub cmdImportuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim Kwerenda As String = ""
    '    Dim Kwe As String = ""
    '    Dim Pos As Integer = 0

    '    PlikSQL = lblPlik.Text
    '    If Len(PlikSQL) = 0 Then
    '        MessageBox.Show("Brak definicji pliku .sql do pobrania", _
    '            "Uwaga :", _
    '            System.Windows.Forms.MessageBoxButtons.OK, _
    '            MessageBoxIcon.Information, _
    '            MessageBoxDefaultButton.Button1)
    '    Else
    '        If Len(Dir(PlikSQL)) = 0 Then
    '            MessageBox.Show("Nie znalaz³em na dysku pliku do pobrania" & vbCrLf & PlikSQL & vbCrLf, _
    '                "Uwaga :", _
    '                System.Windows.Forms.MessageBoxButtons.OK, _
    '                MessageBoxIcon.Information, _
    '                MessageBoxDefaultButton.Button1)
    '        Else
    '            Try
    '                Dim fs As New FileStream(PlikSQL, FileMode.Open)
    '                Dim sr As New StreamReader(fs)
    '                Kwerenda = sr.ReadToEnd()
    '                sr.Close()
    '                fs.Close()
    '            Catch
    '                MessageBox.Show("B³¹d podczas pobierania pliku .sql: " + Err.Description)
    '                Return
    '            Finally
    '            End Try

    '            txtRaport.Text = Now & "  Zak³adam strukturê SQL" & vbCrLf

    '            Dim sConnectionStr As String = ConnectionStrings()
    '            Dim dbConnection As New SqlClient.SqlConnection(sConnectionStr)
    '            Dim dbCommand As SqlClient.SqlCommand = New SqlClient.SqlCommand
    '            Dim Wynik As Integer

    '            System.Windows.Forms.Application.DoEvents()

    '            dbCommand.Connection = dbConnection
    '            dbCommand.CommandType = CommandType.Text


    '            Pos = InStr(Kwerenda, vbCrLf & "GO" & vbCrLf)
    '            Do While Pos > 0
    '                Kwe = Mid(Kwerenda, 1, Pos - 1)
    '                Kwerenda = Mid(Kwerenda, Pos + Len(vbCrLf & "GO" & vbCrLf))

    '                dbCommand.CommandText = Kwe
    '                dbConnection.Open()
    '                Try
    '                    Wynik = dbCommand.ExecuteNonQuery()
    '                    If Wynik = -1 Then
    '                        'txtRaport.Text &= "Kwerenda : " & Kwe & vbCrLf
    '                        'txtRaport.Text &= "------------Mamy problem" & vbCrLf
    '                    End If
    '                    'Catch ex As Exception
    '                    '    'MessageBox.Show(ex.Message)
    '                    '    txtRaport.Text &= "Kwerenda : " & Kwe & vbCrLf
    '                    '    txtRaport.Text &= "B³¹d wykonania : " & ex.Message & vbCrLf
    '                    '    Return
    '                Catch ex As SqlClient.SqlException
    '                    'MessageBox.Show(ex.Message)
    '                    txtRaport.Text &= "Kwerenda : " & Kwe & vbCrLf
    '                    txtRaport.Text &= "B³¹d wykonania : " & ex.Message & vbCrLf
    '                    Return
    '                Finally
    '                    dbConnection.Close()
    '                    Pos = InStr(Kwerenda, vbCrLf & "GO" & vbCrLf)
    '                End Try
    '            Loop
    '            txtRaport.Text &= "Aktualizacja zakoñczona poprawnie" & vbCrLf
    '        End If
    '    End If
    'End Sub

    Private Sub txtNazwaBazy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazwaBazy.TextChanged
        'TestLen(txtNazwaBazy, "NAZWA BAZY DANYCH", 50)
    End Sub
    Private Sub txtNazwaBazy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNazwaBazy.GotFocus
        StartEdycjiTxt(txtNazwaBazy)
    End Sub
    Private Sub txtNazwaBazy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNazwaBazy.LostFocus
        KoniecEdycjiTxt(txtNazwaBazy)
    End Sub

    '**************************************
    '** zakladeamy nowa baze danych
    '**************************************

    Dim sConnectionStr As String = ConnectionStrings()
    Dim dbConnection As SqlClient.SqlConnection
    Dim dbCommand As SqlClient.SqlCommand

    Private Sub cmdZaloz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZaloz.Click
        Dim Kwerenda As String = ""
        Dim Wynik As Integer = 0
        Dim SQLKatalog As String = ""
        Dim SQLBazaDanych As String = ""

        If Len(txtNazwaBazy.Text) = 0 Then
            MessageBox.Show("Proszê wpisaæ nazwê nowej bazy danych", _
               "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            dbConnection = New SqlClient.SqlConnection(sConnectionStr)
            dbCommand = New SqlClient.SqlCommand

            'sprawdz czy jest po³¹czenie z SQL
            If CzyJestPolaczenieZSQL() Then
                'sprawdz czy jest taka baza
                If CzyJestJuzTakaBazaDanych(Trim(txtNazwaBazy.Text)) Then
                    MessageBox.Show("Istnieje ju¿ baza danych o nazwie" & vbCrLf & txtNazwaBazy.Text & vbCrLf & "Proszê wybraæ inn¹ nazwê bazy danych.",
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    SQLBazaDanych = Trim(txtNazwaBazy.Text)
                    SQLKatalog = PobierzKatalogZPlikamiSQL()
                    '-------------------------------------
                    txtRaport.Text = Now & "  Zak³adam bazê danych SQL o nazwie " & txtNazwaBazy.Text & vbCrLf & vbCrLf
                    System.Windows.Forms.Application.DoEvents()
                    'dbCommand.Connection = dbConnection
                    'dbCommand.CommandType = CommandType.Text

                    'USTAL ŒRODOWISKO SQL
                    Kwerenda = "IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'" & SQLBazaDanych & "')" & vbCrLf &
                               "    DROP DATABASE [" & SQLBazaDanych & "]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "CREATE DATABASE [" & SQLBazaDanych & "]  ON (NAME = N'" & SQLBazaDanych & "_Data', FILENAME = N'" & SQLKatalog & SQLBazaDanych & "_Data.MDF' , SIZE = 5, FILEGROWTH = 1)" & vbCrLf &
                               "    LOG ON (NAME = N'" & SQLBazaDanych & "_Log', FILENAME = N'" & SQLKatalog & SQLBazaDanych & "_Log.LDF' , SIZE = 1, FILEGROWTH = 1)" & vbCrLf &
                               "    COLLATE Polish_CI_AS"
                    If Not SQLcommand(Kwerenda) Then Return








                    '-------------------------------------------------------
                    'This post is inspired by a question raised by a reader in the What is New in SQL Server 2012 section
                    'where I posted a note that the system stored procedure sp_dboption is not available in SQL Server 2012 
                    'anymore. Books online says the recommended alternative is to use the ALTER DATABASE command.
                    'Naomi asked me whether the ALTER DATABASE command provides enough options to support all the options 
                    'available with sp_dboption procedure. Though I had seen most of the options that I frequently use available, 
                    'I had never done a full comparison to verify that. Though I was pretty sure that there will be 
                    'enough options available to replace all the previous options, I thought of doing a quick review 
                    'and post back my findings. 
                    '----------------------------
                    ''auto create statistics'
                    'EXEC sp_dboption 'BR', 'auto create statistics', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'auto create statistics', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET AUTO_CREATE_STATISTICS OFF
                    'ALTER DATABASE BR SET AUTO_CREATE_STATISTICS ON
                    '----------------------------
                    ''auto update statistics'
                    'EXEC sp_dboption 'BR', 'auto update statistics', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'auto update statistics', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET AUTO_UPDATE_STATISTICS ON 
                    'ALTER DATABASE BR SET AUTO_UPDATE_STATISTICS OFF
                    '----------------------------
                    ''autoclose'
                    'EXEC sp_dboption 'BR', 'autoclose', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'autoclose', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET AUTO_CLOSE ON 
                    'ALTER DATABASE BR SET AUTO_CLOSE OFF
                    '----------------------------
                    ''autoshrink'
                    'EXEC sp_dboption 'BR', 'autoshrink', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'autoshrink', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET AUTO_SHRINK ON 
                    'ALTER DATABASE BR SET AUTO_SHRINK OFF
                    '----------------------------
                    ''ANSI null default'
                    'EXEC sp_dboption 'BR', 'ANSI null default', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'ANSI null default', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET ANSI_NULL_DEFAULT ON 
                    'ALTER DATABASE BR SET ANSI_NULL_DEFAULT OFF
                    '----------------------------
                    ''ANSI nulls'
                    'EXEC sp_dboption 'BR', 'ANSI nulls', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'ANSI nulls', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET ANSI_NULLS ON 
                    'ALTER DATABASE BR SET ANSI_NULLS OFF
                    '----------------------------
                    ''ANSI warnings'
                    'EXEC sp_dboption 'BR', 'ANSI warnings', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'ANSI warnings', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET ANSI_WARNINGS ON 
                    'ALTER DATABASE BR SET ANSI_WARNINGS OFF
                    '----------------------------
                    ''arithabort'
                    'EXEC sp_dboption 'BR', 'arithabort', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'arithabort', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET ARITHABORT ON 
                    'ALTER DATABASE BR SET ARITHABORT OFF
                    '----------------------------
                    ''concat null yields null'
                    'EXEC sp_dboption 'BR', 'concat null yields null', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'concat null yields null', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET CONCAT_NULL_YIELDS_NULL ON 
                    'ALTER DATABASE BR SET CONCAT_NULL_YIELDS_NULL OFF
                    '----------------------------
                    ''cursor close on commit'
                    'EXEC sp_dboption 'BR', 'cursor close on commit', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'cursor close on commit', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET CURSOR_CLOSE_ON_COMMIT ON 
                    'ALTER DATABASE BR SET CURSOR_CLOSE_ON_COMMIT OFF
                    '----------------------------
                    ''default to local cursor'
                    'EXEC sp_dboption 'BR', 'default to local cursor', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'default to local cursor', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET CURSOR_DEFAULT LOCAL 
                    'ALTER DATABASE BR SET CURSOR_DEFAULT GLOBAL
                    '----------------------------
                    ''numeric roundabort'
                    'EXEC sp_dboption 'BR', 'numeric roundabort', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'numeric roundabort', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET NUMERIC_ROUNDABORT ON 
                    'ALTER DATABASE BR SET NUMERIC_ROUNDABORT OFF
                    '----------------------------
                    ''quoted identifier'
                    'EXEC sp_dboption 'BR', 'quoted identifier', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'quoted identifier', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET QUOTED_IDENTIFIER ON 
                    'ALTER DATABASE BR SET QUOTED_IDENTIFIER OFF
                    '----------------------------
                    ''read only'
                    'EXEC sp_dboption 'BR', 'read only', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'read only', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET READ_ONLY  
                    'ALTER DATABASE BR SET READ_WRITE
                    '----------------------------
                    ''offline'
                    'EXEC sp_dboption 'BR', 'offline', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'offline', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET OFFLINE 
                    'ALTER DATABASE BR SET ONLINE
                    '----------------------------
                    ''recursive triggers'
                    'EXEC sp_dboption 'BR', 'recursive triggers', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'recursive triggers', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET RECURSIVE_TRIGGERS ON 
                    'ALTER DATABASE BR SET RECURSIVE_TRIGGERS OFF
                    '----------------------------
                    ''single user'
                    'EXEC sp_dboption 'BR', 'single user', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'single user', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET SINGLE_USER 
                    'ALTER DATABASE BR SET MULTI_USER
                    '----------------------------
                    ''torn page detection'
                    'EXEC sp_dboption 'BR', 'torn page detection', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'torn page detection', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET TORN_PAGE_DETECTION ON 
                    'ALTER DATABASE BR SET TORN_PAGE_DETECTION OFF
                    '----------------------------
                    ''select into/bulkcopy'
                    'EXEC sp_dboption 'BR', 'select into/bulkcopy', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'select into/bulkcopy', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET RECOVERY BULK_LOGGED  
                    'ALTER DATABASE BR SET RECOVERY FULL 
                    '-- or 
                    'ALTER DATABASE BR SET RECOVERY SIMPLE
                    '----------------------------
                    ''trunc. log on chkpt.'
                    'EXEC sp_dboption 'BR', 'trunc. log on chkpt.', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'trunc. log on chkpt.', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET RECOVERY SIMPLE  
                    'ALTER DATABASE BR SET RECOVERY FULL
                    '-- or 
                    'ALTER DATABASE BR SET RECOVERY BULK_LOGGED
                    '----------------------------
                    ''dbo use only'
                    'EXEC sp_dboption 'BR', 'dbo use only', 'TRUE'; 
                    'EXEC sp_dboption 'BR', 'dbo use only', 'FALSE'; 
                    '-- Replacement 
                    'ALTER DATABASE BR SET RESTRICTED_USER 
                    'ALTER DATABASE BR SET MULTI_USER  
                    '-- or 
                    'ALTER DATABASE BR SET SINGLE_USER
                    '----------------------------
                    ''merge publish', 'published' and 'subscribed'
                    'I have never used these options and I could not find any documentation about the replacement for these options. 
                    'See more information about the new features and changes in SQL Server 2012 at What is new in SQL Server 2012
                    '----------------------------


                    'testuj wersje SQL
                    'SELECT SERVERPROPERTY('productversion'), SERVERPROPERTY ('productlevel'), SERVERPROPERTY ('edition')
                    'SELECT @@version
                    Dim MSSQLversion As String = ""
                    dbConnection = New SqlClient.SqlConnection(sConnectionStr)
                    dbCommand = New SqlClient.SqlCommand
                    Try
                        dbCommand.Connection = dbConnection
                        dbCommand.CommandType = CommandType.Text
                        dbCommand.CommandText = "SELECT @@version"
                        dbConnection.Open()

                        MSSQLversion = dbCommand.ExecuteScalar().ToString
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        txtRaport.Text &= "B³¹d wykonania : " & ex.Message & vbCrLf
                        txtRaport.SelectionStart = txtRaport.Text.Length
                        txtRaport.ScrollToCaret()
                        System.Windows.Forms.Application.DoEvents()
                    Finally
                        dbConnection.Close()
                    End Try
                    dbCommand.Dispose()
                    dbConnection.Dispose()

                    If (InStr(MSSQLversion, "Microsoft SQL Server 2012") = 1) Or
                       (InStr(MSSQLversion, "Microsoft SQL Server 2014") = 1) Or
                       (InStr(MSSQLversion, "Microsoft SQL Server 2016") = 1) Or
                       (InStr(MSSQLversion, "Microsoft SQL Server 2017") = 1) Or
                       (InStr(MSSQLversion, "Microsoft SQL Server 2019") = 1) Then
                        '---------------------------------
                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'autoclose', N'true'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET AUTO_CLOSE ON"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'bulkcopy', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET RECOVERY BULK_LOGGED"
                        If Not SQLcommand(Kwerenda) Then Return
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET RECOVERY FULL"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'torn page detection', N'true'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET TORN_PAGE_DETECTION ON"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'read only', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET READ_WRITE"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'dbo use', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET MULTI_USER"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'autoshrink', N'true'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET AUTO_SHRINK ON"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'ANSI null default', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET ANSI_NULL_DEFAULT OFF"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'recursive triggers', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET RECURSIVE_TRIGGERS OFF"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'ANSI nulls', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET ANSI_NULLS OFF"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'concat null yields null', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET CONCAT_NULL_YIELDS_NULL OFF"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'cursor close on commit', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET CURSOR_CLOSE_ON_COMMIT OFF"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'default to local cursor', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET CURSOR_DEFAULT GLOBAL"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'quoted identifier', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET QUOTED_IDENTIFIER OFF"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'ANSI warnings', N'false'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET ANSI_WARNINGS OFF"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'auto create statistics', N'true'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET AUTO_CREATE_STATISTICS OFF"
                        If Not SQLcommand(Kwerenda) Then Return

                        'Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'auto update statistics', N'true'"
                        Kwerenda = "ALTER DATABASE " & SQLBazaDanych & " SET AUTO_UPDATE_STATISTICS ON"
                        If Not SQLcommand(Kwerenda) Then Return
                        '--------------------
                    Else
                        'wersje poprzednie
                        '---------------------------------
                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'autoclose', N'true'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'bulkcopy', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'trunc. log', N'true'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'torn page detection', N'true'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'read only', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'dbo use', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'single', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'autoshrink', N'true'"
                        'Kwerenda = "ALTER DATABASE [" & SQLBazaDanych & "] SET AUTO_SHRINK ON"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'ANSI null default', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'recursive triggers', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'ANSI nulls', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'concat null yields null', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'cursor close on commit', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'default to local cursor', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'quoted identifier', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'ANSI warnings', N'false'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'auto create statistics', N'true'"
                        If Not SQLcommand(Kwerenda) Then Return

                        Kwerenda = "exec sp_dboption N'" & SQLBazaDanych & "', N'auto update statistics', N'true'"
                        If Not SQLcommand(Kwerenda) Then Return
                        '--------------------
                    End If

                    Kwerenda = "if( (@@microsoftversion / power(2, 24) = 8) and (@@microsoftversion & 0xffff >= 724) )" & vbCrLf &
                            "    exec sp_dboption N'" & SQLBazaDanych & "', N'db chaining', N'false'"
                    If Not SQLcommand(Kwerenda) Then Return







                    'wyczysc strukture
                    Kwerenda = "use [" & SQLBazaDanych & "]"
                    If Not SQLcommand(Kwerenda) Then Return

                    sConnectionStr = "Data Source=" & ParL_Serwer & ";" &
                                   "Integrated Security=SSPI;" &
                                   "Initial Catalog=" & SQLBazaDanych & ""






                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[softart_nazwaprogramu]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & vbCrLf &
                               "   drop procedure [dbo].[softart_nazwaprogramu]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[softart_wersjaprogramu]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & vbCrLf &
                               "   drop procedure [dbo].[softart_wersjaprogramu]"
                    If Not SQLcommand(Kwerenda) Then Return





                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Parametry]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[Parametry]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AkcjeOpis]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[AkcjeOpis]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AkcjeSpec]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[AkcjeSpec]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Klienci]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[Klienci]"
                    If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Kontakty]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
                    '           "   drop table [dbo].[Kontakty]"
                    'If Not SQLcommand(Kwerenda) Then Return
                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[KontaktyHandlowe]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[KontaktyHandlowe]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RodzajeKontaktow]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[RodzajeKontaktow]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Obroty]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[Obroty]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ObrotyMies]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[ObrotyMies]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Osoby]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[Osoby]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Uzytkownicy]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[Uzytkownicy]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AnalizyZakupu]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[AnalizyZakupu]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DaneDoRaportu]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[DaneDoRaportu]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SlownikCoDalej]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[SlownikCoDalej]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CoDalejPlan]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[CoDalejPlan]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LogAktywnosci]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[LogAktywnosci]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Branze]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[Branze]"
                    If Not SQLcommand(Kwerenda) Then Return
                    Kwerenda = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PodBranze]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf &
                               "   drop table [dbo].[PodBranze]"
                    If Not SQLcommand(Kwerenda) Then Return









                    ''zak³adamy u¿ytkownika PRYZMAT
                    'Kwerenda = "if not exists (select * from master.dbo.syslogins where loginname = N'Pryzmat')" & vbCrLf & _
                    '           "BEGIN" & vbCrLf & _
                    '           "   declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'master', @loginlang = N'us_english'" & vbCrLf & _
                    '           "   if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)" & vbCrLf & _
                    '           "      select @logindb = N'master'" & vbCrLf & _
                    '           "   if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'us_english')" & vbCrLf & _
                    '           "      select @loginlang = @@language" & vbCrLf & _
                    '           "   exec sp_addlogin N'Pryzmat', null, @logindb, @loginlang" & vbCrLf & _
                    '           "End"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_addsrvrolemember N'Pryzmat', sysadmin"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_addsrvrolemember N'Pryzmat', securityadmin"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_addsrvrolemember N'Pryzmat', serveradmin"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_addsrvrolemember N'Pryzmat', setupadmin"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_addsrvrolemember N'Pryzmat', processadmin"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_addsrvrolemember N'Pryzmat', diskadmin"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_addsrvrolemember N'Pryzmat', dbcreator"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "if not exists (select * from dbo.sysusers where name = N'Pryzmat')" & vbCrLf & _
                    '        "   EXEC sp_grantdbaccess N'Pryzmat', N'Pryzmat'"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_addrolemember N'db_owner', N'Pryzmat'"
                    'If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "exec sp_password NULL, 'Pryzmat', 'Pryzmat'"
                    'If Not SQLcommand(Kwerenda) Then Return







                    '---------------------------------------------
                    'ZAK£ADAMY TABELE parametry
                    '---------------------------------------------
                    'Public Par_IdentUzytkownika As String = ""
                    'Public Par_NazwaUzytkownika As String = ""
                    'Public Par_AdresUzytkownika As String = ""
                    'Public Par_KontoUzytkownika As String = ""
                    'Public Par_BankUzytkownika As String = ""
                    'Public Par_MiejscowoscUzytkownika As String = ""
                    'Public Par_NIPUzytkownika As String = ""
                    'Public Par_IdentOddzialu As String = ""

                    'Public Par_DataAktAnalizy As String = ""
                    'Public Par_OstOkresAnalizy As String = ""
                    'Public Par_DaneDoAnalizy As String = ""
                    'Public Par_MiesAnalizy1 As String = ""
                    'Public Par_MiesAnalizy2 As String = ""
                    'Public Par_DaneDoPrognozy As String = ""
                    'Public Par_MiesPrognozy As String = ""

                    'Public Par_KatalogOferty As String = ""
                    'Public Par_wwwPAYBACK As String = ""

                    'Public Par_RAKolumna01 As String = ""
                    'Public Par_RAKolumna02 As String = ""
                    'Public Par_RAKolumna03 As String = ""
                    'Public Par_RAKolumna04 As String = ""
                    'Public Par_RAKolumna05 As String = ""
                    'Public Par_RAKolumna06 As String = ""
                    'Public Par_RAKolumna07 As String = ""
                    'Public Par_RAKolumna08 As String = ""
                    'Public Par_RAKolumna09 As String = ""
                    'Public Par_RAKolumna10 As String = ""

                    'Public Par_WartGranWartosc As Double = ""
                    'Public Par_WartGranProcent As Double = ""
                    'Public Par_WartGRANMATWartosc As Double = ""
                    'Public Par_WartGRANMATProcent As Double = ""

                    'Public Par_Wersja As Integer = 1

                    Kwerenda = "CREATE TABLE [dbo].[Parametry] (" & vbCrLf &
                               "   [IDENT] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [IDENTUZYTKOWNIKA] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [NAZWAUZYTKOWNIKA] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [ADRESUZYTKOWNIKA] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KONTOUZYTKOWNIKA] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [BANKUZYTKOWNIKA] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [MIEJSCOWOSCUZYTKOWNIKA] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [NIPUZYTKOWNIKA] [varchar] (15) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [IDENTODDZIALU] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [DATAAKTANALIZY] [varchar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [OSTOKRESANALIZY] [varchar] (14) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [DANEDOANALIZY] [varchar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [MIESANALIZY1] [varchar] (12) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [MIESANALIZY2] [varchar] (12) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [DANEDOPROGNOZY] [varchar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [MIESPROGNOZY] [varchar] (12) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KATALOGOFERTY] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WWWPAYBACK] [varchar] (200) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA01] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA02] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA03] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA04] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA05] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA06] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA07] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA08] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA09] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RAKOLUMNA10] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WARTGRANWARTOSC] [float] NOT NULL ," & vbCrLf &
                               "   [WARTGRANPROCENT] [float] NOT NULL ," & vbCrLf &
                               "   [WARTGRANMATWARTOSC] [float] NOT NULL ," & vbCrLf &
                               "   [WARTGRANMATPROCENT] [float] NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "CREATE TABLE [dbo].[AkcjeOpis] (" & vbCrLf &
                               "   [IDENT] [VarChar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [DATA] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [OPIS] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [UWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [MARKER] [bit] NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "CREATE TABLE [dbo].[AkcjeSpec] (" & vbCrLf &
                               "   [IDENTAKCJI] [VarChar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [IDENTKLI] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return




                    ''---------------------------------------------------------------------
                    ''Katalog Klientow
                    ''-----------------------------------------------
                    'Public oIdentKli As String         'IDENTKLIENTA   Text, 6
                    'Public oNrTOFIKli As Long          'NRTOFI         Long
                    'Public oNrTOFITxtKli As String     'NRTOFITXT      Text 500

                    'Public oKartaPaybackKli As Boolean 'KARTAPAYBACK   Logical
                    'Public oNryKartPaybackKli As String 'NRYKARTPAYBACK   memo

                    'Public oNazwa1Kli As String        'NAZWA1         Text, 100
                    'Public oKodPoczKli As String       'KODPOCZTOWY    Text, 10
                    'Public oMiejscKli As String        'MIEJSCOWOSC    Text, 50
                    'Public oUlicaKli As String         'ULICA          Text, 50
                    'Public oNumNrDomuKli As Integer    'NUMNRDOMU      INTEGER
                    'Public oNrDomuKli As String        'NRDOMU         Text, 10
                    'Public oIDDomuKli As String        'IDDOMU         Text, 10
                    'Public oNIPKli As String           'NIP            Text, 15
                    'Public oTelefonKli As String       'TELEFON        Text, 50
                    'Public oFaxKli As String           'FAX            Text, 50
                    'Public oeMailKli As String         'EMAIL          Text, 50
                    'Public oWpisalKli As String        'KTOWPISAL      Text, 10   uzytkownik
                    'Public oRejonKli As String         'REJKLIENTA     Text, 20   
                    'Public oPKontaktKli As String      'PKONTAKT       Text, 10   data rrrr-mm-dd

                    'Public oBranzaKli As String        'BRANZA         Text, 100
                    'Public oPodBranzaKli As String     'PODBRANZA      Text, 100
                    'Public oLiczbaZaktrudnionychKli As Integer   'LICZBAZATRUDNIONYCH       INTEGER
                    'Public oLiczbaPracZdalnieKli As Integer      'LICZBAPRACZDALNIE         INTEGER
                    ''..............................
                    'Public oWykazUrzadzenKli As String              'WYKLAZURZADZEN    Memo

                    'Public oSposobWyboruDostawcyKli As String       'SPOSOBWYBORUDOSTAWCY    Text 30
                    'Public oCzyZainstalowanoMonitorKli As Boolean   'CZYZAINSTALOWANOMONITOR    Logical
                    'Public oLiczbaUrzadzenKli As Integer            'LICZBAURZADZEN     Integer
                    'Public oLiczbaWydrukowKli As Integer            'LICZBAWYDRUKOW     Integer
                    'Public oPotencjalDrukuKli As String             'POTENCJALDRUKU     Text 2

                    'Public oAktZakupMatEkspKli As Boolean           'AKTZAKUPMATEKSP    Logical
                    'Public oAktRozliczaStronyDrukuKli As Boolean    'AKTROZLICZASTRONYDRUKU    Logical
                    'Public oAktKorzystaZNajmuKli As Boolean         'AKTKORZYSTAZNAJMU  Logical
                    'Public oZaintRozliczaniemStronDrukuKli As Boolean   'ZAINTROZLICZANIEMSTRONDRUKU    Logical
                    'Public oZaintKorzystaniemZNajmuKli As Boolean   'ZAINTKORZYSTANIEMZNAJMU    Logical

                    'Public oDataWeryfSegmentacji As String          'DATAWERYFSEGMENTACJI  Text 10

                    'Public oPlanDlugookresowyKli As Integer         'PLANDLUGOOKRESOWY     Integer
                    'Public oPlanKrotkookresowyKli As Integer        'PLANKROTKOOKRESOWY    Integer
                    ''......................................
                    'Public oRodzSprzetuLKli As Boolean 'RODZSPRZETUL Logical
                    'Public oRodzSprzetuAKli As Boolean 'RODZSPRZETUA Logical
                    'Public oRodzSprzetuTKli As Boolean 'RODZSPRZETUT Logical
                    'Public oIloSprzetuKli As String    'ILOSPRZETU     Text, 100 -> 250/300    !!!!!!!!!!!!
                    'Public oIloSprzetu2Kli As String   'ILOSPRZETU2     Memo
                    'Public oOfertaTZlozeniaKli As String        'OFERTATZLOZENIA         Text, 4
                    'Public oOfertaPlikKli As String    'OFERTAPLIK     Text, 100
                    'Public oUwagikli As String         'UWAGI          Memo

                    'Public oSkupOKontaktRKli As String  'SKUPOKONTAKTR  Text, 4    rok
                    'Public oSkupOKontaktTKli As String  'SKUPOKONTAKTT  Integer    nr tygodnia
                    'Public oSkupOKontaktDKli As String  'SKUPOKONTAKTD  Text, 12   dzien tygodnia
                    'Public oSkupNKontaktRKli As String  'SKUPNKONTAKTR  Text, 4    rok
                    'Public oSkupNKontaktTKli As String  'SKUPNKONTAKTT  Integer    nr tygodnia
                    'Public oSkupNKontaktDKli As String  'SKUPNKONTAKTT  Text, 12    nr tygodnia

                    'Public oSkupUwagikli As String        'SKUPUWAGI      Memo
                    'Public oSprzedOpiekunkli As String    'SPRZEDOPIEKUN    Text, 10   uzytkownik

                    'Public oSprzedOKontaktRKli As String  'SPRZEDOKONTAKTR  Text, 4    rok
                    'Public oSprzedOKontaktTKli As String  'SPRZEDOKONTAKTT  Integer    nr tygodnia
                    'Public oSprzedOKontaktDKli As String  'SPRZEDOKONTAKTD  Text, 12   dzien tygodnia
                    'Public oSprzedNKontaktRKli As String  'SPRZEDNKONTAKTR  Text, 4    rok
                    'Public oSprzedNKontaktTKli As String  'SPRZEDNKONTAKTT  Integer    nr tygodnia
                    'Public oSprzedNKontaktDKli As String  'SKUPNKONTAKTT    Text, 12    nr tygodnia

                    'Public oSprzedUwagiKli As String      'SPRZEDUWAGI      Memo
                    'Public oWersjaKli As Integer          'WERSJA           Integer
                    'Public oMarkerKli As Boolean          'MARKER           boolean
                    'Public oMarkerPKli As Boolean         'MARKERP          boolean
                    'Public oAKTYWNYKli As Boolean         'AKTYWNY          boolean

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





                    Kwerenda = "CREATE TABLE [dbo].[Klienci] (" & vbCrLf &
                               "       [IDENTKLIENTA] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [NRTOFI] [int] NULL ," & vbCrLf &
                               "       [NRTOFITXT] [varchar] (500) COLLATE Polish_CI_AS NULL ," & vbCrLf &
                               "       [KARTAPAYBACK] [bit] NOT NULL ," & vbCrLf &
                               "       [NRYKARTPAYBACK] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [NAZWA1] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [KODPOCZTOWY] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [MIEJSCOWOSC] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [ULICA] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [NUMNRDOMU] [int] NOT NULL ," & vbCrLf &
                               "       [NRDOMU] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [IDDOMU] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [NIP] [VarChar] (15) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [TELEFON] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [FAX] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [EMAIL] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [KTOWPISAL] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [REJKLIENTA] [VarChar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "       [PKONTAKT] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [BRANZA] [VarChar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [PODBRANZA] [VarChar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [LICZBAZATRUDNIONYCH] [int] NOT NULL ," & vbCrLf &
                               "     [LICZBAPRACZDALNIE] [int] NOT NULL ," & vbCrLf &
                               "   [WYKAZURZADZEN] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SPOSOBWYBORUDOSTAWCY] [VarChar] (30) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [ZAINSTALOWANOMONITOR] [bit] NOT NULL ," & vbCrLf &
                               "   [LICZBAURZADZEN] [int] NOT NULL ," & vbCrLf &
                               "   [LICZBAWYDRUKOW] [int] NOT NULL ," & vbCrLf &
                               "   [POTENCJALDRUKU] [VarChar] (2) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [AKTZAKUPMATEKSP] [bit] NOT NULL ," & vbCrLf &
                               "   [AKTROZLICZASTRONYDRUKU] [bit] NOT NULL ," & vbCrLf &
                               "   [AKTKORZYSTAZNAJMU] [bit] NOT NULL ," & vbCrLf &
                               "   [ZAINTROZLICZANIEMSTRONDRUKU] [bit] NOT NULL ," & vbCrLf &
                               "   [ZAINTKORZYSTANIEMZNAJMU] [bit] NOT NULL ," & vbCrLf &
                               "   [DATAWERYFSEGMENTACJI] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [PLANDLUGOOKRESOWY] [int] NOT NULL ," & vbCrLf &
                               "   [PLANKROTKOOKRESOWY] [int] NOT NULL ," & vbCrLf &
                               "     [RODZSPRZETUL] [bit] NOT NULL ," & vbCrLf &
                               "     [RODZSPRZETUA] [bit] NOT NULL ," & vbCrLf &
                               "     [RODZSPRZETUT] [bit] NOT NULL ," & vbCrLf &
                               "     [ILOSPRZETU] [varchar] (300) COLLATE Polish_CI_AS NULL ," & vbCrLf &
                               "     [ILOSPRZETU2] [text] COLLATE Polish_CI_AS NULL ," & vbCrLf &
                               "     [OFERTATZLOZENIA] [varchar] (4) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [OFERTAPLIK] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [UWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SKUPOKONTAKTR] [VarChar] (4) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SKUPOKONTAKTT] [int] NOT NULL ," & vbCrLf &
                               "   [SKUPOKONTAKTD] [VarChar] (12) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SKUPNKONTAKTR] [VarChar] (4) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SKUPNKONTAKTT] [int] NOT NULL ," & vbCrLf &
                               "   [SKUPNKONTAKTD] [VarChar] (12) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SKUPUWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SPRZEDOPIEKUN] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SPRZEDOKONTAKTR] [VarChar] (4) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SPRZEDOKONTAKTT] [int] NOT NULL ," & vbCrLf &
                               "   [SPRZEDOKONTAKTD] [VarChar] (12) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SPRZEDNKONTAKTR] [VarChar] (4) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SPRZEDNKONTAKTT] [int] NOT NULL ," & vbCrLf &
                               "   [SPRZEDNKONTAKTD] [VarChar] (12) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [SPRZEDUWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [MARKER] [bit] NOT NULL ," & vbCrLf &
                               "   [MARKERP] [bit] NOT NULL ," & vbCrLf &
                               "   [AKTYWNY] [bit] NOT NULL ," & vbCrLf &
                               "     [ZAKUPPOPROK] [float] NOT NULL ," & vbCrLf &
                               "     [USLUGIDOD] [varchar] (200) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP01] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP02] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP03] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP04] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP05] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP06] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP07] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP08] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "     [RZURAP09] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "CREATE TABLE [dbo].[Kontakty] (" & vbCrLf & _
                    '           "   [IDENTKLIENTA] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf & _
                    '           "   [DATAKONTAKTU] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf & _
                    '           "   [PRACOWNIK] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf & _
                    '           "   [TEMAT] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf & _
                    '           "   [UWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf & _
                    '           "   [WERSJA] [int] NOT NULL " & vbCrLf & _
                    '           ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
                    'If Not SQLcommand(Kwerenda) Then Return
                    Kwerenda = "CREATE TABLE [dbo].[KontaktyHandlowe] (" & vbCrLf &
                               "   [UNIQID] [VarChar] (40) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [IDENTKLIENTA] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [DATAKONTAKTU] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [PRACOWNIK] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [TEMAT] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [UWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "CREATE TABLE [dbo].[RodzajeKontaktow] (" & vbCrLf &
                               "   [UNIQID] [varchar] (40) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [RODZAJKONTAKTU] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return




                    ''---------------------------------------------------------------------
                    ''Obroty
                    'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
                    'Public oDataObr As String             'DATA             Text,10
                    'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
                    'Public oLIloSprzedObr As Double       'LILOPRZED        Double
                    'Public oLMarSprzedObr As Double       'LMARPRZED        Double
                    'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
                    'Public oAIloSprzedObr As Double       'AILOSPRZED       Double
                    'Public oAMARSprzedObr As Double       'AMARSPRZED       Double
                    'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
                    'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
                    'Public oLOMARSprzedObr As Double       'LOMARPRZED        Double
                    'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
                    'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double
                    'Public oAOMARSprzedObr As Double       'AOMARSPRZED       Double

                    'Public oNAJEMWartSprzedObr As Double      'NAJEMWARTSPRZED      Double
                    'Public oNAJEMIloSprzedObr As Double       'NAJEMILOPRZED        Double
                    'Public oNAJEMMARSprzedObr As Double       'NAJEMMARPRZED        Double

                    'Public oSTRONYWartSprzedObr As Double      'STRONYWARTSPRZED      Double
                    'Public oSTRONYIloSprzedObr As Double       'STRONYILOPRZED        Double
                    'Public oSTRONYMARSprzedObr As Double       'STRONYMARPRZED        Double

                    'Public oDRUKARKIWartSprzedObr As Double      'DRUKARKIWARTSPRZED      Double
                    'Public oDRUKARKIIloSprzedObr As Double       'DRUKARKIILOPRZED        Double
                    'Public oDRUKARKIMARSprzedObr As Double       'DRUKARKIMARPRZED        Double

                    'Public oSKUPWartSprzedObr As Double      'SKUPWARTSPRZED      Double
                    'Public oSKUPIloSprzedObr As Double       'SKUPILOPRZED        Double
                    'Public oSKUPMARSprzedObr As Double       'SKUPMARPRZED        Double

                    'Public oWersjaObr As Integer          'WERSJA           Integer

                    Kwerenda = "CREATE TABLE [dbo].[Obroty] (" & vbCrLf &
                               "   [IDENTKLIENTA] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [DATA] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [LWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LOWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LOILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LOMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AOWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AOILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AOMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [NAJEMWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [NAJEMILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [NAJEMMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [STRONYWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [STRONYILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [STRONYMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [DRUKARKIWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [DRUKARKIILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [DRUKARKIMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [SKUPWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [SKUPILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [SKUPMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return



                    ''---------------------------------------------------------------------
                    ''ObrotyMies
                    'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
                    'Public oMiesiacMies As String          'MIESIAC          Text,7
                    'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
                    'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
                    'Public oLMARSprzedMies As Double       'LMARSPRZED       Double
                    'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
                    'Public oAIloSprzedMies As Double       'AILOSPRZED       Double
                    'Public oAMARSprzedMies As Double       'AMARSPRZED       Double
                    'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
                    'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
                    'Public oLOMARSprzedMies As Double       'LOMARSPRZED       Double
                    'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
                    'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double
                    'Public oAOMARSprzedMies As Double       'AOMARSPRZED       Double

                    'Public oNAJEMWartSprzedMies As Double      'NAJEMWARTSPRZED      Double
                    'Public oNAJEMIloSprzedMies As Double       'NAJEMILOPRZED        Double
                    'Public oNAJEMMARSprzedMies As Double       'NAJEMMARPRZED        Double

                    'Public oSTRONYWartSprzedMies As Double      'STRONYWARTSPRZED      Double
                    'Public oSTRONYIloSprzedMies As Double       'STRONYILOPRZED        Double
                    'Public oSTRONYMARSprzedMies As Double       'STRONYMARPRZED        Double

                    'Public oDRUKARKIWartSprzedMies As Double      'DRUKARKIWARTSPRZED      Double
                    'Public oDRUKARKIIloSprzedMies As Double       'DRUKARKIILOPRZED        Double
                    'Public oDRUKARKIMARSprzedMies As Double       'DRUKARKIMARPRZED        Double

                    'Public oSKUPWartSprzedMies As Double      'SKUPWARTSPRZED      Double
                    'Public oSKUPIloSprzedMies As Double       'SKUPILOPRZED        Double
                    'Public oSKUPMARSprzedMies As Double       'SKUPMARPRZED        Double

                    'Public oWersjaMies As Integer          'WERSJA           Integer

                    Kwerenda = "CREATE TABLE [dbo].[ObrotyMies] (" & vbCrLf &
                               "   [IDENTKLIENTA] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [MIESIAC] [VarChar] (7) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [LWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LOWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LOILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [LOMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AOWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AOILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [AOMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [NAJEMWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [NAJEMILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [NAJEMMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [STRONYWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [STRONYILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [STRONYMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [DRUKARKIWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [DRUKARKIILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [DRUKARKIMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [SKUPWARTSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [SKUPILOSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [SKUPMARSPRZED] [float] NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "CREATE TABLE [dbo].[Osoby] (" & vbCrLf &
                               "   [IDENTKLIENTA] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [OSOBA] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [VIP] [bit] NOT NULL ," & vbCrLf &
                               "   [STANOWISKO] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOMPETENCJE] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [TELEFON] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [TELEFONKOM] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [EMAIL] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "CREATE TABLE [dbo].[Parametry] (" & vbCrLf & _
                    '           "   [PARAMETR] [VarChar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf & _
                    '           "   [OPIS] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf & _
                    '           "   [WARTOSC] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf & _
                    '           "   [WERSJA] [int] NOT NULL " & vbCrLf & _
                    '           ") ON [PRIMARY]"
                    'If Not SQLcommand(Kwerenda) Then Return








                    '-----------------------------
                    ' Uzytkownicy
                    '-----------------------------
                    'Public U_IdentUzytkownika As String   'IDENT          Text    20
                    'Public U_NazwaUzytkownika As String   'NAZWA          Text    100
                    'Public U_FunkcjaUzytkownika As String 'FUNKCJA          Text    100
                    'Public U_HasloUzytkownika As String   'HASLO          Text    100
                    'Public U_UprawnieniaUzytkownika As String 'UPRAWNIENIA          Text    100
                    'Public U_UwagiUzytkownika As String   'UWAGI          Text
                    'Public U_WersjaUzytkownika As Integer 'WERSJA         Integer

                    Kwerenda = "CREATE TABLE [dbo].[Uzytkownicy] (" & vbCrLf &
                               "[IDENT] [varchar] (20) COLLATE Polish_CI_AS NOT NULL," & vbCrLf &
                               "[NAZWA] [varchar] (100) COLLATE Polish_CI_AS," & vbCrLf &
                               "[FUNKCJA] [varchar] (100) COLLATE Polish_CI_AS," & vbCrLf &
                               "[HASLO] [varchar] (100) COLLATE Polish_CI_AS," & vbCrLf &
                               "[UPRAWNIENIA] [varchar] (100) COLLATE Polish_CI_AS," & vbCrLf &
                               "[UWAGI] [text] COLLATE Polish_CI_AS," & vbCrLf &
                               "[WERSJA] [int] " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[Uzytkownicy] WITH NOCHECK ADD " & vbCrLf &
                               "CONSTRAINT [PK_Uzytkownicy] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "(" & vbCrLf &
                               "[IDENT]" & vbCrLf &
                               ")  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "CREATE TABLE [dbo].[AnalizyZakupu] (" & vbCrLf &
                               "   [IDENTKLIENTA] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WARTZA00] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA00] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM00] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA00] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZABM] [float] NOT NULL ," & vbCrLf &
                               "   [MARZABM] [float] NOT NULL ," & vbCrLf &
                               "   [PROCMBM] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZABM] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA01] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA01] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM01] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA01] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA02] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA02] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM02] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA02] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA03] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA03] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM03] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA03] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA04] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA04] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM04] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA04] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA05] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA05] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM05] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA05] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA06] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA06] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM06] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA06] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA07] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA07] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM07] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA07] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA08] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA08] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM08] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA08] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA09] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA09] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM09] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA09] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA10] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA10] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM10] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA10] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA11] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA11] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM11] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA11] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA12] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA12] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM12] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA12] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA13] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA13] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM13] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA13] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA14] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA14] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM14] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA14] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA15] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA15] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM15] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA15] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA16] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA16] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM16] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA16] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA17] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA17] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM17] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA17] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA18] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA18] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM18] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA18] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA19] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA19] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM19] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA19] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA20] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA20] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM20] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA20] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA21] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA21] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM21] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA21] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA22] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA22] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM22] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA22] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA23] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA23] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM23] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA23] [float] NOT NULL ," & vbCrLf &
                               "   [WARTZA24] [float] NOT NULL ," & vbCrLf &
                               "   [MARZA24] [float] NOT NULL ," & vbCrLf &
                               "   [PROCM24] [float] NOT NULL ," & vbCrLf &
                               "   [ILOSZA24] [float] NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return



                    '---------------------------------------------------------------------
                    'DaneDoRaportu
                    '---------------------------------------------------------------------
                    'Public ddrPracownik As String        'PRACOWNIK   Text, 10
                    'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
                    'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
                    'Public ddrOferty As Integer          'OFERTY         Integer
                    'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
                    'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
                    'Public ddrDostawy As Integer         'DOSTAWY        Integer
                    'Public ddrSkup As Integer            'SKUP           Integer

                    'Public ddrKolExcel02 As String            'KolExCEL02           String
                    'Public ddrKolExcel03 As String            'KolExCEL03           String
                    'Public ddrKolExcel04 As String            'KolExCEL04           String
                    'Public ddrKolExcel05 As String            'KolExCEL05           String
                    'Public ddrKolExcel06 As String            'KolExCEL06           String
                    'Public ddrKolExcel07 As String            'KolExCEL07           String
                    'Public ddrKolExcel08 As String            'KolExCEL08           String
                    'Public ddrKolExcel09 As String            'KolExCEL09           String
                    'Public ddrKolExcel10 As String            'KolExCEL10           String

                    'Public ddrExcel01 As Integer            'EXCEL01           Integer
                    'Public ddrExcel02 As Integer            'EXCEL02           Integer
                    'Public ddrExcel03 As Integer            'EXCEL03           Integer
                    'Public ddrExcel04 As Integer            'EXCEL04           Integer
                    'Public ddrExcel05 As Integer            'EXCEL05           Integer
                    'Public ddrExcel06 As Integer            'EXCEL06           Integer
                    'Public ddrExcel07 As Integer            'EXCEL07           Integer
                    'Public ddrExcel08 As Integer            'EXCEL08           Integer
                    'Public ddrExcel09 As Integer            'EXCEL09           Integer
                    'Public ddrExcel10 As Integer            'EXCEL10           Integer

                    'Public ddrWersja As Integer          'WERSJA         Integer
                    Kwerenda = "CREATE TABLE [dbo].[DaneDoRaportu] (" & vbCrLf &
                               "   [PRACOWNIK] [varchar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [DATARAPORTU] [varchar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KLBEZDRUKARKI] [int] NOT NULL ," & vbCrLf &
                               "   [OFERTY] [int] NOT NULL ," & vbCrLf &
                               "   [TESTYWSTAWIONE] [int] NOT NULL ," & vbCrLf &
                               "   [CZYSZCZENIE] [int] NOT NULL ," & vbCrLf &
                               "   [DOSTAWY] [int] NOT NULL ," & vbCrLf &
                               "   [SKUP] [int] NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL01] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL02] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL03] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL04] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL05] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL06] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL07] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL08] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL09] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KOLEXCEL10] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [EXCEL01] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL02] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL03] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL04] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL05] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL06] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL07] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL08] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL09] [int] NOT NULL ," & vbCrLf &
                               "   [EXCEL10] [int] NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return












                    '---------------------------------------------------------------------
                    'SlownikCoDalej
                    '---------------------------------------------------------------------
                    'Public scdIDENT As String             'IDENT       Text, 40
                    'Public scdOPIS As String              'OPIS        memo
                    'Public scdWersja As Integer           'WERSJA      Integer

                    Kwerenda = "CREATE TABLE [dbo].[SlownikCoDalej] (" & vbCrLf &
                               "   [IDENT] [varchar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [OPIS] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return









                    '---------------------------------------------------------------------
                    'CoDalej
                    '---------------------------------------------------------------------
                    'Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
                    'Public cdIDENT As String             'IDENT        Text, 40
                    'Public cdOPIS As String              'OPIS         memo
                    'Public cdWersja As Integer           'WERSJA       Integer

                    Kwerenda = "CREATE TABLE [dbo].[CoDalejPlan] (" & vbCrLf &
                               "   [UNIQID] [varchar] (40) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [IDENTKLIENTA] [varchar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [IDENT] [varchar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [OPIS] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return





                    ''----------------------------------------------------------
                    ''Katalog LogAktywnosci
                    ''----------------------------------------------------------
                    'Public LG_UniqID As String       'UNIQID     STRING 40
                    'Public LG_Temat As String        'TEMAT      STRING 100
                    'Public LG_Katalog As String      'KATALOG    STRING 100
                    'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
                    'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
                    'Public LG_Operacja As String     'OPERACJA   STRING 20
                    'Public LG_Uwagi As String        'UWAGI      Text
                    'Public LG_Wersja As Integer      'WERSJA     Integer

                    Kwerenda = "CREATE TABLE [dbo].[LogAktywnosci] (" & vbCrLf &
                               "   [UNIQID] [varchar] (40) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [TEMAT] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [KATALOG] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [DATAZAPISU] [varchar] (30) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [PRACOWNIK] [varchar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [OPERACJA] [varchar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [UWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[LogAktywnosci] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_LogAktywnosci] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [UNIQID] " & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return








                    '---------------------------------------------------------------------
                    'Branze
                    '---------------------------------------------------------------------
                    'Public brIdBranzy As String            'IDBRANZY         Text, 
                    'Public brWersja As Integer             'WERSJA           Integer

                    Kwerenda = "CREATE TABLE [dbo].[Branze] (" & vbCrLf &
                               "   [IDBRANZY] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[Branze] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_Branze] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDBRANZY]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return






                    '---------------------------------------------------------------------
                    'PodBranze
                    '---------------------------------------------------------------------
                    'Public pbrIdBranzy As String            'IDBRANZY         Text, 
                    'Public pbrIdPodBranzy As String         'IDPODBRANZY         Text, 
                    'Public pbrWersja As Integer             'WERSJA           Integer

                    Kwerenda = "CREATE TABLE [dbo].[PodBranze] (" & vbCrLf &
                               "   [IDBRANZY] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [IDpodBRANZY] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                               "   [WERSJA] [int] NOT NULL " & vbCrLf &
                               ") ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[PodBranze] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_PodBranze] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDBRANZY], " & vbCrLf &
                               "   [IDPODBRANZY]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return










                    'definicja kluczy w tabelach
                    Kwerenda = "ALTER TABLE [dbo].[Parametry] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_Parametry] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENT]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[AkcjeOpis] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_AkcjeOpis] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENT]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[AkcjeSpec] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_AkcjeSpec] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTAKCJI], " & vbCrLf &
                               "   [IDENTKLI]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[Klienci] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_Klienci] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "ALTER TABLE [dbo].[Kontakty] WITH NOCHECK ADD " & vbCrLf & _
                    '           "   CONSTRAINT [PK_Kontakty] PRIMARY KEY  CLUSTERED " & vbCrLf & _
                    '           "   (" & vbCrLf & _
                    '           "   [IDENTKLIENTA]," & vbCrLf & _
                    '           "   [DATAKONTAKTU]" & vbCrLf & _
                    '           "   )  ON [PRIMARY] "
                    'If Not SQLcommand(Kwerenda) Then Return
                    Kwerenda = "ALTER TABLE [dbo].[KontaktyHandlowe] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_KontaktyHandlowe] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [UNIQID]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[RodzajeKontaktow] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_RodzajeKontaktow] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [UNIQID] " & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[Obroty] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_Obroty] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA], " & vbCrLf &
                               "   [DATA]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[ObrotyMies] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_ObrotyMies] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA], " & vbCrLf &
                               "   [MIESIAC]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[Osoby] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_Osoby] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA]," & vbCrLf &
                               "   [OSOBA]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    'Kwerenda = "ALTER TABLE [dbo].[Uzytkownicy] WITH NOCHECK ADD " & vbCrLf &
                    '           "   CONSTRAINT [PK_Uzytkownicy] PRIMARY KEY  CLUSTERED " & vbCrLf &
                    '           "   (" & vbCrLf &
                    '           "   [IDENT]" & vbCrLf &
                    '           "   )  ON [PRIMARY] "
                    'If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[AnalizyZakupu] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_AnalizyZakupu] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA] " & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[DaneDoRaportu] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_DaneDoRaportu] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [PRACOWNIK], " & vbCrLf &
                               "   [DATARAPORTU]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[SlownikCoDalej] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_SlownikCoDalej] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENT] " & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "ALTER TABLE [dbo].[CoDalejPlan] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_CoDalej] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [UNIQID], " & vbCrLf &
                               "   [IDENTKLIENTA] " & vbCrLf &
                               "   )  ON [PRIMARY] "
                    If Not SQLcommand(Kwerenda) Then Return














                    'procedury sk³adowane
                    Kwerenda = "SET QUOTED_IDENTIFIER ON "
                    If Not SQLcommand(Kwerenda) Then Return
                    Kwerenda = "SET ANSI_NULLS ON "
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "/*  Podaje nazwe programu" & vbCrLf &
                               "*/ " & vbCrLf &
                               "create procedure dbo.softart_nazwaprogramu(@name varchar(50) output) " & vbCrLf &
                               "as " & vbCrLf &
                               "   set @name = 'KartotekaKlientowPryzmat' " & vbCrLf &
                               "   Return"
                    If Not SQLcommand(Kwerenda) Then Return


                    Kwerenda = "/*  Podaje wersje programu" & vbCrLf &
                               "*/ " & vbCrLf &
                               "create procedure dbo.softart_wersjaprogramu" & vbCrLf &
                               "@wersja int output " & vbCrLf &
                               "as " & vbCrLf &
                               "   set @wersja = " & Trim(Str(oAktualnaWersjaBazyDanych)) & vbCrLf &
                               "   Return"
                    If Not SQLcommand(Kwerenda) Then Return

                    Kwerenda = "SET QUOTED_IDENTIFIER OFF "
                    If Not SQLcommand(Kwerenda) Then Return
                    Kwerenda = "SET ANSI_NULLS ON "
                    If Not SQLcommand(Kwerenda) Then Return




                    '-------------------------------------
                    txtRaport.Text &= vbCrLf & Now & "  Aktualizacja zakoñczona poprawnie" & vbCrLf
                    txtRaport.SelectionStart = Len(txtRaport.Text)
                    txtRaport.SelectionLength = 0
                    txtRaport.ScrollToCaret()
                    System.Windows.Forms.Application.DoEvents()
                End If
            End If
        End If
    End Sub

    Private Function SQLcommand(ByVal Kwe As String) As Boolean
        Dim Wynik As Integer = 0

        txtRaport.Text &= Kwe & vbCrLf
        txtRaport.SelectionStart = Len(txtRaport.Text)
        txtRaport.SelectionLength = 0
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()
        '---------------------------------
        dbConnection = New SqlClient.SqlConnection(sConnectionStr)
        dbCommand = New SqlClient.SqlCommand
        Try
            dbCommand.Connection = dbConnection
            dbCommand.CommandType = CommandType.Text
            dbCommand.CommandText = Kwe
            dbConnection.Open()
            Wynik = dbCommand.ExecuteNonQuery()
            If Wynik = -1 Then
            End If
            'txtRaport.Text &= "OK, aktualizacja wykonana poprawnie" & vbCrLf
        Catch ex As SqlClient.SqlException
            'MessageBox.Show(ex.Message)
            txtRaport.Text &= vbCrLf
            txtRaport.Text &= "Kwerenda : " & Kwe & vbCrLf
            txtRaport.Text &= Now & "B³¹d wykonania : " & ex.Message & vbCrLf

            txtRaport.SelectionStart = Len(txtRaport.Text)
            txtRaport.SelectionLength = 0
            txtRaport.ScrollToCaret()
            System.Windows.Forms.Application.DoEvents()
            Return (False)
        Finally
            dbConnection.Close()
        End Try
        dbCommand.Dispose()
        dbConnection.Dispose()
        Return (True)
    End Function

    Private Function CzyJestPolaczenieZSQL() As Boolean
        Dim NazwaSerwera As String = ""

        Try
            dbCommand.Connection = dbConnection
            dbCommand.CommandType = CommandType.Text
            dbConnection.Open()
            '------------------------------------------------
            dbCommand.CommandText = "SELECT @@SERVERNAME"
            NazwaSerwera = dbCommand.ExecuteScalar()
            '------------------------------------------------

        Catch ex As Exception
            MessageBox.Show("B³¹d po³¹czenia z baz¹ danych SQL" & vbCrLf & "na serwerze " & ParL_Serwer & "...", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            'MessageBox.Show(ex.Message)
        Finally
            dbConnection.Close()
        End Try

        Return (Len(NazwaSerwera) > 0)
    End Function

    Private Function CzyJestJuzTakaBazaDanych(ByVal NazwaBazy As String) As Boolean
        Dim JestBaza As Boolean = False

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
                        If UCase(db) = UCase(NazwaBazy) Then
                            JestBaza = True
                            Exit For
                        End If
                    Next
                End If
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Finally
                If dbConnection.State = ConnectionState.Open Then
                    dbConnection.Close()
                End If
            End Try
        End If
        '------------------------------------------------
        Return (JestBaza)
    End Function

    Private Function PobierzKatalogZPlikamiSQL() As String
        Dim Katalog As String = ""

        'Dim sSelectSQL As String = "select name, physical_name from sys.database_files"
        Dim sSelectSQL As String = "select name, filename from master.dbo.sysdatabases"
        'Dim dbConnection As New SqlClient.SqlConnection(sConnectionStr)
        Dim dbCommandSelect As New SqlClient.SqlCommand(sSelectSQL, dbConnection)
        Dim dbDataAdapter As New SqlClient.SqlDataAdapter(dbCommandSelect)
        Dim dbDataSet As New DataSet
        Dim dbDataView As New DataView

        dbDataSet.Locale = New System.Globalization.CultureInfo("pl-PL")
        Dim MapowanieTabeli As System.Data.Common.DataTableMapping
        MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Database")
        Try
            dbConnection.Open()
        Catch Ex As System.Exception
        End Try

        If dbConnection.State = ConnectionState.Open Then
            Try
                dbDataAdapter.FillSchema(dbDataSet, SchemaType.Mapped, "TABELA_Database")
                dbDataAdapter.Fill(dbDataSet)
                dbConnection.Close()

                'definiuj dbDataView
                dbDataView = New DataView(dbDataSet.Tables(0))
                If dbDataView.Count > 0 Then
                    Dim dr As DataRow
                    Dim dbn As String = ""
                    Dim pos As Integer = 0
                    For Each dr In dbDataView.Table.Rows
                        dbn = dr.Item("NAME")
                        If Trim(dbn) = "master" Then
                            Katalog = dr.Item("FILENAME")
                            pos = InStr(Katalog, "master")
                            If pos > 0 Then
                                Katalog = Mid(Katalog, 1, pos - 1)
                            End If
                            Exit For
                        End If
                    Next
                End If
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Finally
                If dbConnection.State = ConnectionState.Open Then
                    dbConnection.Close()
                End If
            End Try
        End If
        '------------------------------------------------
        Return (Katalog)
    End Function
End Class

