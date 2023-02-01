Public Class frmZmianaHasla
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    '====================================================
    'zmienne do œledzenia JAK ZAMKNIETO FORME
    Private _closeClick As Boolean
    Private _zapiszHaslo As Boolean
    'Private components As System.ComponentModel.Container
    Public Const SC_CLOSE As Integer = 61536
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPokazNoweHaslo As System.Windows.Forms.Button
    Friend WithEvents cmdPokazHaslo As System.Windows.Forms.Button
    Public Const WM_SYSCOMMAND As Integer = 274
    'Fields
    'Constructors
    'Events
    'Methods

    Public Sub New(ByVal ZapiszHaslo As Boolean)
        MyBase.New()
        _closeClick = False
        _zapiszHaslo = ZapiszHaslo
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
    End Sub

    Protected Overloads Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_SYSCOMMAND AndAlso m.WParam.ToInt32 = SC_CLOSE Then
            'zamkniêto naciskaj¹c krzyzyk w prawym gornym rogu...
            Me._closeClick = True
        End If
        MyBase.WndProc(m)
    End Sub
    '====================================================

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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblData As System.Windows.Forms.Label
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents cmdSprawdz As System.Windows.Forms.Button
    Friend WithEvents cmdZakoncz As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtNoweHaslo2 As System.Windows.Forms.TextBox
    Friend WithEvents txtStareHaslo As System.Windows.Forms.TextBox
    Friend WithEvents lblFunkcja As System.Windows.Forms.Label
    Friend WithEvents lblNazwa As System.Windows.Forms.Label
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents txtNoweHaslo As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmZmianaHasla))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdPokazNoweHaslo = New System.Windows.Forms.Button()
        Me.cmdPokazHaslo = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.txtNoweHaslo2 = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtStareHaslo = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblFunkcja = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblNazwa = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblData = New System.Windows.Forms.Label()
        Me.txtNoweHaslo = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdSprawdz = New System.Windows.Forms.Button()
        Me.cmdZakoncz = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdPokazNoweHaslo)
        Me.Panel1.Controls.Add(Me.cmdPokazHaslo)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.txtNoweHaslo2)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.txtStareHaslo)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.lblFunkcja)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblNazwa)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lblData)
        Me.Panel1.Controls.Add(Me.txtNoweHaslo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(368, 192)
        Me.Panel1.TabIndex = 0
        '
        'cmdPokazNoweHaslo
        '
        Me.cmdPokazNoweHaslo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazNoweHaslo.Location = New System.Drawing.Point(280, 136)
        Me.cmdPokazNoweHaslo.Name = "cmdPokazNoweHaslo"
        Me.cmdPokazNoweHaslo.Size = New System.Drawing.Size(80, 23)
        Me.cmdPokazNoweHaslo.TabIndex = 209
        Me.cmdPokazNoweHaslo.Text = "Poka¿ has³o"
        '
        'cmdPokazHaslo
        '
        Me.cmdPokazHaslo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazHaslo.Location = New System.Drawing.Point(280, 106)
        Me.cmdPokazHaslo.Name = "cmdPokazHaslo"
        Me.cmdPokazHaslo.Size = New System.Drawing.Size(80, 23)
        Me.cmdPokazHaslo.TabIndex = 208
        Me.cmdPokazHaslo.Text = "Poka¿ has³o"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(290, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(61, 51)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 205
        Me.PictureBox1.TabStop = False
        '
        'txtNoweHaslo2
        '
        Me.txtNoweHaslo2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtNoweHaslo2.ForeColor = System.Drawing.Color.Purple
        Me.txtNoweHaslo2.Location = New System.Drawing.Point(120, 160)
        Me.txtNoweHaslo2.Name = "txtNoweHaslo2"
        Me.txtNoweHaslo2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNoweHaslo2.Size = New System.Drawing.Size(150, 20)
        Me.txtNoweHaslo2.TabIndex = 23
        Me.txtNoweHaslo2.Text = "Data"
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(8, 163)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(112, 16)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "Powtórz Nowe has³o. . . . . . . . . . . . ."
        '
        'txtStareHaslo
        '
        Me.txtStareHaslo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtStareHaslo.ForeColor = System.Drawing.Color.Purple
        Me.txtStareHaslo.Location = New System.Drawing.Point(120, 108)
        Me.txtStareHaslo.Name = "txtStareHaslo"
        Me.txtStareHaslo.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtStareHaslo.Size = New System.Drawing.Size(150, 20)
        Me.txtStareHaslo.TabIndex = 12
        Me.txtStareHaslo.Text = "Data"
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(8, 111)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 22
        Me.Label8.Text = "Aktualne has³o. . . . . . . . . . . . ."
        '
        'lblFunkcja
        '
        Me.lblFunkcja.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFunkcja.ForeColor = System.Drawing.Color.Purple
        Me.lblFunkcja.Location = New System.Drawing.Point(120, 80)
        Me.lblFunkcja.Name = "lblFunkcja"
        Me.lblFunkcja.Size = New System.Drawing.Size(240, 16)
        Me.lblFunkcja.TabIndex = 20
        Me.lblFunkcja.Text = "MójKomp"
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(8, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Funkcja u¿ytkownika. . . . . . . . . . . . ."
        '
        'lblNazwa
        '
        Me.lblNazwa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa.Location = New System.Drawing.Point(120, 56)
        Me.lblNazwa.Name = "lblNazwa"
        Me.lblNazwa.Size = New System.Drawing.Size(240, 16)
        Me.lblNazwa.TabIndex = 19
        Me.lblNazwa.Text = "MójKomp"
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(120, 32)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(150, 16)
        Me.lblIdent.TabIndex = 18
        Me.lblIdent.Text = "MójKomp"
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(8, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 16)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Ident.u¿ytkownika. . . . . . . . . . . . ."
        '
        'lblData
        '
        Me.lblData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblData.ForeColor = System.Drawing.Color.Purple
        Me.lblData.Location = New System.Drawing.Point(120, 8)
        Me.lblData.Name = "lblData"
        Me.lblData.Size = New System.Drawing.Size(150, 16)
        Me.lblData.TabIndex = 16
        Me.lblData.Text = "2005-11-16"
        '
        'txtNoweHaslo
        '
        Me.txtNoweHaslo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtNoweHaslo.ForeColor = System.Drawing.Color.Purple
        Me.txtNoweHaslo.Location = New System.Drawing.Point(120, 136)
        Me.txtNoweHaslo.Name = "txtNoweHaslo"
        Me.txtNoweHaslo.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNoweHaslo.Size = New System.Drawing.Size(150, 20)
        Me.txtNoweHaslo.TabIndex = 14
        Me.txtNoweHaslo.Text = "Data"
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 139)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Nowe has³o. . . . . . . . . . . . ."
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Nazwa u¿ytkownika. . . . . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Data i godzina . . . . . . . . . . . . ."
        '
        'cmdSprawdz
        '
        Me.cmdSprawdz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSprawdz.Image = CType(resources.GetObject("cmdSprawdz.Image"), System.Drawing.Image)
        Me.cmdSprawdz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSprawdz.Location = New System.Drawing.Point(365, 241)
        Me.cmdSprawdz.Name = "cmdSprawdz"
        Me.cmdSprawdz.Size = New System.Drawing.Size(89, 23)
        Me.cmdSprawdz.TabIndex = 201
        Me.cmdSprawdz.Text = "Zapamiêtaj"
        Me.cmdSprawdz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdZakoncz
        '
        Me.cmdZakoncz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZakoncz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZakoncz.Location = New System.Drawing.Point(209, 241)
        Me.cmdZakoncz.Name = "cmdZakoncz"
        Me.cmdZakoncz.Size = New System.Drawing.Size(150, 23)
        Me.cmdZakoncz.TabIndex = 202
        Me.cmdZakoncz.Text = "Pozostaw stare has³o"
        Me.cmdZakoncz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 241)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(88, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 203
        Me.picLogo.TabStop = False
        '
        'frmZmianaHasla
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(461, 271)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdZakoncz)
        Me.Controls.Add(Me.cmdSprawdz)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmZmianaHasla"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Zmiana has³a U¿ytkownika"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    '====================================================
    'UWAGA - aby to zadzialalo trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    '====================================================
    Private Sub frmZmianaHasla_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me._closeClick Then
            e.Cancel = True     'nie pozwalaj zamkn¹c forme
            MessageBox.Show("Proszê zamkn¹æ formê klawiszami" & vbCrLf & "ZAPAMIÊTAJ lub POZOSTAW STARE HAS£O", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            Me._closeClick = False
        Else
        End If
    End Sub







    '----------------------------------------------------------
    'Katalog Uzytkownicy
    '----------------------------------------------------------
    'Public U_IdentUzytkownika As String   'IDENT          Text    20
    'Public U_NazwaUzytkownika As String   'NAZWA          Text    100
    'Public U_FunkcjaUzytkownika As String 'FUNKCJA          Text    100
    'Public U_HasloUzytkownika As String   'HASLO          Text    100
    'Public U_UprawnieniaUzytkownika As String 'UPRAWNIENIA          Text    100
    'Public U_MagazynyUzytkownika As String 'UPRAWNIENIA          Text    100

    'Public U_RKRokRaportu As String = ""          'RKROKRAPORTU          TEXT 4
    'Public U_RKNumerRaportu As Integer = 0        'RKNUMERRAPORTU        INTEGER
    'Public U_RKNumerPozycji As Integer = 0        'RKNUMERPOZYCJI        INTEGER
    'Public U_RKNumerDowoduKP As Integer = 0       'RKNUMERDOWODUKP       INTEGER
    'Public U_RKNumerDowoduKW As Integer = 0       'RKNUMERDOWODUKW       INTEGER

    'Public U_OpisUzytkownika As String    'OPIS           Text    100
    'Public U_UwagiUzytkownika As String   'UWAGI          Text
    'Public U_WersjaUzytkownika As Integer 'WERSJA         Integer

    Dim sSelectSQL As String = ""

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

    Dim DataSetUzytkownicy As New DataSet
    Dim DataViewUzytkownicy As New DataView
    Dim ConnectionState As System.Data.ConnectionState


    Private Sub frmZmianaHasla_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.picLogo.Image = My.Resources.logomini
        'Me.cmdSprawdz.Image = My.Resources.ok1.ToBitmap
        'Me.cmdZakoncz.Image = My.Resources._ESCAPE.ToBitmap
        '-------------------------------
        lblData.Text = Now
        lblIdent.Text = Program_IdUzytkownika
        lblNazwa.Text = Program_NazwaUzytkownika
        lblFunkcja.Text = Program_FunkcjaUzytkownika

        txtStareHaslo.Text = ""
        txtNoweHaslo.Text = ""
        txtNoweHaslo2.Text = ""
    End Sub

    Private Sub frmZmianaHasla_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    Private Sub txtStareHaslo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStareHaslo.TextChanged
        TestLen(txtStareHaslo, "STARE HAS£O", 100)
    End Sub
    Private Sub txtStareHaslo_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStareHaslo.GotFocus
        StartEdycjiTxt(txtStareHaslo)
    End Sub
    Private Sub txtStareHaslo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStareHaslo.LostFocus
        KoniecEdycjiTxt(txtStareHaslo)
    End Sub

    Private Sub txtNoweHaslo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNoweHaslo.TextChanged
        TestLen(txtNoweHaslo, "NOWE HAS£O", 100)
    End Sub
    Private Sub txtNoweHaslo_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNoweHaslo.GotFocus
        StartEdycjiTxt(txtNoweHaslo)
    End Sub
    Private Sub txtNoweHaslo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNoweHaslo.LostFocus
        KoniecEdycjiTxt(txtNoweHaslo)
    End Sub

    Private Sub txtNoweHaslo2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNoweHaslo2.TextChanged
        TestLen(txtNoweHaslo2, "POWTÓRZONE NOWE HAS£O", 100)
    End Sub
    Private Sub txtNoweHaslo2_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNoweHaslo2.GotFocus
        StartEdycjiTxt(txtNoweHaslo2)
    End Sub
    Private Sub txtNoweHaslo2_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNoweHaslo2.LostFocus
        KoniecEdycjiTxt(txtNoweHaslo2)
    End Sub





    Private Sub cmdZakoncz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZakoncz.Click
        'nie zmieniamy has³a
        Me.Close()
    End Sub

    Private Sub cmdSprawdz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSprawdz.Click
        'sprawdz czy stare haslo podano poprawnie
        If txtStareHaslo.Text = Program_HasloUzytkownika Then
            'sprawdz czy nowe haslo powtorzono poprawnie
            If txtNoweHaslo.Text = txtNoweHaslo2.Text Then

                Program_HasloUzytkownika = txtNoweHaslo.Text
                If _zapiszHaslo Then
                    ZapiszNoweHaslo(Program_IdUzytkownika, Program_HasloUzytkownika)
                End If

                MessageBox.Show("Zmieni³em has³o u¿ytkownika " & Program_IdUzytkownika, _
                    "Poprawna zmiana has³a :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Exclamation, _
                    MessageBoxDefaultButton.Button1)
                Me.Close()
            Else
                MessageBox.Show("Niepoprawnie powtórzy³eœ nowe has³o", _
                    "B³¹d podczas wprowadzania :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Exclamation, _
                    MessageBoxDefaultButton.Button1)
            End If
        Else
            MessageBox.Show("Poda³eœ b³êdne stare has³o", _
                "Identyfikacja negatywna :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Exclamation, _
                MessageBoxDefaultButton.Button1)
        End If
    End Sub

    Private Sub cmdPokazHaslo_Click(sender As Object, e As EventArgs) Handles cmdPokazHaslo.Click
        If txtStareHaslo.PasswordChar = "*" Then
            txtStareHaslo.PasswordChar = ""
            cmdPokazHaslo.Text = "Ukryj has³o"
        Else
            txtStareHaslo.PasswordChar = "*"
            cmdPokazHaslo.Text = "Poka¿ has³o"
        End If
    End Sub

    Private Sub cmdPokazNoweHaslo_Click(sender As Object, e As EventArgs) Handles cmdPokazNoweHaslo.Click
        If txtNoweHaslo.PasswordChar = "*" Then
            txtNoweHaslo.PasswordChar = ""
            txtNoweHaslo2.PasswordChar = ""
            cmdPokazNoweHaslo.Text = "Ukryj has³o"
        Else
            txtNoweHaslo.PasswordChar = "*"
            txtNoweHaslo2.PasswordChar = "*"
            cmdPokazNoweHaslo.Text = "Poka¿ has³o"
        End If
    End Sub






    '*************************
    '* ZAPISUJEMY NOWE HASLO DO KATALOGU
    '*************************

    Private Sub ZapiszNoweHaslo(ByVal user As String, ByVal pass As String)
        '-------------------------------
        DataSetUzytkownicy.Locale = New System.Globalization.CultureInfo("pl-PL")

        sSelectSQL = sSelectSQLUzytkownicy & " WHERE IDENT='" & user & "'"

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionUzytkownicy)
            'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLUzytkownicy, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLUzytkownicy, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLUzytkownicy, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLUzytkownicy, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
            'DBMapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Uzytkownicy")

            'DBDeleteUzytkownicy(dbCommandDelete, dbDataAdapter)
            'DBUpdateUzytkownicy(dbCommandUpdate, dbDataAdapter)
            'DBInsertUzytkownicy(dbCommandInsert, dbDataAdapter)

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
            '    '    MessageBoxDefaultButton.Button1)\
            'Finally
            '    ConnectionState = dbConnection.State
            'End Try
        Else
            sqlConnection = New SqlClient.SqlConnection(sConnectionUzytkownicy)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQLUzytkownicy, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLUzytkownicy, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLUzytkownicy, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLUzytkownicy, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Uzytkownicy")

            SQLDeleteUzytkownicy(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateUzytkownicy(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertUzytkownicy(sqlCommandInsert, sqlDataAdapter)

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
                '    MessageBoxDefaultButton.Button1)\
            Finally
                ConnectionState = sqlConnection.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapter.FillSchema(DataSetUzytkownicy, SchemaType.Mapped, "TABELA_Uzytkownicy")
                    'dbDataAdapter.Fill(DataSetUzytkownicy)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetUzytkownicy, SchemaType.Mapped, "TABELA_Uzytkownicy")
                    sqlDataAdapter.Fill(DataSetUzytkownicy)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetUzytkownicy.Tables(0).PrimaryKey = New DataColumn() {DataSetUzytkownicy.Tables(0).Columns("IDENT")}

                'definiuj DataView
                DataViewUzytkownicy = New DataView(DataSetUzytkownicy.Tables(0))
                DataViewUzytkownicy.AllowDelete = False
                DataViewUzytkownicy.AllowNew = False

                If DataSetUzytkownicy.Tables(0).Rows.Count <= 0 Then
                    'Raiseevent(Dodaj,Click )
                    MessageBox.Show("Brak zapisu dot. u¿ytkownika " & user & " w katalogu..." & vbCrLf & _
                        "Nie potrafiê zapisaæ nowego has³a do bazy" & user, _
                        "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                Else
                    Dim foundRow As DataRow
                    ' Create an array for the key values to find.
                    Dim findTheseVals(0) As Object
                    findTheseVals(0) = user
                    foundRow = DataSetUzytkownicy.Tables(0).Rows.Find(findTheseVals)
                    ' sprawdz czy jest...
                    If Not (foundRow Is Nothing) Then
                        'foundRow("IDENT") = oIdentUzytkownika
                        'foundRow("NAZWA") = U_NazwaUzytkownika
                        'foundRow("FUNKCJA") = U_FunkcjaUzytkownika
                        foundRow("HASLO") = pass
                        'foundRow("UPRAWNIENIA") = U_UprawnieniaUzytkownika
                        'foundRow("MAGAZYNY") = U_MagazynyUzytkownika

                        'foundRow("RKROKRAPORTU") = U_RKRokRaportu
                        'foundRow("RKNUMERRAPORTU") = U_RKNumerRaportu
                        'foundRow("RKNUMERPOZYCJI") = U_RKNumerPozycji
                        'foundRow("RKNUMERDOWODUKP") = U_RKNumerDowoduKP
                        'foundRow("RKNUMERDOWODUKW") = U_RKNumerDowoduKW

                        'foundRow("OPIS") = U_OpisUzytkownika           'OPIS      Text 50
                        'foundRow("UWAGI") = U_UwagiUzytkownika         'UWAGI     Text
                        foundRow("WERSJA") = ZmienWersje(foundRow("WERSJA"))    'Wersja         Integer, 2
                        AktualizujBaze()
                    End If
                End If

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

        End If
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
            '        dbDataAdapter.Update(DataSetUzytkownicy, "TABELA_Uzytkownicy")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetUzytkownicy)
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
                    sqlDataAdapter.Update(DataSetUzytkownicy, "TABELA_Uzytkownicy")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetUzytkownicy)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub


End Class
