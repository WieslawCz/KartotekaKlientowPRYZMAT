Public Class PobierzObrotyzTOFIdoAnalizy
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    Friend WithEvents lblIleOdrzucono As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblIleZaktualizowano As System.Windows.Forms.Label
    Friend WithEvents lblIleDopisano As System.Windows.Forms.Label
    Friend WithEvents lblIlePrzeczytano As System.Windows.Forms.Label
    Friend WithEvents lblPlik As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdImportuj As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPokazPlik As System.Windows.Forms.Button
    Friend WithEvents rabNiePrzesuwaj As System.Windows.Forms.RadioButton
    Friend WithEvents rabPrzesun As System.Windows.Forms.RadioButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents txtDataPopAkt As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdData As System.Windows.Forms.Button
    Friend WithEvents txtDataAkt As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtPopOkres As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmdOkres As System.Windows.Forms.Button
    Friend WithEvents txtOkres As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PobierzObrotyzTOFIdoAnalizy))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtPopOkres = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cmdOkres = New System.Windows.Forms.Button()
        Me.txtOkres = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtDataPopAkt = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdData = New System.Windows.Forms.Button()
        Me.txtDataAkt = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.rabNiePrzesuwaj = New System.Windows.Forms.RadioButton()
        Me.rabPrzesun = New System.Windows.Forms.RadioButton()
        Me.cmdPokazPlik = New System.Windows.Forms.Button()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.lblIleOdrzucono = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblIleZaktualizowano = New System.Windows.Forms.Label()
        Me.lblIleDopisano = New System.Windows.Forms.Label()
        Me.lblIlePrzeczytano = New System.Windows.Forms.Label()
        Me.lblPlik = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.cmdImportuj = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.txtPopOkres)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.cmdOkres)
        Me.Panel1.Controls.Add(Me.txtOkres)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.txtDataPopAkt)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.cmdData)
        Me.Panel1.Controls.Add(Me.txtDataAkt)
        Me.Panel1.Controls.Add(Me.Label21)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.rabNiePrzesuwaj)
        Me.Panel1.Controls.Add(Me.rabPrzesun)
        Me.Panel1.Controls.Add(Me.cmdPokazPlik)
        Me.Panel1.Controls.Add(Me.ProgressBar)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.cmdWybierz)
        Me.Panel1.Controls.Add(Me.lblIleOdrzucono)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lblIleZaktualizowano)
        Me.Panel1.Controls.Add(Me.lblIleDopisano)
        Me.Panel1.Controls.Add(Me.lblIlePrzeczytano)
        Me.Panel1.Controls.Add(Me.lblPlik)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(613, 310)
        Me.Panel1.TabIndex = 30
        '
        'txtPopOkres
        '
        Me.txtPopOkres.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtPopOkres.ForeColor = System.Drawing.Color.Purple
        Me.txtPopOkres.Location = New System.Drawing.Point(387, 97)
        Me.txtPopOkres.Name = "txtPopOkres"
        Me.txtPopOkres.ReadOnly = True
        Me.txtPopOkres.Size = New System.Drawing.Size(96, 20)
        Me.txtPopOkres.TabIndex = 74
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(261, 100)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(147, 16)
        Me.Label7.TabIndex = 75
        Me.Label7.Text = "Poprzedni okres rozlicz. . . . . . . . . ."
        '
        'cmdOkres
        '
        Me.cmdOkres.Image = CType(resources.GetObject("cmdOkres.Image"), System.Drawing.Image)
        Me.cmdOkres.Location = New System.Drawing.Point(202, 96)
        Me.cmdOkres.Name = "cmdOkres"
        Me.cmdOkres.Size = New System.Drawing.Size(24, 22)
        Me.cmdOkres.TabIndex = 72
        '
        'txtOkres
        '
        Me.txtOkres.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtOkres.ForeColor = System.Drawing.Color.Purple
        Me.txtOkres.Location = New System.Drawing.Point(112, 97)
        Me.txtOkres.Name = "txtOkres"
        Me.txtOkres.Size = New System.Drawing.Size(96, 20)
        Me.txtOkres.TabIndex = 71
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(8, 100)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 73
        Me.Label8.Text = "Okres rozliczeniowy . . . . . . . . ."
        '
        'txtDataPopAkt
        '
        Me.txtDataPopAkt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtDataPopAkt.ForeColor = System.Drawing.Color.Purple
        Me.txtDataPopAkt.Location = New System.Drawing.Point(387, 72)
        Me.txtDataPopAkt.Name = "txtDataPopAkt"
        Me.txtDataPopAkt.ReadOnly = True
        Me.txtDataPopAkt.Size = New System.Drawing.Size(96, 20)
        Me.txtDataPopAkt.TabIndex = 69
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(261, 75)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(147, 16)
        Me.Label1.TabIndex = 70
        Me.Label1.Text = "Poprzednia aktualizacja . . . . . . . . ."
        '
        'cmdData
        '
        Me.cmdData.Image = CType(resources.GetObject("cmdData.Image"), System.Drawing.Image)
        Me.cmdData.Location = New System.Drawing.Point(202, 71)
        Me.cmdData.Name = "cmdData"
        Me.cmdData.Size = New System.Drawing.Size(24, 22)
        Me.cmdData.TabIndex = 67
        '
        'txtDataAkt
        '
        Me.txtDataAkt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtDataAkt.ForeColor = System.Drawing.Color.Purple
        Me.txtDataAkt.Location = New System.Drawing.Point(112, 72)
        Me.txtDataAkt.Name = "txtDataAkt"
        Me.txtDataAkt.Size = New System.Drawing.Size(96, 20)
        Me.txtDataAkt.TabIndex = 66
        '
        'Label21
        '
        Me.Label21.ForeColor = System.Drawing.Color.Navy
        Me.Label21.Location = New System.Drawing.Point(8, 75)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(112, 16)
        Me.Label21.TabIndex = 68
        Me.Label21.Text = "Data aktualizacji . . . . . . . . ."
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Location = New System.Drawing.Point(0, 167)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(611, 1)
        Me.Panel3.TabIndex = 53
        '
        'rabNiePrzesuwaj
        '
        Me.rabNiePrzesuwaj.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rabNiePrzesuwaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabNiePrzesuwaj.ForeColor = System.Drawing.Color.Navy
        Me.rabNiePrzesuwaj.Location = New System.Drawing.Point(11, 145)
        Me.rabNiePrzesuwaj.Name = "rabNiePrzesuwaj"
        Me.rabNiePrzesuwaj.Size = New System.Drawing.Size(583, 16)
        Me.rabNiePrzesuwaj.TabIndex = 52
        Me.rabNiePrzesuwaj.Text = "Dane z pliku z Tofi wczytaj jako poprzedni miesi¹c (bez przesuwania danych w Anal" &
    "izach)"
        '
        'rabPrzesun
        '
        Me.rabPrzesun.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rabPrzesun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabPrzesun.ForeColor = System.Drawing.Color.Navy
        Me.rabPrzesun.Location = New System.Drawing.Point(11, 123)
        Me.rabPrzesun.Name = "rabPrzesun"
        Me.rabPrzesun.Size = New System.Drawing.Size(583, 16)
        Me.rabPrzesun.TabIndex = 51
        Me.rabPrzesun.Text = "Przesuñ dane w Analizach o jeden miesi¹c i plik z TOFI wczytaj jako poprzedni mie" &
    "si¹c."
        '
        'cmdPokazPlik
        '
        Me.cmdPokazPlik.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazPlik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPokazPlik.ForeColor = System.Drawing.Color.Black
        Me.cmdPokazPlik.Image = CType(resources.GetObject("cmdPokazPlik.Image"), System.Drawing.Image)
        Me.cmdPokazPlik.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazPlik.Location = New System.Drawing.Point(376, 32)
        Me.cmdPokazPlik.Name = "cmdPokazPlik"
        Me.cmdPokazPlik.Size = New System.Drawing.Size(95, 24)
        Me.cmdPokazPlik.TabIndex = 50
        Me.cmdPokazPlik.Text = "Poka¿ plik"
        Me.cmdPokazPlik.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ProgressBar
        '
        Me.ProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar.Location = New System.Drawing.Point(5, 285)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(601, 8)
        Me.ProgressBar.TabIndex = 49
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 64)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(611, 1)
        Me.Panel2.TabIndex = 48
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(477, 32)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(112, 24)
        Me.cmdWybierz.TabIndex = 47
        Me.cmdWybierz.Text = "Wybierz plik"
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblIleOdrzucono
        '
        Me.lblIleOdrzucono.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleOdrzucono.ForeColor = System.Drawing.Color.Purple
        Me.lblIleOdrzucono.Location = New System.Drawing.Point(309, 253)
        Me.lblIleOdrzucono.Name = "lblIleOdrzucono"
        Me.lblIleOdrzucono.Size = New System.Drawing.Size(112, 17)
        Me.lblIleOdrzucono.TabIndex = 30
        Me.lblIleOdrzucono.Text = "0"
        Me.lblIleOdrzucono.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(69, 255)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(240, 16)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "Iloœæ rekordów odrzuconych. . . . . . . . . . . . . . . . ."
        '
        'lblIleZaktualizowano
        '
        Me.lblIleZaktualizowano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleZaktualizowano.ForeColor = System.Drawing.Color.Purple
        Me.lblIleZaktualizowano.Location = New System.Drawing.Point(309, 229)
        Me.lblIleZaktualizowano.Name = "lblIleZaktualizowano"
        Me.lblIleZaktualizowano.Size = New System.Drawing.Size(112, 17)
        Me.lblIleZaktualizowano.TabIndex = 28
        Me.lblIleZaktualizowano.Text = "0"
        Me.lblIleZaktualizowano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIleDopisano
        '
        Me.lblIleDopisano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleDopisano.ForeColor = System.Drawing.Color.Purple
        Me.lblIleDopisano.Location = New System.Drawing.Point(309, 205)
        Me.lblIleDopisano.Name = "lblIleDopisano"
        Me.lblIleDopisano.Size = New System.Drawing.Size(112, 17)
        Me.lblIleDopisano.TabIndex = 27
        Me.lblIleDopisano.Text = "0"
        Me.lblIleDopisano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIlePrzeczytano
        '
        Me.lblIlePrzeczytano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlePrzeczytano.ForeColor = System.Drawing.Color.Purple
        Me.lblIlePrzeczytano.Location = New System.Drawing.Point(309, 181)
        Me.lblIlePrzeczytano.Name = "lblIlePrzeczytano"
        Me.lblIlePrzeczytano.Size = New System.Drawing.Size(112, 17)
        Me.lblIlePrzeczytano.TabIndex = 26
        Me.lblIlePrzeczytano.Text = "0"
        Me.lblIlePrzeczytano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblPlik
        '
        Me.lblPlik.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPlik.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlik.ForeColor = System.Drawing.Color.Purple
        Me.lblPlik.Location = New System.Drawing.Point(72, 8)
        Me.lblPlik.Name = "lblPlik"
        Me.lblPlik.Size = New System.Drawing.Size(517, 17)
        Me.lblPlik.TabIndex = 24
        Me.lblPlik.Text = "Katalog Bazy Danych . . . . . . . . . . . "
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(69, 231)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(240, 16)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Iloœæ rekordów zaktualizowanych . . . . . . . . . . . . . . . . . . . . ."
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(69, 207)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(240, 16)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Iloœæ rekordów dopisanych . . . . . . . . . . . . . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(69, 183)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(240, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Iloœæ rekordów przeczytanych . . . . . . . . . . . . . . . . . . "
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 16)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Plik z TOFI . . . . . .. . ."
        '
        'cmdImportuj
        '
        Me.cmdImportuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdImportuj.Image = CType(resources.GetObject("cmdImportuj.Image"), System.Drawing.Image)
        Me.cmdImportuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdImportuj.Location = New System.Drawing.Point(453, 327)
        Me.cmdImportuj.Name = "cmdImportuj"
        Me.cmdImportuj.Size = New System.Drawing.Size(75, 23)
        Me.cmdImportuj.TabIndex = 35
        Me.cmdImportuj.Text = "Importuj"
        Me.cmdImportuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(534, 327)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(82, 23)
        Me.cmdPowrot.TabIndex = 34
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(15, 325)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 25)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 33
        Me.PictureBox2.TabStop = False
        '
        'PobierzObrotyzTOFIdoAnalizy
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(629, 356)
        Me.Controls.Add(Me.cmdImportuj)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "PobierzObrotyzTOFIdoAnalizy"
        Me.ShowInTaskbar = False
        Me.Text = "Pobieranie pliku z obrotami z TOFI do aktualizacji Analizy Zakupów"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    'OleDbConnection-polaczenie

    Dim dbConnectionKlienci As OleDb.OleDbConnection
    Dim dbCommandSelectKlienci As OleDb.OleDbCommand
    Dim dbCommandDeleteKlienci As OleDb.OleDbCommand
    Dim dbCommandUpdateKlienci As OleDb.OleDbCommand
    Dim dbCommandInsertKlienci As OleDb.OleDbCommand
    Dim dbDataAdapterKlienci As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienci As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienci As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienci As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienci As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienci As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienci As SqlClient.SqlDataAdapter

    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView




    Dim dbSelectSQLAnalizyZakupu As String = sSelectSQLAnalizyZakupu

    Dim dbConnectionAnalizyZakupu As OleDb.OleDbConnection
    Dim dbCommandSelectAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandDeleteAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandUpdateAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandInsertAnalizyZakupu As OleDb.OleDbCommand
    Dim dbDataAdapterAnalizyZakupu As OleDb.OleDbDataAdapter

    Dim sqlConnectionAnalizyZakupu As SqlClient.SqlConnection
    Dim sqlCommandSelectAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandDeleteAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandUpdateAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandInsertAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlDataAdapterAnalizyZakupu As SqlClient.SqlDataAdapter

    Dim DataSetAnalizyZakupu As New DataSet
    Dim DataViewAnalizyZakupu As New DataView

    Dim ConnectionState As System.Data.ConnectionState






    '-------------------------------------
    Dim IlePrzeczytano As Long
    Dim IleDopisano As Long
    Dim IleZaktualizowano As Long
    Dim IleOdrzucono As Long
    Dim IleJest As Long

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub



    Private Sub PobierzObrotyzTOFIdoAnalizy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        IlePrzeczytano = 0
        IleDopisano = 0
        IleZaktualizowano = 0
        IleOdrzucono = 0

        IleJest = 0
        lblPlik.Text = ""

        PobierzParametry()
        txtDataAkt.Text = SysData()
        txtDataPopAkt.Text = Par_DataAktAnalizy
        txtOkres.Text = Par_OstOkresAnalizy
        txtPopOkres.Text = Par_OstOkresAnalizy

        lblIlePrzeczytano.Text = IlePrzeczytano.ToString
        lblIleDopisano.Text = IleDopisano.ToString
        lblIleZaktualizowano.Text = IleZaktualizowano.ToString
        lblIleOdrzucono.Text = IleOdrzucono.ToString

        rabPrzesun.Checked = True




        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionAnalizyZakupu = New OleDb.OleDbConnection(sConnectionAnalizyZakupu)
            dbCommandSelectAnalizyZakupu = New OleDb.OleDbCommand(dbSelectSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
            dbCommandDeleteAnalizyZakupu = New OleDb.OleDbCommand(sDeleteSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
            dbCommandUpdateAnalizyZakupu = New OleDb.OleDbCommand(sUpdateSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
            dbCommandInsertAnalizyZakupu = New OleDb.OleDbCommand(sInsertSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
            dbDataAdapterAnalizyZakupu = New OleDb.OleDbDataAdapter(dbCommandSelectAnalizyZakupu)

            'DataSet
            DataSetAnalizyZakupu.Locale = New System.Globalization.CultureInfo("pl-PL")

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliAnalizyZakupu As System.Data.Common.DataTableMapping
            MapowanieTabeliAnalizyZakupu = dbDataAdapterAnalizyZakupu.TableMappings.Add("Table", "TABELA_AnalizyZakupu")

            'komendy DataAdapter
            Dim objParam As OleDb.OleDbParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParam = dbCommandDeleteAnalizyZakupu.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = dbCommandDeleteAnalizyZakupu.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            dbDataAdapterAnalizyZakupu.DeleteCommand = dbCommandDeleteAnalizyZakupu

            '------------------------------------------
            'komenda UPDATE
            'parametry aktualizacji
            'dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIdentklienta", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa00", OleDb.OleDbType.Double, Nothing, "WARTZA00")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa00", OleDb.OleDbType.Double, Nothing, "ILOSZA00")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZaBM", OleDb.OleDbType.Double, Nothing, "WARTZABM")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZaBM", OleDb.OleDbType.Double, Nothing, "ILOSZABM")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa01", OleDb.OleDbType.Double, Nothing, "WARTZA01")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa01", OleDb.OleDbType.Double, Nothing, "ILOSZA01")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa02", OleDb.OleDbType.Double, Nothing, "WARTZA02")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa02", OleDb.OleDbType.Double, Nothing, "ILOSZA02")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa03", OleDb.OleDbType.Double, Nothing, "WARTZA03")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa03", OleDb.OleDbType.Double, Nothing, "ILOSZA03")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa04", OleDb.OleDbType.Double, Nothing, "WARTZA04")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa04", OleDb.OleDbType.Double, Nothing, "ILOSZA04")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa05", OleDb.OleDbType.Double, Nothing, "WARTZA05")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa05", OleDb.OleDbType.Double, Nothing, "ILOSZA05")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa06", OleDb.OleDbType.Double, Nothing, "WARTZA06")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa06", OleDb.OleDbType.Double, Nothing, "ILOSZA06")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa07", OleDb.OleDbType.Double, Nothing, "WARTZA07")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa07", OleDb.OleDbType.Double, Nothing, "ILOSZA07")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa08", OleDb.OleDbType.Double, Nothing, "WARTZA08")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa08", OleDb.OleDbType.Double, Nothing, "ILOSZA08")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa09", OleDb.OleDbType.Double, Nothing, "WARTZA09")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa09", OleDb.OleDbType.Double, Nothing, "ILOSZA09")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa10", OleDb.OleDbType.Double, Nothing, "WARTZA10")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa10", OleDb.OleDbType.Double, Nothing, "ILOSZA10")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa11", OleDb.OleDbType.Double, Nothing, "WARTZA11")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa11", OleDb.OleDbType.Double, Nothing, "ILOSZA11")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa12", OleDb.OleDbType.Double, Nothing, "WARTZA12")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa12", OleDb.OleDbType.Double, Nothing, "ILOSZA12")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa13", OleDb.OleDbType.Double, Nothing, "WARTZA13")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa13", OleDb.OleDbType.Double, Nothing, "ILOSZA13")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa14", OleDb.OleDbType.Double, Nothing, "WARTZA14")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa14", OleDb.OleDbType.Double, Nothing, "ILOSZA14")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa15", OleDb.OleDbType.Double, Nothing, "WARTZA15")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa15", OleDb.OleDbType.Double, Nothing, "ILOSZA15")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa16", OleDb.OleDbType.Double, Nothing, "WARTZA16")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa16", OleDb.OleDbType.Double, Nothing, "ILOSZA16")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa17", OleDb.OleDbType.Double, Nothing, "WARTZA17")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa17", OleDb.OleDbType.Double, Nothing, "ILOSZA17")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa18", OleDb.OleDbType.Double, Nothing, "WARTZA18")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa18", OleDb.OleDbType.Double, Nothing, "ILOSZA18")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa19", OleDb.OleDbType.Double, Nothing, "WARTZA19")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa19", OleDb.OleDbType.Double, Nothing, "ILOSZA19")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa20", OleDb.OleDbType.Double, Nothing, "WARTZA20")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa20", OleDb.OleDbType.Double, Nothing, "ILOSZA20")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa21", OleDb.OleDbType.Double, Nothing, "WARTZA21")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa21", OleDb.OleDbType.Double, Nothing, "ILOSZA21")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa22", OleDb.OleDbType.Double, Nothing, "WARTZA22")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa22", OleDb.OleDbType.Double, Nothing, "ILOSZA22")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa23", OleDb.OleDbType.Double, Nothing, "WARTZA23")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa23", OleDb.OleDbType.Double, Nothing, "ILOSZA23")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa24", OleDb.OleDbType.Double, Nothing, "WARTZA24")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa24", OleDb.OleDbType.Double, Nothing, "ILOSZA24")
            dbCommandUpdateAnalizyZakupu.Parameters.Add("@oWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")

            'parametry filtrowania
            objParam = dbCommandUpdateAnalizyZakupu.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = dbCommandUpdateAnalizyZakupu.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            dbDataAdapterAnalizyZakupu.UpdateCommand = dbCommandUpdateAnalizyZakupu

            '------------------------------------------
            'komenda INSERT
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIdentklienta", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa00", OleDb.OleDbType.Double, Nothing, "WARTZA00")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa00", OleDb.OleDbType.Double, Nothing, "ILOSZA00")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZaBM", OleDb.OleDbType.Double, Nothing, "WARTZABM")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZaBM", OleDb.OleDbType.Double, Nothing, "ILOSZABM")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa01", OleDb.OleDbType.Double, Nothing, "WARTZA01")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa01", OleDb.OleDbType.Double, Nothing, "ILOSZA01")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa02", OleDb.OleDbType.Double, Nothing, "WARTZA02")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa02", OleDb.OleDbType.Double, Nothing, "ILOSZA02")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa03", OleDb.OleDbType.Double, Nothing, "WARTZA03")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa03", OleDb.OleDbType.Double, Nothing, "ILOSZA03")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa04", OleDb.OleDbType.Double, Nothing, "WARTZA04")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa04", OleDb.OleDbType.Double, Nothing, "ILOSZA04")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa05", OleDb.OleDbType.Double, Nothing, "WARTZA05")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa05", OleDb.OleDbType.Double, Nothing, "ILOSZA05")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa06", OleDb.OleDbType.Double, Nothing, "WARTZA06")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa06", OleDb.OleDbType.Double, Nothing, "ILOSZA06")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa07", OleDb.OleDbType.Double, Nothing, "WARTZA07")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa07", OleDb.OleDbType.Double, Nothing, "ILOSZA07")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa08", OleDb.OleDbType.Double, Nothing, "WARTZA08")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa08", OleDb.OleDbType.Double, Nothing, "ILOSZA08")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa09", OleDb.OleDbType.Double, Nothing, "WARTZA09")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa09", OleDb.OleDbType.Double, Nothing, "ILOSZA09")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa10", OleDb.OleDbType.Double, Nothing, "WARTZA10")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa10", OleDb.OleDbType.Double, Nothing, "ILOSZA10")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa11", OleDb.OleDbType.Double, Nothing, "WARTZA11")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa11", OleDb.OleDbType.Double, Nothing, "ILOSZA11")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa12", OleDb.OleDbType.Double, Nothing, "WARTZA12")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa12", OleDb.OleDbType.Double, Nothing, "ILOSZA12")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa13", OleDb.OleDbType.Double, Nothing, "WARTZA13")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa13", OleDb.OleDbType.Double, Nothing, "ILOSZA13")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa14", OleDb.OleDbType.Double, Nothing, "WARTZA14")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa14", OleDb.OleDbType.Double, Nothing, "ILOSZA14")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa15", OleDb.OleDbType.Double, Nothing, "WARTZA15")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa15", OleDb.OleDbType.Double, Nothing, "ILOSZA15")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa16", OleDb.OleDbType.Double, Nothing, "WARTZA16")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa16", OleDb.OleDbType.Double, Nothing, "ILOSZA16")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa17", OleDb.OleDbType.Double, Nothing, "WARTZA17")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa17", OleDb.OleDbType.Double, Nothing, "ILOSZA17")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa18", OleDb.OleDbType.Double, Nothing, "WARTZA18")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa18", OleDb.OleDbType.Double, Nothing, "ILOSZA18")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa19", OleDb.OleDbType.Double, Nothing, "WARTZA19")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa19", OleDb.OleDbType.Double, Nothing, "ILOSZA19")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa20", OleDb.OleDbType.Double, Nothing, "WARTZA20")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa20", OleDb.OleDbType.Double, Nothing, "ILOSZA20")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa21", OleDb.OleDbType.Double, Nothing, "WARTZA21")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa21", OleDb.OleDbType.Double, Nothing, "ILOSZA21")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa22", OleDb.OleDbType.Double, Nothing, "WARTZA22")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa22", OleDb.OleDbType.Double, Nothing, "ILOSZA22")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa23", OleDb.OleDbType.Double, Nothing, "WARTZA23")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa23", OleDb.OleDbType.Double, Nothing, "ILOSZA23")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa24", OleDb.OleDbType.Double, Nothing, "WARTZA24")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa24", OleDb.OleDbType.Double, Nothing, "ILOSZA24")
            objParam.SourceVersion = DataRowVersion.Current

            objParam = dbCommandInsertAnalizyZakupu.Parameters.Add("@oWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Current
            dbDataAdapterAnalizyZakupu.InsertCommand = dbCommandInsertAnalizyZakupu

            ' Add the RowUpdated event handler.
            AddHandler dbDataAdapterAnalizyZakupu.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionAnalizyZakupu.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnectionAnalizyZakupu.State
            End Try

        Else
            sqlConnectionAnalizyZakupu = New SqlClient.SqlConnection(sConnectionAnalizyZakupu)
            sqlCommandSelectAnalizyZakupu = New SqlClient.SqlCommand(dbSelectSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
            sqlCommandDeleteAnalizyZakupu = New SqlClient.SqlCommand(sDeleteSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
            sqlCommandUpdateAnalizyZakupu = New SqlClient.SqlCommand(sUpdateSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
            sqlCommandInsertAnalizyZakupu = New SqlClient.SqlCommand(sInsertSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
            sqlDataAdapterAnalizyZakupu = New SqlClient.SqlDataAdapter(sqlCommandSelectAnalizyZakupu)

            'DataSet
            DataSetAnalizyZakupu.Locale = New System.Globalization.CultureInfo("pl-PL")

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliAnalizyZakupu As System.Data.Common.DataTableMapping
            MapowanieTabeliAnalizyZakupu = sqlDataAdapterAnalizyZakupu.TableMappings.Add("Table", "TABELA_AnalizyZakupu")

            'komendy DataAdapter
            Dim objParam As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParam = sqlCommandDeleteAnalizyZakupu.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandDeleteAnalizyZakupu.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapterAnalizyZakupu.DeleteCommand = sqlCommandDeleteAnalizyZakupu

            '------------------------------------------


            'komenda UPDATE
            'parametry aktualizacji
            'sqlcommandUpdateAnalizyZakupu.Parameters.Add("@oIdentKlienta", sqldbtype.Char, 6, "IDENTKLIENTA")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa00", SqlDbType.Float, Nothing, "WARTZA00")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa00", SqlDbType.Float, Nothing, "ILOSZA00")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZaBM", SqlDbType.Float, Nothing, "WARTZABM")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZaBM", SqlDbType.Float, Nothing, "ILOSZABM")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa01", SqlDbType.Float, Nothing, "WARTZA01")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa01", SqlDbType.Float, Nothing, "ILOSZA01")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa02", SqlDbType.Float, Nothing, "WARTZA02")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa02", SqlDbType.Float, Nothing, "ILOSZA02")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa03", SqlDbType.Float, Nothing, "WARTZA03")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa03", SqlDbType.Float, Nothing, "ILOSZA03")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa04", SqlDbType.Float, Nothing, "WARTZA04")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa04", SqlDbType.Float, Nothing, "ILOSZA04")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa05", SqlDbType.Float, Nothing, "WARTZA05")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa05", SqlDbType.Float, Nothing, "ILOSZA05")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa06", SqlDbType.Float, Nothing, "WARTZA06")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa06", SqlDbType.Float, Nothing, "ILOSZA06")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa07", SqlDbType.Float, Nothing, "WARTZA07")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa07", SqlDbType.Float, Nothing, "ILOSZA07")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa08", SqlDbType.Float, Nothing, "WARTZA08")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa08", SqlDbType.Float, Nothing, "ILOSZA08")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa09", SqlDbType.Float, Nothing, "WARTZA09")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa09", SqlDbType.Float, Nothing, "ILOSZA09")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa10", SqlDbType.Float, Nothing, "WARTZA10")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa10", SqlDbType.Float, Nothing, "ILOSZA10")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa11", SqlDbType.Float, Nothing, "WARTZA11")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa11", SqlDbType.Float, Nothing, "ILOSZA11")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa12", SqlDbType.Float, Nothing, "WARTZA12")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa12", SqlDbType.Float, Nothing, "ILOSZA12")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa13", SqlDbType.Float, Nothing, "WARTZA13")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa13", SqlDbType.Float, Nothing, "ILOSZA13")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa14", SqlDbType.Float, Nothing, "WARTZA14")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa14", SqlDbType.Float, Nothing, "ILOSZA14")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa15", SqlDbType.Float, Nothing, "WARTZA15")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa15", SqlDbType.Float, Nothing, "ILOSZA15")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa16", SqlDbType.Float, Nothing, "WARTZA16")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa16", SqlDbType.Float, Nothing, "ILOSZA16")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa17", SqlDbType.Float, Nothing, "WARTZA17")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa17", SqlDbType.Float, Nothing, "ILOSZA17")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa18", SqlDbType.Float, Nothing, "WARTZA18")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa18", SqlDbType.Float, Nothing, "ILOSZA18")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa19", SqlDbType.Float, Nothing, "WARTZA19")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa19", SqlDbType.Float, Nothing, "ILOSZA19")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa20", SqlDbType.Float, Nothing, "WARTZA20")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa20", SqlDbType.Float, Nothing, "ILOSZA20")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa21", SqlDbType.Float, Nothing, "WARTZA21")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa21", SqlDbType.Float, Nothing, "ILOSZA21")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa22", SqlDbType.Float, Nothing, "WARTZA22")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa22", SqlDbType.Float, Nothing, "ILOSZA22")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa23", SqlDbType.Float, Nothing, "WARTZA23")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa23", SqlDbType.Float, Nothing, "ILOSZA23")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWartZa24", SqlDbType.Float, Nothing, "WARTZA24")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oIlosZa24", SqlDbType.Float, Nothing, "ILOSZA24")
            sqlCommandUpdateAnalizyZakupu.Parameters.Add("@oWersja", SqlDbType.Int, Nothing, "WERSJA")

            'parametry filtrowania
            objParam = sqlCommandUpdateAnalizyZakupu.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandUpdateAnalizyZakupu.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapterAnalizyZakupu.UpdateCommand = sqlCommandUpdateAnalizyZakupu

            '------------------------------------------
            'komenda INSERT
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIdentKlienta", SqlDbType.Char, 6, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa00", SqlDbType.Float, Nothing, "WARTZA00")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa00", SqlDbType.Float, Nothing, "ILOSZA00")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZaBM", SqlDbType.Float, Nothing, "WARTZABM")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZaBM", SqlDbType.Float, Nothing, "ILOSZABM")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa01", SqlDbType.Float, Nothing, "WARTZA01")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa01", SqlDbType.Float, Nothing, "ILOSZA01")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa02", SqlDbType.Float, Nothing, "WARTZA02")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa02", SqlDbType.Float, Nothing, "ILOSZA02")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa03", SqlDbType.Float, Nothing, "WARTZA03")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa03", SqlDbType.Float, Nothing, "ILOSZA03")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa04", SqlDbType.Float, Nothing, "WARTZA04")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa04", SqlDbType.Float, Nothing, "ILOSZA04")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa05", SqlDbType.Float, Nothing, "WARTZA05")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa05", SqlDbType.Float, Nothing, "ILOSZA05")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa06", SqlDbType.Float, Nothing, "WARTZA06")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa06", SqlDbType.Float, Nothing, "ILOSZA06")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa07", SqlDbType.Float, Nothing, "WARTZA07")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa07", SqlDbType.Float, Nothing, "ILOSZA07")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa08", SqlDbType.Float, Nothing, "WARTZA08")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa08", SqlDbType.Float, Nothing, "ILOSZA08")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa09", SqlDbType.Float, Nothing, "WARTZA09")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa09", SqlDbType.Float, Nothing, "ILOSZA09")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa10", SqlDbType.Float, Nothing, "WARTZA10")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa10", SqlDbType.Float, Nothing, "ILOSZA10")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa11", SqlDbType.Float, Nothing, "WARTZA11")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa11", SqlDbType.Float, Nothing, "ILOSZA11")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa12", SqlDbType.Float, Nothing, "WARTZA12")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa12", SqlDbType.Float, Nothing, "ILOSZA12")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa13", SqlDbType.Float, Nothing, "WARTZA13")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa13", SqlDbType.Float, Nothing, "ILOSZA13")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa14", SqlDbType.Float, Nothing, "WARTZA14")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa14", SqlDbType.Float, Nothing, "ILOSZA14")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa15", SqlDbType.Float, Nothing, "WARTZA15")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa15", SqlDbType.Float, Nothing, "ILOSZA15")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa16", SqlDbType.Float, Nothing, "WARTZA16")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa16", SqlDbType.Float, Nothing, "ILOSZA16")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa17", SqlDbType.Float, Nothing, "WARTZA17")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa17", SqlDbType.Float, Nothing, "ILOSZA17")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa18", SqlDbType.Float, Nothing, "WARTZA18")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa18", SqlDbType.Float, Nothing, "ILOSZA18")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa19", SqlDbType.Float, Nothing, "WARTZA19")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa19", SqlDbType.Float, Nothing, "ILOSZA19")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa20", SqlDbType.Float, Nothing, "WARTZA20")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa20", SqlDbType.Float, Nothing, "ILOSZA20")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa21", SqlDbType.Float, Nothing, "WARTZA21")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa21", SqlDbType.Float, Nothing, "ILOSZA21")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa22", SqlDbType.Float, Nothing, "WARTZA22")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa22", SqlDbType.Float, Nothing, "ILOSZA22")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa23", SqlDbType.Float, Nothing, "WARTZA23")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa23", SqlDbType.Float, Nothing, "ILOSZA23")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWartZa24", SqlDbType.Float, Nothing, "WARTZA24")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oIlosZa24", SqlDbType.Float, Nothing, "ILOSZA24")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAnalizyZakupu.Parameters.Add("@oWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Current
            sqlDataAdapterAnalizyZakupu.InsertCommand = sqlCommandInsertAnalizyZakupu

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterAnalizyZakupu.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionAnalizyZakupu.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionAnalizyZakupu.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterAnalizyZakupu.FillSchema(DataSetAnalizyZakupu, SchemaType.Mapped, "TABELA_AnalizyZakupu")
                    dbDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                    dbConnectionAnalizyZakupu.Close()
                Else
                    sqlDataAdapterAnalizyZakupu.FillSchema(DataSetAnalizyZakupu, SchemaType.Mapped, "TABELA_AnalizyZakupu")
                    sqlDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                    sqlConnectionAnalizyZakupu.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetAnalizyZakupu.Tables(0).PrimaryKey = New DataColumn() {DataSetAnalizyZakupu.Tables(0).Columns("IDENTKLIENTA")}
                'definiuj DataView
                DataViewAnalizyZakupu = New DataView(DataSetAnalizyZakupu.Tables(0))

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If
    End Sub


    Private Sub cmdWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierz.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z TOFI..."
            .InitialDirectory = "c:\"
            .DefaultExt = "txt"
            .Filter = "Plik tekstowy (*.txt)|*.txt|Wszystkie pliki|*.*"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                lblPlik.Text = .FileName
            End If
        End With
    End Sub


    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************


    Private Sub AktualizujBazeanalizyZakupu()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnectionAnalizyZakupu.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnectionAnalizyZakupu.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapterAnalizyZakupu.Update(DataSetAnalizyZakupu, "TABELA_analizyZakupu")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                End If
                dbConnectionAnalizyZakupu.Close()
            End If
        Else
            Try
                sqlConnectionAnalizyZakupu.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionAnalizyZakupu.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterAnalizyZakupu.Update(DataSetAnalizyZakupu, "TABELA_analizyZakupu")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                End If
                sqlConnectionAnalizyZakupu.Close()
            End If
        End If
    End Sub



    '*****************************************************
    '** Importujemy dane
    '*****************************************************


    Private Sub cmdImportuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImportuj.Click
        Dim dbSelectSQLKlienci As String = ""
        Dim DataSetKlienci As DataSet
        Dim DataViewKlienci As DataView
        Dim SymbolKlienta As String = ""

        Dim TxtPlik As String
        Dim FileNum As Integer
        Dim NextLine As String
        Dim BiezacyRekord As Long
        Dim IleRec As Long
        Dim ip As Integer
        Dim AnalizyRow As DataRow
        Dim Nierozpoz As String = ""

        Dim SymbTOFI As Long = 0
        Dim SymbKlienta As String = ""
        Dim DataTOFI As String = ""
        Dim MiesiacTOFI As String = ""
        Dim DataObrotu As String = ""

        Dim AIlosc As Integer
        Dim AWartosc As Double
        Dim LIlosc As Integer
        Dim LWartosc As Double

        Dim AoIlosc As Integer
        Dim AoWartosc As Double
        Dim LoIlosc As Integer
        Dim LoWartosc As Double

        Dim ik As Integer = 0
        Dim NewRow As DataRow

        IlePrzeczytano = 0
        IleDopisano = 0
        IleZaktualizowano = 0
        IleOdrzucono = 0
        IleJest = 0

        lblIlePrzeczytano.Text = IlePrzeczytano.ToString
        lblIleDopisano.Text = IleDopisano.ToString
        lblIleZaktualizowano.Text = IleZaktualizowano.ToString
        lblIleOdrzucono.Text = IleOdrzucono.ToString

        TxtPlik = Trim(lblPlik.Text)
        If Len(TxtPlik) = 0 Then
            MessageBox.Show("Brak definicji pliku z TOFI do pobrania", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If Len(Dir(TxtPlik)) = 0 Then
                MessageBox.Show("Nie znalaz³em na dysku pliku do pobrania" & vbCrLf & TxtPlik & vbCrLf, _
                    "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Else
                cmdImportuj.Enabled = False

                Dim idr As Integer = 0
                If rabPrzesun.Checked Then
                    'przesun dane w bazie Analizy
                    If DataSetAnalizyZakupu.Tables(0).Rows.Count > 0 Then
                        For idr = 0 To DataSetAnalizyZakupu.Tables(0).Rows.Count - 1
                            aIdentAnal = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("IDENTKLIENTA")
                            aWart00 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA00")
                            aIlos00 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA00")
                            aWartBM = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZABM")
                            aIlosBM = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZABM")
                            aWart01 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA01")
                            aIlos01 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA01")
                            aWart02 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA02")
                            aIlos02 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA02")
                            aWart03 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA03")
                            aIlos03 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA03")
                            aWart04 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA04")
                            aIlos04 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA04")
                            aWart05 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA05")
                            aIlos05 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA05")
                            aWart06 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA06")
                            aIlos06 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA06")
                            aWart07 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA07")
                            aIlos07 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA07")
                            aWart08 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA08")
                            aIlos08 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA08")
                            aWart09 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA09")
                            aIlos09 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA09")
                            aWart10 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA10")
                            aIlos10 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA10")
                            aWart11 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA11")
                            aIlos11 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA11")
                            aWart12 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA12")
                            aIlos12 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA12")
                            aWart13 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA13")
                            aIlos13 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA13")
                            aWart14 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA14")
                            aIlos14 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA14")
                            aWart15 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA15")
                            aIlos15 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA15")
                            aWart16 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA16")
                            aIlos16 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA16")
                            aWart17 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA17")
                            aIlos17 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA17")
                            aWart18 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA18")
                            aIlos18 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA18")
                            aWart19 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA19")
                            aIlos19 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA19")
                            aWart20 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA20")
                            aIlos20 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA20")
                            aWart21 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA21")
                            aIlos21 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA21")
                            aWart22 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA22")
                            aIlos22 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA22")
                            aWart23 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA23")
                            aIlos23 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA23")
                            aWart24 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA24")
                            aIlos24 = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA24")
                            aWersjaAnal = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WERSJA")

                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA00") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA00") = 0

                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZABM") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZABM") = 0

                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA01") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA01") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA02") = aWart01
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA02") = aIlos01
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA03") = aWart02
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA03") = aIlos02
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA04") = aWart03
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA04") = aIlos03
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA05") = aWart04
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA05") = aIlos04
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA06") = aWart05
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA06") = aIlos05
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA07") = aWart06
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA07") = aIlos06
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA08") = aWart07
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA08") = aIlos07
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA09") = aWart08
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA09") = aIlos08
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA10") = aWart09
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA10") = aIlos09
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA11") = aWart10
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA11") = aIlos10
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA12") = aWart11
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA12") = aIlos11

                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA13") = aWart12
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA13") = aIlos12
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA14") = aWart13
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA14") = aIlos13
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA15") = aWart14
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA15") = aIlos14
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA16") = aWart15
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA16") = aIlos15
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA17") = aWart16
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA17") = aIlos16
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA18") = aWart17
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA18") = aIlos17
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA19") = aWart18
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA19") = aIlos18
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA20") = aWart19
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA20") = aIlos19
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA21") = aWart20
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA21") = aIlos20
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA22") = aWart21
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA22") = aIlos21
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA23") = aWart22
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA23") = aIlos22
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA24") = aWart23
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA24") = aIlos23

                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WERSJA") = ZmienWersje(aWersjaAnal)
                        Next
                        AktualizujBazeanalizyZakupu()
                    End If
                Else
                    'nie przesuwaj - zrób miejsce na nowe dane
                    If DataSetAnalizyZakupu.Tables(0).Rows.Count > 0 Then
                        For idr = 0 To DataSetAnalizyZakupu.Tables(0).Rows.Count - 1
                            aWersjaAnal = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WERSJA")

                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA00") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA00") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZABM") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZABM") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZA01") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZA01") = 0
                            DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WERSJA") = ZmienWersje(aWersjaAnal)
                        Next
                        AktualizujBazeanalizyZakupu()
                    End If
                End If




                'jest plik - otwieramy
                BiezacyRekord = 0
                FileNum = FreeFile()    'kolejny nr otwartego zbioru

                'wylicz ile ma wierszy
                FileOpen(FileNum, TxtPlik, OpenMode.Input)
                IleRec = 0
                Do Until EOF(FileNum)
                    NextLine = LineInput(FileNum)
                    IleRec += 1
                Loop
                FileClose(FileNum)

                'analizuj plik tekstowy
                FileOpen(FileNum, TxtPlik, OpenMode.Input)
                ProgressBar.Maximum = IleRec
                Do Until EOF(FileNum)
                    NextLine = LineInput(FileNum)
                    BiezacyRekord = BiezacyRekord + 1
                    ProgressBar.Value = BiezacyRekord

                    SymbTOFI = 0
                    DataTOFI = ""

                    AIlosc = 0
                    AWartosc = 0
                    LIlosc = 0
                    LWartosc = 0

                    AoIlosc = 0
                    AoWartosc = 0
                    LoIlosc = 0
                    LoWartosc = 0

                    'analizuj pobrany rekord
                    ip = InStr(NextLine, ",")
                    If ip > 0 Then
                        SymbTOFI = CLng(Mid(NextLine, 1, ip - 1))
                        NextLine = Mid(NextLine, ip + 1)
                        ip = InStr(NextLine, ",")
                        If ip > 0 Then
                            DataTOFI = Mid(NextLine, 1, ip - 1)
                            NextLine = Mid(NextLine, ip + 1)
                            ip = InStr(NextLine, ",")
                            If ip > 0 Then
                                AIlosc = Val(Mid(NextLine, 1, ip - 1))
                                NextLine = Mid(NextLine, ip + 1)
                                ip = InStr(NextLine, ",")
                                If ip > 0 Then
                                    AWartosc = Val(Mid(NextLine, 1, ip - 1))
                                    NextLine = Mid(NextLine, ip + 1)
                                    ip = InStr(NextLine, ",")
                                    If ip > 0 Then
                                        LIlosc = Val(Mid(NextLine, 1, ip - 1))
                                        NextLine = Mid(NextLine, ip + 1)
                                        ip = InStr(NextLine, ",")
                                        If ip > 0 Then
                                            LWartosc = Val(Mid(NextLine, 1, ip - 1))
                                            NextLine = Mid(NextLine, ip + 1)
                                            ip = InStr(NextLine, ",")
                                            If ip > 0 Then
                                                AoIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                NextLine = Mid(NextLine, ip + 1)
                                                ip = InStr(NextLine, ",")
                                                If ip > 0 Then
                                                    AoWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                    NextLine = Mid(NextLine, ip + 1)
                                                    ip = InStr(NextLine, ",")
                                                    If ip > 0 Then
                                                        LoIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                        LoWartosc = Val(Mid(NextLine, ip + 1))
                                                    End If
                                                End If
                                            End If
                                        Else
                                            LWartosc = Val(NextLine)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                    DataObrotu = Mid(DataTOFI, 1, 4) & "-" & Mid(DataTOFI, 5, 2) & "-" & Mid(DataTOFI, 7, 2)
                    SymbKlienta = ""
                    IlePrzeczytano += 1

                    'przetlumacz symbol TOFI na symbol klienta
                    If SymbTOFI = 0 Then
                    Else
                        DataSetKlienci = New DataSet
                        DataViewKlienci = New DataView
                        SymbolKlienta = ""

                        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

                        If ParL_DataBaseType = "MS ACCESS" Then
                            dbSelectSQLKlienci = "SELECT * FROM Klienci WHERE NRTOFITXT LIKE '%" & Microsoft.VisualBasic.Right("00000" & Trim(Str(SymbTOFI)), 5) & "%' "
                            dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
                            dbCommandSelectKlienci = New OleDb.OleDbCommand(dbSelectSQLKlienci, dbConnectionKlienci)
                            dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)
                            Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
                            MapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
                            '------------------------------------------
                            Try
                                dbConnectionKlienci.Open()
                            Catch Ex As System.Exception
                            Finally
                                ConnectionState = dbConnectionKlienci.State
                            End Try
                        Else
                            dbSelectSQLKlienci = "SELECT * FROM Klienci WHERE NRTOFITXT LIKE '%" & Microsoft.VisualBasic.Right("00000" & Trim(Str(SymbTOFI)), 5) & "%' "
                            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
                            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectSQLKlienci, sqlConnectionKlienci)
                            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)
                            Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
                            MapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
                            '------------------------------------------
                            Try
                                sqlConnectionKlienci.Open()
                            Catch Ex As System.Exception
                                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                                    System.Windows.Forms.MessageBoxButtons.OK, _
                                    MessageBoxIcon.Information, _
                                    MessageBoxDefaultButton.Button1)
                            Finally
                                ConnectionState = sqlConnectionKlienci.State
                            End Try
                        End If

                        If ConnectionState = ConnectionState.Open Then
                            Try
                                If ParL_DataBaseType = "MS ACCESS" Then
                                    dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                                    dbDataAdapterKlienci.Fill(DataSetKlienci)
                                    dbConnectionKlienci.Close()
                                Else
                                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                                    sqlConnectionKlienci.Close()
                                End If

                                'definiuj DataView
                                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                                If DataViewKlienci.Count > 0 Then
                                    'mamy klientów z takim symbolem TOFI
                                    For ik = 0 To DataViewKlienci.Count - 1
                                        SymbKlienta = DataViewKlienci.Item(ik).Item(0)

                                        '---------------------------------
                                        ' Create an array for the key values to find.
                                        Dim findAnalizy(0) As Object
                                        findAnalizy(0) = SymbKlienta
                                        AnalizyRow = DataSetAnalizyZakupu.Tables(0).Rows.Find(findAnalizy)
                                        ' sprawdz czy jest...
                                        If Not (AnalizyRow Is Nothing) Then
                                            aWersjaAnal = AnalizyRow.Item("WERSJA")
                                            aWart00 = AnalizyRow.Item("WARTZA00")
                                            aIlos00 = AnalizyRow.Item("ILOSZA00")
                                            aWart01 = AnalizyRow.Item("WARTZA01")
                                            aIlos01 = AnalizyRow.Item("ILOSZA01")

                                            AnalizyRow.Item("WARTZA00") = 0
                                            AnalizyRow.Item("ILOSZA00") = 0
                                            AnalizyRow.Item("WARTZABM") = 0
                                            AnalizyRow.Item("ILOSZABM") = 0
                                            AnalizyRow.Item("WARTZA01") = aWart01 + (AWartosc + LWartosc + AoWartosc + LoWartosc)
                                            AnalizyRow.Item("ILOSZA01") = aIlos01 + LIlosc
                                            AnalizyRow.Item("WERSJA") = ZmienWersje(aWersjaAnal)

                                            IleZaktualizowano += 1
                                        Else
                                            oOperacja = "Nowy"
                                            oWersjaKli = 0
                                            'stworz nowy rekord
                                            NewRow = DataSetAnalizyZakupu.Tables(0).NewRow()

                                            NewRow.Item("IDENTKLIENTA") = SymbKlienta

                                            NewRow.Item("WARTZA00") = 0
                                            NewRow.Item("ILOSZA00") = 0

                                            NewRow.Item("WARTZABM") = 0
                                            NewRow.Item("ILOSZABM") = 0

                                            NewRow.Item("WARTZA01") = AWartosc + LWartosc + AoWartosc + LoWartosc
                                            NewRow.Item("ILOSZA01") = LIlosc
                                            NewRow.Item("WARTZA02") = 0
                                            NewRow.Item("ILOSZA02") = 0
                                            NewRow.Item("WARTZA03") = 0
                                            NewRow.Item("ILOSZA03") = 0
                                            NewRow.Item("WARTZA04") = 0
                                            NewRow.Item("ILOSZA04") = 0
                                            NewRow.Item("WARTZA05") = 0
                                            NewRow.Item("ILOSZA05") = 0
                                            NewRow.Item("WARTZA06") = 0
                                            NewRow.Item("ILOSZA06") = 0
                                            NewRow.Item("WARTZA07") = 0
                                            NewRow.Item("ILOSZA07") = 0
                                            NewRow.Item("WARTZA08") = 0
                                            NewRow.Item("ILOSZA08") = 0
                                            NewRow.Item("WARTZA09") = 0
                                            NewRow.Item("ILOSZA09") = 0
                                            NewRow.Item("WARTZA10") = 0
                                            NewRow.Item("ILOSZA10") = 0
                                            NewRow.Item("WARTZA11") = 0
                                            NewRow.Item("ILOSZA11") = 0
                                            NewRow.Item("WARTZA12") = 0
                                            NewRow.Item("ILOSZA12") = 0

                                            NewRow.Item("WARTZA13") = 0
                                            NewRow.Item("ILOSZA13") = 0
                                            NewRow.Item("WARTZA14") = 0
                                            NewRow.Item("ILOSZA14") = 0
                                            NewRow.Item("WARTZA15") = 0
                                            NewRow.Item("ILOSZA15") = 0
                                            NewRow.Item("WARTZA16") = 0
                                            NewRow.Item("ILOSZA16") = 0
                                            NewRow.Item("WARTZA17") = 0
                                            NewRow.Item("ILOSZA17") = 0
                                            NewRow.Item("WARTZA18") = 0
                                            NewRow.Item("ILOSZA18") = 0
                                            NewRow.Item("WARTZA19") = 0
                                            NewRow.Item("ILOSZA19") = 0
                                            NewRow.Item("WARTZA20") = 0
                                            NewRow.Item("ILOSZA20") = 0
                                            NewRow.Item("WARTZA21") = 0
                                            NewRow.Item("ILOSZA21") = 0
                                            NewRow.Item("WARTZA22") = 0
                                            NewRow.Item("ILOSZA22") = 0
                                            NewRow.Item("WARTZA23") = 0
                                            NewRow.Item("ILOSZA23") = 0
                                            NewRow.Item("WARTZA24") = 0
                                            NewRow.Item("ILOSZA24") = 0
                                            NewRow.Item("WERSJA") = 1

                                            'dodaj rekord do DataSet
                                            IleDopisano += 1
                                            DataSetAnalizyZakupu.Tables(0).Rows.Add(NewRow)
                                        End If
                                        AktualizujBazeanalizyZakupu()
                                        '---------------------------------
                                    Next
                                Else
                                    Nierozpoz &= SymbTOFI & " "
                                End If

                            Catch Ex As System.Exception
                            End Try
                        End If
                    End If

                    lblIlePrzeczytano.Text = IlePrzeczytano.ToString
                    lblIleDopisano.Text = IleDopisano.ToString
                    lblIleZaktualizowano.Text = IleZaktualizowano.ToString
                    System.Windows.Forms.Application.DoEvents() 'pozwol na wykonanie
                Loop
                FileClose(FileNum)
                'AktualizujWartoscObrotowwKartKli()
                cmdImportuj.Enabled = True
                '-------------------------------
                Par_DataAktAnalizy = txtDataAkt.Text
                Par_OstOkresAnalizy = txtOkres.Text
                ZapiszParametry(Me)

                MessageBox.Show("Pobra³em informacje o obrotach" & vbCrLf & _
                    "przekazane z programu TOFI " & vbCrLf & _
                    IIf(Len(Nierozpoz) > 0, "Nierozpoznane symbole TOFI :" & vbCrLf & Nierozpoz, ""), _
                    "OK, skoñczy³em", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
                Me.Close()
            End If
        End If
    End Sub

    Private Sub PobierzObrotyzTOFIdoAnalizy_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub



    Private Sub cmdPokazPlik_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokazPlik.Click
        Dim TxtPlik As String = Trim(lblPlik.Text)

        If Len(TxtPlik) = 0 Then
            MessageBox.Show("Brak definicji pliku z TOFI", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If Len(Dir(TxtPlik)) = 0 Then
                MessageBox.Show("Nie znalaz³em na dysku pliku z TOFI" & vbCrLf & TxtPlik & vbCrLf, _
                    "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Else
                'Shell("notepad", AppWinStyle.NormalFocus, False)
                Dim proc As New Process
                System.Windows.Forms.Application.DoEvents()
                proc.StartInfo.FileName = "Notepad.exe"
                proc.StartInfo.Arguments = TxtPlik
                proc.Start()
            End If
        End If
    End Sub

    Private Sub cmdData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdData.Click
        oData = txtDataAkt.Text
        Dim Proc As New WybierzDate
        Proc.ShowDialog()
        txtDataAkt.Text = oData
    End Sub


    Private Sub cmdOkres_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOkres.Click
        If Len(txtOkres.Text) = 7 Then
            oData = txtOkres.Text & "-01"
        Else
            oData = txtOkres.Text
        End If
        Dim Proc As New WybierzDate
        Proc.ShowDialog()
        txtOkres.Text = Microsoft.VisualBasic.Left(oData, 7)
    End Sub

    Private Sub txtDataAkt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDataAkt.TextChanged
        TestLen(txtDataAkt, "DATA AKTUALIZACJI", 10)
    End Sub
    Private Sub txtDataAkt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataAkt.GotFocus
        StartEdycjiTxt(txtDataAkt)
    End Sub
    Private Sub txtDataAkt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataAkt.LostFocus
        KoniecEdycjiTxt(txtDataAkt)
    End Sub

    Private Sub txtOkres_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOkres.TextChanged
        TestLen(txtOkres, "OKRES AKTUALIZACJI", 7)
    End Sub
    Private Sub txtOkres_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOkres.GotFocus
        StartEdycjiTxt(txtOkres)
    End Sub
    Private Sub txtOkres_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOkres.LostFocus
        KoniecEdycjiTxt(txtOkres)
    End Sub

End Class

