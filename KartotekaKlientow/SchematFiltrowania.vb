Public Class SchematFiltrowania
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Dim _Opcja As String = ""
    Public Sub New(ByVal Opcja As String)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        _Opcja = Opcja
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
    Friend WithEvents ListaPol As System.Windows.Forms.ListView
    Friend WithEvents Nazwa As System.Windows.Forms.ColumnHeader
    Friend WithEvents Typ As System.Windows.Forms.ColumnHeader
    Friend WithEvents Filtr As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa As System.Windows.Forms.Label
    Friend WithEvents lblTyp As System.Windows.Forms.Label
    Friend WithEvents pnlText As System.Windows.Forms.Panel
    Friend WithEvents pnlData As System.Windows.Forms.Panel
    Friend WithEvents pnlTakNie As System.Windows.Forms.Panel
    Friend WithEvents pnlLiczba As System.Windows.Forms.Panel
    Friend WithEvents rbtTak As System.Windows.Forms.RadioButton
    Friend WithEvents rbtNie As System.Windows.Forms.RadioButton
    Friend WithEvents txtTextDo As System.Windows.Forms.TextBox
    Friend WithEvents txtDataOd As System.Windows.Forms.TextBox
    Friend WithEvents TxtLiczbaOd As System.Windows.Forms.TextBox
    Friend WithEvents cmdDataOd As System.Windows.Forms.Button
    Friend WithEvents cmdDataDo As System.Windows.Forms.Button
    Friend WithEvents txtLiczbaDo As System.Windows.Forms.TextBox
    Friend WithEvents txtDataDo As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblPath As System.Windows.Forms.Label
    Friend WithEvents txtSchemat As System.Windows.Forms.TextBox
    Friend WithEvents cmdZapisz As System.Windows.Forms.Button
    Friend WithEvents cmdCzysc As System.Windows.Forms.Button
    Friend WithEvents cmdWycofajsie As System.Windows.Forms.Button
    Friend WithEvents chbTestTxt As System.Windows.Forms.CheckBox
    Friend WithEvents lblTextDo As System.Windows.Forms.Label
    Friend WithEvents chbTestNum As System.Windows.Forms.CheckBox
    Friend WithEvents chbTestLog As System.Windows.Forms.CheckBox
    Friend WithEvents chbTestDate As System.Windows.Forms.CheckBox
    Friend WithEvents lblLiczbaDo As System.Windows.Forms.Label
    Friend WithEvents lblLiczbaOd As System.Windows.Forms.Label
    Friend WithEvents lblDataDo As System.Windows.Forms.Label
    Friend WithEvents lblDataOd As System.Windows.Forms.Label
    Friend WithEvents lblTak As System.Windows.Forms.Label
    Friend WithEvents lblNie As System.Windows.Forms.Label
    Friend WithEvents rabtextPuste As System.Windows.Forms.RadioButton
    Friend WithEvents rabTextNiepuste As System.Windows.Forms.RadioButton
    Friend WithEvents rabtextZawiera As System.Windows.Forms.RadioButton
    Friend WithEvents Opis As System.Windows.Forms.ColumnHeader
    Friend WithEvents rabTexOdDo As System.Windows.Forms.RadioButton
    Friend WithEvents txtTextOd As System.Windows.Forms.TextBox
    Friend WithEvents lblTextOd As System.Windows.Forms.Label
    Friend WithEvents rabLiczbaRozne As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaRowne As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaOdDo As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaWieRowne As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaWieksze As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaMniejRowne As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaMniesze As System.Windows.Forms.RadioButton
    Friend WithEvents rabTextNieZawiera As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SchematFiltrowania))
        Me.ListaPol = New System.Windows.Forms.ListView()
        Me.Nazwa = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Opis = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Typ = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Filtr = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblTyp = New System.Windows.Forms.Label()
        Me.lblNazwa = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSchemat = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pnlText = New System.Windows.Forms.Panel()
        Me.rabTexOdDo = New System.Windows.Forms.RadioButton()
        Me.txtTextOd = New System.Windows.Forms.TextBox()
        Me.lblTextOd = New System.Windows.Forms.Label()
        Me.rabTextNieZawiera = New System.Windows.Forms.RadioButton()
        Me.rabtextZawiera = New System.Windows.Forms.RadioButton()
        Me.rabTextNiepuste = New System.Windows.Forms.RadioButton()
        Me.rabtextPuste = New System.Windows.Forms.RadioButton()
        Me.chbTestTxt = New System.Windows.Forms.CheckBox()
        Me.txtTextDo = New System.Windows.Forms.TextBox()
        Me.lblTextDo = New System.Windows.Forms.Label()
        Me.pnlData = New System.Windows.Forms.Panel()
        Me.chbTestDate = New System.Windows.Forms.CheckBox()
        Me.cmdDataDo = New System.Windows.Forms.Button()
        Me.cmdDataOd = New System.Windows.Forms.Button()
        Me.txtDataDo = New System.Windows.Forms.TextBox()
        Me.lblDataDo = New System.Windows.Forms.Label()
        Me.txtDataOd = New System.Windows.Forms.TextBox()
        Me.lblDataOd = New System.Windows.Forms.Label()
        Me.pnlTakNie = New System.Windows.Forms.Panel()
        Me.rbtNie = New System.Windows.Forms.RadioButton()
        Me.lblNie = New System.Windows.Forms.Label()
        Me.chbTestLog = New System.Windows.Forms.CheckBox()
        Me.rbtTak = New System.Windows.Forms.RadioButton()
        Me.lblTak = New System.Windows.Forms.Label()
        Me.pnlLiczba = New System.Windows.Forms.Panel()
        Me.rabLiczbaOdDo = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaWieRowne = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaWieksze = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaMniejRowne = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaMniesze = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaRozne = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaRowne = New System.Windows.Forms.RadioButton()
        Me.chbTestNum = New System.Windows.Forms.CheckBox()
        Me.txtLiczbaDo = New System.Windows.Forms.TextBox()
        Me.lblLiczbaDo = New System.Windows.Forms.Label()
        Me.TxtLiczbaOd = New System.Windows.Forms.TextBox()
        Me.lblLiczbaOd = New System.Windows.Forms.Label()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.cmdZapisz = New System.Windows.Forms.Button()
        Me.cmdCzysc = New System.Windows.Forms.Button()
        Me.cmdWycofajsie = New System.Windows.Forms.Button()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.pnlText.SuspendLayout()
        Me.pnlData.SuspendLayout()
        Me.pnlTakNie.SuspendLayout()
        Me.pnlLiczba.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListaPol
        '
        Me.ListaPol.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListaPol.BackColor = System.Drawing.SystemColors.Control
        Me.ListaPol.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Nazwa, Me.Opis, Me.Typ, Me.Filtr})
        Me.ListaPol.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.ListaPol.ForeColor = System.Drawing.Color.Purple
        Me.ListaPol.FullRowSelect = True
        Me.ListaPol.GridLines = True
        Me.ListaPol.HideSelection = False
        Me.ListaPol.Location = New System.Drawing.Point(8, 64)
        Me.ListaPol.MultiSelect = False
        Me.ListaPol.Name = "ListaPol"
        Me.ListaPol.Size = New System.Drawing.Size(624, 223)
        Me.ListaPol.TabIndex = 1
        Me.ListaPol.UseCompatibleStateImageBehavior = False
        Me.ListaPol.View = System.Windows.Forms.View.Details
        '
        'Nazwa
        '
        Me.Nazwa.Text = "Nazwa pola"
        Me.Nazwa.Width = 0
        '
        'Opis
        '
        Me.Opis.Text = "Nazwa pola"
        Me.Opis.Width = 220
        '
        'Typ
        '
        Me.Typ.Text = "Typ"
        Me.Typ.Width = 80
        '
        'Filtr
        '
        Me.Filtr.Text = "Filtr"
        Me.Filtr.Width = 400
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(554, 369)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 22)
        Me.cmdPowrot.TabIndex = 15
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(16, 369)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 14
        Me.PictureBox2.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblTyp)
        Me.Panel1.Controls.Add(Me.lblNazwa)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 287)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(208, 75)
        Me.Panel1.TabIndex = 16
        '
        'lblTyp
        '
        Me.lblTyp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTyp.ForeColor = System.Drawing.Color.Navy
        Me.lblTyp.Location = New System.Drawing.Point(80, 32)
        Me.lblTyp.Name = "lblTyp"
        Me.lblTyp.Size = New System.Drawing.Size(112, 16)
        Me.lblTyp.TabIndex = 22
        '
        'lblNazwa
        '
        Me.lblNazwa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa.ForeColor = System.Drawing.Color.Navy
        Me.lblNazwa.Location = New System.Drawing.Point(80, 8)
        Me.lblNazwa.Name = "lblNazwa"
        Me.lblNazwa.Size = New System.Drawing.Size(112, 16)
        Me.lblNazwa.TabIndex = 21
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Typ pola . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Nazwa pola . . . . . . . . ."
        '
        'txtSchemat
        '
        Me.txtSchemat.ForeColor = System.Drawing.Color.Purple
        Me.txtSchemat.Location = New System.Drawing.Point(96, 8)
        Me.txtSchemat.Name = "txtSchemat"
        Me.txtSchemat.Size = New System.Drawing.Size(96, 20)
        Me.txtSchemat.TabIndex = 17
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Nazwa szablonu . . . . . . . . ."
        '
        'pnlText
        '
        Me.pnlText.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pnlText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlText.Controls.Add(Me.rabTexOdDo)
        Me.pnlText.Controls.Add(Me.txtTextOd)
        Me.pnlText.Controls.Add(Me.lblTextOd)
        Me.pnlText.Controls.Add(Me.rabTextNieZawiera)
        Me.pnlText.Controls.Add(Me.rabtextZawiera)
        Me.pnlText.Controls.Add(Me.rabTextNiepuste)
        Me.pnlText.Controls.Add(Me.rabtextPuste)
        Me.pnlText.Controls.Add(Me.chbTestTxt)
        Me.pnlText.Controls.Add(Me.txtTextDo)
        Me.pnlText.Controls.Add(Me.lblTextDo)
        Me.pnlText.Location = New System.Drawing.Point(224, 287)
        Me.pnlText.Name = "pnlText"
        Me.pnlText.Size = New System.Drawing.Size(408, 75)
        Me.pnlText.TabIndex = 19
        '
        'rabTexOdDo
        '
        Me.rabTexOdDo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabTexOdDo.ForeColor = System.Drawing.Color.Navy
        Me.rabTexOdDo.Location = New System.Drawing.Point(168, 46)
        Me.rabTexOdDo.Name = "rabTexOdDo"
        Me.rabTexOdDo.Size = New System.Drawing.Size(96, 16)
        Me.rabTexOdDo.TabIndex = 32
        Me.rabTexOdDo.Text = "Od..Do"
        '
        'txtTextOd
        '
        Me.txtTextOd.ForeColor = System.Drawing.Color.Purple
        Me.txtTextOd.Location = New System.Drawing.Point(294, 8)
        Me.txtTextOd.Name = "txtTextOd"
        Me.txtTextOd.Size = New System.Drawing.Size(90, 20)
        Me.txtTextOd.TabIndex = 31
        '
        'lblTextOd
        '
        Me.lblTextOd.ForeColor = System.Drawing.Color.Navy
        Me.lblTextOd.Location = New System.Drawing.Point(237, 11)
        Me.lblTextOd.Name = "lblTextOd"
        Me.lblTextOd.Size = New System.Drawing.Size(149, 16)
        Me.lblTextOd.TabIndex = 30
        Me.lblTextOd.Text = "Tekst od :"
        '
        'rabTextNieZawiera
        '
        Me.rabTextNieZawiera.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabTextNieZawiera.ForeColor = System.Drawing.Color.Navy
        Me.rabTextNieZawiera.Location = New System.Drawing.Point(78, 46)
        Me.rabTextNieZawiera.Name = "rabTextNieZawiera"
        Me.rabTextNieZawiera.Size = New System.Drawing.Size(96, 16)
        Me.rabTextNieZawiera.TabIndex = 29
        Me.rabTextNieZawiera.Text = "Nie zawiera"
        '
        'rabtextZawiera
        '
        Me.rabtextZawiera.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabtextZawiera.ForeColor = System.Drawing.Color.Navy
        Me.rabtextZawiera.Location = New System.Drawing.Point(78, 30)
        Me.rabtextZawiera.Name = "rabtextZawiera"
        Me.rabtextZawiera.Size = New System.Drawing.Size(96, 16)
        Me.rabtextZawiera.TabIndex = 28
        Me.rabtextZawiera.Text = "Zawiera text"
        '
        'rabTextNiepuste
        '
        Me.rabTextNiepuste.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabTextNiepuste.ForeColor = System.Drawing.Color.Navy
        Me.rabTextNiepuste.Location = New System.Drawing.Point(8, 46)
        Me.rabTextNiepuste.Name = "rabTextNiepuste"
        Me.rabTextNiepuste.Size = New System.Drawing.Size(88, 16)
        Me.rabTextNiepuste.TabIndex = 27
        Me.rabTextNiepuste.Text = "Niepuste"
        '
        'rabtextPuste
        '
        Me.rabtextPuste.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabtextPuste.ForeColor = System.Drawing.Color.Navy
        Me.rabtextPuste.Location = New System.Drawing.Point(8, 30)
        Me.rabtextPuste.Name = "rabtextPuste"
        Me.rabtextPuste.Size = New System.Drawing.Size(64, 16)
        Me.rabtextPuste.TabIndex = 26
        Me.rabtextPuste.Text = "Puste"
        '
        'chbTestTxt
        '
        Me.chbTestTxt.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chbTestTxt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chbTestTxt.ForeColor = System.Drawing.Color.Navy
        Me.chbTestTxt.Location = New System.Drawing.Point(8, 8)
        Me.chbTestTxt.Name = "chbTestTxt"
        Me.chbTestTxt.Size = New System.Drawing.Size(128, 16)
        Me.chbTestTxt.TabIndex = 25
        Me.chbTestTxt.Text = "Testuj pole tekstowe"
        '
        'txtTextDo
        '
        Me.txtTextDo.ForeColor = System.Drawing.Color.Purple
        Me.txtTextDo.Location = New System.Drawing.Point(294, 32)
        Me.txtTextDo.Name = "txtTextDo"
        Me.txtTextDo.Size = New System.Drawing.Size(90, 20)
        Me.txtTextDo.TabIndex = 24
        '
        'lblTextDo
        '
        Me.lblTextDo.ForeColor = System.Drawing.Color.Navy
        Me.lblTextDo.Location = New System.Drawing.Point(237, 32)
        Me.lblTextDo.Name = "lblTextDo"
        Me.lblTextDo.Size = New System.Drawing.Size(149, 16)
        Me.lblTextDo.TabIndex = 23
        Me.lblTextDo.Text = "Tekst do :"
        '
        'pnlData
        '
        Me.pnlData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pnlData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlData.Controls.Add(Me.chbTestDate)
        Me.pnlData.Controls.Add(Me.cmdDataDo)
        Me.pnlData.Controls.Add(Me.cmdDataOd)
        Me.pnlData.Controls.Add(Me.txtDataDo)
        Me.pnlData.Controls.Add(Me.lblDataDo)
        Me.pnlData.Controls.Add(Me.txtDataOd)
        Me.pnlData.Controls.Add(Me.lblDataOd)
        Me.pnlData.Location = New System.Drawing.Point(112, 208)
        Me.pnlData.Name = "pnlData"
        Me.pnlData.Size = New System.Drawing.Size(408, 64)
        Me.pnlData.TabIndex = 20
        '
        'chbTestDate
        '
        Me.chbTestDate.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chbTestDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chbTestDate.ForeColor = System.Drawing.Color.Navy
        Me.chbTestDate.Location = New System.Drawing.Point(8, 8)
        Me.chbTestDate.Name = "chbTestDate"
        Me.chbTestDate.Size = New System.Drawing.Size(144, 16)
        Me.chbTestDate.TabIndex = 68
        Me.chbTestDate.Text = "Testuj pole z dat¹ . . . . . . . . . . . . "
        '
        'cmdDataDo
        '
        Me.cmdDataDo.Image = CType(resources.GetObject("cmdDataDo.Image"), System.Drawing.Image)
        Me.cmdDataDo.Location = New System.Drawing.Point(352, 31)
        Me.cmdDataDo.Name = "cmdDataDo"
        Me.cmdDataDo.Size = New System.Drawing.Size(24, 22)
        Me.cmdDataDo.TabIndex = 67
        Me.cmdDataDo.Visible = False
        '
        'cmdDataOd
        '
        Me.cmdDataOd.Image = CType(resources.GetObject("cmdDataOd.Image"), System.Drawing.Image)
        Me.cmdDataOd.Location = New System.Drawing.Point(352, 7)
        Me.cmdDataOd.Name = "cmdDataOd"
        Me.cmdDataOd.Size = New System.Drawing.Size(24, 22)
        Me.cmdDataOd.TabIndex = 66
        Me.cmdDataOd.Visible = False
        '
        'txtDataDo
        '
        Me.txtDataDo.ForeColor = System.Drawing.Color.Purple
        Me.txtDataDo.Location = New System.Drawing.Point(248, 32)
        Me.txtDataDo.Name = "txtDataDo"
        Me.txtDataDo.Size = New System.Drawing.Size(104, 20)
        Me.txtDataDo.TabIndex = 26
        '
        'lblDataDo
        '
        Me.lblDataDo.ForeColor = System.Drawing.Color.Navy
        Me.lblDataDo.Location = New System.Drawing.Point(192, 32)
        Me.lblDataDo.Name = "lblDataDo"
        Me.lblDataDo.Size = New System.Drawing.Size(120, 16)
        Me.lblDataDo.TabIndex = 25
        Me.lblDataDo.Text = "Do daty : "
        '
        'txtDataOd
        '
        Me.txtDataOd.ForeColor = System.Drawing.Color.Purple
        Me.txtDataOd.Location = New System.Drawing.Point(248, 8)
        Me.txtDataOd.Name = "txtDataOd"
        Me.txtDataOd.Size = New System.Drawing.Size(104, 20)
        Me.txtDataOd.TabIndex = 24
        '
        'lblDataOd
        '
        Me.lblDataOd.ForeColor = System.Drawing.Color.Navy
        Me.lblDataOd.Location = New System.Drawing.Point(192, 8)
        Me.lblDataOd.Name = "lblDataOd"
        Me.lblDataOd.Size = New System.Drawing.Size(120, 16)
        Me.lblDataOd.TabIndex = 23
        Me.lblDataOd.Text = "Od daty : "
        '
        'pnlTakNie
        '
        Me.pnlTakNie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pnlTakNie.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlTakNie.Controls.Add(Me.rbtNie)
        Me.pnlTakNie.Controls.Add(Me.lblNie)
        Me.pnlTakNie.Controls.Add(Me.chbTestLog)
        Me.pnlTakNie.Controls.Add(Me.rbtTak)
        Me.pnlTakNie.Controls.Add(Me.lblTak)
        Me.pnlTakNie.Location = New System.Drawing.Point(112, 143)
        Me.pnlTakNie.Name = "pnlTakNie"
        Me.pnlTakNie.Size = New System.Drawing.Size(408, 65)
        Me.pnlTakNie.TabIndex = 21
        '
        'rbtNie
        '
        Me.rbtNie.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtNie.ForeColor = System.Drawing.Color.Navy
        Me.rbtNie.Location = New System.Drawing.Point(296, 32)
        Me.rbtNie.Name = "rbtNie"
        Me.rbtNie.Size = New System.Drawing.Size(104, 16)
        Me.rbtNie.TabIndex = 21
        Me.rbtNie.Text = "= Nie"
        '
        'lblNie
        '
        Me.lblNie.ForeColor = System.Drawing.Color.Navy
        Me.lblNie.Location = New System.Drawing.Point(192, 32)
        Me.lblNie.Name = "lblNie"
        Me.lblNie.Size = New System.Drawing.Size(120, 16)
        Me.lblNie.TabIndex = 28
        Me.lblNie.Text = "Wybierz wartoœæ :"
        '
        'chbTestLog
        '
        Me.chbTestLog.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chbTestLog.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chbTestLog.ForeColor = System.Drawing.Color.Navy
        Me.chbTestLog.Location = New System.Drawing.Point(8, 8)
        Me.chbTestLog.Name = "chbTestLog"
        Me.chbTestLog.Size = New System.Drawing.Size(144, 16)
        Me.chbTestLog.TabIndex = 27
        Me.chbTestLog.Text = "Testuj pole logiczne . . . . . . . . . "
        '
        'rbtTak
        '
        Me.rbtTak.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtTak.ForeColor = System.Drawing.Color.Navy
        Me.rbtTak.Location = New System.Drawing.Point(296, 8)
        Me.rbtTak.Name = "rbtTak"
        Me.rbtTak.Size = New System.Drawing.Size(104, 16)
        Me.rbtTak.TabIndex = 20
        Me.rbtTak.Text = "= Tak"
        '
        'lblTak
        '
        Me.lblTak.ForeColor = System.Drawing.Color.Navy
        Me.lblTak.Location = New System.Drawing.Point(192, 8)
        Me.lblTak.Name = "lblTak"
        Me.lblTak.Size = New System.Drawing.Size(120, 16)
        Me.lblTak.TabIndex = 19
        Me.lblTak.Text = "Wybierz wartoœæ :"
        '
        'pnlLiczba
        '
        Me.pnlLiczba.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pnlLiczba.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlLiczba.Controls.Add(Me.rabLiczbaOdDo)
        Me.pnlLiczba.Controls.Add(Me.rabLiczbaWieRowne)
        Me.pnlLiczba.Controls.Add(Me.rabLiczbaWieksze)
        Me.pnlLiczba.Controls.Add(Me.rabLiczbaMniejRowne)
        Me.pnlLiczba.Controls.Add(Me.rabLiczbaMniesze)
        Me.pnlLiczba.Controls.Add(Me.rabLiczbaRozne)
        Me.pnlLiczba.Controls.Add(Me.rabLiczbaRowne)
        Me.pnlLiczba.Controls.Add(Me.chbTestNum)
        Me.pnlLiczba.Controls.Add(Me.txtLiczbaDo)
        Me.pnlLiczba.Controls.Add(Me.lblLiczbaDo)
        Me.pnlLiczba.Controls.Add(Me.TxtLiczbaOd)
        Me.pnlLiczba.Controls.Add(Me.lblLiczbaOd)
        Me.pnlLiczba.Location = New System.Drawing.Point(112, 79)
        Me.pnlLiczba.Name = "pnlLiczba"
        Me.pnlLiczba.Size = New System.Drawing.Size(408, 64)
        Me.pnlLiczba.TabIndex = 22
        '
        'rabLiczbaOdDo
        '
        Me.rabLiczbaOdDo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaOdDo.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaOdDo.Location = New System.Drawing.Point(152, 42)
        Me.rabLiczbaOdDo.Name = "rabLiczbaOdDo"
        Me.rabLiczbaOdDo.Size = New System.Drawing.Size(73, 16)
        Me.rabLiczbaOdDo.TabIndex = 34
        Me.rabLiczbaOdDo.Text = "Od..Do"
        '
        'rabLiczbaWieRowne
        '
        Me.rabLiczbaWieRowne.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaWieRowne.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaWieRowne.Location = New System.Drawing.Point(104, 42)
        Me.rabLiczbaWieRowne.Name = "rabLiczbaWieRowne"
        Me.rabLiczbaWieRowne.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaWieRowne.TabIndex = 33
        Me.rabLiczbaWieRowne.Text = ">="
        '
        'rabLiczbaWieksze
        '
        Me.rabLiczbaWieksze.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaWieksze.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaWieksze.Location = New System.Drawing.Point(104, 26)
        Me.rabLiczbaWieksze.Name = "rabLiczbaWieksze"
        Me.rabLiczbaWieksze.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaWieksze.TabIndex = 32
        Me.rabLiczbaWieksze.Text = ">"
        '
        'rabLiczbaMniejRowne
        '
        Me.rabLiczbaMniejRowne.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaMniejRowne.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaMniejRowne.Location = New System.Drawing.Point(54, 42)
        Me.rabLiczbaMniejRowne.Name = "rabLiczbaMniejRowne"
        Me.rabLiczbaMniejRowne.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaMniejRowne.TabIndex = 31
        Me.rabLiczbaMniejRowne.Text = "<="
        '
        'rabLiczbaMniesze
        '
        Me.rabLiczbaMniesze.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaMniesze.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaMniesze.Location = New System.Drawing.Point(54, 26)
        Me.rabLiczbaMniesze.Name = "rabLiczbaMniesze"
        Me.rabLiczbaMniesze.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaMniesze.TabIndex = 30
        Me.rabLiczbaMniesze.Text = "<"
        '
        'rabLiczbaRozne
        '
        Me.rabLiczbaRozne.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaRozne.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaRozne.Location = New System.Drawing.Point(8, 42)
        Me.rabLiczbaRozne.Name = "rabLiczbaRozne"
        Me.rabLiczbaRozne.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaRozne.TabIndex = 29
        Me.rabLiczbaRozne.Text = "<>"
        '
        'rabLiczbaRowne
        '
        Me.rabLiczbaRowne.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaRowne.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaRowne.Location = New System.Drawing.Point(8, 26)
        Me.rabLiczbaRowne.Name = "rabLiczbaRowne"
        Me.rabLiczbaRowne.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaRowne.TabIndex = 28
        Me.rabLiczbaRowne.Text = "="
        '
        'chbTestNum
        '
        Me.chbTestNum.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chbTestNum.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chbTestNum.ForeColor = System.Drawing.Color.Navy
        Me.chbTestNum.Location = New System.Drawing.Point(8, 8)
        Me.chbTestNum.Name = "chbTestNum"
        Me.chbTestNum.Size = New System.Drawing.Size(144, 16)
        Me.chbTestNum.TabIndex = 26
        Me.chbTestNum.Text = "Testuj pole numeryczne"
        '
        'txtLiczbaDo
        '
        Me.txtLiczbaDo.ForeColor = System.Drawing.Color.Purple
        Me.txtLiczbaDo.Location = New System.Drawing.Point(296, 32)
        Me.txtLiczbaDo.Name = "txtLiczbaDo"
        Me.txtLiczbaDo.Size = New System.Drawing.Size(88, 20)
        Me.txtLiczbaDo.TabIndex = 22
        Me.txtLiczbaDo.Visible = False
        '
        'lblLiczbaDo
        '
        Me.lblLiczbaDo.ForeColor = System.Drawing.Color.Navy
        Me.lblLiczbaDo.Location = New System.Drawing.Point(222, 35)
        Me.lblLiczbaDo.Name = "lblLiczbaDo"
        Me.lblLiczbaDo.Size = New System.Drawing.Size(120, 16)
        Me.lblLiczbaDo.TabIndex = 21
        Me.lblLiczbaDo.Text = "Do wartoœci :"
        '
        'TxtLiczbaOd
        '
        Me.TxtLiczbaOd.ForeColor = System.Drawing.Color.Purple
        Me.TxtLiczbaOd.Location = New System.Drawing.Point(296, 8)
        Me.TxtLiczbaOd.Name = "TxtLiczbaOd"
        Me.TxtLiczbaOd.Size = New System.Drawing.Size(88, 20)
        Me.TxtLiczbaOd.TabIndex = 20
        Me.TxtLiczbaOd.Visible = False
        '
        'lblLiczbaOd
        '
        Me.lblLiczbaOd.ForeColor = System.Drawing.Color.Navy
        Me.lblLiczbaOd.Location = New System.Drawing.Point(222, 11)
        Me.lblLiczbaOd.Name = "lblLiczbaOd"
        Me.lblLiczbaOd.Size = New System.Drawing.Size(120, 16)
        Me.lblLiczbaOd.TabIndex = 19
        Me.lblLiczbaOd.Text = "Od wartoœci :"
        '
        'lblPath
        '
        Me.lblPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPath.ForeColor = System.Drawing.Color.Navy
        Me.lblPath.Location = New System.Drawing.Point(40, 32)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(560, 20)
        Me.lblPath.TabIndex = 23
        '
        'Label11
        '
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(8, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(24, 16)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Plik z szablonem . . . . . . . . ."
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(600, 32)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(32, 20)
        Me.Button1.TabIndex = 25
        '
        'cmdZapisz
        '
        Me.cmdZapisz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZapisz.Image = CType(resources.GetObject("cmdZapisz.Image"), System.Drawing.Image)
        Me.cmdZapisz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZapisz.Location = New System.Drawing.Point(327, 369)
        Me.cmdZapisz.Name = "cmdZapisz"
        Me.cmdZapisz.Size = New System.Drawing.Size(107, 22)
        Me.cmdZapisz.TabIndex = 27
        Me.cmdZapisz.Text = "Zapisz schemat"
        Me.cmdZapisz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCzysc
        '
        Me.cmdCzysc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCzysc.Image = CType(resources.GetObject("cmdCzysc.Image"), System.Drawing.Image)
        Me.cmdCzysc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCzysc.Location = New System.Drawing.Point(184, 368)
        Me.cmdCzysc.Name = "cmdCzysc"
        Me.cmdCzysc.Size = New System.Drawing.Size(137, 22)
        Me.cmdCzysc.TabIndex = 28
        Me.cmdCzysc.Text = "Wyczyœæ schemat"
        Me.cmdCzysc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdWycofajsie
        '
        Me.cmdWycofajsie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajsie.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWycofajsie.Image = CType(resources.GetObject("cmdWycofajsie.Image"), System.Drawing.Image)
        Me.cmdWycofajsie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajsie.Location = New System.Drawing.Point(443, 369)
        Me.cmdWycofajsie.Name = "cmdWycofajsie"
        Me.cmdWycofajsie.Size = New System.Drawing.Size(104, 22)
        Me.cmdWycofajsie.TabIndex = 46
        Me.cmdWycofajsie.Text = "Wycofaj siê"
        Me.cmdWycofajsie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'SchematFiltrowania
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(641, 396)
        Me.Controls.Add(Me.cmdWycofajsie)
        Me.Controls.Add(Me.cmdCzysc)
        Me.Controls.Add(Me.cmdZapisz)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lblPath)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.pnlLiczba)
        Me.Controls.Add(Me.pnlTakNie)
        Me.Controls.Add(Me.pnlData)
        Me.Controls.Add(Me.pnlText)
        Me.Controls.Add(Me.txtSchemat)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.ListaPol)
        Me.Name = "SchematFiltrowania"
        Me.ShowInTaskbar = False
        Me.Text = "Schemat filtrowania katalogu klientów"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.pnlText.ResumeLayout(False)
        Me.pnlText.PerformLayout()
        Me.pnlData.ResumeLayout(False)
        Me.pnlData.PerformLayout()
        Me.pnlTakNie.ResumeLayout(False)
        Me.pnlLiczba.ResumeLayout(False)
        Me.pnlLiczba.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


#End Region

    Private CurItemIndex As Integer
    Private PlikZSzablonem As String

    Dim Okres1 As String = ""
    Dim Okres2 As String = ""

    Private Sub SchematFiltrowania_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        pnlText.Visible = False

        pnlLiczba.Visible = False
        pnlLiczba.Location = New Point(pnlText.Location.X, pnlText.Location.Y)
        pnlLiczba.Size = New Size(pnlText.Size.Width, pnlText.Size.Height)

        pnlData.Visible = False
        pnlData.Location = New Point(pnlText.Location.X, pnlText.Location.Y)
        pnlData.Size = New Size(pnlText.Size.Width, pnlText.Size.Height)

        pnlTakNie.Visible = False
        pnlTakNie.Location = New Point(pnlText.Location.X, pnlText.Location.Y)
        pnlTakNie.Size = New Size(pnlText.Size.Width, pnlText.Size.Height)


        InitSzablon()
        Dim Plik As String = oKatParametry & "\" & "_ostatni.sf"
        If Len(Dir(Plik)) > 0 Then
            'ZNALAZLEM PLIK Z OSTATNIM SZABLONEM
            txtSchemat.Text = "_ostatni"
            PlikZSzablonem = "_ostatni.sf"
            PobierzSzablon()
            DefListePol()
            AktListePol()
        Else
            'BRAK PLIKU Z OSTATNIM SZABLONEM
            txtSchemat.Text = ""
            InitSzablon()
            DefListePol()
        End If
        ListaPol.Focus()
        ListaPol.Select()

        '-----------------------------------------------------------------
        If oCoMamRobic = "WYBIERAJ" Then
            cmdWycofajsie.Visible = True
            cmdWycofajsie.Enabled = True
            cmdPowrot.Text = "Wybierz"
        Else
            cmdWycofajsie.Visible = False
            cmdWycofajsie.Enabled = False
            cmdPowrot.Text = "Powrót"
        End If
    End Sub


    Private Sub SchematFiltrowania_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub SchematFiltrowania_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.ClosedByControlBox Then
            'e.Cancel = True     'nie pozwalaj zamkn¹c forme
            If MessageBox.Show("Zdecydowa³eœ opuœciæ formê bez zapisania wprowadzonych zmian." & vbCrLf & _
                "Czy mam zapisaæ zmiany w Bazie Danych ?", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question, _
                MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

                'ZAPISZ SCHEMAT....
                PlikZSzablonem = "_ostatni.sf"
                ZapiszSzablon()
                oSchematFiltrowania = BudujSzablon()
                oNazwaSchematu = txtSchemat.Text
            Else
                oNazwaSchematu = ""
                oSchematFiltrowania = ""
            End If
            'Me._closeClick = False
        Else
            'MsgBox("ZAPISZ lub WYCOFAJ SIÊ")
        End If
    End Sub
    '====================================================


    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        'ZAPISZ SCHEMAT....
        PlikZSzablonem = "_ostatni.sf"
        ZapiszSzablon()
        '------------------------------------------------
        oSchematFiltrowania = BudujSzablon()
        oNazwaSchematu = txtSchemat.Text
        Me.Close()
    End Sub

    Private Sub cmdWycofajsie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajsie.Click
        'oNazwaSchematu = ""
        'oSchematFiltrowania = ""
        Me.Close()
    End Sub

    Private Sub cmdDataOd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDataOd.Click
        oData = txtDataOd.Text
        Dim Proc As New WybierzDate
        Proc.ShowDialog()
        txtDataOd.Text = oData
    End Sub

    Private Sub cmdDataDo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDataDo.Click
        oData = txtDataDo.Text
        Dim Proc As New WybierzDate
        Proc.ShowDialog()
        txtDataDo.Text = oData
    End Sub

    Private Sub DefListePol()
        'ListaPol.Clear()
        ListaPol.Visible = True
        ListaPol.Enabled = True
        ListaPol.ForeColor = System.Drawing.Color.Purple


        ''---------------------------------------------------------------------
        ''Katalog Klientow
        'Public oIdentKli As String         'IDENTKLIENTA   Text, 6
        '---------Public oNrTOFIKli As Long          'NRTOFI         Long
        'Public oNrTOFITxtKli As String     'NRTOFITXT      Text 500

        'Public oKartaPaybackKli As Boolean 'KARTAPAYBACK   Logical
        'Public oNryKartPaybackKli As String 'NRYKARTPAYBACK   memo

        'Public oNazwa1Kli As String        'NAZWA1         Text, 100
        'Public oKodPoczKli As String       'KODPOCZTOWY    Text, 10
        'Public oMiejscKli As String        'MIEJSCOWOSC    Text, 50
        'Public oUlicaKli As String         'ULICA          Text, 50
        'Public oNumNrDomuKli As Integer    'NUMNRDOMU      INTEGER
        '--------Public oNrDomuKli As String        'NRDOMU         Text, 10
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
        'Public oZainstalowanoMonitorKli As Boolean      'ZAINSTALOWANOMONITOR    Logical
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
        '--------Public oIloSprzetuKli As String    'ILOSPRZETU     Text, 100 -> 250/300    !!!!!!!!!!!!
        '--------Public oIloSprzetu2Kli As String   'ILOSPRZETU2     Memo
        'Public oOfertaTZlozeniaKli As String        'OFERTATZLOZENIA         Text, 4
        'Public oOfertaPlikKli As String    'OFERTAPLIK     Text, 100
        'Public oUwagikli As String         'UWAGI          Memo

        'Public oSkupUwagikli As String        'SKUPUWAGI      Memo
        'Public oSprzedOpiekunkli As String    'SPRZEDOPIEKUN    Text, 10   uzytkownik

        'Public oSprzedOKontaktRKli As String  'SPRZEDOKONTAKTR  Text, 4    rok
        'Public oSprzedOKontaktTKli As String  'SPRZEDOKONTAKTT  Integer    nr tygodnia
        'Public oSprzedOKontaktDKli As String  'SPRZEDOKONTAKTD  Text, 12   dzien tygodnia
        'Public oSprzedNKontaktRKli As String  'SPRZEDNKONTAKTR  Text, 4    rok
        'Public oSprzedNKontaktTKli As String  'SPRZEDNKONTAKTT  Integer    nr tygodnia
        'Public oSprzedNKontaktDKli As String  'SKUPNKONTAKTT    Text, 12    nr tygodnia

        'Public oSprzedUwagiKli As String      'SPRZEDUWAGI      Memo
        '--------Public oWersjaKli As Integer          'WERSJA           Integer
        '--------Public oMarkerKli As Boolean          'MARKER           boolean
        '--------Public oMarkerPKli As Boolean         'MARKERP          boolean
        'Public oAktywnyKli As Boolean         'AKTYWNY          boolean

        '--------Public oZakupPopRokKli As Double       'ZAKUPPOPROK        Double
        '--------Public oUslugiDodKli As String         'USLUGIDOD          Text, 200
        '--------Public oRZURap01 As String              'RZURAP01     Text, 100
        '--------Public oRZURap02 As String              'RZURAP02     Text, 100
        '--------Public oRZURap03 As String              'RZURAP03     Text, 100
        '--------Public oRZURap04 As String              'RZURAP04     Text, 100
        '--------Public oRZURap05 As String              'RZURAP05     Text, 100
        '--------Public oRZURap06 As String              'RZURAP06     Text, 100
        '--------Public oRZURap07 As String              'RZURAP07     Text, 100
        '--------Public oRZURap08 As String              'RZURAP08     Text, 100
        '--------Public oRZURap09 As String              'RZURAP09     Text, 100
        ''-------------------------------------------------



        ''----------------------------------------
        '-----Public fiIdentKon As String            'IDENTKLIENTA     Text, 6
        'Public fiDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
        'Public fiPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
        'Public fiTematKon As String            'TEMAT            Text, 50
        'Public fiUwagiKon As String            'UWAGI            Memo
        '-----Public fiWersjaKon As Integer          'WERSJA           Integer



        'Public oUniqIdKon As String           'UNIQID            varchar(40)
        'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
        'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
        'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
        'Public oTematKon As String            'TEMAT            Text, 50
        'Public oUwagiKon As String            'UWAGI            Memo
        'Public oWersjaKon As Integer          'WERSJA           Integer

        '"KH.IDENTKLIENTA, " & _
        '"KH.DATAKONTAKTU," & _
        '"KH.PRACOWNIK," & _
        '"KH.TEMAT," & _
        '"KH.UWAGI AS OPIS," & _








        DefItem("PLANDLUGOOKRESOWY", "Plan d³ugookresowy", "Liczba", fiPlanDlugookresowyKli)
        DefItem("PLANKROTKOOKRESOWY", "Plan krótkookresowy", "Liczba", fiPlanKrotkookresowyKli)

        If _Opcja = "AnalizaSprzedazy" Then
            'Par_OSTOkresAnalizy="RRRR-MM" & "RRRR-MM"
            'I okres analizy
            If Len(Trim(Mid(Par_OstOkresAnalizy, 1, 7))) = 0 Then
                Okres1 = Mid(SysData, 1, 7)
            Else
                Okres1 = Mid(Par_OstOkresAnalizy, 1, 7)
            End If

            'II okres analizy
            If Len(Trim(Mid(Par_OstOkresAnalizy, 8, 7))) = 0 Then
                Okres2 = MiesDoAnalizy(Okres1, -12)
            Else
                Okres2 = Mid(Par_OstOkresAnalizy, 8, 7)
            End If


            DefItem("ILOSCAA", "Zakupy Okres I", "Liczba", fiIloscAA, System.Drawing.Color.PaleTurquoise)
            DefItem("ILOSCBB", "Zakupy Okres II", "Liczba", fiIloscBB, System.Drawing.Color.PaleTurquoise)

            DefItem("ILOSCBM", "Zakupy " + InnyMiesiac(Okres1, 0), "Liczba", fiIlZak00, System.Drawing.Color.LightCyan)
            DefItem("ILZAK01", "Zakupy " + InnyMiesiac(Okres1, -1), "Liczba", fiIlZak01, System.Drawing.Color.LightCyan)
            DefItem("ILZAK02", "Zakupy " + InnyMiesiac(Okres1, -2), "Liczba", fiIlZak02, System.Drawing.Color.LightCyan)
            DefItem("ILZAK03", "Zakupy " + InnyMiesiac(Okres1, -3), "Liczba", fiIlZak03, System.Drawing.Color.LightCyan)
            DefItem("ILZAK04", "Zakupy " + InnyMiesiac(Okres1, -4), "Liczba", fiIlZak04, System.Drawing.Color.LightCyan)
            DefItem("ILZAK05", "Zakupy " + InnyMiesiac(Okres1, -5), "Liczba", fiIlZak05, System.Drawing.Color.LightCyan)
            DefItem("ILZAK06", "Zakupy " + InnyMiesiac(Okres1, -6), "Liczba", fiIlZak06, System.Drawing.Color.LightCyan)
            DefItem("ILZAK07", "Zakupy " + InnyMiesiac(Okres1, -7), "Liczba", fiIlZak07, System.Drawing.Color.LightCyan)
            DefItem("ILZAK08", "Zakupy " + InnyMiesiac(Okres1, -8), "Liczba", fiIlZak08, System.Drawing.Color.LightCyan)
            DefItem("ILZAK09", "Zakupy " + InnyMiesiac(Okres1, -9), "Liczba", fiIlZak09, System.Drawing.Color.LightCyan)
            DefItem("ILZAK10", "Zakupy " + InnyMiesiac(Okres1, -10), "Liczba", fiIlZak10, System.Drawing.Color.LightCyan)
            DefItem("ILZAK11", "Zakupy " + InnyMiesiac(Okres1, -11), "Liczba", fiIlZak11, System.Drawing.Color.LightCyan)

            DefItem("ILZAK12", "Zakupy " + InnyMiesiac(Okres2, 0), "Liczba", fiIlZak12, System.Drawing.Color.Azure)
            DefItem("ILZAK13", "Zakupy " + InnyMiesiac(Okres2, -1), "Liczba", fiIlZak13, System.Drawing.Color.Azure)
            DefItem("ILZAK14", "Zakupy " + InnyMiesiac(Okres2, -2), "Liczba", fiIlZak14, System.Drawing.Color.Azure)
            DefItem("ILZAK15", "Zakupy " + InnyMiesiac(Okres2, -3), "Liczba", fiIlZak15, System.Drawing.Color.Azure)
            DefItem("ILZAK16", "Zakupy " + InnyMiesiac(Okres2, -4), "Liczba", fiIlZak16, System.Drawing.Color.Azure)
            DefItem("ILZAK17", "Zakupy " + InnyMiesiac(Okres2, -5), "Liczba", fiIlZak17, System.Drawing.Color.Azure)
            DefItem("ILZAK18", "Zakupy " + InnyMiesiac(Okres2, -6), "Liczba", fiIlZak18, System.Drawing.Color.Azure)
            DefItem("ILZAK19", "Zakupy " + InnyMiesiac(Okres2, -7), "Liczba", fiIlZak19, System.Drawing.Color.Azure)
            DefItem("ILZAK20", "Zakupy " + InnyMiesiac(Okres2, -8), "Liczba", fiIlZak20, System.Drawing.Color.Azure)
            DefItem("ILZAK21", "Zakupy " + InnyMiesiac(Okres2, -9), "Liczba", fiIlZak21, System.Drawing.Color.Azure)
            DefItem("ILZAK22", "Zakupy " + InnyMiesiac(Okres2, -10), "Liczba", fiIlZak22, System.Drawing.Color.Azure)
            DefItem("ILZAK23", "Zakupy " + InnyMiesiac(Okres2, -11), "Liczba", fiIlZak23, System.Drawing.Color.Azure)
        End If

        If _Opcja = "KatalogKlientowIKontaktow" Then
            DefItem("DATAKONTAKTU", "Data kontaktu", "Data", fiDataKon, System.Drawing.SystemColors.ControlLight)
            DefItem("PRACOWNIK", "Pracownik", "Text", fiPracownikKon, System.Drawing.SystemColors.ControlLight)
            DefItem("TEMAT", "Rodzaj kontaktu", "Text", fiTematKon, System.Drawing.SystemColors.ControlLight)
            DefItem("OPIS", "Opis kontaktu", "Text", fiUwagiKon, System.Drawing.SystemColors.ControlLight)
        End If


        ''zdefiniuj poszczegolne pola w bazie KLIENCI
        DefItem("NRKARTY", "Nr karty", "Text", fiIdentKli)
        DefItem("NIP", "NIP klienta", "Text", fiNIPKli)
        DefItem("NRTOFITXT", "Nr TOFI", "Text", fiNrTOFITxtKli)

        DefItem("KARTAPAYBACK", "Karta PayBack", "Tak/Nie", fiKartaPaybackKli)
        DefItem("NRYKARTPAYBACK", "Numery kart PayBacki", "Text", fiNryKartPaybackKli)

        DefItem("NAZWAKLIENTA", "Nazwa klienta", "Text", fiNazwa1Kli)
        DefItem("KODPOCZTOWY", "Kod pocztowy", "Text", fiKodPoczKli)
        DefItem("MIEJSCOWOSC", "Miejscowoœæ", "Text", fiMiejscKli)
        DefItem("ULICA", "Ulica", "Text", fiUlicaKli)
        DefItem("NUMNRDOMU", "Nr domu", "Liczba", fiNumNrDomuKli)
        DefItem("IDDOMU", "Ident domu", "Text", fiIDDomuKli)
        DefItem("TELEFON", "Telefon", "Text", fiTelefonKli)
        DefItem("FAX", "Fax", "Text", fiFaxKli)
        DefItem("EMAIL", "eMail", "Text", fieMailKli)
        DefItem("KTOWPISAL", "Kto wpisa³", "Text", fiWpisalKli)
        DefItem("REJONKLIENTA", "Rejon klienta", "Text", fiRejonKli)

        DefItem("BRANZA", "Bran¿a", "Text", fiBranzaKli)
        DefItem("PODBRANZA", "Podbran¿a", "Text", fiPodBranzaKli)
        DefItem("LICZBAZATRUDNIONYCH", "Liczba zatrudnionych", "Liczba", fiLiczbaUrzadzenKli)
        DefItem("LICZBAPRACZDALNIE", "Liczba prac. zdalnie", "Liczba", fiLiczbaUrzadzenKli)
        ''-------------------------------------------------
        DefItem("WYKAZURZADZEN", "Wykaz urz¹dzeñ", "Text", fiWykazUrzadzenKli)
        DefItem("SPOSOBWYBORUDOSTAWCY", "Sposób wyboru dostawcy", "Text", fiSposobWyboruDostawcyKli)
        DefItem("ZAINSTALOWANOMONITOR", "Zainstalowano minitor", "Tak/Nie", fiZainstalowanoMonitorKli)
        DefItem("LICZBAURZADZEN", "Liczba urz¹dzeñ", "Liczba", fiLiczbaUrzadzenKli)
        DefItem("LICZBAWYDRUKOW", "Liczba wydruków", "Liczba", fiLiczbaWydrukowKli)
        DefItem("POTENCJALDRUKU", "Potencja³ druku", "Text", fiPotencjalDrukuKli)
        DefItem("AKTZAKUPMATEKSP", "Akt. kupuje materia³y eksp.", "Tak/Nie", fiAktZakupMatEkspKli)
        DefItem("AKTROZLICZASTRONYDRUKU", "Akt. rozlicza strony wydruku", "Tak/Nie", fiAktRozliczaStronyDrukuKli)
        DefItem("AKTKORZYSTAZNAJMU", "Akt. korzysta z najmu", "Tak/Nie", fiAktKorzystaZNajmuKli)
        DefItem("ZAINTROZLICZANIEMSTRONDRUKU", "Zaint. rozliczaniem stron", "Tak/Nie", fiZaintRozliczaniemStronDrukuKli)
        DefItem("ZAINTKORZYSTANIEMZNAJMU", "Zaint. korzystaniem z najmu", "Tak/Nie", fiZaintKorzystaniemZNajmuKli)
        DefItem("DATAWERYFSEGMENTACJI", "Data weryfik. segmentacji", "Data", fiDataWeryfSegmentacjiKli)

        ''-------------------------------------------------
        DefItem("RODZSPRZETUL", "Rodzaj sprzêtu L", "Tak/Nie", fiRodzSprzetuLKli)
        DefItem("RODZSPRZETUA", "Rodzaj sprzêtu A", "Tak/Nie", fiRodzSprzetuAKli)
        DefItem("RODZSPRZETUT", "Rodzaj sprzêtu A3", "Tak/Nie", fiRodzSprzetuTKli)
        DefItem("UWAGI", "Uwagi", "Text", fiUwagikli)
        DefItem("OFERTATZLOZENIA", "Termin z³o¿enia oferty", "Text", fiOfertaTZlozeniaKli)
        DefItem("OFERTAPLIK", "Plik z Ofert¹", "Text", fiOfertaPlikKli)
        DefItem("PKONTAKT", "I kontakt", "Data", fiPKontaktKli)
        DefItem("SKUPWARTOSC", "Skup wartoœæ", "Text", fiSkupWartoscKli)
        DefItem("OPIEKUN", "Opiekun", "Text", fiSprzedOpiekunkli)
        DefItem("OPIEKUNOSTKONTAKTROK", "Opiekun ostatni kontakt Rok", "Text", fiSprzedOKontaktRKli)
        DefItem("OPIEKUNOSTKONTAKTTYDZ", "Opiekun ostatni kontakt Tydzieñ", "Liczba", fiSprzedOKontaktTKli)
        DefItem("OPIEKUNOSTKONTAKTDZIEN", "Opiekun ostatni kontakt Dzieñ", "Text", fiSprzedOKontaktDKli)
        DefItem("OPIEKUNKOLKONTAKTROK", "Opiekun kolejny kontakt Rok", "Text", fiSprzedNKontaktRKli)
        DefItem("OPIEKUNKOLKONTAKTTYDZ", "Opiekun kolejny kontakt Tydzieñ", "Liczba", fiSprzedNKontaktTKli)
        DefItem("OPIEKUNKOLKONTAKTDZIEN", "Opiekun kolejny kontakt Dzieñ", "Text", fiSprzedNKontaktDKli)
        DefItem("OPIEKUNUWAGI", "Opiekun Uwagi", "Text", fiSprzedUwagiKli)
        DefItem("AKTYWNY", "Klient Aktywny", "Tak/Nie", fiAktywnyKli)

    End Sub


    Private Sub DefItem(ByVal NazwaPola As String, ByVal OpisPola As String, ByVal TypPola As String, ByVal FiltrPola As String, ByVal bColor As System.Drawing.Color)
        Dim Item00 As New ListViewItem(NazwaPola)
        Item00.SubItems.Add(OpisPola)
        Item00.SubItems.Add(TypPola)
        Item00.SubItems.Add(FiltrPola)
        If Len(FiltrPola) > 0 Then
            Item00.ForeColor = System.Drawing.Color.Purple
        Else
            Item00.ForeColor = System.Drawing.Color.Purple
        End If
        Item00.BackColor = bColor
        ListaPol.Items.Add(Item00)
    End Sub

    Private Sub DefItem(ByVal NazwaPola As String, ByVal OpisPola As String, ByVal TypPola As String, ByVal FiltrPola As String)
        Dim Item00 As New ListViewItem(NazwaPola)
        Item00.SubItems.Add(OpisPola)
        Item00.SubItems.Add(TypPola)
        Item00.SubItems.Add(FiltrPola)
        If Len(FiltrPola) > 0 Then
            Item00.ForeColor = System.Drawing.Color.Purple
        Else
            Item00.ForeColor = System.Drawing.Color.Purple
        End If
        If ListaPol.Items.Count Mod 2 = 1 Then
            'nieparzysta
            Item00.BackColor = System.Drawing.SystemColors.ControlLight
        Else
            'parzysta
            Item00.BackColor = System.Drawing.SystemColors.Control
        End If
        ListaPol.Items.Add(Item00)
    End Sub

    Private Sub AktListePol()
        Dim IleKli As Integer = 0
        Dim IleKon As Integer = 0
        Dim IleAnalSp As Integer = 0







        ListaPol.Items(0).SubItems(3).Text = fiPlanDlugookresowyKli
        ListaPol.Items(1).SubItems(3).Text = fiPlanKrotkookresowyKli
        IleKli = 2

        If _Opcja = "AnalizaSprzedazy" Then
            ListaPol.Items(IleKli + 0).SubItems(3).Text = fiIloscAA
            ListaPol.Items(IleKli + 1).SubItems(3).Text = fiIloscBB

            ListaPol.Items(IleKli + 2).SubItems(3).Text = fiIlZak00
            ListaPol.Items(IleKli + 3).SubItems(3).Text = fiIlZak01
            ListaPol.Items(IleKli + 4).SubItems(3).Text = fiIlZak02
            ListaPol.Items(IleKli + 5).SubItems(3).Text = fiIlZak03
            ListaPol.Items(IleKli + 6).SubItems(3).Text = fiIlZak04
            ListaPol.Items(IleKli + 7).SubItems(3).Text = fiIlZak05
            ListaPol.Items(IleKli + 8).SubItems(3).Text = fiIlZak06
            ListaPol.Items(IleKli + 9).SubItems(3).Text = fiIlZak07
            ListaPol.Items(IleKli + 10).SubItems(3).Text = fiIlZak08
            ListaPol.Items(IleKli + 11).SubItems(3).Text = fiIlZak09
            ListaPol.Items(IleKli + 12).SubItems(3).Text = fiIlZak10
            ListaPol.Items(IleKli + 13).SubItems(3).Text = fiIlZak11

            ListaPol.Items(IleKli + 14).SubItems(3).Text = fiIlZak12
            ListaPol.Items(IleKli + 15).SubItems(3).Text = fiIlZak13
            ListaPol.Items(IleKli + 16).SubItems(3).Text = fiIlZak14
            ListaPol.Items(IleKli + 17).SubItems(3).Text = fiIlZak15
            ListaPol.Items(IleKli + 18).SubItems(3).Text = fiIlZak16
            ListaPol.Items(IleKli + 19).SubItems(3).Text = fiIlZak17
            ListaPol.Items(IleKli + 20).SubItems(3).Text = fiIlZak18
            ListaPol.Items(IleKli + 21).SubItems(3).Text = fiIlZak19
            ListaPol.Items(IleKli + 22).SubItems(3).Text = fiIlZak20
            ListaPol.Items(IleKli + 23).SubItems(3).Text = fiIlZak21
            ListaPol.Items(IleKli + 24).SubItems(3).Text = fiIlZak22
            ListaPol.Items(IleKli + 25).SubItems(3).Text = fiIlZak23

            IleAnalSp = 26
        End If


        If _Opcja = "KatalogKlientowIKontaktow" Then
            'Public fiIdentKon As String            'IDENTKLIENTA     Text, 6
            'Public fiDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
            'Public fiPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
            'Public fiTematKon As String            'TEMAT            Text, 50
            '  Public fiUwagiKon As String            'UWAGI            Memo
            '  Public fiWersjaKon As Integer          'WERSJA           Integer

            ListaPol.Items(IleKli + IleAnalSp + 0).SubItems(3).Text = fiDataKon
            ListaPol.Items(IleKli + IleAnalSp + 1).SubItems(3).Text = fiPracownikKon
            ListaPol.Items(IleKli + IleAnalSp + 2).SubItems(3).Text = fiTematKon
            ListaPol.Items(IleKli + IleAnalSp + 3).SubItems(3).Text = fiUwagiKon
            IleKon = 4
        End If

        ''-------------------------------------------------

        ListaPol.Items(IleKli + IleAnalSp + IleKon + 0).SubItems(3).Text = fiIdentKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 1).SubItems(3).Text = fiNIPKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 2).SubItems(3).Text = fiNrTOFITxtKli

        ListaPol.Items(IleKli + IleAnalSp + IleKon + 3).SubItems(3).Text = fiKartaPaybackKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 4).SubItems(3).Text = fiNryKartPaybackKli

        ListaPol.Items(IleKli + IleAnalSp + IleKon + 5).SubItems(3).Text = fiNazwa1Kli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 6).SubItems(3).Text = fiKodPoczKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 7).SubItems(3).Text = fiMiejscKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 8).SubItems(3).Text = fiUlicaKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 9).SubItems(3).Text = fiNumNrDomuKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 10).SubItems(3).Text = fiIDDomuKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 11).SubItems(3).Text = fiTelefonKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 12).SubItems(3).Text = fiFaxKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 13).SubItems(3).Text = fieMailKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 14).SubItems(3).Text = fiWpisalKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 15).SubItems(3).Text = fiRejonKli

        ListaPol.Items(IleKli + IleAnalSp + IleKon + 16).SubItems(3).Text = fiBranzaKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 17).SubItems(3).Text = fiPodBranzaKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 18).SubItems(3).Text = fiLiczbaZaktrudnionychKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 19).SubItems(3).Text = fiLiczbaPracZdalnieKli

        ListaPol.Items(IleKli + IleAnalSp + IleKon + 20).SubItems(3).Text = fiWykazUrzadzenKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 21).SubItems(3).Text = fiSposobWyboruDostawcyKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 22).SubItems(3).Text = fiZainstalowanoMonitorKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 23).SubItems(3).Text = fiLiczbaUrzadzenKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 24).SubItems(3).Text = fiLiczbaWydrukowKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 25).SubItems(3).Text = fiPotencjalDrukuKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 26).SubItems(3).Text = fiAktZakupMatEkspKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 27).SubItems(3).Text = fiAktRozliczaStronyDrukuKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 28).SubItems(3).Text = fiAktKorzystaZNajmuKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 29).SubItems(3).Text = fiZaintRozliczaniemStronDrukuKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 30).SubItems(3).Text = fiZaintKorzystaniemZNajmuKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 31).SubItems(3).Text = fiDataWeryfSegmentacjiKli

        ListaPol.Items(IleKli + IleAnalSp + IleKon + 32).SubItems(3).Text = fiRodzSprzetuLKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 33).SubItems(3).Text = fiRodzSprzetuAKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 34).SubItems(3).Text = fiRodzSprzetuTKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 35).SubItems(3).Text = fiUwagikli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 36).SubItems(3).Text = fiOfertaTZlozeniaKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 37).SubItems(3).Text = fiOfertaPlikKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 38).SubItems(3).Text = fiPKontaktKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 39).SubItems(3).Text = fiSkupWartoscKli

        ListaPol.Items(IleKli + IleAnalSp + IleKon + 40).SubItems(3).Text = fiSprzedOpiekunkli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 41).SubItems(3).Text = fiSprzedOKontaktRKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 42).SubItems(3).Text = fiSprzedOKontaktTKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 43).SubItems(3).Text = fiSprzedOKontaktDKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 44).SubItems(3).Text = fiSprzedNKontaktRKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 45).SubItems(3).Text = fiSprzedNKontaktTKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 46).SubItems(3).Text = fiSprzedNKontaktDKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 47).SubItems(3).Text = fiSprzedUwagiKli
        ListaPol.Items(IleKli + IleAnalSp + IleKon + 48).SubItems(3).Text = fiAktywnyKli

    End Sub

    Private Sub ListaPol_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListaPol.SelectedIndexChanged
        Dim iCollection As ListView.SelectedListViewItemCollection = Me.ListaPol.SelectedItems
        Dim iRow As ListViewItem
        Dim iIndex As Integer

        Dim iNazwa As String = ""
        Dim iOpis As String = ""
        Dim iTyp As String = ""
        Dim iTest As String = ""
        Dim iFiltr As String = ""

        Dim pos As Integer = 0
        Dim TxtOd As String = ""
        Dim TxtDo As String = ""

        pnlText.Visible = False
        pnlData.Visible = False
        pnlLiczba.Visible = False
        pnlTakNie.Visible = False

        'ZAWIERA KOLEKCJE MARKOWANYCH REKORDOW
        For Each iRow In iCollection
            iIndex = iRow.Index
            iNazwa = iRow.SubItems(0).Text
            iOpis = iRow.SubItems(1).Text
            iTyp = iRow.SubItems(2).Text
            iFiltr = iRow.SubItems(3).Text
        Next

        lblNazwa.Text = iNazwa
        lblTyp.Text = iTyp
        CurItemIndex = iIndex

        Select Case iTyp
            Case "Text"
                pnlText.Visible = True
                If Mid(iFiltr, 1, 3) = "Tak" Then
                    chbTestTxt.CheckState = CheckState.Checked

                    rabtextPuste.Visible = True
                    rabTextNiepuste.Visible = True
                    rabtextZawiera.Visible = True
                    rabTextNieZawiera.Visible = True
                    rabTexOdDo.Visible = True

                    Select Case Mid(iFiltr, 7, 8)
                        Case "Puste   "
                            rabtextPuste.Checked = True
                            rabtextPuste.Visible = True
                            lblTextOd.Visible = False
                            txtTextOd.Visible = False
                            lblTextOd.Text = ""
                            lblTextDo.Visible = False
                            txtTextDo.Visible = False
                            lblTextDo.Text = ""
                        Case "Niepuste"
                            rabTextNiepuste.Checked = True
                            lblTextOd.Visible = False
                            txtTextOd.Visible = False
                            lblTextOd.Text = ""
                            lblTextDo.Visible = False
                            txtTextDo.Visible = False
                            lblTextDo.Text = ""
                        Case "Zawiera "
                            rabtextZawiera.Checked = True
                            lblTextOd.Visible = True
                            txtTextOd.Visible = True
                            lblTextOd.Text = "Tekst : . . . . . . ."
                            txtTextOd.Text = Mid(iFiltr, 16)
                            lblTextDo.Visible = False
                            txtTextDo.Visible = False
                            lblTextDo.Text = ""
                        Case "NieZawi."
                            rabTextNieZawiera.Checked = True
                            lblTextOd.Visible = True
                            txtTextOd.Visible = True
                            lblTextOd.Text = "Tekst : . . . . . . ."
                            txtTextOd.Text = Mid(iFiltr, 16)
                            lblTextDo.Visible = False
                            txtTextDo.Visible = False
                            lblTextDo.Text = ""
                        Case "Od...Do "
                            rabTexOdDo.Checked = True
                            pos = InStr(Mid(iFiltr, 16), " ... ")
                            If pos > 0 Then
                                'TAK_|_Od...Do  999...999
                                TxtOd = Mid(iFiltr, 16, pos - 1)
                                TxtDo = Mid(iFiltr, 16 + pos + 5 - 1)

                                lblTextOd.Visible = True
                                txtTextOd.Visible = True
                                lblTextOd.Text = "Tekst od : . . . . . . ."
                                txtTextOd.Text = TxtOd

                                lblTextDo.Visible = True
                                txtTextDo.Visible = True
                                lblTextDo.Text = "Tekst do : . . . . . . ."
                                txtTextDo.Text = TxtDo
                            Else
                                lblTextOd.Visible = True
                                txtTextOd.Visible = True
                                lblTextOd.Text = "Tekst : . . . . . . ."
                                txtTextOd.Text = Mid(iFiltr, 16)
                                lblTextDo.Visible = False
                                txtTextDo.Visible = False
                                lblTextDo.Text = ""
                            End If

                        Case Else
                            lblTextOd.Visible = False
                            txtTextOd.Visible = False
                            lblTextOd.Text = ""
                            lblTextDo.Visible = False
                            txtTextDo.Visible = False
                            lblTextDo.Text = ""
                    End Select
                Else
                    chbTestTxt.CheckState = CheckState.Unchecked

                    rabtextPuste.Visible = False
                    rabTextNiepuste.Visible = False
                    rabtextZawiera.Visible = False
                    rabTextNieZawiera.Visible = False
                    rabTexOdDo.Visible = False

                    lblTextOd.Visible = False
                    txtTextOd.Visible = False
                    lblTextOd.Text = ""
                    lblTextDo.Visible = False
                    txtTextDo.Visible = False
                    lblTextDo.Text = ""
                End If

            Case "Liczba"
                If Mid(iFiltr, 1, 3) = "Tak" Then
                    pnlLiczba.Visible = True
                    chbTestNum.Visible = True
                    chbTestNum.CheckState = CheckState.Checked

                    rabLiczbaRowne.Visible = True
                    rabLiczbaRozne.Visible = True
                    rabLiczbaMniesze.Visible = True
                    rabLiczbaMniejRowne.Visible = True
                    rabLiczbaWieksze.Visible = True
                    rabLiczbaWieRowne.Visible = True
                    rabLiczbaOdDo.Visible = True

                    Select Case Mid(iFiltr, 7, 3)
                        Case "=  "
                            rabLiczbaRowne.Checked = True
                            TxtLiczbaOd.Text = Mid(iFiltr, 16)
                            lblLiczbaOd.Visible = True
                            lblLiczbaDo.Visible = False
                            TxtLiczbaOd.Visible = True
                            txtLiczbaDo.Visible = False

                        Case "<> "
                            rabLiczbaRozne.Checked = True
                            TxtLiczbaOd.Text = Mid(iFiltr, 16)
                            lblLiczbaOd.Visible = True
                            lblLiczbaDo.Visible = False
                            TxtLiczbaOd.Visible = True
                            txtLiczbaDo.Visible = False

                        Case "<  "
                            rabLiczbaMniesze.Checked = True
                            TxtLiczbaOd.Text = Mid(iFiltr, 16)
                            lblLiczbaOd.Visible = True
                            lblLiczbaDo.Visible = False
                            TxtLiczbaOd.Visible = True
                            txtLiczbaDo.Visible = False

                        Case "<= "
                            rabLiczbaMniejRowne.Checked = True
                            TxtLiczbaOd.Text = Mid(iFiltr, 16)
                            lblLiczbaOd.Visible = True
                            lblLiczbaDo.Visible = False
                            TxtLiczbaOd.Visible = True
                            txtLiczbaDo.Visible = False

                        Case ">  "
                            rabLiczbaWieksze.Checked = True
                            TxtLiczbaOd.Text = Mid(iFiltr, 16)
                            lblLiczbaOd.Visible = True
                            lblLiczbaDo.Visible = False
                            TxtLiczbaOd.Visible = True
                            txtLiczbaDo.Visible = False

                        Case ">= "
                            rabLiczbaWieRowne.Checked = True
                            TxtLiczbaOd.Text = Mid(iFiltr, 16)
                            lblLiczbaOd.Visible = True
                            lblLiczbaDo.Visible = False
                            TxtLiczbaOd.Visible = True
                            txtLiczbaDo.Visible = False

                        Case "Od."
                            rabLiczbaOdDo.Checked = True
                            pos = InStr(Mid(iFiltr, 16), " ... ")
                            If pos > 0 Then
                                TxtOd = Mid(iFiltr, 16, pos - 1)
                                TxtDo = Mid(iFiltr, 16 + pos + 5 - 1)
                            Else
                                TxtOd = ""
                                TxtDo = ""
                            End If
                            TxtLiczbaOd.Text = TxtOd
                            txtLiczbaDo.Text = TxtDo
                            lblLiczbaOd.Visible = True
                            lblLiczbaDo.Visible = True
                            TxtLiczbaOd.Visible = True
                            txtLiczbaDo.Visible = True

                        Case Else
                            rabLiczbaRowne.Checked = True
                            TxtLiczbaOd.Text = Mid(iFiltr, 16)
                            lblLiczbaOd.Visible = True
                            lblLiczbaDo.Visible = False
                            TxtLiczbaOd.Visible = True
                            txtLiczbaDo.Visible = False
                    End Select
                Else
                    pnlLiczba.Visible = True
                    chbTestNum.Visible = True
                    chbTestNum.CheckState = CheckState.Unchecked

                    rabLiczbaRowne.Visible = False
                    rabLiczbaRozne.Visible = False
                    rabLiczbaMniesze.Visible = False
                    rabLiczbaMniejRowne.Visible = False
                    rabLiczbaWieksze.Visible = False
                    rabLiczbaWieRowne.Visible = False
                    rabLiczbaOdDo.Visible = False

                    lblLiczbaOd.Visible = False
                    lblLiczbaDo.Visible = False
                    TxtLiczbaOd.Visible = False
                    txtLiczbaDo.Visible = False
                End If











            Case "Tak/Nie"
                pnlTakNie.Visible = True
                If Mid(iFiltr, 7, 3) = "Tak" Then
                    rbtTak.Checked = True
                Else
                    rbtNie.Checked = True
                End If

                If Mid(iFiltr, 1, 3) = "Tak" Then
                    chbTestLog.CheckState = CheckState.Checked
                    rbtTak.Visible = True
                    rbtNie.Visible = True
                    lblTak.Visible = True
                    lblNie.Visible = True
                Else
                    chbTestLog.CheckState = CheckState.Unchecked
                    rbtTak.Visible = False
                    rbtNie.Visible = False
                    lblTak.Visible = False
                    lblNie.Visible = False
                End If





            Case "Data"
                pnlData.Visible = True
                txtDataOd.Text = Mid(iFiltr, 16, 10)
                txtDataDo.Text = Mid(iFiltr, 16 + 16 - 1, 10)
                If Mid(iFiltr, 1, 3) = "Tak" Then
                    chbTestDate.CheckState = CheckState.Checked
                    lblDataOd.Visible = True
                    lblDataDo.Visible = True
                    txtDataOd.Visible = True
                    txtDataDo.Visible = True
                    cmdDataOd.Visible = True
                    cmdDataDo.Visible = True
                Else
                    chbTestDate.CheckState = CheckState.Unchecked
                    lblDataOd.Visible = False
                    lblDataDo.Visible = False
                    txtDataOd.Visible = False
                    txtDataDo.Visible = False
                    cmdDataOd.Visible = False
                    cmdDataDo.Visible = False
                End If
        End Select
    End Sub

    Private Sub AktualFiltr(ByVal Ind As Integer, ByVal NowyFiltr As String)
        Dim IleKli As Integer = 0
        Dim IleKon As Integer = 0
        Dim IleAnalSp As Integer = 0





        'kATALOG Klientow
        Select Case Ind
            Case IleAnalSp + IleKon + 0
                fiPlanDlugookresowyKli = NowyFiltr
            Case IleAnalSp + IleKon + 1
                fiPlanKrotkookresowyKli = NowyFiltr
        End Select
        IleKli = 2

        If _Opcja = "AnalizaSprzedazy" Then
            Select Case Ind
                Case IleKli + 0
                    fiIloscAA = NowyFiltr
                Case IleKli + 1
                    fiIloscBB = NowyFiltr

                Case IleKli + 2
                    fiIlZak00 = NowyFiltr
                Case IleKli + 3
                    fiIlZak01 = NowyFiltr
                Case IleKli + 4
                    fiIlZak02 = NowyFiltr
                Case IleKli + 5
                    fiIlZak03 = NowyFiltr
                Case IleKli + 6
                    fiIlZak04 = NowyFiltr
                Case IleKli + 7
                    fiIlZak05 = NowyFiltr
                Case IleKli + 8
                    fiIlZak06 = NowyFiltr
                Case IleKli + 9
                    fiIlZak07 = NowyFiltr
                Case IleKli + 10
                    fiIlZak08 = NowyFiltr
                Case IleKli + 11
                    fiIlZak09 = NowyFiltr
                Case IleKli + 12
                    fiIlZak10 = NowyFiltr
                Case IleKli + 13
                    fiIlZak11 = NowyFiltr


                Case IleKli + 14
                    fiIlZak12 = NowyFiltr
                Case IleKli + 15
                    fiIlZak13 = NowyFiltr
                Case IleKli + 16
                    fiIlZak14 = NowyFiltr
                Case IleKli + 17
                    fiIlZak15 = NowyFiltr
                Case IleKli + 18
                    fiIlZak16 = NowyFiltr
                Case IleKli + 19
                    fiIlZak17 = NowyFiltr
                Case IleKli + 20
                    fiIlZak18 = NowyFiltr
                Case IleKli + 21
                    fiIlZak19 = NowyFiltr
                Case IleKli + 22
                    fiIlZak20 = NowyFiltr
                Case IleKli + 23
                    fiIlZak21 = NowyFiltr
                Case IleKli + 24
                    fiIlZak22 = NowyFiltr
                Case IleKli + 25
                    fiIlZak23 = NowyFiltr
            End Select
            IleAnalSp = 26
        End If




        If _Opcja = "KatalogKlientowIKontaktow" Then
            'Public fiIdentKon As String            'IDENTKLIENTA     Text, 6
            'Public fiDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
            'Public fiPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
            'Public fiTematKon As String            'TEMAT            Text, 50
            'Public fiUwagiKon As String            'UWAGI            Memo
            'Public fiWersjaKon As Integer          'WERSJA           Integer


            Select Case Ind
                Case IleKli + IleAnalSp + 0
                    fiDataKon = NowyFiltr
                Case IleKli + IleAnalSp + 1
                    fiPracownikKon = NowyFiltr
                Case IleKli + IleAnalSp + 2
                    fiTematKon = NowyFiltr
                Case IleKli + IleAnalSp + 3
                    fiUwagiKon = NowyFiltr
            End Select
            IleKon = 4
        End If



        'kATALOG Klientow
        Select Case Ind
            Case IleKli + IleAnalSp + IleKon + 0
                fiIdentKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 1
                fiNIPKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 2
                fiNrTOFITxtKli = NowyFiltr

            Case IleKli + IleAnalSp + IleKon + 3
                fiKartaPaybackKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 4
                fiNryKartPaybackKli = NowyFiltr

            Case IleKli + IleAnalSp + IleKon + 5
                fiNazwa1Kli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 6
                fiKodPoczKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 7
                fiMiejscKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 8
                fiUlicaKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 9
                fiNumNrDomuKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 10
                fiIDDomuKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 11
                fiTelefonKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 12
                fiFaxKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 13
                fieMailKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 14
                fiWpisalKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 15
                fiRejonKli = NowyFiltr

            Case IleKli + IleAnalSp + IleKon + 16
                fiBranzaKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 17
                fiPodBranzaKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 18
                fiLiczbaZaktrudnionychKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 19
                fiLiczbaPracZdalnieKli = NowyFiltr




            Case IleKli + IleAnalSp + IleKon + 20
                fiWykazUrzadzenKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 21
                fiSposobWyboruDostawcyKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 22
                fiZainstalowanoMonitorKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 23
                fiLiczbaUrzadzenKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 24
                fiLiczbaWydrukowKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 25
                fiPotencjalDrukuKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 26
                fiAktZakupMatEkspKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 27
                fiAktRozliczaStronyDrukuKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 28
                fiAktKorzystaZNajmuKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 29
                fiZaintRozliczaniemStronDrukuKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 30
                fiZaintKorzystaniemZNajmuKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 31
                fiDataWeryfSegmentacjiKli = NowyFiltr

            Case IleKli + IleAnalSp + IleKon + 32
                fiRodzSprzetuLKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 33
                fiRodzSprzetuAKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 34
                fiRodzSprzetuTKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 35
                fiUwagikli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 36
                fiOfertaTZlozeniaKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 37
                fiOfertaPlikKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 38
                fiPKontaktKli = NowyFiltr

            Case IleKli + IleAnalSp + IleKon + 39
                fiSkupWartoscKli = NowyFiltr

            Case IleKli + IleAnalSp + IleKon + 40
                fiSprzedOpiekunkli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 41
                fiSprzedOKontaktRKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 42
                fiSprzedOKontaktTKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 43
                fiSprzedOKontaktDKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 44
                fiSprzedNKontaktRKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 45
                fiSprzedNKontaktTKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 46
                fiSprzedNKontaktDKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 47
                fiSprzedUwagiKli = NowyFiltr
            Case IleKli + IleAnalSp + IleKon + 48
                fiAktywnyKli = NowyFiltr

        End Select

        '---------------------------------------------
        ListaPol.Items(Ind).SubItems(3).Text = NowyFiltr
    End Sub

    '*********************************************************
    '* obsluga szablonu
    '*********************************************************

    Private Sub InitSzablon()
        fiPlanDlugookresowyKli = "Nie | "
        fiPlanKrotkookresowyKli = "Nie | "
        'analizy
        fiIloscAA = "Nie | "
        fiIloscBB = "Nie | "

        fiIlZak00 = "Nie | "
        fiIlZak01 = "Nie | "
        fiIlZak02 = "Nie | "
        fiIlZak03 = "Nie | "
        fiIlZak04 = "Nie | "
        fiIlZak05 = "Nie | "
        fiIlZak06 = "Nie | "
        fiIlZak07 = "Nie | "
        fiIlZak08 = "Nie | "
        fiIlZak09 = "Nie | "
        fiIlZak10 = "Nie | "
        fiIlZak11 = "Nie | "

        fiIlZak12 = "Nie | "
        fiIlZak13 = "Nie | "
        fiIlZak14 = "Nie | "
        fiIlZak15 = "Nie | "
        fiIlZak16 = "Nie | "
        fiIlZak17 = "Nie | "
        fiIlZak18 = "Nie | "
        fiIlZak19 = "Nie | "
        fiIlZak20 = "Nie | "
        fiIlZak21 = "Nie | "
        fiIlZak22 = "Nie | "
        fiIlZak23 = "Nie | "


        'Kontakty
        fiDataKon = "Nie | "
        fiPracownikKon = "Nie | "
        fiTematKon = "Nie | "
        fiUwagiKon = "Nie | "


        'Klienci
        fiIdentKli = "Nie | "
        fiNIPKli = "Nie | "
        fiNrTOFITxtKli = "Nie | "

        fiKartaPaybackKli = "Nie | "
        fiNryKartPaybackKli = "Nie | "

        fiNazwa1Kli = "Nie | "
        fiKodPoczKli = "Nie | "
        fiMiejscKli = "Nie | "
        fiUlicaKli = "Nie | "
        fiNumNrDomuKli = "Nie | "
        fiIDDomuKli = "Nie | "
        fiTelefonKli = "Nie | "
        fiFaxKli = "Nie | "
        fieMailKli = "Nie | "
        fiWpisalKli = "Nie | "
        fiRejonKli = "Nie | "

        fiBranzaKli = "Nie | "
        fiPodBranzaKli = "Nie | "
        fiLiczbaZaktrudnionychKli = "Nie | "
        fiLiczbaPracZdalnieKli = "Nie | "

        fiWykazUrzadzenKli = "Nie | "
        fiSposobWyboruDostawcyKli = "Nie | "
        fiZainstalowanoMonitorKli = "Nie | "
        fiLiczbaUrzadzenKli = "Nie | "
        fiLiczbaWydrukowKli = "Nie | "
        fiPotencjalDrukuKli = "Nie | "
        fiAktZakupMatEkspKli = "Nie | "
        fiAktRozliczaStronyDrukuKli = "Nie | "
        fiAktKorzystaZNajmuKli = "Nie | "
        fiZaintRozliczaniemStronDrukuKli = "Nie | "
        fiZaintKorzystaniemZNajmuKli = "Nie | "
        fiDataWeryfSegmentacjiKli = "Nie | "

        fiRodzSprzetuLKli = "Nie | "
        fiRodzSprzetuAKli = "Nie | "
        fiRodzSprzetuTKli = "Nie | "
        fiUwagikli = "Nie | "
        fiOfertaTZlozeniaKli = "Nie | "
        fiOfertaPlikKli = "Nie | "
        fiPKontaktKli = "Nie | "
        fiSkupWartoscKli = "Nie | "

        fiSprzedOpiekunkli = "Nie | "
        fiSprzedOKontaktRKli = "Nie | "
        fiSprzedOKontaktTKli = "Nie | "
        fiSprzedOKontaktDKli = "Nie | "
        fiSprzedNKontaktRKli = "Nie | "
        fiSprzedNKontaktTKli = "Nie | "
        fiSprzedNKontaktDKli = "Nie | "
        fiSprzedUwagiKli = "Nie | "
        fiAktywnyKli = "Nie | "
    End Sub

    Private Sub ZapiszSzablon()
        Dim FileNum As Integer

        FileNum = FreeFile()    'kolejny nr otwartego zbioru
        FileOpen(FileNum, oKatParametry & "\" & PlikZSzablonem, OpenMode.Output)
        PrintLine(FileNum, "SOFTART : Szablon filtrowania katalogu klientów")

        PrintLine(FileNum, "PLANDLUGOOKRESOWY=" & IIf(Len(fiPlanDlugookresowyKli) = 0, "NIE | ", fiPlanDlugookresowyKli))
        PrintLine(FileNum, "PLANKROTKOOKRESOWY=" & IIf(Len(fiPlanKrotkookresowyKli) = 0, "NIE | ", fiPlanKrotkookresowyKli))

        'Analizy sprzedazy
        PrintLine(FileNum, "ILOSCAA=" & IIf(Len(fiIloscAA) = 0, "NIE | ", fiIloscAA))
        PrintLine(FileNum, "ILOSCBB=" & IIf(Len(fiIloscBB) = 0, "NIE | ", fiIloscBB))

        PrintLine(FileNum, "ILOSCZAKBM=" & IIf(Len(fiIlZak00) = 0, "NIE | ", fiIlZak00))
        PrintLine(FileNum, "ILOSCZAK01=" & IIf(Len(fiIlZak01) = 0, "NIE | ", fiIlZak01))
        PrintLine(FileNum, "ILOSCZAK02=" & IIf(Len(fiIlZak02) = 0, "NIE | ", fiIlZak02))
        PrintLine(FileNum, "ILOSCZAK03=" & IIf(Len(fiIlZak03) = 0, "NIE | ", fiIlZak03))
        PrintLine(FileNum, "ILOSCZAK04=" & IIf(Len(fiIlZak04) = 0, "NIE | ", fiIlZak04))
        PrintLine(FileNum, "ILOSCZAK05=" & IIf(Len(fiIlZak05) = 0, "NIE | ", fiIlZak05))
        PrintLine(FileNum, "ILOSCZAK06=" & IIf(Len(fiIlZak06) = 0, "NIE | ", fiIlZak06))
        PrintLine(FileNum, "ILOSCZAK07=" & IIf(Len(fiIlZak07) = 0, "NIE | ", fiIlZak07))
        PrintLine(FileNum, "ILOSCZAK08=" & IIf(Len(fiIlZak08) = 0, "NIE | ", fiIlZak08))
        PrintLine(FileNum, "ILOSCZAK09=" & IIf(Len(fiIlZak09) = 0, "NIE | ", fiIlZak09))
        PrintLine(FileNum, "ILOSCZAK10=" & IIf(Len(fiIlZak10) = 0, "NIE | ", fiIlZak10))
        PrintLine(FileNum, "ILOSCZAK11=" & IIf(Len(fiIlZak11) = 0, "NIE | ", fiIlZak11))

        PrintLine(FileNum, "ILOSCZAK12=" & IIf(Len(fiIlZak12) = 0, "NIE | ", fiIlZak12))
        PrintLine(FileNum, "ILOSCZAK13=" & IIf(Len(fiIlZak13) = 0, "NIE | ", fiIlZak13))
        PrintLine(FileNum, "ILOSCZAK14=" & IIf(Len(fiIlZak14) = 0, "NIE | ", fiIlZak14))
        PrintLine(FileNum, "ILOSCZAK15=" & IIf(Len(fiIlZak15) = 0, "NIE | ", fiIlZak15))
        PrintLine(FileNum, "ILOSCZAK16=" & IIf(Len(fiIlZak16) = 0, "NIE | ", fiIlZak16))
        PrintLine(FileNum, "ILOSCZAK17=" & IIf(Len(fiIlZak17) = 0, "NIE | ", fiIlZak17))
        PrintLine(FileNum, "ILOSCZAK18=" & IIf(Len(fiIlZak18) = 0, "NIE | ", fiIlZak18))
        PrintLine(FileNum, "ILOSCZAK19=" & IIf(Len(fiIlZak19) = 0, "NIE | ", fiIlZak19))
        PrintLine(FileNum, "ILOSCZAK20=" & IIf(Len(fiIlZak20) = 0, "NIE | ", fiIlZak20))
        PrintLine(FileNum, "ILOSCZAK21=" & IIf(Len(fiIlZak21) = 0, "NIE | ", fiIlZak21))
        PrintLine(FileNum, "ILOSCZAK22=" & IIf(Len(fiIlZak22) = 0, "NIE | ", fiIlZak22))
        PrintLine(FileNum, "ILOSCZAK23=" & IIf(Len(fiIlZak23) = 0, "NIE | ", fiIlZak23))




        'Kontakty
        'Public fiIdentKon As String            'IDENTKLIENTA     Text, 6
        'Public fiDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
        'Public fiPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
        'Public fiTematKon As String            'TEMAT            Text, 50
        'Public fiUwagiKon As String            'UWAGI            Memo
        'Public fiWersjaKon As Integer          'WERSJA           Integer

        PrintLine(FileNum, "DATAKONTAKTU=" & IIf(Len(fiDataKon) = 0, "NIE | ", fiDataKon))
        PrintLine(FileNum, "PRACOWNIKKONTAKTU=" & IIf(Len(fiPracownikKon) = 0, "NIE | ", fiPracownikKon))
        PrintLine(FileNum, "RODZAJKONTAKTU=" & IIf(Len(fiTematKon) = 0, "NIE | ", fiTematKon))
        PrintLine(FileNum, "OPISKONTAKTU=" & IIf(Len(fiUwagiKon) = 0, "NIE | ", fiUwagiKon))



        ''-------------------------------------------------
        'Katalog Klientow
        PrintLine(FileNum, "IDENTKLIENTA=" & IIf(Len(fiIdentKli) = 0, "NIE | ", fiIdentKli))
        PrintLine(FileNum, "NIP=" & IIf(Len(fiNIPKli) = 0, "NIE | ", fiNIPKli))
        PrintLine(FileNum, "NRTOFI=" & IIf(Len(fiNrTOFITxtKli) = 0, "NIE | ", fiNrTOFITxtKli))

        PrintLine(FileNum, "KARTAPAYBACK=" & IIf(Len(fiKartaPaybackKli) = 0, "NIE | ", fiKartaPaybackKli))
        PrintLine(FileNum, "NRYKARTPAYBACK=" & IIf(Len(fiNryKartPaybackKli) = 0, "NIE | ", fiNryKartPaybackKli))

        PrintLine(FileNum, "NAZWA1=" & IIf(Len(fiNazwa1Kli) = 0, "NIE | ", fiNazwa1Kli))
        PrintLine(FileNum, "KODPOCZTOWY=" & IIf(Len(fiKodPoczKli) = 0, "NIE | ", fiKodPoczKli))
        PrintLine(FileNum, "MIEJSCOWOSC=" & IIf(Len(fiMiejscKli) = 0, "NIE | ", fiMiejscKli))
        PrintLine(FileNum, "ULICA=" & IIf(Len(fiUlicaKli) = 0, "NIE | ", fiUlicaKli))
        PrintLine(FileNum, "NRDOMU=" & IIf(Len(fiNumNrDomuKli) = 0, "NIE | ", fiNumNrDomuKli))
        PrintLine(FileNum, "IDDOMU=" & IIf(Len(fiIDDomuKli) = 0, "NIE | ", fiIDDomuKli))
        PrintLine(FileNum, "TELEFON=" & IIf(Len(fiTelefonKli) = 0, "NIE | ", fiTelefonKli))
        PrintLine(FileNum, "FAX=" & IIf(Len(fiFaxKli) = 0, "NIE | ", fiFaxKli))
        PrintLine(FileNum, "EMAIL=" & IIf(Len(fieMailKli) = 0, "NIE | ", fieMailKli))
        PrintLine(FileNum, "KTOWPISAL=" & IIf(Len(fiWpisalKli) = 0, "NIE | ", fiWpisalKli))
        PrintLine(FileNum, "REJKLIENTA=" & IIf(Len(fiRejonKli) = 0, "NIE | ", fiRejonKli))

        PrintLine(FileNum, "BRANZA=" & IIf(Len(fiBranzaKli) = 0, "NIE | ", fiBranzaKli))
        PrintLine(FileNum, "PODBRANZA=" & IIf(Len(fiPodBranzaKli) = 0, "NIE | ", fiPodBranzaKli))
        PrintLine(FileNum, "LICZBAZATRUDNIONYCH=" & IIf(Len(fiLiczbaZaktrudnionychKli) = 0, "NIE | ", fiLiczbaZaktrudnionychKli))
        PrintLine(FileNum, "LICZBAPRACZDALNIE=" & IIf(Len(fiLiczbaPracZdalnieKli) = 0, "NIE | ", fiLiczbaPracZdalnieKli))

        PrintLine(FileNum, "WYKAZURZADZEN=" & IIf(Len(fiWykazUrzadzenKli) = 0, "NIE | ", fiWykazUrzadzenKli))
        PrintLine(FileNum, "SPOSOBWYBORUDOSTAWCY=" & IIf(Len(fiSposobWyboruDostawcyKli) = 0, "NIE | ", fiSposobWyboruDostawcyKli))
        PrintLine(FileNum, "ZAINSTALOWANOMONITOR=" & IIf(Len(fiZainstalowanoMonitorKli) = 0, "NIE | ", fiZainstalowanoMonitorKli))
        PrintLine(FileNum, "LICZBAURZADZEN=" & IIf(Len(fiLiczbaUrzadzenKli) = 0, "NIE | ", fiLiczbaUrzadzenKli))
        PrintLine(FileNum, "LICZBAWYDRUKOW=" & IIf(Len(fiLiczbaWydrukowKli) = 0, "NIE | ", fiLiczbaWydrukowKli))
        PrintLine(FileNum, "POTENCJALDRUKU=" & IIf(Len(fiPotencjalDrukuKli) = 0, "NIE | ", fiPotencjalDrukuKli))
        PrintLine(FileNum, "AKTZAKUPMATEKSP=" & IIf(Len(fiAktZakupMatEkspKli) = 0, "NIE | ", fiAktZakupMatEkspKli))
        PrintLine(FileNum, "AKTROZLICZASTRONYDRUKU=" & IIf(Len(fiAktRozliczaStronyDrukuKli) = 0, "NIE | ", fiAktRozliczaStronyDrukuKli))
        PrintLine(FileNum, "AKTKORZYSTAZNAJMU=" & IIf(Len(fiAktKorzystaZNajmuKli) = 0, "NIE | ", fiAktKorzystaZNajmuKli))
        PrintLine(FileNum, "ZAINTROZLICZANIEMSTRONDRUKU=" & IIf(Len(fiZaintRozliczaniemStronDrukuKli) = 0, "NIE | ", fiZaintRozliczaniemStronDrukuKli))
        PrintLine(FileNum, "ZAINTKORZYSTANIEMZNAJMU=" & IIf(Len(fiZaintKorzystaniemZNajmuKli) = 0, "NIE | ", fiZaintKorzystaniemZNajmuKli))
        PrintLine(FileNum, "DATAWERYFSEGMENTACJI=" & IIf(Len(fiDataWeryfSegmentacjiKli) = 0, "NIE | ", fiDataWeryfSegmentacjiKli))

        PrintLine(FileNum, "RODZSPRZETUL=" & IIf(Len(fiRodzSprzetuLKli) = 0, "NIE | ", fiRodzSprzetuLKli))
        PrintLine(FileNum, "RODZSPRZETUA=" & IIf(Len(fiRodzSprzetuAKli) = 0, "NIE | ", fiRodzSprzetuAKli))
        PrintLine(FileNum, "RODZSPRZETUT=" & IIf(Len(fiRodzSprzetuTKli) = 0, "NIE | ", fiRodzSprzetuTKli))
        PrintLine(FileNum, "UWAGI=" & IIf(Len(fiUwagikli) = 0, "NIE | ", fiUwagikli))
        PrintLine(FileNum, "OFERTATZLOZENIA=" & IIf(Len(fiOfertaTZlozeniaKli) = 0, "NIE | ", fiOfertaTZlozeniaKli))
        PrintLine(FileNum, "OFERTAPLIK=" & IIf(Len(fiOfertaPlikKli) = 0, "NIE | ", fiOfertaPlikKli))
        PrintLine(FileNum, "PKONTAKT=" & IIf(Len(fiPKontaktKli) = 0, "NIE | ", fiPKontaktKli))

        PrintLine(FileNum, "SKUPWARTOSC=" & IIf(Len(fiSkupWartoscKli) = 0, "NIE | ", fiSkupWartoscKli))
        PrintLine(FileNum, "SPRZEDOPIEKUN=" & IIf(Len(fiSprzedOpiekunkli) = 0, "NIE | ", fiSprzedOpiekunkli))
        PrintLine(FileNum, "SPRZEDOKONTAKTR=" & IIf(Len(fiSprzedOKontaktRKli) = 0, "NIE | ", fiSprzedOKontaktRKli))
        PrintLine(FileNum, "SPRZEDOKONTAKTT=" & IIf(Len(fiSprzedOKontaktTKli) = 0, "NIE | ", fiSprzedOKontaktTKli))
        PrintLine(FileNum, "SPRZEDOKONTAKTD=" & IIf(Len(fiSprzedOKontaktDKli) = 0, "NIE | ", fiSprzedOKontaktDKli))
        PrintLine(FileNum, "SPRZEDNKONTAKTR=" & IIf(Len(fiSprzedNKontaktRKli) = 0, "NIE | ", fiSprzedNKontaktRKli))
        PrintLine(FileNum, "SPRZEDNKONTAKTT=" & IIf(Len(fiSprzedNKontaktTKli) = 0, "NIE | ", fiSprzedNKontaktTKli))
        PrintLine(FileNum, "SPRZEDNKONTAKTD=" & IIf(Len(fiSprzedNKontaktDKli) = 0, "NIE | ", fiSprzedNKontaktDKli))
        PrintLine(FileNum, "SPRZEDUWAGI=" & IIf(Len(fiSprzedUwagiKli) = 0, "NIE | ", fiSprzedUwagiKli))
        PrintLine(FileNum, "AKTYWNY=" & IIf(Len(fiAktywnyKli) = 0, "NIE | ", fiAktywnyKli))

        PrintLine(FileNum, "Koniec szablonu")
        FileClose(FileNum)
    End Sub

    Private Sub PobierzSzablon()
        Dim FileNum As Integer
        Dim ZnakRownaSie As Integer
        Dim NextLine As String
        Dim Klucz As String
        Dim Parametr As String

        If Len(Dir(oKatParametry & "\" & PlikZSzablonem)) > 0 Then
            FileNum = FreeFile()    'kolejny nr otwartego zbioru
            FileOpen(FileNum, oKatParametry & "\" & PlikZSzablonem, OpenMode.Input)
            If Not EOF(FileNum) Then
                NextLine = LineInput(FileNum)

                'sprawdz czy to plik z parametrami programu
                If InStr(NextLine, "SOFTART : Szablon filtrowania katalogu klientów") = 1 Then
                    Do Until EOF(FileNum)
                        NextLine = LineInput(FileNum)
                        ZnakRownaSie = InStr(NextLine, "=")
                        If ZnakRownaSie > 0 Then
                            Klucz = Mid(NextLine, 1, ZnakRownaSie - 1)
                            Parametr = Mid(NextLine, ZnakRownaSie + 1)
                            Select Case Klucz
                                Case "PLANDLUGOOKRESOWY"
                                    fiPlanDlugookresowyKli = Parametr
                                Case "PLANKROTKOOKRESOWY"
                                    fiPlanKrotkookresowyKli = Parametr


                                '-------------------Analizy sprzedazy
                                Case "ILOSCAA"
                                    fiIloscAA = Parametr
                                Case "ILOSCBB"
                                    fiIloscBB = Parametr

                                Case "ILOSCZAKBM"
                                    fiIlZak00 = Parametr
                                Case "ILOSCZAK01"
                                    fiIlZak01 = Parametr
                                Case "ILOSCZAK02"
                                    fiIlZak02 = Parametr
                                Case "ILOSCZAK03"
                                    fiIlZak03 = Parametr
                                Case "ILOSCZAK04"
                                    fiIlZak04 = Parametr
                                Case "ILOSCZAK05"
                                    fiIlZak05 = Parametr
                                Case "ILOSCZAK06"
                                    fiIlZak06 = Parametr
                                Case "ILOSCZAK07"
                                    fiIlZak07 = Parametr
                                Case "ILOSCZAK08"
                                    fiIlZak08 = Parametr
                                Case "ILOSCZAK09"
                                    fiIlZak09 = Parametr
                                Case "ILOSCZAK10"
                                    fiIlZak10 = Parametr
                                Case "ILOSCZAK11"
                                    fiIlZak11 = Parametr

                                Case "ILOSCZAK12"
                                    fiIlZak12 = Parametr
                                Case "ILOSCZAK13"
                                    fiIlZak13 = Parametr
                                Case "ILOSCZAK14"
                                    fiIlZak14 = Parametr
                                Case "ILOSCZAK15"
                                    fiIlZak15 = Parametr
                                Case "ILOSCZAK16"
                                    fiIlZak16 = Parametr
                                Case "ILOSCZAK17"
                                    fiIlZak17 = Parametr
                                Case "ILOSCZAK18"
                                    fiIlZak18 = Parametr
                                Case "ILOSCZAK19"
                                    fiIlZak19 = Parametr
                                Case "ILOSCZAK20"
                                    fiIlZak20 = Parametr
                                Case "ILOSCZAK21"
                                    fiIlZak21 = Parametr
                                Case "ILOSCZAK22"
                                    fiIlZak22 = Parametr
                                Case "ILOSCZAK23"
                                    fiIlZak23 = Parametr


                                    '-----------kONTAKTY HANDLOWE
                                Case "DATAKONTAKTU"
                                    fiDataKon = Parametr
                                Case "PRACOWNIKKONTAKTU"
                                    fiPracownikKon = Parametr
                                Case "RODZAJKONTAKTU"
                                    fiTematKon = Parametr
                                Case "OPISKONTAKTU"
                                    fiUwagiKon = Parametr


                                    '---------Katalog Klientow
                                Case "IDENTKLIENTA"
                                    fiIdentKli = Parametr
                                Case "NIP"
                                    fiNIPKli = Parametr
                                Case "NRTOFI"   'poprzednia wersja
                                    fiNrTOFITxtKli = Parametr
                                Case "NRTOFITXT"
                                    fiNrTOFITxtKli = Parametr

                                Case "KARTAPAYBACK"
                                    fiKartaPaybackKli = Parametr
                                Case "NRYKARTPAYBACK"
                                    fiNryKartPaybackKli = Parametr

                                Case "NAZWA1"
                                    fiNazwa1Kli = Parametr
                                Case "KODPOCZTOWY"
                                    fiKodPoczKli = Parametr
                                Case "MIEJSCOWOSC"
                                    fiMiejscKli = Parametr
                                Case "ULICA"
                                    fiUlicaKli = Parametr
                                Case "NRDOMU"
                                    fiNumNrDomuKli = Parametr
                                Case "IDDOMU"
                                    fiIDDomuKli = Parametr
                                Case "TELEFON"
                                    fiTelefonKli = Parametr
                                Case "FAX"
                                    fiFaxKli = Parametr
                                Case "EMAIL"
                                    fieMailKli = Parametr
                                Case "KTOWPISAL"
                                    fiWpisalKli = Parametr
                                Case "REJKLIENTA"
                                    fiRejonKli = Parametr

                                Case "BRANZA"
                                    fiBranzaKli = Parametr
                                Case "PODBRANZA"
                                    fiPodBranzaKli = Parametr
                                Case "LICZBAZATRUDNIONYCH"
                                    fiLiczbaZaktrudnionychKli = Parametr
                                Case "LICZBAPRACZDALNIE"
                                    fiLiczbaPracZdalnieKli = Parametr

                                Case "WYKAZURZADZEN"
                                    fiWykazUrzadzenKli = Parametr
                                Case "SPOSOBWYBORUDOSTAWCY"
                                    fiSposobWyboruDostawcyKli = Parametr
                                Case "ZAINSTALOWANOMONITOR"
                                    fiZainstalowanoMonitorKli = Parametr
                                Case "LICZBAURZADZEN"
                                    fiLiczbaUrzadzenKli = Parametr
                                Case "LICZBAWYDRUKOW"
                                    fiLiczbaWydrukowKli = Parametr
                                Case "POTENCJALDRUKU"
                                    fiPotencjalDrukuKli = Parametr
                                Case "AKTZAKUPMATEKSP"
                                    fiAktZakupMatEkspKli = Parametr
                                Case "AKTROZLICZASTRONYDRUKU"
                                    fiAktRozliczaStronyDrukuKli = Parametr
                                Case "AKTKORZYSTAZNAJMU"
                                    fiAktKorzystaZNajmuKli = Parametr
                                Case "ZAINTROZLICZANIEMSTRONDRUKU"
                                    fiZaintRozliczaniemStronDrukuKli = Parametr
                                Case "ZAINTKORZYSTANIEMZNAJMU"
                                    fiZaintKorzystaniemZNajmuKli = Parametr
                                Case "DATAWERYFSEGMENTACJI"
                                    fiDataWeryfSegmentacjiKli = Parametr

                                Case "RODZSPRZETUL"
                                    fiRodzSprzetuLKli = Parametr
                                Case "RODZSPRZETUA"
                                    fiRodzSprzetuAKli = Parametr
                                Case "RODZSPRZETUT"
                                    fiRodzSprzetuTKli = Parametr
                                Case "UWAGI"
                                    fiUwagikli = Parametr
                                Case "OFERTATZLOZENIA"
                                    fiOfertaTZlozeniaKli = Parametr
                                Case "OFERTAPLIK"
                                    fiOfertaPlikKli = Parametr
                                Case "PKONTAKT"
                                    fiPKontaktKli = Parametr

                                Case "SKUPWARTOSC"
                                    fiSkupWartoscKli = Parametr
                                Case "SPRZEDOPIEKUN"
                                    fiSprzedOpiekunkli = Parametr
                                Case "SPRZEDOKONTAKTR"
                                    fiSprzedOKontaktRKli = Parametr
                                Case "SPRZEDOKONTAKTT"
                                    fiSprzedOKontaktTKli = Parametr
                                Case "SPRZEDOKONTAKTD"
                                    fiSprzedOKontaktDKli = Parametr
                                Case "SPRZEDNKONTAKTR"
                                    fiSprzedNKontaktRKli = Parametr
                                Case "SPRZEDNKONTAKTT"
                                    fiSprzedNKontaktTKli = Parametr
                                Case "SPRZEDNKONTAKTD"
                                    fiSprzedNKontaktDKli = Parametr
                                Case "SPRZEDUWAGI"
                                    fiSprzedUwagiKli = Parametr
                                Case "AKTYWNY"
                                    fiAktywnyKli = Parametr




                            End Select
                        End If
                    Loop
                End If
                FileClose(FileNum)
            End If
        Else
            InitSzablon()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik ze schematem filtrowania"
            .InitialDirectory = oKatParametry
            .DefaultExt = "sf"
            .Filter = "Schemat filtrowania (*.sf)|*.sf|Wszystkie pliki|*.*"


            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                lblPath.Text = .FileName

                'oddziel plik od katalogow
                Dim Plik As String = .FileName
                Dim ChPos As String
                ChPos = InStr(Plik, "\")
                Do While ChPos > 0
                    Plik = Mid(Plik, ChPos + 1)
                    ChPos = InStr(Plik, "\")
                Loop
                'oddziel rozszerzenie pliku...
                ChPos = InStr(Plik, ".")
                txtSchemat.Text = Mid(Plik, 1, ChPos - 1)

                'pobierz schemat
                PlikZSzablonem = Plik
                PobierzSzablon()
                AktListePol()
            End If
        End With
    End Sub

    Private Sub cmdCzysc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCzysc.Click
        InitSzablon()
        AktListePol()
    End Sub

    Private Sub cmdZapisz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZapisz.Click
        If Len(Trim(txtSchemat.Text)) = 0 Then
            MessageBox.Show("Proszê okreœliæ nazwê szablonu...", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            PlikZSzablonem = Trim(txtSchemat.Text) & ".sf"
            ZapiszSzablon()
        End If
    End Sub


    '**********************************************************
    '** budujemy schemat filtrowania
    '**********************************************************

    Function BudujSzablon() As String
        Dim ind As Integer
        Dim iNazwa As String
        Dim iOpis As String
        Dim iTyp As String
        Dim iTestFiltr As String
        Dim iFiltr As String
        Dim iTest As String
        Dim iOper As String
        Dim TxtOd As String
        Dim TxtDo As String
        Dim pos As Integer

        Dim mySzablon As String = ""

        For ind = 0 To ListaPol.Items.Count - 1
            iNazwa = ListaPol.Items(ind).SubItems(0).Text
            iOpis = ListaPol.Items(ind).SubItems(1).Text
            iTyp = ListaPol.Items(ind).SubItems(2).Text
            iTestFiltr = ListaPol.Items(ind).SubItems(3).Text

            iTest = Mid(iTestFiltr, 1, 3)
            iOper = Mid(iTestFiltr, 7, 8)
            iFiltr = Mid(iTestFiltr, 16)

            If iTest = "Tak" Then
                If Len(mySzablon) > 0 Then
                    mySzablon &= " AND "
                Else
                End If

                Select Case iTyp
                    Case "Text"
                        Select Case iOper
                            Case "Puste   "
                                mySzablon &= iNazwa & " = ''"
                            Case "Niepuste"
                                mySzablon &= iNazwa & " <> ''"

                            Case "Zawiera "
                                If Len(Trim(iFiltr)) = 0 Then
                                    mySzablon &= iNazwa & " = ''"
                                Else
                                    mySzablon &= iNazwa & " LIKE '*" & Trim(iFiltr) & "*'"
                                End If
                            Case "Nie Zawi", "NieZawi."
                                If Len(Trim(iFiltr)) = 0 Then
                                    mySzablon &= iNazwa & " = ''"
                                Else
                                    mySzablon &= "NOT " & iNazwa & " LIKE '*" & Trim(iFiltr) & "*'"
                                End If

                            Case "Od...Do "
                                pos = InStr(iFiltr, " ... ")
                                If pos > 0 Then
                                    TxtOd = Mid(iFiltr, 1, pos - 1)
                                    TxtDo = Mid(iFiltr, pos + 5)
                                    mySzablon &= iNazwa & " >= '" & Trim(TxtOd) & "' AND " & iNazwa & " <= '" & Trim(TxtDo) & "'"
                                Else
                                    If Len(Trim(iFiltr)) = 0 Then
                                        mySzablon &= iNazwa & " = ''"
                                    Else
                                        mySzablon &= iNazwa & " LIKE '" & Trim(iFiltr) & "*'"
                                    End If
                                End If
                            Case Else
                                If Len(Trim(iFiltr)) = 0 Then
                                    mySzablon &= iNazwa & " = ''"
                                Else
                                    mySzablon &= iNazwa & " LIKE '" & Trim(iFiltr) & "*'"
                                End If
                        End Select

                    Case "Liczba"
                        iFiltr = PrzecNaKrop(iFiltr)    'zmien przecinek na kropke....
                        Select Case Mid(iOper, 1, 3)
                            Case "=  "
                                mySzablon &= iNazwa & " = " & iFiltr & ""
                            Case "<> "
                                mySzablon &= iNazwa & " <> " & iFiltr & ""
                            Case "<  "
                                mySzablon &= iNazwa & " < " & iFiltr & ""
                            Case "<= "
                                mySzablon &= iNazwa & " <= " & iFiltr & ""
                            Case ">  "
                                mySzablon &= iNazwa & " > " & iFiltr & ""
                            Case ">= "
                                mySzablon &= iNazwa & " >= " & iFiltr & ""
                            Case "Od."
                                pos = InStr(iFiltr, " ... ")
                                If pos > 0 Then
                                    TxtOd = Mid(iFiltr, 1, pos - 1)
                                    TxtDo = Mid(iFiltr, pos + 5)
                                    If Len(TxtOd) = 0 Then TxtOd = "0"
                                    If Len(TxtDo) = 0 Then TxtDo = "0"
                                    mySzablon &= iNazwa & " >= " & Trim(TxtOd) & " AND " & iNazwa & " <= " & Trim(TxtDo) & ""
                                Else
                                    If Len(Trim(iFiltr)) = 0 Then
                                        mySzablon &= iNazwa & " = 0"
                                    Else
                                        mySzablon &= iNazwa & " = " & Trim(iFiltr) & ""
                                    End If
                                End If
                            Case Else
                                mySzablon &= iNazwa & " = " & iFiltr & ""
                        End Select

                    Case "Tak/Nie"
                        If InStr(iFiltr, "Tak") > 0 Then
                            mySzablon &= iNazwa & " = TRUE"
                        Else
                            mySzablon &= iNazwa & " = FALSE"
                        End If

                    Case "Data"
                        TxtOd = Mid(iFiltr, 1, 10)
                        TxtDo = Mid(iFiltr, 16, 10)
                        mySzablon &= iNazwa & " >= '" & Trim(TxtOd) & "' AND " & iNazwa & " <= '" & Trim(TxtDo) & "'"
                End Select
            End If
        Next ind
        Return (mySzablon)
    End Function

    '***************************************************************
    '*** testuj czy testowac
    '***************************************************************

    Private Sub chbTestTxt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbTestTxt.CheckedChanged
        If chbTestTxt.Checked Then
            rabtextPuste.Visible = True
            rabTextNiepuste.Visible = True
            rabtextZawiera.Visible = True
            rabTextNieZawiera.Visible = True
            rabTexOdDo.Visible = True

            rabtextPuste.Checked = False
            rabTextNiepuste.Checked = False
            rabtextZawiera.Checked = True
            rabTextNieZawiera.Checked = False
            rabTexOdDo.Checked = False

            lblTextOd.Visible = True
            txtTextOd.Visible = True
            txtTextOd.Text = ""
            lblTextOd.Text = "Tekst : . . . . . ."
            lblTextDo.Visible = False
            txtTextDo.Text = ""
            txtTextDo.Visible = False
            lblTextDo.Text = ""
        Else
            rabtextPuste.Visible = False
            rabTextNiepuste.Visible = False
            rabtextZawiera.Visible = False
            rabTextNieZawiera.Visible = False
            rabTexOdDo.Visible = False

            lblTextOd.Visible = False
            txtTextOd.Visible = False
            txtTextOd.Text = ""
            lblTextOd.Text = ""
            lblTextDo.Visible = False
            txtTextDo.Visible = False
            txtTextDo.Text = ""
            lblTextDo.Text = ""
        End If
        txtTextOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub chbTestDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbTestDate.CheckedChanged
        If chbTestDate.Checked Then
            lblDataOd.Visible = True
            lblDataDo.Visible = True
            txtDataOd.Visible = True
            txtDataDo.Visible = True
            cmdDataOd.Visible = True
            cmdDataDo.Visible = True
        Else
            lblDataOd.Visible = False
            lblDataDo.Visible = False
            txtDataOd.Visible = False
            txtDataDo.Visible = False
            cmdDataOd.Visible = False
            cmdDataDo.Visible = False
        End If
        txtDataOd_TextChanged(Me, System.EventArgs.Empty)
        'txtDataDo_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub chbTestNum_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbTestNum.CheckedChanged
        If chbTestNum.Checked Then
            rabLiczbaRowne.Visible = True
            rabLiczbaRozne.Visible = True
            rabLiczbaMniesze.Visible = True
            rabLiczbaMniejRowne.Visible = True
            rabLiczbaWieksze.Visible = True
            rabLiczbaWieRowne.Visible = True
            rabLiczbaOdDo.Visible = True

            rabLiczbaOdDo.Checked = True

            lblLiczbaOd.Visible = True
            lblLiczbaDo.Visible = True
            TxtLiczbaOd.Visible = True
            txtLiczbaDo.Visible = True
        Else
            rabLiczbaRowne.Visible = False
            rabLiczbaRozne.Visible = False
            rabLiczbaMniesze.Visible = False
            rabLiczbaMniejRowne.Visible = False
            rabLiczbaWieksze.Visible = False
            rabLiczbaWieRowne.Visible = False
            rabLiczbaOdDo.Visible = False

            lblLiczbaOd.Visible = False
            lblLiczbaDo.Visible = False
            TxtLiczbaOd.Visible = False
            txtLiczbaDo.Visible = False
        End If
        TxtLiczbaOd_TextChanged(Me, System.EventArgs.Empty)
        'txtLiczbaDo_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub chbTestLog_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbTestLog.CheckedChanged
        If chbTestLog.Checked Then
            lblTak.Visible = True
            lblNie.Visible = True
            rbtTak.Visible = True
            rbtNie.Visible = True
        Else
            lblTak.Visible = False
            lblNie.Visible = False
            rbtTak.Visible = False
            rbtNie.Visible = False
        End If
        rbtTak_CheckedChanged(Me, System.EventArgs.Empty)
        'rbtNie_CheckedChanged(Me, System.EventArgs.Empty)
    End Sub

    '****************************************************
    'test poprawnoœci wprowadzanych wartosci
    '****************************************************

    Private Sub txtSchemat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSchemat.TextChanged
        TestLen(txtSchemat, "NAZWA SZABLONU", 10)
    End Sub

    Private Sub txtTextOd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTextOd.TextChanged
        TestLen(txtTextOd, "WARTOŒÆ FILTRA POLA TEKSTOWEGO OD", 50)
        '---------------------------
        If chbTestTxt.Checked Then
            If rabtextPuste.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "Puste    ")
            ElseIf rabTextNiepuste.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "Niepuste ")
            ElseIf rabtextZawiera.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "Zawiera  " & Trim(txtTextOd.Text))
            ElseIf rabTextNieZawiera.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "NieZawi. " & Trim(txtTextOd.Text))
            ElseIf rabTexOdDo.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "Od...Do  " & Trim(txtTextOd.Text) & " ... " & Trim(txtTextDo.Text))
            End If
        Else
            AktualFiltr(CurItemIndex, "Nie | ")
        End If
    End Sub

    Private Sub txtTextDo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTextDo.TextChanged
        TestLen(txtTextDo, "WARTOŒÆ FILTRA POLA TEKSTOWEGO DO", 50)
        '---------------------------
        If chbTestTxt.Checked Then
            If rabtextPuste.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "Puste    ")
            ElseIf rabTextNiepuste.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "Niepuste ")
            ElseIf rabtextZawiera.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "Zawiera  " & Trim(txtTextOd.Text))
            ElseIf rabTextNieZawiera.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "NieZawi. " & Trim(txtTextOd.Text))
            ElseIf rabTexOdDo.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "Od...Do  " & Trim(txtTextOd.Text) & " ... " & Trim(txtTextDo.Text))
            End If
        Else
            AktualFiltr(CurItemIndex, "Nie | ")
        End If
    End Sub





    Private Sub txtDataOd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDataOd.TextChanged
        TestLen(txtDataOd, "DATA Od", 10)
        If chbTestDate.Checked Then
            AktualFiltr(CurItemIndex, "Tak | " & "Od...Do  " & Trim(txtDataOd.Text) & " ... " & Trim(txtDataDo.Text))
        Else
            AktualFiltr(CurItemIndex, "Nie | ")
        End If
    End Sub

    Private Sub txtDataDo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDataDo.TextChanged
        TestLen(txtDataDo, "DATA Do", 10)
        If chbTestDate.Checked Then
            AktualFiltr(CurItemIndex, "Tak | " & "Od...Do  " & Trim(txtDataOd.Text) & " ... " & Trim(txtDataDo.Text))
        Else
            AktualFiltr(CurItemIndex, "Nie | ")
        End If
    End Sub




    Private Sub TxtLiczbaOd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtLiczbaOd.TextChanged
        'TestLen(TxtLiczbaOd, "LICZBA Od", 2)
        If TestNum(TxtLiczbaOd, "LICZBA Od") Then
            If chbTestNum.Checked Then
                If Len(Trim(txtLiczbaDo.Text)) = 0 Or IsNumeric(txtLiczbaDo.Text) Then
                    If rabLiczbaRowne.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "=        " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaRozne.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "<>       " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaMniesze.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "<        " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaMniejRowne.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "<=       " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaWieksze.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & ">        " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaWieRowne.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & ">=       " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaOdDo.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "Od...Do  " & Trim(TxtLiczbaOd.Text) & " ... " & Trim(txtLiczbaDo.Text))
                    End If
                Else
                    AktualFiltr(CurItemIndex, "Nie | ")
                End If
            Else
                AktualFiltr(CurItemIndex, "Nie | ")
            End If
        End If
    End Sub

    Private Sub txtLiczbaDo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaDo.TextChanged
        'TestLen(txtLiczbaDo, "LICZBA Do", 2)
        If TestNum(txtLiczbaDo, "LICZBA Do") Then
            If chbTestNum.Checked Then
                If Len(Trim(TxtLiczbaOd.Text)) = 0 Or IsNumeric(TxtLiczbaOd.Text) Then
                    If rabLiczbaRowne.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "=        " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaRozne.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "<>       " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaMniesze.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "<        " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaMniejRowne.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "<=       " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaWieksze.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & ">        " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaWieRowne.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & ">=       " & Trim(TxtLiczbaOd.Text))
                    ElseIf rabLiczbaOdDo.Checked Then
                        AktualFiltr(CurItemIndex, "Tak | " & "Od...Do  " & Trim(TxtLiczbaOd.Text) & " ... " & Trim(txtLiczbaDo.Text))
                    End If
                Else
                    AktualFiltr(CurItemIndex, "Nie | ")
                End If
            Else
                AktualFiltr(CurItemIndex, "Nie | ")
            End If
        End If
    End Sub





    Private Sub rbtTak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtTak.CheckedChanged
        If chbTestLog.Checked Then
            If rbtTak.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "         " & "Tak")
            Else
                AktualFiltr(CurItemIndex, "Tak | " & "         " & "Nie")
            End If
        Else
            AktualFiltr(CurItemIndex, "Nie | ")
        End If
    End Sub

    Private Sub rbtNie_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtNie.CheckedChanged
        If chbTestLog.Checked Then
            If rbtNie.Checked Then
                AktualFiltr(CurItemIndex, "Tak | " & "         " & "Nie")
            Else
                AktualFiltr(CurItemIndex, "Tak | " & "         " & "Tak")
            End If
        Else
            AktualFiltr(CurItemIndex, "Nie | ")
        End If
    End Sub












    Private Sub rabtextPuste_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabtextPuste.CheckedChanged
        lblTextOd.Visible = False
        txtTextOd.Visible = False
        txtTextOd.Text = ""
        lblTextDo.Visible = False
        txtTextDo.Visible = False
        txtTextDo.Text = ""
        '------------------------------
        txtTextOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabTextNiepuste_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabTextNiepuste.CheckedChanged
        lblTextOd.Visible = False
        txtTextOd.Visible = False
        txtTextOd.Text = ""
        lblTextDo.Visible = False
        txtTextDo.Visible = False
        txtTextDo.Text = ""
        '------------------------------
        txtTextOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabtextZawiera_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabtextZawiera.CheckedChanged
        lblTextOd.Visible = True
        lblTextOd.Text = "Tekst : . . . . . . . ."
        txtTextOd.Visible = True
        lblTextDo.Visible = False
        txtTextDo.Visible = False
        txtTextDo.Text = ""
        '------------------------------
        txtTextOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabtextNieZawiera_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabTextNieZawiera.CheckedChanged
        lblTextOd.Visible = True
        lblTextOd.Text = "Tekst : . . . . . . . ."
        txtTextOd.Visible = True
        lblTextDo.Visible = False
        txtTextDo.Visible = False
        txtTextDo.Text = ""
        '------------------------------
        txtTextOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabTexOdDo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabTexOdDo.CheckedChanged
        lblTextOd.Visible = True
        lblTextOd.Text = "Tekst od : . . . . . . . ."
        txtTextOd.Visible = True
        lblTextDo.Visible = True
        lblTextDo.Text = "Tekst do : . . . . . . . ."
        txtTextDo.Visible = True
        '------------------------------
        'txtTextOd_TextChanged(Me, System.EventArgs.Empty)
        txtTextDo_TextChanged(Me, System.EventArgs.Empty)
    End Sub




    Private Sub rabLiczbaRowne_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaRowne.CheckedChanged
        lblLiczbaOd.Visible = True
        TxtLiczbaOd.Visible = True
        lblLiczbaDo.Visible = False
        txtLiczbaDo.Visible = False
        TxtLiczbaOd.Text = ""
        TxtLiczbaOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabLiczbaRozne_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaRozne.CheckedChanged
        lblLiczbaOd.Visible = True
        TxtLiczbaOd.Visible = True
        lblLiczbaDo.Visible = False
        txtLiczbaDo.Visible = False
        TxtLiczbaOd.Text = ""
        TxtLiczbaOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabLiczbaMniesze_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaMniesze.CheckedChanged
        lblLiczbaOd.Visible = True
        TxtLiczbaOd.Visible = True
        lblLiczbaDo.Visible = False
        txtLiczbaDo.Visible = False
        TxtLiczbaOd.Text = ""
        TxtLiczbaOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabLiczbaMniejRowne_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaMniejRowne.CheckedChanged
        lblLiczbaOd.Visible = True
        TxtLiczbaOd.Visible = True
        lblLiczbaDo.Visible = False
        txtLiczbaDo.Visible = False
        TxtLiczbaOd.Text = ""
        TxtLiczbaOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabLiczbaWieksze_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaWieksze.CheckedChanged
        lblLiczbaOd.Visible = True
        TxtLiczbaOd.Visible = True
        lblLiczbaDo.Visible = False
        txtLiczbaDo.Visible = False
        TxtLiczbaOd.Text = ""
        TxtLiczbaOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabLiczbaWieRowne_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaWieRowne.CheckedChanged
        lblLiczbaOd.Visible = True
        TxtLiczbaOd.Visible = True
        lblLiczbaDo.Visible = False
        txtLiczbaDo.Visible = False
        TxtLiczbaOd.Text = ""
        TxtLiczbaOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabLiczbaOdDo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaOdDo.CheckedChanged
        lblLiczbaOd.Visible = True
        TxtLiczbaOd.Visible = True
        lblLiczbaDo.Visible = True
        txtLiczbaDo.Visible = True
        TxtLiczbaOd.Text = ""
        txtLiczbaDo.Text = ""
        txtLiczbaDo_TextChanged(Me, System.EventArgs.Empty)
    End Sub
End Class
