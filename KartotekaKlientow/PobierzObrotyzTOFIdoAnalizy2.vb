Public Class PobierzObrotyzTOFIdoAnalizy2
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Friend WithEvents cbxRok2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cbxMies2 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxRok1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmdOdznacz2 As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents cmdZaznacz2 As Button
    Friend WithEvents cmdOdznacz1 As Button
    Friend WithEvents cmdZaznacz1 As Button
    Friend WithEvents chkOrygTonery As CheckBox
    Friend WithEvents chkOrygAtrament As CheckBox
    Friend WithEvents chkPryzmatTonery As CheckBox
    Friend WithEvents chkPryzmatAtrament As CheckBox
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Label16 As Label
    Friend WithEvents cbxCoAnalizowac As ComboBox
    Friend WithEvents cbxMies1 As System.Windows.Forms.ComboBox

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
    Friend WithEvents lblIlePrzeczytano As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdImportuj As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblDataAktualizacji As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chkStrony As System.Windows.Forms.CheckBox
    Friend WithEvents chkNajem As System.Windows.Forms.CheckBox
    Friend WithEvents chkOryg As System.Windows.Forms.CheckBox
    Friend WithEvents chkPryzmat As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkSkup As System.Windows.Forms.CheckBox
    Friend WithEvents chkDrukarki As System.Windows.Forms.CheckBox
    Friend WithEvents Chk112 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk111 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk110 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk109 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk108 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk107 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk106 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk105 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk104 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk103 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk102 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk101 As System.Windows.Forms.CheckBox
    Friend WithEvents lblZakres1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Chk212 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk211 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk210 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk209 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk208 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk207 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk206 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk205 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk204 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk203 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk202 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk201 As System.Windows.Forms.CheckBox
    Friend WithEvents lblZakres2 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PobierzObrotyzTOFIdoAnalizy2))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cbxCoAnalizowac = New System.Windows.Forms.ComboBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.chkOrygTonery = New System.Windows.Forms.CheckBox()
        Me.chkOrygAtrament = New System.Windows.Forms.CheckBox()
        Me.chkPryzmatTonery = New System.Windows.Forms.CheckBox()
        Me.chkPryzmatAtrament = New System.Windows.Forms.CheckBox()
        Me.lblIlePrzeczytano = New System.Windows.Forms.Label()
        Me.lblDataAktualizacji = New System.Windows.Forms.Label()
        Me.cmdOdznacz2 = New System.Windows.Forms.Button()
        Me.cmdZaznacz2 = New System.Windows.Forms.Button()
        Me.cmdOdznacz1 = New System.Windows.Forms.Button()
        Me.cmdZaznacz1 = New System.Windows.Forms.Button()
        Me.cbxRok2 = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cbxMies2 = New System.Windows.Forms.ComboBox()
        Me.cbxRok1 = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cbxMies1 = New System.Windows.Forms.ComboBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Chk212 = New System.Windows.Forms.CheckBox()
        Me.Chk211 = New System.Windows.Forms.CheckBox()
        Me.Chk210 = New System.Windows.Forms.CheckBox()
        Me.Chk209 = New System.Windows.Forms.CheckBox()
        Me.Chk208 = New System.Windows.Forms.CheckBox()
        Me.Chk207 = New System.Windows.Forms.CheckBox()
        Me.Chk206 = New System.Windows.Forms.CheckBox()
        Me.Chk205 = New System.Windows.Forms.CheckBox()
        Me.Chk204 = New System.Windows.Forms.CheckBox()
        Me.Chk203 = New System.Windows.Forms.CheckBox()
        Me.Chk202 = New System.Windows.Forms.CheckBox()
        Me.Chk201 = New System.Windows.Forms.CheckBox()
        Me.lblZakres2 = New System.Windows.Forms.Label()
        Me.Chk112 = New System.Windows.Forms.CheckBox()
        Me.Chk111 = New System.Windows.Forms.CheckBox()
        Me.Chk110 = New System.Windows.Forms.CheckBox()
        Me.Chk109 = New System.Windows.Forms.CheckBox()
        Me.Chk108 = New System.Windows.Forms.CheckBox()
        Me.Chk107 = New System.Windows.Forms.CheckBox()
        Me.Chk106 = New System.Windows.Forms.CheckBox()
        Me.Chk105 = New System.Windows.Forms.CheckBox()
        Me.Chk104 = New System.Windows.Forms.CheckBox()
        Me.Chk103 = New System.Windows.Forms.CheckBox()
        Me.Chk102 = New System.Windows.Forms.CheckBox()
        Me.Chk101 = New System.Windows.Forms.CheckBox()
        Me.lblZakres1 = New System.Windows.Forms.Label()
        Me.chkSkup = New System.Windows.Forms.CheckBox()
        Me.chkDrukarki = New System.Windows.Forms.CheckBox()
        Me.chkStrony = New System.Windows.Forms.CheckBox()
        Me.chkNajem = New System.Windows.Forms.CheckBox()
        Me.chkOryg = New System.Windows.Forms.CheckBox()
        Me.chkPryzmat = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.cmdImportuj = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
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
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.cbxCoAnalizowac)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.chkOrygTonery)
        Me.Panel1.Controls.Add(Me.chkOrygAtrament)
        Me.Panel1.Controls.Add(Me.chkPryzmatTonery)
        Me.Panel1.Controls.Add(Me.chkPryzmatAtrament)
        Me.Panel1.Controls.Add(Me.lblIlePrzeczytano)
        Me.Panel1.Controls.Add(Me.lblDataAktualizacji)
        Me.Panel1.Controls.Add(Me.cmdOdznacz2)
        Me.Panel1.Controls.Add(Me.cmdZaznacz2)
        Me.Panel1.Controls.Add(Me.cmdOdznacz1)
        Me.Panel1.Controls.Add(Me.cmdZaznacz1)
        Me.Panel1.Controls.Add(Me.cbxRok2)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.cbxMies2)
        Me.Panel1.Controls.Add(Me.cbxRok1)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.cbxMies1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Chk212)
        Me.Panel1.Controls.Add(Me.Chk211)
        Me.Panel1.Controls.Add(Me.Chk210)
        Me.Panel1.Controls.Add(Me.Chk209)
        Me.Panel1.Controls.Add(Me.Chk208)
        Me.Panel1.Controls.Add(Me.Chk207)
        Me.Panel1.Controls.Add(Me.Chk206)
        Me.Panel1.Controls.Add(Me.Chk205)
        Me.Panel1.Controls.Add(Me.Chk204)
        Me.Panel1.Controls.Add(Me.Chk203)
        Me.Panel1.Controls.Add(Me.Chk202)
        Me.Panel1.Controls.Add(Me.Chk201)
        Me.Panel1.Controls.Add(Me.lblZakres2)
        Me.Panel1.Controls.Add(Me.Chk112)
        Me.Panel1.Controls.Add(Me.Chk111)
        Me.Panel1.Controls.Add(Me.Chk110)
        Me.Panel1.Controls.Add(Me.Chk109)
        Me.Panel1.Controls.Add(Me.Chk108)
        Me.Panel1.Controls.Add(Me.Chk107)
        Me.Panel1.Controls.Add(Me.Chk106)
        Me.Panel1.Controls.Add(Me.Chk105)
        Me.Panel1.Controls.Add(Me.Chk104)
        Me.Panel1.Controls.Add(Me.Chk103)
        Me.Panel1.Controls.Add(Me.Chk102)
        Me.Panel1.Controls.Add(Me.Chk101)
        Me.Panel1.Controls.Add(Me.lblZakres1)
        Me.Panel1.Controls.Add(Me.chkSkup)
        Me.Panel1.Controls.Add(Me.chkDrukarki)
        Me.Panel1.Controls.Add(Me.chkStrony)
        Me.Panel1.Controls.Add(Me.chkNajem)
        Me.Panel1.Controls.Add(Me.chkOryg)
        Me.Panel1.Controls.Add(Me.chkPryzmat)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.ProgressBar)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(694, 426)
        Me.Panel1.TabIndex = 30
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Navy
        Me.Label16.Location = New System.Drawing.Point(8, 355)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(113, 16)
        Me.Label16.TabIndex = 255
        Me.Label16.Text = "DANE DO ANALIZY :"
        '
        'cbxCoAnalizowac
        '
        Me.cbxCoAnalizowac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCoAnalizowac.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxCoAnalizowac.ForeColor = System.Drawing.Color.Purple
        Me.cbxCoAnalizowac.Location = New System.Drawing.Point(126, 352)
        Me.cbxCoAnalizowac.Name = "cbxCoAnalizowac"
        Me.cbxCoAnalizowac.Size = New System.Drawing.Size(216, 22)
        Me.cbxCoAnalizowac.TabIndex = 254
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Location = New System.Drawing.Point(0, 340)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(693, 1)
        Me.Panel4.TabIndex = 244
        '
        'chkOrygTonery
        '
        Me.chkOrygTonery.ForeColor = System.Drawing.Color.Navy
        Me.chkOrygTonery.Location = New System.Drawing.Point(310, 246)
        Me.chkOrygTonery.Name = "chkOrygTonery"
        Me.chkOrygTonery.Size = New System.Drawing.Size(160, 17)
        Me.chkOrygTonery.TabIndex = 243
        Me.chkOrygTonery.Text = "Oryginalne TONERY"
        Me.chkOrygTonery.UseVisualStyleBackColor = True
        '
        'chkOrygAtrament
        '
        Me.chkOrygAtrament.ForeColor = System.Drawing.Color.Navy
        Me.chkOrygAtrament.Location = New System.Drawing.Point(164, 246)
        Me.chkOrygAtrament.Name = "chkOrygAtrament"
        Me.chkOrygAtrament.Size = New System.Drawing.Size(160, 17)
        Me.chkOrygAtrament.TabIndex = 242
        Me.chkOrygAtrament.Text = "Oryginalne ATRAMENT"
        Me.chkOrygAtrament.UseVisualStyleBackColor = True
        '
        'chkPryzmatTonery
        '
        Me.chkPryzmatTonery.ForeColor = System.Drawing.Color.Navy
        Me.chkPryzmatTonery.Location = New System.Drawing.Point(310, 229)
        Me.chkPryzmatTonery.Name = "chkPryzmatTonery"
        Me.chkPryzmatTonery.Size = New System.Drawing.Size(160, 17)
        Me.chkPryzmatTonery.TabIndex = 241
        Me.chkPryzmatTonery.Text = "Pryzmat TONERY"
        Me.chkPryzmatTonery.UseVisualStyleBackColor = True
        '
        'chkPryzmatAtrament
        '
        Me.chkPryzmatAtrament.ForeColor = System.Drawing.Color.Navy
        Me.chkPryzmatAtrament.Location = New System.Drawing.Point(164, 229)
        Me.chkPryzmatAtrament.Name = "chkPryzmatAtrament"
        Me.chkPryzmatAtrament.Size = New System.Drawing.Size(160, 17)
        Me.chkPryzmatAtrament.TabIndex = 240
        Me.chkPryzmatAtrament.Text = "Pryzmat ATRAMENT"
        Me.chkPryzmatAtrament.UseVisualStyleBackColor = True
        '
        'lblIlePrzeczytano
        '
        Me.lblIlePrzeczytano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlePrzeczytano.ForeColor = System.Drawing.Color.Purple
        Me.lblIlePrzeczytano.Location = New System.Drawing.Point(577, 370)
        Me.lblIlePrzeczytano.Name = "lblIlePrzeczytano"
        Me.lblIlePrzeczytano.Size = New System.Drawing.Size(88, 17)
        Me.lblIlePrzeczytano.TabIndex = 26
        Me.lblIlePrzeczytano.Text = "0"
        Me.lblIlePrzeczytano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDataAktualizacji
        '
        Me.lblDataAktualizacji.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDataAktualizacji.ForeColor = System.Drawing.Color.Purple
        Me.lblDataAktualizacji.Location = New System.Drawing.Point(577, 353)
        Me.lblDataAktualizacji.Name = "lblDataAktualizacji"
        Me.lblDataAktualizacji.Size = New System.Drawing.Size(88, 17)
        Me.lblDataAktualizacji.TabIndex = 56
        Me.lblDataAktualizacji.Text = "0"
        Me.lblDataAktualizacji.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdOdznacz2
        '
        Me.cmdOdznacz2.Image = CType(resources.GetObject("cmdOdznacz2.Image"), System.Drawing.Image)
        Me.cmdOdznacz2.Location = New System.Drawing.Point(642, 167)
        Me.cmdOdznacz2.Name = "cmdOdznacz2"
        Me.cmdOdznacz2.Size = New System.Drawing.Size(24, 21)
        Me.cmdOdznacz2.TabIndex = 239
        Me.cmdOdznacz2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdznacz2, "Odznacz wszystko")
        '
        'cmdZaznacz2
        '
        Me.cmdZaznacz2.Image = CType(resources.GetObject("cmdZaznacz2.Image"), System.Drawing.Image)
        Me.cmdZaznacz2.Location = New System.Drawing.Point(612, 167)
        Me.cmdZaznacz2.Name = "cmdZaznacz2"
        Me.cmdZaznacz2.Size = New System.Drawing.Size(24, 21)
        Me.cmdZaznacz2.TabIndex = 238
        Me.cmdZaznacz2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdZaznacz2, "Zaznacz wszystko")
        '
        'cmdOdznacz1
        '
        Me.cmdOdznacz1.Image = CType(resources.GetObject("cmdOdznacz1.Image"), System.Drawing.Image)
        Me.cmdOdznacz1.Location = New System.Drawing.Point(642, 147)
        Me.cmdOdznacz1.Name = "cmdOdznacz1"
        Me.cmdOdznacz1.Size = New System.Drawing.Size(24, 21)
        Me.cmdOdznacz1.TabIndex = 237
        Me.cmdOdznacz1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdznacz1, "Odznacz wszystko")
        '
        'cmdZaznacz1
        '
        Me.cmdZaznacz1.Image = CType(resources.GetObject("cmdZaznacz1.Image"), System.Drawing.Image)
        Me.cmdZaznacz1.Location = New System.Drawing.Point(612, 147)
        Me.cmdZaznacz1.Name = "cmdZaznacz1"
        Me.cmdZaznacz1.Size = New System.Drawing.Size(24, 21)
        Me.cmdZaznacz1.TabIndex = 236
        Me.cmdZaznacz1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdZaznacz1, "Zaznacz wszystko")
        '
        'cbxRok2
        '
        Me.cbxRok2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxRok2.ForeColor = System.Drawing.Color.Purple
        Me.cbxRok2.Location = New System.Drawing.Point(135, 169)
        Me.cbxRok2.Name = "cbxRok2"
        Me.cbxRok2.Size = New System.Drawing.Size(50, 22)
        Me.cbxRok2.TabIndex = 233
        Me.cbxRok2.Text = "2007"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(8, 173)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(131, 16)
        Me.Label8.TabIndex = 235
        Me.Label8.Text = "II. Start analizy (rok-mies) . . . ."
        '
        'cbxMies2
        '
        Me.cbxMies2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxMies2.ForeColor = System.Drawing.Color.Purple
        Me.cbxMies2.Location = New System.Drawing.Point(191, 169)
        Me.cbxMies2.Name = "cbxMies2"
        Me.cbxMies2.Size = New System.Drawing.Size(37, 22)
        Me.cbxMies2.TabIndex = 234
        Me.cbxMies2.Text = "12"
        '
        'cbxRok1
        '
        Me.cbxRok1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxRok1.ForeColor = System.Drawing.Color.Purple
        Me.cbxRok1.Location = New System.Drawing.Point(135, 148)
        Me.cbxRok1.Name = "cbxRok1"
        Me.cbxRok1.Size = New System.Drawing.Size(50, 22)
        Me.cbxRok1.TabIndex = 230
        Me.cbxRok1.Text = "2007"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(8, 152)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(131, 16)
        Me.Label7.TabIndex = 232
        Me.Label7.Text = "I.  Start analizy (rok-mies) . . . ."
        '
        'cbxMies1
        '
        Me.cbxMies1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxMies1.ForeColor = System.Drawing.Color.Purple
        Me.cbxMies1.Location = New System.Drawing.Point(191, 148)
        Me.cbxMies1.Name = "cbxMies1"
        Me.cbxMies1.Size = New System.Drawing.Size(37, 22)
        Me.cbxMies1.TabIndex = 231
        Me.cbxMies1.Text = "12"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(-1, 198)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(693, 1)
        Me.Panel2.TabIndex = 229
        '
        'Chk212
        '
        Me.Chk212.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk212.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk212.ForeColor = System.Drawing.Color.Navy
        Me.Chk212.Location = New System.Drawing.Point(577, 170)
        Me.Chk212.Name = "Chk212"
        Me.Chk212.Size = New System.Drawing.Size(18, 18)
        Me.Chk212.TabIndex = 228
        Me.Chk212.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk212.UseVisualStyleBackColor = True
        '
        'Chk211
        '
        Me.Chk211.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk211.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk211.ForeColor = System.Drawing.Color.Navy
        Me.Chk211.Location = New System.Drawing.Point(559, 170)
        Me.Chk211.Name = "Chk211"
        Me.Chk211.Size = New System.Drawing.Size(18, 18)
        Me.Chk211.TabIndex = 227
        Me.Chk211.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk211.UseVisualStyleBackColor = True
        '
        'Chk210
        '
        Me.Chk210.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk210.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk210.ForeColor = System.Drawing.Color.Navy
        Me.Chk210.Location = New System.Drawing.Point(541, 170)
        Me.Chk210.Name = "Chk210"
        Me.Chk210.Size = New System.Drawing.Size(18, 18)
        Me.Chk210.TabIndex = 226
        Me.Chk210.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk210.UseVisualStyleBackColor = True
        '
        'Chk209
        '
        Me.Chk209.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk209.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk209.ForeColor = System.Drawing.Color.Navy
        Me.Chk209.Location = New System.Drawing.Point(523, 170)
        Me.Chk209.Name = "Chk209"
        Me.Chk209.Size = New System.Drawing.Size(18, 18)
        Me.Chk209.TabIndex = 225
        Me.Chk209.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk209.UseVisualStyleBackColor = True
        '
        'Chk208
        '
        Me.Chk208.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk208.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk208.ForeColor = System.Drawing.Color.Navy
        Me.Chk208.Location = New System.Drawing.Point(505, 170)
        Me.Chk208.Name = "Chk208"
        Me.Chk208.Size = New System.Drawing.Size(18, 18)
        Me.Chk208.TabIndex = 224
        Me.Chk208.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk208.UseVisualStyleBackColor = True
        '
        'Chk207
        '
        Me.Chk207.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk207.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk207.ForeColor = System.Drawing.Color.Navy
        Me.Chk207.Location = New System.Drawing.Point(487, 170)
        Me.Chk207.Name = "Chk207"
        Me.Chk207.Size = New System.Drawing.Size(18, 18)
        Me.Chk207.TabIndex = 223
        Me.Chk207.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk207.UseVisualStyleBackColor = True
        '
        'Chk206
        '
        Me.Chk206.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk206.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk206.ForeColor = System.Drawing.Color.Navy
        Me.Chk206.Location = New System.Drawing.Point(469, 170)
        Me.Chk206.Name = "Chk206"
        Me.Chk206.Size = New System.Drawing.Size(18, 18)
        Me.Chk206.TabIndex = 222
        Me.Chk206.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk206.UseVisualStyleBackColor = True
        '
        'Chk205
        '
        Me.Chk205.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk205.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk205.ForeColor = System.Drawing.Color.Navy
        Me.Chk205.Location = New System.Drawing.Point(451, 170)
        Me.Chk205.Name = "Chk205"
        Me.Chk205.Size = New System.Drawing.Size(18, 18)
        Me.Chk205.TabIndex = 221
        Me.Chk205.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk205.UseVisualStyleBackColor = True
        '
        'Chk204
        '
        Me.Chk204.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk204.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk204.ForeColor = System.Drawing.Color.Navy
        Me.Chk204.Location = New System.Drawing.Point(433, 170)
        Me.Chk204.Name = "Chk204"
        Me.Chk204.Size = New System.Drawing.Size(18, 18)
        Me.Chk204.TabIndex = 220
        Me.Chk204.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk204.UseVisualStyleBackColor = True
        '
        'Chk203
        '
        Me.Chk203.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk203.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk203.ForeColor = System.Drawing.Color.Navy
        Me.Chk203.Location = New System.Drawing.Point(415, 170)
        Me.Chk203.Name = "Chk203"
        Me.Chk203.Size = New System.Drawing.Size(18, 18)
        Me.Chk203.TabIndex = 219
        Me.Chk203.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk203.UseVisualStyleBackColor = True
        '
        'Chk202
        '
        Me.Chk202.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk202.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk202.ForeColor = System.Drawing.Color.Navy
        Me.Chk202.Location = New System.Drawing.Point(397, 170)
        Me.Chk202.Name = "Chk202"
        Me.Chk202.Size = New System.Drawing.Size(18, 18)
        Me.Chk202.TabIndex = 218
        Me.Chk202.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk202.UseVisualStyleBackColor = True
        '
        'Chk201
        '
        Me.Chk201.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk201.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk201.ForeColor = System.Drawing.Color.Navy
        Me.Chk201.Location = New System.Drawing.Point(379, 170)
        Me.Chk201.Name = "Chk201"
        Me.Chk201.Size = New System.Drawing.Size(18, 18)
        Me.Chk201.TabIndex = 217
        Me.Chk201.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk201.UseVisualStyleBackColor = True
        '
        'lblZakres2
        '
        Me.lblZakres2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblZakres2.ForeColor = System.Drawing.Color.Navy
        Me.lblZakres2.Location = New System.Drawing.Point(234, 173)
        Me.lblZakres2.Name = "lblZakres2"
        Me.lblZakres2.Size = New System.Drawing.Size(160, 16)
        Me.lblZakres2.TabIndex = 216
        Me.lblZakres2.Text = "Zakres analizy 0..-11 . . . ."
        '
        'Chk112
        '
        Me.Chk112.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk112.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk112.ForeColor = System.Drawing.Color.Navy
        Me.Chk112.Location = New System.Drawing.Point(577, 149)
        Me.Chk112.Name = "Chk112"
        Me.Chk112.Size = New System.Drawing.Size(18, 18)
        Me.Chk112.TabIndex = 215
        Me.Chk112.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk112.UseVisualStyleBackColor = True
        '
        'Chk111
        '
        Me.Chk111.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk111.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk111.ForeColor = System.Drawing.Color.Navy
        Me.Chk111.Location = New System.Drawing.Point(559, 149)
        Me.Chk111.Name = "Chk111"
        Me.Chk111.Size = New System.Drawing.Size(18, 18)
        Me.Chk111.TabIndex = 214
        Me.Chk111.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk111.UseVisualStyleBackColor = True
        '
        'Chk110
        '
        Me.Chk110.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk110.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk110.ForeColor = System.Drawing.Color.Navy
        Me.Chk110.Location = New System.Drawing.Point(541, 149)
        Me.Chk110.Name = "Chk110"
        Me.Chk110.Size = New System.Drawing.Size(18, 18)
        Me.Chk110.TabIndex = 213
        Me.Chk110.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk110.UseVisualStyleBackColor = True
        '
        'Chk109
        '
        Me.Chk109.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk109.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk109.ForeColor = System.Drawing.Color.Navy
        Me.Chk109.Location = New System.Drawing.Point(523, 149)
        Me.Chk109.Name = "Chk109"
        Me.Chk109.Size = New System.Drawing.Size(18, 18)
        Me.Chk109.TabIndex = 212
        Me.Chk109.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk109.UseVisualStyleBackColor = True
        '
        'Chk108
        '
        Me.Chk108.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk108.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk108.ForeColor = System.Drawing.Color.Navy
        Me.Chk108.Location = New System.Drawing.Point(505, 149)
        Me.Chk108.Name = "Chk108"
        Me.Chk108.Size = New System.Drawing.Size(18, 18)
        Me.Chk108.TabIndex = 211
        Me.Chk108.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk108.UseVisualStyleBackColor = True
        '
        'Chk107
        '
        Me.Chk107.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk107.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk107.ForeColor = System.Drawing.Color.Navy
        Me.Chk107.Location = New System.Drawing.Point(487, 149)
        Me.Chk107.Name = "Chk107"
        Me.Chk107.Size = New System.Drawing.Size(18, 18)
        Me.Chk107.TabIndex = 210
        Me.Chk107.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk107.UseVisualStyleBackColor = True
        '
        'Chk106
        '
        Me.Chk106.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk106.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk106.ForeColor = System.Drawing.Color.Navy
        Me.Chk106.Location = New System.Drawing.Point(469, 149)
        Me.Chk106.Name = "Chk106"
        Me.Chk106.Size = New System.Drawing.Size(18, 18)
        Me.Chk106.TabIndex = 209
        Me.Chk106.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk106.UseVisualStyleBackColor = True
        '
        'Chk105
        '
        Me.Chk105.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk105.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk105.ForeColor = System.Drawing.Color.Navy
        Me.Chk105.Location = New System.Drawing.Point(451, 149)
        Me.Chk105.Name = "Chk105"
        Me.Chk105.Size = New System.Drawing.Size(18, 18)
        Me.Chk105.TabIndex = 208
        Me.Chk105.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk105.UseVisualStyleBackColor = True
        '
        'Chk104
        '
        Me.Chk104.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk104.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk104.ForeColor = System.Drawing.Color.Navy
        Me.Chk104.Location = New System.Drawing.Point(433, 149)
        Me.Chk104.Name = "Chk104"
        Me.Chk104.Size = New System.Drawing.Size(18, 18)
        Me.Chk104.TabIndex = 207
        Me.Chk104.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk104.UseVisualStyleBackColor = True
        '
        'Chk103
        '
        Me.Chk103.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk103.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk103.ForeColor = System.Drawing.Color.Navy
        Me.Chk103.Location = New System.Drawing.Point(415, 149)
        Me.Chk103.Name = "Chk103"
        Me.Chk103.Size = New System.Drawing.Size(18, 18)
        Me.Chk103.TabIndex = 206
        Me.Chk103.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk103.UseVisualStyleBackColor = True
        '
        'Chk102
        '
        Me.Chk102.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk102.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk102.ForeColor = System.Drawing.Color.Navy
        Me.Chk102.Location = New System.Drawing.Point(397, 149)
        Me.Chk102.Name = "Chk102"
        Me.Chk102.Size = New System.Drawing.Size(18, 18)
        Me.Chk102.TabIndex = 205
        Me.Chk102.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk102.UseVisualStyleBackColor = True
        '
        'Chk101
        '
        Me.Chk101.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk101.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk101.ForeColor = System.Drawing.Color.Navy
        Me.Chk101.Location = New System.Drawing.Point(379, 149)
        Me.Chk101.Name = "Chk101"
        Me.Chk101.Size = New System.Drawing.Size(18, 18)
        Me.Chk101.TabIndex = 204
        Me.Chk101.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk101.UseVisualStyleBackColor = True
        '
        'lblZakres1
        '
        Me.lblZakres1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblZakres1.ForeColor = System.Drawing.Color.Navy
        Me.lblZakres1.Location = New System.Drawing.Point(234, 152)
        Me.lblZakres1.Name = "lblZakres1"
        Me.lblZakres1.Size = New System.Drawing.Size(160, 16)
        Me.lblZakres1.TabIndex = 203
        Me.lblZakres1.Text = "Zakres 12.2018 .. 01-2018 . . . ."
        '
        'chkSkup
        '
        Me.chkSkup.ForeColor = System.Drawing.Color.Navy
        Me.chkSkup.Location = New System.Drawing.Point(38, 315)
        Me.chkSkup.Name = "chkSkup"
        Me.chkSkup.Size = New System.Drawing.Size(160, 17)
        Me.chkSkup.TabIndex = 67
        Me.chkSkup.Text = "Skup"
        Me.chkSkup.UseVisualStyleBackColor = True
        '
        'chkDrukarki
        '
        Me.chkDrukarki.ForeColor = System.Drawing.Color.Navy
        Me.chkDrukarki.Location = New System.Drawing.Point(38, 298)
        Me.chkDrukarki.Name = "chkDrukarki"
        Me.chkDrukarki.Size = New System.Drawing.Size(160, 17)
        Me.chkDrukarki.TabIndex = 66
        Me.chkDrukarki.Text = "Drukarki"
        Me.chkDrukarki.UseVisualStyleBackColor = True
        '
        'chkStrony
        '
        Me.chkStrony.ForeColor = System.Drawing.Color.Navy
        Me.chkStrony.Location = New System.Drawing.Point(38, 280)
        Me.chkStrony.Name = "chkStrony"
        Me.chkStrony.Size = New System.Drawing.Size(160, 17)
        Me.chkStrony.TabIndex = 64
        Me.chkStrony.Text = "Strony"
        Me.chkStrony.UseVisualStyleBackColor = True
        '
        'chkNajem
        '
        Me.chkNajem.ForeColor = System.Drawing.Color.Navy
        Me.chkNajem.Location = New System.Drawing.Point(38, 263)
        Me.chkNajem.Name = "chkNajem"
        Me.chkNajem.Size = New System.Drawing.Size(160, 17)
        Me.chkNajem.TabIndex = 63
        Me.chkNajem.Text = "Najem"
        Me.chkNajem.UseVisualStyleBackColor = True
        '
        'chkOryg
        '
        Me.chkOryg.ForeColor = System.Drawing.Color.Navy
        Me.chkOryg.Location = New System.Drawing.Point(38, 246)
        Me.chkOryg.Name = "chkOryg"
        Me.chkOryg.Size = New System.Drawing.Size(160, 17)
        Me.chkOryg.TabIndex = 62
        Me.chkOryg.Text = "Oryginalne RAZEM"
        Me.chkOryg.UseVisualStyleBackColor = True
        '
        'chkPryzmat
        '
        Me.chkPryzmat.ForeColor = System.Drawing.Color.Navy
        Me.chkPryzmat.Location = New System.Drawing.Point(38, 229)
        Me.chkPryzmat.Name = "chkPryzmat"
        Me.chkPryzmat.Size = New System.Drawing.Size(160, 17)
        Me.chkPryzmat.TabIndex = 60
        Me.chkPryzmat.Text = "Pryzmat RAZEM"
        Me.chkPryzmat.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 212)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(334, 16)
        Me.Label3.TabIndex = 59
        Me.Label3.Text = "ASORTYMENT TOWAROWY DO ANALIZY :"
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(427, 355)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(207, 16)
        Me.Label4.TabIndex = 55
        Me.Label4.Text = "Data aktualizacji . . . . . . . . . . . . . . . . . . "
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(671, 129)
        Me.Label1.TabIndex = 54
        Me.Label1.Text = resources.GetString("Label1.Text")
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Location = New System.Drawing.Point(0, 140)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(692, 1)
        Me.Panel3.TabIndex = 53
        '
        'ProgressBar
        '
        Me.ProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar.Location = New System.Drawing.Point(11, 399)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(668, 8)
        Me.ProgressBar.TabIndex = 49
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(427, 372)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(234, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Iloœæ rekordów przeczytanych . . . . . . . . . . . . . . . . . . "
        '
        'cmdImportuj
        '
        Me.cmdImportuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdImportuj.Image = CType(resources.GetObject("cmdImportuj.Image"), System.Drawing.Image)
        Me.cmdImportuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdImportuj.Location = New System.Drawing.Point(534, 442)
        Me.cmdImportuj.Name = "cmdImportuj"
        Me.cmdImportuj.Size = New System.Drawing.Size(75, 22)
        Me.cmdImportuj.TabIndex = 35
        Me.cmdImportuj.Text = "Importuj"
        Me.cmdImportuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(616, 442)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(81, 22)
        Me.cmdPowrot.TabIndex = 34
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(15, 441)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 23)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 33
        Me.PictureBox2.TabStop = False
        '
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'PobierzObrotyzTOFIdoAnalizy2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(711, 470)
        Me.Controls.Add(Me.cmdImportuj)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "PobierzObrotyzTOFIdoAnalizy2"
        Me.ShowInTaskbar = False
        Me.Text = "Pobieranie danych o obrotch do aktualizacji Analizy Zakupów"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    '---------------------------
    Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies

    Dim dbConnectionObrotyMies As OleDb.OleDbConnection
    Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand
    Dim dbCommandDeleteObrotyMies As OleDb.OleDbCommand
    Dim dbCommandUpdateObrotyMies As OleDb.OleDbCommand
    Dim dbCommandInsertObrotyMies As OleDb.OleDbCommand
    Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter

    Dim sqlConnectionObrotyMies As SqlClient.SqlConnection
    Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandDeleteObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandUpdateObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandInsertObrotyMies As SqlClient.SqlCommand
    Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter

    Dim DataSetObrotyMies As New DataSet
    Dim DataViewObrotyMies As New DataView
    '---------------------------
    Dim dbSelectSQLObroty As String = sSelectSQLObroty

    Dim dbConnectionObroty As OleDb.OleDbConnection
    Dim dbCommandSelectObroty As OleDb.OleDbCommand
    Dim dbCommandDeleteObroty As OleDb.OleDbCommand
    Dim dbCommandUpdateObroty As OleDb.OleDbCommand
    Dim dbCommandInsertObroty As OleDb.OleDbCommand
    Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter

    Dim sqlConnectionObroty As SqlClient.SqlConnection
    Dim sqlCommandSelectObroty As SqlClient.SqlCommand
    Dim sqlCommandDeleteObroty As SqlClient.SqlCommand
    Dim sqlCommandUpdateObroty As SqlClient.SqlCommand
    Dim sqlCommandInsertObroty As SqlClient.SqlCommand
    Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter

    Dim DataSetObroty As New DataSet
    Dim DataViewObroty As New DataView
    '---------------------------
    Dim dbSelectSQLKlienci As String = ""        ' sSelectSQLKlienci & " ORDER BY IDENTKLIENTA"

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
    '---------------------------
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
    '---------------------------
    Dim ConnectionState As System.Data.ConnectionState






    '-------------------------------------
    Dim IlePrzeczytano As Long
    Dim IleJest As Long

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        'zapamietaj warunki wyboru do analizy
        Par_OstOkresAnalizy = Microsoft.VisualBasic.Right("0000" & Trim(cbxRok1.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies1.Text), 2) &
                              Microsoft.VisualBasic.Right("0000" & Trim(cbxRok2.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies2.Text), 2)

        Par_DataAktAnalizy = SysData()
        Par_DaneDoAnalizy = IIf(chkPryzmat.Checked, "T", IIf(chkPryzmatAtrament.Checked, "A", IIf(chkPryzmatTonery.Checked, "L", "N"))) &
                            IIf(chkOryg.Checked, "T", IIf(chkOrygAtrament.Checked, "A", IIf(chkOrygTonery.Checked, "L", "N"))) &
                            IIf(chkNajem.Checked, "T", "N") &
                            IIf(chkStrony.Checked, "T", "N") &
                            IIf(chkDrukarki.Checked, "T", "N") &
                            IIf(chkSkup.Checked, "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj procent Mar¿y", "T", "N")


        Par_MiesAnalizy1 = IIf(Chk101.Checked, "T", "N") &
                           IIf(Chk102.Checked, "T", "N") &
                           IIf(Chk103.Checked, "T", "N") &
                           IIf(Chk104.Checked, "T", "N") &
                           IIf(Chk105.Checked, "T", "N") &
                           IIf(Chk106.Checked, "T", "N") &
                           IIf(Chk107.Checked, "T", "N") &
                           IIf(Chk108.Checked, "T", "N") &
                           IIf(Chk109.Checked, "T", "N") &
                           IIf(Chk110.Checked, "T", "N") &
                           IIf(Chk111.Checked, "T", "N") &
                           IIf(Chk112.Checked, "T", "N")

        Par_MiesAnalizy2 = IIf(Chk201.Checked, "T", "N") &
                           IIf(Chk202.Checked, "T", "N") &
                           IIf(Chk203.Checked, "T", "N") &
                           IIf(Chk204.Checked, "T", "N") &
                           IIf(Chk205.Checked, "T", "N") &
                           IIf(Chk206.Checked, "T", "N") &
                           IIf(Chk207.Checked, "T", "N") &
                           IIf(Chk208.Checked, "T", "N") &
                           IIf(Chk209.Checked, "T", "N") &
                           IIf(Chk210.Checked, "T", "N") &
                           IIf(Chk211.Checked, "T", "N") &
                           IIf(Chk212.Checked, "T", "N")

        ZapiszParametry(Me)
        ''-------------------------------
        Me.Close()
    End Sub



    Private Sub PobierzObrotyzTOFIdoAnalizy2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        InicjujListeLat(cbxRok1)
        InicjujListeMiesiecy(cbxMies1)
        InicjujListeLat(cbxRok2)
        InicjujListeMiesiecy(cbxMies2)
        '-------------------------
        'Par_OstOkresAnalizy = RokMies1 & _   1,7zn  RRRRMM
        '                      RokMies2       8,7zn  RRRRMM

        'Par_DataAktAnalizy = SysData()
        'Par_DaneDoAnalizy = IIf(chkToneryOryg.Checked, "T", "N") & _
        '                    IIf(chkToneryPryzmat.Checked, "T", "N") & _
        '                    IIf(chkAtramentOryg.Checked, "T", "N") & _
        '                    IIf(chkAtramentPryzmat.Checked, "T", "N") & _
        '                    IIf(chkIlosci.Checked, "T", "N") & _
        '                    IIf(chkWartosci.Checked, "T", "N") & _
        'ZapiszParametry(Me)
        '-------------------------------
        PobierzParametry()
        lblDataAktualizacji.Text = Par_DataAktAnalizy       'SysData()

        chkPryzmat.Checked = (Mid(Par_DaneDoAnalizy, 1, 1) = "T")
        chkPryzmatAtrament.Checked = (Mid(Par_DaneDoAnalizy, 1, 1) = "A")
        chkPryzmatTonery.Checked = (Mid(Par_DaneDoAnalizy, 1, 1) = "L")

        chkOryg.Checked = (Mid(Par_DaneDoAnalizy, 2, 1) = "T")
        chkOrygAtrament.Checked = (Mid(Par_DaneDoAnalizy, 2, 1) = "A")
        chkOrygTonery.Checked = (Mid(Par_DaneDoAnalizy, 2, 1) = "L")

        chkNajem.Checked = (Mid(Par_DaneDoAnalizy, 3, 1) = "T")
        chkStrony.Checked = (Mid(Par_DaneDoAnalizy, 4, 1) = "T")
        chkDrukarki.Checked = (Mid(Par_DaneDoAnalizy, 5, 1) = "T")
        chkSkup.Checked = (Mid(Par_DaneDoAnalizy, 6, 1) = "T")

        Dim Okres As String = ""
        'I okres analizy
        If Len(Trim(Mid(Par_OstOkresAnalizy, 1, 7))) = 0 Then
            Okres = Mid(SysData, 1, 7)
        Else
            Okres = Mid(Par_OstOkresAnalizy, 1, 7)
        End If
        cbxRok1.Text = Mid(Okres, 1, 4)
        cbxMies1.Text = Mid(Okres, 6, 2)

        'II okres analizy
        If Len(Trim(Mid(Par_OstOkresAnalizy, 8, 7))) = 0 Then
            Okres = MiesDoAnalizy(Okres, -12)
        Else
            Okres = Mid(Par_OstOkresAnalizy, 8, 7)
        End If
        cbxRok2.Text = Mid(Okres, 1, 4)
        cbxMies2.Text = Mid(Okres, 6, 2)


        Chk101.Checked = (Mid(Par_MiesAnalizy1, 1, 1) = "T")
        Chk102.Checked = (Mid(Par_MiesAnalizy1, 2, 1) = "T")
        Chk103.Checked = (Mid(Par_MiesAnalizy1, 3, 1) = "T")
        Chk104.Checked = (Mid(Par_MiesAnalizy1, 4, 1) = "T")
        Chk105.Checked = (Mid(Par_MiesAnalizy1, 5, 1) = "T")
        Chk106.Checked = (Mid(Par_MiesAnalizy1, 6, 1) = "T")
        Chk107.Checked = (Mid(Par_MiesAnalizy1, 7, 1) = "T")
        Chk108.Checked = (Mid(Par_MiesAnalizy1, 8, 1) = "T")
        Chk109.Checked = (Mid(Par_MiesAnalizy1, 9, 1) = "T")
        Chk110.Checked = (Mid(Par_MiesAnalizy1, 10, 1) = "T")
        Chk111.Checked = (Mid(Par_MiesAnalizy1, 11, 1) = "T")
        Chk112.Checked = (Mid(Par_MiesAnalizy1, 12, 1) = "T")

        Chk201.Checked = (Mid(Par_MiesAnalizy2, 1, 1) = "T")
        Chk202.Checked = (Mid(Par_MiesAnalizy2, 2, 1) = "T")
        Chk203.Checked = (Mid(Par_MiesAnalizy2, 3, 1) = "T")
        Chk204.Checked = (Mid(Par_MiesAnalizy2, 4, 1) = "T")
        Chk205.Checked = (Mid(Par_MiesAnalizy2, 5, 1) = "T")
        Chk206.Checked = (Mid(Par_MiesAnalizy2, 6, 1) = "T")
        Chk207.Checked = (Mid(Par_MiesAnalizy2, 7, 1) = "T")
        Chk208.Checked = (Mid(Par_MiesAnalizy2, 8, 1) = "T")
        Chk209.Checked = (Mid(Par_MiesAnalizy2, 9, 1) = "T")
        Chk210.Checked = (Mid(Par_MiesAnalizy2, 10, 1) = "T")
        Chk211.Checked = (Mid(Par_MiesAnalizy2, 11, 1) = "T")
        Chk212.Checked = (Mid(Par_MiesAnalizy2, 12, 1) = "T")

        cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
        cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
        cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
        cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
        cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
        cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")

        'Par_DaneDoAnalizy = Mid(Par_DaneDoAnalizy, 1, 6) &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj procent Mar¿y", "T", "N")

        If Mid(Par_DaneDoAnalizy, 7, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne"
        If Mid(Par_DaneDoAnalizy, 8, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne"
        If Mid(Par_DaneDoAnalizy, 9, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci"
        If Mid(Par_DaneDoAnalizy, 10, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci"
        If Mid(Par_DaneDoAnalizy, 11, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y"
        If Mid(Par_DaneDoAnalizy, 12, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj procent Mar¿y"

        '-------------------------------
        IlePrzeczytano = 0
        IleJest = 0

        lblIlePrzeczytano.Text = IlePrzeczytano.ToString

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

            DBDeleteAnalizyZakupu(dbCommandDeleteAnalizyZakupu, dbDataAdapterAnalizyZakupu)
            DBUpdateAnalizyZakupu(dbCommandUpdateAnalizyZakupu, dbDataAdapterAnalizyZakupu)
            DBInsertAnalizyZakupu(dbCommandInsertAnalizyZakupu, dbDataAdapterAnalizyZakupu)

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

            SQLDeleteAnalizyZakupu(sqlCommandDeleteAnalizyZakupu, sqlDataAdapterAnalizyZakupu)
            SQLUpdateAnalizyZakupu(sqlCommandUpdateAnalizyZakupu, sqlDataAdapterAnalizyZakupu)
            SQLInsertAnalizyZakupu(sqlCommandInsertAnalizyZakupu, sqlDataAdapterAnalizyZakupu)

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

                DataSetAnalizyZakupu.Tables(0).PrimaryKey = New DataColumn() {DataSetAnalizyZakupu.Tables(0).Columns("IDENTKLIENTA")}
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


    Private Sub PobierzObrotyzTOFIdoAnalizy2_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************


    Private Sub AktualizujBazeAnalizyZakupu()
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
        Dim SymbolKlienta As String = ""
        Dim IleRecAnalizy As Long = 0
        Dim IleRecKlienci As Long = 0
        Dim IleRecObroty As Long = 0

        IlePrzeczytano = 0
        IleJest = 0
        lblIlePrzeczytano.Text = IlePrzeczytano.ToString





        '-----------------------------
        'usuwamy zapisy z analizy
        '-----------------------------
        Dim drAnalizy As DataRow
        Dim IrecAnalizy As Integer
        IleRecAnalizy = DataSetAnalizyZakupu.Tables(0).Rows.Count
        ProgressBar.Maximum = IleRecAnalizy
        For IrecAnalizy = IleRecAnalizy - 1 To 0 Step -1
            ProgressBar.Value = IleRecAnalizy - IrecAnalizy
            drAnalizy = DataSetAnalizyZakupu.Tables(0).Rows.Item(IrecAnalizy)
            drAnalizy.Delete()
            System.Windows.Forms.Application.DoEvents()
        Next
        AktualizujBazeAnalizyZakupu()



        '-----------------------------
        'analizuj klientow
        '-----------------------------

        'DataSet
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelectKlienci = New OleDb.OleDbCommand(dbSelectSQLKlienci, dbConnectionKlienci)
            ''dbCommandDeleteKlienci = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnectionKlienci)
            ''dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
            ''dbCommandInsertKlienci = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnectionKlienci)
            'dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

            ''DataSet
            'DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            'MapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterKlienci.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionKlienci.State
            'End Try

        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienciAktywni & " ORDER BY IDENTKLIENTA", sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            'sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            MapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienci.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
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

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If






        '-----------------------------
        'analizuj Obroty Mies
        '-----------------------------

        'DataSet
        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionObrotyMies)
            'dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
            ''dbCommandDeleteObrotyMies = New OleDb.OleDbCommand(sDeleteSQLObrotyMies, dbConnectionObrotyMies)
            ''dbCommandUpdateObrotyMies = New OleDb.OleDbCommand(sUpdateSQLObrotyMies, dbConnectionObrotyMies)
            ''dbCommandInsertObrotyMies = New OleDb.OleDbCommand(sInsertSQLObrotyMies, dbConnectionObrotyMies)
            'dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

            ''DataSet
            'DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            'MapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterObrotyMies.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionObrotyMies.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionObrotyMies.State
            'End Try

        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionObrotyMies)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandDeleteObrotyMies = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandUpdateObrotyMies = New SqlClient.SqlCommand(sUpdateSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandInsertObrotyMies = New SqlClient.SqlCommand(sInsertSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            MapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterObrotyMies.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionObrotyMies.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    dbConnectionObrotyMies.Close()
                Else
                    sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionObrotyMies.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"), DataSetObrotyMies.Tables(0).Columns("MIESIAC")}
                DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If






        '-----------------------------
        'analizuj Obroty w ostatnim miesiacu
        '-----------------------------

        'DataSet
        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionObroty = New OleDb.OleDbConnection(sConnectionObroty)
            'dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
            ''dbCommandDeleteObroty = New OleDb.OleDbCommand(sDeleteSQLObroty, dbConnectionObroty)
            ''dbCommandUpdateObroty = New OleDb.OleDbCommand(sUpdateSQLObroty, dbConnectionObroty)
            ''dbCommandInsertObroty = New OleDb.OleDbCommand(sInsertSQLObroty, dbConnectionObroty)
            'dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

            ''DataSet
            'DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
            'MapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterObroty.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionObroty.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionObroty.State
            'End Try

        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionObroty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
            'sqlCommandDeleteObroty = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnectionObroty)
            'sqlCommandUpdateObroty = New SqlClient.SqlCommand(sUpdateSQLObroty, sqlConnectionObroty)
            'sqlCommandInsertObroty = New SqlClient.SqlCommand(sInsertSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
            MapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterObroty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionObroty.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    dbDataAdapterObroty.Fill(DataSetObroty)
                    dbConnectionObroty.Close()
                Else
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetObroty.Tables(0).PrimaryKey = New DataColumn() {DataSetObroty.Tables(0).Columns("IDENTKLIENTA"), DataSetObroty.Tables(0).Columns("DATA")}
                DataViewObroty = New DataView(DataSetObroty.Tables(0))

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If











        ' analizuj poszczegolnych klientow
        ' a dla klientow analizuj obroty miesieczne
        ' dopisuj do bazy AnalizyZakupu
        Dim drKlienci As DataRow
        Dim nrAnalizy As DataRow
        Dim IrecKlienci As Integer = 0
        Dim SymbKlienta As String = ""

        Dim ILBM As Integer = 0
        Dim IL01 As Integer = 0
        Dim IL02 As Integer = 0
        Dim IL03 As Integer = 0
        Dim IL04 As Integer = 0
        Dim IL05 As Integer = 0
        Dim IL06 As Integer = 0
        Dim IL07 As Integer = 0
        Dim IL08 As Integer = 0
        Dim IL09 As Integer = 0
        Dim IL10 As Integer = 0
        Dim IL11 As Integer = 0

        Dim IL12 As Integer = 0
        Dim IL13 As Integer = 0
        Dim IL14 As Integer = 0
        Dim IL15 As Integer = 0
        Dim IL16 As Integer = 0
        Dim IL17 As Integer = 0
        Dim IL18 As Integer = 0
        Dim IL19 As Integer = 0
        Dim IL20 As Integer = 0
        Dim IL21 As Integer = 0
        Dim IL22 As Integer = 0
        Dim IL23 As Integer = 0



        Dim WABM As Double = 0
        Dim WA01 As Double = 0
        Dim WA02 As Double = 0
        Dim WA03 As Double = 0
        Dim WA04 As Double = 0
        Dim WA05 As Double = 0
        Dim WA06 As Double = 0
        Dim WA07 As Double = 0
        Dim WA08 As Double = 0
        Dim WA09 As Double = 0
        Dim WA10 As Double = 0
        Dim WA11 As Double = 0

        Dim WA12 As Double = 0
        Dim WA13 As Double = 0
        Dim WA14 As Double = 0
        Dim WA15 As Double = 0
        Dim WA16 As Double = 0
        Dim WA17 As Double = 0
        Dim WA18 As Double = 0
        Dim WA19 As Double = 0
        Dim WA20 As Double = 0
        Dim WA21 As Double = 0
        Dim WA22 As Double = 0
        Dim WA23 As Double = 0



        Dim MARBM As Double = 0
        Dim MAR01 As Double = 0
        Dim MAR02 As Double = 0
        Dim MAR03 As Double = 0
        Dim MAR04 As Double = 0
        Dim MAR05 As Double = 0
        Dim MAR06 As Double = 0
        Dim MAR07 As Double = 0
        Dim MAR08 As Double = 0
        Dim MAR09 As Double = 0
        Dim MAR10 As Double = 0
        Dim MAR11 As Double = 0

        Dim MAR12 As Double = 0
        Dim MAR13 As Double = 0
        Dim MAR14 As Double = 0
        Dim MAR15 As Double = 0
        Dim MAR16 As Double = 0
        Dim MAR17 As Double = 0
        Dim MAR18 As Double = 0
        Dim MAR19 As Double = 0
        Dim MAR20 As Double = 0
        Dim MAR21 As Double = 0
        Dim MAR22 As Double = 0
        Dim MAR23 As Double = 0



        Dim PROBM As Double = 0
        Dim PRO01 As Double = 0
        Dim PRO02 As Double = 0
        Dim PRO03 As Double = 0
        Dim PRO04 As Double = 0
        Dim PRO05 As Double = 0
        Dim PRO06 As Double = 0
        Dim PRO07 As Double = 0
        Dim PRO08 As Double = 0
        Dim PRO09 As Double = 0
        Dim PRO10 As Double = 0
        Dim PRO11 As Double = 0

        Dim PRO12 As Double = 0
        Dim PRO13 As Double = 0
        Dim PRO14 As Double = 0
        Dim PRO15 As Double = 0
        Dim PRO16 As Double = 0
        Dim PRO17 As Double = 0
        Dim PRO18 As Double = 0
        Dim PRO19 As Double = 0
        Dim PRO20 As Double = 0
        Dim PRO21 As Double = 0
        Dim PRO22 As Double = 0
        Dim PRO23 As Double = 0

        Dim IL24 As Integer = 0

        Dim OkresAn As String = ""

        IleRecKlienci = DataSetKlienci.Tables(0).Rows.Count
        ProgressBar.Maximum = IleRecKlienci
        For IrecKlienci = 0 To IleRecKlienci - 1
            ProgressBar.Value = IrecKlienci
            drKlienci = DataSetKlienci.Tables(0).Rows.Item(IrecKlienci)
            SymbKlienta = DataSetKlienci.Tables(0).Rows.Item(IrecKlienci).Item("NRKARTY")
            System.Windows.Forms.Application.DoEvents()

            'dopisz tego klienta do bazy Analizy
            nrAnalizy = DataSetAnalizyZakupu.Tables(0).NewRow()

            nrAnalizy.Item("IDENTKLIENTA") = SymbKlienta
            nrAnalizy.Item("WARTZA00") = 0
            nrAnalizy.Item("MARZA00") = 0
            nrAnalizy.Item("PROCM00") = 0
            nrAnalizy.Item("ILOSZA00") = 0

            '--------------------------------------
            'pierwszy okres do analizy,,,
            OkresAn = Microsoft.VisualBasic.Right("0000" & Trim(cbxRok1.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies1.Text), 2)

            If Not Chk101.Checked Then
                ILBM = 0
                WABM = 0
                MARBM = 0
                PROBM = 0
            Else
                ILBM = AnalizujObrotyIlo(SymbKlienta, OkresAn)
                WABM = AnalizujObrotyWar(SymbKlienta, OkresAn)
                MARBM = AnalizujObrotyMar(SymbKlienta, OkresAn)
                PROBM = IIf(WABM = 0, 0, MARBM / WABM * 100)         'AnalizujObrotyPrM(SymbKlienta, OkresAn)
            End If
            nrAnalizy.Item("WARTZABM") = WABM
            nrAnalizy.Item("MARZABM") = MARBM
            nrAnalizy.Item("PROCMBM") = PROBM
            nrAnalizy.Item("ILOSZABM") = ILBM

            'analizuj poprzednie miesiace...
            If Not Chk102.Checked Then
                IL01 = 0
                WA01 = 0
                MAR01 = 0
                PRO01 = 0
            Else
                IL01 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
                WA01 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
                MAR01 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
                PRO01 = IIf(WA01 = 0, 0, MAR01 / WA01 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
            End If
            nrAnalizy.Item("WARTZA01") = WA01
            nrAnalizy.Item("MARZA01") = MAR01
            nrAnalizy.Item("PROCM01") = PRO01
            nrAnalizy.Item("ILOSZA01") = IL01

            If Not Chk103.Checked Then
                IL02 = 0
                WA02 = 0
                MAR02 = 0
                PRO02 = 0
            Else
                IL02 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
                WA02 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
                MAR02 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
                PRO02 = IIf(WA02 = 0, 0, MAR02 / WA02 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
            End If
            nrAnalizy.Item("WARTZA02") = WA02
            nrAnalizy.Item("MARZA02") = MAR02
            nrAnalizy.Item("PROCM02") = PRO02
            nrAnalizy.Item("ILOSZA02") = IL02

            If Not Chk104.Checked Then
                IL03 = 0
                WA03 = 0
                MAR03 = 0
                PRO03 = 0
            Else
                IL03 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
                WA03 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
                MAR03 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
                PRO03 = IIf(WA03 = 0, 0, MAR03 / WA03 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
            End If
            nrAnalizy.Item("WARTZA03") = WA03
            nrAnalizy.Item("MARZA03") = MAR03
            nrAnalizy.Item("PROCM03") = PRO03
            nrAnalizy.Item("ILOSZA03") = IL03

            If Not Chk105.Checked Then
                IL04 = 0
                WA04 = 0
                MAR04 = 0
                PRO04 = 0
            Else
                IL04 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
                WA04 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
                MAR04 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
                PRO05 = IIf(WA04 = 0, 0, MAR04 / WA04 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
            End If
            nrAnalizy.Item("WARTZA04") = WA04
            nrAnalizy.Item("MARZA04") = MAR04
            nrAnalizy.Item("PROCM04") = PRO04
            nrAnalizy.Item("ILOSZA04") = IL04

            If Not Chk106.Checked Then
                IL05 = 0
                WA05 = 0
                MAR05 = 0
                PRO05 = 0
            Else
                IL05 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
                WA05 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
                MAR05 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
                PRO05 = IIf(WA05 = 0, 0, MAR05 / WA05 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
            End If
            nrAnalizy.Item("WARTZA05") = WA05
            nrAnalizy.Item("MARZA05") = MAR05
            nrAnalizy.Item("PROCM05") = PRO05
            nrAnalizy.Item("ILOSZA05") = IL05

            If Not Chk107.Checked Then
                IL06 = 0
                WA06 = 0
                MAR06 = 0
                PRO06 = 0
            Else
                IL06 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
                WA06 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
                MAR06 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
                PRO06 = IIf(WA06 = 0, 0, MAR06 / WA06 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
            End If
            nrAnalizy.Item("WARTZA06") = WA06
            nrAnalizy.Item("MARZA06") = MAR06
            nrAnalizy.Item("PROCM06") = PRO06
            nrAnalizy.Item("ILOSZA06") = IL06

            If Not Chk108.Checked Then
                IL07 = 0
                WA07 = 0
                MAR07 = 0
                PRO07 = 0
            Else
                IL07 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
                WA07 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
                MAR07 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
                PRO07 = IIf(WA07 = 0, 0, MAR07 / WA07 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
            End If
            nrAnalizy.Item("WARTZA07") = WA07
            nrAnalizy.Item("MARZA07") = MAR07
            nrAnalizy.Item("PROCM07") = PRO07
            nrAnalizy.Item("ILOSZA07") = IL07

            If Not Chk109.Checked Then
                IL08 = 0
                WA08 = 0
                MAR08 = 0
                PRO08 = 0
            Else
                IL08 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
                WA08 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
                MAR08 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
                PRO08 = IIf(WA08 = 0, 0, MAR08 / WA08 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
            End If
            nrAnalizy.Item("WARTZA08") = WA08
            nrAnalizy.Item("MARZA08") = MAR08
            nrAnalizy.Item("PROCM08") = PRO08
            nrAnalizy.Item("ILOSZA08") = IL08

            If Not Chk110.Checked Then
                IL09 = 0
                WA09 = 0
                MAR09 = 0
                PRO09 = 0
            Else
                IL09 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
                WA09 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
                MAR09 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
                PRO09 = IIf(WA09 = 0, 0, MAR09 / WA09 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
            End If
            nrAnalizy.Item("WARTZA09") = WA09
            nrAnalizy.Item("MARZA09") = MAR09
            nrAnalizy.Item("PROCM09") = PRO09
            nrAnalizy.Item("ILOSZA09") = IL09

            If Not Chk111.Checked Then
                IL10 = 0
                WA10 = 0
                MAR10 = 0
                PRO10 = 0
            Else
                IL10 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
                WA10 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
                MAR10 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
                PRO10 = IIf(WA10 = 0, 0, MAR10 / WA10 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
            End If
            nrAnalizy.Item("WARTZA10") = WA10
            nrAnalizy.Item("MARZA10") = MAR10
            nrAnalizy.Item("PROCM10") = PRO10
            nrAnalizy.Item("ILOSZA10") = IL10

            If Not Chk112.Checked Then
                IL11 = 0
                WA11 = 0
                MAR11 = 0
                PRO11 = 0
            Else
                IL11 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
                WA11 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
                MAR11 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
                PRO11 = IIf(WA11 = 0, 0, MAR11 / WA11 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
            End If
            nrAnalizy.Item("WARTZA11") = WA11
            nrAnalizy.Item("MARZA11") = MAR11
            nrAnalizy.Item("PROCM11") = PRO11
            nrAnalizy.Item("ILOSZA11") = IL11



            '--------------------------------------
            'drugi okres do analizy,,,
            OkresAn = Microsoft.VisualBasic.Right("0000" & Trim(cbxRok2.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies2.Text), 2)

            If Not Chk201.Checked Then
                IL12 = 0
                WA12 = 0
                MAR12 = 0
                PRO12 = 0
            Else
                IL12 = AnalizujObrotyIlo(SymbKlienta, OkresAn)
                WA12 = AnalizujObrotyWar(SymbKlienta, OkresAn)
                MAR12 = AnalizujObrotyMar(SymbKlienta, OkresAn)
                PRO12 = IIf(WA12 = 0, 0, MAR12 / WA12 * 100)         'AnalizujObrotyPrM(SymbKlienta, OkresAn)
            End If
            nrAnalizy.Item("WARTZA12") = WA12
            nrAnalizy.Item("ILOSZA12") = IL12
            nrAnalizy.Item("MARZA12") = MAR12
            nrAnalizy.Item("PROCM12") = PRO12

            If Not Chk202.Checked Then
                IL13 = 0
                WA13 = 0
                MAR13 = 0
                PRO13 = 0
            Else
                IL13 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
                WA13 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
                MAR13 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
                PRO13 = IIf(WA13 = 0, 0, MAR13 / WA13 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
            End If
            nrAnalizy.Item("WARTZA13") = WA13
            nrAnalizy.Item("MARZA13") = MAR13
            nrAnalizy.Item("PROCM13") = PRO13
            nrAnalizy.Item("ILOSZA13") = IL13

            If Not Chk203.Checked Then
                IL14 = 0
                WA14 = 0
                MAR14 = 0
                PRO14 = 0
            Else
                IL14 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
                WA14 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
                MAR14 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
                PRO14 = IIf(WA14 = 0, 0, MAR14 / WA14 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
            End If
            nrAnalizy.Item("WARTZA14") = WA14
            nrAnalizy.Item("MARZA14") = MAR14
            nrAnalizy.Item("PROCM14") = PRO14
            nrAnalizy.Item("ILOSZA14") = IL14

            If Not Chk204.Checked Then
                IL15 = 0
                WA15 = 0
                MAR15 = 0
                PRO15 = 0
            Else
                IL15 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
                WA15 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
                MAR15 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
                PRO15 = IIf(WA15 = 0, 0, MAR15 / WA15 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
            End If
            nrAnalizy.Item("WARTZA15") = WA15
            nrAnalizy.Item("MARZA15") = MAR15
            nrAnalizy.Item("PROCM15") = PRO15
            nrAnalizy.Item("ILOSZA15") = IL15

            If Not Chk205.Checked Then
                IL16 = 0
                WA16 = 0
                MAR16 = 0
                PRO16 = 0
            Else
                IL16 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
                WA16 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
                MAR16 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
                PRO16 = IIf(WA16 = 0, 0, MAR16 / WA16 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
            End If
            nrAnalizy.Item("WARTZA16") = WA16
            nrAnalizy.Item("MARZA16") = MAR16
            nrAnalizy.Item("PROCM16") = PRO16
            nrAnalizy.Item("ILOSZA16") = IL16

            If Not Chk206.Checked Then
                IL17 = 0
                WA17 = 0
                MAR17 = 0
                PRO17 = 0
            Else
                IL17 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
                WA17 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
                MAR17 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
                PRO17 = IIf(WA17 = 0, 0, MAR17 / WA17 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
            End If
            nrAnalizy.Item("WARTZA17") = WA17
            nrAnalizy.Item("MARZA17") = MAR17
            nrAnalizy.Item("PROCM17") = PRO17
            nrAnalizy.Item("ILOSZA17") = IL17

            If Not Chk207.Checked Then
                IL18 = 0
                WA18 = 0
                MAR18 = 0
                PRO18 = 0
            Else
                IL18 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
                WA18 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
                MAR18 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
                PRO18 = IIf(WA18 = 0, 0, MAR18 / WA18 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
            End If
            nrAnalizy.Item("WARTZA18") = WA18
            nrAnalizy.Item("MARZA18") = MAR18
            nrAnalizy.Item("PROCM18") = PRO18
            nrAnalizy.Item("ILOSZA18") = IL18

            If Not Chk208.Checked Then
                IL19 = 0
                WA19 = 0
                MAR19 = 0
                PRO19 = 0
            Else
                IL19 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
                WA19 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
                MAR19 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
                PRO19 = IIf(WA19 = 0, 0, MAR19 / WA19 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
            End If
            nrAnalizy.Item("WARTZA19") = WA19
            nrAnalizy.Item("MARZA19") = MAR19
            nrAnalizy.Item("PROCM19") = PRO19
            nrAnalizy.Item("ILOSZA19") = IL19

            If Not Chk209.Checked Then
                IL20 = 0
                WA20 = 0
                MAR20 = 0
                PRO20 = 0
            Else
                IL20 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
                WA20 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
                MAR20 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
                PRO20 = IIf(WA20 = 0, 0, MAR20 / WA20 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
            End If
            nrAnalizy.Item("WARTZA20") = WA20
            nrAnalizy.Item("MARZA20") = MAR20
            nrAnalizy.Item("PROCM20") = PRO20
            nrAnalizy.Item("ILOSZA20") = IL20

            If Not Chk210.Checked Then
                IL21 = 0
                WA21 = 0
                MAR21 = 0
                PRO21 = 0
            Else
                IL21 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
                WA21 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
                MAR21 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
                PRO21 = IIf(WA21 = 0, 0, MAR21 / WA21 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
            End If
            nrAnalizy.Item("WARTZA21") = WA21
            nrAnalizy.Item("MARZA21") = MAR21
            nrAnalizy.Item("PROCM21") = PRO21
            nrAnalizy.Item("ILOSZA21") = IL21

            If Not Chk211.Checked Then
                IL22 = 0
                WA22 = 0
                MAR22 = 0
                PRO22 = 0
            Else
                IL22 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
                WA22 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
                MAR22 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
                PRO22 = IIf(WA22 = 0, 0, MAR22 / WA22 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
            End If
            nrAnalizy.Item("WARTZA22") = WA22
            nrAnalizy.Item("MARZA22") = MAR22
            nrAnalizy.Item("PROCM22") = PRO22
            nrAnalizy.Item("ILOSZA22") = IL22

            If Not Chk212.Checked Then
                IL23 = 0
                WA23 = 0
                MAR23 = 0
                PRO23 = 0
            Else
                IL23 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
                WA23 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
                MAR23 = AnalizujObrotyMar(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
                PRO23 = IIf(WA23 = 0, 0, MAR23 / WA23 * 100)         'AnalizujObrotyPrM(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
            End If
            nrAnalizy.Item("WARTZA23") = WA23
            nrAnalizy.Item("MARZA23") = MAR23
            nrAnalizy.Item("PROCM23") = PRO23
            nrAnalizy.Item("ILOSZA23") = IL23

            'IL24 = AnalizujObrotyKlienta(SymbKlienta, MiesDoAnalizy(-24))
            'WA24 = AnalizujWartObrotyKlienta(SymbKlienta, MiesDoAnalizy(-24))
            nrAnalizy.Item("WARTZA24") = 0
            nrAnalizy.Item("MARZA24") = 0
            nrAnalizy.Item("PROCM24") = 0
            nrAnalizy.Item("ILOSZA24") = 0      'IL24

            nrAnalizy.Item("WERSJA") = 1







            'ilosc miesiecy w ktorych byla sprzedaz 0...-11
            nrAnalizy.Item("WARTZA00") = IIf((ILBM <> 0) Or (WABM <> 0), 1, 0) +
                                         IIf((IL01 <> 0) Or (WA01 <> 0), 1, 0) +
                                         IIf((IL02 <> 0) Or (WA02 <> 0), 1, 0) +
                                         IIf((IL03 <> 0) Or (WA03 <> 0), 1, 0) +
                                         IIf((IL04 <> 0) Or (WA04 <> 0), 1, 0) +
                                         IIf((IL05 <> 0) Or (WA05 <> 0), 1, 0) +
                                         IIf((IL06 <> 0) Or (WA06 <> 0), 1, 0) +
                                         IIf((IL07 <> 0) Or (WA07 <> 0), 1, 0) +
                                         IIf((IL08 <> 0) Or (WA08 <> 0), 1, 0) +
                                         IIf((IL09 <> 0) Or (WA09 <> 0), 1, 0) +
                                         IIf((IL10 <> 0) Or (WA10 <> 0), 1, 0) +
                                         IIf((IL11 <> 0) Or (WA11 <> 0), 1, 0)

            'ilosc miesiecy w ktorych byla sprzedaz -12..-23
            nrAnalizy.Item("ILOSZA00") = IIf((IL12 <> 0) Or (WA12 <> 0), 1, 0) +
                                         IIf((IL13 <> 0) Or (WA13 <> 0), 1, 0) +
                                         IIf((IL14 <> 0) Or (WA14 <> 0), 1, 0) +
                                         IIf((IL15 <> 0) Or (WA15 <> 0), 1, 0) +
                                         IIf((IL16 <> 0) Or (WA16 <> 0), 1, 0) +
                                         IIf((IL17 <> 0) Or (WA17 <> 0), 1, 0) +
                                         IIf((IL18 <> 0) Or (WA18 <> 0), 1, 0) +
                                         IIf((IL19 <> 0) Or (WA19 <> 0), 1, 0) +
                                         IIf((IL20 <> 0) Or (WA20 <> 0), 1, 0) +
                                         IIf((IL21 <> 0) Or (WA21 <> 0), 1, 0) +
                                         IIf((IL22 <> 0) Or (WA22 <> 0), 1, 0) +
                                         IIf((IL23 <> 0) Or (WA23 <> 0), 1, 0)





            'dodaj rekord do DataSet
            DataSetAnalizyZakupu.Tables(0).Rows.Add(nrAnalizy)
            AktualizujBazeAnalizyZakupu()
            lblIlePrzeczytano.Text = IrecKlienci.ToString

            System.Windows.Forms.Application.DoEvents()
        Next
        '-------------------------------
        'zapamietaj warunki wyboru do analizy
        Par_OstOkresAnalizy = Microsoft.VisualBasic.Right("0000" & Trim(cbxRok1.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies1.Text), 2) &
                              Microsoft.VisualBasic.Right("0000" & Trim(cbxRok2.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies2.Text), 2)

        Par_DataAktAnalizy = SysData()
        Par_DaneDoAnalizy = IIf(chkPryzmat.Checked, "T", IIf(chkPryzmatAtrament.Checked, "A", IIf(chkPryzmatTonery.Checked, "L", "N"))) &
                            IIf(chkOryg.Checked, "T", IIf(chkOrygAtrament.Checked, "A", IIf(chkOrygTonery.Checked, "L", "N"))) &
                            IIf(chkNajem.Checked, "T", "N") &
                            IIf(chkStrony.Checked, "T", "N") &
                            IIf(chkDrukarki.Checked, "T", "N") &
                            IIf(chkSkup.Checked, "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj procent Mar¿y", "T", "N")


        Par_MiesAnalizy1 = IIf(Chk101.Checked, "T", "N") &
                           IIf(Chk102.Checked, "T", "N") &
                           IIf(Chk103.Checked, "T", "N") &
                           IIf(Chk104.Checked, "T", "N") &
                           IIf(Chk105.Checked, "T", "N") &
                           IIf(Chk106.Checked, "T", "N") &
                           IIf(Chk107.Checked, "T", "N") &
                           IIf(Chk108.Checked, "T", "N") &
                           IIf(Chk109.Checked, "T", "N") &
                           IIf(Chk110.Checked, "T", "N") &
                           IIf(Chk111.Checked, "T", "N") &
                           IIf(Chk112.Checked, "T", "N")

        Par_MiesAnalizy2 = IIf(Chk201.Checked, "T", "N") &
                           IIf(Chk202.Checked, "T", "N") &
                           IIf(Chk203.Checked, "T", "N") &
                           IIf(Chk204.Checked, "T", "N") &
                           IIf(Chk205.Checked, "T", "N") &
                           IIf(Chk206.Checked, "T", "N") &
                           IIf(Chk207.Checked, "T", "N") &
                           IIf(Chk208.Checked, "T", "N") &
                           IIf(Chk209.Checked, "T", "N") &
                           IIf(Chk210.Checked, "T", "N") &
                           IIf(Chk211.Checked, "T", "N") &
                           IIf(Chk212.Checked, "T", "N")

        ZapiszParametry(Me)




        '-------------------------------
        MessageBox.Show("Pobra³em informacje o obrotach" & vbCrLf &
            "klientów w poszczególnych miesi¹cach" & vbCrLf &
            "i przepisa³em do tabeli Analizy Zakupu...",
            "OK, skoñczy³em",
            System.Windows.Forms.MessageBoxButtons.OK,
            MessageBoxIcon.Information,
            MessageBoxDefaultButton.Button1)
        Me.Close()
    End Sub












    Private Function AnalizujObrotyIlo(ByVal pKlient As String,
                                       ByVal pMiesiac As String) As Double
        If (pMiesiac = Mid(SysData, 1, 7)) Then
            Return AnalizujIloObrotyBiezMies(pKlient)
            'Return AnalizujWartObrotyBiezMies(pKlient)
        Else
            'analizuj poprzednie miesiace...
            Return AnalizujIloObrotyKlienta(pKlient, pMiesiac)
            'Return AnalizujWartObrotyKlienta(pKlient, pMiesiac)
        End If

        'Return AnalizujObrotyKlienta(pKlient, pMiesiac) + AnalizujObrotyBiezMies(pKlient)
    End Function


    Private Function AnalizujObrotyWar(ByVal pKlient As String,
                                       ByVal pMiesiac As String) As Double
        If (pMiesiac = Mid(SysData, 1, 7)) Then
            'Return AnalizujObrotyBiezMies(pKlient)
            Return AnalizujWartObrotyBiezMies(pKlient)
        Else
            'analizuj poprzednie miesiace...
            'Return AnalizujObrotyKlienta(pKlient, pMiesiac)
            Return AnalizujWartObrotyKlienta(pKlient, pMiesiac)
        End If

        'Return AnalizujWartObrotyKlienta(pKlient, pMiesiac) + AnalizujWartObrotyBiezMies(pKlient)
    End Function


    Private Function AnalizujObrotyMar(ByVal pKlient As String,
                                       ByVal pMiesiac As String) As Double
        If (pMiesiac = Mid(SysData, 1, 7)) Then
            'Return AnalizujObrotyBiezMies(pKlient)
            Return AnalizujMarzeObrotyBiezMies(pKlient)
        Else
            'analizuj poprzednie miesiace...
            'Return AnalizujObrotyKlienta(pKlient, pMiesiac)
            Return AnalizujMarzeObrotyKlienta(pKlient, pMiesiac)
        End If

        'Return AnalizujWartObrotyKlienta(pKlient, pMiesiac) + AnalizujWartObrotyBiezMies(pKlient)
    End Function


    Private Function AnalizujObrotyPrM(ByVal pKlient As String,
                                       ByVal pMiesiac As String) As Double
        If (pMiesiac = Mid(SysData, 1, 7)) Then
            'Return AnalizujObrotyBiezMies(pKlient)
            Return AnalizujMarzeObrotyBiezMies(pKlient)
        Else
            'analizuj poprzednie miesiace...
            'Return AnalizujObrotyKlienta(pKlient, pMiesiac)
            Return AnalizujMarzeObrotyKlienta(pKlient, pMiesiac)
        End If

        'Return AnalizujWartObrotyKlienta(pKlient, pMiesiac) + AnalizujWartObrotyBiezMies(pKlient)
    End Function











    '************************************
    ' Analizuj obroty klienta w poszczególnych miesiacach...
    '************************************

    Private Function AnalizujIloObrotyKlienta(ByVal pKlient As String,
                                           ByVal pMiesiac As String) As Double
        Dim foundRow As DataRow
        Dim findTheseVals(1) As Object
        Dim ObrMiesieczny As Double = 0

        'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
        'Public oMiesiacMies As String          'MIESIAC          Text,10

        'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
        'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedMies As Double       'AILOSPRZED       Double

        'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
        'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double

        findTheseVals(0) = pKlient
        findTheseVals(1) = pMiesiac
        foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)

        If Not (foundRow Is Nothing) Then
            'jest obrot w tym miesiacu
            ObrMiesieczny = 0
            If chkPryzmat.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LILOSPRZED") + GetDblDRField(foundRow, "AILOSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If
            If chkPryzmatAtrament.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AILOSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If
            If chkPryzmatTonery.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LILOSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If

            If chkOryg.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOILOSPRZED") + GetDblDRField(foundRow, "LOILOSPRZED")   'sprzedaz atranametow i laserów oryg
            End If
            If chkOrygAtrament.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOILOSPRZED")   'sprzedaz atranametow i laserów oryg
            End If
            If chkOrygTonery.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LOILOSPRZED")   'sprzedaz atranametow i laserów oryg
            End If

            If chkNajem.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "NAJEMILOSPRZED")
            End If
            If chkStrony.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "STRONYILOSPRZED")
            End If
            If chkDrukarki.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "DRUKARKIILOSPRZED")
            End If
            If chkSkup.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "SKUPILOSPRZED")
            End If
        End If
        Return (ObrMiesieczny)
    End Function



    Private Function AnalizujWartObrotyKlienta(ByVal pKlient As String,
                                       ByVal pMiesiac As String) As Double
        Dim foundRow As DataRow
        Dim findTheseVals(1) As Object
        Dim ObrMiesieczny As Double = 0

        'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
        'Public oMiesiacMies As String          'MIESIAC          Text,10

        'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
        'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedMies As Double       'AILOSPRZED       Double

        'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
        'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double

        findTheseVals(0) = pKlient
        findTheseVals(1) = pMiesiac
        foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)

        If Not (foundRow Is Nothing) Then
            'jest obrot w tym miesiacu
            ObrMiesieczny = 0
            If chkPryzmat.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LWARTSPRZED") + GetDblDRField(foundRow, "AWARTSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If
            If chkPryzmatAtrament.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AWARTSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If
            If chkPryzmatTonery.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LWARTSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If

            If chkOryg.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOWARTSPRZED") + GetDblDRField(foundRow, "LOWARTSPRZED")   'sprzedaz atranametow i laserów oryg
            End If
            If chkOrygAtrament.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOWARTSPRZED")   'sprzedaz atranametow i laserów oryg
            End If
            If chkOrygTonery.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LOWARTSPRZED")   'sprzedaz atranametow i laserów oryg
            End If

            If chkNajem.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "NAJEMWARTSPRZED")
            End If
            If chkStrony.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "STRONYWARTSPRZED")
            End If
            If chkDrukarki.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "DRUKARKIWARTSPRZED")
            End If
            If chkSkup.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "SKUPWARTSPRZED")
            End If
        End If
        Return (ObrMiesieczny)
    End Function



    Private Function AnalizujMarzeObrotyKlienta(ByVal pKlient As String,
                                       ByVal pMiesiac As String) As Double
        Dim foundRow As DataRow
        Dim findTheseVals(1) As Object
        Dim ObrMiesieczny As Double = 0

        'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
        'Public oMiesiacMies As String          'MIESIAC          Text,10

        'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
        'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedMies As Double       'AILOSPRZED       Double

        'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
        'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double

        findTheseVals(0) = pKlient
        findTheseVals(1) = pMiesiac
        foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)

        If Not (foundRow Is Nothing) Then
            'jest obrot w tym miesiacu
            ObrMiesieczny = 0
            If chkPryzmat.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LMARSPRZED") + GetDblDRField(foundRow, "AMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If
            If chkPryzmatAtrament.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If
            If chkPryzmatTonery.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If

            If chkOryg.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOMARSPRZED") + GetDblDRField(foundRow, "LOMARSPRZED")   'sprzedaz atranametow i laserów oryg
            End If
            If chkOrygAtrament.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOMARSPRZED")   'sprzedaz atranametow i laserów oryg
            End If
            If chkOrygTonery.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LOMARSPRZED")   'sprzedaz atranametow i laserów oryg
            End If

            If chkNajem.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "NAJEMMARSPRZED")
            End If
            If chkStrony.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "STRONYMARSPRZED")
            End If
            If chkDrukarki.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "DRUKARKIMARSPRZED")
            End If
            If chkSkup.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "SKUPMARSPRZED")
            End If
        End If
        Return (ObrMiesieczny)
    End Function




    Private Function AnalizujProcMObrotyKlienta(ByVal pKlient As String,
                                       ByVal pMiesiac As String) As Double
        Dim foundRow As DataRow
        Dim findTheseVals(1) As Object
        Dim ObrMiesieczny As Double = 0

        'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
        'Public oMiesiacMies As String          'MIESIAC          Text,10

        'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
        'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedMies As Double       'AILOSPRZED       Double

        'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
        'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double

        findTheseVals(0) = pKlient
        findTheseVals(1) = pMiesiac
        foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)

        If Not (foundRow Is Nothing) Then
            'jest obrot w tym miesiacu
            ObrMiesieczny = 0
            If chkPryzmat.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LMARSPRZED") + GetDblDRField(foundRow, "AMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If
            If chkPryzmatAtrament.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If
            If chkPryzmatTonery.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
            End If

            If chkOryg.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOMARSPRZED") + GetDblDRField(foundRow, "LOMARSPRZED")   'sprzedaz atranametow i laserów oryg
            End If
            If chkOrygAtrament.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOMARSPRZED")   'sprzedaz atranametow i laserów oryg
            End If
            If chkOrygTonery.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "LOMARSPRZED")   'sprzedaz atranametow i laserów oryg
            End If

            If chkNajem.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "NAJEMMARSPRZED")
            End If
            If chkStrony.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "STRONYMARSPRZED")
            End If
            If chkDrukarki.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "DRUKARKIMARSPRZED")
            End If
            If chkSkup.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "SKUPMARSPRZED")
            End If
        End If
        Return (ObrMiesieczny)
    End Function








    Private Function AnalizujIloObrotyBiezMies(ByVal pKlient As String) As Double
        Dim drv As DataRowView
        Dim i As Integer = 0
        Dim Obrot As Double = 0

        'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
        'Public oDataObr As String             'DATA             Text,10

        'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedObr As Double       'LILOPRZED        Double
        'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedObr As Double       'AILOSPRZED       Double

        'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
        'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double


        DataViewObroty.RowFilter = "IDENTKLIENTA='" & pKlient & "'"

        If DataViewObroty.Count > 0 Then
            For i = 0 To DataViewObroty.Count - 1
                drv = DataViewObroty.Item(i)

                If chkPryzmat.Checked Then
                    Obrot += GetDblDRVField(drv, "LILOSPRZED") + GetDblDRVField(drv, "AILOSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If
                If chkPryzmatAtrament.Checked Then
                    Obrot += GetDblDRVField(drv, "AILOSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If
                If chkPryzmatTonery.Checked Then
                    Obrot += GetDblDRVField(drv, "LILOSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If

                If chkOryg.Checked Then
                    Obrot += GetDblDRVField(drv, "AOILOSPRZED") + GetDblDRVField(drv, "LOILOSPRZED")   'sprzedaz atranametow i laserów oryg
                End If
                If chkOrygAtrament.Checked Then
                    Obrot += GetDblDRVField(drv, "AOILOSPRZED")   'sprzedaz atranametow i laserów oryg
                End If
                If chkOrygTonery.Checked Then
                    Obrot += GetDblDRVField(drv, "LOILOSPRZED")   'sprzedaz atranametow i laserów oryg
                End If

                If chkNajem.Checked Then
                    Obrot += GetDblDRVField(drv, "NAJEMILOSPRZED")
                End If
                If chkStrony.Checked Then
                    Obrot += GetDblDRVField(drv, "STRONYILOSPRZED")
                End If
                If chkDrukarki.Checked Then
                    Obrot += GetDblDRVField(drv, "DRUKARKIILOSPRZED")
                End If
                If chkSkup.Checked Then
                    Obrot += GetDblDRVField(drv, "SKUPILOSPRZED")
                End If

            Next
        End If

        DataViewObroty.RowFilter = ""
        Return (Obrot)
    End Function



    Private Function AnalizujWartObrotyBiezMies(ByVal pKlient As String) As Double
        Dim drv As DataRowView
        Dim i As Integer = 0
        Dim Obrot As Double = 0

        'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
        'Public oDataObr As String             'DATA             Text,10

        'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedObr As Double       'LILOPRZED        Double
        'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedObr As Double       'AILOSPRZED       Double

        'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
        'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double


        DataViewObroty.RowFilter = "IDENTKLIENTA='" & pKlient & "'"

        If DataViewObroty.Count > 0 Then
            For i = 0 To DataViewObroty.Count - 1
                drv = DataViewObroty.Item(i)

                If chkPryzmat.Checked Then
                    Obrot += GetDblDRVField(drv, "LWARTSPRZED") + GetDblDRVField(drv, "AWARTSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If
                If chkPryzmatAtrament.Checked Then
                    Obrot += GetDblDRVField(drv, "AWARTSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If
                If chkPryzmatTonery.Checked Then
                    Obrot += GetDblDRVField(drv, "LWARTSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If

                If chkOryg.Checked Then
                    Obrot += GetDblDRVField(drv, "AOWARTSPRZED") + GetDblDRVField(drv, "LOWARTSPRZED")   'sprzedaz atranametow i laserów oryg
                End If
                If chkOrygAtrament.Checked Then
                    Obrot += GetDblDRVField(drv, "AOWARTSPRZED")   'sprzedaz atranametow i laserów oryg
                End If
                If chkOrygTonery.Checked Then
                    Obrot += GetDblDRVField(drv, "LOWARTSPRZED")   'sprzedaz atranametow i laserów oryg
                End If

                If chkNajem.Checked Then
                    Obrot += GetDblDRVField(drv, "NAJEMWARTSPRZED")
                End If
                If chkStrony.Checked Then
                    Obrot += GetDblDRVField(drv, "STRONYWARTSPRZED")
                End If
                If chkDrukarki.Checked Then
                    Obrot += GetDblDRVField(drv, "DRUKARKIWARTSPRZED")
                End If
                If chkSkup.Checked Then
                    Obrot += GetDblDRVField(drv, "SKUPWARTSPRZED")
                End If
            Next
        End If

        DataViewObroty.RowFilter = ""
        Return (Obrot)
    End Function





    Private Function AnalizujMarzeObrotyBiezMies(ByVal pKlient As String) As Double
        Dim drv As DataRowView
        Dim i As Integer = 0
        Dim Obrot As Double = 0

        'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
        'Public oDataObr As String             'DATA             Text,10

        'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedObr As Double       'LILOPRZED        Double
        'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedObr As Double       'AILOSPRZED       Double

        'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
        'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double


        DataViewObroty.RowFilter = "IDENTKLIENTA='" & pKlient & "'"

        If DataViewObroty.Count > 0 Then
            For i = 0 To DataViewObroty.Count - 1
                drv = DataViewObroty.Item(i)

                If chkPryzmat.Checked Then
                    Obrot += GetDblDRVField(drv, "LMARSPRZED") + GetDblDRVField(drv, "AMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If
                If chkPryzmatAtrament.Checked Then
                    Obrot += GetDblDRVField(drv, "AMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If
                If chkPryzmatTonery.Checked Then
                    Obrot += GetDblDRVField(drv, "LMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If

                If chkOryg.Checked Then
                    Obrot += GetDblDRVField(drv, "AOMARSPRZED") + GetDblDRVField(drv, "LOMARSPRZED")   'sprzedaz atranametow i laserów oryg
                End If
                If chkOrygAtrament.Checked Then
                    Obrot += GetDblDRVField(drv, "AOMARSPRZED")   'sprzedaz atranametow i laserów oryg
                End If
                If chkOrygTonery.Checked Then
                    Obrot += GetDblDRVField(drv, "LOMARSPRZED")   'sprzedaz atranametow i laserów oryg
                End If

                If chkNajem.Checked Then
                    Obrot += GetDblDRVField(drv, "NAJEMMARSPRZED")
                End If
                If chkStrony.Checked Then
                    Obrot += GetDblDRVField(drv, "STRONYMARSPRZED")
                End If
                If chkDrukarki.Checked Then
                    Obrot += GetDblDRVField(drv, "DRUKARKIMARSPRZED")
                End If
                If chkSkup.Checked Then
                    Obrot += GetDblDRVField(drv, "SKUPMARSPRZED")
                End If
            Next
        End If

        DataViewObroty.RowFilter = ""
        Return (Obrot)
    End Function



    Private Function AnalizujProcMObrotyBiezMies(ByVal pKlient As String) As Double
        Dim drv As DataRowView
        Dim i As Integer = 0
        Dim Obrot As Double = 0

        'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
        'Public oDataObr As String             'DATA             Text,10

        'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedObr As Double       'LILOPRZED        Double
        'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedObr As Double       'AILOSPRZED       Double

        'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
        'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double


        DataViewObroty.RowFilter = "IDENTKLIENTA='" & pKlient & "'"

        If DataViewObroty.Count > 0 Then
            For i = 0 To DataViewObroty.Count - 1
                drv = DataViewObroty.Item(i)

                If chkPryzmat.Checked Then
                    Obrot += GetDblDRVField(drv, "LMARSPRZED") + GetDblDRVField(drv, "AMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If
                If chkPryzmatAtrament.Checked Then
                    Obrot += GetDblDRVField(drv, "AMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If
                If chkPryzmatTonery.Checked Then
                    Obrot += GetDblDRVField(drv, "LMARSPRZED")   'sprzedaz laserów i atramentu PRYZMAT
                End If

                If chkOryg.Checked Then
                    Obrot += GetDblDRVField(drv, "AOMARSPRZED") + GetDblDRVField(drv, "LOMARSPRZED")   'sprzedaz atranametow i laserów oryg
                End If
                If chkOrygAtrament.Checked Then
                    Obrot += GetDblDRVField(drv, "AOMARSPRZED")   'sprzedaz atranametow i laserów oryg
                End If
                If chkOrygTonery.Checked Then
                    Obrot += GetDblDRVField(drv, "LOMARSPRZED")   'sprzedaz atranametow i laserów oryg
                End If

                If chkNajem.Checked Then
                    Obrot += GetDblDRVField(drv, "NAJEMMARSPRZED")
                End If
                If chkStrony.Checked Then
                    Obrot += GetDblDRVField(drv, "STRONYMARSPRZED")
                End If
                If chkDrukarki.Checked Then
                    Obrot += GetDblDRVField(drv, "DRUKARKIMARSPRZED")
                End If
                If chkSkup.Checked Then
                    Obrot += GetDblDRVField(drv, "SKUPMARSPRZED")
                End If
            Next
        End If

        DataViewObroty.RowFilter = ""
        Return (Obrot)
    End Function







    Private Sub cmdZaznacz1_Click(sender As Object, e As EventArgs) Handles cmdZaznacz1.Click
        Chk101.Checked = True
        Chk102.Checked = True
        Chk103.Checked = True
        Chk104.Checked = True
        Chk105.Checked = True
        Chk106.Checked = True
        Chk107.Checked = True
        Chk108.Checked = True
        Chk109.Checked = True
        Chk110.Checked = True
        Chk111.Checked = True
        Chk112.Checked = True
    End Sub

    Private Sub cmdZaznacz2_Click(sender As Object, e As EventArgs) Handles cmdZaznacz2.Click
        Chk201.Checked = True
        Chk202.Checked = True
        Chk203.Checked = True
        Chk204.Checked = True
        Chk205.Checked = True
        Chk206.Checked = True
        Chk207.Checked = True
        Chk208.Checked = True
        Chk209.Checked = True
        Chk210.Checked = True
        Chk211.Checked = True
        Chk212.Checked = True
    End Sub

    Private Sub cmdOdznacz1_Click(sender As Object, e As EventArgs) Handles cmdOdznacz1.Click
        Chk101.Checked = False
        Chk102.Checked = False
        Chk103.Checked = False
        Chk104.Checked = False
        Chk105.Checked = False
        Chk106.Checked = False
        Chk107.Checked = False
        Chk108.Checked = False
        Chk109.Checked = False
        Chk110.Checked = False
        Chk111.Checked = False
        Chk112.Checked = False
    End Sub

    Private Sub cmdOdznacz2_Click(sender As Object, e As EventArgs) Handles cmdOdznacz2.Click
        Chk201.Checked = False
        Chk202.Checked = False
        Chk203.Checked = False
        Chk204.Checked = False
        Chk205.Checked = False
        Chk206.Checked = False
        Chk207.Checked = False
        Chk208.Checked = False
        Chk209.Checked = False
        Chk210.Checked = False
        Chk211.Checked = False
        Chk212.Checked = False
    End Sub





    Private Sub cbxRok1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxRok1.SelectedIndexChanged
        ZmienOpis()
    End Sub

    Private Sub cbxMies1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxMies1.SelectedIndexChanged
        ZmienOpis()
    End Sub

    Private Sub cbxRok2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxRok2.SelectedIndexChanged
        ZmienOpis()
    End Sub

    Private Sub cbxMies2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxMies2.SelectedIndexChanged
        ZmienOpis()
    End Sub


    Private Sub ZmienOpis()
        Dim Okres1 As String = cbxRok1.Text & "-" & cbxMies1.Text
        Dim Okres2 As String = cbxRok2.Text & "-" & cbxMies2.Text

        lblZakres1.Text = "Zakres " & InnyMiesiac(Okres1, 0) & ".." & InnyMiesiac(Okres1, -11)
        lblZakres2.Text = "Zakres " & InnyMiesiac(Okres2, 0) & ".." & InnyMiesiac(Okres2, -11)

        ToolTip1.SetToolTip(Me.Chk101, InnyMiesiac(Okres1, 0))
        ToolTip1.SetToolTip(Me.Chk102, InnyMiesiac(Okres1, -1))
        ToolTip1.SetToolTip(Me.Chk103, InnyMiesiac(Okres1, -2))
        ToolTip1.SetToolTip(Me.Chk104, InnyMiesiac(Okres1, -3))
        ToolTip1.SetToolTip(Me.Chk105, InnyMiesiac(Okres1, -4))
        ToolTip1.SetToolTip(Me.Chk106, InnyMiesiac(Okres1, -5))
        ToolTip1.SetToolTip(Me.Chk107, InnyMiesiac(Okres1, -6))
        ToolTip1.SetToolTip(Me.Chk108, InnyMiesiac(Okres1, -7))
        ToolTip1.SetToolTip(Me.Chk109, InnyMiesiac(Okres1, -8))
        ToolTip1.SetToolTip(Me.Chk110, InnyMiesiac(Okres1, -9))
        ToolTip1.SetToolTip(Me.Chk111, InnyMiesiac(Okres1, -10))
        ToolTip1.SetToolTip(Me.Chk112, InnyMiesiac(Okres1, -11))

        ToolTip1.SetToolTip(Me.Chk201, InnyMiesiac(Okres2, 0))
        ToolTip1.SetToolTip(Me.Chk202, InnyMiesiac(Okres2, -1))
        ToolTip1.SetToolTip(Me.Chk203, InnyMiesiac(Okres2, -2))
        ToolTip1.SetToolTip(Me.Chk204, InnyMiesiac(Okres2, -3))
        ToolTip1.SetToolTip(Me.Chk205, InnyMiesiac(Okres2, -4))
        ToolTip1.SetToolTip(Me.Chk206, InnyMiesiac(Okres2, -5))
        ToolTip1.SetToolTip(Me.Chk207, InnyMiesiac(Okres2, -6))
        ToolTip1.SetToolTip(Me.Chk208, InnyMiesiac(Okres2, -7))
        ToolTip1.SetToolTip(Me.Chk209, InnyMiesiac(Okres2, -8))
        ToolTip1.SetToolTip(Me.Chk210, InnyMiesiac(Okres2, -9))
        ToolTip1.SetToolTip(Me.Chk211, InnyMiesiac(Okres2, -10))
        ToolTip1.SetToolTip(Me.Chk212, InnyMiesiac(Okres2, -11))

    End Sub




    Private Sub chkPryzmat_CheckedChanged(sender As Object, e As EventArgs) Handles chkPryzmat.CheckedChanged
        If chkPryzmat.Checked = True Then
            chkPryzmatAtrament.Checked = False
            chkPryzmatTonery.Checked = False
        End If
    End Sub

    Private Sub chkPryzmatAtrament_CheckedChanged(sender As Object, e As EventArgs) Handles chkPryzmatAtrament.CheckedChanged
        If chkPryzmatAtrament.Checked = True Then
            chkPryzmat.Checked = False
            chkPryzmatTonery.Checked = False
        End If
    End Sub

    Private Sub chkPryzmatTonery_CheckedChanged(sender As Object, e As EventArgs) Handles chkPryzmatTonery.CheckedChanged
        If chkPryzmatTonery.Checked = True Then
            chkPryzmatAtrament.Checked = False
            chkPryzmat.Checked = False
        End If
    End Sub





    Private Sub chkOryg_CheckedChanged(sender As Object, e As EventArgs) Handles chkOryg.CheckedChanged
        If chkOryg.Checked = True Then
            chkOrygAtrament.Checked = False
            chkOrygTonery.Checked = False
        End If
    End Sub

    Private Sub chkOrygAtrament_CheckedChanged(sender As Object, e As EventArgs) Handles chkOrygAtrament.CheckedChanged
        If chkOrygAtrament.Checked = True Then
            chkOryg.Checked = False
            chkOrygTonery.Checked = False
        End If
    End Sub

    Private Sub chkOrygTonery_CheckedChanged(sender As Object, e As EventArgs) Handles chkOrygTonery.CheckedChanged
        If chkOrygTonery.Checked = True Then
            chkOrygAtrament.Checked = False
            chkOryg.Checked = False
        End If
    End Sub





End Class

