Public Class PobierzPrognozydoAnalizy2
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
    Friend WithEvents Label17 As System.Windows.Forms.Label
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
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PobierzPrognozydoAnalizy2))
        Me.Panel1 = New System.Windows.Forms.Panel()
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
        Me.Label9 = New System.Windows.Forms.Label()
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
        Me.Label17 = New System.Windows.Forms.Label()
        Me.chkSkup = New System.Windows.Forms.CheckBox()
        Me.chkDrukarki = New System.Windows.Forms.CheckBox()
        Me.chkStrony = New System.Windows.Forms.CheckBox()
        Me.chkNajem = New System.Windows.Forms.CheckBox()
        Me.chkOryg = New System.Windows.Forms.CheckBox()
        Me.chkPryzmat = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblDataAktualizacji = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.lblIlePrzeczytano = New System.Windows.Forms.Label()
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
        Me.Panel1.Controls.Add(Me.Label9)
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
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.chkSkup)
        Me.Panel1.Controls.Add(Me.chkDrukarki)
        Me.Panel1.Controls.Add(Me.chkStrony)
        Me.Panel1.Controls.Add(Me.chkNajem)
        Me.Panel1.Controls.Add(Me.chkOryg)
        Me.Panel1.Controls.Add(Me.chkPryzmat)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblDataAktualizacji)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.ProgressBar)
        Me.Panel1.Controls.Add(Me.lblIlePrzeczytano)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(666, 390)
        Me.Panel1.TabIndex = 30
        '
        'cmdOdznacz2
        '
        Me.cmdOdznacz2.Image = CType(resources.GetObject("cmdOdznacz2.Image"), System.Drawing.Image)
        Me.cmdOdznacz2.Location = New System.Drawing.Point(615, 168)
        Me.cmdOdznacz2.Name = "cmdOdznacz2"
        Me.cmdOdznacz2.Size = New System.Drawing.Size(24, 21)
        Me.cmdOdznacz2.TabIndex = 239
        Me.cmdOdznacz2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdznacz2, "Odznacz wszystko")
        '
        'cmdZaznacz2
        '
        Me.cmdZaznacz2.Image = CType(resources.GetObject("cmdZaznacz2.Image"), System.Drawing.Image)
        Me.cmdZaznacz2.Location = New System.Drawing.Point(585, 168)
        Me.cmdZaznacz2.Name = "cmdZaznacz2"
        Me.cmdZaznacz2.Size = New System.Drawing.Size(24, 21)
        Me.cmdZaznacz2.TabIndex = 238
        Me.cmdZaznacz2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdZaznacz2, "Zaznacz wszystko")
        '
        'cmdOdznacz1
        '
        Me.cmdOdznacz1.Image = CType(resources.GetObject("cmdOdznacz1.Image"), System.Drawing.Image)
        Me.cmdOdznacz1.Location = New System.Drawing.Point(615, 148)
        Me.cmdOdznacz1.Name = "cmdOdznacz1"
        Me.cmdOdznacz1.Size = New System.Drawing.Size(24, 21)
        Me.cmdOdznacz1.TabIndex = 237
        Me.cmdOdznacz1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdznacz1, "Odznacz wszystko")
        '
        'cmdZaznacz1
        '
        Me.cmdZaznacz1.Image = CType(resources.GetObject("cmdZaznacz1.Image"), System.Drawing.Image)
        Me.cmdZaznacz1.Location = New System.Drawing.Point(585, 148)
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
        Me.Panel2.Size = New System.Drawing.Size(665, 1)
        Me.Panel2.TabIndex = 229
        '
        'Chk212
        '
        Me.Chk212.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk212.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk212.ForeColor = System.Drawing.Color.Navy
        Me.Chk212.Location = New System.Drawing.Point(550, 171)
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
        Me.Chk211.Location = New System.Drawing.Point(532, 171)
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
        Me.Chk210.Location = New System.Drawing.Point(514, 171)
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
        Me.Chk209.Location = New System.Drawing.Point(496, 171)
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
        Me.Chk208.Location = New System.Drawing.Point(478, 171)
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
        Me.Chk207.Location = New System.Drawing.Point(460, 171)
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
        Me.Chk206.Location = New System.Drawing.Point(442, 171)
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
        Me.Chk205.Location = New System.Drawing.Point(424, 171)
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
        Me.Chk204.Location = New System.Drawing.Point(406, 171)
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
        Me.Chk203.Location = New System.Drawing.Point(388, 171)
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
        Me.Chk202.Location = New System.Drawing.Point(370, 171)
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
        Me.Chk201.Location = New System.Drawing.Point(352, 171)
        Me.Chk201.Name = "Chk201"
        Me.Chk201.Size = New System.Drawing.Size(18, 18)
        Me.Chk201.TabIndex = 217
        Me.Chk201.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk201.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(234, 173)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(131, 16)
        Me.Label9.TabIndex = 216
        Me.Label9.Text = "Zakres analizy 0..-11 . . . ."
        '
        'Chk112
        '
        Me.Chk112.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk112.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk112.ForeColor = System.Drawing.Color.Navy
        Me.Chk112.Location = New System.Drawing.Point(550, 150)
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
        Me.Chk111.Location = New System.Drawing.Point(532, 150)
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
        Me.Chk110.Location = New System.Drawing.Point(514, 150)
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
        Me.Chk109.Location = New System.Drawing.Point(496, 150)
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
        Me.Chk108.Location = New System.Drawing.Point(478, 150)
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
        Me.Chk107.Location = New System.Drawing.Point(460, 150)
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
        Me.Chk106.Location = New System.Drawing.Point(442, 150)
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
        Me.Chk105.Location = New System.Drawing.Point(424, 150)
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
        Me.Chk104.Location = New System.Drawing.Point(406, 150)
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
        Me.Chk103.Location = New System.Drawing.Point(388, 150)
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
        Me.Chk102.Location = New System.Drawing.Point(370, 150)
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
        Me.Chk101.Location = New System.Drawing.Point(352, 150)
        Me.Chk101.Name = "Chk101"
        Me.Chk101.Size = New System.Drawing.Size(18, 18)
        Me.Chk101.TabIndex = 204
        Me.Chk101.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk101.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Navy
        Me.Label17.Location = New System.Drawing.Point(234, 152)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(131, 16)
        Me.Label17.TabIndex = 203
        Me.Label17.Text = "Zakres analizy 0..-11 . . . ."
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
        Me.chkOryg.Text = "Oryginalne"
        Me.chkOryg.UseVisualStyleBackColor = True
        '
        'chkPryzmat
        '
        Me.chkPryzmat.ForeColor = System.Drawing.Color.Navy
        Me.chkPryzmat.Location = New System.Drawing.Point(38, 229)
        Me.chkPryzmat.Name = "chkPryzmat"
        Me.chkPryzmat.Size = New System.Drawing.Size(160, 17)
        Me.chkPryzmat.TabIndex = 60
        Me.chkPryzmat.Text = "Pryzmat"
        Me.chkPryzmat.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 212)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(190, 16)
        Me.Label3.TabIndex = 59
        Me.Label3.Text = "Asortyment towarowy do analizy :"
        '
        'lblDataAktualizacji
        '
        Me.lblDataAktualizacji.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDataAktualizacji.ForeColor = System.Drawing.Color.Purple
        Me.lblDataAktualizacji.Location = New System.Drawing.Point(508, 246)
        Me.lblDataAktualizacji.Name = "lblDataAktualizacji"
        Me.lblDataAktualizacji.Size = New System.Drawing.Size(88, 17)
        Me.lblDataAktualizacji.TabIndex = 56
        Me.lblDataAktualizacji.Text = "0"
        Me.lblDataAktualizacji.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(325, 248)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(240, 16)
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
        Me.Label1.Size = New System.Drawing.Size(643, 129)
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
        Me.Panel3.Size = New System.Drawing.Size(664, 1)
        Me.Panel3.TabIndex = 53
        '
        'ProgressBar
        '
        Me.ProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar.Location = New System.Drawing.Point(11, 356)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(640, 8)
        Me.ProgressBar.TabIndex = 49
        '
        'lblIlePrzeczytano
        '
        Me.lblIlePrzeczytano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlePrzeczytano.ForeColor = System.Drawing.Color.Purple
        Me.lblIlePrzeczytano.Location = New System.Drawing.Point(508, 298)
        Me.lblIlePrzeczytano.Name = "lblIlePrzeczytano"
        Me.lblIlePrzeczytano.Size = New System.Drawing.Size(88, 17)
        Me.lblIlePrzeczytano.TabIndex = 26
        Me.lblIlePrzeczytano.Text = "0"
        Me.lblIlePrzeczytano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(325, 300)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(240, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Iloœæ rekordów przeczytanych . . . . . . . . . . . . . . . . . . "
        '
        'cmdImportuj
        '
        Me.cmdImportuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdImportuj.Image = CType(resources.GetObject("cmdImportuj.Image"), System.Drawing.Image)
        Me.cmdImportuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdImportuj.Location = New System.Drawing.Point(506, 406)
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
        Me.cmdPowrot.Location = New System.Drawing.Point(588, 406)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(81, 22)
        Me.cmdPowrot.TabIndex = 34
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(15, 405)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 23)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 33
        Me.PictureBox2.TabStop = False
        '
        'PobierzPrognozydoAnalizy2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(683, 434)
        Me.Controls.Add(Me.cmdImportuj)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "PobierzPrognozydoAnalizy2"
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
    Dim dbSelectSQLKlienci As String = ""

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
        ''-------------------------------
        ''zapamietaj warunki wyboru do analizy
        'Par_OstOkresAnalizy = Microsoft.VisualBasic.Right("0000" & Trim(cbxRok1.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies1.Text), 2) & _
        '                      Microsoft.VisualBasic.Right("0000" & Trim(cbxRok2.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies2.Text), 2)

        'Par_DataAktAnalizy = SysData()
        'Par_DaneDoAnalizy = IIf(chkToneryOryg.Checked, "T", "N") & _
        '                    IIf(chkToneryPryzmat.Checked, "T", "N") & _
        '                    IIf(chkAtramentOryg.Checked, "T", "N") & _
        '                    IIf(chkAtramentPryzmat.Checked, "T", "N") & _
        '                    IIf(chkIlosci.Checked, "T", "N") & _
        '                    IIf(chkWartosci.Checked, "T", "N")

        'Par_MiesAnalizy1 = IIf(Chk101.Checked, "T", "N") & _
        '                   IIf(Chk102.Checked, "T", "N") & _
        '                   IIf(Chk103.Checked, "T", "N") & _
        '                   IIf(Chk104.Checked, "T", "N") & _
        '                   IIf(Chk105.Checked, "T", "N") & _
        '                   IIf(Chk106.Checked, "T", "N") & _
        '                   IIf(Chk107.Checked, "T", "N") & _
        '                   IIf(Chk108.Checked, "T", "N") & _
        '                   IIf(Chk109.Checked, "T", "N") & _
        '                   IIf(Chk110.Checked, "T", "N") & _
        '                   IIf(Chk111.Checked, "T", "N") & _
        '                   IIf(Chk112.Checked, "T", "N")
        'Par_MiesAnalizy2 = IIf(Chk201.Checked, "T", "N") & _
        '                   IIf(Chk202.Checked, "T", "N") & _
        '                   IIf(Chk203.Checked, "T", "N") & _
        '                   IIf(Chk204.Checked, "T", "N") & _
        '                   IIf(Chk205.Checked, "T", "N") & _
        '                   IIf(Chk206.Checked, "T", "N") & _
        '                   IIf(Chk207.Checked, "T", "N") & _
        '                   IIf(Chk208.Checked, "T", "N") & _
        '                   IIf(Chk209.Checked, "T", "N") & _
        '                   IIf(Chk210.Checked, "T", "N") & _
        '                   IIf(Chk211.Checked, "T", "N") & _
        '                   IIf(Chk212.Checked, "T", "N")

        'ZapiszParametry(Me)
        ''-------------------------------
        Me.Close()
    End Sub



    Private Sub PobierzPrognozydoAnalizy2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        chkOryg.Checked = (Mid(Par_DaneDoAnalizy, 2, 1) = "T")
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


    Private Sub PobierzPrognozydoAnalizy2_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
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

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelectKlienci = New OleDb.OleDbCommand(sSelectSQLKlienci & " ORDER BY IDENTKLIENTA", dbConnectionKlienci)
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

            'DataSet
            DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

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

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionObrotyMies)
            dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
            'dbCommandDeleteObrotyMies = New OleDb.OleDbCommand(sDeleteSQLObrotyMies, dbConnectionObrotyMies)
            'dbCommandUpdateObrotyMies = New OleDb.OleDbCommand(sUpdateSQLObrotyMies, dbConnectionObrotyMies)
            'dbCommandInsertObrotyMies = New OleDb.OleDbCommand(sInsertSQLObrotyMies, dbConnectionObrotyMies)
            dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

            'DataSet
            DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            MapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

            ' Add the RowUpdated event handler.
            AddHandler dbDataAdapterObrotyMies.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObrotyMies.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnectionObrotyMies.State
            End Try

        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionObrotyMies)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandDeleteObrotyMies = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandUpdateObrotyMies = New SqlClient.SqlCommand(sUpdateSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandInsertObrotyMies = New SqlClient.SqlCommand(sInsertSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

            'DataSet
            DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")

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

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObroty = New OleDb.OleDbConnection(sConnectionObroty)
            dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
            'dbCommandDeleteObroty = New OleDb.OleDbCommand(sDeleteSQLObroty, dbConnectionObroty)
            'dbCommandUpdateObroty = New OleDb.OleDbCommand(sUpdateSQLObroty, dbConnectionObroty)
            'dbCommandInsertObroty = New OleDb.OleDbCommand(sInsertSQLObroty, dbConnectionObroty)
            dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

            'DataSet
            DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
            MapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

            ' Add the RowUpdated event handler.
            AddHandler dbDataAdapterObroty.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnectionObroty.State
            End Try

        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionObroty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
            'sqlCommandDeleteObroty = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnectionObroty)
            'sqlCommandUpdateObroty = New SqlClient.SqlCommand(sUpdateSQLObroty, sqlConnectionObroty)
            'sqlCommandInsertObroty = New SqlClient.SqlCommand(sInsertSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

            'DataSet
            DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

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
            nrAnalizy.Item("ILOSZA00") = 0


            '--------------------------------------
            'pierwszy okres do analizy,,,
            OkresAn = Microsoft.VisualBasic.Right("0000" & Trim(cbxRok1.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies1.Text), 2)

            If Not Chk101.Checked Then
                ILBM = 0
                WABM = 0
            Else
                ILBM = AnalizujObrotyIlo(SymbKlienta, OkresAn)
                WABM = AnalizujObrotyWar(SymbKlienta, OkresAn)
            End If
            nrAnalizy.Item("WARTZABM") = WABM
            nrAnalizy.Item("ILOSZABM") = ILBM

            'analizuj poprzednie miesiace...
            If Not Chk102.Checked Then
                IL01 = 0
                WA01 = 0
            Else
                IL01 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
                WA01 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
            End If
            nrAnalizy.Item("WARTZA01") = WA01
            nrAnalizy.Item("ILOSZA01") = IL01

            If Not Chk103.Checked Then
                IL02 = 0
                WA02 = 0
            Else
                IL02 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
                WA02 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
            End If
            nrAnalizy.Item("WARTZA02") = WA02
            nrAnalizy.Item("ILOSZA02") = IL02

            If Not Chk104.Checked Then
                IL03 = 0
                WA03 = 0
            Else
                IL03 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
                WA03 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
            End If
            nrAnalizy.Item("WARTZA03") = WA03
            nrAnalizy.Item("ILOSZA03") = IL03

            If Not Chk105.Checked Then
                IL04 = 0
                WA04 = 0
            Else
                IL04 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
                WA04 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
            End If
            nrAnalizy.Item("WARTZA04") = WA04
            nrAnalizy.Item("ILOSZA04") = IL04

            If Not Chk106.Checked Then
                IL05 = 0
                WA05 = 0
            Else
                IL05 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
                WA05 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
            End If
            nrAnalizy.Item("WARTZA05") = WA05
            nrAnalizy.Item("ILOSZA05") = IL05

            If Not Chk107.Checked Then
                IL06 = 0
                WA06 = 0
            Else
                IL06 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
                WA06 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
            End If
            nrAnalizy.Item("WARTZA06") = WA06
            nrAnalizy.Item("ILOSZA06") = IL06

            If Not Chk108.Checked Then
                IL07 = 0
                WA07 = 0
            Else
                IL07 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
                WA07 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
            End If
            nrAnalizy.Item("WARTZA07") = WA07
            nrAnalizy.Item("ILOSZA07") = IL07

            If Not Chk109.Checked Then
                IL08 = 0
                WA08 = 0
            Else
                IL08 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
                WA08 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
            End If
            nrAnalizy.Item("WARTZA08") = WA08
            nrAnalizy.Item("ILOSZA08") = IL08

            If Not Chk110.Checked Then
                IL09 = 0
                WA09 = 0
            Else
                IL09 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
                WA09 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
            End If
            nrAnalizy.Item("WARTZA09") = WA09
            nrAnalizy.Item("ILOSZA09") = IL09

            If Not Chk111.Checked Then
                IL10 = 0
                WA10 = 0
            Else
                IL10 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
                WA10 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
            End If
            nrAnalizy.Item("WARTZA10") = WA10
            nrAnalizy.Item("ILOSZA10") = IL10

            If Not Chk112.Checked Then
                IL11 = 0
                WA11 = 0
            Else
                IL11 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
                WA11 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
            End If
            nrAnalizy.Item("WARTZA11") = WA11
            nrAnalizy.Item("ILOSZA11") = IL11



            '--------------------------------------
            'drugi okres do analizy,,,
            OkresAn = Microsoft.VisualBasic.Right("0000" & Trim(cbxRok2.Text), 4) & "-" & Microsoft.VisualBasic.Right("00" & Trim(cbxMies2.Text), 2)

            If Not Chk201.Checked Then
                IL12 = 0
                WA12 = 0
            Else
                IL12 = AnalizujObrotyIlo(SymbKlienta, OkresAn)
                WA12 = AnalizujObrotyWar(SymbKlienta, OkresAn)
            End If
            nrAnalizy.Item("WARTZA12") = WA12
            nrAnalizy.Item("ILOSZA12") = IL12

            If Not Chk202.Checked Then
                IL13 = 0
                WA13 = 0
            Else
                IL13 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
                WA13 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -1))
            End If
            nrAnalizy.Item("WARTZA13") = WA13
            nrAnalizy.Item("ILOSZA13") = IL13

            If Not Chk203.Checked Then
                IL14 = 0
                WA14 = 0
            Else
                IL14 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
                WA14 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -2))
            End If
            nrAnalizy.Item("WARTZA14") = WA14
            nrAnalizy.Item("ILOSZA14") = IL14

            If Not Chk204.Checked Then
                IL15 = 0
                WA15 = 0
            Else
                IL15 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
                WA15 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -3))
            End If
            nrAnalizy.Item("WARTZA15") = WA15
            nrAnalizy.Item("ILOSZA15") = IL15

            If Not Chk205.Checked Then
                IL16 = 0
                WA16 = 0
            Else
                IL16 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
                WA16 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -4))
            End If
            nrAnalizy.Item("WARTZA16") = WA16
            nrAnalizy.Item("ILOSZA16") = IL16

            If Not Chk206.Checked Then
                IL17 = 0
                WA17 = 0
            Else
                IL17 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
                WA17 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -5))
            End If
            nrAnalizy.Item("WARTZA17") = WA17
            nrAnalizy.Item("ILOSZA17") = IL17

            If Not Chk207.Checked Then
                IL18 = 0
                WA18 = 0
            Else
                IL18 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
                WA18 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -6))
            End If
            nrAnalizy.Item("WARTZA18") = WA18
            nrAnalizy.Item("ILOSZA18") = IL18

            If Not Chk208.Checked Then
                IL19 = 0
                WA19 = 0
            Else
                IL19 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
                WA19 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -7))
            End If
            nrAnalizy.Item("WARTZA19") = WA19
            nrAnalizy.Item("ILOSZA19") = IL19

            If Not Chk209.Checked Then
                IL20 = 0
                WA20 = 0
            Else
                IL20 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
                WA20 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -8))
            End If
            nrAnalizy.Item("WARTZA20") = WA20
            nrAnalizy.Item("ILOSZA20") = IL20

            If Not Chk210.Checked Then
                IL21 = 0
                WA21 = 0
            Else
                IL21 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
                WA21 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -9))
            End If
            nrAnalizy.Item("WARTZA21") = WA21
            nrAnalizy.Item("ILOSZA21") = IL21

            If Not Chk211.Checked Then
                IL22 = 0
                WA22 = 0
            Else
                IL22 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
                WA22 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -10))
            End If
            nrAnalizy.Item("WARTZA22") = WA22
            nrAnalizy.Item("ILOSZA22") = IL22

            If Not Chk212.Checked Then
                IL23 = 0
                WA23 = 0
            Else
                IL23 = AnalizujObrotyIlo(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
                WA23 = AnalizujObrotyWar(SymbKlienta, MiesDoAnalizy(OkresAn, -11))
            End If
            nrAnalizy.Item("WARTZA23") = WA23
            nrAnalizy.Item("ILOSZA23") = IL23

            'IL24 = AnalizujObrotyKlienta(SymbKlienta, MiesDoAnalizy(-24))
            'WA24 = AnalizujWartObrotyKlienta(SymbKlienta, MiesDoAnalizy(-24))
            nrAnalizy.Item("WARTZA24") = 0
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
        Par_DaneDoAnalizy = IIf(chkPryzmat.Checked, "T", "N") &
                            IIf(chkOryg.Checked, "T", "N") &
                            IIf(chkNajem.Checked, "T", "N") &
                            IIf(chkStrony.Checked, "T", "N") &
                            IIf(chkDrukarki.Checked, "T", "N") &
                            IIf(chkSkup.Checked, "T", "N")

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
            Return AnalizujObrotyBiezMies(pKlient)
            'Return AnalizujWartObrotyBiezMies(pKlient)
        Else
            'analizuj poprzednie miesiace...
            Return AnalizujObrotyKlienta(pKlient, pMiesiac)
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






    '************************************
    ' Analizuj obroty klienta w poszczególnych miesiacach...
    '************************************

    Private Function AnalizujObrotyKlienta(ByVal pKlient As String,
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
            If chkOryg.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOILOSPRZED") + GetDblDRField(foundRow, "LOILOSPRZED")   'sprzedaz atranametow i laserów oryg
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
        If chkDrukarki.Checked Then
            Return (ObrMiesieczny)
        Else
            Return (0)
        End If
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
            If chkOryg.Checked Then
                ObrMiesieczny += GetDblDRField(foundRow, "AOWARTSPRZED") + GetDblDRField(foundRow, "LOWARTSPRZED")   'sprzedaz atranametow i laserów oryg
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
        If chkSkup.Checked Then
            Return (ObrMiesieczny)
        Else
            Return (0)
        End If
    End Function




    Private Function AnalizujObrotyBiezMies(ByVal pKlient As String) As Double
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
                    Obrot += GetTxtDRVField(drv, "LOILOSPRZED")   'sprzedaz laserów PRYZMAT
                End If
                If chkNajem.Checked Then
                    Obrot += GetDblDRVField(drv, "LILOSPRZED")   'sprzedaz laserów PRYZMAT
                End If
                If chkOryg.Checked Then
                    Obrot += GetDblDRVField(drv, "AOILOSPRZED")   'sprzedaz atranametow oryg
                End If
                If chkStrony.Checked Then
                    Obrot += GetDblDRVField(drv, "AILOSPRZED")   'sprzedaz atramentow PRYZMAT
                End If
            Next
        End If

        DataViewObroty.RowFilter = ""
        If chkDrukarki.Checked Then
            Return (Obrot)
        Else
            Return (0)
        End If
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
                    Obrot += GetTxtDRVField(drv, "LOWARTSPRZED")   'sprzedaz laserów PRYZMAT
                End If
                If chkNajem.Checked Then
                    Obrot += GetDblDRVField(drv, "LWARTSPRZED")   'sprzedaz laserów PRYZMAT
                End If
                If chkOryg.Checked Then
                    Obrot += GetDblDRVField(drv, "AOWARTSPRZED")   'sprzedaz atranametow oryg
                End If
                If chkStrony.Checked Then
                    Obrot += GetDblDRVField(drv, "AWARTSPRZED")   'sprzedaz atramentow PRYZMAT
                End If
            Next
        End If

        DataViewObroty.RowFilter = ""
        If chkSkup.Checked Then
            Return (Obrot)
        Else
            Return (0)
        End If
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
End Class

