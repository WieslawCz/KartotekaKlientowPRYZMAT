<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAnalizaCzestKontaktowGrupowa
    Inherits System.Windows.Forms.Form

    Dim _dsKlenci As DataSet
    Dim _AktualizujKartKlientow As delegateAktualizuj

    Public Sub New(ByVal ds As DataSet, _
                   ByVal Aktualizuj As delegateAktualizuj)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _dsKlenci = ds
        _AktualizujKartKlientow = Aktualizuj
    End Sub




    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAnalizaCzestKontaktowGrupowa))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dtpDoDaty = New System.Windows.Forms.DateTimePicker()
        Me.cmdM10 = New System.Windows.Forms.Button()
        Me.dtpOdDaty = New System.Windows.Forms.DateTimePicker()
        Me.cmdM9 = New System.Windows.Forms.Button()
        Me.cmdM8 = New System.Windows.Forms.Button()
        Me.cmdM7 = New System.Windows.Forms.Button()
        Me.cmdM6 = New System.Windows.Forms.Button()
        Me.cmdM5 = New System.Windows.Forms.Button()
        Me.cmdM4 = New System.Windows.Forms.Button()
        Me.cmdM3 = New System.Windows.Forms.Button()
        Me.cmdM2 = New System.Windows.Forms.Button()
        Me.cmdM1 = New System.Windows.Forms.Button()
        Me.cmdT10 = New System.Windows.Forms.Button()
        Me.cmdT9 = New System.Windows.Forms.Button()
        Me.cmdT8 = New System.Windows.Forms.Button()
        Me.cmdT7 = New System.Windows.Forms.Button()
        Me.cmdT6 = New System.Windows.Forms.Button()
        Me.cmdT5 = New System.Windows.Forms.Button()
        Me.cmdT4 = New System.Windows.Forms.Button()
        Me.cmdT3 = New System.Windows.Forms.Button()
        Me.cmdT2 = New System.Windows.Forms.Button()
        Me.cmdT1 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rabLiczbaWieksze = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaMniesze = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaRowne = New System.Windows.Forms.RadioButton()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rabLiczbaOdDo = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaWieRowne = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaMniejRowne = New System.Windows.Forms.RadioButton()
        Me.rabLiczbaRozne = New System.Windows.Forms.RadioButton()
        Me.cbxIleKontaktowDo = New System.Windows.Forms.ComboBox()
        Me.lblDo = New System.Windows.Forms.Label()
        Me.cmdAnalizuj = New System.Windows.Forms.Button()
        Me.cbxIleKontaktowOd = New System.Windows.Forms.ComboBox()
        Me.lblOd = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.menuRaport2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuWybierzAM = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdZapamietaj = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.rabRKLista = New System.Windows.Forms.RadioButton()
        Me.rabRKWszystkie = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdRodzajeKontaktow = New System.Windows.Forms.Button()
        Me.txtRodzajeKontaktow = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.stbRaport = New System.Windows.Forms.StatusBar()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.dagRaport = New System.Windows.Forms.DataGridView()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuRaport2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.dagRaport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.dtpDoDaty)
        Me.Panel1.Controls.Add(Me.cmdM10)
        Me.Panel1.Controls.Add(Me.dtpOdDaty)
        Me.Panel1.Controls.Add(Me.cmdM9)
        Me.Panel1.Controls.Add(Me.cmdM8)
        Me.Panel1.Controls.Add(Me.cmdM7)
        Me.Panel1.Controls.Add(Me.cmdM6)
        Me.Panel1.Controls.Add(Me.cmdM5)
        Me.Panel1.Controls.Add(Me.cmdM4)
        Me.Panel1.Controls.Add(Me.cmdM3)
        Me.Panel1.Controls.Add(Me.cmdM2)
        Me.Panel1.Controls.Add(Me.cmdM1)
        Me.Panel1.Controls.Add(Me.cmdT10)
        Me.Panel1.Controls.Add(Me.cmdT9)
        Me.Panel1.Controls.Add(Me.cmdT8)
        Me.Panel1.Controls.Add(Me.cmdT7)
        Me.Panel1.Controls.Add(Me.cmdT6)
        Me.Panel1.Controls.Add(Me.cmdT5)
        Me.Panel1.Controls.Add(Me.cmdT4)
        Me.Panel1.Controls.Add(Me.cmdT3)
        Me.Panel1.Controls.Add(Me.cmdT2)
        Me.Panel1.Controls.Add(Me.cmdT1)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(758, 58)
        Me.Panel1.TabIndex = 59
        '
        'dtpDoDaty
        '
        Me.dtpDoDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDoDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpDoDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDoDaty.Location = New System.Drawing.Point(118, 25)
        Me.dtpDoDaty.Name = "dtpDoDaty"
        Me.dtpDoDaty.Size = New System.Drawing.Size(112, 20)
        Me.dtpDoDaty.TabIndex = 367
        Me.dtpDoDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'cmdM10
        '
        Me.cmdM10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM10.ForeColor = System.Drawing.Color.Black
        Me.cmdM10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM10.Location = New System.Drawing.Point(703, 26)
        Me.cmdM10.Name = "cmdM10"
        Me.cmdM10.Size = New System.Drawing.Size(40, 23)
        Me.cmdM10.TabIndex = 101
        Me.cmdM10.Text = "-10M"
        '
        'dtpOdDaty
        '
        Me.dtpOdDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpOdDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpOdDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOdDaty.Location = New System.Drawing.Point(118, 3)
        Me.dtpOdDaty.Name = "dtpOdDaty"
        Me.dtpOdDaty.Size = New System.Drawing.Size(112, 20)
        Me.dtpOdDaty.TabIndex = 366
        Me.dtpOdDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'cmdM9
        '
        Me.cmdM9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM9.ForeColor = System.Drawing.Color.Black
        Me.cmdM9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM9.Location = New System.Drawing.Point(661, 26)
        Me.cmdM9.Name = "cmdM9"
        Me.cmdM9.Size = New System.Drawing.Size(38, 23)
        Me.cmdM9.TabIndex = 100
        Me.cmdM9.Text = "-9M"
        '
        'cmdM8
        '
        Me.cmdM8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM8.ForeColor = System.Drawing.Color.Black
        Me.cmdM8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM8.Location = New System.Drawing.Point(620, 26)
        Me.cmdM8.Name = "cmdM8"
        Me.cmdM8.Size = New System.Drawing.Size(38, 23)
        Me.cmdM8.TabIndex = 99
        Me.cmdM8.Text = "-8M"
        '
        'cmdM7
        '
        Me.cmdM7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM7.ForeColor = System.Drawing.Color.Black
        Me.cmdM7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM7.Location = New System.Drawing.Point(579, 26)
        Me.cmdM7.Name = "cmdM7"
        Me.cmdM7.Size = New System.Drawing.Size(38, 23)
        Me.cmdM7.TabIndex = 98
        Me.cmdM7.Text = "-7M"
        '
        'cmdM6
        '
        Me.cmdM6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM6.ForeColor = System.Drawing.Color.Black
        Me.cmdM6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM6.Location = New System.Drawing.Point(538, 26)
        Me.cmdM6.Name = "cmdM6"
        Me.cmdM6.Size = New System.Drawing.Size(38, 23)
        Me.cmdM6.TabIndex = 97
        Me.cmdM6.Text = "-6M"
        '
        'cmdM5
        '
        Me.cmdM5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM5.ForeColor = System.Drawing.Color.Black
        Me.cmdM5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM5.Location = New System.Drawing.Point(497, 26)
        Me.cmdM5.Name = "cmdM5"
        Me.cmdM5.Size = New System.Drawing.Size(38, 23)
        Me.cmdM5.TabIndex = 96
        Me.cmdM5.Text = "-5M"
        '
        'cmdM4
        '
        Me.cmdM4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM4.ForeColor = System.Drawing.Color.Black
        Me.cmdM4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM4.Location = New System.Drawing.Point(456, 26)
        Me.cmdM4.Name = "cmdM4"
        Me.cmdM4.Size = New System.Drawing.Size(38, 23)
        Me.cmdM4.TabIndex = 95
        Me.cmdM4.Text = "-4M"
        '
        'cmdM3
        '
        Me.cmdM3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM3.ForeColor = System.Drawing.Color.Black
        Me.cmdM3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM3.Location = New System.Drawing.Point(415, 26)
        Me.cmdM3.Name = "cmdM3"
        Me.cmdM3.Size = New System.Drawing.Size(38, 23)
        Me.cmdM3.TabIndex = 94
        Me.cmdM3.Text = "-3M"
        '
        'cmdM2
        '
        Me.cmdM2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM2.ForeColor = System.Drawing.Color.Black
        Me.cmdM2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM2.Location = New System.Drawing.Point(374, 26)
        Me.cmdM2.Name = "cmdM2"
        Me.cmdM2.Size = New System.Drawing.Size(38, 23)
        Me.cmdM2.TabIndex = 93
        Me.cmdM2.Text = "-2M"
        '
        'cmdM1
        '
        Me.cmdM1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM1.ForeColor = System.Drawing.Color.Black
        Me.cmdM1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM1.Location = New System.Drawing.Point(333, 26)
        Me.cmdM1.Name = "cmdM1"
        Me.cmdM1.Size = New System.Drawing.Size(38, 23)
        Me.cmdM1.TabIndex = 92
        Me.cmdM1.Text = "-1M"
        '
        'cmdT10
        '
        Me.cmdT10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT10.ForeColor = System.Drawing.Color.Black
        Me.cmdT10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT10.Location = New System.Drawing.Point(703, 4)
        Me.cmdT10.Name = "cmdT10"
        Me.cmdT10.Size = New System.Drawing.Size(40, 23)
        Me.cmdT10.TabIndex = 91
        Me.cmdT10.Text = "-10T"
        '
        'cmdT9
        '
        Me.cmdT9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT9.ForeColor = System.Drawing.Color.Black
        Me.cmdT9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT9.Location = New System.Drawing.Point(661, 4)
        Me.cmdT9.Name = "cmdT9"
        Me.cmdT9.Size = New System.Drawing.Size(38, 23)
        Me.cmdT9.TabIndex = 90
        Me.cmdT9.Text = "-9T"
        '
        'cmdT8
        '
        Me.cmdT8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT8.ForeColor = System.Drawing.Color.Black
        Me.cmdT8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT8.Location = New System.Drawing.Point(620, 4)
        Me.cmdT8.Name = "cmdT8"
        Me.cmdT8.Size = New System.Drawing.Size(38, 23)
        Me.cmdT8.TabIndex = 89
        Me.cmdT8.Text = "-8T"
        '
        'cmdT7
        '
        Me.cmdT7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT7.ForeColor = System.Drawing.Color.Black
        Me.cmdT7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT7.Location = New System.Drawing.Point(579, 4)
        Me.cmdT7.Name = "cmdT7"
        Me.cmdT7.Size = New System.Drawing.Size(38, 23)
        Me.cmdT7.TabIndex = 88
        Me.cmdT7.Text = "-7T"
        '
        'cmdT6
        '
        Me.cmdT6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT6.ForeColor = System.Drawing.Color.Black
        Me.cmdT6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT6.Location = New System.Drawing.Point(538, 4)
        Me.cmdT6.Name = "cmdT6"
        Me.cmdT6.Size = New System.Drawing.Size(38, 23)
        Me.cmdT6.TabIndex = 87
        Me.cmdT6.Text = "-6T"
        '
        'cmdT5
        '
        Me.cmdT5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT5.ForeColor = System.Drawing.Color.Black
        Me.cmdT5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT5.Location = New System.Drawing.Point(497, 4)
        Me.cmdT5.Name = "cmdT5"
        Me.cmdT5.Size = New System.Drawing.Size(38, 23)
        Me.cmdT5.TabIndex = 86
        Me.cmdT5.Text = "-5T"
        '
        'cmdT4
        '
        Me.cmdT4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT4.ForeColor = System.Drawing.Color.Black
        Me.cmdT4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT4.Location = New System.Drawing.Point(456, 4)
        Me.cmdT4.Name = "cmdT4"
        Me.cmdT4.Size = New System.Drawing.Size(38, 23)
        Me.cmdT4.TabIndex = 85
        Me.cmdT4.Text = "-4T"
        '
        'cmdT3
        '
        Me.cmdT3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT3.ForeColor = System.Drawing.Color.Black
        Me.cmdT3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT3.Location = New System.Drawing.Point(415, 4)
        Me.cmdT3.Name = "cmdT3"
        Me.cmdT3.Size = New System.Drawing.Size(38, 23)
        Me.cmdT3.TabIndex = 84
        Me.cmdT3.Text = "-3T"
        '
        'cmdT2
        '
        Me.cmdT2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT2.ForeColor = System.Drawing.Color.Black
        Me.cmdT2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT2.Location = New System.Drawing.Point(374, 4)
        Me.cmdT2.Name = "cmdT2"
        Me.cmdT2.Size = New System.Drawing.Size(38, 23)
        Me.cmdT2.TabIndex = 83
        Me.cmdT2.Text = "-2T"
        '
        'cmdT1
        '
        Me.cmdT1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT1.ForeColor = System.Drawing.Color.Black
        Me.cmdT1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT1.Location = New System.Drawing.Point(333, 4)
        Me.cmdT1.Name = "cmdT1"
        Me.cmdT1.Size = New System.Drawing.Size(38, 23)
        Me.cmdT1.TabIndex = 82
        Me.cmdT1.Text = "-1T"
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(270, 30)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 16)
        Me.Label5.TabIndex = 74
        Me.Label5.Text = "Miesiące . . . . . . . . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(270, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 16)
        Me.Label4.TabIndex = 72
        Me.Label4.Text = "Tygodnie . . . . . . . . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "Do daty . . . . . . . . . . . . . ."
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "Od daty . . . . . . . . . . . . . . . ."
        '
        'rabLiczbaWieksze
        '
        Me.rabLiczbaWieksze.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaWieksze.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaWieksze.Location = New System.Drawing.Point(429, 9)
        Me.rabLiczbaWieksze.Name = "rabLiczbaWieksze"
        Me.rabLiczbaWieksze.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaWieksze.TabIndex = 108
        Me.rabLiczbaWieksze.Text = ">"
        '
        'rabLiczbaMniesze
        '
        Me.rabLiczbaMniesze.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaMniesze.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaMniesze.Location = New System.Drawing.Point(379, 9)
        Me.rabLiczbaMniesze.Name = "rabLiczbaMniesze"
        Me.rabLiczbaMniesze.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaMniesze.TabIndex = 106
        Me.rabLiczbaMniesze.Text = "<"
        '
        'rabLiczbaRowne
        '
        Me.rabLiczbaRowne.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaRowne.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaRowne.Location = New System.Drawing.Point(333, 9)
        Me.rabLiczbaRowne.Name = "rabLiczbaRowne"
        Me.rabLiczbaRowne.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaRowne.TabIndex = 104
        Me.rabLiczbaRowne.Text = "="
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(270, 6)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(83, 16)
        Me.Label7.TabIndex = 111
        Me.Label7.Text = "Relacje . . . . . . . . . . . . . . . ."
        '
        'rabLiczbaOdDo
        '
        Me.rabLiczbaOdDo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaOdDo.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaOdDo.Location = New System.Drawing.Point(477, 25)
        Me.rabLiczbaOdDo.Name = "rabLiczbaOdDo"
        Me.rabLiczbaOdDo.Size = New System.Drawing.Size(73, 16)
        Me.rabLiczbaOdDo.TabIndex = 110
        Me.rabLiczbaOdDo.Text = "Od..Do"
        '
        'rabLiczbaWieRowne
        '
        Me.rabLiczbaWieRowne.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaWieRowne.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaWieRowne.Location = New System.Drawing.Point(429, 25)
        Me.rabLiczbaWieRowne.Name = "rabLiczbaWieRowne"
        Me.rabLiczbaWieRowne.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaWieRowne.TabIndex = 109
        Me.rabLiczbaWieRowne.Text = ">="
        '
        'rabLiczbaMniejRowne
        '
        Me.rabLiczbaMniejRowne.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaMniejRowne.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaMniejRowne.Location = New System.Drawing.Point(379, 25)
        Me.rabLiczbaMniejRowne.Name = "rabLiczbaMniejRowne"
        Me.rabLiczbaMniejRowne.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaMniejRowne.TabIndex = 107
        Me.rabLiczbaMniejRowne.Text = "<="
        '
        'rabLiczbaRozne
        '
        Me.rabLiczbaRozne.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabLiczbaRozne.ForeColor = System.Drawing.Color.Navy
        Me.rabLiczbaRozne.Location = New System.Drawing.Point(333, 25)
        Me.rabLiczbaRozne.Name = "rabLiczbaRozne"
        Me.rabLiczbaRozne.Size = New System.Drawing.Size(48, 16)
        Me.rabLiczbaRozne.TabIndex = 105
        Me.rabLiczbaRozne.Text = "<>"
        '
        'cbxIleKontaktowDo
        '
        Me.cbxIleKontaktowDo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxIleKontaktowDo.ForeColor = System.Drawing.Color.Purple
        Me.cbxIleKontaktowDo.Location = New System.Drawing.Point(118, 27)
        Me.cbxIleKontaktowDo.Name = "cbxIleKontaktowDo"
        Me.cbxIleKontaktowDo.Size = New System.Drawing.Size(112, 22)
        Me.cbxIleKontaktowDo.TabIndex = 102
        '
        'lblDo
        '
        Me.lblDo.ForeColor = System.Drawing.Color.Navy
        Me.lblDo.Location = New System.Drawing.Point(8, 30)
        Me.lblDo.Name = "lblDo"
        Me.lblDo.Size = New System.Drawing.Size(112, 16)
        Me.lblDo.TabIndex = 103
        Me.lblDo.Text = "Ilość kontaktów Do . . . . . . . . ."
        '
        'cmdAnalizuj
        '
        Me.cmdAnalizuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAnalizuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdAnalizuj.ForeColor = System.Drawing.Color.Black
        Me.cmdAnalizuj.Image = CType(resources.GetObject("cmdAnalizuj.Image"), System.Drawing.Image)
        Me.cmdAnalizuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAnalizuj.Location = New System.Drawing.Point(224, 479)
        Me.cmdAnalizuj.Name = "cmdAnalizuj"
        Me.cmdAnalizuj.Size = New System.Drawing.Size(355, 23)
        Me.cmdAnalizuj.TabIndex = 70
        Me.cmdAnalizuj.Text = "Analizuj częstość kontaktów z klientami w podanym zakresie dat"
        Me.cmdAnalizuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbxIleKontaktowOd
        '
        Me.cbxIleKontaktowOd.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxIleKontaktowOd.ForeColor = System.Drawing.Color.Purple
        Me.cbxIleKontaktowOd.Location = New System.Drawing.Point(118, 3)
        Me.cbxIleKontaktowOd.Name = "cbxIleKontaktowOd"
        Me.cbxIleKontaktowOd.Size = New System.Drawing.Size(112, 22)
        Me.cbxIleKontaktowOd.TabIndex = 68
        '
        'lblOd
        '
        Me.lblOd.ForeColor = System.Drawing.Color.Navy
        Me.lblOd.Location = New System.Drawing.Point(8, 6)
        Me.lblOd.Name = "lblOd"
        Me.lblOd.Size = New System.Drawing.Size(112, 16)
        Me.lblOd.TabIndex = 69
        Me.lblOd.Text = "Ilość kontaktów Od . . . . . . . . ."
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 478)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 65
        Me.PictureBox1.TabStop = False
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(678, 479)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 23)
        Me.cmdStop.TabIndex = 66
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'menuRaport2
        '
        Me.menuRaport2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuWybierzAM})
        Me.menuRaport2.Name = "menuRaport2"
        Me.menuRaport2.Size = New System.Drawing.Size(228, 26)
        '
        'mnuWybierzAM
        '
        Me.mnuWybierzAM.Name = "mnuWybierzAM"
        Me.mnuWybierzAM.Size = New System.Drawing.Size(227, 22)
        Me.mnuWybierzAM.Text = "Wybierz Akcję Marketingową"
        '
        'cmdZapamietaj
        '
        Me.cmdZapamietaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZapamietaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdZapamietaj.ForeColor = System.Drawing.Color.Black
        Me.cmdZapamietaj.Image = CType(resources.GetObject("cmdZapamietaj.Image"), System.Drawing.Image)
        Me.cmdZapamietaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZapamietaj.Location = New System.Drawing.Point(585, 479)
        Me.cmdZapamietaj.Name = "cmdZapamietaj"
        Me.cmdZapamietaj.Size = New System.Drawing.Size(87, 23)
        Me.cmdZapamietaj.TabIndex = 68
        Me.cmdZapamietaj.Text = "Pobierz"
        Me.cmdZapamietaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.rabLiczbaOdDo)
        Me.Panel2.Controls.Add(Me.rabLiczbaWieRowne)
        Me.Panel2.Controls.Add(Me.rabLiczbaWieksze)
        Me.Panel2.Controls.Add(Me.rabLiczbaMniejRowne)
        Me.Panel2.Controls.Add(Me.rabLiczbaMniesze)
        Me.Panel2.Controls.Add(Me.cbxIleKontaktowDo)
        Me.Panel2.Controls.Add(Me.cbxIleKontaktowOd)
        Me.Panel2.Controls.Add(Me.lblOd)
        Me.Panel2.Controls.Add(Me.rabLiczbaRowne)
        Me.Panel2.Controls.Add(Me.lblDo)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.rabLiczbaRozne)
        Me.Panel2.Location = New System.Drawing.Point(8, 65)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(758, 56)
        Me.Panel2.TabIndex = 69
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.rabRKLista)
        Me.Panel3.Controls.Add(Me.rabRKWszystkie)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.cmdRodzajeKontaktow)
        Me.Panel3.Controls.Add(Me.txtRodzajeKontaktow)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Location = New System.Drawing.Point(8, 120)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(758, 84)
        Me.Panel3.TabIndex = 71
        '
        'rabRKLista
        '
        Me.rabRKLista.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabRKLista.ForeColor = System.Drawing.Color.Navy
        Me.rabRKLista.Location = New System.Drawing.Point(333, 21)
        Me.rabRKLista.Name = "rabRKLista"
        Me.rabRKLista.Size = New System.Drawing.Size(202, 16)
        Me.rabRKLista.TabIndex = 114
        Me.rabRKLista.Text = "Tylko Rodzaje Kontaktów z listy"
        '
        'rabRKWszystkie
        '
        Me.rabRKWszystkie.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rabRKWszystkie.ForeColor = System.Drawing.Color.Navy
        Me.rabRKWszystkie.Location = New System.Drawing.Point(333, 5)
        Me.rabRKWszystkie.Name = "rabRKWszystkie"
        Me.rabRKWszystkie.Size = New System.Drawing.Size(202, 16)
        Me.rabRKWszystkie.TabIndex = 113
        Me.rabRKWszystkie.Text = "Wszystkie Rodzaje Kontaktów"
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(270, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 16)
        Me.Label3.TabIndex = 112
        Me.Label3.Text = "Wybieraj . . . . . . . . . . . . . . . ."
        '
        'cmdRodzajeKontaktow
        '
        Me.cmdRodzajeKontaktow.Image = CType(resources.GetObject("cmdRodzajeKontaktow.Image"), System.Drawing.Image)
        Me.cmdRodzajeKontaktow.Location = New System.Drawing.Point(232, 6)
        Me.cmdRodzajeKontaktow.Name = "cmdRodzajeKontaktow"
        Me.cmdRodzajeKontaktow.Size = New System.Drawing.Size(32, 22)
        Me.cmdRodzajeKontaktow.TabIndex = 105
        '
        'txtRodzajeKontaktow
        '
        Me.txtRodzajeKontaktow.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRodzajeKontaktow.ForeColor = System.Drawing.Color.Purple
        Me.txtRodzajeKontaktow.Location = New System.Drawing.Point(73, 6)
        Me.txtRodzajeKontaktow.Multiline = True
        Me.txtRodzajeKontaktow.Name = "txtRodzajeKontaktow"
        Me.txtRodzajeKontaktow.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtRodzajeKontaktow.Size = New System.Drawing.Size(157, 69)
        Me.txtRodzajeKontaktow.TabIndex = 104
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(8, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(71, 35)
        Me.Label6.TabIndex = 103
        Me.Label6.Text = "Rodzaje kontaktów :"
        '
        'stbRaport
        '
        Me.stbRaport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbRaport.Dock = System.Windows.Forms.DockStyle.None
        Me.stbRaport.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbRaport.Location = New System.Drawing.Point(8, 450)
        Me.stbRaport.Name = "stbRaport"
        Me.stbRaport.ShowPanels = True
        Me.stbRaport.Size = New System.Drawing.Size(758, 22)
        Me.stbRaport.TabIndex = 72
        Me.stbRaport.Text = "StatusBar1"
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(587, 210)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 183
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(608, 208)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 23)
        Me.cmdWszystko.TabIndex = 181
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 208)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(758, 22)
        Me.stbFiltry.TabIndex = 182
        Me.stbFiltry.Text = "StatusBar1"
        '
        'dagRaport
        '
        Me.dagRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagRaport.ContextMenuStrip = Me.menuRaport2
        Me.dagRaport.Location = New System.Drawing.Point(8, 230)
        Me.dagRaport.Name = "dagRaport"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagRaport.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagRaport.Size = New System.Drawing.Size(758, 219)
        Me.dagRaport.TabIndex = 180
        '
        'HorizScrollTimer
        '
        '
        'Timer1
        '
        '
        'frmAnalizaCzestKontaktowGrupowa
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(778, 509)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.dagRaport)
        Me.Controls.Add(Me.stbRaport)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.cmdZapamietaj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdAnalizuj)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAnalizaCzestKontaktowGrupowa"
        Me.Text = "Analiza częstości kontaktów z klientami"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuRaport2.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.dagRaport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdAnalizuj As System.Windows.Forms.Button
    Friend WithEvents cbxIleKontaktowOd As System.Windows.Forms.ComboBox
    Friend WithEvents lblOd As System.Windows.Forms.Label
    Friend WithEvents cmdM10 As System.Windows.Forms.Button
    Friend WithEvents cmdM9 As System.Windows.Forms.Button
    Friend WithEvents cmdM8 As System.Windows.Forms.Button
    Friend WithEvents cmdM7 As System.Windows.Forms.Button
    Friend WithEvents cmdM6 As System.Windows.Forms.Button
    Friend WithEvents cmdM5 As System.Windows.Forms.Button
    Friend WithEvents cmdM4 As System.Windows.Forms.Button
    Friend WithEvents cmdM3 As System.Windows.Forms.Button
    Friend WithEvents cmdM2 As System.Windows.Forms.Button
    Friend WithEvents cmdM1 As System.Windows.Forms.Button
    Friend WithEvents cmdT10 As System.Windows.Forms.Button
    Friend WithEvents cmdT9 As System.Windows.Forms.Button
    Friend WithEvents cmdT8 As System.Windows.Forms.Button
    Friend WithEvents cmdT7 As System.Windows.Forms.Button
    Friend WithEvents cmdT6 As System.Windows.Forms.Button
    Friend WithEvents cmdT5 As System.Windows.Forms.Button
    Friend WithEvents cmdT4 As System.Windows.Forms.Button
    Friend WithEvents cmdT3 As System.Windows.Forms.Button
    Friend WithEvents cmdT2 As System.Windows.Forms.Button
    Friend WithEvents cmdT1 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdZapamietaj As System.Windows.Forms.Button
    Friend WithEvents cbxIleKontaktowDo As System.Windows.Forms.ComboBox
    Friend WithEvents lblDo As System.Windows.Forms.Label
    Friend WithEvents rabLiczbaWieksze As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaMniesze As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaRowne As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents rabLiczbaOdDo As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaWieRowne As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaMniejRowne As System.Windows.Forms.RadioButton
    Friend WithEvents rabLiczbaRozne As System.Windows.Forms.RadioButton
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents rabRKLista As System.Windows.Forms.RadioButton
    Friend WithEvents rabRKWszystkie As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdRodzajeKontaktow As System.Windows.Forms.Button
    Friend WithEvents txtRodzajeKontaktow As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents stbRaport As System.Windows.Forms.StatusBar
    Friend WithEvents menuRaport2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuWybierzAM As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystko As Button
    Friend WithEvents stbFiltry As StatusBar
    Friend WithEvents dagRaport As DataGridView
    Friend WithEvents HorizScrollTimer As Timer
    Friend WithEvents dtpDoDaty As DateTimePicker
    Friend WithEvents dtpOdDaty As DateTimePicker
    Friend WithEvents Timer1 As Timer
End Class
