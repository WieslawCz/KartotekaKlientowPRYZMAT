Imports System.Drawing.Printing

Public Class frmDrukujRaportAktywnosci
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    'wywolanie z menu
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
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents cmdDrukuj As System.Windows.Forms.Button
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdZapiszExcel As System.Windows.Forms.Button
    Friend WithEvents dtpDoDaty As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
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
    Friend WithEvents dtpOdDaty As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDrukujRaportAktywnosci))
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.cmdDrukuj = New System.Windows.Forms.Button()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdM10 = New System.Windows.Forms.Button()
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
        Me.dtpDoDaty = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpOdDaty = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdZapiszExcel = New System.Windows.Forms.Button()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 81)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 23)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 84
        Me.picLogo.TabStop = False
        '
        'cmdDrukuj
        '
        Me.cmdDrukuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDrukuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDrukuj.Image = CType(resources.GetObject("cmdDrukuj.Image"), System.Drawing.Image)
        Me.cmdDrukuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDrukuj.Location = New System.Drawing.Point(631, 82)
        Me.cmdDrukuj.Name = "cmdDrukuj"
        Me.cmdDrukuj.Size = New System.Drawing.Size(70, 22)
        Me.cmdDrukuj.TabIndex = 83
        Me.cmdDrukuj.Text = "Raport"
        Me.cmdDrukuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(707, 82)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(80, 22)
        Me.cmdStop.TabIndex = 82
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdM10)
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
        Me.Panel1.Controls.Add(Me.dtpDoDaty)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.dtpOdDaty)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(784, 68)
        Me.Panel1.TabIndex = 81
        '
        'cmdM10
        '
        Me.cmdM10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM10.ForeColor = System.Drawing.Color.Black
        Me.cmdM10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM10.Location = New System.Drawing.Point(734, 32)
        Me.cmdM10.Name = "cmdM10"
        Me.cmdM10.Size = New System.Drawing.Size(40, 23)
        Me.cmdM10.TabIndex = 393
        Me.cmdM10.Text = "-10M"
        '
        'cmdM9
        '
        Me.cmdM9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM9.ForeColor = System.Drawing.Color.Black
        Me.cmdM9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM9.Location = New System.Drawing.Point(692, 32)
        Me.cmdM9.Name = "cmdM9"
        Me.cmdM9.Size = New System.Drawing.Size(38, 23)
        Me.cmdM9.TabIndex = 392
        Me.cmdM9.Text = "-9M"
        '
        'cmdM8
        '
        Me.cmdM8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM8.ForeColor = System.Drawing.Color.Black
        Me.cmdM8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM8.Location = New System.Drawing.Point(651, 32)
        Me.cmdM8.Name = "cmdM8"
        Me.cmdM8.Size = New System.Drawing.Size(38, 23)
        Me.cmdM8.TabIndex = 391
        Me.cmdM8.Text = "-8M"
        '
        'cmdM7
        '
        Me.cmdM7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM7.ForeColor = System.Drawing.Color.Black
        Me.cmdM7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM7.Location = New System.Drawing.Point(610, 32)
        Me.cmdM7.Name = "cmdM7"
        Me.cmdM7.Size = New System.Drawing.Size(38, 23)
        Me.cmdM7.TabIndex = 390
        Me.cmdM7.Text = "-7M"
        '
        'cmdM6
        '
        Me.cmdM6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM6.ForeColor = System.Drawing.Color.Black
        Me.cmdM6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM6.Location = New System.Drawing.Point(569, 32)
        Me.cmdM6.Name = "cmdM6"
        Me.cmdM6.Size = New System.Drawing.Size(38, 23)
        Me.cmdM6.TabIndex = 389
        Me.cmdM6.Text = "-6M"
        '
        'cmdM5
        '
        Me.cmdM5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM5.ForeColor = System.Drawing.Color.Black
        Me.cmdM5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM5.Location = New System.Drawing.Point(528, 32)
        Me.cmdM5.Name = "cmdM5"
        Me.cmdM5.Size = New System.Drawing.Size(38, 23)
        Me.cmdM5.TabIndex = 388
        Me.cmdM5.Text = "-5M"
        '
        'cmdM4
        '
        Me.cmdM4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM4.ForeColor = System.Drawing.Color.Black
        Me.cmdM4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM4.Location = New System.Drawing.Point(487, 32)
        Me.cmdM4.Name = "cmdM4"
        Me.cmdM4.Size = New System.Drawing.Size(38, 23)
        Me.cmdM4.TabIndex = 387
        Me.cmdM4.Text = "-4M"
        '
        'cmdM3
        '
        Me.cmdM3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM3.ForeColor = System.Drawing.Color.Black
        Me.cmdM3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM3.Location = New System.Drawing.Point(446, 32)
        Me.cmdM3.Name = "cmdM3"
        Me.cmdM3.Size = New System.Drawing.Size(38, 23)
        Me.cmdM3.TabIndex = 386
        Me.cmdM3.Text = "-3M"
        '
        'cmdM2
        '
        Me.cmdM2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM2.ForeColor = System.Drawing.Color.Black
        Me.cmdM2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM2.Location = New System.Drawing.Point(405, 32)
        Me.cmdM2.Name = "cmdM2"
        Me.cmdM2.Size = New System.Drawing.Size(38, 23)
        Me.cmdM2.TabIndex = 385
        Me.cmdM2.Text = "-2M"
        '
        'cmdM1
        '
        Me.cmdM1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdM1.ForeColor = System.Drawing.Color.Black
        Me.cmdM1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdM1.Location = New System.Drawing.Point(364, 32)
        Me.cmdM1.Name = "cmdM1"
        Me.cmdM1.Size = New System.Drawing.Size(38, 23)
        Me.cmdM1.TabIndex = 384
        Me.cmdM1.Text = "-1M"
        '
        'cmdT10
        '
        Me.cmdT10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT10.ForeColor = System.Drawing.Color.Black
        Me.cmdT10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT10.Location = New System.Drawing.Point(734, 10)
        Me.cmdT10.Name = "cmdT10"
        Me.cmdT10.Size = New System.Drawing.Size(40, 23)
        Me.cmdT10.TabIndex = 383
        Me.cmdT10.Text = "-10T"
        '
        'cmdT9
        '
        Me.cmdT9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT9.ForeColor = System.Drawing.Color.Black
        Me.cmdT9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT9.Location = New System.Drawing.Point(692, 10)
        Me.cmdT9.Name = "cmdT9"
        Me.cmdT9.Size = New System.Drawing.Size(38, 23)
        Me.cmdT9.TabIndex = 382
        Me.cmdT9.Text = "-9T"
        '
        'cmdT8
        '
        Me.cmdT8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT8.ForeColor = System.Drawing.Color.Black
        Me.cmdT8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT8.Location = New System.Drawing.Point(651, 10)
        Me.cmdT8.Name = "cmdT8"
        Me.cmdT8.Size = New System.Drawing.Size(38, 23)
        Me.cmdT8.TabIndex = 381
        Me.cmdT8.Text = "-8T"
        '
        'cmdT7
        '
        Me.cmdT7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT7.ForeColor = System.Drawing.Color.Black
        Me.cmdT7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT7.Location = New System.Drawing.Point(610, 10)
        Me.cmdT7.Name = "cmdT7"
        Me.cmdT7.Size = New System.Drawing.Size(38, 23)
        Me.cmdT7.TabIndex = 380
        Me.cmdT7.Text = "-7T"
        '
        'cmdT6
        '
        Me.cmdT6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT6.ForeColor = System.Drawing.Color.Black
        Me.cmdT6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT6.Location = New System.Drawing.Point(569, 10)
        Me.cmdT6.Name = "cmdT6"
        Me.cmdT6.Size = New System.Drawing.Size(38, 23)
        Me.cmdT6.TabIndex = 379
        Me.cmdT6.Text = "-6T"
        '
        'cmdT5
        '
        Me.cmdT5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT5.ForeColor = System.Drawing.Color.Black
        Me.cmdT5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT5.Location = New System.Drawing.Point(528, 10)
        Me.cmdT5.Name = "cmdT5"
        Me.cmdT5.Size = New System.Drawing.Size(38, 23)
        Me.cmdT5.TabIndex = 378
        Me.cmdT5.Text = "-5T"
        '
        'cmdT4
        '
        Me.cmdT4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT4.ForeColor = System.Drawing.Color.Black
        Me.cmdT4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT4.Location = New System.Drawing.Point(487, 10)
        Me.cmdT4.Name = "cmdT4"
        Me.cmdT4.Size = New System.Drawing.Size(38, 23)
        Me.cmdT4.TabIndex = 377
        Me.cmdT4.Text = "-4T"
        '
        'cmdT3
        '
        Me.cmdT3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT3.ForeColor = System.Drawing.Color.Black
        Me.cmdT3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT3.Location = New System.Drawing.Point(446, 10)
        Me.cmdT3.Name = "cmdT3"
        Me.cmdT3.Size = New System.Drawing.Size(38, 23)
        Me.cmdT3.TabIndex = 376
        Me.cmdT3.Text = "-3T"
        '
        'cmdT2
        '
        Me.cmdT2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT2.ForeColor = System.Drawing.Color.Black
        Me.cmdT2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT2.Location = New System.Drawing.Point(405, 10)
        Me.cmdT2.Name = "cmdT2"
        Me.cmdT2.Size = New System.Drawing.Size(38, 23)
        Me.cmdT2.TabIndex = 375
        Me.cmdT2.Text = "-2T"
        '
        'cmdT1
        '
        Me.cmdT1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdT1.ForeColor = System.Drawing.Color.Black
        Me.cmdT1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdT1.Location = New System.Drawing.Point(364, 10)
        Me.cmdT1.Name = "cmdT1"
        Me.cmdT1.Size = New System.Drawing.Size(38, 23)
        Me.cmdT1.TabIndex = 374
        Me.cmdT1.Text = "-1T"
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(301, 36)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 16)
        Me.Label5.TabIndex = 373
        Me.Label5.Text = "Miesi¹ce . . . . . . . . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(301, 14)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 16)
        Me.Label4.TabIndex = 372
        Me.Label4.Text = "Tygodnie . . . . . . . . . . . . . . . ."
        '
        'dtpDoDaty
        '
        Me.dtpDoDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDoDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpDoDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDoDaty.Location = New System.Drawing.Point(164, 32)
        Me.dtpDoDaty.Name = "dtpDoDaty"
        Me.dtpDoDaty.Size = New System.Drawing.Size(112, 20)
        Me.dtpDoDaty.TabIndex = 365
        Me.dtpDoDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(118, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(164, 16)
        Me.Label2.TabIndex = 364
        Me.Label2.Text = "Do dnia . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . ." &
    " . . . . "
        '
        'dtpOdDaty
        '
        Me.dtpOdDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpOdDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpOdDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOdDaty.Location = New System.Drawing.Point(164, 10)
        Me.dtpOdDaty.Name = "dtpOdDaty"
        Me.dtpOdDaty.Size = New System.Drawing.Size(112, 20)
        Me.dtpOdDaty.TabIndex = 363
        Me.dtpOdDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(260, 16)
        Me.Label1.TabIndex = 336
        Me.Label1.Text = "Analizowany okres - Od dnia . . . . . . . . . . . . . . . . . . . . . . . . . . ." &
    " . . . . . . . . . . . . . . "
        '
        'cmdZapiszExcel
        '
        Me.cmdZapiszExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZapiszExcel.Image = CType(resources.GetObject("cmdZapiszExcel.Image"), System.Drawing.Image)
        Me.cmdZapiszExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZapiszExcel.Location = New System.Drawing.Point(555, 82)
        Me.cmdZapiszExcel.Name = "cmdZapiszExcel"
        Me.cmdZapiszExcel.Size = New System.Drawing.Size(70, 22)
        Me.cmdZapiszExcel.TabIndex = 85
        Me.cmdZapiszExcel.Text = "Raport"
        Me.cmdZapiszExcel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmDrukujRaportAktywnosci
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(800, 109)
        Me.Controls.Add(Me.cmdZapiszExcel)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdDrukuj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frmDrukujRaportAktywnosci"
        Me.Text = "Raport Aktywnoœci Pracowników"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Dim StartFormy As Boolean = True
    Dim Tytul As String
    Dim LicznikRec As Long = 0

    '------------------------------------------------------------------
    'Kontakty Handlowe
    '------------------------------------------------------------------
    Dim dbSelectSQLKontakty As String = "SELECT * " & _
                               "FROM KontaktyHandlowe "

    Dim dbConnectionKontakty As OleDb.OleDbConnection
    Dim dbCommandSelectKontakty As OleDb.OleDbCommand
    Dim dbCommandDeleteKontakty As OleDb.OleDbCommand
    Dim dbCommandUpdateKontakty As OleDb.OleDbCommand
    Dim dbCommandInsertKontakty As OleDb.OleDbCommand
    Dim dbDataAdapterKontakty As OleDb.OleDbDataAdapter

    Dim sqlConnectionKontakty As SqlClient.SqlConnection
    Dim sqlCommandSelectKontakty As SqlClient.SqlCommand
    Dim sqlCommandDeleteKontakty As SqlClient.SqlCommand
    Dim sqlCommandUpdateKontakty As SqlClient.SqlCommand
    Dim sqlCommandInsertKontakty As SqlClient.SqlCommand
    Dim sqlDataAdapterKontakty As SqlClient.SqlDataAdapter

    Dim DataSetKontakty As DataSet
    Dim DataViewKontakty As DataView


    '------------------------------------------------------------------
    'Katalog Uzytkownicy
    'Public oIdentUzytkownika As String             'IDENT  Text    10
    'Public oNazwaUzytkownika As String             'NAZWA             Text    100
    'Public oFunkcjaUzytkownika As String           'FUNKCJA           Text    100
    'Public oHasloUzytkownika As String             'HASLO             Text    100
    'Public oUprawnieniaUzytkownika As String       'UPRAWNIENIA       Text    100
    'Public oUwagiUzytkownika As String   'UWAGI          Text
    'Public oWersjaUzytkownika As Integer 'WERSJA         Integer
    '------------------------------------------------------------------
    Dim dbSelectSQLUzytkownicy As String = sSelectSQLUzytkownicy

    Dim dbConnectionUzytkownicy As OleDb.OleDbConnection
    Dim dbCommandSelectUzytkownicy As OleDb.OleDbCommand
    Dim dbCommandDeleteUzytkownicy As OleDb.OleDbCommand
    Dim dbCommandUpdateUzytkownicy As OleDb.OleDbCommand
    Dim dbCommandInsertUzytkownicy As OleDb.OleDbCommand
    Dim dbDataAdapterUzytkownicy As OleDb.OleDbDataAdapter

    Dim sqlConnectionUzytkownicy As SqlClient.SqlConnection
    Dim sqlCommandSelectUzytkownicy As SqlClient.SqlCommand
    Dim sqlCommandDeleteUzytkownicy As SqlClient.SqlCommand
    Dim sqlCommandUpdateUzytkownicy As SqlClient.SqlCommand
    Dim sqlCommandInsertUzytkownicy As SqlClient.SqlCommand
    Dim sqlDataAdapterUzytkownicy As SqlClient.SqlDataAdapter

    Dim DataSetUzytkownicy As DataSet
    Dim DataViewUzytkownicy As DataView




    '------------------------------------------------------------------------
    'Katalog DaneDoRaportu
    'Public ddrPracownik As String        'PRACOWNIK   Text, 10
    'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
    'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
    'Public ddrOferty As Integer          'OFERTY         Integer
    'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
    'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
    'Public ddrDostawy As Integer         'DOSTAWY        Integer
    'Public ddrSkup As Integer            'SKUP           Integer

    'Public ddrKolExcel01 As String            'KolExCEL01           String
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
    '------------------------------------------------------------------
    Dim dbSelectSQLDaneDoRaportu As String = sSelectSQLDaneDoRaportu

    Dim dbConnectionDaneDoRaportu As OleDb.OleDbConnection
    Dim dbCommandSelectDaneDoRaportu As OleDb.OleDbCommand
    Dim dbCommandDeleteDaneDoRaportu As OleDb.OleDbCommand
    Dim dbCommandUpdateDaneDoRaportu As OleDb.OleDbCommand
    Dim dbCommandInsertDaneDoRaportu As OleDb.OleDbCommand
    Dim dbDataAdapterDaneDoRaportu As OleDb.OleDbDataAdapter

    Dim sqlConnectionDaneDoRaportu As SqlClient.SqlConnection
    Dim sqlCommandSelectDaneDoRaportu As SqlClient.SqlCommand
    Dim sqlCommandDeleteDaneDoRaportu As SqlClient.SqlCommand
    Dim sqlCommandUpdateDaneDoRaportu As SqlClient.SqlCommand
    Dim sqlCommandInsertDaneDoRaportu As SqlClient.SqlCommand
    Dim sqlDataAdapterDaneDoRaportu As SqlClient.SqlDataAdapter

    Dim DataSetDaneDoRaportu As DataSet
    Dim DataViewDaneDoRaportu As DataView
    '------------------------------------------------------------------------
    Dim ConnectionState As System.Data.ConnectionState



    Dim DataSetTotal As DataSet = Nothing
    Dim DataViewTotal As DataView = Nothing

    Dim DataSetRoboczy As DataSet = Nothing
    Dim DataViewRoboczy As DataView = Nothing
    'zakladamy tablice
    Dim TableRoboczy As System.Data.DataTable = Nothing
    Dim KeyKolumn1 As DataColumn = Nothing
    Dim KeyKolumn2 As DataColumn = Nothing
    Dim findRobo(1) As Object

    Dim DataSetObsluzeni As DataSet = Nothing
    Dim DataViewObsluzeni As DataView = Nothing
    'zakladamy tablice
    Dim TableObsluzeni As System.Data.DataTable = Nothing



    Dim pPracownik As String = ""
    Dim pNazwa As String = ""
    Dim pOpis As String = ""
    Dim pData As String = ""
    Dim pTelefon As Integer = 0
    Dim pWizytyO As Integer = 0
    Dim pWizytyProm As Integer = 0
    Dim pPozostale As Integer = 0
    Dim pKlBezDr As Integer = 0
    Dim pOferty As Integer = 0
    Dim pTesty As Integer = 0
    Dim pCzyszczenie As Integer = 0
    Dim pDostawy As Integer = 0
    Dim pSkup As Integer = 0
    Dim pRazemKon As Integer = 0
    Dim pKlLPoten As Integer = 0
    Dim pKlLPonow As Integer = 0
    Dim pKlAPoten As Integer = 0
    Dim pKlAPonow As Integer = 0
    Dim pNowiKl As Integer = 0

    Dim pExcel01 As Integer = 0
    Dim pExcel02 As Integer = 0
    Dim pExcel03 As Integer = 0
    Dim pExcel04 As Integer = 0
    Dim pExcel05 As Integer = 0
    Dim pExcel06 As Integer = 0
    Dim pExcel07 As Integer = 0
    Dim pExcel08 As Integer = 0
    Dim pExcel09 As Integer = 0
    Dim pExcel10 As Integer = 0

    Dim SumTelefon As Integer = 0
    Dim SumWizytyO As Integer = 0
    Dim SumWizytyProm As Integer = 0
    Dim SumPozostale As Integer = 0
    Dim SumKlBezDr As Integer = 0
    Dim SumOferty As Integer = 0
    Dim SumTesty As Integer = 0
    Dim SumCzyszczenie As Integer = 0
    Dim SumDostawy As Integer = 0
    Dim SumSkup As Integer = 0
    Dim SumRazemKon As Integer = 0
    Dim SumKlLPoten As Integer = 0
    Dim SumKlLPonow As Integer = 0
    Dim SumKlAPoten As Integer = 0
    Dim SumKlAPonow As Integer = 0
    Dim SumNowiKl As Integer = 0

    Dim SumExcel01 As Integer = 0
    Dim SumExcel02 As Integer = 0
    Dim SumExcel03 As Integer = 0
    Dim SumExcel04 As Integer = 0
    Dim SumExcel05 As Integer = 0
    Dim SumExcel06 As Integer = 0
    Dim SumExcel07 As Integer = 0
    Dim SumExcel08 As Integer = 0
    Dim SumExcel09 As Integer = 0
    Dim SumExcel10 As Integer = 0

    Dim TotTelefon As Integer = 0
    Dim TotWizytyO As Integer = 0
    Dim TotWizytyProm As Integer = 0
    Dim TotPozostale As Integer = 0
    Dim TotKlBezDr As Integer = 0
    Dim TotOferty As Integer = 0
    Dim TotTesty As Integer = 0
    Dim TotCzyszczenie As Integer = 0
    Dim TotDostawy As Integer = 0
    Dim TotSkup As Integer = 0
    Dim TotRazemKon As Integer = 0
    Dim TotKlLPoten As Integer = 0
    Dim TotKlLPonow As Integer = 0
    Dim TotKlAPoten As Integer = 0
    Dim TotKlAPonow As Integer = 0
    Dim TotNowiKl As Integer = 0

    Dim TotExcel01 As Integer = 0
    Dim TotExcel02 As Integer = 0
    Dim TotExcel03 As Integer = 0
    Dim TotExcel04 As Integer = 0
    Dim TotExcel05 As Integer = 0
    Dim TotExcel06 As Integer = 0
    Dim TotExcel07 As Integer = 0
    Dim TotExcel08 As Integer = 0
    Dim TotExcel09 As Integer = 0
    Dim TotExcel10 As Integer = 0


    Dim _CoMamRobic As String = ""
    Dim OdDaty As String = ""
    Dim DoDaty As String = ""
    Dim TakenDate As Date = Nothing
    Dim LiRow As Integer = 0
    Dim IleRecPrac As Integer = 0


    Private Sub DrukujZestawienie_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.picLogo.Image = My.Resources.logomini
        Me.Icon = My.Resources.MrJones
        '------------------------------------
        dtpOdDaty.Value = Mid(SysData(), 1, 7) & "-01"
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub frmDrukujRaportAktywnosci_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub




    '--------------------------------
    ' automatyczne ustalanie zakresu dat
    '--------------------------------

    Private Sub cmdT1_Click(sender As Object, e As EventArgs) Handles cmdT1.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -1 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT2_Click(sender As Object, e As EventArgs) Handles cmdT2.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -2 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT3_Click(sender As Object, e As EventArgs) Handles cmdT3.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -3 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT4_Click(sender As Object, e As EventArgs) Handles cmdT4.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -4 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT5_Click(sender As Object, e As EventArgs) Handles cmdT5.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -5 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT6_Click(sender As Object, e As EventArgs) Handles cmdT6.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -6 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT7_Click(sender As Object, e As EventArgs) Handles cmdT7.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -7 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT8_Click(sender As Object, e As EventArgs) Handles cmdT8.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -8 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT9_Click(sender As Object, e As EventArgs) Handles cmdT9.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -9 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT10_Click(sender As Object, e As EventArgs) Handles cmdT10.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -10 * 7)
        dtpDoDaty.Value = SysData()
    End Sub




    Private Sub cmdM1_Click(sender As Object, e As EventArgs) Handles cmdM1.Click
        dtpOdDaty.Value = PopMiesiac(SysData())
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM2_Click(sender As Object, e As EventArgs) Handles cmdM2.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(SysData()))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM3_Click(sender As Object, e As EventArgs) Handles cmdM3.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(SysData())))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM4_Click(sender As Object, e As EventArgs) Handles cmdM4.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData()))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM5_Click(sender As Object, e As EventArgs) Handles cmdM5.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData())))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM6_Click(sender As Object, e As EventArgs) Handles cmdM6.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData()))))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM7_Click(sender As Object, e As EventArgs) Handles cmdM7.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData())))))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM8_Click(sender As Object, e As EventArgs) Handles cmdM8.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData()))))))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM9_Click(sender As Object, e As EventArgs) Handles cmdM9.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData())))))))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM10_Click(sender As Object, e As EventArgs) Handles cmdM10.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData()))))))))))
        dtpDoDaty.Value = SysData()
    End Sub







    '-----------------------------------
    ' pobierz dane do zestawienia
    '-----------------------------------

    Private Function PrzygotujDane(ByVal pDodajTotal As Boolean) As Boolean
        Dim RetValue As Boolean = False
        Dim tx As String = ""
        Dim iU As Integer = 0

        Dim RoboOK As Boolean = False
        Dim UzytkownicyOK As Boolean = False
        Dim DaneDoRaportuOK As Boolean = False
        Dim KontaktyOK As Boolean = False

        TakenDate = dtpOdDaty.Value
        OdDaty = TakenDate.ToString("yyyy-MM-dd")
        TakenDate = dtpDoDaty.Value
        DoDaty = TakenDate.ToString("yyyy-MM-dd")


        '------------------------------------
        'przygotowujemy strukture na dane syntetyczne
        DataSetObsluzeni = New DataSet
        'zakladamy tablice
        TableObsluzeni = DataSetObsluzeni.Tables.Add("TabObsluzeni")
        'zakladamy kolumny
        KeyKolumn1 = New DataColumn
        KeyKolumn1 = TableObsluzeni.Columns.Add("IdKlienta", Type.GetType("System.String"))
        'wybierz klucz glowny...
        TableObsluzeni.PrimaryKey = New DataColumn() {KeyKolumn1}
        Dim DataViewObsluzeni As New DataView(DataSetObsluzeni.Tables(0))
        DataViewObsluzeni.Sort = "IdKlienta"




        '------------------------------------
        'Dim i As Integer = 0
        'If DataSetRoboczy.Tables(0).Rows.Count > 0 Then
        '    For i = DataSetRoboczy.Tables(0).Rows.Count To 0 Step -1
        '        DataSetRoboczy.Tables(0).Rows.Item(i).Delete()
        '    Next
        'End If
        '------------------------------------
        'przygotowujemy strukture na dane syntetyczne
        DataSetRoboczy = New DataSet
        'zakladamy tablice
        TableRoboczy = DataSetRoboczy.Tables.Add("Raport")
        'zakladamy kolumny
        'Pracownik|NazwaP|OpisP|Data|Telefony|WizytyO|WizytyProm|Pozostale|KlBezDr|Oferty|Testy|Czyszczenie|Dostawy|Skup|SumaKon|KlLPoten|KlLPonow|KlAPoten|KlAPonow|NowiKl

        KeyKolumn1 = New DataColumn
        KeyKolumn2 = New DataColumn

        KeyKolumn1 = TableRoboczy.Columns.Add("Pracownik", Type.GetType("System.String"))
        TableRoboczy.Columns.Item("Pracownik").MaxLength = 50
        TableRoboczy.Columns.Add("NazwaP", Type.GetType("System.String"))
        TableRoboczy.Columns.Item("NazwaP").MaxLength = 100
        TableRoboczy.Columns.Add("opisP", Type.GetType("System.String"))
        TableRoboczy.Columns.Item("opisP").MaxLength = 100
        KeyKolumn2 = TableRoboczy.Columns.Add("Data", Type.GetType("System.String"))
        TableRoboczy.Columns.Item("Data").MaxLength = 80
        TableRoboczy.Columns.Add("DzienTyg", Type.GetType("System.String"))
        TableRoboczy.Columns.Item("DzienTyg").MaxLength = 100
        TableRoboczy.Columns.Add("Telefony", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("WizytyO", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("WizytyProm", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Pozostale", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("KlBezDr", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Oferty", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Testy", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Czyszczenie", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Dostawy", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Skup", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("SumaKon", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("KlLPoten", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("KlLPonow", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("KlAPoten", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("KlAPonow", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("NowiKl", Type.GetType("System.Int32"))

        TableRoboczy.Columns.Add("Kolumna01", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna02", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna03", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna04", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna05", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna06", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna07", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna08", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna09", Type.GetType("System.Int32"))
        TableRoboczy.Columns.Add("Kolumna10", Type.GetType("System.Int32"))

        'wybierz klucz glowny...
        TableRoboczy.PrimaryKey = New DataColumn() {KeyKolumn1, KeyKolumn2}
        DataViewRoboczy = New DataView(DataSetRoboczy.Tables(0))
        DataViewRoboczy.Sort = "Pracownik,Data"
        '------------------------------------

        Dim drv As DataRowView = Nothing
        Dim dr As DataRow = Nothing
        Dim rowNo As Integer = 0

        TotTelefon = 0
        TotWizytyO = 0
        TotWizytyProm = 0
        TotPozostale = 0
        TotKlBezDr = 0
        TotOferty = 0
        TotTesty = 0
        TotCzyszczenie = 0
        TotDostawy = 0
        TotSkup = 0
        TotRazemKon = 0
        TotKlLPoten = 0
        TotKlLPonow = 0
        TotKlAPoten = 0
        TotKlAPonow = 0
        TotNowiKl = 0

        TotExcel01 = 0
        TotExcel02 = 0
        TotExcel03 = 0
        TotExcel04 = 0
        TotExcel05 = 0
        TotExcel06 = 0
        TotExcel07 = 0
        TotExcel08 = 0
        TotExcel09 = 0
        TotExcel10 = 0

        RoboOK = True


        '-------------------------------
        dbSelectSQLUzytkownicy = sSelectSQLUzytkownicy

        DataSetUzytkownicy = New DataSet
        DataViewUzytkownicy = New DataView

        DataSetUzytkownicy.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionUzytkownicy = New OleDb.OleDbConnection(sConnectionUzytkownicy)
            dbCommandSelectUzytkownicy = New OleDb.OleDbCommand(dbSelectSQLUzytkownicy, dbConnectionUzytkownicy)
            'dbCommandDeleteUzytkownicy = New OleDb.OleDbCommand(sDeleteSQLUzytkownicy, dbConnectionUzytkownicy)
            'dbCommandUpdateUzytkownicy = New OleDb.OleDbCommand(sUpdateSQLUzytkownicy, dbConnectionUzytkownicy)
            'dbCommandInsertUzytkownicy = New OleDb.OleDbCommand(sInsertSQLUzytkownicy, dbConnectionUzytkownicy)
            dbDataAdapterUzytkownicy = New OleDb.OleDbDataAdapter(dbCommandSelectUzytkownicy)

            Dim DBMapowanieTabeliUzytkownicy As System.Data.Common.DataTableMapping
            DBMapowanieTabeliUzytkownicy = dbDataAdapterUzytkownicy.TableMappings.Add("Table", "TABELA_Uzytkownicy")

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionUzytkownicy.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnectionUzytkownicy.State
            End Try
        Else
            sqlConnectionUzytkownicy = New SqlClient.SqlConnection(sConnectionUzytkownicy)
            sqlCommandSelectUzytkownicy = New SqlClient.SqlCommand(dbSelectSQLUzytkownicy, sqlConnectionUzytkownicy)
            'sqlCommandDeleteUzytkownicy = New sqlclient.sqlCommand(sDeleteSQLUzytkownicy, sqlconnectionUzytkownicy)
            'sqlCommandUpdateUzytkownicy = New sqlclient.sqlCommand(sUpdateSQLUzytkownicy, sqlconnectionUzytkownicy)
            'sqlCommandInsertUzytkownicy = New sqlclient.sqlCommand(sInsertSQLUzytkownicy, sqlconnectionUzytkownicy)
            sqlDataAdapterUzytkownicy = New SqlClient.SqlDataAdapter(sqlCommandSelectUzytkownicy)

            Dim sqlMapowanieTabeliUzytkownicy As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliUzytkownicy = sqlDataAdapterUzytkownicy.TableMappings.Add("Table", "TABELA_Uzytkownicy")

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionUzytkownicy.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionUzytkownicy.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterUzytkownicy.FillSchema(DataSetUzytkownicy, SchemaType.Mapped, "TABELA_Uzytkownicy")
                    dbDataAdapterUzytkownicy.Fill(DataSetUzytkownicy)
                    dbConnectionUzytkownicy.Close()
                Else
                    sqlDataAdapterUzytkownicy.FillSchema(DataSetUzytkownicy, SchemaType.Mapped, "TABELA_Uzytkownicy")
                    sqlDataAdapterUzytkownicy.Fill(DataSetUzytkownicy)
                    sqlConnectionUzytkownicy.Close()
                End If
                'DataSetUzytkownicy.Tables(0).PrimaryKey = New DataColumn() _
                '            {DataSetUzytkownicy.Tables(0).Columns("IDENTTRA")}
                DataViewUzytkownicy = New DataView(DataSetUzytkownicy.Tables(0))
                DataViewUzytkownicy.AllowDelete = False
                DataViewUzytkownicy.AllowNew = False

                UzytkownicyOK = True

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)

            End Try
        End If



        '-------------------------------
        If RoboOK And UzytkownicyOK Then
            dbSelectSQLDaneDoRaportu = sSelectSQLDaneDoRaportu & _
                                        " WHERE (DATARAPORTU>='" & OdDaty & "') AND (DATARAPORTU<='" & DoDaty & "') " & _
                                        " ORDER BY DATARAPORTU"


            DataSetDaneDoRaportu = New DataSet
            DataViewDaneDoRaportu = New DataView

            DataSetDaneDoRaportu.Locale = New System.Globalization.CultureInfo("pl-PL")

            If ParL_DataBaseType = "MS ACCESS" Then
                dbConnectionDaneDoRaportu = New OleDb.OleDbConnection(sConnectionDaneDoRaportu)
                dbCommandSelectDaneDoRaportu = New OleDb.OleDbCommand(dbSelectSQLDaneDoRaportu, dbConnectionDaneDoRaportu)
                'dbCommandDeleteDaneDoRaportu = New OleDb.OleDbCommand(sDeleteSQLDaneDoRaportu, dbConnectionDaneDoRaportu)
                'dbCommandUpdateDaneDoRaportu = New OleDb.OleDbCommand(sUpdateSQLDaneDoRaportu, dbConnectionDaneDoRaportu)
                'dbCommandInsertDaneDoRaportu = New OleDb.OleDbCommand(sInsertSQLDaneDoRaportu, dbConnectionDaneDoRaportu)
                dbDataAdapterDaneDoRaportu = New OleDb.OleDbDataAdapter(dbCommandSelectDaneDoRaportu)

                Dim DBMapowanieTabeliDaneDoRaportu As System.Data.Common.DataTableMapping
                DBMapowanieTabeliDaneDoRaportu = dbDataAdapterDaneDoRaportu.TableMappings.Add("Table", "TABELA_DaneDoRaportu")

                '------------------------------------------
                'wypelnij DATASET
                Try
                    dbConnectionDaneDoRaportu.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = dbConnectionDaneDoRaportu.State
                End Try
            Else
                sqlConnectionDaneDoRaportu = New SqlClient.SqlConnection(sConnectionDaneDoRaportu)
                sqlCommandSelectDaneDoRaportu = New SqlClient.SqlCommand(dbSelectSQLDaneDoRaportu, sqlConnectionDaneDoRaportu)
                'sqlCommandDeleteDaneDoRaportu = New sqlclient.sqlCommand(sDeleteSQLDaneDoRaportu, sqlconnectionDaneDoRaportu)
                'sqlCommandUpdateDaneDoRaportu = New sqlclient.sqlCommand(sUpdateSQLDaneDoRaportu, sqlconnectionDaneDoRaportu)
                'sqlCommandInsertDaneDoRaportu = New sqlclient.sqlCommand(sInsertSQLDaneDoRaportu, sqlconnectionDaneDoRaportu)
                sqlDataAdapterDaneDoRaportu = New SqlClient.SqlDataAdapter(sqlCommandSelectDaneDoRaportu)

                Dim sqlMapowanieTabeliDaneDoRaportu As System.Data.Common.DataTableMapping
                sqlMapowanieTabeliDaneDoRaportu = sqlDataAdapterDaneDoRaportu.TableMappings.Add("Table", "TABELA_DaneDoRaportu")

                '------------------------------------------
                'wypelnij DATASET
                Try
                    sqlConnectionDaneDoRaportu.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = sqlConnectionDaneDoRaportu.State
                End Try
            End If

            If ConnectionState = ConnectionState.Open Then
                Try
                    If ParL_DataBaseType = "MS ACCESS" Then
                        dbDataAdapterDaneDoRaportu.FillSchema(DataSetDaneDoRaportu, SchemaType.Mapped, "TABELA_DaneDoRaportu")
                        dbDataAdapterDaneDoRaportu.Fill(DataSetDaneDoRaportu)
                        dbConnectionDaneDoRaportu.Close()
                    Else
                        sqlDataAdapterDaneDoRaportu.FillSchema(DataSetDaneDoRaportu, SchemaType.Mapped, "TABELA_DaneDoRaportu")
                        sqlDataAdapterDaneDoRaportu.Fill(DataSetDaneDoRaportu)
                        sqlConnectionDaneDoRaportu.Close()
                    End If

                    'DataSetDaneDoRaportu.Tables(0).PrimaryKey = New DataColumn() _
                    '            {DataSetDaneDoRaportu.Tables(0).Columns("IDENTTRA")}

                    DataViewDaneDoRaportu = New DataView(DataSetDaneDoRaportu.Tables(0))
                    DataViewDaneDoRaportu.AllowDelete = False
                    DataViewDaneDoRaportu.AllowNew = False
                    DataViewDaneDoRaportu.Sort = "PRACOWNIK, DATARAPORTU"

                    DaneDoRaportuOK = True

                Catch Ex As System.Exception
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)

                End Try
            End If


            If RoboOK And UzytkownicyOK And DaneDoRaportuOK Then
                dbSelectSQLKontakty = "SELECT " & _
                                                         "IDENTKLIENTA, " & _
                                                         "DATAKONTAKTU, " & _
                                                         "PRACOWNIK, " & _
                                                         "TEMAT AS RODZAJ, " & _
                                                         "UWAGI AS OPIS, " & _
                                                         "WERSJA " & _
                                        "FROM KontaktyHandlowe " & _
                                        " WHERE (DATAKONTAKTU>='" & OdDaty & "') AND (DATAKONTAKTU<='" & DoDaty & "') " & _
                                        " ORDER BY DATAKONTAKTU"

                DataSetKontakty = New DataSet
                DataViewKontakty = New DataView
                DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbConnectionKontakty = New OleDb.OleDbConnection(sConnectionKontakty)
                    dbCommandSelectKontakty = New OleDb.OleDbCommand(dbSelectSQLKontakty, dbConnectionKontakty)
                    'dbCommandDeleteKontakty = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnectionKontakty)
                    'dbCommandUpdateKontakty = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnectionKontakty)
                    'dbCommandInsertKontakty = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnectionKontakty)
                    dbDataAdapterKontakty = New OleDb.OleDbDataAdapter(dbCommandSelectKontakty)

                    Dim DBMapowanieTabeliKontakty As System.Data.Common.DataTableMapping
                    DBMapowanieTabeliKontakty = dbDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

                    '------------------------------------------
                    'wypelnij DATASET
                    Try
                        dbConnectionKontakty.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    Finally
                        ConnectionState = dbConnectionKontakty.State
                    End Try
                Else
                    sqlConnectionKontakty = New SqlClient.SqlConnection(sConnectionKontakty)
                    sqlCommandSelectKontakty = New SqlClient.SqlCommand(dbSelectSQLKontakty, sqlConnectionKontakty)
                    'sqlCommandDeleteKontakty = New sqlclient.sqlCommand(sDeleteSQLKontakty, sqlconnectionKontakty)
                    'sqlCommandUpdateKontakty = New sqlclient.sqlCommand(sUpdateSQLKontakty, sqlconnectionKontakty)
                    'sqlCommandInsertKontakty = New sqlclient.sqlCommand(sInsertSQLKontakty, sqlconnectionKontakty)
                    sqlDataAdapterKontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectKontakty)

                    Dim sqlMapowanieTabeliKontakty As System.Data.Common.DataTableMapping
                    sqlMapowanieTabeliKontakty = sqlDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

                    '------------------------------------------
                    'wypelnij DATASET
                    Try
                        sqlConnectionKontakty.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    Finally
                        ConnectionState = sqlConnectionKontakty.State
                    End Try
                End If

                If ConnectionState = ConnectionState.Open Then
                    Try
                        If ParL_DataBaseType = "MS ACCESS" Then
                            dbDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                            dbDataAdapterKontakty.Fill(DataSetKontakty)
                            dbConnectionKontakty.Close()
                        Else
                            sqlDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                            sqlDataAdapterKontakty.Fill(DataSetKontakty)
                            sqlConnectionKontakty.Close()
                        End If
                        'DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() _
                        '            {DataSetKontakty.Tables(0).Columns("IDENTTRA")}
                        DataViewKontakty = New DataView(DataSetKontakty.Tables(0))
                        DataViewKontakty.AllowDelete = False
                        DataViewKontakty.AllowNew = False

                        KontaktyOK = True


                    Catch Ex As System.Exception
                        MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                            System.Windows.Forms.MessageBoxButtons.OK, _
                            MessageBoxIcon.Information, _
                            MessageBoxDefaultButton.Button1)

                    End Try
                End If




                If RoboOK And UzytkownicyOK And DaneDoRaportuOK And KontaktyOK Then
                    'ANALIZUJEMY WSZYSTKICH UZYTKOWNIKOW PO KOLEI
                    If DataViewUzytkownicy.Count > 0 Then
                        For iU = 0 To DataViewUzytkownicy.Count - 1
                            'Katalog Uzytkownicy
                            'Public oIdentUzytkownika As String             'IDENT  Text    10
                            'Public oNazwaUzytkownika As String             'NAZWA             Text    100
                            'Public oFunkcjaUzytkownika As String           'FUNKCJA           Text    100
                            'Public oHasloUzytkownika As String             'HASLO             Text    100
                            'Public oUprawnieniaUzytkownika As String       'UPRAWNIENIA       Text    100
                            'Public oUwagiUzytkownika As String   'UWAGI          Text
                            'Public oWersjaUzytkownika As Integer 'WERSJA         Integer

                            pPracownik = GetTxtDRVField(DataViewUzytkownicy.Item(iU), "IDENT")
                            pNazwa = GetTxtDRVField(DataViewUzytkownicy.Item(iU), "NAZWA")
                            pOpis = GetTxtDRVField(DataViewUzytkownicy.Item(iU), "FUNKCJA")

                            SumTelefon = 0
                            SumWizytyO = 0
                            SumWizytyProm = 0
                            SumPozostale = 0
                            SumKlBezDr = 0
                            SumOferty = 0
                            SumTesty = 0
                            SumCzyszczenie = 0
                            SumDostawy = 0
                            SumSkup = 0
                            SumRazemKon = 0
                            SumKlLPoten = 0
                            SumKlLPonow = 0
                            SumKlAPoten = 0
                            SumKlAPonow = 0
                            SumNowiKl = 0

                            SumExcel01 = 0
                            SumExcel02 = 0
                            SumExcel03 = 0
                            SumExcel04 = 0
                            SumExcel05 = 0
                            SumExcel06 = 0
                            SumExcel07 = 0
                            SumExcel08 = 0
                            SumExcel09 = 0
                            SumExcel10 = 0

                            IleRecPrac = 0
                            '---------------------------
                            'ANALIZUJEMY KONTAKTY TEGO PRACOWNIKA W WYBRANYM OKRESIE
                            DataViewKontakty.RowFilter = "PRACOWNIK='" & pPracownik & "'"
                            If DataViewKontakty.Count > 0 Then
                                For LiRow = 0 To DataViewKontakty.Count - 1
                                    PobierzDaneZKontaktow(LiRow)

                                    'oIdentKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "IDENTKLIENTA")
                                    'oDataKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "DATAKONTAKTU")
                                    'oPracownikKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "PRACOWNIK")
                                    'oTematKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "RODZAJ")
                                    'oUwagiKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "OPIS")
                                    'oWersjaKon = GetIntDRVField(DataViewKontakty.Item(LiRow), "WERSJA")
                                    '--------------------------------
                                    IdentKlienta(oIdentKon)
                                    '--------------------------------
                                    findRobo(0) = pPracownik
                                    findRobo(1) = oDataKon
                                    rowNo = DataViewRoboczy.Find(findRobo)
                                    If rowNo < 0 Then
                                        ' wypelnij dataset informacjami o programie...
                                        dr = DataSetRoboczy.Tables(0).NewRow()
                                        'Pracownik|NazwaP|OpisP|Data|Telefony|WizytyO|WizytyProm|Pozostale|KlBezDr|Oferty|Testy|Czyszczenie|Dostawy|Skup|SumaKon|KlLPoten|KlLPonow|KlAPoten|KlAPonow|NowiKl
                                        dr("Pracownik") = pPracownik
                                        dr("NazwaP") = pNazwa
                                        dr("Opisp") = pOpis
                                        dr("Data") = oDataKon
                                        dr("DzienTyg") = DzienTygodnia(oDataKon)

                                        dr("Telefony") = 0
                                        dr("WizytyO") = 0
                                        dr("WizytyProm") = 0
                                        dr("Pozostale") = 0

                                        dr("KlBezDr") = 0
                                        dr("Oferty") = 0
                                        dr("Testy") = 0
                                        dr("Czyszczenie") = 0
                                        dr("Dostawy") = 0
                                        dr("Skup") = 0

                                        dr("SumaKon") = 0

                                        dr("KlLPoten") = 0
                                        dr("KlLPonow") = 0
                                        dr("KlAPoten") = 0
                                        dr("KlAPonow") = 0
                                        dr("NowiKl") = 0

                                        dr("Kolumna01") = 0
                                        dr("Kolumna02") = 0
                                        dr("Kolumna03") = 0
                                        dr("Kolumna04") = 0
                                        dr("Kolumna05") = 0
                                        dr("Kolumna06") = 0
                                        dr("Kolumna07") = 0
                                        dr("Kolumna08") = 0
                                        dr("Kolumna09") = 0
                                        dr("Kolumna10") = 0

                                        DataSetRoboczy.Tables(0).Rows.Add(dr)
                                        IleRecPrac += 1

                                        rowNo = DataViewRoboczy.Find(findRobo)
                                    End If
                                    drv = DataViewRoboczy.Item(rowNo)

                                    pTelefon = 0
                                    pWizytyO = 0
                                    pWizytyProm = 0
                                    pPozostale = 0
                                    pKlBezDr = 0
                                    pOferty = 0
                                    pTesty = 0
                                    pCzyszczenie = 0
                                    pDostawy = 0
                                    pSkup = 0
                                    pRazemKon = 0
                                    pKlLPoten = 0
                                    pKlLPonow = 0
                                    pKlAPoten = 0
                                    pKlAPonow = 0
                                    pNowiKl = 0

                                    pExcel01 = 0
                                    pExcel02 = 0
                                    pExcel03 = 0
                                    pExcel04 = 0
                                    pExcel05 = 0
                                    pExcel06 = 0
                                    pExcel07 = 0
                                    pExcel08 = 0
                                    pExcel09 = 0
                                    pExcel10 = 0

                                    'rodzaje wizyt
                                    If InStr(UCase(oTematKon), "TEL") > 0 Then
                                        'Telefony
                                        pTelefon += 1
                                        pRazemKon += 1

                                    ElseIf InStr(UCase(oTematKon), "WIZ") > 0 Then
                                        'Wizyty O
                                        pWizytyO += 1
                                        pRazemKon += 1

                                    ElseIf InStr(UCase(oTematKon), "PROM") > 0 Then
                                        'Wizyty O
                                        pWizytyProm += 1
                                        pRazemKon += 1

                                    Else
                                        'Pozostale
                                        pPozostale += 1
                                        pRazemKon += 1

                                    End If
                                    '----------------------------
                                    'Sprawdz czy klient byl juz obslugiwany w analizowanych kontaktach
                                    rowNo = DataViewObsluzeni.Find(oIdentKon)
                                    If rowNo < 0 Then
                                        'NIE BYL -  dopisz do listy
                                        dr = DataSetObsluzeni.Tables(0).NewRow()
                                        dr("IdKlienta") = oIdentKon
                                        DataSetObsluzeni.Tables(0).Rows.Add(dr)

                                        If Len((oNrTOFITxtKli)) = 0 And (oRodzSprzetuLKli = True) Then
                                            pKlLPoten += 1
                                        ElseIf Len((oNrTOFITxtKli)) > 0 And (oRodzSprzetuLKli = True) Then
                                            pKlLPonow += 1
                                        ElseIf Len((oNrTOFITxtKli)) = 0 And (oRodzSprzetuAKli = True) Then
                                            pKlAPoten += 1
                                        ElseIf Len((oNrTOFITxtKli)) > 0 And (oRodzSprzetuAKli = True) Then
                                            pKlAPonow += 1
                                        Else
                                        End If
                                    End If
                                    '----------------------------
                                    'analizuj czy to nie jest NOWY KLIENT
                                    If oPKontaktKli = oDataKon Then
                                        pNowiKl += 1
                                    End If
                                    '-------------------------
                                    'aktualizacja zmiennych
                                    drv("Telefony") += pTelefon
                                    drv("WizytyO") += pWizytyO
                                    drv("WizytyProm") += pWizytyProm
                                    drv("Pozostale") += pPozostale
                                    drv("KlBezdr") += pKlBezDr
                                    drv("Oferty") += pOferty
                                    drv("Testy") += pTesty
                                    drv("Czyszczenie") += pCzyszczenie
                                    drv("Dostawy") += pDostawy
                                    drv("Skup") += pSkup
                                    drv("SumaKon") += pRazemKon
                                    drv("KlLPoten") += pKlLPoten
                                    drv("KlLPonow") += pKlLPonow
                                    drv("KlAPoten") += pKlAPoten
                                    drv("KlAPonow") += pKlAPonow
                                    drv("NowiKl") += pNowiKl
                                    drv("Kolumna01") += pExcel01
                                    drv("Kolumna02") += pExcel02
                                    drv("Kolumna03") += pExcel03
                                    drv("Kolumna04") += pExcel04
                                    drv("Kolumna05") += pExcel05
                                    drv("Kolumna06") += pExcel06
                                    drv("Kolumna07") += pExcel07
                                    drv("Kolumna08") += pExcel08
                                    drv("Kolumna09") += pExcel09
                                    drv("Kolumna10") += pExcel10

                                    SumTelefon += pTelefon
                                    SumWizytyO += pWizytyO
                                    SumWizytyProm += pWizytyProm
                                    SumPozostale += pPozostale
                                    SumKlBezDr += pKlBezDr
                                    SumOferty += pOferty
                                    SumTesty += pTesty
                                    SumCzyszczenie += pCzyszczenie
                                    SumDostawy += pDostawy
                                    SumSkup += pSkup
                                    SumRazemKon += pRazemKon
                                    SumKlLPoten += pKlLPoten
                                    SumKlLPonow += pKlLPonow
                                    SumKlAPoten += pKlAPoten
                                    SumKlAPonow += pKlAPonow
                                    SumNowiKl += pNowiKl

                                    SumExcel01 += pExcel01
                                    SumExcel02 += pExcel02
                                    SumExcel03 += pExcel03
                                    SumExcel04 += pExcel04
                                    SumExcel05 += pExcel05
                                    SumExcel06 += pExcel06
                                    SumExcel07 += pExcel07
                                    SumExcel08 += pExcel08
                                    SumExcel09 += pExcel09
                                    SumExcel10 += pExcel10

                                    TotTelefon += pTelefon
                                    TotWizytyO += pWizytyO
                                    TotWizytyProm += pWizytyProm
                                    TotPozostale += pPozostale
                                    TotKlBezDr += pKlBezDr
                                    TotOferty += pOferty
                                    TotTesty += pTesty
                                    TotCzyszczenie += pCzyszczenie
                                    TotDostawy += pDostawy
                                    TotSkup += pSkup
                                    TotRazemKon += pRazemKon
                                    TotKlLPoten += pKlLPoten
                                    TotKlLPonow += pKlLPonow
                                    TotKlAPoten += pKlAPoten
                                    TotKlAPonow += pKlAPonow
                                    TotNowiKl += pNowiKl

                                    TotExcel01 += pExcel01
                                    TotExcel02 += pExcel02
                                    TotExcel03 += pExcel03
                                    TotExcel04 += pExcel04
                                    TotExcel05 += pExcel05
                                    TotExcel06 += pExcel06
                                    TotExcel07 += pExcel07
                                    TotExcel08 += pExcel08
                                    TotExcel09 += pExcel09
                                    TotExcel10 += pExcel10
                                Next
                            End If




                            '---------------------------
                            'ANALIZUJEMY INFORMACJE DODATKOWE
                            'Public ddrPracownik As String        'PRACOWNIK   Text, 10
                            'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
                            'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
                            'Public ddrOferty As Integer          'OFERTY         Integer
                            'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
                            'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
                            'Public ddrDostawy As Integer         'DOSTAWY        Integer
                            'Public ddrSkup As Integer            'SKUP           Integer

                            'Public ddrKolExcel01 As String            'KolExCEL01           String
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

                            DataViewDaneDoRaportu.RowFilter = "PRACOWNIK='" & pPracownik & "'"
                            If DataViewDaneDoRaportu.Count > 0 Then
                                For LiRow = 0 To DataViewDaneDoRaportu.Count - 1
                                    ddrPracownik = GetTxtDRVField(DataViewDaneDoRaportu.Item(LiRow), "PRACOWNIK")
                                    ddrDataRaportu = GetTxtDRVField(DataViewDaneDoRaportu.Item(LiRow), "DATARAPORTU")
                                    ddrKlBezDrukarki = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "KLBEZDRUKARKI")
                                    ddrOferty = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "OFERTY")
                                    ddrTestyWstawionw = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "TESTYWSTAWIONE")
                                    ddrCzyszczenie = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "CZYSZCZENIE")
                                    ddrDostawy = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "DOSTAWY")
                                    ddrSkup = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "SKUP")

                                    ddrExcel01 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL01")
                                    ddrExcel02 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL02")
                                    ddrExcel03 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL03")
                                    ddrExcel04 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL04")
                                    ddrExcel05 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL05")
                                    ddrExcel06 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL06")
                                    ddrExcel07 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL07")
                                    ddrExcel08 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL08")
                                    ddrExcel09 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL09")
                                    ddrExcel10 = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "EXCEL10")

                                    ddrWersja = GetIntDRVField(DataViewDaneDoRaportu.Item(LiRow), "WERSJA")
                                    '--------------------------------
                                    findRobo(0) = ddrPracownik
                                    findRobo(1) = ddrDataRaportu
                                    rowNo = DataViewRoboczy.Find(findRobo)
                                    If rowNo < 0 Then
                                        ' wypelnij dataset informacjami o programie...
                                        dr = DataSetRoboczy.Tables(0).NewRow()
                                        'Pracownik|NazwaP|OpisP|Data|Telefony|WizytyO|WizytyPROM|Pozostale|KlBezDr|Oferty|Testy|Czyszczenie|Dostawy|Skup|SumaKon|KlLPoten|KlLPonow|KlAPoten|KlAPonow|NowiKl
                                        dr("Pracownik") = pPracownik
                                        dr("NazwaP") = pNazwa
                                        dr("Opisp") = pOpis
                                        dr("Data") = ddrDataRaportu
                                        dr("DzienTyg") = DzienTygodnia(ddrDataRaportu)

                                        dr("Telefony") = 0
                                        dr("WizytyO") = 0
                                        dr("WizytyProm") = 0
                                        dr("Pozostale") = 0

                                        dr("KlBezDr") = 0
                                        dr("Oferty") = 0
                                        dr("Testy") = 0
                                        dr("Czyszczenie") = 0
                                        dr("Dostawy") = 0
                                        dr("Skup") = 0

                                        dr("SumaKon") = 0

                                        dr("KlLPoten") = 0
                                        dr("KlLPonow") = 0
                                        dr("KlAPoten") = 0
                                        dr("KlAPonow") = 0
                                        dr("NowiKl") = 0

                                        dr("Kolumna01") = 0
                                        dr("Kolumna02") = 0
                                        dr("Kolumna03") = 0
                                        dr("Kolumna04") = 0
                                        dr("Kolumna05") = 0
                                        dr("Kolumna06") = 0
                                        dr("Kolumna07") = 0
                                        dr("Kolumna08") = 0
                                        dr("Kolumna09") = 0
                                        dr("Kolumna10") = 0

                                        DataSetRoboczy.Tables(0).Rows.Add(dr)
                                        IleRecPrac += 1

                                        rowNo = DataViewRoboczy.Find(findRobo)
                                    End If
                                    drv = DataViewRoboczy.Item(rowNo)

                                    pTelefon = 0
                                    pWizytyO = 0
                                    pWizytyProm = 0
                                    pPozostale = 0
                                    pKlBezDr = ddrKlBezDrukarki
                                    pOferty = ddrOferty
                                    pTesty = ddrTestyWstawionw
                                    pCzyszczenie = ddrCzyszczenie
                                    pDostawy = ddrDostawy
                                    pSkup = ddrSkup
                                    pRazemKon = 0
                                    pKlLPoten = 0
                                    pKlLPonow = 0
                                    pKlAPoten = 0
                                    pKlAPonow = 0
                                    pNowiKl = 0

                                    pExcel01 = ddrExcel01
                                    pExcel02 = ddrExcel02
                                    pExcel03 = ddrExcel03
                                    pExcel04 = ddrExcel04
                                    pExcel05 = ddrExcel05
                                    pExcel06 = ddrExcel06
                                    pExcel07 = ddrExcel07
                                    pExcel08 = ddrExcel08
                                    pExcel09 = ddrExcel09
                                    pExcel10 = ddrExcel10
                                    '-------------------------
                                    'aktualizacja zmiennych
                                    'drv("Telefony") += pTelefon
                                    'drv("WizytyO") += pWizytyO
                                    'drv("WizytyProm") += pWizytyProm
                                    'drv("Pozostale") += pPozostale
                                    drv("KlBezdr") += pKlBezDr
                                    drv("Oferty") += pOferty
                                    drv("Testy") += pTesty
                                    drv("Czyszczenie") += pCzyszczenie
                                    drv("Dostawy") += pDostawy
                                    drv("Skup") += pSkup
                                    'drv("SumaKon") += pRazemKon
                                    'drv("KlLPoten") += pKlLPoten
                                    'drv("KlLPonow") += pKlLPonow
                                    'drv("KlAPoten") += pKlAPoten
                                    'drv("KlAPonow") += pKlAPonow
                                    'drv("NowiKl") += pNowiKl
                                    drv("Kolumna01") += pExcel01
                                    drv("Kolumna02") += pExcel02
                                    drv("Kolumna03") += pExcel03
                                    drv("Kolumna04") += pExcel04
                                    drv("Kolumna05") += pExcel05
                                    drv("Kolumna06") += pExcel06
                                    drv("Kolumna07") += pExcel07
                                    drv("Kolumna08") += pExcel08
                                    drv("Kolumna09") += pExcel09
                                    drv("Kolumna10") += pExcel10



                                    'SumTelefon += pTelefon
                                    'SumWizytyO += pWizytyO
                                    'SumWizytyProm += pWizytyProm
                                    'SumPozostale += pPozostale
                                    SumKlBezDr += pKlBezDr
                                    SumOferty += pOferty
                                    SumTesty += pTesty
                                    SumCzyszczenie += pCzyszczenie
                                    SumDostawy += pDostawy
                                    SumSkup += pSkup
                                    'SumRazemKon += pRazemKon
                                    'SumKlLPoten += pKlLPoten
                                    'SumKlLPonow += pKlLPonow
                                    'SumKlAPoten += pKlAPoten
                                    'SumKlAPonow += pKlAPonow
                                    'SumNowiKl += pNowiKl
                                    SumExcel01 += pExcel01
                                    SumExcel02 += pExcel02
                                    SumExcel03 += pExcel03
                                    SumExcel04 += pExcel04
                                    SumExcel05 += pExcel05
                                    SumExcel06 += pExcel06
                                    SumExcel07 += pExcel07
                                    SumExcel08 += pExcel08
                                    SumExcel09 += pExcel09
                                    SumExcel10 += pExcel10

                                    'TotTelefon += pTelefon
                                    'TotWizytyO += pWizytyO
                                    'TotWizytyProm += pWizytyProm
                                    'TotPozostale += pPozostale
                                    TotKlBezDr += pKlBezDr
                                    TotOferty += pOferty
                                    TotTesty += pTesty
                                    TotCzyszczenie += pCzyszczenie
                                    TotDostawy += pDostawy
                                    TotSkup += pSkup
                                    'TotRazemKon += pRazemKon
                                    'TotKlLPoten += pKlLPoten
                                    'TotKlLPonow += pKlLPonow
                                    'TotKlAPoten += pKlAPoten
                                    'TotKlAPonow += pKlAPonow
                                    'TotNowiKl += pNowiKl
                                    TotExcel01 += pExcel01
                                    TotExcel02 += pExcel02
                                    TotExcel03 += pExcel03
                                    TotExcel04 += pExcel04
                                    TotExcel05 += pExcel05
                                    TotExcel06 += pExcel06
                                    TotExcel07 += pExcel07
                                    TotExcel08 += pExcel08
                                    TotExcel09 += pExcel09
                                    TotExcel10 += pExcel10
                                Next
                            End If

                            If IleRecPrac > 0 Then
                                'dodaj wiersz SUMA DLA PRACOWNIKA
                                dr = DataSetRoboczy.Tables(0).NewRow()
                                'Pracownik|NazwaP|OpisP|Data|Telefony|WizytyO|WizytyPRom|Pozostale|KlBezDr|Oferty|Testy|Czyszczenie|Dostawy|Skup|SumaKon|KlLPoten|KlLPonow|KlAPoten|KlAPonow|NowiKl
                                dr("Pracownik") = pPracownik
                                dr("NazwaP") = pNazwa
                                dr("Opisp") = pOpis
                                dr("Data") = "RAZEM:"
                                dr("DzienTyg") = ""
                                dr("Telefony") = SumTelefon
                                dr("WizytyO") = SumWizytyO
                                dr("WizytyProm") = SumWizytyProm
                                dr("Pozostale") = SumPozostale
                                dr("KlBezdr") = SumKlBezDr
                                dr("Oferty") = SumOferty
                                dr("Testy") = SumTesty
                                dr("Czyszczenie") = SumCzyszczenie
                                dr("Dostawy") = SumDostawy
                                dr("Skup") = SumSkup
                                dr("SumaKon") = SumRazemKon
                                dr("KlLPoten") = SumKlLPoten
                                dr("KlLPonow") = SumKlLPonow
                                dr("KlAPoten") = SumKlAPoten
                                dr("KlAPonow") = SumKlAPonow
                                dr("NowiKl") = SumNowiKl
                                dr("Kolumna01") = SumExcel01
                                dr("Kolumna02") = SumExcel02
                                dr("Kolumna03") = SumExcel03
                                dr("Kolumna04") = SumExcel04
                                dr("Kolumna05") = SumExcel05
                                dr("Kolumna06") = SumExcel06
                                dr("Kolumna07") = SumExcel07
                                dr("Kolumna08") = SumExcel08
                                dr("Kolumna09") = SumExcel09
                                dr("Kolumna10") = SumExcel10
                                DataSetRoboczy.Tables(0).Rows.Add(dr)
                            End If
                        Next
                        '--------------------------
                        If pDodajTotal Then
                            'wiersz TOTAL MUSI byc na koncu....odrêbny DATASET na TOTAL
                            DataSetTotal = DataSetRoboczy.Clone
                            dr = DataSetTotal.Tables(0).NewRow()
                            'Pracownik|NazwaP|OpisP|Data|Telefony|WizytyO|WizytyProm|Pozostale|KlBezDr|Oferty|Testy|Czyszczenie|Dostawy|Skup|SumaKon|KlLPoten|KlLPonow|KlAPoten|KlAPonow|NowiKl
                            dr("Pracownik") = ""
                            dr("NazwaP") = ""
                            dr("Opisp") = ""
                            dr("Data") = "RAZEM RAPORT:"
                            dr("DzienTyg") = ""
                            dr("Telefony") = TotTelefon
                            dr("WizytyO") = TotWizytyO
                            dr("WizytyProm") = TotWizytyProm
                            dr("Pozostale") = TotPozostale
                            dr("KlBezdr") = TotKlBezDr
                            dr("Oferty") = TotOferty
                            dr("Testy") = TotTesty
                            dr("Czyszczenie") = TotCzyszczenie
                            dr("Dostawy") = TotDostawy
                            dr("Skup") = TotSkup
                            dr("SumaKon") = TotRazemKon
                            dr("KlLPoten") = TotKlLPoten
                            dr("KlLPonow") = TotKlLPonow
                            dr("KlAPoten") = TotKlAPoten
                            dr("KlAPonow") = TotKlAPonow
                            dr("NowiKl") = TotNowiKl
                            dr("Kolumna01") = TotExcel01
                            dr("Kolumna02") = TotExcel02
                            dr("Kolumna03") = TotExcel03
                            dr("Kolumna04") = TotExcel04
                            dr("Kolumna05") = TotExcel05
                            dr("Kolumna06") = TotExcel06
                            dr("Kolumna07") = TotExcel07
                            dr("Kolumna08") = TotExcel08
                            dr("Kolumna09") = TotExcel09
                            dr("Kolumna10") = TotExcel10
                            DataSetTotal.Tables(0).Rows.Add(dr)
                            DataViewTotal = New DataView(DataSetTotal.Tables(0))
                        End If
                        '--------------------------
                        RetValue = True
                    End If
                End If
            End If
        End If
        Return RetValue
    End Function





    'Public oUniqIdKon As String           'UNIQID            varchar(40)
    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
    'Public oTematKon As String            'TEMAT            Text, 50
    'Public oUwagiKon As String            'UWAGI            Memo
    'Public oWersjaKon As Integer          'WERSJA           Integer

    Private Sub PobierzDaneZKontaktow(ByVal LiRow As Integer)
        oIdentKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "IDENTKLIENTA")
        oDataKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "DATAKONTAKTU")
        oPracownikKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "PRACOWNIK")
        oTematKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "RODZAJ")  'alias
        oUwagiKon = GetTxtDRVField(DataViewKontakty.Item(LiRow), "OPIS")    'alias
        oWersjaKon = GetIntDRVField(DataViewKontakty.Item(LiRow), "WERSJA")
    End Sub












    '**************************************************************
    '** wydruk z bazy danych
    '**************************************************************

    Dim AnalPrac As String = ""

    Private Sub cmdDrukuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDrukuj.Click
        PrzygotujDane(False)
        LicznikRec = 0
        AnalPrac = ""

        ' Tworzymy dokument i do³¹czmy obs³ugê zdarzenia PrintPage
        Dim MyDoc As New Softart.myPrintDocument

        MyDoc.DocumentName = "Raport Aktywnoœci Pracowników"

        MyDoc.LineNumber = 0
        MyDoc.PageNumber = 0
        MyDoc.NextRowNumber = 0
        MyDoc.DefaultPageSettings.Landscape = True
        MyDoc.DefaultPageSettings.Margins.Bottom = cm2pts(0.5)
        MyDoc.DefaultPageSettings.Margins.Top = cm2pts(0.5)
        MyDoc.DefaultPageSettings.Margins.Left = cm2pts(0.5)
        MyDoc.DefaultPageSettings.Margins.Right = cm2pts(0.5)
        AddHandler MyDoc.PrintPage, AddressOf MyDoc_ZestawienieTowarow
        '--------------------------------------

        Try
            'wybor drukarki...
            Dim PDialog As New PrintDialog
            PDialog.Document = MyDoc
            PDialog.AllowPrintToFile = True
            PDialog.PrintToFile = False
            PDialog.AllowSelection = True
            PDialog.AllowSomePages = True
            PDialog.ShowHelp = True
            PDialog.ShowNetwork = True
            PDialog.UseEXDialog = True
            Dim Result As DialogResult = PDialog.ShowDialog()
            ' Drukuj po nacisnieciu OK
            If Result = Windows.Forms.DialogResult.OK Then
                ' This method returns immediately, before the print job starts.
                ' The PrintPage event will fire asynchronously.
                'MyDoc.Print()

                'podiroglad wydruku
                Dim dlgPreview As New PrintPreviewDialog
                dlgPreview.Document = MyDoc
                dlgPreview.WindowState = FormWindowState.Maximized
                dlgPreview.UseAntiAlias = True  'lepsza jakosc wydruku
                dlgPreview.ShowDialog()
            End If
        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Error, _
                MessageBoxDefaultButton.Button1)
        Finally
            'MyDoc.Dispose()
        End Try
    End Sub






    '**************************************************************
    '** wydruk z bazy danych
    '**************************************************************

    Private Sub MyDoc_ZestawienieTowarow(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        ' Retrieve the document that sent this event.
        ' You could store the document in a class member variable,
        ' but this approach allows you to use the same event handler
        ' to handle multiple print documents at once.
        Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
        Dim Text As String
        Dim LiRow As Integer
        Dim drv As DataRowView = Nothing

        e.PageSettings.Landscape = True

        ' Define the font and determine the line height.
        Dim FontSize As Integer = 8
        Dim FontSizeTitle As Integer = 12
        Dim MyFontTitle As New Font("Arial", FontSizeTitle, FontStyle.Bold)
        Dim LineHeightTitle As Single = MyFontTitle.GetHeight(e.Graphics)
        Dim MyFont As New Font("Arial", FontSize, FontStyle.Regular)
        Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)
        Dim DrawingPen As New Pen(Color.Black, 1)

        ' Create variables to hold position on page.
        Dim x As Single = e.MarginBounds.Left
        Dim y As Single = e.MarginBounds.Top

        Dim PageWidth As Single = cm2pts(29)      'e.MarginBounds.Width
        Dim PageHeight As Single = cm2pts(19)   'e.MarginBounds.Height
        Dim PageLeft As Single = cm2pts(0.2)       ' e.MarginBounds.Left
        Dim PageTop As Single = cm2pts(1)          'e.MarginBounds.Top

        Dim Tekst As String = ""
        Dim Drukuj As Boolean = True
        Dim xBase As Single = 0
        Dim yBase As Single = 0

        Do While True
            ' Naglowek strony
            Doc.PageNumber += 1
            y = e.MarginBounds.Top

            x = PageLeft
            LeftTxt(Par_NazwaUzytkownika, x, y, cm2pts(29), MyFont, e)
            RightTxt(Trim(Par_MiejscowoscUzytkownika) & " " & SysData(), x, y, cm2pts(29), MyFont, e)
            y += LineHeight
            LeftTxt(Par_AdresUzytkownika, x, y, cm2pts(29), MyFont, e)
            RightTxt("Str. " & Trim(Str(Doc.PageNumber)), x, y, cm2pts(29), MyFont, e)
            y += LineHeight

            Text = "Raport Aktywnoœci Pracowników w okresie od " & dtpOdDaty.Value.ToString("yyyy-MM-dd") & " do " & dtpDoDaty.Value.ToString("yyyy-MM-dd")
            CentrTxt(Text, x, y, cm2pts(29), MyFontTitle, e)
            y += MyFontTitle.GetHeight(e.Graphics)
            y += LineHeight

            'nag³ówek tabeli
            x = PageLeft
            xBase = x
            yBase = y

            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(3.5), CType(5 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(5.5), CType(5 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(6.5), CType(5 * MyFont.GetHeight(e.Graphics), Single))

            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(6.5), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(1.4), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(6.5), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(2.8), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(6.5), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(4.2), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(6.5), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(5.6), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(6.5), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(7.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))

            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(6.5), CType(2 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(7.9), CType(2 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(9.3), CType(2 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(10.7), CType(2 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(12.1), CType(2 * MyFont.GetHeight(e.Graphics), Single))

            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(13.5), CType(5 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(14.9), CType(5 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(16.3), CType(5 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(17.7), CType(5 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(19.1), CType(5 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(20.5), CType(5 * MyFont.GetHeight(e.Graphics), Single))

            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(21.9), CType(5 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(23.3), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(1.4), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(23.3), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(2.8), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(23.3), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(4.2), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(23.3), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(5.6), CType(3 * MyFont.GetHeight(e.Graphics), Single))

            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(21.9), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(1.4), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(21.9), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(2.8), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(21.9), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(4.2), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(21.9), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(5.6), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x + cm2pts(21.9), y + CType(2 * MyFont.GetHeight(e.Graphics), Single), cm2pts(7.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))


            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(23.3), CType(2 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(24.7), CType(2 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(26.1), CType(2 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(27.5), CType(2 * MyFont.GetHeight(e.Graphics), Single))

            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(28.9), CType(5 * MyFont.GetHeight(e.Graphics), Single))




            y += 0.5 * MyFont.GetHeight(e.Graphics)
            CentrTxt("Pracownik", x, y, cm2pts(3.5), MyFont, e)
            CentrTxt("Data", x + cm2pts(3.5), y, cm2pts(2.0), MyFont, e)
            CentrTxt("DzTyg.", x + cm2pts(5.5), y, cm2pts(1.0), MyFont, e)

            CentrTxt("Kontakt", x + cm2pts(6.5), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Kontakt", x + cm2pts(7.9), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Wizyty", x + cm2pts(9.3), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Kontakt", x + cm2pts(10.7), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Kontakt", x + cm2pts(12.1), y, cm2pts(1.4), MyFont, e)

            CentrTxt("", x + cm2pts(13.5), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(14.9), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(16.3), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(17.7), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(19.1), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(20.5), y, cm2pts(1.4), MyFont, e)

            CentrTxt("Aktywny kontakt", x + cm2pts(21.9), y, cm2pts(4 * 1.4), MyFont, e)
            CentrTxt("", x + cm2pts(27.5), y, cm2pts(1.4), MyFont, e)
            y += MyFont.GetHeight(e.Graphics)
            y += MyFont.GetHeight(e.Graphics)


            CentrTxt("", x, y, cm2pts(3.5), MyFont, e)
            CentrTxt("", x + cm2pts(3.5), y, cm2pts(2.0), MyFont, e)
            CentrTxt("", x + cm2pts(5.5), y, cm2pts(1.0), MyFont, e)

            CentrTxt("Telefony", x + cm2pts(6.5), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Opieka", x + cm2pts(7.9), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Prom.", x + cm2pts(9.3), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Pozosta³e", x + cm2pts(10.7), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Kl. bez", x + cm2pts(12.1), y, cm2pts(1.4), MyFont, e)

            CentrTxt("Oferty", x + cm2pts(13.5), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Testy", x + cm2pts(14.9), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Czyszczenie", x + cm2pts(16.3), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Dostawy", x + cm2pts(17.7), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Skup", x + cm2pts(19.1), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Suma", x + cm2pts(20.5), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Kl. laser", x + cm2pts(21.9), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Kl. laser", x + cm2pts(23.3), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Kl. atram.", x + cm2pts(24.7), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Kl. atram.", x + cm2pts(26.1), y, cm2pts(1.4), MyFont, e)
            CentrTxt("Pierwszy", x + cm2pts(27.5), y, cm2pts(1.4), MyFont, e)
            y += MyFont.GetHeight(e.Graphics)



            CentrTxt("", x, y, cm2pts(3.5), MyFont, e)
            CentrTxt("", x + cm2pts(3.5), y, cm2pts(2.0), MyFont, e)
            CentrTxt("", x + cm2pts(5.5), y, cm2pts(1.0), MyFont, e)

            CentrTxt("", x + cm2pts(6.5), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(7.9), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(9.3), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(10.7), y, cm2pts(1.4), MyFont, e)
            CentrTxt("drukarki", x + cm2pts(12.1), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(13.5), y, cm2pts(1.4), MyFont, e)
            CentrTxt("wstawione", x + cm2pts(14.9), y, cm2pts(1.4), MyFont, e)
            CentrTxt("klienci", x + cm2pts(16.3), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(17.7), y, cm2pts(1.4), MyFont, e)
            CentrTxt("", x + cm2pts(19.1), y, cm2pts(1.4), MyFont, e)
            CentrTxt("kontaktów", x + cm2pts(20.5), y, cm2pts(1.4), MyFont, e)
            CentrTxt("potenc.", x + cm2pts(21.9), y, cm2pts(1.4), MyFont, e)
            CentrTxt("ponow.", x + cm2pts(23.3), y, cm2pts(1.4), MyFont, e)
            CentrTxt("potenc.", x + cm2pts(24.7), y, cm2pts(1.4), MyFont, e)
            CentrTxt("ponow.", x + cm2pts(26.1), y, cm2pts(1.4), MyFont, e)
            CentrTxt("kontakt", x + cm2pts(27.5), y, cm2pts(1.4), MyFont, e)
            y += 2 * MyFont.GetHeight(e.Graphics)

            ' petla do konca strony
            If Doc.NextRowNumber < DataViewRoboczy.Count Then
                Do
                    'ustaw kursor na wierszu DataGrid
                    LiRow = Doc.NextRowNumber
                    drv = DataViewRoboczy.Item(LiRow)

                    pPracownik = drv("Pracownik")
                    pNazwa = drv("NazwaP")
                    pOpis = drv("Opisp")
                    pData = drv("Data")
                    pTelefon = drv("Telefony")
                    pWizytyO = drv("WizytyO")
                    pWizytyProm = drv("WizytyProm")
                    pPozostale = drv("Pozostale")
                    pKlBezDr = drv("KlBezdr")
                    pOferty = drv("Oferty")
                    pTesty = drv("Testy")
                    pCzyszczenie = drv("Czyszczenie")
                    pDostawy = drv("Dostawy")
                    pSkup = drv("Skup")
                    pRazemKon = drv("SumaKon")
                    pKlLPoten = drv("KlLPoten")
                    pKlLPonow = drv("KlLPonow")
                    pKlAPoten = drv("KlAPoten")
                    pKlAPonow = drv("KlAPonow")
                    pNowiKl = drv("NowiKl")

                    If (AnalPrac <> pPracownik) And (Doc.LineNumber > 0) Then
                        y += Int(0.5 * LineHeight)
                        e.Graphics.DrawLine(DrawingPen, x, y, x + cm2pts(29), y)
                        y += Int(0.5 * LineHeight)
                    End If
                    AnalPrac = pPracownik

                    '--------------------------------------
                    x = PageLeft 'od marginesu...
                    Doc.LineNumber += 1
                    '------------------------------------
                    'Tekst = Trim(Str(Doc.LineNumber))
                    'RightTxt(Tekst, x, y, cm2pts(0.8), MyFont, e)

                    LeftTxt(pPracownik & " " & pNazwa, x + cm2pts(0.0), y, cm2pts(3.5), MyFont, e)
                    LeftTxt(pData, x + cm2pts(3.5), y, cm2pts(2.0), MyFont, e)
                    If pData = "RAZEM:" Then
                        'LeftTxt(pData, x + cm2pts(5.5), y, cm2pts(1.0), MyFont, e)
                    Else
                        CentrTxt(Microsoft.VisualBasic.Left(DzienTygodnia(pData), 3), x + cm2pts(5.5), y, cm2pts(1.0), MyFont, e)
                    End If
                    RightTxt(pTelefon.ToString("N0"), x + cm2pts(6.5), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pWizytyO.ToString("N0"), x + cm2pts(7.9), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pWizytyProm.ToString("N0"), x + cm2pts(9.3), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pPozostale.ToString("N0"), x + cm2pts(10.7), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pKlBezDr.ToString("N0"), x + cm2pts(12.1), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pOferty.ToString("N0"), x + cm2pts(13.5), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pTesty.ToString("N0"), x + cm2pts(14.9), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pCzyszczenie.ToString("N0"), x + cm2pts(16.3), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pDostawy.ToString("N0"), x + cm2pts(17.7), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pSkup.ToString("N0"), x + cm2pts(19.1), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pRazemKon.ToString("N0"), x + cm2pts(20.5), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pKlLPoten.ToString("N0"), x + cm2pts(21.9), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pKlLPonow.ToString("N0"), x + cm2pts(23.3), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pKlAPoten.ToString("N0"), x + cm2pts(24.7), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pKlAPonow.ToString("N0"), x + cm2pts(26.1), y, cm2pts(1.4), MyFont, e)
                    RightTxt(pNowiKl.ToString("N0"), x + cm2pts(27.5), y, cm2pts(1.4), MyFont, e)
                    y += LineHeight

                    'y += Int(0.5 * LineHeight)
                    'e.Graphics.DrawLine(DrawingPen, x, y, x + cm2pts(29), y)
                    'y += Int(0.5 * LineHeight)

                    'Text = Str(Doc.LineNumber)
                    'e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + CType(0.0 * (100 / 2.54), Single), y)
                    Doc.NextRowNumber += 1
                Loop Until ((y + 3 * LineHeight) > e.MarginBounds.Bottom) Or (Doc.NextRowNumber >= DataViewRoboczy.Count)
            End If


            If (Doc.NextRowNumber < DataViewRoboczy.Count) Then
                ' There is still at least one more page.
                ' Signal this event to fire again.
                e.HasMorePages = True

                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(6.5), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(13.5), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(21.9), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(28.9), y - yBase)

            Else
                ' Printing is complete.
                y += Int(0.5 * LineHeight)
                e.Graphics.DrawLine(DrawingPen, x, y, x + cm2pts(29), y)
                y += Int(0.5 * LineHeight)

                LeftTxt("Razem Raport :", x + cm2pts(3.5), y, cm2pts(3.0), MyFont, e)

                RightTxt(TotTelefon.ToString("N0"), x + cm2pts(6.5), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotWizytyO.ToString("N0"), x + cm2pts(7.9), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotWizytyProm.ToString("N0"), x + cm2pts(9.3), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotPozostale.ToString("N0"), x + cm2pts(10.7), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotKlBezDr.ToString("N0"), x + cm2pts(12.1), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotOferty.ToString("N0"), x + cm2pts(13.5), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotTesty.ToString("N0"), x + cm2pts(14.9), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotCzyszczenie.ToString("N0"), x + cm2pts(16.3), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotDostawy.ToString("N0"), x + cm2pts(17.7), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotSkup.ToString("N0"), x + cm2pts(19.1), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotRazemKon.ToString("N0"), x + cm2pts(20.5), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotKlLPoten.ToString("N0"), x + cm2pts(21.9), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotKlLPonow.ToString("N0"), x + cm2pts(23.3), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotKlAPoten.ToString("N0"), x + cm2pts(24.7), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotKlAPonow.ToString("N0"), x + cm2pts(26.1), y, cm2pts(1.4), MyFont, e)
                RightTxt(TotNowiKl.ToString("N0"), x + cm2pts(27.5), y, cm2pts(1.4), MyFont, e)

                y += LineHeight

                y += Int(0.5 * LineHeight)
                e.Graphics.DrawLine(DrawingPen, x, y, x + cm2pts(29), y)

                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(6.5), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(13.5), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(21.9), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(28.9), y - yBase)

                y += Int(0.5 * LineHeight)
                x = PageLeft 'od marginesu...
                Text = "Koniec Raportu"
                'Text = "Koniec Kontaktya na pozycji " & Str(Doc.LineNumber)
                e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + cm2pts(0.2), y)

                e.HasMorePages = False
            End If

            '------------------------------
            'sprawdz czy strona do wydrukowania
            If Doc.PrinterSettings.ToPage = 0 Then
                'nie ma ograniczenia stron
                Exit Do
            Else
                If Doc.PageNumber = Doc.PrinterSettings.ToPage Then
                    'ostatnia drukowana strona
                    e.HasMorePages = False
                    Exit Do
                Else
                    If Doc.PageNumber < Doc.PrinterSettings.FromPage Or Doc.PageNumber > Doc.PrinterSettings.ToPage Then
                        'nie drukuj
                        e.Graphics.Clear(Color.White)
                        If e.HasMorePages = False Then
                            Exit Do
                        End If
                    Else
                        'drukuj
                        Exit Do
                    End If
                End If
            End If
        Loop

        If e.HasMorePages = False Then
            Doc.PageNumber = 0
            Doc.LineNumber = 0
            Doc.NextRowNumber = 0

            AnalPrac = ""
            '----------------------------
        End If
    End Sub






    ''----------------------
    '' generujemy arkusz w Excel
    ''----------------------

    Private Sub cmdZapiszExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZapiszExcel.Click
        PrzygotujDane(True) 'dataviewroboczy
        '--------------------------------------
        'KeyKolumn1 = TableRoboczy.Columns.Add("Pracownik", Type.GetType("System.String"))
        'TableRoboczy.Columns.Add("NazwaP", Type.GetType("System.String"))
        'TableRoboczy.Columns.Add("opisP", Type.GetType("System.String"))
        'KeyKolumn2 = TableRoboczy.Columns.Add("Data", Type.GetType("System.String"))
        'TableRoboczy.Columns.Add("DzienTyg", Type.GetType("System.String"))
        'TableRoboczy.Columns.Add("Telefony", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("WizytyO", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("WizytyProm", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Pozostale", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("KlBezDr", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Oferty", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Testy", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Czyszczenie", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Dostawy", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Skup", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("SumaKon", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("KlLPoten", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("KlLPonow", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("KlAPoten", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("KlAPonow", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("NowiKl", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna01", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna02", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna03", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna04", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna05", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna06", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna07", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna08", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna09", Type.GetType("System.Int32"))
        'TableRoboczy.Columns.Add("Kolumna10", Type.GetType("System.Int32"))
        '---------------------------------------
        Try
            'UWAGA : FormatRaport zawiera format pól liczbowych dla EXCEL
            'EXCEL po Polsku u¿ywa ',' jaoo separator liczb dziesietnych !!!
            Dim FormatRaport(21 + 10) As String
            FormatRaport(0) = "@"
            FormatRaport(1) = "@"
            FormatRaport(2) = "@"
            FormatRaport(3) = "@"
            FormatRaport(4) = "@"
            FormatRaport(5) = "########0"
            FormatRaport(6) = "########0"
            FormatRaport(7) = "########0"
            FormatRaport(8) = "########0"
            FormatRaport(9) = "########0"
            FormatRaport(10) = "########0"
            FormatRaport(11) = "########0"
            FormatRaport(12) = "########0"
            FormatRaport(13) = "########0"
            FormatRaport(14) = "########0"
            FormatRaport(15) = "########0"
            FormatRaport(16) = "########0"
            FormatRaport(17) = "########0"
            FormatRaport(18) = "########0"
            FormatRaport(19) = "########0"
            FormatRaport(20) = "########0"
            FormatRaport(20 + 1) = "@"
            FormatRaport(20 + 2) = "@"
            FormatRaport(20 + 3) = "@"
            FormatRaport(20 + 4) = "@"
            FormatRaport(20 + 5) = "@"
            FormatRaport(20 + 6) = "@"
            FormatRaport(20 + 7) = "@"
            FormatRaport(20 + 8) = "@"
            FormatRaport(20 + 9) = "@"
            FormatRaport(20 + 10) = "@"

            Dim FormatNaglowek1(21 + 10) As String
            FormatNaglowek1(0) = "Pracownik"
            FormatNaglowek1(1) = "Nazwisko i Imie"
            FormatNaglowek1(2) = "Funkcja"
            FormatNaglowek1(3) = "Data"
            FormatNaglowek1(4) = "Dzien Tyg."
            FormatNaglowek1(5) = "Kontakt"
            FormatNaglowek1(6) = "Kontakt"
            FormatNaglowek1(7) = "Wizyty"
            FormatNaglowek1(8) = "Kontakt"
            FormatNaglowek1(9) = "Kontakt"
            FormatNaglowek1(10) = "Oferty"
            FormatNaglowek1(11) = "Testy"
            FormatNaglowek1(12) = "Czyszczenie"
            FormatNaglowek1(13) = "Dostawy"
            FormatNaglowek1(14) = "Skup"
            FormatNaglowek1(15) = "Suma"
            FormatNaglowek1(16) = "Klienci L"
            FormatNaglowek1(17) = "Klienci L"
            FormatNaglowek1(18) = "Klienci A"
            FormatNaglowek1(19) = "Klienci A"
            FormatNaglowek1(20) = "Nowi"
            FormatNaglowek1(20 + 1) = Trim(Par_RAKolumna01)
            FormatNaglowek1(20 + 2) = Trim(Par_RAKolumna02)
            FormatNaglowek1(20 + 3) = Trim(Par_RAKolumna03)
            FormatNaglowek1(20 + 4) = Trim(Par_RAKolumna04)
            FormatNaglowek1(20 + 5) = Trim(Par_RAKolumna05)
            FormatNaglowek1(20 + 6) = Trim(Par_RAKolumna06)
            FormatNaglowek1(20 + 7) = Trim(Par_RAKolumna07)
            FormatNaglowek1(20 + 8) = Trim(Par_RAKolumna08)
            FormatNaglowek1(20 + 9) = Trim(Par_RAKolumna09)
            FormatNaglowek1(20 + 10) = Trim(Par_RAKolumna10)

            Dim FormatNaglowek2(21 + 10) As String
            FormatNaglowek2(0) = ""
            FormatNaglowek2(1) = ""
            FormatNaglowek2(2) = ""
            FormatNaglowek2(3) = ""
            FormatNaglowek2(4) = ""
            FormatNaglowek2(5) = "Telefony"
            FormatNaglowek2(6) = "Opieka"
            FormatNaglowek2(7) = "Prom."
            FormatNaglowek2(8) = "Pozosta³e"
            FormatNaglowek2(9) = "Kl. bez dr."
            FormatNaglowek2(10) = ""
            FormatNaglowek2(11) = "wstawione"
            FormatNaglowek2(12) = "klienci"
            FormatNaglowek2(13) = ""
            FormatNaglowek2(14) = ""
            FormatNaglowek2(15) = "kontaktów"
            FormatNaglowek2(16) = "potenc."
            FormatNaglowek2(17) = "ponow."
            FormatNaglowek2(18) = "potenc."
            FormatNaglowek2(19) = "ponow."
            FormatNaglowek2(20) = "klienci"
            FormatNaglowek2(20 + 1) = ""
            FormatNaglowek2(20 + 2) = ""
            FormatNaglowek2(20 + 3) = ""
            FormatNaglowek2(20 + 4) = ""
            FormatNaglowek2(20 + 5) = ""
            FormatNaglowek2(20 + 6) = ""
            FormatNaglowek2(20 + 7) = ""
            FormatNaglowek2(20 + 8) = ""
            FormatNaglowek2(20 + 9) = ""
            FormatNaglowek2(20 + 10) = ""

            'przekazujemy dataview do excela
            Me.Cursor = Cursors.WaitCursor
            ExportDataView2Excel(DataViewRoboczy, _
                                 DataViewTotal, _
                                 FormatRaport, _
                                 FormatNaglowek1, _
                                 FormatNaglowek2, _
                                "Raport Aktywnoœci Pracowników w okresie od " & dtpOdDaty.Value.ToString("yyyy-MM-dd") & " do " & dtpDoDaty.Value.ToString("yyyy-MM-dd"))
            Me.Cursor = Cursors.Default

            '--------------------------------------


        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Error, _
                MessageBoxDefaultButton.Button1)
        Finally
            'MyDoc.Dispose()
            '--------------------------------------
        End Try
    End Sub






End Class
