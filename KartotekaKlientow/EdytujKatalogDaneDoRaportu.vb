Public Class EdytujKatalogDaneDoRaportu
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPrzywroc As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblStanowisko As System.Windows.Forms.Label
    Friend WithEvents lblNazwisko As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtDostawy As System.Windows.Forms.TextBox
    Friend WithEvents txtCzyszczenie As System.Windows.Forms.TextBox
    Friend WithEvents txtTesty As System.Windows.Forms.TextBox
    Friend WithEvents txtOferty As System.Windows.Forms.TextBox
    Friend WithEvents txtSkup As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtPracownik As System.Windows.Forms.TextBox
    Friend WithEvents cmdPracownik As System.Windows.Forms.Button
    Friend WithEvents txtEXCEL10 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL09 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL08 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL07 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL06 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL05 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL04 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL03 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL02 As System.Windows.Forms.TextBox
    Friend WithEvents txtEXCEL01 As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL10 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL09 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL08 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL07 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL06 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL05 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL04 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL03 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL02 As System.Windows.Forms.Label
    Friend WithEvents lblEXCEL01 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents dtpDataRaportu As DateTimePicker
    Friend WithEvents txtKlBezDrukarek As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogDaneDoRaportu))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dtpDataRaportu = New System.Windows.Forms.DateTimePicker()
        Me.txtEXCEL10 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL09 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL08 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL07 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL06 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL05 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL04 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL03 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL02 = New System.Windows.Forms.TextBox()
        Me.txtEXCEL01 = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.lblEXCEL10 = New System.Windows.Forms.Label()
        Me.lblEXCEL09 = New System.Windows.Forms.Label()
        Me.lblEXCEL08 = New System.Windows.Forms.Label()
        Me.lblEXCEL07 = New System.Windows.Forms.Label()
        Me.lblEXCEL06 = New System.Windows.Forms.Label()
        Me.lblEXCEL05 = New System.Windows.Forms.Label()
        Me.lblEXCEL04 = New System.Windows.Forms.Label()
        Me.lblEXCEL03 = New System.Windows.Forms.Label()
        Me.lblEXCEL02 = New System.Windows.Forms.Label()
        Me.lblEXCEL01 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.cmdPracownik = New System.Windows.Forms.Button()
        Me.txtSkup = New System.Windows.Forms.TextBox()
        Me.txtDostawy = New System.Windows.Forms.TextBox()
        Me.txtCzyszczenie = New System.Windows.Forms.TextBox()
        Me.txtTesty = New System.Windows.Forms.TextBox()
        Me.txtOferty = New System.Windows.Forms.TextBox()
        Me.txtKlBezDrukarek = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtPracownik = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblStanowisko = New System.Windows.Forms.Label()
        Me.lblNazwisko = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdWycofajSie = New System.Windows.Forms.Button()
        Me.cmdPrzywroc = New System.Windows.Forms.Button()
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
        Me.Panel1.Controls.Add(Me.dtpDataRaportu)
        Me.Panel1.Controls.Add(Me.txtEXCEL10)
        Me.Panel1.Controls.Add(Me.txtEXCEL09)
        Me.Panel1.Controls.Add(Me.txtEXCEL08)
        Me.Panel1.Controls.Add(Me.txtEXCEL07)
        Me.Panel1.Controls.Add(Me.txtEXCEL06)
        Me.Panel1.Controls.Add(Me.txtEXCEL05)
        Me.Panel1.Controls.Add(Me.txtEXCEL04)
        Me.Panel1.Controls.Add(Me.txtEXCEL03)
        Me.Panel1.Controls.Add(Me.txtEXCEL02)
        Me.Panel1.Controls.Add(Me.txtEXCEL01)
        Me.Panel1.Controls.Add(Me.Label22)
        Me.Panel1.Controls.Add(Me.lblEXCEL10)
        Me.Panel1.Controls.Add(Me.lblEXCEL09)
        Me.Panel1.Controls.Add(Me.lblEXCEL08)
        Me.Panel1.Controls.Add(Me.lblEXCEL07)
        Me.Panel1.Controls.Add(Me.lblEXCEL06)
        Me.Panel1.Controls.Add(Me.lblEXCEL05)
        Me.Panel1.Controls.Add(Me.lblEXCEL04)
        Me.Panel1.Controls.Add(Me.lblEXCEL03)
        Me.Panel1.Controls.Add(Me.lblEXCEL02)
        Me.Panel1.Controls.Add(Me.lblEXCEL01)
        Me.Panel1.Controls.Add(Me.Label21)
        Me.Panel1.Controls.Add(Me.cmdPracownik)
        Me.Panel1.Controls.Add(Me.txtSkup)
        Me.Panel1.Controls.Add(Me.txtDostawy)
        Me.Panel1.Controls.Add(Me.txtCzyszczenie)
        Me.Panel1.Controls.Add(Me.txtTesty)
        Me.Panel1.Controls.Add(Me.txtOferty)
        Me.Panel1.Controls.Add(Me.txtKlBezDrukarek)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.txtPracownik)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblStanowisko)
        Me.Panel1.Controls.Add(Me.lblNazwisko)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(589, 355)
        Me.Panel1.TabIndex = 2
        '
        'dtpDataRaportu
        '
        Me.dtpDataRaportu.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDataRaportu.CustomFormat = "yyyy-MM-dd"
        Me.dtpDataRaportu.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDataRaportu.Location = New System.Drawing.Point(477, 9)
        Me.dtpDataRaportu.Name = "dtpDataRaportu"
        Me.dtpDataRaportu.Size = New System.Drawing.Size(100, 20)
        Me.dtpDataRaportu.TabIndex = 447
        Me.dtpDataRaportu.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'txtEXCEL10
        '
        Me.txtEXCEL10.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL10.Location = New System.Drawing.Point(477, 315)
        Me.txtEXCEL10.Name = "txtEXCEL10"
        Me.txtEXCEL10.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL10.TabIndex = 446
        Me.txtEXCEL10.Text = "0"
        '
        'txtEXCEL09
        '
        Me.txtEXCEL09.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL09.Location = New System.Drawing.Point(477, 293)
        Me.txtEXCEL09.Name = "txtEXCEL09"
        Me.txtEXCEL09.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL09.TabIndex = 445
        Me.txtEXCEL09.Text = "0"
        '
        'txtEXCEL08
        '
        Me.txtEXCEL08.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL08.Location = New System.Drawing.Point(477, 271)
        Me.txtEXCEL08.Name = "txtEXCEL08"
        Me.txtEXCEL08.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL08.TabIndex = 444
        Me.txtEXCEL08.Text = "0"
        '
        'txtEXCEL07
        '
        Me.txtEXCEL07.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL07.Location = New System.Drawing.Point(477, 249)
        Me.txtEXCEL07.Name = "txtEXCEL07"
        Me.txtEXCEL07.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL07.TabIndex = 443
        Me.txtEXCEL07.Text = "0"
        '
        'txtEXCEL06
        '
        Me.txtEXCEL06.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL06.Location = New System.Drawing.Point(477, 227)
        Me.txtEXCEL06.Name = "txtEXCEL06"
        Me.txtEXCEL06.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL06.TabIndex = 442
        Me.txtEXCEL06.Text = "0"
        '
        'txtEXCEL05
        '
        Me.txtEXCEL05.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL05.Location = New System.Drawing.Point(477, 205)
        Me.txtEXCEL05.Name = "txtEXCEL05"
        Me.txtEXCEL05.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL05.TabIndex = 441
        Me.txtEXCEL05.Text = "0"
        '
        'txtEXCEL04
        '
        Me.txtEXCEL04.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL04.Location = New System.Drawing.Point(477, 183)
        Me.txtEXCEL04.Name = "txtEXCEL04"
        Me.txtEXCEL04.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL04.TabIndex = 440
        Me.txtEXCEL04.Text = "0"
        '
        'txtEXCEL03
        '
        Me.txtEXCEL03.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL03.Location = New System.Drawing.Point(477, 161)
        Me.txtEXCEL03.Name = "txtEXCEL03"
        Me.txtEXCEL03.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL03.TabIndex = 439
        Me.txtEXCEL03.Text = "0"
        '
        'txtEXCEL02
        '
        Me.txtEXCEL02.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL02.Location = New System.Drawing.Point(477, 139)
        Me.txtEXCEL02.Name = "txtEXCEL02"
        Me.txtEXCEL02.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL02.TabIndex = 438
        Me.txtEXCEL02.Text = "0"
        '
        'txtEXCEL01
        '
        Me.txtEXCEL01.ForeColor = System.Drawing.Color.Purple
        Me.txtEXCEL01.Location = New System.Drawing.Point(477, 117)
        Me.txtEXCEL01.Name = "txtEXCEL01"
        Me.txtEXCEL01.Size = New System.Drawing.Size(100, 20)
        Me.txtEXCEL01.TabIndex = 437
        Me.txtEXCEL01.Text = "0"
        '
        'Label22
        '
        Me.Label22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Navy
        Me.Label22.Location = New System.Drawing.Point(8, 80)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(306, 16)
        Me.Label22.TabIndex = 436
        Me.Label22.Text = "Dane uzupe³naij¹ce do Raportu :"
        '
        'lblEXCEL10
        '
        Me.lblEXCEL10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL10.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL10.Location = New System.Drawing.Point(324, 318)
        Me.lblEXCEL10.Name = "lblEXCEL10"
        Me.lblEXCEL10.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL10.TabIndex = 434
        Me.lblEXCEL10.Text = "Dodatkowa kolumna Nr 10. . . . . . . "
        '
        'lblEXCEL09
        '
        Me.lblEXCEL09.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL09.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL09.Location = New System.Drawing.Point(324, 296)
        Me.lblEXCEL09.Name = "lblEXCEL09"
        Me.lblEXCEL09.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL09.TabIndex = 432
        Me.lblEXCEL09.Text = "Dodatkowa kolumna Nr 09. . . . . . . "
        '
        'lblEXCEL08
        '
        Me.lblEXCEL08.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL08.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL08.Location = New System.Drawing.Point(324, 274)
        Me.lblEXCEL08.Name = "lblEXCEL08"
        Me.lblEXCEL08.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL08.TabIndex = 430
        Me.lblEXCEL08.Text = "Dodatkowa kolumna Nr 08. . . . . . . "
        '
        'lblEXCEL07
        '
        Me.lblEXCEL07.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL07.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL07.Location = New System.Drawing.Point(324, 252)
        Me.lblEXCEL07.Name = "lblEXCEL07"
        Me.lblEXCEL07.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL07.TabIndex = 428
        Me.lblEXCEL07.Text = "Dodatkowa kolumna Nr 07. . . . . . . "
        '
        'lblEXCEL06
        '
        Me.lblEXCEL06.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL06.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL06.Location = New System.Drawing.Point(324, 230)
        Me.lblEXCEL06.Name = "lblEXCEL06"
        Me.lblEXCEL06.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL06.TabIndex = 426
        Me.lblEXCEL06.Text = "Dodatkowa kolumna Nr 06. . . . . . . "
        '
        'lblEXCEL05
        '
        Me.lblEXCEL05.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL05.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL05.Location = New System.Drawing.Point(324, 208)
        Me.lblEXCEL05.Name = "lblEXCEL05"
        Me.lblEXCEL05.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL05.TabIndex = 424
        Me.lblEXCEL05.Text = "Dodatkowa kolumna Nr 05. . . . . . . "
        '
        'lblEXCEL04
        '
        Me.lblEXCEL04.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL04.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL04.Location = New System.Drawing.Point(324, 186)
        Me.lblEXCEL04.Name = "lblEXCEL04"
        Me.lblEXCEL04.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL04.TabIndex = 422
        Me.lblEXCEL04.Text = "Dodatkowa kolumna Nr 04. . . . . . . "
        '
        'lblEXCEL03
        '
        Me.lblEXCEL03.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL03.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL03.Location = New System.Drawing.Point(324, 164)
        Me.lblEXCEL03.Name = "lblEXCEL03"
        Me.lblEXCEL03.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL03.TabIndex = 420
        Me.lblEXCEL03.Text = "Dodatkowa kolumna Nr 03. . . . . . . "
        '
        'lblEXCEL02
        '
        Me.lblEXCEL02.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL02.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL02.Location = New System.Drawing.Point(324, 142)
        Me.lblEXCEL02.Name = "lblEXCEL02"
        Me.lblEXCEL02.Size = New System.Drawing.Size(199, 16)
        Me.lblEXCEL02.TabIndex = 418
        Me.lblEXCEL02.Text = "Dodatkowa kolumna Nr 02. . . . . . . "
        '
        'lblEXCEL01
        '
        Me.lblEXCEL01.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblEXCEL01.ForeColor = System.Drawing.Color.Navy
        Me.lblEXCEL01.Location = New System.Drawing.Point(324, 120)
        Me.lblEXCEL01.Name = "lblEXCEL01"
        Me.lblEXCEL01.Size = New System.Drawing.Size(196, 16)
        Me.lblEXCEL01.TabIndex = 416
        Me.lblEXCEL01.Text = "Dodatkowa kolumna Nr 01. . . . . . . "
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Navy
        Me.Label21.Location = New System.Drawing.Point(324, 80)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(250, 40)
        Me.Label21.TabIndex = 415
        Me.Label21.Text = "Dodatkowe kolumny dodane do Raportu  emitowanego w EXCEL'u :"
        '
        'cmdPracownik
        '
        Me.cmdPracownik.Image = CType(resources.GetObject("cmdPracownik.Image"), System.Drawing.Image)
        Me.cmdPracownik.Location = New System.Drawing.Point(219, 7)
        Me.cmdPracownik.Name = "cmdPracownik"
        Me.cmdPracownik.Size = New System.Drawing.Size(32, 23)
        Me.cmdPracownik.TabIndex = 2
        Me.cmdPracownik.Text = "2"
        '
        'txtSkup
        '
        Me.txtSkup.ForeColor = System.Drawing.Color.Purple
        Me.txtSkup.Location = New System.Drawing.Point(157, 227)
        Me.txtSkup.Name = "txtSkup"
        Me.txtSkup.Size = New System.Drawing.Size(100, 20)
        Me.txtSkup.TabIndex = 10
        Me.txtSkup.Text = "0"
        '
        'txtDostawy
        '
        Me.txtDostawy.ForeColor = System.Drawing.Color.Purple
        Me.txtDostawy.Location = New System.Drawing.Point(157, 205)
        Me.txtDostawy.Name = "txtDostawy"
        Me.txtDostawy.Size = New System.Drawing.Size(100, 20)
        Me.txtDostawy.TabIndex = 9
        Me.txtDostawy.Text = "0"
        '
        'txtCzyszczenie
        '
        Me.txtCzyszczenie.ForeColor = System.Drawing.Color.Purple
        Me.txtCzyszczenie.Location = New System.Drawing.Point(157, 183)
        Me.txtCzyszczenie.Name = "txtCzyszczenie"
        Me.txtCzyszczenie.Size = New System.Drawing.Size(100, 20)
        Me.txtCzyszczenie.TabIndex = 8
        Me.txtCzyszczenie.Text = "0"
        '
        'txtTesty
        '
        Me.txtTesty.ForeColor = System.Drawing.Color.Purple
        Me.txtTesty.Location = New System.Drawing.Point(157, 161)
        Me.txtTesty.Name = "txtTesty"
        Me.txtTesty.Size = New System.Drawing.Size(100, 20)
        Me.txtTesty.TabIndex = 7
        Me.txtTesty.Text = "0"
        '
        'txtOferty
        '
        Me.txtOferty.ForeColor = System.Drawing.Color.Purple
        Me.txtOferty.Location = New System.Drawing.Point(157, 139)
        Me.txtOferty.Name = "txtOferty"
        Me.txtOferty.Size = New System.Drawing.Size(100, 20)
        Me.txtOferty.TabIndex = 6
        Me.txtOferty.Text = "0"
        '
        'txtKlBezDrukarek
        '
        Me.txtKlBezDrukarek.ForeColor = System.Drawing.Color.Purple
        Me.txtKlBezDrukarek.Location = New System.Drawing.Point(157, 117)
        Me.txtKlBezDrukarek.Name = "txtKlBezDrukarek"
        Me.txtKlBezDrukarek.Size = New System.Drawing.Size(100, 20)
        Me.txtKlBezDrukarek.TabIndex = 5
        Me.txtKlBezDrukarek.Text = "0"
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(11, 230)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(183, 16)
        Me.Label9.TabIndex = 30
        Me.Label9.Text = "Skup . . . . . . . . . . . . . . . . . . . . . ."
        '
        'Label10
        '
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(324, 12)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(179, 16)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "Data Raportu Aktywnoœci . . . . . . . . ."
        '
        'txtPracownik
        '
        Me.txtPracownik.ForeColor = System.Drawing.Color.Purple
        Me.txtPracownik.Location = New System.Drawing.Point(120, 9)
        Me.txtPracownik.Name = "txtPracownik"
        Me.txtPracownik.Size = New System.Drawing.Size(131, 20)
        Me.txtPracownik.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 70)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(589, 1)
        Me.Panel2.TabIndex = 23
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(8, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 22
        Me.Label8.Text = "Stanowisko . . . . . . . . ."
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(8, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Imie i Nazwisko . . . . . . . . ."
        '
        'lblStanowisko
        '
        Me.lblStanowisko.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStanowisko.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStanowisko.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblStanowisko.ForeColor = System.Drawing.Color.Purple
        Me.lblStanowisko.Location = New System.Drawing.Point(120, 48)
        Me.lblStanowisko.Name = "lblStanowisko"
        Me.lblStanowisko.Size = New System.Drawing.Size(457, 16)
        Me.lblStanowisko.TabIndex = 19
        '
        'lblNazwisko
        '
        Me.lblNazwisko.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwisko.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwisko.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwisko.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwisko.Location = New System.Drawing.Point(120, 32)
        Me.lblNazwisko.Name = "lblNazwisko"
        Me.lblNazwisko.Size = New System.Drawing.Size(457, 16)
        Me.lblNazwisko.TabIndex = 18
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(11, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(183, 16)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Dostawy . . . . . . . . . . . . . . . . . . . . . ."
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(11, 186)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(183, 16)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Czyszczenie Klienci . . . . . . . . . . . . . "
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(11, 164)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(183, 16)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Testy wstawione . . . . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(11, 142)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(183, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Iloœæ z³o¿onych ofert . . . . . . . . . . . . . . . . "
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(11, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(183, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Iloœæ klientów bez drukarek . . . . . . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Pracownik . . . . . . . . ."
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(409, 370)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 25)
        Me.cmdWycofajSie.TabIndex = 12
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPrzywroc
        '
        Me.cmdPrzywroc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrzywroc.Image = CType(resources.GetObject("cmdPrzywroc.Image"), System.Drawing.Image)
        Me.cmdPrzywroc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrzywroc.Location = New System.Drawing.Point(312, 370)
        Me.cmdPrzywroc.Name = "cmdPrzywroc"
        Me.cmdPrzywroc.Size = New System.Drawing.Size(91, 25)
        Me.cmdPrzywroc.TabIndex = 13
        Me.cmdPrzywroc.Text = "Przywróæ"
        Me.cmdPrzywroc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(519, 370)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 25)
        Me.cmdPowrot.TabIndex = 11
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 370)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 21
        Me.PictureBox2.TabStop = False
        '
        'EdytujKatalogDaneDoRaportu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(604, 401)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPrzywroc)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EdytujKatalogDaneDoRaportu"
        Me.ShowInTaskbar = False
        Me.Text = "Dane uzupe³niaj¹ce do Raportu Aktywnoœci"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub EdytujKatalogDaneDoRaportu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtPracownik.Enabled = False
            cmdPracownik.Enabled = False
            dtpDataRaportu.Enabled = False
        Else
            txtPracownik.Enabled = True
            cmdPracownik.Enabled = True
            dtpDataRaportu.Enabled = True
        End If
    End Sub

    Private Sub EdytujKatalogDaneDoRaportu_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtKlBezDrukarek.Focus()
        Else
            txtPracownik.Focus()
        End If
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        ZapiszDane()
        oAktualizuj = True
        Me.Close()
    End Sub


    Private Sub cmdWycofajSie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajSie.Click
        oAktualizuj = False
        Me.Close()
    End Sub

    Private Sub cmdPrzywroc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrzywroc.Click
        PobierzDane()
    End Sub

    Private Sub EdytujKatalogDaneDoRaportu_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub EdytujKatalogDaneDoRaportu_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.ClosedByControlBox Then
            'e.Cancel = True     'nie pozwalaj zamkn¹c forme
            If MessageBox.Show("Zdecydowa³eœ opuœciæ formê bez zapisania wprowadzonych zmian." & vbCrLf & _
                "Czy mam zapisaæ zmiany w Bazie Danych ?", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question, _
                MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

                ZapiszDane()
                oAktualizuj = True
            Else
                oAktualizuj = False
            End If
            'Me._closeClick = False
        Else
            'MsgBox("ZAPISZ lub WYCOFAJ SIÊ")
        End If
    End Sub
    '====================================================


    '***************************************************
    '* procedury pobierania/zapisywania danych
    '***************************************************

    Private Sub PobierzDane()
        txtPracownik.Text = ddrPracownik
        IdentUzytkownika(txtPracownik.Text)
        lblNazwisko.Text = oNazwaUzytkownika
        lblStanowisko.Text = oFunkcjaUzytkownika

        dtpDataRaportu.Value = ddrDataRaportu

        txtKlBezDrukarek.Text = ddrKlBezDrukarki.ToString("F0")
        txtOferty.Text = ddrOferty.ToString("F0")
        txtTesty.Text = ddrTestyWstawionw.ToString("F0")
        txtCzyszczenie.Text = ddrCzyszczenie.ToString("F0")
        txtDostawy.Text = ddrDostawy.ToString("F0")
        txtSkup.Text = ddrSkup.ToString("F0")

        lblEXCEL01.Text = Par_RAKolumna01 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL02.Text = Par_RAKolumna02 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL03.Text = Par_RAKolumna03 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL04.Text = Par_RAKolumna04 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL05.Text = Par_RAKolumna05 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL06.Text = Par_RAKolumna06 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL07.Text = Par_RAKolumna07 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL08.Text = Par_RAKolumna08 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL09.Text = Par_RAKolumna09 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
        lblEXCEL10.Text = Par_RAKolumna10 & " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "

        txtEXCEL01.Text = ddrExcel01.ToString("F0")
        txtEXCEL02.Text = ddrExcel02.ToString("F0")
        txtEXCEL03.Text = ddrExcel03.ToString("F0")
        txtEXCEL04.Text = ddrExcel04.ToString("F0")
        txtEXCEL05.Text = ddrExcel05.ToString("F0")
        txtEXCEL06.Text = ddrExcel06.ToString("F0")
        txtEXCEL07.Text = ddrExcel07.ToString("F0")
        txtEXCEL08.Text = ddrExcel08.ToString("F0")
        txtEXCEL09.Text = ddrExcel09.ToString("F0")
        txtEXCEL10.Text = ddrExcel10.ToString("F0")

    End Sub

    Private Sub ZapiszDane()
        ddrPracownik = txtPracownik.Text
        ddrDataRaportu = dtpDataRaportu.Value.ToString("yyyy-MM-dd")

        ddrKlBezDrukarki = GetNumFromText(txtKlBezDrukarek.Text)
        ddrOferty = GetNumFromText(txtOferty.Text)
        ddrTestyWstawionw = GetNumFromText(txtTesty.Text)
        ddrCzyszczenie = GetNumFromText(txtCzyszczenie.Text)
        ddrDostawy = GetNumFromText(txtDostawy.Text)
        ddrSkup = GetNumFromText(txtSkup.Text)

        ddrExcel01 = GetNumFromText(txtEXCEL01.Text)
        ddrExcel02 = GetNumFromText(txtEXCEL02.Text)
        ddrExcel03 = GetNumFromText(txtEXCEL03.Text)
        ddrExcel04 = GetNumFromText(txtEXCEL04.Text)
        ddrExcel05 = GetNumFromText(txtEXCEL05.Text)
        ddrExcel06 = GetNumFromText(txtEXCEL06.Text)
        ddrExcel07 = GetNumFromText(txtEXCEL07.Text)
        ddrExcel08 = GetNumFromText(txtEXCEL08.Text)
        ddrExcel09 = GetNumFromText(txtEXCEL09.Text)
        ddrExcel10 = GetNumFromText(txtEXCEL10.Text)
    End Sub

    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    Private Sub txtPracownik_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPracownik.TextChanged
        TestLen(txtPracownik, "PRACOWNIK", 10)
    End Sub
    Private Sub txtPracownik_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPracownik.GotFocus
        StartEdycjiTxt(txtPracownik)
    End Sub
    Private Sub txtPracownik_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPracownik.LostFocus
        KoniecEdycjiTxt(txtPracownik)
    End Sub

    'Private Sub txtDataRaportu_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDataRaportu.TextChanged
    '    TestLen(txtDataRaportu, "DataRaportu", 10)
    'End Sub
    'Private Sub txtDataRaportu_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataRaportu.GotFocus
    '    StartEdycjiTxt(txtDataRaportu)
    'End Sub
    'Private Sub txtDataRaportu_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataRaportu.LostFocus
    '    KoniecEdycjiTxt(txtDataRaportu)
    'End Sub

    Private Sub txtKlBezDrukarek_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDostawy.TextChanged
        If TestNum(txtDostawy, "KlBezDrukarek") Then
        End If
    End Sub
    Private Sub txtKlBezDrukarek_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDostawy.GotFocus
        StartEdycjiTxt(txtDostawy)
    End Sub
    Private Sub txtKlBezDrukarek_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDostawy.LostFocus
        KoniecEdycjiTxt(txtDostawy)
    End Sub

    Private Sub txtOferty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOferty.TextChanged
        If TestNum(txtOferty, "Oferty") Then
        End If
    End Sub
    Private Sub txtOferty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOferty.GotFocus
        StartEdycjiTxt(txtOferty)
    End Sub
    Private Sub txtOferty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOferty.LostFocus
        KoniecEdycjiTxt(txtOferty)
    End Sub

    Private Sub txtTesty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTesty.TextChanged
        If TestNum(txtTesty, "Testy") Then
        End If
    End Sub
    Private Sub txtTesty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTesty.GotFocus
        StartEdycjiTxt(txtTesty)
    End Sub
    Private Sub txtTesty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTesty.LostFocus
        KoniecEdycjiTxt(txtTesty)
    End Sub

    Private Sub txtCzyszczenie_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCzyszczenie.TextChanged
        If TestNum(txtCzyszczenie, "Czyszczenie") Then
        End If
    End Sub
    Private Sub txtCzyszczenie_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCzyszczenie.GotFocus
        StartEdycjiTxt(txtCzyszczenie)
    End Sub
    Private Sub txtCzyszczenie_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCzyszczenie.LostFocus
        KoniecEdycjiTxt(txtCzyszczenie)
    End Sub

    Private Sub txtDostawy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKlBezDrukarek.TextChanged
        If TestNum(txtKlBezDrukarek, "Dostawy") Then
        End If
    End Sub
    Private Sub txtDostawy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKlBezDrukarek.GotFocus
        StartEdycjiTxt(txtKlBezDrukarek)
    End Sub
    Private Sub txtDostawy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKlBezDrukarek.LostFocus
        KoniecEdycjiTxt(txtKlBezDrukarek)
    End Sub

    Private Sub txtSkup_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSkup.TextChanged
        If TestNum(txtSkup, "Skup") Then
        End If
    End Sub
    Private Sub txtSkup_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSkup.GotFocus
        StartEdycjiTxt(txtSkup)
    End Sub
    Private Sub txtSkup_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSkup.LostFocus
        KoniecEdycjiTxt(txtSkup)
    End Sub





    Private Sub txtExcel01_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL01.TextChanged
        If TestNum(txtExcel01, "Excel01") Then
        End If
    End Sub
    Private Sub txtExcel01_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL01.GotFocus
        StartEdycjiTxt(txtExcel01)
    End Sub
    Private Sub txtExcel01_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL01.LostFocus
        KoniecEdycjiTxt(txtExcel01)
    End Sub

    Private Sub txtExcel02_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL02.TextChanged
        If TestNum(txtExcel02, "Excel02") Then
        End If
    End Sub
    Private Sub txtExcel02_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL02.GotFocus
        StartEdycjiTxt(txtExcel02)
    End Sub
    Private Sub txtExcel02_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL02.LostFocus
        KoniecEdycjiTxt(txtExcel02)
    End Sub

    Private Sub txtExcel03_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL03.TextChanged
        If TestNum(txtExcel03, "Excel03") Then
        End If
    End Sub
    Private Sub txtExcel03_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL03.GotFocus
        StartEdycjiTxt(txtExcel03)
    End Sub
    Private Sub txtExcel03_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL03.LostFocus
        KoniecEdycjiTxt(txtExcel03)
    End Sub

    Private Sub txtExcel04_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL04.TextChanged
        If TestNum(txtExcel04, "Excel04") Then
        End If
    End Sub
    Private Sub txtExcel04_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL04.GotFocus
        StartEdycjiTxt(txtExcel04)
    End Sub
    Private Sub txtExcel04_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL04.LostFocus
        KoniecEdycjiTxt(txtExcel04)
    End Sub

    Private Sub txtExcel05_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL05.TextChanged
        If TestNum(txtExcel05, "Excel05") Then
        End If
    End Sub
    Private Sub txtExcel05_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL05.GotFocus
        StartEdycjiTxt(txtExcel05)
    End Sub
    Private Sub txtExcel05_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL05.LostFocus
        KoniecEdycjiTxt(txtExcel05)
    End Sub

    Private Sub txtExcel06_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL06.TextChanged
        If TestNum(txtExcel06, "Excel06") Then
        End If
    End Sub
    Private Sub txtExcel06_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL06.GotFocus
        StartEdycjiTxt(txtExcel06)
    End Sub
    Private Sub txtExcel06_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL06.LostFocus
        KoniecEdycjiTxt(txtExcel06)
    End Sub

    Private Sub txtExcel07_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL07.TextChanged
        If TestNum(txtExcel07, "Excel07") Then
        End If
    End Sub
    Private Sub txtExcel07_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL07.GotFocus
        StartEdycjiTxt(txtExcel07)
    End Sub
    Private Sub txtExcel07_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL07.LostFocus
        KoniecEdycjiTxt(txtExcel07)
    End Sub

    Private Sub txtExcel08_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL08.TextChanged
        If TestNum(txtExcel08, "Excel08") Then
        End If
    End Sub
    Private Sub txtExcel08_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL08.GotFocus
        StartEdycjiTxt(txtExcel08)
    End Sub
    Private Sub txtExcel08_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL08.LostFocus
        KoniecEdycjiTxt(txtExcel08)
    End Sub

    Private Sub txtExcel09_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL09.TextChanged
        If TestNum(txtExcel09, "Excel09") Then
        End If
    End Sub
    Private Sub txtExcel09_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL09.GotFocus
        StartEdycjiTxt(txtExcel09)
    End Sub
    Private Sub txtExcel09_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL09.LostFocus
        KoniecEdycjiTxt(txtExcel09)
    End Sub

    Private Sub txtExcel10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEXCEL10.TextChanged
        If TestNum(txtExcel10, "Excel10") Then
        End If
    End Sub
    Private Sub txtExcel10_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL10.GotFocus
        StartEdycjiTxt(txtExcel10)
    End Sub
    Private Sub txtExcel10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEXCEL10.LostFocus
        KoniecEdycjiTxt(txtExcel10)
    End Sub








    Private Sub cmdPracownik_Click(sender As Object, e As EventArgs) Handles cmdPracownik.Click
        oCoMamRobic = "WYBIERAJ"
        oPracownik = Trim(txtPracownik.Text)
        Dim form As New KatalogUzytkownikow
        form.ShowDialog()
        If Len(Trim(oPracownik)) > 0 Then
            txtPracownik.Text = oPracownik
            IdentUzytkownika(txtPracownik.Text)
            lblNazwisko.Text = oNazwaUzytkownika
            lblStanowisko.Text = oFunkcjaUzytkownika
        End If
    End Sub


End Class
