Imports System.Drawing.Printing

Imports System.Data.Oledb
Imports System
Imports System.Reflection ' For Missing.Value and BindingFlags
Imports System.Runtime.InteropServices ' For COMException
'------------------------------------------
'UWAGA :
'Do referencji trzeba DODAC
'   Microsoft Excel 11.0 Object Library
'   Microsoft Office 11.0 Object Library
'-------------------------------------------
Imports Microsoft.Office.Interop.excel
Imports Microsoft.Office.Interop.Outlook






Public Class KatalogKlientowIKontaktow
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Dim _dsKlenci As DataSet
    Friend WithEvents dtpOdDaty As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents dtpDoDaty As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents cmdHarmonogramWizyt As System.Windows.Forms.Button
    Friend WithEvents menuKlienci As ContextMenuStrip
    Friend WithEvents menuEEdytuj As ToolStripMenuItem
    Friend WithEvents menuEDodaj As ToolStripMenuItem
    Friend WithEvents menuEUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents menuEAkcja As ToolStripMenuItem
    Friend WithEvents dagKlienci As DataGridView
    Dim _AktualizujKartKlientow As delegateAktualizuj

    Public Sub New(ByVal ds As DataSet,
                   ByVal Aktualizuj As delegateAktualizuj)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _dsKlenci = ds
        _AktualizujKartKlientow = Aktualizuj
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
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdDrukuj As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblKod As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblMiejscowosc As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lbleMail As System.Windows.Forms.Label
    Friend WithEvents stbKlienci As System.Windows.Forms.StatusBar
    'Friend WithEvents dagKlienci As KartotekaKlientow.Softart.MyDataGrid
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents lblNrDomu As System.Windows.Forms.Label
    Friend WithEvents lblTelefon As System.Windows.Forms.Label
    Friend WithEvents lblUlica As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHandlowa As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdWybierzSchemat As System.Windows.Forms.Button
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblRodzaj As System.Windows.Forms.Label
    Friend WithEvents lblTransakcje As System.Windows.Forms.Label
    Friend WithEvents lblNastKontakt As System.Windows.Forms.Label
    Friend WithEvents lblOstKontakt As System.Windows.Forms.Label
    Friend WithEvents lblOpiekun As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblPotencjal As System.Windows.Forms.Label
    Friend WithEvents stbFiltry As System.Windows.Forms.StatusBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdFi As System.Windows.Forms.Button
    Friend WithEvents txtTOFI As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdOdswiez As System.Windows.Forms.Button
    Friend WithEvents lblIdDomu As System.Windows.Forms.Label
    Friend WithEvents cmdClearColWidth As System.Windows.Forms.Button
    Friend WithEvents HorizScrollTimer As System.Windows.Forms.Timer
    Friend WithEvents lblTemat As System.Windows.Forms.Label
    Friend WithEvents lblData As System.Windows.Forms.Label
    Friend WithEvents lblNotatka As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogKlientowIKontaktow))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbKlienci = New System.Windows.Forms.StatusBar()
        Me.cmdDrukuj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblTemat = New System.Windows.Forms.Label()
        Me.lblData = New System.Windows.Forms.Label()
        Me.lblNotatka = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.lbleMail = New System.Windows.Forms.Label()
        Me.lblTelefon = New System.Windows.Forms.Label()
        Me.lblIdDomu = New System.Windows.Forms.Label()
        Me.txtTOFI = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblPotencjal = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblTransakcje = New System.Windows.Forms.Label()
        Me.lblNastKontakt = New System.Windows.Forms.Label()
        Me.lblOstKontakt = New System.Windows.Forms.Label()
        Me.lblOpiekun = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.lblRodzaj = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa = New System.Windows.Forms.Label()
        Me.lblUlica = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblMiejscowosc = New System.Windows.Forms.Label()
        Me.lblNrDomu = New System.Windows.Forms.Label()
        Me.lblKod = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdWybierzSchemat = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdClearColWidth = New System.Windows.Forms.Button()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.cmdOdswiez = New System.Windows.Forms.Button()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.dtpOdDaty = New System.Windows.Forms.DateTimePicker()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.dtpDoDaty = New System.Windows.Forms.DateTimePicker()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmdHarmonogramWizyt = New System.Windows.Forms.Button()
        Me.menuKlienci = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuESingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEAkcja = New System.Windows.Forms.ToolStripMenuItem()
        Me.dagKlienci = New System.Windows.Forms.DataGridView()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.menuKlienci.SuspendLayout()
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(927, 430)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(80, 23)
        Me.cmdStop.TabIndex = 38
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdStop, "Koniec dzia³ania funkcji." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Powrót do ekranu z którego wywo³ano tê funkcjê.")
        '
        'stbKlienci
        '
        Me.stbKlienci.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbKlienci.Dock = System.Windows.Forms.DockStyle.None
        Me.stbKlienci.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbKlienci.Location = New System.Drawing.Point(8, 371)
        Me.stbKlienci.Name = "stbKlienci"
        Me.stbKlienci.ShowPanels = True
        Me.stbKlienci.Size = New System.Drawing.Size(999, 22)
        Me.stbKlienci.TabIndex = 43
        Me.stbKlienci.Text = "StatusBar1"
        '
        'cmdDrukuj
        '
        Me.cmdDrukuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDrukuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDrukuj.ForeColor = System.Drawing.Color.Black
        Me.cmdDrukuj.Image = CType(resources.GetObject("cmdDrukuj.Image"), System.Drawing.Image)
        Me.cmdDrukuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDrukuj.Location = New System.Drawing.Point(756, 430)
        Me.cmdDrukuj.Name = "cmdDrukuj"
        Me.cmdDrukuj.Size = New System.Drawing.Size(165, 23)
        Me.cmdDrukuj.TabIndex = 42
        Me.cmdDrukuj.Text = "Drukuj Raport Kont. Handl."
        Me.cmdDrukuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDrukuj, "Wydruk Raportu Kontaktów Handlowych.")
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(10, 429)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 39
        Me.PictureBox1.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblTemat)
        Me.Panel1.Controls.Add(Me.lblData)
        Me.Panel1.Controls.Add(Me.lblNotatka)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.lbleMail)
        Me.Panel1.Controls.Add(Me.lblTelefon)
        Me.Panel1.Controls.Add(Me.lblIdDomu)
        Me.Panel1.Controls.Add(Me.txtTOFI)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.lblPotencjal)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.lblTransakcje)
        Me.Panel1.Controls.Add(Me.lblNastKontakt)
        Me.Panel1.Controls.Add(Me.lblOstKontakt)
        Me.Panel1.Controls.Add(Me.lblOpiekun)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.lblRodzaj)
        Me.Panel1.Controls.Add(Me.lblNazwaHandlowa)
        Me.Panel1.Controls.Add(Me.lblUlica)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.lblMiejscowosc)
        Me.Panel1.Controls.Add(Me.lblNrDomu)
        Me.Panel1.Controls.Add(Me.lblKod)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(999, 111)
        Me.Panel1.TabIndex = 46
        '
        'lblTemat
        '
        Me.lblTemat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTemat.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTemat.ForeColor = System.Drawing.Color.Purple
        Me.lblTemat.Location = New System.Drawing.Point(961, 7)
        Me.lblTemat.Name = "lblTemat"
        Me.lblTemat.Size = New System.Drawing.Size(253, 16)
        Me.lblTemat.TabIndex = 42
        '
        'lblData
        '
        Me.lblData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblData.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblData.ForeColor = System.Drawing.Color.Purple
        Me.lblData.Location = New System.Drawing.Point(787, 7)
        Me.lblData.Name = "lblData"
        Me.lblData.Size = New System.Drawing.Size(80, 16)
        Me.lblData.TabIndex = 40
        '
        'lblNotatka
        '
        Me.lblNotatka.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNotatka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNotatka.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNotatka.ForeColor = System.Drawing.Color.Purple
        Me.lblNotatka.Location = New System.Drawing.Point(787, 23)
        Me.lblNotatka.Name = "lblNotatka"
        Me.lblNotatka.Size = New System.Drawing.Size(253, 80)
        Me.lblNotatka.TabIndex = 44
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Navy
        Me.Label17.Location = New System.Drawing.Point(742, 24)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(56, 16)
        Me.Label17.TabIndex = 45
        Me.Label17.Text = "Opis . . . . . . . . "
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(883, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(102, 16)
        Me.Label9.TabIndex = 43
        Me.Label9.Text = "Rodz.kontaktu . . . . . . . . "
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(742, 8)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(56, 16)
        Me.Label18.TabIndex = 41
        Me.Label18.Text = "Data . . . . . . . . "
        '
        'lbleMail
        '
        Me.lbleMail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbleMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lbleMail.ForeColor = System.Drawing.Color.Purple
        Me.lbleMail.Location = New System.Drawing.Point(168, 72)
        Me.lbleMail.Name = "lbleMail"
        Me.lbleMail.Size = New System.Drawing.Size(232, 16)
        Me.lbleMail.TabIndex = 20
        '
        'lblTelefon
        '
        Me.lblTelefon.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTelefon.ForeColor = System.Drawing.Color.Purple
        Me.lblTelefon.Location = New System.Drawing.Point(168, 56)
        Me.lblTelefon.Name = "lblTelefon"
        Me.lblTelefon.Size = New System.Drawing.Size(232, 16)
        Me.lblTelefon.TabIndex = 9
        '
        'lblIdDomu
        '
        Me.lblIdDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblIdDomu.Location = New System.Drawing.Point(355, 40)
        Me.lblIdDomu.Name = "lblIdDomu"
        Me.lblIdDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblIdDomu.TabIndex = 39
        '
        'txtTOFI
        '
        Me.txtTOFI.Location = New System.Drawing.Point(64, 24)
        Me.txtTOFI.Multiline = True
        Me.txtTOFI.Name = "txtTOFI"
        Me.txtTOFI.ReadOnly = True
        Me.txtTOFI.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtTOFI.Size = New System.Drawing.Size(56, 48)
        Me.txtTOFI.TabIndex = 38
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(128, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Telefon . . "
        '
        'lblPotencjal
        '
        Me.lblPotencjal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPotencjal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPotencjal.ForeColor = System.Drawing.Color.Purple
        Me.lblPotencjal.Location = New System.Drawing.Point(504, 72)
        Me.lblPotencjal.Name = "lblPotencjal"
        Me.lblPotencjal.Size = New System.Drawing.Size(232, 16)
        Me.lblPotencjal.TabIndex = 33
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(408, 72)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(104, 16)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Potencja³ . . . ."
        '
        'lblTransakcje
        '
        Me.lblTransakcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTransakcje.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTransakcje.ForeColor = System.Drawing.Color.Purple
        Me.lblTransakcje.Location = New System.Drawing.Point(504, 56)
        Me.lblTransakcje.Name = "lblTransakcje"
        Me.lblTransakcje.Size = New System.Drawing.Size(232, 16)
        Me.lblTransakcje.TabIndex = 19
        '
        'lblNastKontakt
        '
        Me.lblNastKontakt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNastKontakt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNastKontakt.ForeColor = System.Drawing.Color.Purple
        Me.lblNastKontakt.Location = New System.Drawing.Point(504, 40)
        Me.lblNastKontakt.Name = "lblNastKontakt"
        Me.lblNastKontakt.Size = New System.Drawing.Size(232, 16)
        Me.lblNastKontakt.TabIndex = 18
        '
        'lblOstKontakt
        '
        Me.lblOstKontakt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOstKontakt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOstKontakt.ForeColor = System.Drawing.Color.Purple
        Me.lblOstKontakt.Location = New System.Drawing.Point(504, 24)
        Me.lblOstKontakt.Name = "lblOstKontakt"
        Me.lblOstKontakt.Size = New System.Drawing.Size(232, 16)
        Me.lblOstKontakt.TabIndex = 17
        '
        'lblOpiekun
        '
        Me.lblOpiekun.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOpiekun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOpiekun.ForeColor = System.Drawing.Color.Purple
        Me.lblOpiekun.Location = New System.Drawing.Point(504, 8)
        Me.lblOpiekun.Name = "lblOpiekun"
        Me.lblOpiekun.Size = New System.Drawing.Size(232, 16)
        Me.lblOpiekun.TabIndex = 16
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Navy
        Me.Label15.Location = New System.Drawing.Point(8, 72)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(56, 16)
        Me.Label15.TabIndex = 32
        Me.Label15.Text = "Rodz.sp. . . . . . . . . "
        '
        'lblRodzaj
        '
        Me.lblRodzaj.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRodzaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblRodzaj.ForeColor = System.Drawing.Color.Purple
        Me.lblRodzaj.Location = New System.Drawing.Point(64, 72)
        Me.lblRodzaj.Name = "lblRodzaj"
        Me.lblRodzaj.Size = New System.Drawing.Size(56, 16)
        Me.lblRodzaj.TabIndex = 31
        Me.lblRodzaj.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblNazwaHandlowa
        '
        Me.lblNazwaHandlowa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa.Location = New System.Drawing.Point(168, 8)
        Me.lblNazwaHandlowa.Name = "lblNazwaHandlowa"
        Me.lblNazwaHandlowa.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwaHandlowa.TabIndex = 2
        '
        'lblUlica
        '
        Me.lblUlica.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUlica.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUlica.ForeColor = System.Drawing.Color.Purple
        Me.lblUlica.Location = New System.Drawing.Point(168, 40)
        Me.lblUlica.Name = "lblUlica"
        Me.lblUlica.Size = New System.Drawing.Size(144, 16)
        Me.lblUlica.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(128, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Ulica . . "
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Nr TOFI . . . . . . . . "
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(128, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Nazwa . . . . . . . . "
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Navy
        Me.Label14.Location = New System.Drawing.Point(128, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(56, 16)
        Me.Label14.TabIndex = 24
        Me.Label14.Text = "eMail . ."
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Navy
        Me.Label13.Location = New System.Drawing.Point(408, 56)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(104, 16)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "Sprzed.Ost.Trans."
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(408, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(96, 16)
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "Sprzed.Nast.Kont."
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(408, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(104, 16)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "Sprzed.Ost.Kont."
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(408, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 16)
        Me.Label10.TabIndex = 15
        Me.Label10.Text = "Opiekun . . . "
        '
        'lblMiejscowosc
        '
        Me.lblMiejscowosc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMiejscowosc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblMiejscowosc.ForeColor = System.Drawing.Color.Purple
        Me.lblMiejscowosc.Location = New System.Drawing.Point(240, 24)
        Me.lblMiejscowosc.Name = "lblMiejscowosc"
        Me.lblMiejscowosc.Size = New System.Drawing.Size(160, 16)
        Me.lblMiejscowosc.TabIndex = 12
        '
        'lblNrDomu
        '
        Me.lblNrDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNrDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNrDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblNrDomu.Location = New System.Drawing.Point(311, 40)
        Me.lblNrDomu.Name = "lblNrDomu"
        Me.lblNrDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblNrDomu.TabIndex = 11
        '
        'lblKod
        '
        Me.lblKod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKod.ForeColor = System.Drawing.Color.Purple
        Me.lblKod.Location = New System.Drawing.Point(168, 24)
        Me.lblKod.Name = "lblKod"
        Me.lblKod.Size = New System.Drawing.Size(64, 16)
        Me.lblKod.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(128, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Adres"
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(64, 8)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(56, 16)
        Me.lblIdent.TabIndex = 1
        Me.lblIdent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Nr karty . . . . . . . . "
        '
        'cmdWybierzSchemat
        '
        Me.cmdWybierzSchemat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzSchemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierzSchemat.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierzSchemat.Location = New System.Drawing.Point(152, 370)
        Me.cmdWybierzSchemat.Name = "cmdWybierzSchemat"
        Me.cmdWybierzSchemat.Size = New System.Drawing.Size(56, 23)
        Me.cmdWybierzSchemat.TabIndex = 57
        Me.cmdWybierzSchemat.Text = "Wybierz"
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 124)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(999, 22)
        Me.stbFiltry.TabIndex = 67
        Me.stbFiltry.Text = "StatusBar1"
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWszystko.ForeColor = System.Drawing.Color.Black
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(724, 127)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystko.TabIndex = 115
        Me.ToolTip1.SetToolTip(Me.cmdWszystko, "Wyczyœæ wszystkie filtry kolumnowe.")
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(706, 126)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 171
        Me.ToolTip1.SetToolTip(Me.cmdFi, "Unikalne wartoœci pól tej kolumny.")
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 5000
        Me.ToolTip1.InitialDelay = 100
        Me.ToolTip1.IsBalloon = True
        Me.ToolTip1.ReshowDelay = 100
        '
        'cmdClearColWidth
        '
        Me.cmdClearColWidth.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdClearColWidth.ForeColor = System.Drawing.Color.Black
        Me.cmdClearColWidth.Location = New System.Drawing.Point(8, 152)
        Me.cmdClearColWidth.Name = "cmdClearColWidth"
        Me.cmdClearColWidth.Size = New System.Drawing.Size(46, 20)
        Me.cmdClearColWidth.TabIndex = 174
        Me.cmdClearColWidth.Text = "|<==>|"
        Me.ToolTip1.SetToolTip(Me.cmdClearColWidth, "Powrót do pierwotnej szerokoœci kolumn")
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.ForeColor = System.Drawing.Color.Black
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(927, 401)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytuj.TabIndex = 182
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdEdytuj, "Edycja informacji o kontakcie wskazywanym przez kursor.")
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Enabled = False
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.ForeColor = System.Drawing.Color.Black
        Me.cmdUsuñ.Image = CType(resources.GetObject("cmdUsuñ.Image"), System.Drawing.Image)
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(841, 401)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 23)
        Me.cmdUsuñ.TabIndex = 181
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdUsuñ, "Usuniecie informacji o kontakcie handlowym wskazywanym przez kursor" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " Kartoteki.")
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.ForeColor = System.Drawing.Color.Black
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(756, 401)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodaj.TabIndex = 180
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDodaj, "Dopisanie (dodanie) nowego kontaktu handlowego do Kartoteki.")
        '
        'cmdOdswiez
        '
        Me.cmdOdswiez.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdOdswiez.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdOdswiez.ForeColor = System.Drawing.Color.Black
        Me.cmdOdswiez.Image = CType(resources.GetObject("cmdOdswiez.Image"), System.Drawing.Image)
        Me.cmdOdswiez.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOdswiez.Location = New System.Drawing.Point(461, 430)
        Me.cmdOdswiez.Name = "cmdOdswiez"
        Me.cmdOdswiez.Size = New System.Drawing.Size(88, 23)
        Me.cmdOdswiez.TabIndex = 173
        Me.cmdOdswiez.Text = "Pobierz"
        Me.cmdOdswiez.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdswiez, "Wybierz z Kartoteki Klientów klientów analizowanych kontaktów handlowych.")
        '
        'HorizScrollTimer
        '
        '
        'dtpOdDaty
        '
        Me.dtpOdDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpOdDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpOdDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOdDaty.Location = New System.Drawing.Point(59, 1)
        Me.dtpOdDaty.Name = "dtpOdDaty"
        Me.dtpOdDaty.Size = New System.Drawing.Size(91, 20)
        Me.dtpOdDaty.TabIndex = 175
        Me.dtpOdDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Navy
        Me.Label19.Location = New System.Drawing.Point(10, 5)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(56, 16)
        Me.Label19.TabIndex = 176
        Me.Label19.Text = "Od daty . . . . . . . . . "
        '
        'dtpDoDaty
        '
        Me.dtpDoDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDoDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpDoDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDoDaty.Location = New System.Drawing.Point(209, 1)
        Me.dtpDoDaty.Name = "dtpDoDaty"
        Me.dtpDoDaty.Size = New System.Drawing.Size(91, 20)
        Me.dtpDoDaty.TabIndex = 177
        Me.dtpDoDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Navy
        Me.Label20.Location = New System.Drawing.Point(160, 5)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(56, 16)
        Me.Label20.TabIndex = 178
        Me.Label20.Text = "Do daty . . . . . . . . . "
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.dtpOdDaty)
        Me.Panel2.Controls.Add(Me.dtpDoDaty)
        Me.Panel2.Controls.Add(Me.Label19)
        Me.Panel2.Controls.Add(Me.Label20)
        Me.Panel2.Location = New System.Drawing.Point(150, 429)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(305, 24)
        Me.Panel2.TabIndex = 179
        '
        'cmdHarmonogramWizyt
        '
        Me.cmdHarmonogramWizyt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdHarmonogramWizyt.ForeColor = System.Drawing.Color.Black
        Me.cmdHarmonogramWizyt.Image = CType(resources.GetObject("cmdHarmonogramWizyt.Image"), System.Drawing.Image)
        Me.cmdHarmonogramWizyt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdHarmonogramWizyt.Location = New System.Drawing.Point(120, 430)
        Me.cmdHarmonogramWizyt.Name = "cmdHarmonogramWizyt"
        Me.cmdHarmonogramWizyt.Size = New System.Drawing.Size(24, 23)
        Me.cmdHarmonogramWizyt.TabIndex = 205
        Me.cmdHarmonogramWizyt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'menuKlienci
        '
        Me.menuKlienci.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuEEdytuj, Me.menuEDodaj, Me.menuEUsun, Me.ToolStripSeparator2, Me.menuESingleL, Me.menuEMultiL, Me.ToolStripSeparator3, Me.menuEOdswiez, Me.menuEAkcja})
        Me.menuKlienci.Name = "menuEdytuj"
        Me.menuKlienci.Size = New System.Drawing.Size(226, 170)
        '
        'menuEEdytuj
        '
        Me.menuEEdytuj.Name = "menuEEdytuj"
        Me.menuEEdytuj.Size = New System.Drawing.Size(225, 22)
        Me.menuEEdytuj.Text = "Edytuj"
        '
        'menuEDodaj
        '
        Me.menuEDodaj.Name = "menuEDodaj"
        Me.menuEDodaj.Size = New System.Drawing.Size(225, 22)
        Me.menuEDodaj.Text = "Dodaj"
        '
        'menuEUsun
        '
        Me.menuEUsun.Name = "menuEUsun"
        Me.menuEUsun.Size = New System.Drawing.Size(225, 22)
        Me.menuEUsun.Text = "Usuñ"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(222, 6)
        '
        'menuESingleL
        '
        Me.menuESingleL.Name = "menuESingleL"
        Me.menuESingleL.Size = New System.Drawing.Size(225, 22)
        Me.menuESingleL.Text = "Poka¿ w jednej linii"
        '
        'menuEMultiL
        '
        Me.menuEMultiL.Name = "menuEMultiL"
        Me.menuEMultiL.Size = New System.Drawing.Size(225, 22)
        Me.menuEMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(222, 6)
        '
        'menuEOdswiez
        '
        Me.menuEOdswiez.Name = "menuEOdswiez"
        Me.menuEOdswiez.Size = New System.Drawing.Size(225, 22)
        Me.menuEOdswiez.Text = "Odœwie¿ Tabelê"
        '
        'menuEAkcja
        '
        Me.menuEAkcja.Name = "menuEAkcja"
        Me.menuEAkcja.Size = New System.Drawing.Size(225, 22)
        Me.menuEAkcja.Text = "Wybierz akcjê marketingow¹"
        '
        'dagKlienci
        '
        Me.dagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagKlienci.ContextMenuStrip = Me.menuKlienci
        Me.dagKlienci.Location = New System.Drawing.Point(8, 150)
        Me.dagKlienci.Name = "dagKlienci"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagKlienci.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagKlienci.Size = New System.Drawing.Size(999, 221)
        Me.dagKlienci.TabIndex = 207
        '
        'KatalogKlientowIKontaktow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 459)
        Me.Controls.Add(Me.cmdClearColWidth)
        Me.Controls.Add(Me.dagKlienci)
        Me.Controls.Add(Me.cmdHarmonogramWizyt)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.cmdOdswiez)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdDrukuj)
        Me.Controls.Add(Me.cmdWybierzSchemat)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.stbKlienci)
        Me.Controls.Add(Me.PictureBox1)
        Me.ForeColor = System.Drawing.Color.Teal
        Me.Location = New System.Drawing.Point(8, 8)
        Me.Name = "KatalogKlientowIKontaktow"
        Me.Text = "Kontakty Handlowe z Klientami"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.menuKlienci.ResumeLayout(False)
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim StartFormy As Boolean = True
    Dim _OCoMamRobic As String = oCoMamRobic

    '---------------------------------------------
    'SCROLL-parametry obslugi scrollingu poziomego
    Dim StartPointInCell00 As System.Drawing.Point             'pocz¹tek ektanu przed scrollingiem
    Dim ScrollXOffset As Integer = 0            'bie¿¹cy punkt ekranu wskazywany przez kursor podczas scrollingu
    Dim HorizScrollLastTime As String = ""      'Chwila rozpoczecia scrollingu - filtry wyswietlaner bêd¹ po 1 sekundzie...
    ' HorizScrollTimer - zegar odliczaj¹cy czas od poczatku scrollingu
    '---------------------------------------------




    Dim dbSelectKlienci As String = "SELECT * FROM Klienci"
    Dim dbSelectKlienciSorted As String = "SELECT * FROM Klienci"

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

    Dim dSetKlienci As New DataSet
    Dim dViewKlienci As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim _CoMamRobic As String = ""
    Dim OdDaty As String = ""
    Dim DoDaty As String = ""
    Dim TakenDate As Date = Nothing

    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub KatalogKlientowIKontaktow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.ClosedByControlBox Then
            'e.Cancel = True     'nie pozwalaj zamkn¹c forme
            'MessageBox.Show("Proszê zamkn¹æ formê klawiszami ZAPISZ lub WYCOFAJ SIÊ", _
            '    "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    MessageBoxIcon.Information, _
            '    MessageBoxDefaultButton.Button1)

            ZapamietajSzerokosciKolumn()
        Else
            'MsgBox("ZAPISZ lub WYCOFAJ SIÊ")
        End If
    End Sub
    '====================================================



    ''---------------------------------------------------------------------
    ''Klienci
    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    'Public oPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
    'Public oTematKon As String            'TEMAT            Text, 50
    'Public oUwagiKon As String            'UWAGI            Memo
    'Public oWersjaKon As Integer          'WERSJA           Integer

    Private Sub KatalogKlientowIKontaktow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        StartFormy = True
        '-------------------------------
        If _dsKlenci Is Nothing Then
            cmdOdswiez.Visible = False
        End If
        '-------------------------------
        _CoMamRobic = oCoMamRobic
        dtpOdDaty.Value = Mid(SysData(), 1, 7) & "-01"
        dtpDoDaty.Value = SysData()
        '---------------------------------------
        TakenDate = dtpOdDaty.Value
        OdDaty = TakenDate.ToString("yyyy-MM-dd")
        TakenDate = dtpDoDaty.Value
        DoDaty = TakenDate.ToString("yyyy-MM-dd")
        PokazKontakty(OdDaty, DoDaty)
        '---------------------------------------
        IdentUzytkownika(Program_IdUzytkownika)
        If OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownik Then
            cmdUsuñ.Enabled = False
            menuEUsun.Enabled = False

            cmdEdytuj.Text = "Poka¿"
        End If

        StartFormy = False
        'RysujGradientWPanelu(Panel2)
        'RysujGradientWPanelu(Panel3)
        'RysujGradientWPanelu(Panel4)
        'RysujGradientWPanelu(Panel5)
        Me.WindowState = FormWindowState.Normal
        InicjujFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        'RowHeightSetter.MultiHightsView()
        dagKlienci_CurrentCellChanged(sender, e)
        '--------------------------
        'zegar do scrollingu poziomego
        HorizScrollLastTime = ""
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
    End Sub





    Private Sub PokazKontakty(ByVal OdDaty As String, ByVal DoDaty As String)
        StartFormy = True


        dbSelectKlienci = "SELECT " &
                                   "KH.IDENTKLIENTA, " &
                                   "KH.DATAKONTAKTU," &
                                   "KH.PRACOWNIK," &
                                   "KH.TEMAT," &
                                   "KH.UWAGI," &
                                       "KL.IDENTKLIENTA AS NRKARTY, " &
                                       "KL.NIP," &
                                       "KL.NRTOFITXT," &
                                       "KL.KARTAPAYBACK," &
                                       "KL.NRYKARTPAYBACK," &
                                       "KL.NAZWA1 AS NAZWAKLIENTA," &
                                       "KL.MIEJSCOWOSC," &
                                       "KL.KODPOCZTOWY," &
                                       "KL.ULICA," &
                                       "KL.NUMNRDOMU, "

        If ParL_DataBaseType = "MS ACCESS" Then
            dbSelectKlienci &= "IIF(NUMNRDOMU MOD 2=0,'True','False') AS PARZYSTE, "
        Else
            dbSelectKlienci &= "CASE WHEN NUMNRDOMU%2=0 THEN CAST('True' AS BIT) ELSE CAST('False' AS BIT) END AS PARZYSTE, "
        End If


        dbSelectKlienci &=
                                               "KL.NRDOMU," &
                                               "KL.IDDOMU," &
                                               "KL.REJKLIENTA AS REJONKLIENTA," &
                                               "KL.TELEFON," &
                                               "KL.FAX," &
                                               "KL.EMAIL," &
                                       "KL.WYKAZURZADZEN," &
                                         "KL.SPOSOBWYBORUDOSTAWCY," &
                                         "KL.ZAINSTALOWANOMONITOR," &
                                         "KL.LICZBAURZADZEN," &
                                         "KL.LICZBAWYDRUKOW," &
                                         "KL.POTENCJALDRUKU," &
                                       "KL.AKTZAKUPMATEKSP," &
                                       "KL.AKTROZLICZASTRONYDRUKU," &
                                       "KL.AKTKORZYSTAZNAJMU," &
                                       "KL.ZAINTROZLICZANIEMSTRONDRUKU," &
                                       "KL.ZAINTKORZYSTANIEMZNAJMU," &
                                         "KL.DATAWERYFSEGMENTACJI," &
                                         "KL.PLANDLUGOOKRESOWY," &
                                         "KL.PLANKROTKOOKRESOWY," &
                                               "KL.RODZSPRZETUL," &
                                               "KL.RODZSPRZETUA," &
                                               "KL.RODZSPRZETUT," &
                                               "KL.OFERTATZLOZENIA," &
                                               "KL.OFERTAPLIK," &
                                               "KL.PKONTAKT," &
                                               "KL.SKUPUWAGI AS PROMOTORUWAGI," &
                                               "KL.SPRZEDOPIEKUN AS OPIEKUN," &
                                               "KL.SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," &
                                               "KL.SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," &
                                               "KL.SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," &
                                               "KL.SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," &
                                               "KL.SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," &
                                               "KL.SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," &
                                               "KL.SPRZEDUWAGI AS OPIEKUNUWAGI," &
                                               "KL.KTOWPISAL," &
                                               "KL.UWAGI AS UWAGIKLI," &
                                               "KL.MARKER, " &
                                               "KL.MARKERP, " &
                                               "KL.AKTYWNY, " &
                                             "KH.UNIQID, " &
                                             "KH.WERSJA " &
                                            "FROM KontaktyHandlowe as KH INNER JOIN Klienci AS KL ON KH.IDENTKLIENTA=KL.IDENTKLIENTA " &
                                            "WHERE (KH.DATAKONTAKTU>='" & OdDaty & "') AND (KH.DATAKONTAKTU<='" & DoDaty & "') AND (KL.AKTYWNY='TRUE') "
        dbSelectKlienciSorted = dbSelectKlienci & " ORDER BY KH.DATAKONTAKTU DESC, KH.IDENTKLIENTA"




        dSetKlienci = New DataSet
        dViewKlienci = New DataView
        dSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectKlienciSorted, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            'DBDeleteKontaktyHandlowe(dbCommandDelete, dbDataAdapter)
            'DBUpdateKontaktyHandlowe(dbCommandUpdate, dbDataAdapter)
            'DBInsertKontaktyHandlowe(dbCommandInsert, dbDataAdapter)

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
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnection.State
            'End Try


        Else
            sqlConnection = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelect = New SqlClient.SqlCommand(dbSelectKlienciSorted, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLKontakty, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLKontakty, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLKontakty, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            SQLDeleteKontaktyHandlowe(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateKontaktyHandlowe(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertKontaktyHandlowe(sqlCommandInsert, sqlDataAdapter)

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
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnection.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapter.FillSchema(dSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapter.Fill(dSetKlienci)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(dSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapter.Fill(dSetKlienci)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                dSetKlienci.Tables(0).PrimaryKey = New DataColumn() {dSetKlienci.Tables(0).Columns("UNIQID")}

                'definiuj DataView
                dViewKlienci = New DataView(dSetKlienci.Tables(0))
                dViewKlienci.AllowDelete = False
                dViewKlienci.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagKlienci.BackgroundColor = System.Drawing.SystemColors.Control
                dagKlienci.GridColor = System.Drawing.SystemColors.ControlDark
                dagKlienci.ForeColor = System.Drawing.Color.Purple
                dagKlienci.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagKlienci.BorderStyle = BorderStyle.Fixed3D
                'dagKlienci.Dock = DockStyle.Fill
                dagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagKlienci.AllowUserToAddRows = False
                dagKlienci.AllowUserToDeleteRows = False
                dagKlienci.AllowUserToOrderColumns = True
                dagKlienci.AllowUserToResizeColumns = True
                dagKlienci.AllowUserToResizeRows = True
                dagKlienci.ReadOnly = True
                dagKlienci.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagKlienci.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagKlienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagKlienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagKlienci.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagKlienci.ColumnHeadersVisible = True
                dagKlienci.ColumnHeadersHeight = 40
                dagKlienci.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagKlienci.RowHeadersVisible = True
                dagKlienci.RowHeadersWidth = 50
                dagKlienci.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagKlienci.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagKlienci.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagKlienci.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagKlienci.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagKlienci.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagKlienci.DefaultCellStyle.NullValue = ""
                dagKlienci.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagKlienci.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagKlienci.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagKlienci.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagKlienci.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagKlienci.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagKlienci.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagKlienci.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagKlienci.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagKlienci.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                'dagKlienci.DataMember = dSetKlienci.Tables(0).TableName
                'dagKlienci.DataSource = dViewKlienci
                '-----------------------------------

                'dbSelectKlienci = "SELECT " &
                '                   "KH.IDENTKLIENTA, " &
                '                   "KH.DATAKONTAKTU," &
                '                   "KH.PRACOWNIK," &
                '                   "KH.TEMAT," &
                '                   "KH.UWAGI," &
                '                       "KL.IDENTKLIENTA AS NRKARTY, " &
                '                       "KL.NIP," &
                '                       "KL.NRTOFITXT," &
                '                       "KL.KARTAPAYBACK," &
                '                       "KL.NRYKARTPAYBACK," &
                '                       "KL.NAZWA1 AS NAZWAKLIENTA," &
                '                       "KL.MIEJSCOWOSC," &
                '                       "KL.KODPOCZTOWY," &
                '                       "KL.ULICA," &
                '                       "KL.NUMNRDOMU, "

                'If ParL_DataBaseType = "MS ACCESS" Then
                '    dbSelectKlienci &= "IIF(NUMNRDOMU MOD 2=0,'True','False') AS PARZYSTE, "
                'Else
                '    dbSelectKlienci &= "CASE WHEN NUMNRDOMU%2=0 THEN CAST('True' AS BIT) ELSE CAST('False' AS BIT) END AS PARZYSTE, "
                'End If


                'dbSelectKlienci &=
                '                               "KL.NRDOMU," &
                '                               "KL.IDDOMU," &
                '                               "KL.REJKLIENTA AS REJONKLIENTA," &
                '                               "KL.TELEFON," &
                '                               "KL.FAX," &
                '                               "KL.EMAIL," &
                '                       "KL.WYKAZURZADZEN," &
                '                         "KL.SPOSOBWYBORUDOSTAWCY," &
                '                         "KL.ZAINSTALOWANOMONITOR," &
                '                         "KL.LICZBAURZADZEN," &
                '                         "KL.LICZBAWYDRUKOW," &
                '                         "KL.POTENCJALDRUKU," &
                '                       "KL.AKTZAKUPMATEKSP," &
                '                       "KL.AKTROZLICZASTRONYDRUKU," &
                '                       "KL.AKTKORZYSTAZNAJMU," &
                '                       "KL.ZAINTROZLICZANIEMSTRONDRUKU," &
                '                       "KL.ZAINTKORZYSTANIEMZNAJMU," &
                '                         "KL.DATAWERYFSEGMENTACJI," &
                '                         "KL.PLANDLUGOOKRESOWY," &
                '                         "KL.PLANKROTKOOKRESOWY," &
                '                               "KL.RODZSPRZETUL," &
                '                               "KL.RODZSPRZETUA," &
                '                               "KL.RODZSPRZETUT," &
                '                               "KL.OFERTATZLOZENIA," &
                '                               "KL.OFERTAPLIK," &
                '                               "KL.PKONTAKT," &
                '                               "KL.SKUPUWAGI AS PROMOTORUWAGI," &
                '                               "KL.SPRZEDOPIEKUN AS OPIEKUN," &
                '                               "KL.SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," &
                '                               "KL.SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," &
                '                               "KL.SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," &
                '                               "KL.SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," &
                '                               "KL.SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," &
                '                               "KL.SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," &
                '                               "KL.SPRZEDUWAGI AS OPIEKUNUWAGI," &
                '                               "KL.KTOWPISAL," &
                '                               "KL.UWAGI AS UWAGIKLI," &
                '                               "KL.MARKER, " &
                '                               "KL.MARKERP, " &
                '                               "KL.AKTYWNY, " &
                '                             "KH.UNIQID, " &
                '                             "KH.WERSJA " &
                '                            "FROM KontaktyHandlowe as KH INNER JOIN Klienci AS KL ON KH.IDENTKLIENTA=KL.IDENTKLIENTA " &
                '                            "WHERE (KH.DATAKONTAKTU>='" & OdDaty & "') AND (KH.DATAKONTAKTU<='" & DoDaty & "') AND (KL.AKTYWNY='TRUE') "
                'dbSelectKlienciSorted = dbSelectKlienci & " ORDER BY KH.DATAKONTAKTU DESC, KH.IDENTKLIENTA"



                dagKlienci.Columns.Clear()

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(1).ColumnName, "Data kontaktu", 70, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(2).ColumnName, "Pracownik", 50, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(3).ColumnName, "Rodzaj kontaktu", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(4).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                '-------------------------------------
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(5).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(6).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(7).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(8).ColumnName, "PayBack", 50, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(9).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(10).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(11).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(12).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(13).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(14).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(15).ColumnName, "Parzysty", 40, True)

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(16).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(17).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(18).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(19).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(20).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(21).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(22).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(23).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(24).ColumnName, "Zainst. monitor", 50, True)
                NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(25).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(26).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(27).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(28).ColumnName, "Akt. zakup mat.eksp.", 50, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(29).ColumnName, "Akt. rozlicza strony druku", 50, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(30).ColumnName, "Akt. korzysta z najmu", 50, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(31).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(32).ColumnName, "Zaint. korzyst. z najmu", 50, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(33).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)
                NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(34).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(35).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(36).ColumnName, "Rodz.Sprzêtu L", 100, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(37).ColumnName, "Rodz.Sprzêtu A", 100, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(38).ColumnName, "Rodz.Sprzêtu A3", 100, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(39).ColumnName, "Z³o¿ono", 40, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(40).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(41).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(42).ColumnName, "Promotor Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(43).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(44).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(45).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(46).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(47).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(48).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(49).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(50).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(51).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(52).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(53).ColumnName, "Marker  ", 50, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(54).ColumnName, "Marker P.", 50, True)
                LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(55).ColumnName, "", 0, False)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(56).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(57).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)



                Me.Cursor = Cursors.WaitCursor
                dagKlienci.DataSource = dViewKlienci
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If dSetKlienci.Tables(0).Rows.Count > 0 Then
                    dagKlienci.CurrentCell = dagKlienci(1, 0)
                    dagKlienci.CurrentCell.Selected = True
                End If

                UstalSzerokosciKolumn()




            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End Try

            '------------------------------------
            'dodaj StatusBar na koncu
            stbKlienci.Panels.Clear()

            stbKlienci.BackColor = System.Drawing.SystemColors.Control
            stbKlienci.ForeColor = System.Drawing.Color.DarkGreen
            stbKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbKlienci.Panels.Add("stbPanelIleRec")
            stbKlienci.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(0).Width = 80
            stbKlienci.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(0).Text = "Iloœæ rec.: " & dSetKlienci.Tables(0).Rows.Count.ToString

            stbKlienci.Panels.Add("stbPanelWiersz")
            stbKlienci.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(1).Width = 80
            stbKlienci.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(1).Text = "Wi: "
                'stbKlienci.Panels(2).Text = "Ko: "
            Else
                stbKlienci.Panels(1).Text = "Wi: " & dagKlienci.CurrentCell.RowIndex.ToString
                'stbKlienci.Panels(2).Text = "Ko: " & dagKlienci.CurrentCell.ColumnIndex.ToString
            End If


            stbKlienci.Panels.Add("stbPanelKolumna")
            stbKlienci.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(2).Width = 80
            stbKlienci.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagKlienci.CurrentCell Is Nothing Then
                'stbKlienci.Panels(1).Text = "Wi: "
                stbKlienci.Panels(2).Text = "Ko: "
            Else
                'stbKlienci.Panels(1).Text = "Wi: " & dagKlienci.CurrentCell.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & dagKlienci.CurrentCell.ColumnIndex.ToString
            End If

            stbKlienci.Panels.Add("stbPanelSort")
            stbKlienci.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(3).Width = 150
            stbKlienci.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(3).Text = "Sort : "

            stbKlienci.Panels.Add("stbPanelFiltr")
            stbKlienci.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(4).Width = 150
            stbKlienci.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(4).Text = "Filtr : "




            stbKlienci.Panels.Add("stbPanelSzablon")
            stbKlienci.Panels(5).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(5).Width = 150
            stbKlienci.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(5).Text = "Szablon: "

            'dodaj ctrl do pobierania szablonu
            cmdWybierzSchemat.Location = New System.Drawing.Point(stbKlienci.Location.X +
                                                stbKlienci.Panels(0).Width +
                                                stbKlienci.Panels(1).Width +
                                                stbKlienci.Panels(2).Width +
                                                stbKlienci.Panels(3).Width +
                                                stbKlienci.Panels(4).Width, stbKlienci.Location.Y + 2)
            cmdWybierzSchemat.Size = New Size(60, 20)
            cmdWybierzSchemat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)



            stbKlienci.Panels.Add("stbPanelReszta")
            stbKlienci.Panels(6).AutoSize = StatusBarPanelAutoSize.Spring
            'stbKlienci.Panels(6).Width = 150
            stbKlienci.Panels(6).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(6).Text = ""

            stbKlienci.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.Panels.Clear()

            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagKlienci.TableStyles(0).RowHeaderWidth
            stbFiltry.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbFiltry.Panels(0).Text = "Filtry"

            stbFiltry.Panels.Add("stbFiltryReszta")
            stbFiltry.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
            'stbFiltry.Panels(1).Width = 150
            stbFiltry.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbFiltry.Panels(1).Text = ""

            stbFiltry.ShowPanels = True

            'dodaj ctrl do pobierania szablonu
            'cmdWszystko.Size = New Size(20, 20)
            'cmdWszystko.Location = New system.drawing.Point(stbFiltry.Location.X + _
            '                                 stbFiltry.Size.Width - _
            '                                 cmdWszystko.Size.Width, _
            '                                 stbFiltry.Location.Y + 2)
            'cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
            '                            System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            cmdWszystko.Size = New Size(20, 22)
            cmdWszystko.Location = New System.Drawing.Point(stbFiltry.Location.X + 50 - 20,
                                             stbFiltry.Location.Y)
            cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

        End If
        InitKlienci() 'inicjuj zmienne
        '-----------------------------------------------------------------
        'If _OCoMamRobic = "WYBIERAJ" Then
        '    dagKlienci.ContextMenuStrip = Me.menuWybieraj
        '    cmdStop.Text = "Wybierz"
        '    menuWybieraj.Visible = False
        '    menuWybieraj.Enabled = True
        '    MenuEdytuj.Visible = False
        '    MenuEdytuj.Enabled = False
        'Else
        '    dagKlienci.ContextMenuStrip = Me.MenuEdytuj
        '    cmdStop.Text = "Powrót"
        '    menuWybieraj.Visible = False
        '    menuWybieraj.Enabled = False
        '    MenuEdytuj.Visible = False
        '    MenuEdytuj.Enabled = True
        'End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagKlienci.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagKlienci.RowHeaderWidth = 80 'use if no tablestyle
        cmdClearColWidth.Size = New Size(50, 40)
        cmdClearColWidth.Location = New System.Drawing.Point(dagKlienci.Location.X,
                                                             dagKlienci.Location.Y)
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If dSetKlienci.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagKlienci.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagKlienci.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                        dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagKlienci.Left + 4), (dagKlienci.Top + 4))
            ScrollXOffset = 0
        End If
        '--------------------------------------------------
        'inicjowanie zegara dla scrollingu poziomego
        HorizScrollLastTime = ""
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
        '--------------------------------------------------
        'Stworz zmienne do filtrowania
        Dim ii As Integer = 0
        For ii = 0 To dViewKlienci.Table.Columns.Count - 1
            'stworz zmienn¹ TXTBOX
            Dim CtrlT As New System.Windows.Forms.TextBox
            CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlT.Visible = False
            Me.Controls.Add(CtrlT)
            AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXX_TextChanged)
            AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXX_GotFocus)
            AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXX_LostFocus)
            AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

            'stworz zmienn¹ BUTTON
            Dim CtrlB As New System.Windows.Forms.Button
            CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlB.Image = cmdFi.Image
            CtrlB.ImageAlign = ContentAlignment.MiddleCenter
            CtrlB.Visible = False
            Me.Controls.Add(CtrlB)
            AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXX_Click)
            AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXX_GotFocus)
            AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXX_LostFocus)
        Next
        Me.Refresh()
        '--------------------------------------------------
        StartFormy = False
        InicjujFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        dagKlienci_CurrentCellChanged(Me, Nothing)
    End Sub

    Private Sub KatalogKlientowIKontaktow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub





















    Private Sub dagKlienci_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.CurrentCellChanged
        Dim Txt As String = ""
        Dim words() As String
        Dim splitChars() As Char = {","c}
        Dim iwords As Integer = 0

        If Not StartFormy Then
            If dagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(1).Text = "Wi: "
                stbKlienci.Panels(2).Text = "Ko: "

                lblIdent.Text = ""
                lblNazwaHandlowa.Text = ""
                txtTOFI.Text = ""
                lblRodzaj.Text = ""

                lblKod.Text = ""
                lblMiejscowosc.Text = ""
                lblUlica.Text = ""
                lblNrDomu.Text = ""
                lblIdDomu.Text = ""
                lblTelefon.Text = ""
                lbleMail.Text = ""

                lblOstKontakt.Text = ""
                lblNastKontakt.Text = ""
                lblOpiekun.Text = ""
                lblPotencjal.Text = ""
                lblTransakcje.Text = ""

                lblData.Text = ""
                lblTemat.Text = ""
                lblNotatka.Text = ""
            Else
                stbKlienci.Panels(1).Text = "Wi: " & dagKlienci.CurrentCell.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & dagKlienci.CurrentCell.ColumnIndex.ToString

                If dViewKlienci.Count = 0 Then
                    lblIdent.Text = ""
                    lblNazwaHandlowa.Text = ""
                    txtTOFI.Text = ""
                    lblRodzaj.Text = ""

                    lblKod.Text = ""
                    lblMiejscowosc.Text = ""
                    lblUlica.Text = ""
                    lblNrDomu.Text = ""
                    lblIdDomu.Text = ""
                    lblTelefon.Text = ""
                    lbleMail.Text = ""

                    lblOstKontakt.Text = ""
                    lblNastKontakt.Text = ""
                    lblOpiekun.Text = ""
                    lblPotencjal.Text = ""
                    lblTransakcje.Text = ""

                    lblData.Text = ""
                    lblTemat.Text = ""
                    lblNotatka.Text = ""
                Else
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(1).ColumnName, "Data kontaktu", 70, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(2).ColumnName, "Pracownik", 50, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(3).ColumnName, "Rodzaj kontaktu", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(4).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    ''-------------------------------------
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(5).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(6).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(7).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(8).ColumnName, "PayBack", 50, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(9).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(10).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(11).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(12).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(13).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(14).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(15).ColumnName, "Parzysty", 40, True)

                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(16).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(17).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(18).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(19).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(20).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(21).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(22).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(23).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(24).ColumnName, "Zainst. monitor", 50, True)
                    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(25).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(26).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(27).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(28).ColumnName, "Akt. zakup mat.eksp.", 50, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(29).ColumnName, "Akt. rozlicza strony druku", 50, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(30).ColumnName, "Akt. korzysta z najmu", 50, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(31).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(32).ColumnName, "Zaint. korzyst. z najmu", 50, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(33).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)
                    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(34).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(35).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(36).ColumnName, "Rodz.Sprzêtu L", 100, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(37).ColumnName, "Rodz.Sprzêtu A", 100, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(38).ColumnName, "Rodz.Sprzêtu A3", 100, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(41).ColumnName, "Z³o¿ono", 40, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(42).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(43).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(44).ColumnName, "Promotor Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(45).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(46).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(47).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(48).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(49).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(50).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(51).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(52).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(53).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(54).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(55).ColumnName, "Marker  ", 50, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(56).ColumnName, "Marker P.", 50, True)
                    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(57).ColumnName, "", 0, False)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(58).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(59).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)




                    lblIdent.Text = GetTxtField(dagKlienci, 0)

                    'podziel linie na poszczególne czêœci
                    Txt = GetTxtField(dagKlienci, 7)
                    If Len(Txt) = 0 Then
                        txtTOFI.Text = ""
                    Else
                        words = Txt.Split(splitChars, StringSplitOptions.None)
                        If words.Length = 0 Then
                            txtTOFI.Text = ""
                        ElseIf words.Length = 1 Then
                            txtTOFI.Text = Txt
                        Else
                            For iwords = 0 To words.Length - 1
                                If Len(words(iwords)) > 0 Then
                                    txtTOFI.Text &= words(iwords) & vbCrLf
                                End If
                            Next
                        End If
                    End If

                    lblNazwaHandlowa.Text = GetTxtField(dagKlienci, 10)


                    oRodzSprzetuLKli = GetLogField(dagKlienci, 36)
                    oRodzSprzetuAKli = GetLogField(dagKlienci, 37)
                    oRodzSprzetuTKli = GetLogField(dagKlienci, 38)
                    lblRodzaj.Text = IIf(oRodzSprzetuLKli, "L", "") & IIf(oRodzSprzetuAKli, "A", "") & IIf(oRodzSprzetuTKli, "T", "")

                    lblKod.Text = GetTxtField(dagKlienci, 12)
                    lblMiejscowosc.Text = GetTxtField(dagKlienci, 11)
                    lblUlica.Text = GetTxtField(dagKlienci, 13)
                    lblNrDomu.Text = GetIntField(dagKlienci, 14).ToString("F0")
                    lblIdDomu.Text = GetTxtField(dagKlienci, 17)
                    lblTelefon.Text = GetTxtField(dagKlienci, 19)
                    lbleMail.Text = GetTxtField(dagKlienci, 21)


                    lblOpiekun.Text = GetTxtField(dagKlienci, 43)
                    lblOstKontakt.Text = "Tydzieñ " & Trim(GetTxtField(dagKlienci, 45)) & " / " & GetTxtField(dagKlienci, 44)
                    lblNastKontakt.Text = "Tydzieñ " & Trim(GetTxtField(dagKlienci, 48)) & " / " & GetTxtField(dagKlienci, 47)
                    lblTransakcje.Text = OstTransakcje(lblIdent.Text, True)
                    lblPotencjal.Text = GetTxtField(dagKlienci, 42)

                    lblData.Text = GetTxtField(dagKlienci, 1)
                    lblTemat.Text = GetTxtField(dagKlienci, 3)
                    lblNotatka.Text = GetTxtField(dagKlienci, 4)
                End If
            End If
        End If
    End Sub










    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
        HorizScrollTimer.Enabled = False
        If Len(HorizScrollLastTime) = 0 Then
            'nie zainicjowano scrollingu
        Else
            'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
            If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
                If dagKlienci.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                                    dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                                    dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub



    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagKlienci.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagKlienci.Rows(nextrow).Selected = True
            End If
        End If
    End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXX_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormy Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagKlienci, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            'FiltrujDataViewDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, dViewKlienci, stbKlienci)
            FiltrujDataViewDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, dViewKlienci, stbKlienci, oSchematFiltrowania)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagKlienci, Mid(sender.name, 6, 3), "Klienci")
    End Sub
    Private Sub cmdFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me, sender)
    End Sub
    Private Sub cmdFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me, sender)
    End Sub




    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystko.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormy = True
            CzyscFiltryDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, dViewKlienci, stbKlienci)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbKlienci.Panels(0).Text = "Il.zap.: " & dViewKlienci.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagKlienci_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagKlienci.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagKlienci.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagKlienci.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltry_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltry.PanelClick
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub


    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    Public Function Koloruj(ByVal iRow As Long) As Integer
        Dim Kolor As Integer = System.Drawing.Color.Purple.ToArgb
        Dim NrTofi As String = ""

        Try
            NrTofi = IIf(IsDBNull(dViewKlienci.Item(iRow).Item("NRTOFITXT")), "", dViewKlienci.Item(iRow).Item("NRTOFITXT"))
            If Len(Trim(NrTofi)) = 0 Then
                'brak numeru TOFI
                Kolor = System.Drawing.Color.Red.ToArgb
            Else
                Dim words() As String
                Dim splitChars() As Char = {","c}
                words = NrTofi.Split(splitChars, StringSplitOptions.None)
                If words.Length = 1 Then
                    'jeden numer TOFI
                    Kolor = System.Drawing.Color.Purple.ToArgb
                Else
                    'wiêcej ni¿ jeden numer TOFI
                    Kolor = System.Drawing.Color.Green.ToArgb
                    'Kolor = System.Drawing.Color.Navy.ToArgb
                End If
            End If

        Catch ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace,
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
        End Try

        Return (Kolor)
    End Function



    Private Sub dagKlienci_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagKlienci.CellFormatting
        Dim NrTofi As String = ""
        If Not StartFormy Then
            '-----------------------
            'zmieniamy ForeColor
            NrTofi = IIf(IsDBNull(dViewKlienci.Item(e.RowIndex).Item("NRTOFITXT")), "", dViewKlienci.Item(e.RowIndex).Item("NRTOFITXT"))
            If Len(Trim(NrTofi)) = 0 Then
                'brak numeru TOFI
                e.CellStyle.ForeColor = Color.Red
            Else
                Dim words() As String
                Dim splitChars() As Char = {","c}
                words = NrTofi.Split(splitChars, StringSplitOptions.None)
                If words.Length = 1 Then
                    'jeden numer TOFI
                    e.CellStyle.ForeColor = Color.Purple
                Else
                    'wiêcej ni¿ jeden numer TOFI
                    e.CellStyle.ForeColor = Color.Green
                End If
            End If
            '-----------------------
            ''zmieniamy BackColor
            'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
            'Select Case Wal
            '    Case "EUR"
            '        e.CellStyle.BackColor = Color.White
            '    Case Else
            'End Select
            '-----------------------
        End If
    End Sub




    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagKlienci_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagKlienci.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKlienci_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKlienci_Scroll(sender As Object, e As ScrollEventArgs) Handles dagKlienci.Scroll
        'If (Not StartFormy) And (dViewKlienci.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (dViewKlienci.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If
    End Sub

    Private Sub dagKlienci_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub


    Private Sub dagKlienci_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKlienci.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormy Then
                '    PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagKlienci_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKlienci.MouseUp
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Dim message As String = "You clicked "

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None
                message &= "the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                message &= "cell at row " & hti.RowIndex & ", col " & hti.ColumnIndex
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagKlienci(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbKlienci.Panels(3).Text = "Sort: " & dSetKlienci.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagKlienci(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub




    Private Sub dagKlienci_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.DoubleClick
        If _CoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagKlienci_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagKlienci.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _CoMamRobic = "WYBIERAJ" Then
                    cmdStop_Click(sender, e)
                Else
                    cmdEdytuj_Click(sender, e)
                End If
            Case Keys.Insert
                cmdDodaj_Click(sender, e)
            Case Keys.Delete
                cmdUsuñ_Click(sender, e)
            Case Else
        End Select
    End Sub







    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************

    Private Sub InitKlienci()
        oUniqIdKon = Guid.NewGuid.ToString
        oIdentKon = ""
        oDataKon = SysData()
        oPracownikKon = ""
        oTematKon = ""
        oUwagiKon = ""
        oWersjaKon = 1          'WERSJA         Integer, 2
        '-------------------------
        oIdentKli = ""
        oNIPKli = ""
        oNrTOFIKli = 0
        oNrTOFITxtKli = ""
        oKartaPaybackKli = False
        oNryKartPaybackKli = ""
        oNazwa1Kli = ""
        oMiejscKli = ""
        oKodPoczKli = ""
        oUlicaKli = ""
        oNumNrDomuKli = 0
        oNrDomuKli = ""
        oIDDomuKli = ""
        oRejonKli = ""
        oTelefonKli = ""
        oFaxKli = ""
        oeMailKli = ""

        oWykazUrzadzenKli = ""
        oSposobWyboruDostawcyKli = ""
        oZainstalowanoMonitorKli = False
        oLiczbaUrzadzenKli = 0
        oLiczbaWydrukowKli = 0
        oPotencjalDrukuKli = ""
        oAktZakupMatEkspKli = False
        oAktRozliczaStronyDrukuKli = False
        oAktKorzystaZNajmuKli = False
        oZaintRozliczaniemStronDrukuKli = False
        oZaintKorzystaniemZNajmuKli = False
        oDataWeryfSegmentacjiKli = ""
        oPlanDlugookresowyKli = 0
        oPlanKrotkookresowyKli = 0

        oRodzSprzetuLKli = False
        oRodzSprzetuAKli = False
        oRodzSprzetuTKli = False
        oOfertaTZlozeniaKli = ""
        oOfertaPlikKli = ""
        oPKontaktKli = ""
        oSkupUwagikli = ""
        oSprzedOpiekunkli = ""
        oSprzedOKontaktRKli = Mid(SysData, 1, 4)
        oSprzedOKontaktTKli = ""
        oSprzedOKontaktDKli = ""
        oSprzedNKontaktRKli = Mid(SysData, 1, 4)
        oSprzedNKontaktTKli = ""
        oSprzedNKontaktDKli = ""
        oSprzedUwagiKli = ""
        oWpisalKli = ""
        oUwagikli = ""
        oMarkerKli = False
        oMarkerPKli = False
        oAktywnyKli = True
        oWersjaKli = 1          'WERSJA         Integer, 2
    End Sub

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(1).ColumnName, "Data kontaktu", 70, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(2).ColumnName, "Pracownik", 50, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(3).ColumnName, "Rodzaj kontaktu", 100, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(4).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
    ''-------------------------------------
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(5).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(6).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(7).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(8).ColumnName, "PayBack", 50, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(9).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(10).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(11).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(12).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(13).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(14).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(15).ColumnName, "Parzysty", 40, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(16).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(17).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(18).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(19).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(20).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(21).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(22).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(23).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(24).ColumnName, "Zainst. monitor", 50, True)
    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(25).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(26).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(27).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(28).ColumnName, "Akt. zakup mat.eksp.", 50, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(29).ColumnName, "Akt. rozlicza strony druku", 50, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(30).ColumnName, "Akt. korzysta z najmu", 50, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(31).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(32).ColumnName, "Zaint. korzyst. z najmu", 50, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(33).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)
    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(34).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
    'NumColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(35).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(36).ColumnName, "Rodz.Sprzêtu L", 100, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(37).ColumnName, "Rodz.Sprzêtu A", 100, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(38).ColumnName, "Rodz.Sprzêtu A3", 100, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(39).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(40).ColumnName, "Sprzêt - iloœæ", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(41).ColumnName, "Z³o¿ono", 40, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(42).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(43).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(44).ColumnName, "Promotor Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(45).ColumnName, "Promotor Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(46).ColumnName, "Promotor Ost.kontakt dzien", 150, DataGridViewContentAlignment.MiddleCenter, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(47).ColumnName, "Promotor Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(48).ColumnName, "Promotor Kol.kontakt tydz", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(49).ColumnName, "Promotor Kol.kontakt dzien", 150, DataGridViewContentAlignment.MiddleCenter, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(50).ColumnName, "Promotor Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(51).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(52).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(53).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(54).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(55).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(56).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(57).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(58).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(59).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(60).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(61).ColumnName, "Marker  ", 50, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(62).ColumnName, "Marker P.", 50, True)
    'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(63).ColumnName, "", 0, False)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(64).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(65).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)


    Private Sub PobierzKlienci()
        'pobierz wartosci przed aktualizacja
        oIdentKon = GetTxtField(dagKlienci, 0)
        oDataKon = GetTxtField(dagKlienci, 1)
        oPracownikKon = GetTxtField(dagKlienci, 2)
        oTematKon = GetTxtField(dagKlienci, 3)
        oUwagiKon = GetTxtField(dagKlienci, 4)
        '-------------------------

        oIdentKli = GetTxtField(dagKlienci, 5)
        oNIPKli = GetTxtField(dagKlienci, 6)
        oNrTOFITxtKli = GetTxtField(dagKlienci, 7)
        oKartaPaybackKli = GetLogField(dagKlienci, 8)
        oNryKartPaybackKli = GetTxtField(dagKlienci, 9)
        oNazwa1Kli = GetTxtField(dagKlienci, 10)
        oMiejscKli = GetTxtField(dagKlienci, 11)
        oKodPoczKli = GetTxtField(dagKlienci, 12)
        oUlicaKli = GetTxtField(dagKlienci, 13)
        oNumNrDomuKli = GetIntField(dagKlienci, 14)
        '15 parzyste/nieparzyste
        oNrDomuKli = GetTxtField(dagKlienci, 16)
        oIDDomuKli = GetTxtField(dagKlienci, 17)
        oRejonKli = GetTxtField(dagKlienci, 18)
        oTelefonKli = GetTxtField(dagKlienci, 19)
        oFaxKli = GetTxtField(dagKlienci, 20)
        oeMailKli = GetTxtField(dagKlienci, 21)

        oWykazUrzadzenKli = GetTxtField(dagKlienci, 22)
        oSposobWyboruDostawcyKli = GetTxtField(dagKlienci, 23)
        oZainstalowanoMonitorKli = GetLogField(dagKlienci, 24)
        oLiczbaUrzadzenKli = GetNumField(dagKlienci, 25)
        oLiczbaWydrukowKli = GetNumField(dagKlienci, 26)
        oPotencjalDrukuKli = GetTxtField(dagKlienci, 27)
        oAktZakupMatEkspKli = GetLogField(dagKlienci, 28)
        oAktRozliczaStronyDrukuKli = GetLogField(dagKlienci, 29)
        oAktKorzystaZNajmuKli = GetLogField(dagKlienci, 30)
        oZaintRozliczaniemStronDrukuKli = GetLogField(dagKlienci, 31)
        oZaintKorzystaniemZNajmuKli = GetLogField(dagKlienci, 32)
        oDataWeryfSegmentacjiKli = GetTxtField(dagKlienci, 33)
        oPlanDlugookresowyKli = GetNumField(dagKlienci, 34)
        oPlanKrotkookresowyKli = GetNumField(dagKlienci, 35)

        oRodzSprzetuLKli = GetLogField(dagKlienci, 36)
        oRodzSprzetuAKli = GetLogField(dagKlienci, 37)
        oRodzSprzetuTKli = GetLogField(dagKlienci, 38)
        oOfertaTZlozeniaKli = GetTxtField(dagKlienci, 39)
        oOfertaPlikKli = GetTxtField(dagKlienci, 40)
        oPKontaktKli = GetTxtField(dagKlienci, 41)
        oSkupUwagikli = GetTxtField(dagKlienci, 42)
        oSprzedOpiekunkli = GetTxtField(dagKlienci, 43)
        oSprzedOKontaktRKli = GetTxtField(dagKlienci, 44)
        oSprzedOKontaktTKli = GetTxtField(dagKlienci, 45)
        oSprzedOKontaktDKli = GetTxtField(dagKlienci, 46)
        oSprzedNKontaktRKli = GetTxtField(dagKlienci, 47)
        oSprzedNKontaktTKli = GetTxtField(dagKlienci, 48)
        oSprzedNKontaktDKli = GetTxtField(dagKlienci, 49)
        oSprzedUwagiKli = GetTxtField(dagKlienci, 50)
        oWpisalKli = GetTxtField(dagKlienci, 51)
        oUwagikli = GetTxtField(dagKlienci, 52)
        oMarkerKli = GetLogField(dagKlienci, 53)
        oMarkerPKli = GetLogField(dagKlienci, 54)
        oAktywnyKli = GetLogField(dagKlienci, 55)
        oUniqIdKon = GetTxtField(dagKlienci, 56)

        oWersjaKli = GetTxtField(dagKlienci, 57)
        oWersjaKon = GetTxtField(dagKlienci, 57)
    End Sub







    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdWybierzSchemat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierzSchemat.Click
        oCoMamRobic = "WYBIERAJ"
        oLokalnyFiltr = ""
        oNazwaSchematu = ""
        oSchematFiltrowania = ""
        Dim Form As New SchematFiltrowania("KatalogKlientowIKontaktow")
        Form.ShowDialog()

        'If Len(Trim(oSchematFiltrowania)) > 0 Then
        stbKlienci.Panels(4).Text = "Szablon : " & oNazwaSchematu
        Try
            'dViewKlienci.RowFilter = ""
            ''dViewKlienci.RowFilter = Trim(cbxFilter.SelectedItem.ToString) & " LIKE '" & Trim(txtFilter.Text) & "*'"
            'dViewKlienci.RowFilter = BudujFiltr(Trim(oSchematFiltrowania), oLokalnyFiltr)
            'dagKlienci.DataSource = dViewKlienci
            'stbKlienci.Panels(0).Text = "Il.rec.: " & dViewKlienci.Count.ToString

            FiltrujDataViewDGV(Me,
                            dagKlienci,
                            dViewKlienci.Table.Columns.Count,
                            dViewKlienci,
                            stbKlienci,
                            Trim(oSchematFiltrowania))
            '--------------------------------------------------------------------

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        End Try
        'Else
        '   stbKlienci.Panels(4).Text = "Szablon : "
        'End If

        If dViewKlienci.Count > 0 Then
            'ustaw sie na poczatku DataGrid
            dagKlienci.CurrentCell = dagKlienci(1, 0)
            dagKlienci.CurrentCell.Selected = True
        End If
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        ZapamietajSzerokosciKolumn()
        '-------------------------------
        If dViewKlienci.Count > 0 Then
            oKlient = GetTxtField(dagKlienci, 0)
            'oKlient = IIf(IsDBNull(dagKlienci.Item(dagKlienci.CurrentCell.RowNumber, 0)), "", dagKlienci.Item(dagKlienci.CurrentCell.RowNumber, 0))        'IDENT       Text    10
        Else
            oKlient = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub




    Private Sub cmdOdswiez_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOdswiez.Click
        If dSetKlienci.Tables(0).Rows.Count = 0 Then
            MessageBox.Show("Brak informacji o Kontaktach Handlowych.",
                    "Uwaga:",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
        Else

            '...........................
            'definiuj delegat
            Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
            Dim FormRaport As New RaportWybierzAnalizeGrupowaKontaktow(_dsKlenci,
                                                                 dViewKlienci,
                                                                 deleg,
                                                                 True)
            FormRaport.ShowDialog()
            '...........................
            MessageBox.Show("Wybra³em klientów na podstawie wyników analizy.",
                "OK - skoñczy³em",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)


        End If
    End Sub







    '*******************************
    ' edycja katalogu klientów
    '**********************************

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If dSetKlienci.Tables(0).Rows.Count <= 0 Then
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzKlienci()
            Dim Form As New EdytujKatalogKontaktow
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                'Dim findTheseVals(1) As Object
                'findTheseVals(0) = oIdentKon
                'findTheseVals(1) = oDataKon
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oUniqIdKon
                foundRow = dSetKlienci.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("UNIQID") = oUniqIdKon
                    foundRow("IDENTKLIENTA") = oIdentKli
                    foundRow("DATAKONTAKTU") = oDataKon
                    foundRow("PRACOWNIK") = oPracownikKon
                    'foundRow("RODZAJ") = oTematKon
                    foundRow("TEMAT") = oTematKon
                    foundRow("UWAGI") = oUwagiKon
                    foundRow("WERSJA") = ZmienWersje(oWersjaKon) 'Wersja         Integer, 2
                    '------------------------------
                    'foundRow("NRKARTY") = oIdentKli
                    foundRow("NIP") = oNIPKli
                    foundRow("NRTOFITXT") = oNrTOFITxtKli
                    foundRow("KARTAPAYBACK") = oKartaPaybackKli
                    foundRow("NRYKARTPAYBACK") = oNryKartPaybackKli
                    foundRow("NAZWAKLIENTA") = oNazwa1Kli
                    foundRow("MIEJSCOWOSC") = oMiejscKli
                    foundRow("KODPOCZTOWY") = oKodPoczKli
                    foundRow("ULICA") = oUlicaKli
                    foundRow("NUMNRDOMU") = oNumNrDomuKli
                    'PARZYSTE
                    foundRow("NRDOMU") = oNrDomuKli
                    foundRow("IDDOMU") = oIDDomuKli
                    foundRow("REJONKLIENTA") = oRejonKli
                    foundRow("TELEFON") = oTelefonKli
                    foundRow("FAX") = oFaxKli
                    foundRow("EMAIL") = oeMailKli

                    foundRow("WYKAZURZADZEN") = oWykazUrzadzenKli
                    foundRow("SPOSOBWYBORUDOSTAWCY") = oSposobWyboruDostawcyKli
                    foundRow("ZAINSTALOWANOMONITOR") = oZainstalowanoMonitorKli
                    foundRow("LICZBAURZADZEN") = oLiczbaUrzadzenKli
                    foundRow("LICZBAWYDRUKOW") = oLiczbaWydrukowKli
                    foundRow("POTENCJALDRUKU") = oPotencjalDrukuKli
                    foundRow("AKTZAKUPMATEKSP") = oAktZakupMatEkspKli
                    foundRow("AKTROZLICZASTRONYDRUKU") = oAktRozliczaStronyDrukuKli
                    foundRow("AKTKORZYSTAZNAJMU") = oAktKorzystaZNajmuKli
                    foundRow("ZAINTROZLICZANIEMSTRONDRUKU") = oZaintRozliczaniemStronDrukuKli
                    foundRow("ZAINTKORZYSTANIEMZNAJMU") = oZaintKorzystaniemZNajmuKli
                    foundRow("DATAWERYFSEGMENTACJI") = oDataWeryfSegmentacjiKli
                    foundRow("PLANDLUGOOKRESOWY") = oPlanDlugookresowyKli
                    foundRow("PLANKROTKOOKRESOWY") = oPlanKrotkookresowyKli

                    foundRow("RODZSPRZETUL") = oRodzSprzetuLKli
                    foundRow("RODZSPRZETUA") = oRodzSprzetuAKli
                    foundRow("RODZSPRZETUT") = oRodzSprzetuTKli
                    foundRow("OFERTATZLOZENIA") = oOfertaTZlozeniaKli
                    foundRow("OFERTAPLIK") = oOfertaPlikKli
                    foundRow("PKONTAKT") = oPKontaktKli
                    foundRow("PROMOTORUWAGI") = oSkupUwagikli
                    foundRow("OPIEKUN") = oSprzedOpiekunkli
                    foundRow("OPIEKUNOSTKONTAKTROK") = oSprzedOKontaktRKli
                    foundRow("OPIEKUNOSTKONTAKTTYDZ") = Val(oSprzedOKontaktTKli)
                    foundRow("OPIEKUNOSTKONTAKTDZIEN") = oSprzedOKontaktDKli
                    foundRow("OPIEKUNKOLKONTAKTROK") = oSprzedNKontaktRKli
                    foundRow("OPIEKUNKOLKONTAKTTYDZ") = Val(oSprzedNKontaktTKli)
                    foundRow("OPIEKUNKOLKONTAKTDZIEN") = oSprzedNKontaktDKli
                    foundRow("OPIEKUNUWAGI") = oSprzedUwagiKli
                    foundRow("KTOWPISAL") = oWpisalKli
                    foundRow("UWAGIKLI") = oUwagikli
                    foundRow("MARKER") = oMarkerKli
                    foundRow("MARKERP") = oMarkerPKli
                    foundRow("AKTYWNY") = oAktywnyKli
                    'foundRow("WERSJA") = ZmienWersje(oWersjaKli) 'Wersja         Integer, 2



                    'wyswietl ilosc rekordow
                    stbKlienci.Panels(0).Text = "Il.rec.: " & dViewKlienci.Count.ToString
                    AktualizujBaze()
                    DopiszDoLoguAktywnosci("",
                                           "Katalog Kontaktów Handlowych",
                                           SysTime(),
                                           Program_IdUzytkownika,
                                           LOG_OperacjaEdytuj,
                                           "Edytowano zapis kontaktu pracownika " & Trim(oPracownikKon) & " z klientem " & oIdentKon & vbCrLf &
                                           "Temat : " & oTematKon & vbCrLf &
                                           "Data kontaktu : " & oDataKon & "     Funkcja : KatalogKlientówiKontaktów" & vbCrLf &
                                           "")



                    'aktualizuj DataGrid
                    dagKlienci.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If dSetKlienci.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis ?",
                        "Proszê o potwierdzenie :",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
                oOperacja = "USUN"
                PobierzKlienci()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                'Dim findTheseVals(1) As Object
                'findTheseVals(0) = oIdentKon
                'findTheseVals(1) = oDataKon

                Dim findTheseVals(0) As Object
                findTheseVals(0) = oUniqIdKon
                foundRow = dSetKlienci.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbKlienci.Panels(0).Text = "Il.rec.: " & dViewKlienci.Count.ToString
                    AktualizujBaze()
                    DopiszDoLoguAktywnosci("",
                                           "Katalog Kontaktów Handlowych",
                                           SysTime(),
                                           Program_IdUzytkownika,
                                           LOG_OperacjaUsun,
                                           "Usuniêto zapis kontaktu pracownika " & Trim(oPracownikKon) & " z klientem " & oIdentKon & vbCrLf &
                                           "Temat : " & oTematKon & vbCrLf &
                                           "Data kontaktu : " & oDataKon & "     Funkcja : KatalogKlientówiKontaktów" & vbCrLf &
                                           oUwagiKon)
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitKlienci()
        '----------------------------------------------
        Dim Form As New EdytujKatalogKontaktow
        Do While True
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                'Dim findTheseVals(1) As Object
                'findTheseVals(0) = oIdentKon
                'findTheseVals(1) = oDataKon

                Dim findTheseVals(0) As Object
                findTheseVals(0) = oUniqIdKon
                foundRow = dSetKlienci.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "UniqID = " & oUniqIdKon,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = dSetKlienci.Tables(0).NewRow()

                    NewRow("UNIQID") = oUniqIdKon

                    NewRow("IDENTKLIENTA") = oIdentKon
                    NewRow("DATAKONTAKTU") = oDataKon
                    NewRow("PRACOWNIK") = oPracownikKon
                    'NewRow("RODZAJ") = oTematKon
                    NewRow("TEMAT") = oTematKon
                    NewRow("UWAGI") = oUwagiKon
                    NewRow("WERSJA") = 1 'Wersja         Integer, 2
                    '---------------------------
                    NewRow("NRKARTY") = oIdentKli
                    NewRow("NIP") = oNIPKli
                    NewRow("NRTOFITXT") = oNrTOFITxtKli
                    NewRow("KARTAPAYBACK") = oKartaPaybackKli
                    NewRow("NRYKARTPAYBACK") = oNryKartPaybackKli
                    NewRow("NAZWAKLIENTA") = oNazwa1Kli
                    NewRow("MIEJSCOWOSC") = oMiejscKli
                    NewRow("KODPOCZTOWY") = oKodPoczKli
                    NewRow("ULICA") = oUlicaKli
                    NewRow("NUMNRDOMU") = oNumNrDomuKli
                    'PARZYSTE
                    NewRow("NRDOMU") = oNrDomuKli
                    NewRow("IDDOMU") = oIDDomuKli
                    NewRow("REJONKLIENTA") = oRejonKli
                    NewRow("TELEFON") = oTelefonKli
                    NewRow("FAX") = oFaxKli
                    NewRow("EMAIL") = oeMailKli

                    NewRow("WYKAZURZADZEN") = oWykazUrzadzenKli
                    NewRow("SPOSOBWYBORUDOSTAWCY") = oSposobWyboruDostawcyKli
                    NewRow("ZAINSTALOWANOMONITOR") = oZainstalowanoMonitorKli
                    NewRow("LICZBAURZADZEN") = oLiczbaUrzadzenKli
                    NewRow("LICZBAWYDRUKOW") = oLiczbaWydrukowKli
                    NewRow("POTENCJALDRUKU") = oPotencjalDrukuKli
                    NewRow("AKTZAKUPMATEKSP") = oAktZakupMatEkspKli
                    NewRow("AKTROZLICZASTRONYDRUKU") = oAktRozliczaStronyDrukuKli
                    NewRow("AKTKORZYSTAZNAJMU") = oAktKorzystaZNajmuKli
                    NewRow("ZAINTROZLICZANIEMSTRONDRUKU") = oZaintRozliczaniemStronDrukuKli
                    NewRow("ZAINTKORZYSTANIEMZNAJMU") = oZaintKorzystaniemZNajmuKli
                    NewRow("DATAWERYFSEGMENTACJI") = oDataWeryfSegmentacjiKli
                    NewRow("PLANDLUGOOKRESOWY") = oPlanDlugookresowyKli
                    NewRow("PLANKROTKOOKRESOWY") = oPlanKrotkookresowyKli

                    NewRow("RODZSPRZETUL") = oRodzSprzetuLKli
                    NewRow("RODZSPRZETUA") = oRodzSprzetuAKli
                    NewRow("RODZSPRZETUT") = oRodzSprzetuTKli
                    NewRow("OFERTATZLOZENIA") = oOfertaTZlozeniaKli
                    NewRow("OFERTAPLIK") = oOfertaPlikKli
                    NewRow("PKONTAKT") = oPKontaktKli
                    NewRow("PROMOTORUWAGI") = oSkupUwagikli
                    NewRow("OPIEKUN") = oSprzedOpiekunkli
                    NewRow("OPIEKUNOSTKONTAKTROK") = oSprzedOKontaktRKli
                    NewRow("OPIEKUNOSTKONTAKTTYDZ") = Val(oSprzedOKontaktTKli)
                    NewRow("OPIEKUNOSTKONTAKTDZIEN") = oSprzedOKontaktDKli
                    NewRow("OPIEKUNKOLKONTAKTROK") = oSprzedNKontaktRKli
                    NewRow("OPIEKUNKOLKONTAKTTYDZ") = Val(oSprzedNKontaktTKli)
                    NewRow("OPIEKUNKOLKONTAKTDZIEN") = oSprzedNKontaktDKli
                    NewRow("OPIEKUNUWAGI") = oSprzedUwagiKli
                    NewRow("KTOWPISAL") = oWpisalKli
                    NewRow("UWAGIKLI") = oUwagikli
                    NewRow("MARKER") = False
                    NewRow("MARKERP") = False
                    NewRow("AKTYWNY") = True
                    'NewRow("WERSJA") = 1                     'Wersja         Integer, 2


                    'dodaj rekord do DataSet
                    dSetKlienci.Tables(0).Rows.Add(NewRow)
                    'wyswietl ilosc rekordow
                    stbKlienci.Panels(0).Text = "Il.rec.: " & dViewKlienci.Count.ToString
                    AktualizujBaze()
                    DopiszDoLoguAktywnosci("",
                                           "Katalog Kontaktów Handlowych",
                                           SysTime(),
                                           Program_IdUzytkownika,
                                           LOG_OperacjaDodaj,
                                           "Dodano zapis kontaktu pracownika " & Trim(oPracownikKon) & " z klientem " & oIdentKon & vbCrLf &
                                           "Temat : " & oTematKon & vbCrLf &
                                           "Data kontaktu : " & oDataKon & "     Funkcja : KatalogKlientówiKontaktów" & vbCrLf &
                                           "")

                    'aktualizuj DataGrid
                    dagKlienci.Update()
                    Exit Do
                End If
            End If
        Loop
    End Sub

    'Private Sub cmdWysylaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWysylaj.Click
    '    Dim Form As New eMailing(dViewKlienci, "", "")
    '    Form.ShowDialog()
    'End Sub




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
            '        dbDataAdapter.Update(dSetKlienci, "TABELA_Klienci")
            '    Catch Ex As System.Exception
            '        'ViewInLog(ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(dSetKlienci)
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
                    sqlDataAdapter.Update(dSetKlienci, "TABELA_Klienci")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(dSetKlienci)
                End If
                sqlConnection.Close()
            End If
        End If

        stbKlienci.Panels(0).Text = "Il.rec.: " & dViewKlienci.Count.ToString
    End Sub

    Private Sub OdswiezBaze()
        dSetKlienci.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(dSetKlienci)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(dSetKlienci)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub





































    Private Sub dtpOdDaty_LostFocus(sender As Object, e As EventArgs) Handles dtpOdDaty.LostFocus
        '---------------------------------------
        If Not StartFormy Then
            TakenDate = dtpOdDaty.Value
            OdDaty = TakenDate.ToString("yyyy-MM-dd")
            TakenDate = dtpDoDaty.Value
            DoDaty = TakenDate.ToString("yyyy-MM-dd")
            PokazKontakty(OdDaty, DoDaty)
        End If
        '---------------------------------------
    End Sub

    Private Sub dtpDoDaty_LostFocus(sender As Object, e As EventArgs) Handles dtpDoDaty.LostFocus
        '---------------------------------------
        If Not StartFormy Then
            TakenDate = dtpOdDaty.Value
            OdDaty = TakenDate.ToString("yyyy-MM-dd")
            TakenDate = dtpDoDaty.Value
            DoDaty = TakenDate.ToString("yyyy-MM-dd")
            PokazKontakty(OdDaty, DoDaty)
        End If
        '---------------------------------------
    End Sub
















    '------------------------
    ' wydruk Raportu
    '--------------------------

    Private Sub cmdDrukuj_Click(sender As Object, e As EventArgs) Handles cmdDrukuj.Click
        ' Tworzymy dokument i do³¹czmy obs³ugê zdarzenia PrintPage
        Dim MyDoc As New Softart.myPrintDocument
        MyDoc.DocumentName = "Kontakty Handlowe"
        MyDoc.LineNumber = 0
        MyDoc.PageNumber = 0
        MyDoc.NextRowNumber = 0
        MyDoc.DefaultPageSettings.Landscape = False
        AddHandler MyDoc.PrintPage, AddressOf MyDoc_Zestawienie

        Try
            'wybor drukarki...
            Dim PDialog As New PrintDialog
            PDialog.Document = MyDoc
            PDialog.AllowPrintToFile = True
            PDialog.PrintToFile = False
            PDialog.AllowSelection = True
            PDialog.ShowHelp = True
            PDialog.ShowNetwork = True
            PDialog.UseEXDialog = True

            Dim Result As DialogResult = PDialog.ShowDialog()
            ' Drukuj po nacisnieciu OK
            If Result = System.Windows.Forms.DialogResult.OK Then
                ' This method returns immediately, before the print job starts.
                ' The PrintPage event will fire asynchronously.
                'MyDoc.Print()

                'podglad wydruku
                Dim dlgPreview As New PrintPreviewDialog
                dlgPreview.Document = MyDoc
                dlgPreview.WindowState = FormWindowState.Maximized
                dlgPreview.UseAntiAlias = True  'lepsza jakosc wydruku
                dlgPreview.ShowDialog()
            End If
        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1)
        Finally
        End Try
    End Sub



    Private Sub MyDoc_Zestawienie(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        ' Retrieve the document that sent this event.
        ' You could store the document in a class member variable,
        ' but this approach allows you to use the same event handler
        ' to handle multiple print documents at once.
        Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
        Dim Text As String
        Dim LiRow As Integer

        e.PageSettings.Landscape = True

        ' Define the font and determine the line height.
        Dim FontSizeTitle As Integer = 12
        Dim MyFontTitle As New System.Drawing.Font("Arial", FontSizeTitle, FontStyle.Bold)
        Dim LineHeightTitle As Single = MyFontTitle.GetHeight(e.Graphics)

        Dim FontSize As Integer = 8
        Dim MyFont As New System.Drawing.Font("Arial", FontSize, FontStyle.Regular)
        Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)

        Dim DrawingPen As New Pen(Color.Black, 1)

        ' Create variables to hold position on page.
        Dim x As Single = e.MarginBounds.Left
        Dim y As Single = e.MarginBounds.Top

        Dim PageWidth As Single = cm2pts(19)        'e.MarginBounds.Width
        Dim PageHeight As Single = cm2pts(28)       'e.MarginBounds.Height
        Dim PageLeft As Single = cm2pts(0.5)          'e.MarginBounds.Left
        Dim PageTop As Single = cm2pts(1)           'e.MarginBounds.Top

        Dim words() As String
        Dim splitChars() As Char = {","c}
        Dim Tofi1 As String = ""
        Dim Tofi2 As String = ""
        Dim Tofi3 As String = ""
        Dim Tofi4 As String = ""
        Dim Tofi5 As String = ""
        Dim Tofi6 As String = ""

        Dim IleLiOpis As Integer = 0
        Dim II As Integer = 0



        ' Naglowek strony
        Doc.PageNumber += 1
        x = PageLeft
        y = PageTop

        RightTxt("Str. " & Trim(Str(Doc.PageNumber)), x + cm2pts(0), y, PageWidth, MyFont, e)
        y += MyFont.GetHeight(e.Graphics)

        Text = "Zestawienie Kontaktów Handlowych"
        x = PageLeft + (PageWidth - e.Graphics.MeasureString(Text, MyFontTitle).Width) / 2
        e.Graphics.DrawString(Text, MyFontTitle, Brushes.Black, x, y)
        y += 2 * MyFontTitle.GetHeight(e.Graphics)

        x = PageLeft
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(2), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(8), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(12), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(19), 4 * LineHeight)
        y += 0.5 * MyFont.GetHeight(e.Graphics)

        CentrTxt("Nr Karty", x + cm2pts(0), y, cm2pts(2), MyFont, e)
        CentrTxt("Nazwa Klienta", x + cm2pts(2.0), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Przedstawiciel", x + cm2pts(8.0), y, cm2pts(4.0), MyFont, e)
        CentrTxt("Kontakt Handlowy", x + cm2pts(12.0), y, cm2pts(7.0), MyFont, e)
        y += 1.0 * MyFont.GetHeight(e.Graphics)

        CentrTxt("Nr TOFI", x + cm2pts(0), y, cm2pts(2), MyFont, e)
        CentrTxt("Kod.P. i Miejscowoœæ", x + cm2pts(2.0), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Data kontaktu", x + cm2pts(8.0), y, cm2pts(4.0), MyFont, e)
        CentrTxt("", x + cm2pts(12.0), y, cm2pts(7.0), MyFont, e)
        y += 1.0 * MyFont.GetHeight(e.Graphics)

        CentrTxt("", x + cm2pts(0), y, cm2pts(2), MyFont, e)
        CentrTxt("Ulica i Nr Domu", x + cm2pts(2.0), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Rodzaj Kontaktu", x + cm2pts(8.0), y, cm2pts(4.0), MyFont, e)
        CentrTxt("", x + cm2pts(12.0), y, cm2pts(7.0), MyFont, e)
        y += 2 * MyFont.GetHeight(e.Graphics)


        ' petla do konca strony
        If Doc.NextRowNumber < dViewKlienci.Count Then
            Do
                'ustaw kursor na wierszu DataGrid
                LiRow = Doc.NextRowNumber
                dagKlienci.CurrentCell = dagKlienci(LiRow, 1)

                PobierzKlienci()

                'oIdentKon = GetTxtField(dagKlienci, 0)
                'oDataKon = GetTxtField(dagKlienci, 1)
                'oPracownikKon = GetTxtField(dagKlienci, 2)
                'oTematKon = GetTxtField(dagKlienci, 3)
                'oUwagiKon = GetTxtField(dagKlienci, 4)
                ''-------------------------
                'oIdentKli = GetTxtField(dagKlienci, 5)
                'oNIPKli = GetTxtField(dagKlienci, 6)
                'oNrTOFITxtKli = GetTxtField(dagKlienci, 7)
                'oKartaPaybackKli = GetLogField(dagKlienci, 8)
                'oNryKartPaybackKli = GetTxtField(dagKlienci, 9)
                'oNazwa1Kli = GetTxtField(dagKlienci, 13)
                'oMiejscKli = GetTxtField(dagKlienci, 14)
                'oKodPoczKli = GetTxtField(dagKlienci, 15)
                'oUlicaKli = GetTxtField(dagKlienci, 16)
                'oNumNrDomuKli = GetIntegerField(dagKlienci, 17)
                ''parzyste/nieparzyste
                'oNrDomuKli = GetTxtField(dagKlienci, 19)
                'oIDDomuKli = GetTxtField(dagKlienci, 20)
                'oRejonKli = GetTxtField(dagKlienci, 21)
                'oTelefonKli = GetTxtField(dagKlienci, 22)
                'oFaxKli = GetTxtField(dagKlienci, 23)
                'oeMailKli = GetTxtField(dagKlienci, 24)
                'oRodzSprzetuLKli = GetLogField(dagKlienci, 25)
                'oRodzSprzetuAKli = GetLogField(dagKlienci, 26)
                'oRodzSprzetuTKli = GetLogField(dagKlienci, 27)
                'oOfertaTZlozeniaKli = GetTxtField(dagKlienci, 30)
                'oOfertaPlikKli = GetTxtField(dagKlienci, 31)
                'oPKontaktKli = GetTxtField(dagKlienci, 33)
                'oSkupUwagikli = GetTxtField(dagKlienci, 43)
                'oSprzedOpiekunkli = GetTxtField(dagKlienci, 45)
                'oSprzedOKontaktRKli = GetTxtField(dagKlienci, 47)
                'oSprzedOKontaktTKli = GetTxtField(dagKlienci, 48)
                'oSprzedOKontaktDKli = GetTxtField(dagKlienci, 49)
                'oSprzedNKontaktRKli = GetTxtField(dagKlienci, 50)
                'oSprzedNKontaktTKli = GetTxtField(dagKlienci, 51)
                'oSprzedNKontaktDKli = GetTxtField(dagKlienci, 52)
                'oSprzedUwagiKli = GetTxtField(dagKlienci, 53)
                'oWpisalKli = GetTxtField(dagKlienci, 56)
                'oUwagikli = GetTxtField(dagKlienci, 58)
                'oMarkerKli = GetLogField(dagKlienci, 59)
                'oMarkerPKli = GetLogField(dagKlienci, 60)
                ''uniqid
                'oWersjaKli = GetTxtField(dagKlienci, 62)
                'oWersjaKon = GetTxtField(dagKlienci, 62)

                Tofi1 = ""
                Tofi2 = ""
                Tofi3 = ""
                Tofi4 = ""
                Tofi5 = ""
                Tofi6 = ""
                words = oNrTOFITxtKli.Split(splitChars, StringSplitOptions.None)
                If words.Length = 0 Then
                ElseIf words.Length = 1 Then
                    Tofi1 = oNrTOFITxtKli
                ElseIf words.Length = 2 Then
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                ElseIf words.Length = 3 Then
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                    Tofi3 = words(2)
                ElseIf words.Length = 4 Then
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                    Tofi3 = words(2)
                    Tofi4 = words(3)
                ElseIf words.Length = 5 Then
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                    Tofi3 = words(2)
                    Tofi4 = words(3)
                    Tofi5 = words(4)
                Else
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                    Tofi3 = words(2)
                    Tofi4 = words(3)
                    Tofi5 = words(4)
                    Tofi6 = words(5)
                End If

                IleLiOpis = IleLiniiPotrzebaNaText(oUwagiKon, cm2pts(7.0), MyFont, e)

                x = PageLeft 'od marginesu...
                Doc.LineNumber += 1
                Doc.NextRowNumber += 1


                'wiersz 1
                LeftTxt(oIdentKli, x + cm2pts(0.0), y, cm2pts(2.0), MyFont, e)
                LeftTxt(oNazwa1Kli, x + cm2pts(2.0), y, cm2pts(6.0), MyFont, e)
                LeftTxt(oPracownikKon, x + cm2pts(8.0), y, cm2pts(4.0), MyFont, e)
                LeftTxt(DajLinieTextuNr(oUwagiKon, 1, cm2pts(7.0), MyFont, e), x + cm2pts(12.0), y, cm2pts(7.0), MyFont, e)
                y += LineHeight

                'wiersz 2
                LeftTxt(Tofi1, x + cm2pts(0.0), y, cm2pts(2.0), MyFont, e)
                LeftTxt(Trim(oKodPoczKli) & " " & Trim(oMiejscKli), x + cm2pts(2.0), y, cm2pts(6.0), MyFont, e)
                LeftTxt(oDataKon, x + cm2pts(8.0), y, cm2pts(4.0), MyFont, e)
                If IleLiOpis > 1 Then
                    LeftTxt(DajLinieTextuNr(oUwagiKon, 2, cm2pts(7.0), MyFont, e), x + cm2pts(12.0), y, cm2pts(7.0), MyFont, e)
                Else
                End If
                y += LineHeight

                'wiersz 3
                LeftTxt(Tofi2, x + cm2pts(0.0), y, cm2pts(2.0), MyFont, e)
                LeftTxt(Trim(oUlicaKli) & " " & Trim(oNrDomuKli), x + cm2pts(2.0), y, cm2pts(6.0), MyFont, e)
                LeftTxt(oTematKon, x + cm2pts(8.0), y, cm2pts(4.0), MyFont, e)
                If IleLiOpis > 2 Then
                    LeftTxt(DajLinieTextuNr(oUwagiKon, 3, cm2pts(7.0), MyFont, e), x + cm2pts(12.0), y, cm2pts(7.0), MyFont, e)
                Else
                End If
                y += LineHeight

                If IleLiOpis > 3 Then
                    For II = 4 To IleLiOpis
                        If ((y + 4 * LineHeight) < e.MarginBounds.Bottom) Then
                            LeftTxt(DajLinieTextuNr(oUwagiKon, 3, cm2pts(7.0), MyFont, e), x + cm2pts(12.0), y, cm2pts(7.0), MyFont, e)
                            y += LineHeight
                        Else
                            Exit For
                        End If
                    Next II
                End If
                '--------------------------------------
                y += CType(0.5 * LineHeight, Single)
                e.Graphics.DrawLine(DrawingPen, x, y, x + cm2pts(19), y)
                'e.Graphics.DrawLine(DrawingPen, x, y + CType(0.2 * LineHeight, Single), x + CType(19 * (100 / 2.54), Single), y + CType(0.5 * LineHeight, Single))
                y += CType(0.5 * LineHeight, Single)
            Loop Until ((y + 3 * LineHeight) > e.MarginBounds.Bottom) Or (Doc.NextRowNumber >= dViewKlienci.Count)
        End If

        If (Doc.NextRowNumber < dViewKlienci.Count) Then
            ' There is still at least one more page.
            ' Signal this event to fire again.
            e.HasMorePages = True
        Else
            ' Printing is complete.
            x = PageLeft 'od marginesu...
            'y += CType(0.5 * LineHeight, Single)
            'e.Graphics.DrawLine(DrawingPen, x, y + CType(0.2 * LineHeight, Single), x + CType(19 * (100 / 2.54), Single), y + CType(0.5 * LineHeight, Single))
            'y += CType(0.5 * LineHeight, Single)
            Text = "Koniec zestawienia na pozycji " & Str(Doc.LineNumber)
            e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + CType(0.2 * (100 / 2.54), Single), y)

            Doc.PageNumber = 0
            Doc.LineNumber = 0
            Doc.NextRowNumber = 0

            e.HasMorePages = False
        End If
    End Sub







    Private Sub cmdHarmonogramWizyt_Click(sender As Object, e As EventArgs) Handles cmdHarmonogramWizyt.Click
        Dim Form As New KatalogHarmonogramWizyt
        Form.Show()
    End Sub

    '*************************************************
    '** obs³uga menu kontekstowego
    '*************************************************

    Private Sub menuEEdytuj_Click(sender As Object, e As EventArgs) Handles menuEEdytuj.Click
        cmdEdytuj_Click(sender, e)
    End Sub

    Private Sub menuEDodaj_Click(sender As Object, e As EventArgs) Handles menuEDodaj.Click
        cmdDodaj_Click(sender, e)
    End Sub

    Private Sub menuEUsun_Click(sender As Object, e As EventArgs) Handles menuEUsun.Click
        cmdUsuñ_Click(sender, e)
    End Sub

    Private Sub menuESingleL_Click(sender As Object, e As EventArgs) Handles menuESingleL.Click
        dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagKlienci.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagKlienci.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub

    Private Sub menuEAkcja_Click(sender As Object, e As EventArgs) Handles menuEAkcja.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        'Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        'Dim Form As New AkcjaMarketingowaWybierz(dSetKlienci, deleg)
        Dim Form As New AkcjaMarketingowaWybierz(dSetKlienci, Nothing)
        Form.ShowDialog()

        If oWybranoAkcjeMarketingowa Then
            'aktualizuj Katalog klientow
            StartFormy = True
            CzyscFiltryDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, dViewKlienci, stbKlienci)
            StartFormy = False

            'CzyscFiltryDGV(Me,
            '            dagKlienci,
            '            dViewKlienci.Table.Columns.Count,
            '            dViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(61).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, dSetKlienci.Tables(0).Columns(62).ColumnName, "Marker P.", 50, True)
            'Dim nazwa As String = "txtFi57"
            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi061" Or ctrl.Name = "txtFi062" Then
                        ctrl.Text = "T"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            'ustaw sie na pierwszym zapisie
            'If dSetKlienci.Tables(0).Rows.Count > 0 Then
            If dViewKlienci.Count > 0 Then
                dagKlienci.CurrentCell = dagKlienci(1, 0)
                dagKlienci.CurrentCell.Selected = True
            End If
            '--------------------------------------------------------------------
        End If
    End Sub




















    '**************************************************************
    '** Zapamietaj szerokosci kolumn
    '**************************************************************

    Public Sub ZapamietajSzerokosciKolumn()
        Dim ii As Integer = 0
        Dim Txt As String = ""
        Dim FileNum As Integer

        For ii = 0 To dViewKlienci.Table.Columns.Count - 1
            Txt &= Trim(Str(dagKlienci.Columns.Item(ii).Width)) & "|"
        Next

        If IO.File.Exists(oKatParametry & "\" & oPlikSzerokosciKolumnKontakty) Then
            IO.File.Delete(oKatParametry & "\" & oPlikSzerokosciKolumnKontakty)
        End If
        FileNum = FreeFile()    'kolejny nr otwartego zbioru
        FileOpen(FileNum, oKatParametry & "\" & oPlikSzerokosciKolumnKontakty, OpenMode.Output)
        PrintLine(FileNum, "KartotekaKlientowPryzmat" & oAktualnaWersjaBazyDanych.ToString)
        PrintLine(FileNum, Txt)
        FileClose(FileNum)
    End Sub

    Public Sub UstalSzerokosciKolumn()
        Dim ii As Integer = 0
        Dim Txt As String = ""
        Dim FileNum As Integer
        Dim pos As Integer = 0
        Dim NrWersjiPliku As Integer = 0
        Dim NowyTXT As String = ""

        If IO.File.Exists(oKatParametry & "\" & oPlikSzerokosciKolumnKontakty) Then
            FileNum = FreeFile()    'kolejny nr otwartego zbioru
            FileOpen(FileNum, oKatParametry & "\" & oPlikSzerokosciKolumnKontakty, OpenMode.Input)
            If Not EOF(FileNum) Then
                Txt = LineInput(FileNum)

                '-------------obsluga zmiany pliku
                If InStr(Txt, "KartotekaKlientowPryzmat") = 1 Then
                    NrWersjiPliku = CInt(Mid(Txt, 25, 3))
                    Txt = LineInput(FileNum)
                Else
                    'ident wersji w pliku od wersji 2.40
                    NrWersjiPliku = 230
                End If

                'sprawdzenie koniecznosci korekty
                'sprawdzenie koniecznosci korekty
                If NrWersjiPliku <> oAktualnaWersjaBazyDanych Then
                    'pozostawiamy domyœlne...
                Else
                    'pobieramy z pliku
                    For ii = 0 To dViewKlienci.Table.Columns.Count - 1
                        pos = InStr(Txt, "|")
                        If pos > 0 Then
                            dagKlienci.Columns.Item(ii).Width = CInt(Mid(Txt, 1, pos - 1))
                            Txt = Mid(Txt, pos + 1)
                        End If
                    Next
                End If
                '---------------------------------
            End If
            FileClose(FileNum)
        End If
    End Sub


    Private Sub clrClearColWidth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearColWidth.Click
        Dim ii As Integer = 0
        StartFormy = True
        Me.Cursor = Cursors.WaitCursor
        For ii = 0 To dViewKlienci.Table.Columns.Count - 1
            'dagKlienci.Columns.Item(ii).Width = GetIntFromText(dagKlienci.Columns.Item(ii).Tag)
            If IsNumeric(dagKlienci.Columns.Item(ii).Tag) Then
                dagKlienci.Columns.Item(ii).Width = dagKlienci.Columns.Item(ii).Tag
            Else
                dagKlienci.Columns.Item(ii).Width = 0
            End If
        Next
        StartFormy = False
        dagKlienci_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagKlienci, dViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        Me.Cursor = Cursors.Default
        ''--------------------------------
        ZapamietajSzerokosciKolumn()
    End Sub


End Class
