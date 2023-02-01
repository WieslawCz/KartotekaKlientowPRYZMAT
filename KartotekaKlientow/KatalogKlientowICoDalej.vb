
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






Public Class KatalogKlientowICoDalej
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Dim _dsKlenci As DataSet
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents cmdHarmonogramWizyt As System.Windows.Forms.Button
    Friend WithEvents DagKlienci As DataGridView
    Friend WithEvents menuWybieraj As ContextMenuStrip
    Friend WithEvents MenuWEdytuj As ToolStripMenuItem
    Friend WithEvents MenuWDodaj As ToolStripMenuItem
    Friend WithEvents MenuWUsuñ As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents MenuWSingleL As ToolStripMenuItem
    Friend WithEvents menuWMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents MenuWWybierzA As ToolStripMenuItem
    Friend WithEvents menuWOdswiez As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
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
    Friend WithEvents lblNotatka As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogKlientowICoDalej))
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbKlienci = New System.Windows.Forms.StatusBar()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblTemat = New System.Windows.Forms.Label()
        Me.lblNotatka = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
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
        Me.cmdHarmonogramWizyt = New System.Windows.Forms.Button()
        Me.DagKlienci = New System.Windows.Forms.DataGridView()
        Me.menuWybieraj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MenuWEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuWDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuWUsuñ = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuWSingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuWMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuWWybierzA = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.DagKlienci, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuWybieraj.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(911, 431)
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
        Me.stbKlienci.Location = New System.Drawing.Point(8, 372)
        Me.stbKlienci.Name = "stbKlienci"
        Me.stbKlienci.ShowPanels = True
        Me.stbKlienci.Size = New System.Drawing.Size(983, 22)
        Me.stbKlienci.TabIndex = 43
        Me.stbKlienci.Text = "StatusBar1"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(10, 430)
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
        Me.Panel1.Controls.Add(Me.lblNotatka)
        Me.Panel1.Controls.Add(Me.Label17)
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
        Me.Panel1.Size = New System.Drawing.Size(983, 111)
        Me.Panel1.TabIndex = 46
        '
        'lblTemat
        '
        Me.lblTemat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTemat.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTemat.ForeColor = System.Drawing.Color.Purple
        Me.lblTemat.Location = New System.Drawing.Point(787, 7)
        Me.lblTemat.Name = "lblTemat"
        Me.lblTemat.Size = New System.Drawing.Size(411, 16)
        Me.lblTemat.TabIndex = 42
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
        Me.lblNotatka.Size = New System.Drawing.Size(237, 80)
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
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(742, 8)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(56, 16)
        Me.Label18.TabIndex = 41
        Me.Label18.Text = "Ident. . . . . . . . . "
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
        Me.Label8.Text = "Potencja³ . . . . . ."
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
        Me.Label10.Text = "Opiekun . . . . . . ."
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
        Me.stbFiltry.Location = New System.Drawing.Point(8, 125)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(983, 21)
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
        Me.cmdClearColWidth.Location = New System.Drawing.Point(8, 147)
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
        Me.cmdEdytuj.Location = New System.Drawing.Point(911, 402)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytuj.TabIndex = 182
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdEdytuj, "Edycja informacji o planach wskazywanych przez kursor.")
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Enabled = False
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.ForeColor = System.Drawing.Color.Black
        Me.cmdUsuñ.Image = CType(resources.GetObject("cmdUsuñ.Image"), System.Drawing.Image)
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(825, 402)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 23)
        Me.cmdUsuñ.TabIndex = 181
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdUsuñ, "Usuniecie informacji o planach wskazywanych przez kursor" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " Kartoteki.")
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Enabled = False
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.ForeColor = System.Drawing.Color.Black
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(739, 402)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodaj.TabIndex = 180
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDodaj, "Dopisanie (dodanie) nowego planu do kartoteki.")
        '
        'cmdOdswiez
        '
        Me.cmdOdswiez.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdOdswiez.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdOdswiez.ForeColor = System.Drawing.Color.Black
        Me.cmdOdswiez.Image = CType(resources.GetObject("cmdOdswiez.Image"), System.Drawing.Image)
        Me.cmdOdswiez.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOdswiez.Location = New System.Drawing.Point(150, 432)
        Me.cmdOdswiez.Name = "cmdOdswiez"
        Me.cmdOdswiez.Size = New System.Drawing.Size(88, 22)
        Me.cmdOdswiez.TabIndex = 173
        Me.cmdOdswiez.Text = "Pobierz"
        Me.cmdOdswiez.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdswiez, "Wybierz z Kartoteki Klientów klientów analizowanych planów Co Dalej.")
        '
        'HorizScrollTimer
        '
        '
        'cmdHarmonogramWizyt
        '
        Me.cmdHarmonogramWizyt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdHarmonogramWizyt.ForeColor = System.Drawing.Color.Black
        Me.cmdHarmonogramWizyt.Image = CType(resources.GetObject("cmdHarmonogramWizyt.Image"), System.Drawing.Image)
        Me.cmdHarmonogramWizyt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdHarmonogramWizyt.Location = New System.Drawing.Point(120, 431)
        Me.cmdHarmonogramWizyt.Name = "cmdHarmonogramWizyt"
        Me.cmdHarmonogramWizyt.Size = New System.Drawing.Size(24, 23)
        Me.cmdHarmonogramWizyt.TabIndex = 205
        Me.cmdHarmonogramWizyt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DagKlienci
        '
        Me.DagKlienci.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DagKlienci.Location = New System.Drawing.Point(8, 147)
        Me.DagKlienci.Name = "DagKlienci"
        Me.DagKlienci.Size = New System.Drawing.Size(983, 225)
        Me.DagKlienci.TabIndex = 206
        '
        'menuWybieraj
        '
        Me.menuWybieraj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuWEdytuj, Me.MenuWDodaj, Me.MenuWUsuñ, Me.ToolStripSeparator1, Me.MenuWSingleL, Me.menuWMultiL, Me.ToolStripSeparator2, Me.menuWOdswiez, Me.ToolStripSeparator3, Me.MenuWWybierzA})
        Me.menuWybieraj.Name = "menuWybieraj"
        Me.menuWybieraj.Size = New System.Drawing.Size(226, 176)
        '
        'MenuWEdytuj
        '
        Me.MenuWEdytuj.Name = "MenuWEdytuj"
        Me.MenuWEdytuj.Size = New System.Drawing.Size(225, 22)
        Me.MenuWEdytuj.Text = "Edytuj"
        '
        'MenuWDodaj
        '
        Me.MenuWDodaj.Enabled = False
        Me.MenuWDodaj.Name = "MenuWDodaj"
        Me.MenuWDodaj.Size = New System.Drawing.Size(225, 22)
        Me.MenuWDodaj.Text = "Dodaj"
        '
        'MenuWUsuñ
        '
        Me.MenuWUsuñ.Enabled = False
        Me.MenuWUsuñ.Name = "MenuWUsuñ"
        Me.MenuWUsuñ.Size = New System.Drawing.Size(225, 22)
        Me.MenuWUsuñ.Text = "Usuñ"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(222, 6)
        '
        'MenuWSingleL
        '
        Me.MenuWSingleL.Name = "MenuWSingleL"
        Me.MenuWSingleL.Size = New System.Drawing.Size(225, 22)
        Me.MenuWSingleL.Text = "Poka¿ w jednej linii"
        '
        'menuWMultiL
        '
        Me.menuWMultiL.Name = "menuWMultiL"
        Me.menuWMultiL.Size = New System.Drawing.Size(225, 22)
        Me.menuWMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(222, 6)
        '
        'menuWOdswiez
        '
        Me.menuWOdswiez.Name = "menuWOdswiez"
        Me.menuWOdswiez.Size = New System.Drawing.Size(225, 22)
        Me.menuWOdswiez.Text = "Odœwie¿ Tabele"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(222, 6)
        '
        'MenuWWybierzA
        '
        Me.MenuWWybierzA.Name = "MenuWWybierzA"
        Me.MenuWWybierzA.Size = New System.Drawing.Size(225, 22)
        Me.MenuWWybierzA.Text = "Wybierz akcjê marketingow¹"
        '
        'KatalogKlientowICoDalej
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(999, 460)
        Me.Controls.Add(Me.cmdClearColWidth)
        Me.Controls.Add(Me.DagKlienci)
        Me.Controls.Add(Me.cmdHarmonogramWizyt)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.cmdOdswiez)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdWybierzSchemat)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.stbKlienci)
        Me.Controls.Add(Me.PictureBox1)
        Me.ForeColor = System.Drawing.Color.Teal
        Me.Location = New System.Drawing.Point(8, 8)
        Me.Name = "KatalogKlientowICoDalej"
        Me.Text = "CoDalej  z Klientami"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DagKlienci, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuWybieraj.ResumeLayout(False)
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

    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim _CoMamRobic As String = ""
    Dim OdDaty As String = ""
    Dim DoDaty As String = ""
    Dim TakenDate As Date = Nothing

    ''---------------------------------------------------------------------
    ''Klienci
    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    'Public oPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
    'Public oTematKon As String            'TEMAT            Text, 50
    'Public oUwagiKon As String            'UWAGI            Memo
    'Public oWersjaKon As Integer          'WERSJA           Integer

    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub KatalogKlientowICoDalej_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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


    Private Sub KatalogKlientowICoDalej_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        PokazCoDalej()
        '---------------------------------------
        StartFormy = False
        Me.WindowState = FormWindowState.Normal
        InicjujFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        dagKlienci_CurrentCellChanged(sender, e)
        '--------------------------
        'zegar do scrollingu poziomego
        HorizScrollLastTime = ""
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
    End Sub





    Private Sub PokazCoDalej()

        ''---------------------------------------------------------------------
        ''CoDalej
        'Public cdUNIQID As String      'UNIQID Text, 40
        ' Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
        ' Public cdIDENT As String             'IDENT        Text, 40
        ' Public cdOPIS As String              'OPIS         memo
        'Public cdWersja As Integer           'WERSJA       Integer

        dbSelectKlienci = "SELECT " &
                                   "CD.IDENTKLIENTA, " &
                                   "CD.IDENT," &
                                   "CD.OPIS," &
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
                                             "CD.UNIQID, " &
                                             "CD.WERSJA " &
                                            "FROM CoDalejPlan as CD INNER JOIN Klienci AS KL ON CD.IDENTKLIENTA=KL.IDENTKLIENTA " &
                                            "WHERE (KL.AKTYWNY='TRUE') "

        dbSelectKlienciSorted = dbSelectKlienci & " ORDER BY CD.IDENTKLIENTA"



        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionCoDalej)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectKlienciSorted, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLCoDalej, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLCoDalej, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLCoDalej, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            'DBDeleteCoDalej(dbCommandDelete, dbDataAdapter)
            'DBUpdateCoDalej(dbCommandUpdate, dbDataAdapter)
            'DBInsertCoDalej(dbCommandInsert, dbDataAdapter)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionCoDalej)
            sqlCommandSelect = New SqlClient.SqlCommand(dbSelectKlienciSorted, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLCoDalej, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLCoDalej, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLCoDalej, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            SQLDeleteCoDalej(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateCoDalej(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertCoDalej(sqlCommandInsert, sqlDataAdapter)

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
                    'dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapter.Fill(DataSetKlienci)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapter.Fill(DataSetKlienci)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("UNIQID")}

                'definiuj DataView
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                DataViewKlienci.AllowDelete = False
                DataViewKlienci.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                DagKlienci.BackgroundColor = System.Drawing.SystemColors.Control
                DagKlienci.GridColor = System.Drawing.SystemColors.ControlDark
                DagKlienci.ForeColor = System.Drawing.Color.Purple
                DagKlienci.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                DagKlienci.BorderStyle = BorderStyle.Fixed3D
                'dagKlienci.Dock = DockStyle.Fill
                DagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                DagKlienci.AllowUserToAddRows = False
                DagKlienci.AllowUserToDeleteRows = False
                DagKlienci.AllowUserToOrderColumns = True
                DagKlienci.AllowUserToResizeColumns = True
                DagKlienci.AllowUserToResizeRows = True
                DagKlienci.ReadOnly = True
                DagKlienci.SelectionMode = DataGridViewSelectionMode.CellSelect
                DagKlienci.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagKlienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                DagKlienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                DagKlienci.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                DagKlienci.ColumnHeadersVisible = True
                DagKlienci.ColumnHeadersHeight = 40
                DagKlienci.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                DagKlienci.RowHeadersVisible = True
                DagKlienci.RowHeadersWidth = 50
                DagKlienci.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                DagKlienci.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                DagKlienci.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                DagKlienci.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                DagKlienci.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                DagKlienci.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                DagKlienci.DefaultCellStyle.NullValue = ""
                DagKlienci.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                DagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.False   'single line...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                DagKlienci.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                DagKlienci.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                DagKlienci.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                DagKlienci.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                DagKlienci.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                DagKlienci.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                DagKlienci.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                DagKlienci.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                DagKlienci.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                DagKlienci.DataMember = DataSetKlienci.Tables(0).TableName
                'dagKlienci.DataSource = DataViewKlienci
                '-----------------------------------



                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "Ident.", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "Opis", 500, DataGridViewContentAlignment.MiddleLeft, True)
                '-------------------------------------
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "PayBack", 50, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Parzysty", 40, True)

                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Zainst. monitor", 50, True)
                NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "Akt. zakup mat.eksp.", 50, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "Akt. rozlicza strony druku", 50, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "Akt. korzysta z najmu", 50, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "Zaint. korzyst. z najmu", 50, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)
                NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "Rodz.Sprzêtu L", 100, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Rodz.Sprzêtu A", 100, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Rodz.Sprzêtu A3", 100, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, "Z³o¿ono", 40, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, "Promotor Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, "Marker  ", 50, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker P.", 50, True)
                LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "", 0, False)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)




                Me.Cursor = Cursors.WaitCursor
                DagKlienci.DataSource = DataViewKlienci
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                    DagKlienci.CurrentCell = DagKlienci(1, 0)
                    DagKlienci.CurrentCell.Selected = True
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
            stbKlienci.BackColor = System.Drawing.SystemColors.Control
            stbKlienci.ForeColor = System.Drawing.Color.DarkGreen
            stbKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbKlienci.Panels.Add("stbPanelIleRec")
            stbKlienci.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(0).Width = 80
            stbKlienci.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(0).Text = "Iloœæ rec.: " & DataSetKlienci.Tables(0).Rows.Count.ToString

            stbKlienci.Panels.Add("stbPanelWiersz")
            stbKlienci.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(1).Width = 80
            stbKlienci.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If DagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(1).Text = "Wiersz : "
            Else
                stbKlienci.Panels(1).Text = "Wiersz : " & DagKlienci.CurrentCell.RowIndex.ToString
            End If

            stbKlienci.Panels.Add("stbPanelKolumna")
            stbKlienci.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(2).Width = 80
            stbKlienci.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If DagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(2).Text = "Kolumna : "
            Else
                stbKlienci.Panels(2).Text = "Kolumna : " & DagKlienci.CurrentCell.ColumnIndex.ToString
            End If

            stbKlienci.Panels.Add("stbPanelSort")
            stbKlienci.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(3).Width = 150
            stbKlienci.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(3).Text = "Sort : "

            stbKlienci.Panels.Add("stbPanelSzablon")
            stbKlienci.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(4).Width = 150
            stbKlienci.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(4).Text = "Szablon: "

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
            stbKlienci.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbKlienci.Panels(5).Width = 150
            stbKlienci.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(5).Text = ""

            stbKlienci.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagklienci.TableStyles(0).RowHeaderWidth
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
            '---------------------------------
            DagKlienci.Focus()
            dagKlienci_CurrentCellChanged(Me, System.EventArgs.Empty)
        End If

        InitInfoUzytkownika() 'inicjuj zmienne
        '-----------------------------------------------------------------
        'set the header width to something apporpriate
        DagKlienci.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagklienci.RowHeaderWidth = 80 'use if no tablestyle
        cmdClearColWidth.Size = New Size(50, 40)
        cmdClearColWidth.Location = New System.Drawing.Point(DagKlienci.Location.X,
                                                             DagKlienci.Location.Y)
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetKlienci.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((DagKlienci.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (DagKlienci.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * DagKlienci.FirstDisplayedScrollingColumnIndex +
                        DagKlienci.GetCellDisplayRectangle(DagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((DagKlienci.Left + 4), (DagKlienci.Top + 4))
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
        For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
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
        dagKlienci_CurrentCellChanged(Me, Nothing)
        InicjujFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        '--------------------------------------------------
    End Sub


    Private Sub KatalogKlientowICoDalej_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub



    Private Sub cmdHarmonogramWizyt_Click(sender As Object, e As EventArgs) Handles cmdHarmonogramWizyt.Click
        Dim Form As New KatalogHarmonogramWizyt
        Form.Show()
    End Sub










    Private Sub dagKlienci_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DagKlienci.CurrentCellChanged
        Dim Txt As String = ""
        Dim words() As String
        Dim splitChars() As Char = {","c}
        Dim iwords As Integer = 0

        If Not StartFormy Then
            If DagKlienci.CurrentCell Is Nothing Then
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

                lblTemat.Text = ""
                lblNotatka.Text = ""
            Else
                stbKlienci.Panels(1).Text = "Wi: " & DagKlienci.CurrentCell.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & DagKlienci.CurrentCell.ColumnIndex.ToString

                If DataViewKlienci.Count = 0 Then
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

                    lblTemat.Text = ""
                    lblNotatka.Text = ""
                Else

                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "Ident.", 80, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "Opis", 500, DataGridViewContentAlignment.MiddleLeft, True)
                    ''-------------------------------------
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "PayBack", 50, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Parzysty", 40, True)

                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Zainst. monitor", 50, True)
                    'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "Akt. zakup mat.eksp.", 50, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "Akt. rozlicza strony druku", 50, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "Akt. korzysta z najmu", 50, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "Zaint. korzyst. z najmu", 50, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)
                    'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "Rodz.Sprzêtu L", 100, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Rodz.Sprzêtu A", 100, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Rodz.Sprzêtu A3", 100, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, "Sprzêt - iloœæ", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, "Z³o¿ono", 40, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, "Promotor Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker  ", 50, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Marker P.", 50, True)
                    'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, "", 0, False)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(56).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(57).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)


                    lblIdent.Text = GetTxtField(DagKlienci, 3)

                    'podziel linie na poszczególne czêœci
                    Txt = GetTxtField(DagKlienci, 5)
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

                    lblNazwaHandlowa.Text = GetTxtField(DagKlienci, 8)

                    oRodzSprzetuLKli = GetLogField(DagKlienci, 34)
                    oRodzSprzetuAKli = GetLogField(DagKlienci, 35)
                    oRodzSprzetuTKli = GetLogField(DagKlienci, 36)
                    lblRodzaj.Text = IIf(oRodzSprzetuLKli, "L", "") & IIf(oRodzSprzetuAKli, "A", "") & IIf(oRodzSprzetuTKli, "T", "")

                    lblKod.Text = GetTxtField(DagKlienci, 10)
                    lblMiejscowosc.Text = GetTxtField(DagKlienci, 9)
                    lblUlica.Text = GetTxtField(DagKlienci, 11)
                    lblNrDomu.Text = GetIntField(DagKlienci, 12).ToString("F0")
                    lblIdDomu.Text = GetTxtField(DagKlienci, 15)
                    lblTelefon.Text = GetTxtField(DagKlienci, 17)
                    lbleMail.Text = GetTxtField(DagKlienci, 19) 'OsobaKontaktowa(oIdentKli)

                    lblOpiekun.Text = GetTxtField(DagKlienci, 41)
                    lblOstKontakt.Text = "Tydzieñ " & Trim(GetTxtField(DagKlienci, 43)) & " / " & GetTxtField(DagKlienci, 42)
                    lblNastKontakt.Text = "Tydzieñ " & Trim(GetTxtField(DagKlienci, 46)) & " / " & GetTxtField(DagKlienci, 45)
                    lblTransakcje.Text = OstTransakcje(lblIdent.Text, True)
                    lblPotencjal.Text = GetTxtField(DagKlienci, 40)

                    lblTemat.Text = GetTxtField(DagKlienci, 1)
                    lblNotatka.Text = GetTxtField(DagKlienci, 2)
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
                If DagKlienci.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * DagKlienci.FirstDisplayedScrollingColumnIndex +
                                    DagKlienci.GetCellDisplayRectangle(DagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * DagKlienci.FirstDisplayedScrollingColumnIndex +
                                    DagKlienci.GetCellDisplayRectangle(DagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = DagKlienci.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                DagKlienci.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(DagKlienci, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            'FiltrujDataViewDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, DataViewKlienci, stbKlienci)
            FiltrujDataViewDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, DataViewKlienci, stbKlienci, oSchematFiltrowania)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, DagKlienci, Mid(sender.name, 6, 3), "klienci")
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
            CzyscFiltryDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, DataViewKlienci, stbKlienci)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbKlienci.Panels(0).Text = "Il.zap.: " & DataViewKlienci.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub
































    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagklienci_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DagKlienci.RowPostPaint
        Using b As SolidBrush = New SolidBrush(DagKlienci.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      DagKlienci.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub




    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagklienci_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagklienci.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagklienci, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagklienci, e.RowIndex, 4)

    '        Select Case UserName = Program_IdUzytkownika
    '            Case True
    '                e.CellStyle.ForeColor = Color.Red
    '            Case Else
    '                Select Case Mid(Upr, _Uprawnnienia_NrZnaku, 1)
    '                    Case "A"
    '                        e.CellStyle.ForeColor = Color.Green
    '                    Case "S"
    '                        e.CellStyle.ForeColor = Color.Purple
    '                    Case "U"
    '                        e.CellStyle.ForeColor = Color.Navy
    '                    Case Else
    '                        e.CellStyle.ForeColor = Color.Black
    '                End Select
    '        End Select
    '        '-----------------------
    '        ''zmieniamy BackColor
    '        'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
    '        'Select Case Wal
    '        '    Case "EUR"
    '        '        e.CellStyle.BackColor = Color.White
    '        '    Case Else
    '        'End Select
    '        '-----------------------
    '    End If
    'End Sub


    '**************************************************************
    '** procedura okreslania koloru w DataGrid
    '**************************************************************

    'Public Function Koloruj(ByVal iRow As Long) As Integer
    '    Dim Kolor As Integer = System.Drawing.Color.Purple.ToArgb
    '    Dim NrTofi As String = ""

    '    Try
    '        NrTofi = IIf(IsDBNull(DataViewKlienci.Item(iRow).Item("NRTOFITXT")), "", DataViewKlienci.Item(iRow).Item("NRTOFITXT"))
    '        If Len(Trim(NrTofi)) = 0 Then
    '            'brak numeru TOFI
    '            Kolor = System.Drawing.Color.Red.ToArgb
    '        Else
    '            Dim words() As String
    '            Dim splitChars() As Char = {","c}
    '            words = NrTofi.Split(splitChars, StringSplitOptions.None)
    '            If words.Length = 1 Then
    '                'jeden numer TOFI
    '                Kolor = System.Drawing.Color.Purple.ToArgb
    '            Else
    '                'wiêcej ni¿ jeden numer TOFI
    '                Kolor = System.Drawing.Color.Green.ToArgb
    '                'Kolor = System.Drawing.Color.Navy.ToArgb
    '            End If
    '        End If

    '    Catch ex As System.Exception
    '        MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace,
    '            "Uwaga :",
    '            System.Windows.Forms.MessageBoxButtons.OK,
    '            MessageBoxIcon.Information,
    '            MessageBoxDefaultButton.Button1)
    '    Finally
    '    End Try

    '    Return (Kolor)
    'End Function




    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagklienci_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DagKlienci.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagklienci_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles DagKlienci.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagklienci_Scroll(sender As Object, e As ScrollEventArgs) Handles DagKlienci.Scroll
        'If (Not StartFormy) And (DataViewklienci.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagklienci, DataViewklienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagklienci, DataViewklienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewKlienci.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagklienci_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DagKlienci.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagklienci_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DagKlienci.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagklienci, DataViewklienci.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagklienci_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DagKlienci.MouseUp
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
                myGridView.CurrentCell = DagKlienci(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbKlienci.Panels(3).Text = "Sort: " & DataSetKlienci.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagklienci(hti.ColumnIndex, hti.RowIndex)
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




    Private Sub dagKlienci_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DagKlienci.DoubleClick
        If _CoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
            'cmdEdytuj_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagKlienci_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DagKlienci.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _CoMamRobic = "WYBIERAJ" Then
                    cmdStop_Click(sender, e)
                    'cmdEdytuj_Click(sender, e)
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

    'Public cdUNIQID As String      'UNIQID Text, 40
    ' Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
    ' Public cdIDENT As String             'IDENT        Text, 40
    ' Public cdOPIS As String              'OPIS         memo
    'Public cdWersja As Integer           'WERSJA       Integer

    Private Sub InitKlienci()
        cdUNIQID = Guid.NewGuid.ToString
        cdIDENTKLIENTA = ""
        cdIDENT = ""
        cdOPIS = ""
        cdWersja = 1          'WERSJA         Integer, 2
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
    End Sub

    'Public cdUNIQID As String      'UNIQID Text, 40
    ' Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
    ' Public cdIDENT As String             'IDENT        Text, 40
    ' Public cdOPIS As String              'OPIS         memo
    'Public cdWersja As Integer           'WERSJA       Integer

    Private Sub PobierzKlienci()
        'pobierz wartosci przed aktualizacja
        cdUNIQID = GetTxtField(DagKlienci, 62)
        cdIDENTKLIENTA = GetTxtField(DagKlienci, 0)
        cdIDENT = GetTxtField(DagKlienci, 1)
        cdOPIS = GetTxtField(DagKlienci, 2)
        cdWersja = GetIntField(DagKlienci, 63)
        '-------------------------
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "Ident.", 80, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "Opis", 500, DataGridViewContentAlignment.MiddleLeft, True)
        ''-------------------------------------
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "PayBack", 50, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
        'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Parzysty", 40, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Zainst. monitor", 50, True)
        'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
        'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "Akt. zakup mat.eksp.", 50, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "Akt. rozlicza strony druku", 50, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "Akt. korzysta z najmu", 50, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "Zaint. korzyst. z najmu", 50, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)
        'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
        'NumColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "Rodz.Sprzêtu L", 100, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Rodz.Sprzêtu A", 100, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Rodz.Sprzêtu A3", 100, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, "Sprzêt - iloœæ", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, "Z³o¿ono", 40, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, "Promotor Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, "Promotor Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, "Promotor Ost.kontakt dzien", 150, DataGridViewContentAlignment.MiddleCenter, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, "Promotor Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, "Promotor Kol.kontakt tydz", 150, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, "Promotor Kol.kontakt dzien", 150, DataGridViewContentAlignment.MiddleCenter, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, "Promotor Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(56).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(57).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(58).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(59).ColumnName, "Marker  ", 50, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(60).ColumnName, "Marker P.", 50, True)
        'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(61).ColumnName, "", 0, False)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(62).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
        'TxtColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(63).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)



        oIdentKli = GetTxtField(DagKlienci, 3)
        oNIPKli = GetTxtField(DagKlienci, 4)
        oNrTOFITxtKli = GetTxtField(DagKlienci, 5)
        oKartaPaybackKli = GetLogField(DagKlienci, 6)
        oNryKartPaybackKli = GetTxtField(DagKlienci, 7)

        oNazwa1Kli = GetTxtField(DagKlienci, 8)
        oMiejscKli = GetTxtField(DagKlienci, 9)
        oKodPoczKli = GetTxtField(DagKlienci, 10)
        oUlicaKli = GetTxtField(DagKlienci, 11)
        oNumNrDomuKli = GetIntField(DagKlienci, 12)
        '13 parzyste/nieparzyste
        oNrDomuKli = GetTxtField(DagKlienci, 14)
        oIDDomuKli = GetTxtField(DagKlienci, 15)
        oRejonKli = GetTxtField(DagKlienci, 16)
        oTelefonKli = GetTxtField(DagKlienci, 17)
        oFaxKli = GetTxtField(DagKlienci, 18)
        oeMailKli = GetTxtField(DagKlienci, 19)

        oWykazUrzadzenKli = GetTxtField(DagKlienci, 20)
        oSposobWyboruDostawcyKli = GetTxtField(DagKlienci, 21)
        oZainstalowanoMonitorKli = GetLogField(DagKlienci, 222)
        oLiczbaUrzadzenKli = GetNumField(DagKlienci, 23)
        oLiczbaWydrukowKli = GetNumField(DagKlienci, 24)
        oPotencjalDrukuKli = GetTxtField(DagKlienci, 25)
        oAktZakupMatEkspKli = GetLogField(DagKlienci, 26)
        oAktRozliczaStronyDrukuKli = GetLogField(DagKlienci, 27)
        oAktKorzystaZNajmuKli = GetLogField(DagKlienci, 28)
        oZaintRozliczaniemStronDrukuKli = GetLogField(DagKlienci, 29)
        oZaintKorzystaniemZNajmuKli = GetLogField(DagKlienci, 30)
        oDataWeryfSegmentacjiKli = GetTxtField(DagKlienci, 31)
        oPlanDlugookresowyKli = GetNumField(DagKlienci, 32)
        oPlanKrotkookresowyKli = GetNumField(DagKlienci, 33)

        oRodzSprzetuLKli = GetLogField(DagKlienci, 34)
        oRodzSprzetuAKli = GetLogField(DagKlienci, 35)
        oRodzSprzetuTKli = GetLogField(DagKlienci, 36)
        oOfertaTZlozeniaKli = GetTxtField(DagKlienci, 37)
        oOfertaPlikKli = GetTxtField(DagKlienci, 38)
        oPKontaktKli = GetTxtField(DagKlienci, 39)
        oSkupUwagikli = GetTxtField(DagKlienci, 40)
        oSprzedOpiekunkli = GetTxtField(DagKlienci, 41)
        oSprzedOKontaktRKli = GetTxtField(DagKlienci, 42)
        oSprzedOKontaktTKli = GetTxtField(DagKlienci, 43)
        oSprzedOKontaktDKli = GetTxtField(DagKlienci, 44)
        oSprzedNKontaktRKli = GetTxtField(DagKlienci, 45)
        oSprzedNKontaktTKli = GetTxtField(DagKlienci, 46)
        oSprzedNKontaktDKli = GetTxtField(DagKlienci, 47)
        oSprzedUwagiKli = GetTxtField(DagKlienci, 48)
        oWpisalKli = GetTxtField(DagKlienci, 49)
        oUwagikli = GetTxtField(DagKlienci, 50)
        oMarkerKli = GetLogField(DagKlienci, 51)
        oMarkerPKli = GetLogField(DagKlienci, 52)
        oAktywnyKli = GetLogField(DagKlienci, 53)
        'uniqid
        oWersjaKli = GetTxtField(DagKlienci, 55)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdWybierzSchemat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierzSchemat.Click
        oCoMamRobic = "WYBIERAJ"
        oLokalnyFiltr = ""
        oNazwaSchematu = ""
        oSchematFiltrowania = ""
        Dim Form As New SchematFiltrowania("KatalogKlientowICoDalej")
        Form.ShowDialog()

        'If Len(Trim(oSchematFiltrowania)) > 0 Then
        stbKlienci.Panels(4).Text = "Szablon : " & oNazwaSchematu
        Try
            'DataViewKlienci.RowFilter = ""
            ''DataViewKlienci.RowFilter = Trim(cbxFilter.SelectedItem.ToString) & " LIKE '" & Trim(txtFilter.Text) & "*'"
            'DataViewKlienci.RowFilter = BudujFiltr(Trim(oSchematFiltrowania), oLokalnyFiltr)
            'dagKlienci.DataSource = DataViewKlienci
            'stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString

            FiltrujDataViewDGV(Me,
                            DagKlienci,
                            DataViewKlienci.Table.Columns.Count,
                            DataViewKlienci,
                            stbKlienci,
                            Trim(oSchematFiltrowania))

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        End Try
        'Else
        '   stbKlienci.Panels(4).Text = "Szablon : "
        'End If

        If DataViewKlienci.Count > 0 Then
            'ustaw sie na poczatku DataGrid
            DagKlienci.CurrentCell = DagKlienci(1, 0)
            DagKlienci.CurrentCell.Selected = True
        End If
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        ZapamietajSzerokosciKolumn()
        '-------------------------------
        If DataViewKlienci.Count > 0 Then
            oKlient = GetTxtField(DagKlienci, 1)        'IDENT       Text    10
        Else
            oKlient = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub






    'Public cdUNIQID As String      'UNIQID Text, 40
    ' Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
    ' Public cdIDENT As String             'IDENT        Text, 40
    ' Public cdOPIS As String              'OPIS         memo
    'Public cdWersja As Integer           'WERSJA       Integer

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzKlienci()
            Dim Form As New EdytujKatalogCoDalej
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
                findTheseVals(0) = cdUNIQID
                foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("UNIQID") = cdUNIQID
                    foundRow("IDENTKLIENTA") = cdIDENTKLIENTA
                    foundRow("IDENT") = cdIDENT
                    foundRow("OPIS") = cdOPIS
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


                    'wyswietl ilosc rekordow
                    stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    DagKlienci.Update()
                    dagKlienci_CurrentCellChanged(sender, e)
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
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
                findTheseVals(0) = cdUNIQID
                foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
                    AktualizujBaze()
                    dagKlienci_CurrentCellChanged(sender, e)
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitKlienci()
        '----------------------------------------------
        Dim Form As New EdytujKatalogCoDalej
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
                findTheseVals(0) = cdUNIQID
                foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "UniqID = " & cdUNIQID,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetKlienci.Tables(0).NewRow()

                    NewRow("UNIQID") = cdUNIQID
                    NewRow("IDENTKLIENTA") = cdIDENTKLIENTA
                    NewRow("IDENT") = cdIDENT
                    NewRow("OPIS") = cdOPIS
                    NewRow("WERSJA") = 1
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
                    DataSetKlienci.Tables(0).Rows.Add(NewRow)
                    'wyswietl ilosc rekordow
                    stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    DagKlienci.Update()
                    '--------------------------------------------------------------------
                    'reinit HeightSetter
                    'RowHeightSetter.Dispose()
                    'RowHeightSetter = New Softart.myDataGridRowHeightSetter(Me.dagKlienci, 1, 2)
                    dagKlienci_CurrentCellChanged(sender, e)
                    Exit Do
                End If
            End If
        Loop
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
            '        dbDataAdapter.Update(DataSetKlienci, "TABELA_Klienci")
            '    Catch Ex As System.Exception
            '        'ViewInLog(ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetKlienci)
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
                    sqlDataAdapter.Update(DataSetKlienci, "TABELA_Klienci")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetKlienci)
                End If
                sqlConnection.Close()
            End If
        End If

        stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
    End Sub

    Private Sub OdswiezBaze()
        DataSetKlienci.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetKlienci)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetKlienci)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub














    Private Sub cmdOdswiez_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOdswiez.Click
        If DataSetKlienci.Tables(0).Rows.Count = 0 Then
            MessageBox.Show("Brak informacji o planach Co Dalej...",
                    "Uwaga:",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
        Else

            '...........................
            'definiuj delegat
            Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
            Dim FormRaport As New RaportWybierzAnalizeGrupowaKontaktow(_dsKlenci,
                                                                 DataViewKlienci,
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



















    '*************************************************
    '** obs³uga menu kontekstowego
    '*************************************************

    Private Sub MenuWEdytuj_Click(sender As Object, e As EventArgs) Handles MenuWEdytuj.Click
        cmdEdytuj_Click(sender, e)
    End Sub

    Private Sub MenuWDodaj_Click(sender As Object, e As EventArgs) Handles MenuWDodaj.Click
        cmdDodaj_Click(sender, e)
    End Sub

    Private Sub MenuWUsuñ_Click(sender As Object, e As EventArgs) Handles MenuWUsuñ.Click
        cmdUsuñ_Click(sender, e)
    End Sub

    Private Sub MenuWSingleL_Click(sender As Object, e As EventArgs) Handles MenuWSingleL.Click
        DagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        DagKlienci.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        DagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DagKlienci.Refresh()
    End Sub

    Private Sub menuWOdswiez_Click(sender As Object, e As EventArgs) Handles menuWOdswiez.Click
        OdswiezBaze()
    End Sub








    '------------------------
    ' wybierz na podstawie Akcji Marketingowej
    '--------------------------

    Private Sub MenuWWybierzA_Click(sender As Object, e As EventArgs) Handles MenuWWybierzA.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        'Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        'Dim Form As New AkcjaMarketingowaWybierz(DataSetKlienci, deleg)
        Dim Form As New AkcjaMarketingowaWybierz(DataSetKlienci, Nothing)
        Form.ShowDialog()

        If oWybranoAkcjeMarketingowa Then
            'aktualizuj Katalog klientow
            'CzyscFiltryDGV(Me,
            '            DagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'CzyscFiltryDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, DataViewKlienci, stbKlienci)


            'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(59).ColumnName, "Marker  ", 50, True)
            'LogColumnView(DagKlienci, DataSetKlienci.Tables(0).Columns(60).ColumnName, "Marker P.", 50, True)
            'pole MARKER - nr 59 / MARKERP = nr 60
            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi060" Or ctrl.Name = "txtFi059" Then
                        ctrl.Text = "T"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            DagKlienci.Update()
            'ustaw sie na pierwszym zapisie
            'If DataSetKlienci.Tables(0).Rows.Count > 0 Then
            If DataViewKlienci.Count > 0 Then
                DagKlienci.CurrentCell = DagKlienci(1, 0)
                DagKlienci.CurrentCell.Selected = True
            End If
            '--------------------------------------------------------------------
            'reinit HeightSetter
            'RowHeightSetter.Dispose()
            'RowHeightSetter = New Softart.myDataGridRowHeightSetter(Me.dagKlienci, 3, 4)
            dagKlienci_CurrentCellChanged(sender, e)
        End If
    End Sub







    '**************************************************************
    '** Zapamietaj szerokosci kolumn
    '**************************************************************

    Public Sub ZapamietajSzerokosciKolumn()
        Dim ii As Integer = 0
        Dim Txt As String = ""
        Dim FileNum As Integer

        For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
            Txt &= Trim(Str(DagKlienci.Columns(ii).Width)) & "|"
        Next

        If IO.File.Exists(oKatParametry & "\" & oPlikSzerokosciKolumnKlienciICoDalej) Then
            IO.File.Delete(oKatParametry & "\" & oPlikSzerokosciKolumnKlienciICoDalej)
        End If
        FileNum = FreeFile()    'kolejny nr otwartego zbioru
        FileOpen(FileNum, oKatParametry & "\" & oPlikSzerokosciKolumnKlienciICoDalej, OpenMode.Output)
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

        If IO.File.Exists(oKatParametry & "\" & oPlikSzerokosciKolumnKlienciICoDalej) Then
            FileNum = FreeFile()    'kolejny nr otwartego zbioru
            FileOpen(FileNum, oKatParametry & "\" & oPlikSzerokosciKolumnKlienciICoDalej, OpenMode.Input)
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
                If NrWersjiPliku <> oAktualnaWersjaBazyDanych Then
                Else
                    For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
                        pos = InStr(Txt, "|")
                        If pos > 0 Then
                            DagKlienci.Columns(ii).Width = CInt(Mid(Txt, 1, pos - 1))
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
        For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
            'dagKlienci.Columns.Item(ii).Width = GetIntFromText(dagKlienci.Columns.Item(ii).Tag)
            If IsNumeric(DagKlienci.Columns.Item(ii).Tag) Then
                DagKlienci.Columns.Item(ii).Width = DagKlienci.Columns.Item(ii).Tag
            Else
                DagKlienci.Columns.Item(ii).Width = 0
            End If
        Next
        StartFormy = False
        dagKlienci_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, DagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        Me.Cursor = Cursors.Default
        ''--------------------------------
        ZapamietajSzerokosciKolumn()
    End Sub

End Class
